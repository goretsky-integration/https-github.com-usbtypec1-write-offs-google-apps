class UnitsManager {
  constructor(units) {
    this.units = units;
  }

  get names() {
    return this.units.map((unit) => unit.name);
  }

  getByNameOrNull(name) {
    const filtered = this.units.filter((unit) => unit.name == name);
    if (filtered.length === 0) return
    return filtered[0];
  }
}

function enumerate(iterable, start = 0) {
  return iterable.map((elem) => [start++, elem]);
}

class TimeBeforeExpireFilter {
  constructor(eventType, rangeStart, rangeStop) {
    this.eventType = eventType;
    this.rangeStart = rangeStart;
    this.rangeStop = rangeStop;
  }

  isSatisfied(timeBeforeExpireInSeconds) {
    return this.rangeStart <= timeBeforeExpireInSeconds && timeBeforeExpireInSeconds <= this.rangeStop
  }
}

class AlreadyExpiredFilter {
  constructor(intervalInSeconds, deviationInSeconds) {
    this.eventType = "ALREADY_EXPIRED";
    this.intervalInSeconds = intervalInSeconds;
    this.deviationInSeconds = deviationInSeconds;
  }

  isSatisfied(passedSeconds) {
    if (passedSeconds > this.deviationInSeconds) return false
    const intervalMultiplier = Math.round(
      passedSeconds / this.intervalInSeconds
    );
    const comparableThreshold = intervalMultiplier * this.intervalInSeconds;
    const fromSecondsThreshold = comparableThreshold - this.deviationInSeconds;
    const toSecondsThreshold = comparableThreshold + this.deviationInSeconds;
    return (
      fromSecondsThreshold <= passedSeconds &&
      passedSeconds <= toSecondsThreshold
    );
  }
}

class WriteOffsAPI {
  constructor(serverUrl) {
    this.serverUrl = serverUrl;
  }

  createEvents(events) {
    const response = UrlFetchApp.fetch(`${this.serverUrl}/events/`, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(events),
    });
    Logger.log(response.getResponseCode())
  }
}

class DatabaseAPI {
  constructor(serverUrl) {
    this.serverUrl = serverUrl;
  }

  getUnits() {
    const url = `${this.serverUrl}/units/`;
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      contentType: "application/json",
    });
    if (response.getResponseCode() !== 200) {
      throw Error("Could not get units from API");
    }
    return JSON.parse(response.getContentText());
  }
}

class WorksheetManager {
  constructor(worksheet) {
    this.rowsRange = "A2:O";
    this.worksheet = worksheet;
  }

  static getAllSheets() {
    return SpreadsheetApp.getActive().getSheets();
  }

  static isSheetNameAllowed(worksheet, allowedNames) {
    return allowedNames.includes(worksheet.getName());
  }

  getRows() {
    return this.worksheet.getRange(this.rowsRange).getValues();
  }

  getEnumeratedRows() {
		const rows = this.getRows()
      return enumerate(rows, 2).map(enumeratedRow => {
        const [rowNumber, rowData] = enumeratedRow
        return [rowNumber, ...rowData]
      })
	}
}

class TimeHelper {

  static getMoscowNow() {
    const date = new Date();
    date.setHours(date.getHours() + 3);
    return date;
  }

  static getWeekdayNumber() {
    const day = this.getMoscowNow().getDay();
    return day === 0 ? 7 : day
  }

  static normalizeDate(date) {
    const now = this.getMoscowNow();
    return new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds()
    );
  }
}

class DailyWriteOff {
  constructor(
    ingredientName,
    writeOffAt,
    isWrittenOff,
    rowNumber,
    weekdayNumber
  ) {
    this.ingredientName = ingredientName;
    this._writeOffAt = writeOffAt;
    this.isWrittenOff = isWrittenOff;
    this.rowNumber = rowNumber;
    this.weekdayNumber = weekdayNumber;
  }

  get writeOffAt() {
    return TimeHelper.normalizeDate(this._writeOffAt);
  }

  calculateTimeBeforeExpire(timeNow) {
    return (this.writeOffAt.getTime() - timeNow) / 1000
  }

  isValid() {
    if (!(typeof this.ingredientName === "string")) return false;
    if (!this.ingredientName) return false;
    if (!(typeof this.isWrittenOff === "boolean")) return false;
    if (this.isWrittenOff) return
    if (!(this._writeOffAt instanceof Date)) return false;
    return true;
  }
}

class WeeklyWriteOff {
  constructor(row) {
    this.row = row;
  }

  get rowNumber() {
    return this.row[0]
  }

  get ingredientName() {
    return this.row[1]
  }

  getByWeekdayNumber(weekdayNumber) {
    const [writeOffAt, isWrittenOff] = this.row.slice(
      weekdayNumber * 2,
      weekdayNumber * 2 + 2
    );
    return new DailyWriteOff(
      this.ingredientName,
      writeOffAt,
      isWrittenOff,
      this.rowNumber,
      weekdayNumber
    );
  }
}


// Does not save order of the elements
function removeDublicates(array) {
  return [...new Set(array)]
}


function preprareEventsData(data, unitsManager) {
  const result = []
  for (const {worksheet, events} of data) {
    const worksheetName = worksheet.getName()
    const unit = unitsManager.getByNameOrNull(worksheetName)
    if (unit === null) continue
    if (events.length === 0) continue
    result.push({unit_id: unit.id, unit_name: unit.name, events: removeDublicates(events)})
  }
  return result
}


const main = () => {
  const dateNow = TimeHelper.getMoscowNow()
  const filters = [
    new AlreadyExpiredFilter(600, 25),
    new TimeBeforeExpireFilter("EXPIRE_AT_5_MINUTES", 275, 325),
    new TimeBeforeExpireFilter("EXPIRE_AT_10_MINUTES", 575, 625),
    new TimeBeforeExpireFilter("EXPIRE_AT_15_MINUTES", 875, 925),
  ]
  const writeOffsAPI = new WriteOffsAPI("YOUR_API_URL")
  const databaseAPI = new DatabaseAPI("YOUR_API_URL")

  const unitsManager = new UnitsManager(databaseAPI.getUnits())
  const worksheets = WorksheetManager.getAllSheets().filter((worksheet) =>
    WorksheetManager.isSheetNameAllowed(worksheet, unitsManager.names)
  ) 

  const allEventTypes = worksheets.map(worksheet => {
    const worksheetManager = new WorksheetManager(worksheet)
    const worksheetRows = worksheetManager.getEnumeratedRows()

    const weeklyWriteOffs = worksheetRows.map(worksheetRow => new WeeklyWriteOff(worksheetRow))
    const dailyWriteOffs = weeklyWriteOffs.map(weeklyWriteOff => (weeklyWriteOff.getByWeekdayNumber(TimeHelper.getWeekdayNumber())))
                           .filter(dailyWriteOff => dailyWriteOff.isValid())

    const worksheetEvents = dailyWriteOffs.map(dailyWriteOff => {
      const timeBeforeExpire = dailyWriteOff.calculateTimeBeforeExpire(dateNow.getTime())
      return filters.filter(eventsFilter => eventsFilter.isSatisfied(timeBeforeExpire)).map(eventsFilter => eventsFilter.eventType)
    })
    return {worksheet: worksheet, events: worksheetEvents.flat()}
  })
  const data = preprareEventsData(allEventTypes, unitsManager)
  if (data.length === 0) return
  Logger.log(data)
  writeOffsAPI.createEvents(data)
};

