class ValidationError extends Error {
}

class SheetCursor {
    constructor({columnNumber, rowNumber}) {
        this.rowNumber = rowNumber
        this.columnNumber = columnNumber
    }

    static fromEvent({range}) {
        return new SheetCursor({columnNumber: range.getColumn(), rowNumber: range.getRow()})
    }

    get range() {
        return SpreadsheetApp.getActiveSheet().getRange(this.rowNumber, this.columnNumber)
    }

    clear() {
        this.range.clearContent()
    }

    get right() {
        return new SheetCursor({columnNumber: this.columnNumber + 1, rowNumber: this.rowNumber})
    }

    get left() {
        return new SheetCursor({columnNumber: this.columnNumber - 1, rowNumber: this.rowNumber})
    }

    get firstOfRow() {
        return new SheetCursor({columnNumber: 1, rowNumber: this.rowNumber})
    }

    get value() {
        return this.range.getValue()
    }

    get sheetName() {
        return SpreadsheetApp.getActiveSheet().getName()
    }
}

const columnNumberFactory = allowedColumns => event => allowedColumns.includes(event.range.getColumn())

const isTypeOfDate = value => value instanceof Date

const isDateInFuture = value => {
    return TimeUtilities.now().getTime() < value.getTime()
}

const validateDate = value => {
    const UI = SpreadsheetApp.getUi()
    console.log(value)
    if (!isTypeOfDate(value)) {
        throw new ValidationError('Неправильный формат времени (HH:MM, HH:MM:SS)')
    }
    if (!isDateInFuture(TimeHelper.normalizeDate(value))) {
        const userChoice = UI.alert('Вы точно хотите ввести уже прошедшее время списания?', UI.ButtonSet.YES_NO)
        if (userChoice == UI.Button.NO) {
            throw new ValidationError()
        }
    }
}

const isCellCleaned = event => typeof event.value === 'undefined'

const handleEdit = (event) => {
    const UI = SpreadsheetApp.getUi()
    const isWriteOffAtColumn = columnNumberFactory([2, 4, 6, 8, 10, 12, 14])
    const cursor = SheetCursor.fromEvent(event)
    if (!isWriteOffAtColumn(event)) return

    try {
        if (!isCellCleaned(event)) {
            validateDate(cursor.value)
        }
    } catch (error) {
        if (error instanceof ValidationError) {
            cursor.clear()
            if (error.message) UI.alert(error.message)
            return
        }
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


class DatabaseAPI {
    constructor(serverUrl) {
        this.serverUrl = serverUrl
    }

    getUnits() {
        const response = UrlFetchApp.fetch(`${this.serverUrl}/units/`)
        return JSON.parse(response.getContentText())
    }

    getUnitByName(name) {
        const response = UrlFetchApp.fetch(`${this.serverUrl}/units/name/${name}/`)
        return JSON.parse(response.getContentText())
    }
}


class WriteOffsAPI {
    constructor(serverUrl, token) {
        this.serverUrl = serverUrl;
        this.token = token
    }

    createEvents(events) {
        const response = UrlFetchApp.fetch(`${this.serverUrl}/events/`, {
            method: "POST",
            contentType: "application/json",
            headers: {
                Authorization: `Bearer ${this.token}`,
            },
            payload: JSON.stringify(events),
        });
        Logger.log(response.getResponseCode())
    }
}


class TimeUtilities {
    // All provided time and date attached to Moscow timezone (GMT+3)

    static now() {
        const secondsInHour = 3600
        const timezoneOffset = 3
        return new Date((new Date).getTime() + timezoneOffset * secondsInHour * 1000)
    }

    static getWeekdayFromDate(date) {
        const weekday = date.getDay()
        return weekday === 0 ? 7 : weekday
    }

    static weekday() {
        return TimeUtilities.getWeekdayFromDate(TimeUtilities.now())
    }

    static normalizeDate(date) {
        const now = TimeUtilities.now();
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

function calculateTimeBeforeExpire({toBeWrittenOffAtt, now}) {
    return (toBeWrittenOffAtt.getTime() - now) / 1000
}


function calculateColumnNumberByWeekday(weekday) {
    return {
        writeOffDatesColumnNumber: weekday * 2,
        checkboxesColumnNumber: weekday * 2 + 1,
    }
}

const columnNumberToChar = [null, 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']

const isWriteOffDateValid = ({toBeWrittenOffAtt}) => toBeWrittenOffAtt instanceof Date

const isCheckboxValid = ({isChecked}) => isChecked === false

const filterWriteOffs = writeOffs => {
    return writeOffs.filter(isWriteOffDateValid).filter(isCheckboxValid).map(writeOff => {
        return {...writeOff, toBeWrittenOffAtt: TimeUtilities.normalizeDate(writeOff.toBeWrittenOffAtt)}
    })
}


class WorksheetSelector {
    constructor(worksheet) {
        this.worksheet = worksheet
    }

    getWriteOffsByWeekday(weekdayNumber) {
        const {writeOffDatesColumnNumber, checkboxesColumnNumber} = calculateColumnNumberByWeekday(weekdayNumber)
        const dateColumnChar = columnNumberToChar[writeOffDatesColumnNumber]
        const checkboxColumnChar = columnNumberToChar[checkboxesColumnNumber]
        const worksheetName = this.worksheet.getName()
        const range = `${worksheetName}!${dateColumnChar}2:${checkboxColumnChar}`
        const rows = this.worksheet.getRange(range).getValues().map(row => {
            const [toBeWrittenOffAtt, isChecked] = row
            return {toBeWrittenOffAtt, isChecked}
        })
        return enumerate(rows, 2).map(([rowNumber, writeOff]) => {
            return {...writeOff, row: rowNumber, column: writeOffDatesColumnNumber}
        })
    }
}


class WorksheetWriteOffsHandler {
    constructor({eventFilters, worksheet}) {
        this.eventFilters = eventFilters
        this.worksheet = worksheet
        this.worksheetSelector = new WorksheetSelector(worksheet)
    }

    findWriteOffs({now}) {
        const weekday = TimeUtilities.getWeekdayFromDate(now)

        const worksheetName = this.worksheet.getName()

        const rawWriteOffs = this.worksheetSelector.getWriteOffsByWeekday(weekday)

        const filteredWriteOffs = filterWriteOffs(rawWriteOffs)

        const worksheetWriteOffsWithEvents = []

        filteredWriteOffs.forEach(writeOff => {

            const timeBeforeExpire = calculateTimeBeforeExpire({
                toBeWrittenOffAtt: writeOff.toBeWrittenOffAtt,
                now: now,
            })

            this.eventFilters.forEach(eventFilter => {

                if (eventFilter.isSatisfied(timeBeforeExpire)) {
                    worksheetWriteOffsWithEvents.push({
                        ...writeOff,
                        event: eventFilter.eventType,
                        unitName: worksheetName,
                    })
                }

            })
        })
        return worksheetWriteOffsWithEvents

    }

}

function findWriteOffsInWorksheets({writeOffHandlers, now}) {
    return writeOffHandlers.map(handler => handler.findWriteOffs({now})).flat()
}


class PaintSpecification {
    constructor({event, color}) {
        this.event = event
        this.color = color
    }
}


function paintBySpecification({worksheets, specifications, writeOffs}) {
    writeOffs.forEach(({event, unitName, row, column}) => {
        const specification = specifications.find(specification => specification.event === event)
        const worksheet = worksheets.find(worksheet => worksheet.getName() === unitName)
        if (typeof specification === 'undefined' || typeof worksheet === 'undefined') return
        worksheet.getRange(row, column).setBackground(specification.color)
    })
}


function prepareEventsPayload({writeOffs, units}) {
    const unitNameToEvents = {}
    writeOffs.forEach(writeOff => {
        if (!(writeOff.unitName in unitNameToEvents)) {
            unitNameToEvents[writeOff.unitName] = []
        }
        unitNameToEvents[writeOff.unitName].push(writeOff.event)
    })
    const payload = []
    Object.entries(unitNameToEvents).forEach(([unitName, events]) => {
        const unit = units.find(({name}) => name === unitName)
        if (typeof unit === 'undefined') return
        payload.push({unit_id: unit.id, unit_name: unit.name, events: events})
    })
    return payload
}


function main() {
    const now = TimeUtilities.now()

    const writeOffsAPI = new WriteOffsAPI('')
    const databaseAPI = new DatabaseAPI('')

    const units = databaseAPI.getUnits()
    const allowedSheetNames = units.map(({name}) => name)

    const eventFilters = [
        new AlreadyExpiredFilter(600, 30),
        new TimeBeforeExpireFilter("EXPIRE_AT_5_MINUTES", 270, 330),
        new TimeBeforeExpireFilter("EXPIRE_AT_10_MINUTES", 570, 630),
        new TimeBeforeExpireFilter("EXPIRE_AT_15_MINUTES", 870, 930),
    ]

    const paintSpecifications = [
        new PaintSpecification({event: 'ALREADY_EXPIRED', color: '#f45252'}),
        new PaintSpecification({event: 'EXPIRE_AT_5_MINUTES', color: '#f46052'}),
        new PaintSpecification({event: 'EXPIRE_AT_10_MINUTES', color: '#f47b52'}),
        new PaintSpecification({event: 'EXPIRE_AT_15_MINUTES', color: '#f4aa52'}),
    ]

    const worksheets = SpreadsheetApp
        .getActive()
        .getSheets()
        .filter(worksheet => allowedSheetNames.includes(worksheet.getName()))

    const writeOffHandlers = worksheets.map(worksheet => {
        return new WorksheetWriteOffsHandler({eventFilters, worksheet})
    })

    const worksheetsWriteOffs = findWriteOffsInWorksheets({writeOffHandlers, now})
    const eventsPayload = prepareEventsPayload({writeOffs: worksheetsWriteOffs, units: units})

    if (eventsPayload.length === 0) return

    try {
        writeOffsAPI.createEvents(eventsPayload)
    } catch (error) {
        console.log(error.message)
    }

    try {
        paintBySpecification({
            worksheets: worksheets,
            specifications: paintSpecifications,
            writeOffs: worksheetsWriteOffs,
        })

    } catch (error) {
        console.log(error)
    }
}
