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


const isWriteOffDateValid = ({toBeWrittenOffAtt}) => {
    if (typeof toBeWrittenOffAtt === 'string') return false
    return toBeWrittenOffAtt instanceof Date;
}

const filterWriteOffs = writeOffs => {
    return writeOffs.filter(isWriteOffDateValid).map(({toBeWrittenOffAtt, isChecked}) => {
        return {isChecked: isChecked, toBeWrittenOffAtt: TimeUtilities.normalizeDate(toBeWrittenOffAtt)}
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
        return this.worksheet.getRange(range).getValues().map(row => {
            const [toBeWrittenOffAtt, isChecked] = row
            return {toBeWrittenOffAtt, isChecked}
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

        const rawWriteOffs = this.worksheetSelector.getWriteOffsByWeekday(weekday)

        const filteredWriteOffs = filterWriteOffs(rawWriteOffs)

        const worksheetEvents = new Set()

        filteredWriteOffs.forEach(writeOff => {

            const timeBeforeExpire = calculateTimeBeforeExpire({
                toBeWrittenOffAtt: writeOff.toBeWrittenOffAtt,
                now: now,
            })

            this.eventFilters.forEach(eventFilter => {

                if (eventFilter.isSatisfied(timeBeforeExpire)) {
                    worksheetEvents.add(eventFilter.eventType)
                }

            })
        })

        return {unitName: this.worksheet.getName(), events: Array.from(worksheetEvents)}

    }

}

function findWriteOffsInWorksheets({writeOffHandlers, now}) {
    return writeOffHandlers
        .map(handler => handler.findWriteOffs({now}))
        .filter(({events}) => events.length !== 0)
}


function main() {
    const now = TimeUtilities.now()
    const eventFilters = [
        new AlreadyExpiredFilter(600, 30),
        new TimeBeforeExpireFilter("EXPIRE_AT_5_MINUTES", 270, 330),
        new TimeBeforeExpireFilter("EXPIRE_AT_10_MINUTES", 570, 630),
        new TimeBeforeExpireFilter("EXPIRE_AT_15_MINUTES", 870, 930),
    ]

    const worksheets = SpreadsheetApp.getActive().getSheets()

    const writeOffHandlers = worksheets.map(worksheet => {
        return new WorksheetWriteOffsHandler({eventFilters, worksheet})
    })

    const worksheetsWriteOffs = findWriteOffsInWorksheets({writeOffHandlers, now})

    console.log(worksheetsWriteOffs)
}
