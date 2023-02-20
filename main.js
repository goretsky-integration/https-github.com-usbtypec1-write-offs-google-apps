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
        return enumerate(rows, 2).map(enumeratedRow => {
            const [rowNumber, writeOff] = enumeratedRow
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
        worksheet.getRange({row, column}).setBackground(specification.color)
    })
}


function main() {
    const now = TimeUtilities.now()
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

    const worksheets = SpreadsheetApp.getActive().getSheets()

    const writeOffHandlers = worksheets.map(worksheet => {
        return new WorksheetWriteOffsHandler({eventFilters, worksheet})
    })

    const worksheetsWriteOffs = findWriteOffsInWorksheets({writeOffHandlers, now})

    paintBySpecification({
        worksheets: worksheets,
        specifications: paintSpecifications,
        writeOffs: worksheetsWriteOffs},
    )
}
