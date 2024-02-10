function __AddIncomeRow() {
    const UI = SpreadsheetApp.getUi()
    const YEAR_INPUT = UI.prompt("What year is this income for?", UI.ButtonSet.OK_CANCEL).getResponseText()
    const INCOME_LABEL = UI.prompt("What is the income label?", UI.ButtonSet.OK_CANCEL).getResponseText()

    const BUDGET_PLANNER_TAB = new GoogleSheetTabs(BUDGET_PLANNER_TAB_NAME)
    const JAN_LABEL_COL = BUDGET_PLANNER_TAB.GetHeaderIndex("January")

    const INCOME_YEAR = Number(YEAR_INPUT)
    const INCOME_PER_MONTH_COL = 1
    const INCOME_STREAM_COL = 1
    const INCOME_YEAR_COL = 0

    const __IncomePerMonthRow = () => BUDGET_PLANNER_TAB.IndexOfRow(row => row[INCOME_PER_MONTH_COL] === "Income Per Paycheck")

    const __IncomeStreamRow = () => BUDGET_PLANNER_TAB.IndexOfRow(row => row[INCOME_STREAM_COL] === "Income Stream")

    const __AddValidationToRow = function () {
        const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        const VALIDATION_OPTIONS = [...WEEKDAYS.map(day => PAYMENT_SCHEDULE.map(pay => `${day} : ${pay}`)).flat(), "Custom", "N/A"]
        const VALIDATION = SpreadsheetApp
            .newDataValidation()
            .requireValueInList(VALIDATION_OPTIONS)
            .build()

        for (let i = __IncomePerMonthRow(); i <= __IncomeStreamRow(); i++) {
            BUDGET_PLANNER_TAB.GetRowRange(i)?.setDataValidation(null)
        }

        for (let j = __IncomeStreamRow() + 1; j < BUDGET_PLANNER_TAB.NumberOfRows(); j++) {
            const ROW = BUDGET_PLANNER_TAB.GetRow(j)!
            for (let i = JAN_LABEL_COL - 1; i < ROW.length; i += 2) {
                BUDGET_PLANNER_TAB.GetTab().getRange(j + 1, i + 1).setDataValidation(VALIDATION)
                if (ROW[i] !== "") { continue }
                ROW[i] = VALIDATION_OPTIONS[VALIDATION_OPTIONS.length - 1]
                ROW[i + 1] = "-"
            }
            BUDGET_PLANNER_TAB.WriteRow(j, ROW)
        }
    }

    const __ValidateUserInput = function () {
        if (!isNaN(INCOME_YEAR) && INCOME_LABEL !== "") { return true }
        let alert = ""
        if (isNaN(INCOME_YEAR)) { alert = `Please enter a valid year instead of '${YEAR_INPUT}'` }
        else if (INCOME_LABEL === "") { alert = "Please enter a valid income label (anything but an empty string)" }
        UI.alert(alert)
        return false
    }

    const __InsertRow = function (row: DataArrayEntry, row_index: number) {
        if (row[INCOME_PER_MONTH_COL] === "") {
            BUDGET_PLANNER_TAB.InsertRow(row_index, [INCOME_YEAR, INCOME_LABEL], { should_fill: true })
            return true
        }
        if (Number(row[INCOME_YEAR_COL]) <= INCOME_YEAR) { return false }
        BUDGET_PLANNER_TAB.InsertRow(row_index, [INCOME_YEAR, INCOME_LABEL], { should_fill: true })
        return true
    }

    if (!__ValidateUserInput() || __IncomePerMonthRow() === -1 || __IncomeStreamRow() === -1) { return }

    const STOP = __IncomeStreamRow()
    for (let i = __IncomePerMonthRow() + 1; i < STOP; i++) {
        if (__InsertRow(BUDGET_PLANNER_TAB.GetRow(i)!, i)) { break }
    }

    const __ShouldLoop = (i: number) => i < BUDGET_PLANNER_TAB.NumberOfRows()
    const START = __IncomeStreamRow() + 1
    let inserted = false
    for (let i = START; __ShouldLoop(i); i++) {
        inserted = __InsertRow(BUDGET_PLANNER_TAB.GetRow(i)!, i)
        if (inserted) { break }
    }

    if (!__ShouldLoop(START) || !inserted) {
        BUDGET_PLANNER_TAB.AppendRow([INCOME_YEAR, INCOME_LABEL], true)
    }

    __AddValidationToRow()

    BUDGET_PLANNER_TAB.SaveToTab()
}

function AddIncomeToPlanner() {
    __AddIncomeRow()
}