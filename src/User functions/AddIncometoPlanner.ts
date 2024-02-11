function __AddIncomeRow() {
    const UI = SpreadsheetApp.getUi()
    const YEAR_INPUT = Number(UI.prompt("What year is this income for?", UI.ButtonSet.OK).getResponseText())
    const INCOME_LABEL = UI.prompt("What is the income label?", UI.ButtonSet.OK).getResponseText()
    const INCOME_SHEET = new GoogleSheetTabs("Settings")
    const TRUNCED_MONTHS = MONTHS.map(month => [month.slice(0, 3)])

    const NEW_INCOME_TEMPLATE: DataArray = [
        [INCOME_LABEL],
        [YEAR_INPUT, "Pay Schedule Start Date", "Frequency", "Take Home Pay", "Pay for Saving", "Pay for LE", "Pay for Personal Use"],
        ...TRUNCED_MONTHS
    ];

    if (isNaN(YEAR_INPUT)) {
        UI.alert("Please enter a valid year")
        return
    }

    while (INCOME_SHEET.NumberOfRows() < 6) { INCOME_SHEET.AppendRow([]) }

    let row = INCOME_SHEET.GetRow(5)!

    if (row.length === 0) {
        INCOME_SHEET.AppendToRow(5, INCOME_LABEL)
    }
    else if (row.indexOf(INCOME_LABEL) === -1) {
        INCOME_SHEET.AppendToRow(5, " ", INCOME_LABEL)
    }
    NEW_INCOME_TEMPLATE.shift()

    const START_OF_ROW = INCOME_SHEET.GetRow(5)!.indexOf(INCOME_LABEL)
    
    for (let i = 0; i < INCOME_SHEET.NumberOfRows(); i++) {
        let row = INCOME_SHEET.GetRow(i)!
        if (isNaN(Number(row[START_OF_ROW]))) { continue }
        if (Number(row[START_OF_ROW]) === YEAR_INPUT) { break }
        if (Number(row[START_OF_ROW]) > YEAR_INPUT) {
            for (let j = 0; j < NEW_INCOME_TEMPLATE.length; j++) {
                INCOME_SHEET.InsertRow(i + j, [], {should_fill: true})
                INCOME_SHEET.WriteRowAt(i + j, START_OF_ROW, NEW_INCOME_TEMPLATE[j])
            }
            break
        }
    }

    let index = INCOME_SHEET.IndexOfRow(row => row.indexOf(YEAR_INPUT) !== -1)
    if (index === -1) {
        for (let i = 0; i < NEW_INCOME_TEMPLATE.length; i++) {
            INCOME_SHEET.AppendRow([])
            INCOME_SHEET.WriteRowAt(INCOME_SHEET.NumberOfRows() - 1, START_OF_ROW, NEW_INCOME_TEMPLATE[i])
        }
    }
    else {
        for (let i = 0; i < NEW_INCOME_TEMPLATE.length; i++) {
            INCOME_SHEET.WriteRowAt(index + i, START_OF_ROW, NEW_INCOME_TEMPLATE[i])
        }
    }

    INCOME_SHEET.SaveToTab()
}

function AddIncomeToPlanner() {
    __AddIncomeRow()
}