function __AITP_CreateNewIncomeTemplate(year: number): DataArray {
    const TRUNCED_MONTHS = MONTHS.map(month => [month.slice(0, 3)])
    return [
        [year, "Pay Schedule Start Date", "Frequency", "Take Home Pay", "Pay for Saving", "Pay for LE", "Pay for Personal Use"],
        ...TRUNCED_MONTHS
    ];
}

function __AITP_FindAllYearsInIncomeSheet(income_sheet: GoogleSheetTabs) {
    let years: number[] = []
    for (let i = 6; i < income_sheet.NumberOfRows(); i++) {
        let row = Number(income_sheet.GetRow(i)![0])
        if (isNaN(row)) { continue }
        years.push(row)
    }
    return years
}

function __AITP_AddToIncomeEarnerRow(income_earner_name: string, income_sheet: GoogleSheetTabs) {
    const INCOME_EARNERS_ROW_INDEX = 5
    while (income_sheet.NumberOfRows() <= INCOME_EARNERS_ROW_INDEX) { 
        income_sheet.AppendRow([]) 
    }

    let row = income_sheet.GetRow(INCOME_EARNERS_ROW_INDEX)!
    if (row.length === 0) {
        income_sheet.AppendToRow(INCOME_EARNERS_ROW_INDEX, income_earner_name)
    }
    else if (row.indexOf(income_earner_name) === -1) {
        income_sheet.AppendToRow(INCOME_EARNERS_ROW_INDEX, " ", income_earner_name)
        let years = __AITP_FindAllYearsInIncomeSheet(income_sheet)
        for (let year of years) {
            let template = __AITP_CreateNewIncomeTemplate(year)
            let index = income_sheet.IndexOfRow(row => row.indexOf(year) !== -1)
            __AITP_AddYearNotInEarner(index, income_earner_name, income_sheet, template)
        }
    }

    return income_sheet.GetRow(INCOME_EARNERS_ROW_INDEX)!
}

function __AITP_FindIncomeYear(year: number, earner: string, income_sheet: GoogleSheetTabs) {
    const EARNER_INDEX = income_sheet.GetRow(5)!.indexOf(earner)
    for (let i = 0; i < income_sheet.NumberOfRows(); i++) {
        let row = income_sheet.GetRow(i)!
        if (isNaN(Number(row[EARNER_INDEX]))) { continue }
        if (Number(row[EARNER_INDEX]) === year) { return i }
        if (Number(row[EARNER_INDEX]) > year) { break }
    }
    return -1
}

function __AITP_AddYearNotInIncomeSheet(row: string[], year_input: number, income_sheet: GoogleSheetTabs, template: DataArray) {
    const INCOME_EARNERS_INDEXES = row.map(val => val.trim().length > 0).map((val, i) => val ? i : -1).filter(val => val !== -1)

    for (let i of INCOME_EARNERS_INDEXES) {
        let year_index = __AITP_FindIncomeYear(year_input, row[i], income_sheet)
        let year_exists_for_other_users = income_sheet.IndexOfRow(row => row.indexOf(year_input) !== -1)
        
        const __GetRowIndex = function(offset: number) {
            if (year_exists_for_other_users !== -1 && year_exists_for_other_users + offset < income_sheet.NumberOfRows()) {
                return year_exists_for_other_users + offset
            }
            income_sheet.AppendRow([], true)
            return income_sheet.NumberOfRows() - 1
        }

        for (let j = 0; j < template.length && year_index === -1; j++) {
            income_sheet.WriteRowAt(__GetRowIndex(j), i, template[j])
        }
    }
}

function __AITP_AddYearNotInEarner(year_row: number, earner: string, income_sheet: GoogleSheetTabs, template: DataArray) {
    const EARNER_INDEX = income_sheet.GetRow(5)!.indexOf(earner)

    for (let i = 0; i < template.length; i++) {
        income_sheet.WriteRowAt(year_row + i, EARNER_INDEX, template[i])
    }
}

function __AITP_AddIncomeRow() {
    const UI = SpreadsheetApp.getUi()
    const YEAR_INPUT = Number(UI.prompt("What year is this income for?", UI.ButtonSet.OK).getResponseText())
    const INCOME_LABEL = UI.prompt("What is the income label?", UI.ButtonSet.OK).getResponseText()
    const INCOME_SHEET = new GoogleSheetTabs("Settings")
    const NEW_INCOME_TEMPLATE = __AITP_CreateNewIncomeTemplate(YEAR_INPUT)

    if (isNaN(YEAR_INPUT)) {
        UI.alert("Please enter a valid year")
        return
    }

    let row = __AITP_AddToIncomeEarnerRow(INCOME_LABEL, INCOME_SHEET).map(val => String(val))
    const START_OF_ROW = row.indexOf(INCOME_LABEL)
    const YEAR_ROW = INCOME_SHEET.IndexOfRow(row => row.indexOf(YEAR_INPUT) !== -1)

    if (YEAR_ROW > -1) {
        __AITP_AddYearNotInEarner(YEAR_ROW, INCOME_LABEL, INCOME_SHEET, NEW_INCOME_TEMPLATE)
    }
    
    
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


    if (YEAR_ROW === -1) {
        __AITP_AddYearNotInIncomeSheet(row, YEAR_INPUT, INCOME_SHEET, NEW_INCOME_TEMPLATE)
    }
    
    
    INCOME_SHEET.SaveToTab()
}

function AddIncomeToPlanner() {
    __AITP_AddIncomeRow()
}