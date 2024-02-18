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

function __AITP_CreateConditionalFormatRuleList(col_index: number, row_index: number) {
    let too_big = SpreadsheetApp.newConditionalFormatRule()
    let too_small = SpreadsheetApp.newConditionalFormatRule()
    let just_right = SpreadsheetApp.newConditionalFormatRule()

    const DARK_RED_3 = "#3f0000"
    const LIGHT_RED_3 = "#f3d3d9"
    const DARK_GREEN_3 = "#003f00"
    const LIGHT_GREEN_3 = "#d3f3d9"
    const DARK_ORANGE_3 = "#673800"
    const LIGHT_ORANGE_3 = "#ffd19a"

    const INDEX = __Util_IndexToColLetter
    const COL_1 = `$${INDEX(col_index)}${row_index}`
    const COL_2 = `$${INDEX(col_index + 1)}${row_index}`
    const COL_3 = `$${INDEX(col_index + 2)}${row_index}`

    too_big.whenFormulaSatisfied(`=AND(ISNUMBER(${COL_1}), ${COL_1}<${COL_2}+${COL_3})`)
        .setBackground(DARK_RED_3)
        .setFontColor(LIGHT_RED_3)
    too_small.whenFormulaSatisfied(`=AND(ISNUMBER(${COL_1}), ${COL_1}>${COL_2}+${COL_3})`)
        .setBackground(DARK_ORANGE_3)
        .setFontColor(LIGHT_ORANGE_3)
    just_right.whenFormulaSatisfied(`=AND(ISNUMBER(${COL_1}), ${COL_1}=${COL_2}+${COL_3})`)
        .setBackground(DARK_GREEN_3)
        .setFontColor(LIGHT_GREEN_3)
    
    return [too_big, too_small, just_right]
}

function __AITP_AddConditionalFormatToEarner() {
    const SHEET = new GoogleSheetTabs("Settings")
    const RULES: GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder[] = []
    const START_ROW = SHEET.IndexOfRow(row => row.some(val => !isNaN(Number(val))))

    for (let i = START_ROW; i < SHEET.NumberOfRows(); i++) {
        if (i === -1) { break }

        let row = SHEET.GetRow(i)!
        for (let j = 0; j < row.length; j++) {
            if (row[j] !== "Take Home Pay") { continue }
            const SHEET_ROW = i + 2
            const FORMAT_RULES = __AITP_CreateConditionalFormatRuleList(j, SHEET_ROW)
            const RANGE_STR = `${__Util_IndexToColLetter(j)}${SHEET_ROW}:${__Util_IndexToColLetter(j)}${SHEET_ROW+11}`
            const RANGE = SHEET.GetTab().getRange(RANGE_STR)
            RULES.push(...FORMAT_RULES.map(rule => rule.setRanges([RANGE])))
        }
    }

    SHEET.GetTab().setConditionalFormatRules(RULES.map(rule => rule.build()))
}

function __AITP_BoldAndCenterYearRows() {
    const SHEET = new GoogleSheetTabs("Settings")
    const FIRST_COL = SHEET.GetColByIndex(0)!
    const BOLD_FONT = SpreadsheetApp.newTextStyle()
        .setBold(true)
        .build()

    for (let i = 0; i < FIRST_COL.length; i++) {
        if (isNaN(Number(FIRST_COL[i]))) { continue }
        SHEET.GetRowRange(i)
            ?.setTextStyle(BOLD_FONT)
            ?.setHorizontalAlignment("center")
    }
}

function __AITP_AddDataValidationToEarner() {
    const SHEET = new GoogleSheetTabs("Settings")
    const VALIDATION = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Weekly", "Bi-Weekly", "Semi-Monthly", "Monthly"])
        .build()

    for (let i = 6; i < SHEET.NumberOfRows(); i+=13) {
        let row = SHEET.GetRow(i)!
        for (let j = 0; j < row.length; j++) {
            let cell = row[j]
            if (cell !== "Frequency") { continue }
            const COL = __Util_IndexToColLetter(j)
            const RANGE_STR = `${COL}${i+2}:${COL}${i+13}`
            SHEET.GetTab().getRange(RANGE_STR).setDataValidation(VALIDATION)
        }
    }
}

function AddIncomeToPlanner() {
    __AITP_AddIncomeRow()
    __AITP_AddConditionalFormatToEarner()
    __AITP_BoldAndCenterYearRows()
    __AITP_AddDataValidationToEarner()
}