const PAY_AMT = [126]

function __WCCOE_GetSumFormula(start_range: string, end_range: string) {
    return `=SUM(ARRAYFORMULA(ROUNDUP(${start_range}:${end_range})))`
}

function __WCCOE_SetLastRowToHaveSum(sheet: GoogleSheetTabs, start_range: string, amt_col_index: number, total_col_index: number) {
    const LAST_ROW = sheet.GetRow(sheet.NumberOfRows() - 1)!
    const END_RANGE = `${__Util_IndexToColLetter(amt_col_index)}${sheet.NumberOfRows()}`
    LAST_ROW[total_col_index] = __WCCOE_GetSumFormula(start_range, END_RANGE)
    sheet.OverWriteRow(LAST_ROW)
    return END_RANGE
}

function __WCCOE_GetWeeklyCharges(sheet: GoogleSheetTabs, start: number, purchase_col: number, amt_col: number) {
    let cells = new Array<string>()
    for (let i = start; i < sheet.NumberOfRows(); i++) {
        const ROW = sheet.GetRow(i)!
        const PURCHASES = ROW[purchase_col].toString()

        if (PURCHASES.startsWith("Purchases for")) {
            break
        }

        cells.push(`${__Util_IndexToColLetter(amt_col)}${i+1}`)
    }

    return `${cells[0]}:${cells.at(-1)}`
}


function WeeklyCreditChargesOnEdit() {
    const WEEKLY_CHARGES_SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)

    const TOTAL_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Total")
    const AMT_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Amount")
    const PURCHASE_LOC_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Location")
    const DUE_DATE_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Due Date")
    const PURCHASE_DATE_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Date")
    const TIPS_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Tips")
    const MONEY_LEFT_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Money Left")

    let start_range = ''
    let due_date = ''
    let tip_cell = ''
    let total_charge_cells = new Array<Array<string>>()
    let money_left = 0
    let in_month = 0

    const InPayPeriod = (check: string) => {
        const TODAY = new Date()

        let year = TODAY.getFullYear()
        let month = TODAY.getMonth() - 1

        if (month < 0) {
            month = 11
            year -= 1
        }

        const FROM_DATE = new Date(`${month + 1}/28/${year}`)
        const TO_DATE = new Date(`${TODAY.getMonth() + 1}/27/${TODAY.getFullYear()}`)
        const CHECK_DATE = new Date(check)

        return CHECK_DATE >= FROM_DATE && CHECK_DATE <= TO_DATE
    }


    for (let i = 1; i < WEEKLY_CHARGES_SHEET.NumberOfRows(); i++) {
        const ROW = WEEKLY_CHARGES_SHEET.GetRow(i)!
        const PURCHASES = ROW[PURCHASE_LOC_COL_INDEX].toString()

        if (PURCHASES.startsWith("Purchases for")) {
            due_date = __Util_GetDateFromDateHeader(PURCHASES)
            const LAST_ROW = WEEKLY_CHARGES_SHEET.GetRow(i - 1)!
            const IN_PAY_PERIOD = InPayPeriod(due_date)
            const end_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i}`
            const tip_range = `${__Util_IndexToColLetter(TIPS_COL_INDEX)}${i+1}`
            

            if (IN_PAY_PERIOD) {
                in_month++
            }
            else if (in_month > 0 && !IN_PAY_PERIOD) {
                const MONTHLY_TOTALS = total_charge_cells.map(cell => cell[1]).join(',')
                const MONTHLY_TIPS = total_charge_cells.map(cell => cell[0]).join(',')
                LAST_ROW[MONEY_LEFT_COL_INDEX] = `=${money_left} + SUM(${MONTHLY_TIPS}) - SUM(${MONTHLY_TOTALS})`
                in_month = 0
            }

            if (in_month > 0) { 
                money_left += PAY_AMT.at(-1)!
                total_charge_cells.push([tip_range, __WCCOE_GetWeeklyCharges(WEEKLY_CHARGES_SHEET, i+1, PURCHASE_LOC_COL_INDEX, AMT_COL_INDEX)]) 
            }

            if (start_range === '') {
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
            }
            else {
                const TOTAL = __WCCOE_GetSumFormula(start_range, end_range)
                LAST_ROW[TOTAL_COL_INDEX] = TOTAL
                
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
            }

            WEEKLY_CHARGES_SHEET.OverWriteRow(LAST_ROW)
        }
        else {
            ROW[TOTAL_COL_INDEX] = ''
            ROW[MONEY_LEFT_COL_INDEX] = ''

            if (ROW[DUE_DATE_COL_INDEX] === '') {
                ROW[DUE_DATE_COL_INDEX] = due_date
            }
    
            if (ROW[PURCHASE_DATE_COL_INDEX] === '') {
                ROW[PURCHASE_DATE_COL_INDEX] = __Util_CreateDateString(new Date())
            }
        }

        WEEKLY_CHARGES_SHEET.OverWriteRow(ROW)
    }

    __WCCOE_SetLastRowToHaveSum(WEEKLY_CHARGES_SHEET, start_range, AMT_COL_INDEX, TOTAL_COL_INDEX)

    WEEKLY_CHARGES_SHEET.SaveToTab()
}