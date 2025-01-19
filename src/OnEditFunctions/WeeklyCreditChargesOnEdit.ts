
function __WCCOE_GetSumFormula(start_range: string, end_range: string) {
    return `=SUM(ARRAYFORMULA(ROUNDUP(${start_range}:${end_range})))`
}


function WeeklyCreditChargesOnEdit() {
    const WEEKLY_CHARGES_SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)

    const TOTAL_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Total")
    const AMT_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Amount")
    const PURCHASE_LOC_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Location")
    const DUE_DATE_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Due Date")
    const PURCHASE_DATE_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Date")

    let start_range = ''
    let due_date = ''


    for (let i = 1; i < WEEKLY_CHARGES_SHEET.NumberOfRows(); i++) {
        const ROW = WEEKLY_CHARGES_SHEET.GetRow(i)!
        const PURCHASES = ROW[PURCHASE_LOC_COL_INDEX].toString()

        if (PURCHASES.startsWith("Purchases for")) {
            due_date = __Util_GetDateFromDateHeader(PURCHASES)

            if (start_range === '') {
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
            }
            else {
                const end_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i}`
                const TOTAL = __WCCOE_GetSumFormula(start_range, end_range)
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
                const LAST_ROW = WEEKLY_CHARGES_SHEET.GetRow(i - 1)!
                LAST_ROW[TOTAL_COL_INDEX] = TOTAL
                WEEKLY_CHARGES_SHEET.OverWriteRow(LAST_ROW)
            }
        }
        else {
            ROW[TOTAL_COL_INDEX] = ''

            if (ROW[DUE_DATE_COL_INDEX] === '') {
                ROW[DUE_DATE_COL_INDEX] = due_date
            }
    
            if (ROW[PURCHASE_DATE_COL_INDEX] === '') {
                ROW[PURCHASE_DATE_COL_INDEX] = __Util_CreateDateString(new Date())
            }
        }

        

        WEEKLY_CHARGES_SHEET.OverWriteRow(ROW)
    }

    const LAST_ROW = WEEKLY_CHARGES_SHEET.GetRow(WEEKLY_CHARGES_SHEET.NumberOfRows() - 1)!
    const END_RANGE = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${WEEKLY_CHARGES_SHEET.NumberOfRows()}`
    LAST_ROW[TOTAL_COL_INDEX] = __WCCOE_GetSumFormula(start_range, END_RANGE)
    WEEKLY_CHARGES_SHEET.OverWriteRow(LAST_ROW)



    WEEKLY_CHARGES_SHEET.SaveToTab()
}