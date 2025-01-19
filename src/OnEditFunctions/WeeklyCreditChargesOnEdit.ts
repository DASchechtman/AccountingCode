

function WeeklyCreditChargesOnEdit() {
    const WEEKLY_CHARGES_SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    const TOTAL_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Total")
    const AMT_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Amount")
    const PURCHASE_LOC_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Location")

    let start_range = ''


    for (let i = 1; i < WEEKLY_CHARGES_SHEET.NumberOfRows(); i++) {
        const ROW = WEEKLY_CHARGES_SHEET.GetRow(i)!
        const PURCHASES = ROW[PURCHASE_LOC_COL_INDEX].toString()

        if (PURCHASES.startsWith("Purchases for")) {
            if (start_range === '') {
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
            }
            else {
                const end_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i}`
                const RANGE = `${start_range}:${end_range}`
                const TOTAL = `=SUM(ARRAYFORMULA(ROUNDUP(${RANGE})))`
                start_range = `${__Util_IndexToColLetter(AMT_COL_INDEX)}${i + 2}`
                const LAST_ROW = WEEKLY_CHARGES_SHEET.GetRow(i - 1)!
                LAST_ROW[TOTAL_COL_INDEX] = TOTAL
                WEEKLY_CHARGES_SHEET.OverWriteRow(LAST_ROW)
            }
        }
        else {
            ROW[TOTAL_COL_INDEX] = ''
            WEEKLY_CHARGES_SHEET.OverWriteRow(ROW)
        }
    }

    WEEKLY_CHARGES_SHEET.SaveToTab()
}