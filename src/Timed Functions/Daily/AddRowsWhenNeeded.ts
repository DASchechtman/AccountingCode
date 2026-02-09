
function AddRowsWhenNeeded() {
    const SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    const PURCHASE_LOCATION_INDEX = SHEET.GetHeaderIndex("Purchase Location")
    const DUE_DATE_INDEX = SHEET.GetHeaderIndex("Due Date")
    const NEW_ROWS = new Array<Array<string>>()
    let today = new Date()

    if (today.getDate() !== 28) { return }

    while (today.getDay() !== 3) {
        today.setDate(today.getDate() + 1)
    }

    while (__Util_DateInCurrentPayPeriod(today)) {
        const LAST_DATE = __Util_CreateDateString(today)
        const HEADER = `${PURCHASE_HEADER} ${LAST_DATE}`
        const HEADER_ARR: string[] = []
        const DUE_DATE_ARR: string[] = []
        HEADER_ARR[PURCHASE_LOCATION_INDEX] = HEADER
        DUE_DATE_ARR[DUE_DATE_INDEX] = LAST_DATE
        NEW_ROWS.push(HEADER_ARR)
        NEW_ROWS.push(DUE_DATE_ARR)
        today.setDate(today.getDate() + 7)
    }

    for(let new_row of NEW_ROWS) {
        SHEET.AppendRow(new_row)
    }

    SHEET.SaveToTab()


}