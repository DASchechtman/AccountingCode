
function __ARWN_AddNextMonth(sheet: GoogleSheetTabs, date_str: string) {
    const DATE = new Date(date_str)

    for (let i = 0; i < 5; i++) {
        DATE.setDate(DATE.getDate() + 7)
        const NEW_DATE = __Util_CreateDateString(DATE)
        sheet.AppendRow(["", "", NEW_DATE])
        sheet.AppendRow(["", "", "", "", NEW_DATE])
    }

    sheet.SaveToTab()
}


function AddRowsWhenNeeded() {
    const SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    const LAST_CACHED_DATE = __Cache_Utils_QueryLastWeek('date')

    const PURCHASE_LOC_INDEX = SHEET.GetHeaderIndex('Purchase Location')

    for (let i = SHEET.NumberOfRows() - 1; i >= 0; i--) {
        const ROW = SHEET.GetRow(i)!
        const HEADER = String(ROW[PURCHASE_LOC_INDEX])
        const IS_HEADER = HEADER.startsWith(PURCHASE_HEADER)

        if (IS_HEADER) {
            const HEADER_DATE = __Util_GetDateFromDateHeader(HEADER)
            if (HEADER_DATE === LAST_CACHED_DATE) { __ARWN_AddNextMonth(SHEET, LAST_CACHED_DATE) }
            break
        }
    }
}