

function GroupWeeklyCharges() {
    const QUERY = __Cache_Utils_QueryAllAttrs()

    if (!QUERY) { return }

    for (let query of QUERY) {
        const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEEKLY_CREDIT_CHARGES_TAB_NAME)
        const COLOR_RANGE = SHEET?.getRange(query.range_str)
        const GROUP_RANGE = SHEET?.getRange(query.grouping_range)
        const COLOR = query.range_shade
        let group

        try {
            group = SHEET!.getRowGroup(GROUP_RANGE!.getRowIndex(), 1)
            group?.remove()
            GROUP_RANGE?.shiftRowGroupDepth(1)
            group = SHEET!.getRowGroup(GROUP_RANGE!.getRowIndex(), 1)
        } catch {
            GROUP_RANGE?.shiftRowGroupDepth(1)
            group = SHEET!.getRowGroup(GROUP_RANGE!.getRowIndex(), 1)
        }

        if (new Date() > new Date(query.date) && !group?.isCollapsed()) {
            GROUP_RANGE?.collapseGroups()
            COLOR_RANGE?.setBackground(COLOR)
        }
    }

}