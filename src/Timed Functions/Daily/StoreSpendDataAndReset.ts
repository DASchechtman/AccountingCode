function __SSDAR_StoreData(spending_tab_name: string) {
    const SPENDING_TAB = new GoogleSheetTabs(spending_tab_name)

    const START_INDEX = SPENDING_TAB.FindRowIndex(row => row.includes("Purchase Category"))
    const PURCHASE_CAT_INDEX = 0
    const PURCHASE_DATE_INDEX = 3
    const CSV = new Array<String>()

    SPENDING_TAB.ForEachRow((row, i) => {
        CSV.push(row.filter(n => n !== "").join("<->"))
        if (i > START_INDEX && row[PURCHASE_CAT_INDEX] !== "Subscriptions" && row[PURCHASE_CAT_INDEX] !== "") {
            let new_row = ["", "", "", ""]
            return new_row
        }
        else if (i > START_INDEX && row[PURCHASE_CAT_INDEX] === "Subscriptions") {
            row[PURCHASE_DATE_INDEX] = __Util_CreateDateString(new Date())
            return row
        } 
    }, START_INDEX)

    SPENDING_TAB.SaveToTab()
    return CSV.join("\n")
}

function StoreSpendDataAndReset() {
    const STORAGE_TAB = new GoogleSheetTabs(PERSONAL_SPEND_STORAGE_TAB)
    const STORAGE = [__SSDAR_StoreData(RO_PERSONAL_SPEND_TAB),
                     __SSDAR_StoreData(DAN_PERSONAL_SPEND_TAB)]

    STORAGE_TAB.AppendRow(STORAGE)
    STORAGE_TAB.SaveToTab()
}