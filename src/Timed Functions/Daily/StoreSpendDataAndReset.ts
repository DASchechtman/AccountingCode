function __SSDAR_StoreData(spending_tab_name: string): [string, number, number] {
    const SPENDING_TAB = new GoogleSheetTabs(spending_tab_name)

    const ALLOWANCE = SPENDING_TAB.GetRow(1)!.at(0) as number

    const START_INDEX = SPENDING_TAB.FindRowIndex(row => row.includes("Purchase Category"))
    const PURCHASE_CAT_INDEX = 0
    const PURCHASE_ATM_INDEX = 1
    const PURCHASE_DATE_INDEX = 3
    const CSV = new Array<String>()

    let total_spent = 0

    SPENDING_TAB.ForEachRow((row, i) => {
        CSV.push(row.filter(n => n !== "").join("<->"))
        const COST = Number(row[PURCHASE_ATM_INDEX])

        if (!isNaN(COST)) { total_spent += COST }
        
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
    return [CSV.join("\n"), ALLOWANCE, total_spent]
}

function __SSDAR_UpdateLoanTab(loan_tab: GoogleSheetTabs, who: string, allowance: number, total_spent: number) {
    let start_index = 0
    let new_row = new Array<string | number>()
    let found_empty_cell = false

    if (who.toLowerCase() === "dan") {
        start_index = 4
    }

    const CHARGES_INDEX = start_index + 1
    const PAYMENT_INDEX = start_index + 2

    for (let i = 0; i < loan_tab.NumberOfRows(); i++) {
        const ROW = loan_tab.GetRow(i)!
        
        if (new_row.length === 0) {
            new_row.length = ROW.length
            new_row.fill("")
        }

        if (allowance - total_spent < 0 && ROW[CHARGES_INDEX] === "") {
            ROW[CHARGES_INDEX] = allowance - total_spent
            found_empty_cell = true
            loan_tab.OverWriteRow(ROW)
            break
        }
        else if (allowance - total_spent > 0 && ROW[PAYMENT_INDEX] === "") {
            ROW[PAYMENT_INDEX] = allowance - total_spent
            found_empty_cell = true
            loan_tab.OverWriteRow(ROW)
            break
        }
    }

    if (!found_empty_cell) {
        if (allowance - total_spent < 0) {
            new_row[CHARGES_INDEX] = allowance - total_spent
        }
        else if (allowance - total_spent > 0) {
            new_row[PAYMENT_INDEX] = allowance - total_spent
        }

        loan_tab.AppendRow(new_row)
    }
}

function StoreSpendDataAndReset() {
    const STORAGE_TAB = new GoogleSheetTabs(PERSONAL_SPEND_STORAGE_TAB)
    const LOAN_TAB = new GoogleSheetTabs(LOAN_FROM_HOUSE_TAB)

    const RO_SPEND_DATA = __SSDAR_StoreData(RO_PERSONAL_SPEND_TAB)
    const DAN_SPEND_DATA = __SSDAR_StoreData(DAN_PERSONAL_SPEND_TAB)

    const STORAGE = [
        RO_SPEND_DATA[0],
        DAN_SPEND_DATA[0]
    ]

    STORAGE_TAB.AppendRow(STORAGE)
    STORAGE_TAB.SaveToTab()

    __SSDAR_UpdateLoanTab(LOAN_TAB, "Ro", RO_SPEND_DATA[1], RO_SPEND_DATA[2])
    __SSDAR_UpdateLoanTab(LOAN_TAB, "Dan", DAN_SPEND_DATA[1], DAN_SPEND_DATA[2])
    LOAN_TAB.SaveToTab()
}