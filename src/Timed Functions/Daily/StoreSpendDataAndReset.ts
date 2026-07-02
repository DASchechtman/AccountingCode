function __SSDAR_StoreData(spending_tab_name: string): [string, () => number] {
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
            let blank_row = new Array<string>(row.length).fill("")
            return blank_row
        }
        else if (i > START_INDEX && row[PURCHASE_CAT_INDEX] === "Subscriptions") {
            row[PURCHASE_DATE_INDEX] = __Util_CreateDateString(new Date())
            return row
        } 
    }, START_INDEX)

    SPENDING_TAB.SaveToTab()
    return [CSV.join("\n"), () => ALLOWANCE - total_spent]
}

function __SSDAR_UpdateLoanTab(loan_tab: GoogleSheetTabs, who: string, BudgetLeft: () => number) {
    const START_ROW = loan_tab.GetRow(0)!

    let start_index = START_ROW.findIndex(n => n.toString().toLowerCase().includes(who.toLowerCase()))
    let new_row = new Array<string | number>()
    let found_empty_cell = false

    const REMAINING_BUDGET = BudgetLeft()

    if (start_index < 0) { return }

    const CHARGES_INDEX = start_index + 1
    const PAYMENT_INDEX = start_index + 2

    for (let i = 0; i < loan_tab.NumberOfRows(); i++) {
        const ROW = loan_tab.GetRow(i)!
        
        if (new_row.length === 0) {
            new_row.length = ROW.length
            new_row.fill("")
        }

        if (REMAINING_BUDGET < 0 && ROW[CHARGES_INDEX] === "") {
            ROW[CHARGES_INDEX] = REMAINING_BUDGET
            found_empty_cell = true
            loan_tab.OverWriteRow(ROW)
            break
        }
        else if (REMAINING_BUDGET > 0 && ROW[PAYMENT_INDEX] === "") {
            ROW[PAYMENT_INDEX] = REMAINING_BUDGET
            found_empty_cell = true
            loan_tab.OverWriteRow(ROW)
            break
        }
    }

    if (!found_empty_cell && REMAINING_BUDGET !== 0) {
        if (REMAINING_BUDGET < 0) {
            new_row[CHARGES_INDEX] = REMAINING_BUDGET
        }
        else if (REMAINING_BUDGET > 0) {
            new_row[PAYMENT_INDEX] = REMAINING_BUDGET
        }

        loan_tab.AppendRow(new_row)
    }
}

function StoreSpendDataAndReset() {
    const STORAGE_TAB = new GoogleSheetTabs(PERSONAL_SPEND_STORAGE_TAB)
    const LOAN_TAB = new GoogleSheetTabs(LOAN_FROM_HOUSE_TAB)

    const RO_SPEND_DATA = __SSDAR_StoreData(RO_PERSONAL_SPEND_TAB)
    const DAN_SPEND_DATA = __SSDAR_StoreData(DAN_PERSONAL_SPEND_TAB)

    const DAN_CSV = DAN_SPEND_DATA[0]
    const RO_CSV = RO_SPEND_DATA[0]

    const DanBudgetLeft = DAN_SPEND_DATA[1]
    const RoBudgetLeft = RO_SPEND_DATA[1]

    const STORAGE = [RO_CSV, DAN_CSV]

    STORAGE_TAB.AppendRow(STORAGE)
    STORAGE_TAB.SaveToTab()

    __SSDAR_UpdateLoanTab(LOAN_TAB, "Ro", RoBudgetLeft)
    __SSDAR_UpdateLoanTab(LOAN_TAB, "Dan", DanBudgetLeft)
    LOAN_TAB.SaveToTab()
}