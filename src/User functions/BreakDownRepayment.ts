function __BDR_Record(bdr_tab: GoogleSheetTabs, bdr_due_date_index: number, date: string, bdr_offset: number, bdr_hhloc: number, bdr_card: number, bdr_total: number, card_sum: number, hhloc_sum: number, total: number) {
    const UPDATE_ROW = bdr_tab.FindRow(row => row[bdr_due_date_index] === date)
    if (UPDATE_ROW === undefined) {
        const ROW: DataArrayEntry = []
        ROW[bdr_due_date_index] = date
        ROW[bdr_hhloc] = hhloc_sum
        ROW[bdr_card] = card_sum
        ROW[bdr_total] = total
        bdr_tab.AppendRow(ROW.filter(x => x != null))
    }
    else {
        const OFFSET = Number(UPDATE_ROW[bdr_offset])
        UPDATE_ROW[bdr_due_date_index] = date
        UPDATE_ROW[bdr_hhloc] = hhloc_sum - OFFSET + (total - (card_sum + hhloc_sum))
        UPDATE_ROW[bdr_card] = card_sum + OFFSET
        UPDATE_ROW[bdr_total] = total
        bdr_tab.OverWriteRow(UPDATE_ROW)
    }
}


function BreakDownRepayment() {
    const ONE_WEEK_LOAN_TAB = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME)
    const ONE_WEEK_BREAKDOWN_TAB = new GoogleSheetTabs(WEEKLY_PAYMENT_BREAK_DOWN_TAB_NAME)
    const ONE_WEEK_INTERPETER = new FormulaInterpreter(ONE_WEEK_LOAN_TAB)

    const PAY_WHERE_COL = ONE_WEEK_LOAN_TAB.GetHeaderIndex("Pay Where?")
    const PAY_AMT_COL = ONE_WEEK_LOAN_TAB.GetHeaderIndex("Amount")
    const PURCHASE_LOC_COL = ONE_WEEK_LOAN_TAB.GetHeaderIndex("Purchase Location")
    const TOTAL_COL = ONE_WEEK_LOAN_TAB.GetHeaderIndex("Total")
    const BREAK_DOWN_DUE_DATE = ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Payment Date")
    const BREAK_DOWN_HHLOC = ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Amount for HHLOC")
    const BREAK_DOWN_CARD = ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Amount for Card")
    const BREAK_DOWN_OFFSET = ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Offset")
    const BREAK_DOWN_TOTAL = ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Total")

    let date = ""
    let total = 0
    let card_sum = 0
    let hhloc_sum = 0

    for (let i = 1; i < ONE_WEEK_LOAN_TAB.NumberOfRows(); i++) {
        const ROW = ONE_WEEK_LOAN_TAB.GetRow(i)!
        const PURCHASE_LOC = String(ROW[PURCHASE_LOC_COL])
        const REPAY_LOC = String(ROW[PAY_WHERE_COL])
        const TOTAL = Number(ROW[TOTAL_COL])


        if (TOTAL !== 0) {
            total = TOTAL
        }

        let purchase_amt = Number(ROW[PAY_AMT_COL])

        if (typeof ROW[PAY_AMT_COL] === 'string') {
            const PARSE_RESULTS = ONE_WEEK_INTERPETER.ParseInput(ROW[PAY_AMT_COL] as string)
            if (PARSE_RESULTS == null || typeof PARSE_RESULTS !== 'number') {
                purchase_amt = 0
            }
            else {
                purchase_amt = PARSE_RESULTS
            }
        }

        if (date === "" && PURCHASE_LOC.startsWith(PURCHASE_HEADER)) {
            date = PURCHASE_LOC.split(" ")[2]
        }
        else if (PURCHASE_LOC.startsWith(PURCHASE_HEADER)) {
            __BDR_Record(ONE_WEEK_BREAKDOWN_TAB, BREAK_DOWN_DUE_DATE, date, BREAK_DOWN_OFFSET, BREAK_DOWN_HHLOC, BREAK_DOWN_CARD, BREAK_DOWN_TOTAL, card_sum, hhloc_sum, total)
            date = PURCHASE_LOC.split(" ")[2]
            card_sum = 0
            hhloc_sum = 0
        }

        if (REPAY_LOC === "Card") {
            card_sum += purchase_amt
        }
        else if (REPAY_LOC === "") {
            card_sum += purchase_amt
            ROW[PAY_WHERE_COL] = PURCHASE_LOC.startsWith(PURCHASE_HEADER) ? "" : "Card"
        }
        else if (REPAY_LOC === "HHLOC") {
            hhloc_sum += purchase_amt
        }

        ONE_WEEK_LOAN_TAB.WriteRow(i, ROW)
    }

    __BDR_Record(ONE_WEEK_BREAKDOWN_TAB, BREAK_DOWN_DUE_DATE, date, BREAK_DOWN_OFFSET, BREAK_DOWN_HHLOC, BREAK_DOWN_CARD, BREAK_DOWN_TOTAL, card_sum, hhloc_sum, total)

    ONE_WEEK_LOAN_TAB.SaveToTab()
    ONE_WEEK_BREAKDOWN_TAB.SaveToTab()
}