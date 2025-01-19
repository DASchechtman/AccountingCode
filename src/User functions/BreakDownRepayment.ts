class __BDR_BreakDownExpenses {
    private readonly ONE_WEEK_LOAN_TAB = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    private readonly ONE_WEEK_BREAKDOWN_TAB = new GoogleSheetTabs(WEEKLY_PAYMENT_BREAK_DOWN_TAB_NAME)
    private readonly ONE_WEEK_INTERPRETER = new FormulaInterpreter(this.ONE_WEEK_LOAN_TAB)
    private readonly ONE_WEEK_BREAKDOWN_INTERPRETER = new FormulaInterpreter(this.ONE_WEEK_BREAKDOWN_TAB)

    private readonly PAY_WHERE_COL = this.ONE_WEEK_LOAN_TAB.GetHeaderIndex("Pay Where?")
    private readonly PAY_AMT_COL = this.ONE_WEEK_LOAN_TAB.GetHeaderIndex("Amount")
    private readonly PURCHASE_LOC_COL = this.ONE_WEEK_LOAN_TAB.GetHeaderIndex("Purchase Location")
    private readonly TOTAL_COL = this.ONE_WEEK_LOAN_TAB.GetHeaderIndex("Total")

    private readonly BREAK_DOWN_DUE_DATE = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Payment Date")
    private readonly BREAK_DOWN_HHLOC = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Amount for HHLOC")
    private readonly BREAK_DOWN_CARD = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Amount for Card")
    private readonly BREAK_DOWN_OFFSET = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Offset")
    private readonly BREAK_DOWN_TOTAL = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Total")
    private readonly BREAK_DOWN_CHANGE_ADD_TO_CARD = this.ONE_WEEK_BREAKDOWN_TAB.GetHeaderIndex("Change added to Card?")

    private Record(date: string, hhloc_sum: number, card_sum: number, total: number) {
        const UPDATE_ROW = this.ONE_WEEK_BREAKDOWN_TAB.FindRow(row => row[this.BREAK_DOWN_DUE_DATE] === date)
        if (UPDATE_ROW === undefined) {
            const ROW = this.ONE_WEEK_BREAKDOWN_TAB.FindRow(row => row[this.BREAK_DOWN_DUE_DATE] === "")
            if (ROW === undefined) { return }
            ROW[this.BREAK_DOWN_DUE_DATE] = date
            ROW[this.BREAK_DOWN_HHLOC] = hhloc_sum
            ROW[this.BREAK_DOWN_CARD] = card_sum
            ROW[this.BREAK_DOWN_TOTAL] = total
            this.ONE_WEEK_BREAKDOWN_TAB.OverWriteRow(ROW)
        }
        else {
            let [did_parse, data] = this.ONE_WEEK_BREAKDOWN_INTERPRETER.AttemptToParseInput(UPDATE_ROW[this.BREAK_DOWN_OFFSET])
            let offset = 0

            if (did_parse && typeof data === 'number') {
                offset = data
            }

            UPDATE_ROW[this.BREAK_DOWN_DUE_DATE] = date
            UPDATE_ROW[this.BREAK_DOWN_HHLOC] = hhloc_sum - offset + (total - (card_sum + hhloc_sum))
            UPDATE_ROW[this.BREAK_DOWN_CARD] = card_sum + offset
            UPDATE_ROW[this.BREAK_DOWN_TOTAL] = total

            if (UPDATE_ROW[this.BREAK_DOWN_CHANGE_ADD_TO_CARD]) {
                (UPDATE_ROW[this.BREAK_DOWN_CARD] as number) += UPDATE_ROW[this.BREAK_DOWN_HHLOC] as number
                UPDATE_ROW[this.BREAK_DOWN_HHLOC] = 0
            }

            this.ONE_WEEK_BREAKDOWN_TAB.OverWriteRow(UPDATE_ROW)
        }
    }

    public BreakDownRepayment() {
        let date = ""
        let total = 0
        let card_sum = 0
        let hhloc_sum = 0

        this.ONE_WEEK_LOAN_TAB.ForEachRow((row, i) => {
            if (i === 0) { return 'continue' }

            const ROW = row
            const PURCHASE_LOC = String(ROW[this.PURCHASE_LOC_COL])
            const REPAY_LOC = String(ROW[this.PAY_WHERE_COL])
            const TOTAL = Number(ROW[this.TOTAL_COL])

            if (TOTAL !== 0) {
                total = TOTAL
            }

            let purchase_amt = Number(ROW[this.PAY_AMT_COL])

            if (isNaN(purchase_amt)) {
                const PARSE_RESULTS = this.ONE_WEEK_INTERPRETER.ParseInput(ROW[this.PAY_AMT_COL] as string)
                if (PARSE_RESULTS == null || typeof PARSE_RESULTS !== 'number') {
                    purchase_amt = 0
                }
                else {
                    purchase_amt = PARSE_RESULTS
                }
            }

            if (date === "" && PURCHASE_LOC.startsWith(PURCHASE_HEADER)) {
                date = PURCHASE_LOC.split(" ")[2]
                return 'continue'
            }
            else if (PURCHASE_LOC.startsWith(PURCHASE_HEADER)) {
                this.Record(date, hhloc_sum, card_sum, total)
                date = PURCHASE_LOC.split(" ")[2]
                card_sum = 0
                hhloc_sum = 0
                total = 0
                return 'continue'
            }

            if (REPAY_LOC === "Card") {
                card_sum += purchase_amt
            }
            else if (REPAY_LOC === "") {
                card_sum += purchase_amt
                ROW[this.PAY_WHERE_COL] = PURCHASE_LOC.startsWith(PURCHASE_HEADER) ? "" : "Card"
            }
            else if (REPAY_LOC === "HHLOC") {
                hhloc_sum += purchase_amt
            }

            return ROW.map(cell => {
                if (typeof cell === 'string') { return cell.trim() }
                return cell
            })
        })

        this.Record(date, hhloc_sum, card_sum, total)

        this.ONE_WEEK_LOAN_TAB.SaveToTab()
        this.ONE_WEEK_BREAKDOWN_TAB.SaveToTab()
    }

}


function BreakDownRepayment() {
    new __BDR_BreakDownExpenses().BreakDownRepayment()
}