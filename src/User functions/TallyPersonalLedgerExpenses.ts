function __TPLE_IsBorrowed(ledger: GoogleSheetTabs, debit_row_index: number, credit_row_index: number, amt_row_index: number, cur_row: number) {
    if (cur_row + 1 >= ledger.NumberOfRows()) { return false }
    const CUR_ROW = ledger.GetRow(cur_row)!
    const NEXT_ROW = ledger.GetRow(cur_row + 1)!

    let is_borrowed = false

    if (CUR_ROW[debit_row_index] === "Checkings" && CUR_ROW[credit_row_index] === "HHLOC") {
        const NOT_GOING_TO_OWNED_ACCOUNT = NEXT_ROW[debit_row_index] !== "Checkings" && NEXT_ROW[debit_row_index] !== "HHLOC"
        if (NOT_GOING_TO_OWNED_ACCOUNT && NEXT_ROW[credit_row_index] === "Checkings") {
            is_borrowed = CUR_ROW[amt_row_index] === NEXT_ROW[amt_row_index]
        }
    }

    return is_borrowed
}

function ___TPLE_ComputeMonthCashflow(row_start: number, col_start: number, ledger: GoogleSheetTabs, date_header_test: Parser, inter: FormulaInterpreter): void {
    let category_index = -1
    let transaction_index = -1
    let debit_index = -1
    let debit_amt_index = -1
    let credit_index = -1
    let credit_amt_index = -1

    let income_total = 0
    let hhloc_total = 0
    let expense_total = 0
    let borrowed_total = 0

    for(let i = row_start; i < ledger.NumberOfRows(); i++) {
        const ROW = ledger.GetRow(i)!

        if (!date_header_test.Run(String(ROW[col_start])).is_error) { break }
        if (ROW[col_start] === "Category") {
            category_index = col_start
            transaction_index = col_start + 1
            debit_index = col_start + 2
            debit_amt_index = col_start + 3
            credit_index = col_start + 4
            credit_amt_index = col_start + 5
            continue
        }
        let amt = Number(ROW[debit_amt_index])

        if (isNaN(amt)) {
            let parse_attempt = inter.AttemptToParseInput(ROW[debit_amt_index])
            if (parse_attempt[0] && typeof parse_attempt[1] === "number") {
                amt = parse_attempt[1]
            }
        }

        const IS_EXPENSE = (
            String(ROW[credit_index]).toLowerCase().trim() === "income"
            && String(ROW[debit_index]).toLowerCase().trim() !== "hhloc"
            && String(ROW[debit_index]).toLowerCase().trim() !== "checkings"
            && !isNaN(amt)
        )

        if (__TPLE_IsBorrowed(ledger, debit_index, credit_index, debit_amt_index, i)) {
            borrowed_total += amt
            income_total += amt
            i++
        }
        else if (IS_EXPENSE) {
            expense_total += amt
        }
        else if (/^(G|g)ift from.+$/.test(String(ROW[category_index]))) {
            const GIFT_FROM = String(ROW[category_index]).split("from")[1].trim()
            const GIFT_ROW = ledger.FindRow(row => row[credit_index] === GIFT_FROM)
            if (GIFT_ROW) {
                const GIFT_AMT = Number(GIFT_ROW[credit_amt_index])
                income_total -= !isNaN(GIFT_AMT) ? GIFT_AMT : 0
            }
        }
        else if (ROW[debit_index] === "Checkings" && !isNaN(amt)) {
            income_total += amt
        }
        else if (ROW[debit_index] === "HHLOC" && !isNaN(amt)) {
            hhloc_total += amt
        }
    }

    const FIRST_ROW = ledger.GetRow(row_start)
    if (FIRST_ROW === undefined) { return }
    FIRST_ROW[col_start] = `Total Income: $${income_total.toFixed(2)}`
    FIRST_ROW[col_start+1] = `Total Expenses: $${expense_total.toFixed(2)}`
    FIRST_ROW[col_start+2] = `Total HH Repayments: $${hhloc_total.toFixed(2)}`
    FIRST_ROW[col_start+3] = `Total Borrowed: $${borrowed_total.toFixed(2)}`
    FIRST_ROW[col_start+4] = `Remaining Spend Power: $${(income_total - (expense_total + hhloc_total + borrowed_total)).toFixed(2)}`
    ledger.OverWriteRow(FIRST_ROW)
}

function __TPLE_CreateDateHeaderParser() {
    const Months = __SFI_Choice(
        __SFI_Str("Jan"),
        __SFI_Str("Feb"),
        __SFI_Str("Mar"),
        __SFI_Str("Apr"),
        __SFI_Str("May"),
        __SFI_Str("Jun"),
        __SFI_Str("Jul"),
        __SFI_Str("Aug"),
        __SFI_Str("Sep"),
        __SFI_Str("Oct"),
        __SFI_Str("Nov"),
        __SFI_Str("Dec")
    )

    const DateSeg = __SFI_Regex(/\d{1,2}/)
    const Date = __SFI_SeqOf(DateSeg, __SFI_Str("/"), DateSeg, __SFI_Str("/"), DateSeg, DateSeg)
    const DateRange = __SFI_SeqOf(__SFI_Str("("), Date, __SFI_Str(" - "), Date, __SFI_Str(")"))
    return new Parser(__SFI_SeqOf(Months, __SFI_Str(" "), DateRange))
}

function TallyPersonalLedgerExpenses() {
    const LEDGER = new GoogleSheetTabs(PERSONAL_LEDGER_TAB_NAME)
    const LEDGER_FORMULA_INTERPRETER = new FormulaInterpreter(LEDGER)
    const DATE_HEADER = __TPLE_CreateDateHeaderParser()
    const NUM_OF_LEDGER_COLS = 6

    for (let i = 0; i < LEDGER.NumberOfRows(); i++) {
        const ROW = LEDGER.GetRow(i)!
        let found_headers = false

        for (let j = 0; j < ROW.length; j++) {
            let cell = ROW[j]
            if (DATE_HEADER.Run(String(cell)).is_error) { continue }
            ___TPLE_ComputeMonthCashflow(i+1, j, LEDGER, DATE_HEADER, LEDGER_FORMULA_INTERPRETER)
            j += NUM_OF_LEDGER_COLS - 1
            found_headers = true
        }
    }

    LEDGER.SaveToTab()
}