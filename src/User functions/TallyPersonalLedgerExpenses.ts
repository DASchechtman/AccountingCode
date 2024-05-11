function ___TPLE_ComputeMonthCashflow(row_start: number, col_start: number, col_end: number, ledger: GoogleSheetTabs, date_header_test: Parser, inter: FormulaInterpreter): void {
    let category_index = -1
    let transaction_index = -1
    let debit_index = -1
    let debit_amt_index = -1
    let credit_index = -1
    let credit_amt_index = -1

    let income_total = 0
    let hhloc_total = 0
    let expense_total = 0

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

        if (ROW[debit_index] === "Checkings" && !isNaN(amt)) {
            income_total += amt
        }
        else if (ROW[debit_index] === "HHLOC" && !isNaN(amt)) {
            hhloc_total += amt
        }
        else if (!isNaN(amt)) {
            expense_total += amt
        }
    }

    const FIRST_ROW = ledger.GetRow(row_start)
    if (FIRST_ROW === undefined) { return }
    FIRST_ROW[col_start] = `Total Income: $${income_total}`
    FIRST_ROW[col_start+1] = `Total Expenses: $${expense_total}`
    FIRST_ROW[col_start+2] = `Total HH Loans: $${hhloc_total}`
    FIRST_ROW[col_start+3] = `Remaining Spend Power: $${(income_total -(expense_total + hhloc_total)).toFixed(2)}`
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
            ___TPLE_ComputeMonthCashflow(i+1, j, NUM_OF_LEDGER_COLS, LEDGER, DATE_HEADER, LEDGER_FORMULA_INTERPRETER)
            j += NUM_OF_LEDGER_COLS - 1
            found_headers = true
        }
    }

    LEDGER.SaveToTab()
}