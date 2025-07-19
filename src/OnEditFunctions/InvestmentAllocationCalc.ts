
function InvestmentAllocationCalcOnEdit() {
    const SHEET = new GoogleSheetTabs(INVESTMENT_ALLOC_TAB, "H2:M")
    let i = 0

    const TICKER_INDEX = 0
    const TICKER_PAY_AMT_INDEX = 1
    const TICKER_PERCENT_INDEX = 2
    const TOTAL_PAY_INDEX = 5

    let row = SHEET.GetRow(i)!
    const TOTAL_PAY = Number(row[TOTAL_PAY_INDEX])

    do {
        row = SHEET.GetRow(i)!

        if (row[TICKER_INDEX] === "Total") {
            break
        }
        else if (row[TICKER_PAY_AMT_INDEX] === "" && row[TICKER_PERCENT_INDEX] === "") {
            continue
        }
        else if (row[TICKER_PAY_AMT_INDEX] === "") {
            row[TICKER_PAY_AMT_INDEX] = TOTAL_PAY * Number(row[TICKER_PERCENT_INDEX])
        }
        else if (row[TICKER_PERCENT_INDEX] === "") {
            row[TICKER_PERCENT_INDEX] = Number(row[TICKER_PAY_AMT_INDEX]) / TOTAL_PAY
        }
        else {
            row[TICKER_PAY_AMT_INDEX] = TOTAL_PAY * Number(row[TICKER_PERCENT_INDEX])
        }

        SHEET.OverWriteRow(row)
    } while(++i)

    SHEET.SaveToTab()
}