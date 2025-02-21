
function __MSU_FindOpenStat(html: string) {
    const DATA_KEY = "data-value="
    let i = html.indexOf('title="Open"')
    i = html.indexOf(DATA_KEY, i)

    let digits = new Array<string>()
    for (let j = i + DATA_KEY.length + 1; j < html.length; j++) {
        const DIGIT = Number(html[j])
        if (isNaN(DIGIT) && html[j] !== '.') {
            break
        }
        else {
            digits.push(html[j])
        }
    }

    return Number(digits.join(''))
}

function __MSU_FindCurrentPPSStat(html: string) {
    const DATA_KEY = 'data-testid="qsp-price"'
    let i = html.indexOf(DATA_KEY)
    let found_digits = false
    let digits = new Array<string>()

    for (let j = i + DATA_KEY.length; j < html.length; j++) {
        const DIGIT = Number(html[j])
        if (found_digits && (!isNaN(DIGIT) || html[j] === '.')) {
            digits.push(html[j])
        }
        else if (!isNaN(DIGIT) && !found_digits) {
            found_digits = true
            digits.push(html[j])
        }
        else if (isNaN(DIGIT) && !found_digits) {
            continue
        }
        else if (isNaN(DIGIT) && found_digits) {
            break
        }


    }

    return Number(digits.join(''))
}

function __MSU_YahooFinance(ticker: string) {
    const url = `https://finance.yahoo.com/quote/${ticker}?p=${ticker}`
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true })
    const contentText = res.getContentText()
    const open = __MSU_FindOpenStat(contentText)
    const current = __MSU_FindCurrentPPSStat(contentText)
    return [open, current]
}

function StockUpdates() {
    const INVEST_TAB = new GoogleSheetTabs(HOUSE_SAVINGS_TAB_NAME, "G2:Q10")
    const SHUT_OFF_TAB = new GoogleSheetTabs(HOUSE_SAVINGS_TAB_NAME, "K1")
    const START_COL = __Util_ColLetterToIndex("G")
    const START_ROW = 2

    const TICKER_INDEX = INVEST_TAB.GetHeaderIndex("Ticker")
    const PRICE_PER_SHARE_INDEX = INVEST_TAB.GetHeaderIndex("Price per Share")
    const TODAY_OPEN_INDEX = INVEST_TAB.GetHeaderIndex("Today's Open")
    const TOTAL_RETURN_INDEX = INVEST_TAB.GetHeaderIndex("Total Returns")
    const TODAYS_RETURN_INDEX = INVEST_TAB.GetHeaderIndex("Today's Loss/Gain")
    const TODAYS_RETURN_DOLLAR_INDEX = INVEST_TAB.GetHeaderIndex("Today's Loss/Gain $")
    const CUR_VALUE_INDEX = INVEST_TAB.GetHeaderIndex("Current Value")
    const COST_BASIS_INDEX = INVEST_TAB.GetHeaderIndex("Cost Basis")

    const SHOULD_SHUT_DOWN_UPDATE = Boolean(SHUT_OFF_TAB.GetRow(0)?.at(0))

    const CreatePercentDifBetween2NumbersFormula = (ticker: string, from: string, to: string) => {
        return `=IF(ISBLANK(${ticker}), 0, IFERROR((${from}-${to})/${to}, 0))`
    }

    const ToCellStr = (col_index: number, row: number) => {
        return `${__Util_IndexToColLetter(START_COL + col_index)}${row + START_ROW}`
    }

    INVEST_TAB.ForEachRow((row, i) => {

        if (SHOULD_SHUT_DOWN_UPDATE) {
            row[TOTAL_RETURN_INDEX] = 0
            row[TODAYS_RETURN_INDEX] = 0
            row[TODAYS_RETURN_DOLLAR_INDEX] = "$0"
        }
        else {
            row[TOTAL_RETURN_INDEX] = CreatePercentDifBetween2NumbersFormula(
                ToCellStr(TICKER_INDEX, i),
                ToCellStr(CUR_VALUE_INDEX, i),
                ToCellStr(COST_BASIS_INDEX, i)
            )

            row[TODAYS_RETURN_INDEX] = CreatePercentDifBetween2NumbersFormula(
                ToCellStr(TICKER_INDEX, i),
                ToCellStr(PRICE_PER_SHARE_INDEX, i),
                ToCellStr(TODAY_OPEN_INDEX, i)
            )

            row[TODAYS_RETURN_DOLLAR_INDEX] = `=IFERROR(DOLLAR(${ToCellStr(TODAY_OPEN_INDEX, i)}*${ToCellStr(TODAYS_RETURN_INDEX, i)}), 0)`
        }

        const [open, current] = __MSU_YahooFinance(String(row[TICKER_INDEX]))
        row[PRICE_PER_SHARE_INDEX] = current
        row[TODAY_OPEN_INDEX] = open

        return row
    }, 1)

    INVEST_TAB.SaveToTab()
}