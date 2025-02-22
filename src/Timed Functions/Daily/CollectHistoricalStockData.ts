class __DCHSD_ChainNode {
    private ticker: string
    private value: number
    public next_node?: __DCHSD_ChainNode

    constructor(ticker: string, value: number) {
        this.ticker = ticker
        this.value = value
    }

    public toString(): string {
        let next_val = this.next_node ? this.next_node.toString() : 0
        return `IF(EQ("${this.ticker}", '${HOUSE_SAVINGS_TAB_NAME}'!T9), ${this.value}, ${next_val})`
    }
}

class __DCHSD_DynamicFormulaChain {
    private cur_node?: __DCHSD_ChainNode
    private head_node?: __DCHSD_ChainNode

    public AddNode(ticker: string, value: number) {
        if (this.head_node == null) {
            this.head_node = new __DCHSD_ChainNode(ticker, value)
            this.cur_node = this.head_node
        }
        else {
            this.cur_node!.next_node = new __DCHSD_ChainNode(ticker, value)
            this.cur_node = this.cur_node!.next_node
        }
    }

    public toString() {
        return this.head_node ? this.head_node.toString() : ""
    }
}

function __DCHSD_IsNotOnMarketDay() {
    const TODAY = new Date()
    return TODAY.getDay() === 0 || TODAY.getDay() === 6 || __Util_IsMarketHoliday(TODAY)
}

function __DCHSD_GetInvestmentData() {
    const INVEST_TAB = new GoogleSheetTabs(HOUSE_SAVINGS_TAB_NAME, INVESTMENT_TRACKER_RANGE)

    const PRICE_PER_SHARE_INDEX = INVEST_TAB.GetHeaderIndex("Price per Share")
    const TODAYS_OPEN_INDEX = INVEST_TAB.GetHeaderIndex("Today's Open")
    const TICKER_INDEX = INVEST_TAB.GetHeaderIndex("Ticker")

    const CLOSING_PRICES = new Array<number>()
    const RETURNS_PER_INVESTMENT = new Array<number>()
    const CLOSE_PRICE_CHAIN = new __DCHSD_DynamicFormulaChain()
    const RETURN_CHAIN = new __DCHSD_DynamicFormulaChain()

    const RoundDecTo = (n: number, places: number) => {
        const ROUND_PLACE = Math.pow(10, places)
        const RET = Math.round(n * ROUND_PLACE) / ROUND_PLACE
        return isNaN(RET) ? 0 : RET
    }

    const ReturnOnInvestment = (from: number, to: number) => {
        const FULL_PERCENTAGE =  ((to - from) / from) * 100
        return RoundDecTo(FULL_PERCENTAGE, 2)
    }

    INVEST_TAB.ForEachRow(row => {
        const STOCK_RETURN = ReturnOnInvestment(
            Number(row[TODAYS_OPEN_INDEX]),
            Number(row[PRICE_PER_SHARE_INDEX])
        )

        CLOSING_PRICES.push(Number(row[PRICE_PER_SHARE_INDEX]))
        RETURNS_PER_INVESTMENT.push(STOCK_RETURN)

        CLOSE_PRICE_CHAIN.AddNode(String(row[TICKER_INDEX]), Number(row[PRICE_PER_SHARE_INDEX]))
        RETURN_CHAIN.AddNode(String(row[TICKER_INDEX]), STOCK_RETURN)
    }, 1)

    return {
        date: __Util_CreateDateString(new Date()),
        avg_close: RoundDecTo((CLOSING_PRICES.reduce((p, c) => p + c, 0) / CLOSING_PRICES.length), 2),
        avg_ret_rage: RoundDecTo((RETURNS_PER_INVESTMENT.reduce((p, c) => p + c, 0) / RETURNS_PER_INVESTMENT.length), 2),
        single_stocks: CLOSE_PRICE_CHAIN,
        single_return: RETURN_CHAIN
    }
}

function CollectHistoricalStockData() {
    if (__DCHSD_IsNotOnMarketDay()) { return }
    const HISTORICAL_TAB = new GoogleSheetTabs(INVESTMENT_DATA_TAB_NAME)
    const DATA = __DCHSD_GetInvestmentData()
    const DATA_ARR = [
        DATA.date,
        `=IF('${HOUSE_SAVINGS_TAB_NAME}'!K1, 0, ${DATA.avg_close})`,
        `=IF('${HOUSE_SAVINGS_TAB_NAME}'!K1, 0, ${DATA.avg_ret_rage})`,
        `=IF('${HOUSE_SAVINGS_TAB_NAME}'!K1, 0, ${DATA.single_stocks})`,
        `=IF('${HOUSE_SAVINGS_TAB_NAME}'!K1, 0, ${DATA.single_return})`,
    ]
    HISTORICAL_TAB.AppendRow(DATA_ARR)
    HISTORICAL_TAB.SaveToTab()
}