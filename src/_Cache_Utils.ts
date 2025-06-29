type WeekInfo = {
    date: string,
    start_row: number,
    end_row: number,
    range_str: string,
    range_shade: string,
    grouping_range: string
}

type WeekObj = {
    week_info_list: WeekInfo[],
    last_date: string
}

type QueryAttrs = 'date' | 'start' | 'end' | 'full range' | 'group range' | 'shade'

const CURRENT_MONTH_KEY = "ONE_WEEK_CUR_MONTH"
const [LIGHT_RED_1, LIGHT_RED_2] = ["#FF7F7F", "#FF9F9F"]

function __CU_CreateEmptyWeekObject(): WeekObj {
    return {
        week_info_list: [],
        last_date: "",
    }
}

function __CU_CreateEmptyWeekInfoObj(): WeekInfo {
    return {
        date: "",
        start_row: -1,
        end_row: -1,
        range_str: "",
        range_shade: "",
        grouping_range: ""
    }
}

function __CU_GetNextShade(shade: string | undefined | null) {
    if (shade === LIGHT_RED_1) { return LIGHT_RED_2 }
    else if (shade === LIGHT_RED_2) { return LIGHT_RED_1 }
    else { return LIGHT_RED_1 }
}

function __Cache_Utils_StoreOneWeekLoanCurrentMonthInfo(start = -1, end = -1) {
    const SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)

    const PURCHASE_LOC_INDEX = SHEET.GetHeaderIndex("Purchase Location")

    const OBJ = __CU_CreateEmptyWeekObject()
    let date = ""
    let is_current_month = false
    const WEEK_MAP = new Map<string, number[]>()

    if (start < 0) { start = 1 }
    if (end < 0) { end = SHEET.NumberOfRows() }

    SHEET.ForEachRow((row, i) => {
        if (i > end) { return 'break' }

        const IS_HEADER = String(row[PURCHASE_LOC_INDEX]).startsWith(PURCHASE_HEADER)
        if (IS_HEADER) {
            date = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOC_INDEX]))
            is_current_month = __Util_DateInCurrentPayPeriod(date)
        }

        if (is_current_month) {
            if (!WEEK_MAP.has(date)) {
                WEEK_MAP.set(date, [])
            }

            WEEK_MAP.get(date)!.push(i)
        }
    }, start)

    for (let [key, index] of WEEK_MAP) {
        const WEEK_INFO = __CU_CreateEmptyWeekInfoObj()
        WEEK_INFO.date = key
        WEEK_INFO.start_row = index[0]
        WEEK_INFO.end_row = index.at(-1)!

        let row_mod = Number(WEEK_INFO.start_row + 1 === WEEK_INFO.end_row)

        WEEK_INFO.range_str = `A${WEEK_INFO.start_row + 1}:J${WEEK_INFO.end_row + row_mod}`
        WEEK_INFO.range_shade = __CU_GetNextShade(OBJ.week_info_list.at(-1)?.range_shade)
        WEEK_INFO.grouping_range = row_mod === 1 ? `A${WEEK_INFO.end_row + row_mod}:J${WEEK_INFO.end_row + row_mod}` : `A${WEEK_INFO.start_row + 2}:J${WEEK_INFO.end_row + 1}`
        OBJ.week_info_list.push(WEEK_INFO)
    }

    OBJ.last_date = OBJ.week_info_list.at(-1)!.date

    PropertiesService.getDocumentProperties().setProperty(CURRENT_MONTH_KEY, JSON.stringify(OBJ))
}

function __Cache_Utils_HasOneWeekLoanInfo() {
    const INFO = PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)
    return typeof INFO === 'string'
}

function __Cache_Utils_QueryAll(attr: QueryAttrs): Array<number | string> | null {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return null }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    const RES = new Array<number | string>()

    for (let info of INFO.week_info_list) {

        switch (attr) {
            case 'date': { 
                RES.push(info.date) 
                break
            }
            case 'start': { 
                RES.push(info.start_row)
                break
            }
            case 'end': { 
                RES.push(info.end_row) 
                break
            }
            case 'full range': { 
                RES.push(info.range_str) 
                break
            }
            case 'group range': { 
                RES.push(info.grouping_range) 
                break
            }
            case 'shade': { 
                RES.push(info.range_shade) 
                break
            }
        }

    }

    return RES
}

function __Cache_Utils_QueryAllAttrs(): Array<WeekInfo> | null {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return null }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    return INFO.week_info_list
}

function __Cache_Utils_QueryFirstWeek(attr: QueryAttrs): number | string | null {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return null }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    const FIRST = INFO.week_info_list[0]

    if (!FIRST) { return null }

    switch (attr) {
        case 'date': { return FIRST.date }
        case 'start': { return FIRST.start_row }
        case 'end': { return FIRST.end_row }
        case 'full range': { return FIRST.range_str }
        case 'group range': { return FIRST.grouping_range }
        case 'shade': { return FIRST.range_shade }
    }
}

function __Cache_Utils_QueryLastWeek(attr: QueryAttrs): number | string | null {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return null }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    const LAST = INFO.week_info_list.at(-1)

    if (!LAST) { return null }

    switch (attr) {
        case 'date': { return LAST.date }
        case 'start': { return LAST.start_row }
        case 'end': { return LAST.end_row }
        case 'full range': { return LAST.range_str }
        case 'group range': { return LAST.grouping_range }
        case 'shade': { return LAST.range_shade }
    }
}

function __Cache_Utils_RecacheCurrentMonthInfo() {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    const START = INFO.week_info_list[0].start_row

    __Cache_Utils_StoreOneWeekLoanCurrentMonthInfo(START)
}

function __Cache_Utils_ShouldRecacheCurrentMonthInfo() {
    if (!__Cache_Utils_HasOneWeekLoanInfo()) { return -1 }
    const INFO = JSON.parse(PropertiesService.getDocumentProperties().getProperty(CURRENT_MONTH_KEY)!)
    const DATE = INFO.week_info_list.at(-1)!.date

    if (!DATE) { return -1 }

    const TODAY = new Date()
    const COMP_DATE = new Date(DATE)

    if (TODAY > COMP_DATE && (TODAY.getDate() >= 28 || TODAY.getMonth() > COMP_DATE.getMonth())) { return 1 }

    return 0
}