function __Util_ConvertToStrOrNumOrBool(val: unknown) {
  let ret: number | string | boolean = ""

  if (val instanceof Date) {
    ret = __Util_CreateDateString(val)
  }
  else if (typeof val === "number") {
    ret = val
    if (isNaN(ret)) { ret = "" }
  }
  else if (typeof val === "string" || typeof val === 'boolean') {
    ret = val
  }
  else if (val != null) {
    ret = String(val)
  }

  return ret
}

function __Util_ColLetterToIndex(col: string) {
  let letters = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  ];

  if (col.split("").some(x => !letters.includes(x.toUpperCase()))) { return 0 }

  let result = 0
  const OFFSET = Number(col.length > 1)
  col = col.toUpperCase()

  for (let i = col.length; i > 0; i--) {
    let c = col.charAt(col.length - i)
    let num = letters.indexOf(c) + OFFSET
    result += num * Math.pow(26, i - 1)
  }

  return result - OFFSET
}

function __Util_IndexToColLetter(index: number) {
  const DIGITS = new Array<string>();
  const BASE = 26;
  const CHAR_CODE = "A".charCodeAt(0);

  if (index < 0) { index = 0; }

  while (true) {
    const LETTER_CODE = index % BASE;
    DIGITS.unshift(String.fromCharCode(LETTER_CODE + CHAR_CODE));
    if (index < 26) {
      break;
    }
    index = ~~(index / BASE) - 1;
  }

  return DIGITS.join("");
}

function __Util_GetDateWhenCellEmpty(cell: any) {
  if (!cell) {
    return __Util_CreateDateString(new Date(), true);
  }
  return cell;
}

function __Util_AddToFixed(num: number, add_val: number, Round?: (x: number) => number) {
  let ret = ~~((num + add_val) * 100) / 100;
  if (Round) { ret = Round(ret); }
  return ret
}

function __Util_SetDateToNextWeds(date: Date) {
  const WEDSDAY_INDEX = 3
  while (date.getUTCDay() !== WEDSDAY_INDEX) {
    date.setUTCDate(date.getUTCDate() + 1);
  }
  return date;
}

function __Util_SetDateToNextFri(year: number) {
  return function (date: Date) {
    while (date.getUTCDay() !== 5) {
      date.setUTCDate(date.getUTCDate() + 1);
    }
    if (year === date.getUTCFullYear()) {
      date.setUTCDate(date.getUTCDate() + 7);
    }
    return date;
  };
}

function __Util_CreateDateString(date: Date | string, local: boolean = false) {
  if (typeof date === "string") {
    date = new Date(date);
  }

  let date_str = `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`;

  if (local) {
    date_str = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  }

  return date_str;
}

function __Util_SetMonthDates(tab: GoogleSheetTabs) {
  const DATES_ROW = tab.GetRow(2)
  if (DATES_ROW === undefined) { return }

  const YEAR = tab.GetTab().getName().split(" ")[2]
  const START_END_DATES = [
    ["12/28/", "1/27/"],
    ["1/28/", "2/27/"],
    ["2/28/", "3/27/"],
    ["3/28/", "4/27/"],
    ["4/28/", "5/27/"],
    ["5/28/", "6/27/"],
    ["6/28/", "7/27/"],
    ["7/28/", "8/27/"],
    ["8/28/", "9/27/"],
    ["9/28/", "10/27/"],
    ["10/28/", "11/27/"],
    ["11/28/", "12/27/"],
  ]
  let i = 1

  for (const DATES of START_END_DATES) {
    let year_start = Number(YEAR)
    let year_end = year_start
    if (i === 1) { year_start-- }
    DATES_ROW[i] = `${DATES[0]}${year_start}`
    DATES_ROW[i + 1] = `${DATES[1]}${year_end}`
    i += 2
  }
  tab.WriteRow(2, DATES_ROW)
  tab.SaveToTab()
}

function __Util_EventObjectIsEditEventObject(e: any): e is {
  authMode: GoogleAppsScript.Script.AuthMode,
  range: GoogleAppsScript.Spreadsheet.Range,
  source: Spreadsheet,
  user: GoogleAppsScript.Base.User,
  oldValue?: string,
  triggerUid?: string,
  value?: string,
} {
  return e.authMode !== undefined && e.range !== undefined && e.source !== undefined && e.user !== undefined
}

/**
 * @returns {boolean} checks if date1 is greater than date2
 */
function __Util_CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(__Util_CreateDateString(date1));
  date2 = new Date(__Util_CreateDateString(date2));
  return date1.getTime() > date2.getTime();
}

function __Util_CacheSheets() {
  __CacheTab(WEEKLY_CREDIT_CHARGES_TAB_NAME)
  __CacheTab(MULTI_WEEK_LOANS_TAB_NAME)
}

function __CacheTab(tab_name: string) {
  const TAB = new GoogleSheetTabs(tab_name)
  const DATA: DataArray = []

  for (let i = 1; i < TAB.NumberOfRows(); i++) {
    const ROW = TAB.GetRow(i)
    if (ROW === undefined) { continue }
    DATA.push(ROW)
  }

  PropertiesService.getDocumentProperties().setProperty(tab_name, JSON.stringify(DATA))
}

function __Util_GetCachedOneWeekLoansData(tab_name: string) {
  let data = PropertiesService.getDocumentProperties().getProperty(tab_name)

  if (data === null) {
    __Util_CacheSheets()
    data = PropertiesService.getDocumentProperties().getProperty(tab_name)!
  }

  return JSON.parse(data) as DataArray
}

function __Util_CheckAllAreNotUndefined<T>(vals: T[]): vals is (T extends undefined ? never : T)[] {
  return __Util_CheckAllAreNot(undefined, vals)
}

function __Util_CheckAllAreNotInvalidIndex<T>(vals: T[]): vals is (T extends -1 ? never : T)[] {
  return __Util_CheckAllAreNot(-1, vals)
}

function __Util_CheckAllAreNot<T>(check_type: T, vals: unknown[]) {
  return vals.every(el => el !== check_type)
}

function __Util_GetDateFromDateHeader(date_header: string) {
  const HEADER_TEST = new RegExp(`^${PURCHASE_HEADER} \\d{1,2}/\\d{1,2}/\\d{1,2}\\d{1,2}$`)
  if (HEADER_TEST.test(date_header)) { return date_header.split(" ")[2] }
  return ""
}

function __Util_ComputeTotal() {
  const ONE_WEEK_LOAN_SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
  const ONE_WEEK_INTERPRETER = new FormulaInterpreter(ONE_WEEK_LOAN_SHEET)

  const PURCHASE_LOC_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Purchase Location")
  const AMOUNT_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Amount")
  const DUE_DATE_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Due Date")
  const TOTAL_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Total")
  const MONEY_LEFT_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Money Left")
  const PURCHASE_DATE_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Purchase Date")
  const WEEKLY_TOTALS = new Map<string, { total: number }>()
  const IN_CUR_MONTH_DATES = new Array<string>()
  const WEEKLY_SPENDING_LIMIT = [150, 126].at(-1)!

  const GetMoneyLeft = (money_left: number) => {
    return money_left - IN_CUR_MONTH_DATES.reduce((p, c) => p + WEEKLY_TOTALS.get(c)!.total, 0)
  }

  let last_date = ""
  let money_left = 0
  let in_cur_month = false

  ONE_WEEK_LOAN_SHEET.ForEachRow((row, i) => {
    const HEADER_DATE = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOC_COL_INDEX]))
    const DUE_DATE = String(row[DUE_DATE_COL_INDEX])

    let date = HEADER_DATE
    if (date === "") { date = DUE_DATE }

    if (date === "") {
      date = last_date
      row[DUE_DATE_COL_INDEX] = last_date
    }
    else if (date.toLowerCase() === "nw") {
      const LAST_DATE = new Date(last_date)
      LAST_DATE.setDate(LAST_DATE.getDate() + 7)
      date = __Util_CreateDateString(LAST_DATE)
      row[DUE_DATE_COL_INDEX] = date
    }

    const DetermineIfHeaderIsInCurMonthForMonthlyMoneyLeft = (last_row: DataArrayEntry) => {
      if (!in_cur_month && __Util_DateInCurrentPayPeriod(date)) {
        in_cur_month = true
      }
      else if (in_cur_month && !__Util_DateInCurrentPayPeriod(date)) {
        in_cur_month = false
        last_row[MONEY_LEFT_COL_INDEX] = GetMoneyLeft(money_left)
        money_left = 0
        IN_CUR_MONTH_DATES.splice(0, IN_CUR_MONTH_DATES.length)
      }

      if (in_cur_month) {
        money_left += WEEKLY_SPENDING_LIMIT
        IN_CUR_MONTH_DATES.push(date)
      }
    }

    const HeaderTotalSetup = () => {
      if (date !== "" && !WEEKLY_TOTALS.has(date)) {
        WEEKLY_TOTALS.set(date, { total: 0 })

        if (last_date === "") {
          last_date = date
        }
        else {
          const LAST_ROW = ONE_WEEK_LOAN_SHEET.GetRow(i - 1)!
          const LAST_TOTAL = WEEKLY_TOTALS.get(last_date)!
          LAST_ROW[TOTAL_COL_INDEX] = LAST_TOTAL.total

          DetermineIfHeaderIsInCurMonthForMonthlyMoneyLeft(LAST_ROW)

          ONE_WEEK_LOAN_SHEET.OverWriteRow(LAST_ROW)
          last_date = date
        }
      }
    }
    HeaderTotalSetup()

    const ComputeHeaderTotal = () => {
      if (!WEEKLY_TOTALS.has(date)) { return 'continue' }

      const TOTALS = WEEKLY_TOTALS.get(date)!
      let purchase_total = Number(row[AMOUNT_COL_INDEX])

      if (isNaN(purchase_total)) {
        const [did_parse, parse_res] = ONE_WEEK_INTERPRETER.AttemptToParseInput(row[AMOUNT_COL_INDEX])
        purchase_total = did_parse ? Number(parse_res) : 0
      }

      purchase_total = Math.ceil(purchase_total)


      TOTALS.total += purchase_total

      if (HEADER_DATE === "") {
        row[PURCHASE_DATE_COL_INDEX] = __Util_GetDateWhenCellEmpty(row[PURCHASE_DATE_COL_INDEX])
        row[TOTAL_COL_INDEX] = ""
        row[MONEY_LEFT_COL_INDEX] = in_cur_month ? "" : row[MONEY_LEFT_COL_INDEX]
      }

      if (i + 1 === ONE_WEEK_LOAN_SHEET.NumberOfRows()) {
        row[TOTAL_COL_INDEX] = TOTALS.total
        row[MONEY_LEFT_COL_INDEX] = in_cur_month ? GetMoneyLeft(money_left) : row[MONEY_LEFT_COL_INDEX]
      }
    }
    if (ComputeHeaderTotal() === 'continue') { return 'continue' }

    return row
  }, true)

  ONE_WEEK_LOAN_SHEET.SaveToTab()
}

function __Util_GroupAndHighlightOneWeekLoans(should_shade_red: boolean = true) {
  const SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
  const CUR_DATE = new Date()
  const PURCHASE_LOCATION_INDEX = SHEET.GetHeaderIndex("Purchase Location")
  const LIGHT_RED_SHADES = ["#FF7F7F", "#FF9F9F"]
  let color_index = 0
  let date = ""
  let last_date_header = -1

  SHEET.ForEachRow((row, i, range) => {

    if (last_date_header === -1 && String(row[PURCHASE_LOCATION_INDEX]).startsWith(PURCHASE_HEADER)) {
      date = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOCATION_INDEX]))
      last_date_header = i
    }
    else if (String(row[PURCHASE_LOCATION_INDEX]).startsWith(PURCHASE_HEADER) || i === SHEET.NumberOfRows() - 1) {
      const TAB = SHEET.GetTab()
      const DATE = new Date(String(row[PURCHASE_LOCATION_INDEX]).split(" ")[2])
      const RANGE_STR = `A${last_date_header + 2}:A${i + Number(i === SHEET.NumberOfRows() - 1)}`
      const RANGE = TAB.getRange(RANGE_STR)
      let group

      if (true) {
        try {
          group = TAB.getRowGroup(last_date_header + 2, 1)
          group?.remove()
          RANGE.shiftRowGroupDepth(1)
          group = TAB.getRowGroup(last_date_header + 2, 1)
        } catch {
          RANGE.shiftRowGroupDepth(1)
          group = TAB.getRowGroup(last_date_header + 2, 1)
        }
      }

      if (!group?.isCollapsed() && __Util_CompareDates(CUR_DATE, date)) {
        RANGE.collapseGroups()
      }

      date = __Util_CreateDateString(DATE)
      last_date_header = i
      color_index++
    }
    const BG_COLOR = range.getBackground().toUpperCase()
    if (__Util_CompareDates(CUR_DATE, date) && !LIGHT_RED_SHADES.includes(BG_COLOR)) {
      range.setBackground(LIGHT_RED_SHADES[color_index % LIGHT_RED_SHADES.length])
    }
  }, true)
}

function __Util_CreateHeadersForOneWeekLoans(date_header: string, tab_name: string) {
  const TAB = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
  const DATE_HEADER_INDEX = TAB.GetHeaderIndex(date_header)
  const PURCHASE_LOCATION_INDEX = TAB.GetHeaderIndex("Purchase Location")
  const TIPS_INDEX = TAB.GetHeaderIndex("Tips")
  const GROUPS = new Map<string, DataArray>()
  const HEADER_KEY = "HEADERS"

  GROUPS.set(HEADER_KEY, [TAB.GetRow(0)!])

  const CacheTipValue = (date: string, val: any) => {
    PropertiesService.getDocumentProperties().setProperty(`TIPS - ${date}`, String(val))
  }

  const GetTipValue = (date: string) => {
    const TIP = PropertiesService.getDocumentProperties().getProperty(`TIPS - ${date}`)
    PropertiesService.getDocumentProperties().deleteProperty(`TIPS - ${date}`)
    if (TIP == null) { return "" }
    return TIP
  }

  TAB.ForEachRow((row, i) => {
    if (i === 0) { return 'continue' }
    const ROW = row

    const DATE = String(ROW[DATE_HEADER_INDEX])
    if (String(row[PURCHASE_LOCATION_INDEX]).includes(PURCHASE_HEADER)) {
      const HEADER_DATE = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOCATION_INDEX]))
      CacheTipValue(HEADER_DATE, row[TIPS_INDEX])
    }
    
    if (DATE === "") { 
      return 'continue' 
    }

    if (!GROUPS.has(DATE)) {
      GROUPS.set(DATE, [])
    }

    

    GROUPS.get(DATE)!.push(ROW)
  })
  TAB.MakeInternalCopy()
  TAB.EraseTab()

  const CreateDateHeader = (len: number, date: string) => {
    const HEADER = new Array<string>()
    for (let i = 0; i < len; i++) {
      if (i === PURCHASE_LOCATION_INDEX) {
        HEADER.push(`${PURCHASE_HEADER} ${date}`)
      }
      else if (i === TIPS_INDEX) {
        HEADER.push(GetTipValue(date))
      }
      else {
        HEADER.push("")
      }
    }
    return HEADER
  }

  for (let [date_key, group] of GROUPS) {
    if (date_key === HEADER_KEY) {
      TAB.AppendRow(group[0])
      continue
    }

    let start_of_grouping = true
    for (let row of group) {
      if (start_of_grouping) {
        TAB.AppendRow(CreateDateHeader(row.length, date_key))
        start_of_grouping = false
      }
      TAB.AppendRow(row)
    }
  }

  TAB.RestoreFromInternalCopy()
  TAB.SaveToTab()
}

function __Util_GroupByDate(
  date_header: string,
  tab_name: string,
  shade_red: boolean = true
) {
  __Util_CreateHeadersForOneWeekLoans(date_header, tab_name)
  __Util_GroupAndHighlightOneWeekLoans(shade_red)
}

function __Util_DateInCurrentPayPeriod(compare_date: string) {
  const COMPARE_DATE = new Date(compare_date)
  const CUR_DAY = new Date()
  if (COMPARE_DATE.toString() === "Invalid Date") { return false }

  let test_month = COMPARE_DATE.getMonth()
  let test_day = COMPARE_DATE.getDate()
  let test_year = COMPARE_DATE.getFullYear()

  if (test_day >= 28) {
    test_month = (test_month + 1) % 12
    test_year = test_month === 0 ? test_year + 1 : test_year
  }

  let cur_month = CUR_DAY.getMonth()
  let cur_day = CUR_DAY.getDate()
  let cur_year = CUR_DAY.getFullYear()

  if (cur_day >= 28) {
    cur_month = (cur_month + 1) % 12
    cur_year = cur_month === 0 ? cur_year + 1 : cur_year
  }

  const SAME_MONTH = cur_month === test_month
  const SAME_YEAR = cur_year === test_year

  return SAME_MONTH && SAME_YEAR
}

function __Util_IsMarketHoliday(date: Date) {
  const MARKET_HOLIDAYS = [
      "New Year's Day",
      "Martin Luther King Jr. Day",
      "Presidents' Day",
      "Good Friday",
      "Memorial Day",
      "Juneteenth National Independence Day",
      "Independence Day",
      "Labor Day",
      "Thanksgiving Day",
      "Christmas Day",
  ].map(e => e.toLowerCase())

  const HOLIDAYS = CalendarApp.getCalendarsByName("Holidays in United States")
  const CHRISTIAN_HOLIDAYS = CalendarApp.getCalendarsByName("Christian Holidays")
  const TODAYS_EVENTS = [...HOLIDAYS[0].getEventsForDay(date), ...CHRISTIAN_HOLIDAYS[0].getEventsForDay(date)]
      .map(e => e.getTitle().toLowerCase())
  return TODAYS_EVENTS.some(e => MARKET_HOLIDAYS.includes(e))
}