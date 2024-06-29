function __Util_ConvertToStrOrNumOrBool(val: unknown) {
  let ret: number | string | boolean = ""

  if (val instanceof Date) {
    ret = __Util_CreateDateString(val)
  }
  else if (typeof val === "number") {
    ret = Number(val)
    if (isNaN(ret)) { ret = "" }
  }
  else if (typeof val === "string") {
    ret = val
  }
  else if (typeof val === 'boolean') {
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
  __CacheTab(ONE_WEEK_LOANS_TAB_NAME)
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

function __Util_ComputeTotal_Archived() {
  const TAB_NAME = "One Week Loans";
  const SHEET = new GoogleSheetTabs(TAB_NAME);
  const FORMULA_INTERPRETER = new FormulaInterpreter(SHEET);

  const PURCHASE_COL_HEADER = "Purchase Location";
  const DUE_DATE_COL_HEADER = "Due Date";
  const AMOUNT_COL_HEADER = "Amount";
  const TOTAL_COL_HEADER = "Total";
  const PURCHASE_DATE_COL_HEADER = "Purchase Date";

  const COLS = [
    SHEET.GetCol(PURCHASE_COL_HEADER),
    SHEET.GetCol(DUE_DATE_COL_HEADER),
    SHEET.GetCol(AMOUNT_COL_HEADER),
    SHEET.GetCol(TOTAL_COL_HEADER),
    SHEET.GetCol(PURCHASE_DATE_COL_HEADER)
  ]

  if (!__Util_CheckAllAreNotUndefined(COLS)) {
    COLS
    return;
  }

  const [
    PURCHASE_LOCATION_INDEX,
    DUE_DATE_INDEX,
    AMOUNT_INDEX,
    TOTAL_INDEX,
    PURCHASE_DATE_INDEX
  ] = COLS

  let total = 0;
  let last_amt = 0
  let last_recorded_date = "";

  for (let i = 1; i < SHEET.NumberOfRows(); i++) {
    const AMOUNT_VAL = typeof AMOUNT_INDEX[i] === 'string' ? FORMULA_INTERPRETER.ParseInput(AMOUNT_INDEX[i] as string) : AMOUNT_INDEX[i]
    const PURCHASE_LOCATION = String(PURCHASE_LOCATION_INDEX[i]);
    const DUE_DATE = String(DUE_DATE_INDEX[i])
    const AMOUNT = typeof AMOUNT_VAL === "number" ? AMOUNT_VAL : -1;

    if (PURCHASE_LOCATION.includes(PURCHASE_HEADER)) {
      continue;
    }

    if (DUE_DATE === "" || AMOUNT === -1 || PURCHASE_LOCATION === "") {
      TOTAL_INDEX[i] = "";
      continue;
    }

    if (i + 1 === SHEET.NumberOfRows()) {
      if (last_recorded_date !== DUE_DATE) {
        TOTAL_INDEX[last_amt] = __Util_AddToFixed(total, 0, Math.ceil)
        TOTAL_INDEX[i] = AMOUNT
      }
      else {
        TOTAL_INDEX[i] = __Util_AddToFixed(total, AMOUNT, Math.ceil)
      }
    }
    else if (last_recorded_date === "" || last_recorded_date !== DUE_DATE) {
      if (last_recorded_date !== "") {
        TOTAL_INDEX[last_amt] = __Util_AddToFixed(total, 0, Math.ceil)
      }
      last_recorded_date = DUE_DATE;
      total = Math.ceil(AMOUNT);
      last_amt = i
      TOTAL_INDEX[i] = ""
    }
    else {
      total += Math.ceil(AMOUNT);
      last_amt = i
      TOTAL_INDEX[i] = ""
    }

    PURCHASE_DATE_INDEX[i] = __Util_GetDateWhenCellEmpty(PURCHASE_DATE_INDEX[i]);
  }

  SHEET.WriteCol(PURCHASE_COL_HEADER, PURCHASE_LOCATION_INDEX)
  SHEET.WriteCol(DUE_DATE_COL_HEADER, DUE_DATE_INDEX)
  SHEET.WriteCol(AMOUNT_COL_HEADER, AMOUNT_INDEX)
  SHEET.WriteCol(TOTAL_COL_HEADER, TOTAL_INDEX.map(x => typeof x === "number" ? __Util_AddToFixed(x, 0, Math.ceil) : x))
  SHEET.WriteCol(PURCHASE_DATE_COL_HEADER, PURCHASE_DATE_INDEX)
  SHEET.SaveToTab();
}

function __Util_ComputeTotal() {
  const ONE_WEEK_LOAN_SHEET = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME)

  const PURCHASE_LOC_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Purchase Location")
  const AMOUNT_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Amount")
  const DUE_DATE_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Due Date")
  const TOTAL_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Total")
  const MONEY_LEFT_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Money Left")
  const PURCHASE_DATE_COL_INDEX = ONE_WEEK_LOAN_SHEET.GetHeaderIndex("Purchase Date")
  const WEEKLY_TOTALS = new Map<string, {total: number}>()
  const WEEKLY_SPENDING_LIMIT = 150
  
  let last_date = ""
  let money_left = 0
  let in_cur_month = false

  ONE_WEEK_LOAN_SHEET.ForEachRow((row, i) => {
    const HEADER_DATE = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOC_COL_INDEX]))
    if (HEADER_DATE !== "" && !WEEKLY_TOTALS.has(HEADER_DATE)) {
      WEEKLY_TOTALS.set(HEADER_DATE, {total: 0})
      
      if (last_date === "") {
        last_date = HEADER_DATE
      }
      else {
        const LAST_ROW = ONE_WEEK_LOAN_SHEET.GetRow(i-1)!
        const LAST_TOTAL = WEEKLY_TOTALS.get(last_date)!
        LAST_ROW[TOTAL_COL_INDEX] = LAST_TOTAL.total

        if (!in_cur_month && __Util_DateInCurrentPayPeriod(HEADER_DATE)) {
          in_cur_month = true
        }
        else if (in_cur_month && !__Util_DateInCurrentPayPeriod(HEADER_DATE)) {
          in_cur_month = false
          LAST_ROW[MONEY_LEFT_COL_INDEX] = money_left
          money_left = 0
        }

        if (in_cur_month) {
          money_left += WEEKLY_SPENDING_LIMIT
        }

        ONE_WEEK_LOAN_SHEET.OverWriteRow(LAST_ROW)
        last_date = HEADER_DATE
      }
    }

    const DUE_DATE = String(row[DUE_DATE_COL_INDEX])
    if (!WEEKLY_TOTALS.has(DUE_DATE)) { return 'continue' }
    
    const TOTALS = WEEKLY_TOTALS.get(DUE_DATE)!
    const PURCHASE_TOTAL = Number(row[AMOUNT_COL_INDEX])
    TOTALS.total += Math.ceil(PURCHASE_TOTAL)

    if (in_cur_month) {
      money_left -= Math.ceil(PURCHASE_TOTAL)
    }

    if (HEADER_DATE === "") {
      row[PURCHASE_DATE_COL_INDEX] = __Util_GetDateWhenCellEmpty(row[PURCHASE_DATE_COL_INDEX])
      row[TOTAL_COL_INDEX] = ""
      row[MONEY_LEFT_COL_INDEX] = ""
    }

    if (i + 1 === ONE_WEEK_LOAN_SHEET.NumberOfRows()) {
      row[TOTAL_COL_INDEX] = TOTALS.total
      row[MONEY_LEFT_COL_INDEX] = in_cur_month ? money_left : row[MONEY_LEFT_COL_INDEX]
    }

    return row
  }, true)

  ONE_WEEK_LOAN_SHEET.SaveToTab()
}

function __Util_GroupAndHighlightOneWeekLoans(should_shade_red: boolean = true) {
  const SHEET = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME)
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

      if (__Util_CompareDates(date, CUR_DATE)) {
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
  const TAB = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME)
  const DATE_HEADER_INDEX = TAB.GetHeaderIndex(date_header)
  const PURCHASE_LOCATION_INDEX = TAB.GetHeaderIndex("Purchase Location")
  const GROUPS = new Map<string, DataArray>()
  const HEADER_KEY = "HEADERS"

  GROUPS.set(HEADER_KEY, [TAB.GetRow(0)!])

  TAB.ForEachRow((row, i) => {
    if (i === 0) { return 'continue' }
    const ROW = row

    const DATE = String(ROW[DATE_HEADER_INDEX])
    if (DATE === "") { return 'continue' }

    if (!GROUPS.has(DATE)) {
      GROUPS.set(DATE, [])
    }

    GROUPS.get(DATE)!.push(ROW)
  })

  TAB.EraseTab()

  const CreateDateHeader = (len: number, date: string) => {
    const HEADER = new Array<string>()
    for (let i = 0; i < len; i++) {
      if (i === PURCHASE_LOCATION_INDEX) {
        HEADER.push(`${PURCHASE_HEADER} ${date}`)
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