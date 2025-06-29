// currently used
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

// currently used
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

// currently used
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

// currently used
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

// currently used
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
 * currently used
 * @returns {boolean} checks if date1 is greater than date2
 */
function __Util_CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(__Util_CreateDateString(date1));
  date2 = new Date(__Util_CreateDateString(date2));
  return date1.getTime() > date2.getTime();
}

// currently used
function __Util_GetDateFromDateHeader(date_header: string) {
  const HEADER_TEST = new RegExp(`^${PURCHASE_HEADER} \\d{1,2}/\\d{1,2}/\\d{1,2}\\d{1,2}$`)
  if (HEADER_TEST.test(date_header)) { return date_header.split(" ")[2] }
  return ""
}

// currently used
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

// currently used
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

// currently used
function __Util_GroupByDate(
  date_header: string,
  tab_name: string,
  shade_red: boolean = true
) {
  __Util_CreateHeadersForOneWeekLoans(date_header, tab_name)
  __Util_GroupAndHighlightOneWeekLoans(shade_red)
}

// currently used
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