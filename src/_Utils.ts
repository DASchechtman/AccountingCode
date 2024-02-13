class PayDay {
  private pay_out_amt: number;
  private pay_date: Date;
  private ShouldPayOut: CheckPayOut;
  private months: string[]
  private day_inc = 7;
  private total_days = 0;

  constructor(pay_out_amt: number, pay_date: Date, ShouldPayOut: CheckPayOut) {
    this.pay_out_amt = pay_out_amt;
    this.ShouldPayOut = ShouldPayOut;
    this.pay_date = new Date(pay_date);
    this.months = MONTHS
  }

  public SetPayoutDate(PayOutDate: (date: Date) => Date) {
    this.pay_date = PayOutDate(this.pay_date);
  }

  public SetPayoutAmount(pay_out_amt: number) {
    this.pay_out_amt = pay_out_amt;
  }

  public SetPayoutCheck(ShouldPayOut: CheckPayOut) {
    this.ShouldPayOut = ShouldPayOut;
  }

  public PayOut() {
    let pay_amt = this.pay_out_amt;
    const SHOULD_PAY = this.ShouldPayOut({
      date: this.pay_date, 
      total_days: this.total_days, 
      inc: this.day_inc, 
      pay_month: this.PayMonth()
    })

    if (!SHOULD_PAY) {
      pay_amt = 0;
    }

    this.pay_date.setUTCDate(this.pay_date.getUTCDate() + this.day_inc);
    this.total_days += this.day_inc;
    return pay_amt;
  }

  public PayMonth() {
    return this.months[this.GetMonthIndex()];
  }

  private GetMonthIndex() {
    let month = this.pay_date.getUTCMonth();
    const MONTH_DAY = this.pay_date.getUTCDate();
    if (MONTH_DAY >= 28) {
      month = (month + 1) % this.months.length;
    }
    return month
  }
}

class GoogleSheetTabs {
    private tab: Tab
    private headers: Map<string, number>
    private data: DataArray

    constructor(tab: Tab | string) {
        if (typeof tab === "string") {
            const SHEET_TAB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab)
            if (SHEET_TAB === null) { throw new Error("Tab does not exist") }
            tab = SHEET_TAB
        }

        this.tab = tab
        this.data = []
        this.InitSheetData()

        const HEADERS = this.data[0]

        this.headers = new Map<string, number>()
        for (let i = 0; i < HEADERS.length; i++) {
            const HEADER = HEADERS[i]
            if (typeof HEADER !== "string") { continue }
            this.headers.set(HEADER, i)
        }
    }

    public GetHeaderIndex(header_name: string) {
        return this.headers.get(header_name) === undefined ? -1 : this.headers.get(header_name)!
    }

    public GetHeaderNames() {
        return Array.from(this.headers.keys())
    }

    public GetCol(header_name: string) {
        const COL: DataArrayEntry = []
        const COL_INDEX = this.headers.get(header_name)

        if (COL_INDEX === undefined) { return undefined }

        for (let i = 0; i < this.data.length; i++) {
            COL.push(this.data[i][COL_INDEX])
        }

        return COL
    }

    public GetColByIndex(col_index: number) {
      if (col_index < 0 || col_index >= this.data[0].length) { return undefined }
      const COL: DataArrayEntry = []
      for (let i = 0; i < this.data.length; i++) {
        COL.push(this.data[i][col_index])
      }
      return COL
    }

    public WriteCol(header_name: string, col: DataArrayEntry) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return }
        const LONGEST_ROW = this.FindLongestRowLength()

        for (let i = col.length-1; i >= 0; i--) {
            if (this.data[i] === undefined) { this.data[i] = new Array(LONGEST_ROW).fill("") }
            this.data[i][COL_INDEX] = col[i]
        }
    }

    public GetRow(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        return this.CreateRowCopy(this.data[row_index])
    }

    public WriteRow(row_index: number, row: DataArrayEntry) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        this.data[row_index] = this.CreateRowCopy(row)
    }

    public WriteRowAt(row_index: number, start: number, row: DataArrayEntry) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        if (start < 0) { start = 0 }
        while (start + row.length >= this.data[row_index].length) { this.data[row_index].push("") }

        for (let i = 0; i < row.length; i++) {
            this.data[row_index][start + i] = row[i]
        }
    }

    public AppendRow(row: DataArrayEntry, should_fill: boolean = false) {
        row = this.CreateRowCopy(row)
        this.data.push(row)
        if (should_fill) {
          const LONGEST_ROW = this.FindLongestRowLength()
          while (row.length < LONGEST_ROW) {
            row.push("")
          }
        }
        return row
    }

    public InsertRow(row_index: number, row: DataArrayEntry, { AlterRow, should_fill }: { 
      AlterRow?: (row: DataArrayEntry) => DataArrayEntry, 
      should_fill?: boolean 
    } = {}) {
        if (row_index < 0) { row_index = 0 }
        row = this.CreateRowCopy(row)
        if (AlterRow) { row = AlterRow(row) }

        const LONGEST_ROW = this.FindLongestRowLength()
        while (row.length < LONGEST_ROW && should_fill) {
          row.push("")
        }

        if (row_index >= this.data.length) { return this.AppendRow(row) }
        this.data.splice(row_index, 0, row)

        return row
    }

    public AppendToRow(row_index: number, ...row: DataArrayElement[]) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        this.data[row_index].push(...row.map(__Util_ConvertToStrOrNum))
        return row
    }

    public FindRow(func: (row: DataArrayEntry) => boolean) {
        return this.data.find(func)
    }

    public IndexOfRow(row?: DataArrayEntry | ((row: DataArrayEntry) => boolean), index_from?: number) {
      let search_row = row
      if (typeof search_row === "function") { search_row = this.FindRow(search_row) }
      if (search_row === undefined) { return -1 }
      return this.data.indexOf(search_row, index_from)
    }

    public GetRowRange(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        const RANGE_NOTATION = `A${row_index + 1}:${__Util_IndexToColLetter(this.data[row_index].length)}${row_index + 1}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public GetRowSubRange(row_index: number, start: number, end: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }

        if (start > end || end < start) { start = end }
        if (start < 0) { start = 0 }
        if (end < 0) { end = 0 }

        const RANGE_NOTATION = `${__Util_IndexToColLetter(start)}${row_index + 1}:${__Util_IndexToColLetter(end)}${row_index + 1}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public GetRange(start_row: number, end_row: number, start_col: number, end_col: number) {
        const RANGE_1 = this.GetRowSubRange(start_row, start_col, end_col)
        const RANGE_2 = this.GetRowSubRange(end_row, start_col, end_col)

        if (RANGE_1 === undefined || RANGE_2 === undefined) { return undefined }
        const FIRST_NOTATION_PART = RANGE_1.getA1Notation().split(":")[0]
        const SECOND_NOTATION_PART = RANGE_2.getA1Notation().split(":")[1]
        const RANGE_NOTATION = `${FIRST_NOTATION_PART}:${SECOND_NOTATION_PART}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public NumberOfRows() {
        return this.data.length
    }

    public SaveToTab() {
        this.SetAllRowsToSameLength()
        const WRITE_RANGE = this.tab.getRange(1, 1, this.data.length, this.data[0].length)
        WRITE_RANGE.setValues(this.data)
    }

    public GetTab() {
        return this.tab
    }

    public CopyTo(tab: GoogleSheetTabs) {
        for(let i = 0; i < this.data.length; i++) {
          if (i >= tab.NumberOfRows()) {
            tab.AppendRow(this.data[i])
          }
          else {
            tab.WriteRow(i, this.data[i])
          }

          const ROW_RANGE = this.GetRowRange(i)
          const TAB_ROW_RANGE = tab.GetRowRange(i)
          if (ROW_RANGE === undefined || TAB_ROW_RANGE === undefined) { continue }
          ROW_RANGE.copyTo(TAB_ROW_RANGE, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
        }

        for (let i = 0; i < tab.data[0].length; i++) {
          tab.GetTab().autoResizeColumn(i+1)
          const width = tab.GetTab().getColumnWidth(i+1)
          tab.GetTab().setColumnWidth(i+1, width+25)
        }
    }

    public ClearTab() {
      this.data.map(row => row.fill(""))
    }

    private FindLongestRowLength() {
        let longest_row = -1
        for (let i = 0; i < this.data.length; i++) {
            if (this.data[i].length > longest_row) {
                longest_row = this.data[i].length
            }
        }
        return longest_row
    }

    private SetAllRowsToSameLength() {
        const LONGEST_ROW = this.FindLongestRowLength()
        for (let i = 0; i < this.data.length; i++) {
            while (this.data[i].length < LONGEST_ROW) {
                this.data[i].push("")
            }
            this.data[i] = this.data[i].map(__Util_ConvertToStrOrNum)
        }
    }

    private CreateRowCopy(row: any[]) {
        return [...row].map(__Util_ConvertToStrOrNum)
    }

    private InitSheetData() {
      const RANGE_DATA = this.tab.getDataRange().getValues().map(row => row.map(__Util_ConvertToStrOrNum))
      this.data = this.tab.getDataRange().getFormulas()

      for (let row = 0; row < RANGE_DATA.length; row++) {
        for (let col = 0; col < RANGE_DATA[row].length; col++) {
          if (this.data[row][col] !== "") { continue }
          this.data[row][col] = RANGE_DATA[row][col]
        }
      }

    }
}

function __Util_ConvertToStrOrNum(val: unknown) {
    let ret: number | string = ""

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
    else if (val != null) {
      ret = String(val)
    }

    return ret
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
    DATES_ROW[i+1] = `${DATES[1]}${year_end}`
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

function __Util_ComputeTotal() {
    const TAB_NAME = "One Week Loans";
    const SHEET = new GoogleSheetTabs(TAB_NAME);
  
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
      const PURCHASE_LOCATION = String(PURCHASE_LOCATION_INDEX[i]);
      const DUE_DATE = String(DUE_DATE_INDEX[i])
      const AMOUNT = typeof AMOUNT_INDEX[i] === "number" ? Number(AMOUNT_INDEX[i]) : -1;
  
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
        total = AMOUNT;
        last_amt = i
        TOTAL_INDEX[i] = ""
      }
      else {
        total += AMOUNT;
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

function __Util_GroupByDate(
  date_header: string,
  tab_name: string,
  shade_red: boolean = true
) {
  const TAB = new GoogleSheetTabs(tab_name);
  const CURRENT_ONE_WEEK_TAB = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME);
  const ROW_COMPARE = new Map<string, [number, number]>()
  const CACHE_INDEX = 0
  const CURRENT_INDEX = 1

  const __InsertGroupingRow = function (date: string) {
    return function (arr: DataArrayEntry) {
      arr[PURCHASE_LOCATION_INDEX] = `${PURCHASE_HEADER} ${date}`;
      if (CARD_INDEX >= 0) { arr[CARD_INDEX] = " " }
      return arr;
    };
  };

  const __GetNextDateGroup = function (date_header: string, row_index: number) {
    const SEARCH_DATE = date_header.split(" ")[2];
    const DateFound = function (row: (string | number)[]) {
      return (
        String(row[PURCHASE_LOCATION_INDEX]).includes(SEARCH_DATE) ||
        String(row[DATE_COL_INDEX]).includes(SEARCH_DATE)
      );
    };
    let row = TAB.GetRow(row_index);

    while (row && DateFound(row)) {
      row_index++;
      row = TAB.GetRow(row_index);
    }
    return row_index - 1;
  };

  const __StoreCompResults = function (key: string, index: number) {
    if (!ROW_COMPARE.has(key)) {
      ROW_COMPARE.set(key, [0, 0])
      ROW_COMPARE.get(key)![index]++
    }
    else {
      ROW_COMPARE.get(key)![index]++
    }
  }

  const __GetCompResults = function(key: string) {
    const COMP = ROW_COMPARE.get(key)
    if (!COMP) { return false }
    return COMP[0] !== COMP[1]
  }

  const __GetCachedData = function(key: string, index: typeof CACHE_INDEX | typeof CURRENT_INDEX) {
    const DATA = ROW_COMPARE.get(key)
    if (!DATA) { return -1 }
    return DATA[index]
  }

  const __CheckIfDateEntriesAltered = function (date: string) {
    const CACHED_DATA = __Util_GetCachedOneWeekLoansData(tab_name)

    if (ROW_COMPARE.size > 0) { return __GetCompResults(date) }

    for (let i = 0; i < CACHED_DATA.length; i++) {
      __StoreCompResults(String(CACHED_DATA[i]![DATE_COL_INDEX]), CACHE_INDEX)
    }

    for (let i = 0; i < CURRENT_ONE_WEEK_TAB.NumberOfRows(); i++) {
      __StoreCompResults(String(CURRENT_ONE_WEEK_TAB.GetRow(i)![DATE_COL_INDEX]), CURRENT_INDEX)
    }
    
    return __GetCompResults(date)
  }

  const GenerateLoanGroupHeader = function () {
    let last_recorded_date = "";
    const FIRST_ROW_PAST_HEADERS = 1

    for (let i = FIRST_ROW_PAST_HEADERS; i < TAB.NumberOfRows(); i++) {
      const ROW = TAB.GetRow(i)!.map(x => String(x));
      const PURCHASE = ROW[PURCHASE_LOCATION_INDEX];
      const DATE_VAL = ROW[DATE_COL_INDEX];
      const DATE = PURCHASE.includes(PURCHASE_HEADER) ? PURCHASE.split(" ")[2] : "";

      if (DATE !== "" && !__CheckIfDateEntriesAltered(DATE)) { 
        const RES = __GetCachedData(DATE, CURRENT_INDEX)
        i += RES
        continue
      }
      else if (PURCHASE.includes(PURCHASE_HEADER)) {
        i = __GetNextDateGroup(PURCHASE, i);
        continue;
      } else if (DATE_VAL === "") {
        continue;
      }

      const NEW_DATE = __Util_CreateDateString(DATE_VAL);

      if (last_recorded_date === "" || last_recorded_date !== NEW_DATE) {
        last_recorded_date = NEW_DATE;
        TAB.InsertRow(i, [], { AlterRow: __InsertGroupingRow(NEW_DATE) });
      }
    }
  };

  const GetGroupBoundries = function () {
    const BOUNDRIES = new Map<string, [number, number, string]>();
    for (let i = 1; i < TAB.NumberOfRows(); i++) {
      const ROW = TAB.GetRow(i)
      if (!ROW) { continue }

      if (String(ROW[PURCHASE_LOCATION_INDEX]).includes(PURCHASE_HEADER)) {
        const DATE = String(ROW[PURCHASE_LOCATION_INDEX]).split(" ")[2];
        const ARR: [number, number, string] = [i + 1, 0, DATE]
        i = __GetNextDateGroup(String(ROW[PURCHASE_LOCATION_INDEX]), i);
        ARR[1] = i - (ARR[0] as number) + 1;
        BOUNDRIES.set(DATE, ARR);
      }
    }

    return BOUNDRIES
  }

  const GroupRowsInSheet = function () {
    const LIGHT_RED_SHADES = ["#FF7F7F", "#FF9F9F"]
    let i = 0

    for (const [_, val] of BOUNDRIES) {
      const DUE_DATE = new Date(val[2])
      const CUR_DATE = new Date()
      const DUE_DATE_HAS_PASSED = __Util_CompareDates(CUR_DATE, DUE_DATE)
      const GROUP_RANGE = TAB.GetTab().getRange(val[0] + 1, 1, val[1], TAB.GetTab().getLastColumn())
      const COLOR_RANGE = TAB.GetTab().getRange(val[0], 1, val[1] + 1, TAB.GetTab().getLastColumn())

      if (DUE_DATE_HAS_PASSED && shade_red) {
        COLOR_RANGE.setBackground(LIGHT_RED_SHADES[i++ % LIGHT_RED_SHADES.length])
      }

      try {
        let GROUP = TAB.GetTab().getRowGroup(val[0], 1)
        if (__CheckIfDateEntriesAltered(__Util_CreateDateString(DUE_DATE))) {
          GROUP?.remove()
          GROUP_RANGE.shiftRowGroupDepth(1)
          GROUP = TAB.GetTab().getRowGroup(val[0], 1)
        }
        if (DUE_DATE_HAS_PASSED) {
          GROUP?.collapse()
        }
      } catch {
        GROUP_RANGE.shiftRowGroupDepth(1)
      }
    }
  }

  const COL_INDEXES = [
    TAB.GetHeaderIndex(date_header),
    TAB.GetHeaderIndex("Purchase Location"),
    TAB.GetHeaderIndex("Card")
  ]
 

  if (!__Util_CheckAllAreNotInvalidIndex(COL_INDEXES)) {
    return;
  }

  const [
    DATE_COL_INDEX,
    PURCHASE_LOCATION_INDEX,
    CARD_INDEX
  ] = COL_INDEXES

  GenerateLoanGroupHeader()
  const BOUNDRIES = GetGroupBoundries()
  GroupRowsInSheet()

  TAB.SaveToTab();
}