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

  public PayOut() {
    let pay_amt = this.pay_out_amt;
    if (!this.ShouldPayOut(this.pay_date, this.total_days, this.day_inc)) {
      pay_amt = 0;
    }
    this.pay_date.setUTCDate(this.pay_date.getUTCDate() + this.day_inc);
    this.total_days += this.day_inc;
    return pay_amt;
  }

  public PayMonth() {
    const MONTH_DAY = this.pay_date.getUTCDate();
    const MONTH = this.pay_date.getUTCMonth();
    if (MONTH_DAY >= 28) {
      return this.months[(MONTH + 1) % this.months.length];
    }
    return this.months[MONTH];
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
      for (let i = 0; i < this.data[col_index].length; i++) {
        COL.push(this.data[col_index][i])
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

    public FindRow(func: (row: DataArrayEntry) => boolean) {
        return this.data.find(func)
    }

    public IndexOfRow(row?: DataArrayEntry) {
      if (row === undefined) { return -1 }
      return this.data.indexOf(row)
    }

    public GetRowRange(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        const RANGE_NOTATION = `A${row_index + 1}:${__IndexToColLetter(this.data[row_index].length)}${row_index + 1}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public GetRowSubRange(row_index: number, start: number, end: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }

        if (start > end || end < start) { start = end }
        if (start < 0) { start = 0 }
        if (end < 0) { end = 0 }

        const RANGE_NOTATION = `${__IndexToColLetter(start)}${row_index + 1}:${__IndexToColLetter(end)}${row_index + 1}`
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
            this.data[i] = this.data[i].map(__ConvertToStrOrNum)
        }
    }

    private CreateRowCopy(row: any[]) {
        return [...row].map(__ConvertToStrOrNum)
    }

    private InitSheetData() {
      const RANGE_DATA = this.tab.getDataRange().getValues().map(row => row.map(__ConvertToStrOrNum))
      this.data = this.tab.getDataRange().getFormulas()

      for (let row = 0; row < RANGE_DATA.length; row++) {
        for (let col = 0; col < RANGE_DATA[row].length; col++) {
          if (this.data[row][col] !== "") { continue }
          this.data[row][col] = RANGE_DATA[row][col]
        }
      }

    }
}

class Console {
  public static Log(...val: any[]) {
    console.log(val.join(" "))
  }
}

function __ConvertToStrOrNum(val: unknown) {
    let ret: number | string = ""

    if (val instanceof Date) { 
      ret = __CreateDateString(val)
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

function __IndexToColLetter(index: number) {
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

function __GetDateWhenCellEmpty(cell: any) {
  if (!cell) {
    return __CreateDateString(new Date(), true);
  }
  return cell;
}

function __AddToFixed(num: number, add_val: number, Round?: (x: number) => number) {
  let ret = ~~((num + add_val) * 100) / 100;
  if (Round) { ret = Round(ret); }
  return ret
}

function __SetDateToNextWeds(date: Date) {
  const WEDSDAY_INDEX = 3
  while (date.getUTCDay() !== WEDSDAY_INDEX) {
    date.setUTCDate(date.getUTCDate() + 1);
  }
  return date;
}

function __SetDateToNextFri(year: number) {
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

function __CreateDateString(date: Date | string, local: boolean = false) {
  if (typeof date === "string") {
    date = new Date(date);
  }

  let date_str = `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`;

  if (local) {
    date_str = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  }

  return date_str;
}

function __CreateBudgetTab() {
  const TEMPLATE_TAB = new GoogleSheetTabs("Household Budget Template");
  const SHEET = SpreadsheetApp.getActiveSpreadsheet();
  let year = new Date().getUTCFullYear();
  let budget_tab: GoogleSheetTabs;

  while (true) {
    try{
      budget_tab = new GoogleSheetTabs(`Household Budget ${year}`);
      year++;
    }
    catch(err) {
      break;
    }
  }

  SHEET.insertSheet(`Household Budget ${year}`);
  budget_tab = new GoogleSheetTabs(`Household Budget ${year}`);

  TEMPLATE_TAB.CopyTo(budget_tab);

  budget_tab.GetTab().setFrozenRows(3);
  budget_tab.GetTab().setFrozenColumns(1);
  
  const MISC_HOUSEHOLD_PURCHASES_INDEX = budget_tab.IndexOfRow(budget_tab.FindRow(row => row[0] === "Misc Household Purchases"))

  if (MISC_HOUSEHOLD_PURCHASES_INDEX !== -1) {
    budget_tab.GetTab().getRange(`B${MISC_HOUSEHOLD_PURCHASES_INDEX+1}:B`).setNumberFormat("mm/dd/yyyy")
    budget_tab.GetTab().getRange(`C${MISC_HOUSEHOLD_PURCHASES_INDEX+1}:C`).setNumberFormat("$#,##0.00")
  }

  return budget_tab
}

function __SetMonthDates(tab: GoogleSheetTabs) {
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

function __CreateNewHouseholdBudgetTab() {
  const TAB = __CreateBudgetTab()
  __SetMonthDates(TAB)
  ComputeMonthlyIncome()
}

function __EventObjectIsEditEventObject(e: any): e is {
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
function __CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(__CreateDateString(date1));
  date2 = new Date(__CreateDateString(date2));
  return date1.getTime() > date2.getTime();
}

function __Test() {
  let start_date = new Date("12/28/2023")
  let cur_year = start_date.getUTCFullYear()
  const __RosPayDay = function (_: Date, __: number, ___: number) {
    return true
  }

  const __DansPayDay = function (_: Date, total_days: number, inc: number) {
    const SHOULD_PAY =  total_days % (inc * 2) === 0;
    return SHOULD_PAY;
  }

  const ROS_PAY_DAY = new PayDay(350.95, start_date, __RosPayDay);
  ROS_PAY_DAY.SetPayoutDate(__SetDateToNextWeds);

  const MY_PAY_DAY = new PayDay(880.78, start_date, __DansPayDay);
  MY_PAY_DAY.SetPayoutDate(__SetDateToNextFri(cur_year - 1));
  let total = 0
  let total2 = 0

  while (MY_PAY_DAY.PayMonth() === "January" || ROS_PAY_DAY.PayMonth() === "January") {
    if (MY_PAY_DAY.PayMonth() === "January") {
      total2 = __AddToFixed(total2, MY_PAY_DAY.PayOut());
    }
    if (ROS_PAY_DAY.PayMonth() === "January") {
      total = __AddToFixed(total, ROS_PAY_DAY.PayOut())
    }
  }
}
