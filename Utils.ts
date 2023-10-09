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

    public WriteCol(header_name: string, col: DataArrayEntry) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return }
        const LONGEST_ROW = this.FindLongestRowLength()

        for (let i = col.length-1; i >= 0; i--) {
            if (i >= this.data.length) { this.data.push(new Array(LONGEST_ROW).fill("")) }
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

    public AppendRow(row: DataArrayEntry) {
        this.data.push(this.CreateRowCopy(row))
    }

    public InsertRow(row_index: number, row: DataArrayEntry, AlterRow?: (row: DataArrayEntry) => DataArrayEntry) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        row = this.CreateRowCopy(row)
        if (AlterRow) { row = AlterRow(row) }
        this.data.splice(row_index, 0, row)
    }

    public FindRow(func: (row: DataArrayEntry) => boolean) {
        return this.data.find(func)
    }

    public GetRowRange(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        const RANGE_NOTATION = `A${row_index + 1}:${__IndexToColLetter(this.data[row_index].length)}${row_index + 1}`
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
  while (date.getUTCDay() !== 3) {
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
}

function __CreateNewHouseholdBudgetTab() {
  __CreateBudgetTab()
  ComputeMonthlyIncome()
}

/**
 * @returns {boolean} checks if date1 is greater than date2
 */
function __CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(__CreateDateString(date1));
  date2 = new Date(__CreateDateString(date2));
  return date1.getTime() > date2.getTime();
}
