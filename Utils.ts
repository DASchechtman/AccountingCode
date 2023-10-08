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
        this.data = this.tab.getDataRange().getValues().map(row => row.map(__ConvertToStrOrNum))

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

    public HasHeader(header_name: string) {
        return this.headers.has(header_name)
    }

    public GetCol(header_name: string) {
        const COL = new Array<number | string>()
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

        for (let i = col.length-1; i >= 0; i--) {
            if (!this.data[i]) { this.data[i] = [] }
            this.data[i][COL_INDEX] = col[i]
        }
    }

    public AppendCol(col: DataArrayEntry) {
        this.SetAllRowsToSameLength()
        for (let i = 0; i < this.data.length; i++) {
            if (i >= col.length) { 
                this.data[i].push("") 
            }
            else { 
                this.data[i].push(col[i])
            }
        }
    }

    public FindCol(func: (col: DataArrayEntry) => boolean) {
        const HEADERS = this.GetHeaderNames()
        for (const HEADER of HEADERS) {
            const COL = this.GetCol(HEADER)!
            if (func(COL)) { return COL }
        }
        return undefined
    }

    public GetColRange(header_name: string) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return undefined }
        const COL_LETTER = __IndexToColLetter(COL_INDEX)
        const RANGE_NOTATION = `${COL_LETTER}1:${COL_LETTER}${this.data.length}`
        return this.tab.getRange(RANGE_NOTATION)
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

    public GetCell(row_index: number, col_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return "" }
        if (col_index < 0 || col_index >= this.data[row_index].length) { return "" }
        return this.data[row_index][col_index]
    }

    public WriteCell(row_index: number, col_index: number, val: any) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        if (col_index < 0 || col_index >= this.data[row_index].length) { return }
        this.data[row_index][col_index] = val
    }

    public GetCellByHeader(row_index: number, header_name: string) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return "" }
        return this.data[row_index][COL_INDEX]
    }

    public WriteCellByHeader(row_index: number, header_name: string, val: any) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return }
        this.data[row_index][COL_INDEX] = val
    }

    public SaveToTab() {
        this.SetAllRowsToSameLength()
        const WRITE_RANGE = this.tab.getRange(1, 1, this.data.length, this.data[0].length)
        WRITE_RANGE.setValues(this.data)
    }

    public GetTab() {
        return this.tab
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
            this.data[i] = this.data[i].map(x => x == null ? "" : x)
        }
    }

    private CreateRowCopy(row: any[]) {
        return [...row].map(x => x == null ? "" : __ConvertToStrOrNum(x))
    }
}

function __ConvertToStrOrNum(val: unknown) {
    if (val instanceof Date) { return __CreateDateString(val) }
    if (typeof val === "number") { return isNaN(val) ? "" : val }
    return String(val) === "NaN" ? "" : String(val)
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

function __RoundUpToNearestDollar(table_val: string) {
  if (table_val === "") { return ""; }
  const NUM = Number(table_val);
  if (isNaN(NUM)) { return ""; }
  return Math.ceil(NUM);
}

function __FindMultiWeekRepayment(date_string: string) {
  const TAB = new GoogleSheetTabs("Multi Week Loans");
  const DATE = new Date(date_string);
  const DATE_COL = TAB.GetHeaderIndex("Repayment Date");
  const LOANEE_COL = TAB.GetHeaderIndex("Loanee");
  const REPAYMENT_COL = TAB.GetHeaderIndex("Repayment Amount");
  let repayment = 0;

  const REPAYMENT_ROW = TAB.FindRow(
    (x) => new Date(x[DATE_COL]).toDateString() === DATE.toDateString()
  );

  if (!REPAYMENT_ROW || REPAYMENT_ROW[LOANEE_COL] === "Dan") {
    return repayment;
  }
  return Number(REPAYMENT_ROW[REPAYMENT_COL]);
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

function __GetFirstNonEmptyCellInCol(
  tab: GoogleAppsScript.Spreadsheet.Sheet,
  Col: () => number
) {
  const TBL = tab.getDataRange().getValues();
  let i = 1;

  while (i < TBL.length && TBL[i][Col()] === "") {
    i++;
  }

  return i + 1;
}

function __CreateNewHouseholdBudgetTab() {
  const HOUSE_HOLD_BUDGET_TAB_NAME = "Household Budget";
  let date_year = new Date().getUTCFullYear();

  const SHEET = SpreadsheetApp.getActiveSpreadsheet();
  const BUDGET_TEMPLATE = SHEET.getSheetByName(
    `${HOUSE_HOLD_BUDGET_TAB_NAME} Template`
  );

  if (!BUDGET_TEMPLATE) {
    return;
  }

  while (SHEET.getSheetByName(`${HOUSE_HOLD_BUDGET_TAB_NAME} ${date_year}`)) {
    date_year++;
  }

  const NEW_TAB = SHEET.insertSheet(
    `${HOUSE_HOLD_BUDGET_TAB_NAME} ${date_year}`
  );

  for (let i = 1; i <= BUDGET_TEMPLATE.getLastColumn(); i++) {
    for (let j = 1; j <= BUDGET_TEMPLATE.getLastRow(); j++) {
      const TEMPLATE_CELL = BUDGET_TEMPLATE.getRange(j, i);
      let new_cell = NEW_TAB.getRange(j, i);
      const CELL_FORMULA = TEMPLATE_CELL.getFormula();
      const CELL_VALUE = TEMPLATE_CELL.getValue();
      const IS_MERGED = TEMPLATE_CELL.getMergedRanges().length > 0;

      if (IS_MERGED && i % 2 !== 0) {
        continue;
      }

      if (IS_MERGED) {
        new_cell = NEW_TAB.getRange(
          TEMPLATE_CELL.getMergedRanges()[0].getA1Notation()
        );
        new_cell.merge();
      }

      if (CELL_FORMULA !== "") {
        new_cell.setFormula(CELL_FORMULA);
      } else {
        new_cell.setValue(CELL_VALUE);
      }

      TEMPLATE_CELL.copyTo(
        new_cell,
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
    }
    NEW_TAB.autoResizeColumn(i);
  }

  NEW_TAB.setFrozenRows(3);
  NEW_TAB.setFrozenColumns(1);
}

/**
 * @returns {boolean} checks if date1 is greater than date2
 */
function __CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(__CreateDateString(date1));
  date2 = new Date(__CreateDateString(date2));
  return date1.getTime() > date2.getTime();
}
