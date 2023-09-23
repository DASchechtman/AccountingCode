type CheckPayOut = (date: Date, dir: number, inc: number) => boolean;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Tab = GoogleAppsScript.Spreadsheet.Sheet;

const PURCHASE_HEADER = "Purchases for"

class PayDay {
  private pay_out_amt: number;
  private pay_date: Date;
  private ShouldPayOut: CheckPayOut;
  private months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  private day_inc = 7;
  private total_days = 0;

  constructor(pay_out_amt: number, pay_date: Date, ShouldPayOut: CheckPayOut) {
    this.pay_out_amt = pay_out_amt;
    this.ShouldPayOut = ShouldPayOut;
    this.pay_date = new Date(pay_date);
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

function IndexToColLetter(index: number) {
  const DIGITS = new Array<number>();
  const BASE = 26;

  if (index < 0) {
    index = 0;
  }

  while (true) {
    const LETTER_CODE = index % BASE;
    DIGITS.push(LETTER_CODE);
    if (index < 26) {
      break;
    }
    index = ~~(index / BASE) - 1;
  }

  return DIGITS.map((x) => String.fromCharCode("A".charCodeAt(0) + x))
    .reverse()
    .join("");
}

function RoundUpToNearestDollar(table_val: string) {
  if (table_val === "") {
    return "";
  }
  const NUM = Number(table_val);

  if (isNaN(NUM)) {
    return "";
  }

  return Math.ceil(NUM);
}

function FindMultiWeekRepayment(date_string: string) {
  const TAB =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Multi Week Loans");
  const HEADERS = TAB!.getRange(1, 1, 1, TAB!.getLastColumn()).getValues()[0];
  const DATE = new Date(date_string);
  const DATE_COL = HEADERS.indexOf("Repayment Date");
  const LOANEE_COL = HEADERS.indexOf("Loanee");
  const REPAYMENT_COL = HEADERS.indexOf("Repayment Amount");
  let repayment = 0;

  if (!TAB) {
    return repayment;
  }

  const VALS = TAB.getDataRange().getValues();
  const REPAYMENT_ROW = VALS.find(
    (x) => new Date(x[DATE_COL]).toDateString() === DATE.toDateString()
  );

  if (!REPAYMENT_ROW || REPAYMENT_ROW[LOANEE_COL] === "Dan") {
    return repayment;
  }
  return Number(REPAYMENT_ROW[REPAYMENT_COL]);
}

function GetDateWhenCellEmpty(cell: any) {
    if (!cell) { return CreateDateString(new Date(), true); }
    return cell
}

function ComputeTotal() {
  const TAB_NAME = "One Week Loans";
  const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAME);

  if (!SHEET) {
    return;
  }

  const DATA = SHEET.getDataRange().getValues();
  const PURCHASE_LOCATION_INDEX = DATA[0].indexOf("Purchase Location");
  const DUE_DATE_INDEX = DATA[0].indexOf("Due Date");
  const AMOUNT_INDEX = DATA[0].indexOf("Amount");
  const TOTAL_INDEX = DATA[0].indexOf("Total");
  const PURCHASE_DATE_INDEX = DATA[0].indexOf("Purchase Date");
  let total = 0

  let found_change = false

  for (let col = 1; col < DATA.length; col++) {
    const ROW = DATA[col];

    if (ROW[DUE_DATE_INDEX] === "" || ROW[AMOUNT_INDEX] === "") {
        continue
    }
    else {
        found_change = true
    }

    if (ROW[PURCHASE_LOCATION_INDEX].toString().includes(PURCHASE_HEADER)) { continue; }

    if (col + 1 === DATA.length) {
        ROW[TOTAL_INDEX] = RoundUpToNearestDollar(String(total + ROW[AMOUNT_INDEX] + FindMultiWeekRepayment(ROW[DUE_DATE_INDEX])))
        ROW[PURCHASE_DATE_INDEX] = GetDateWhenCellEmpty(ROW[PURCHASE_DATE_INDEX])
        continue
    }

    const NEXT_PURCHASE_IS_HEADER = DATA[col + 1][PURCHASE_LOCATION_INDEX].toString().includes(PURCHASE_HEADER);
    const NEXT_DATE_NOT_MATCH = DATA[col + 1][DUE_DATE_INDEX].toString() !== ROW[DUE_DATE_INDEX].toString();

    
    if (NEXT_PURCHASE_IS_HEADER || NEXT_DATE_NOT_MATCH) {
        ROW[TOTAL_INDEX] = RoundUpToNearestDollar(String(total + ROW[AMOUNT_INDEX] + FindMultiWeekRepayment(ROW[DUE_DATE_INDEX])))
        total = 0
    }
    else {
        ROW[TOTAL_INDEX] = ""
        total += ROW[AMOUNT_INDEX]
    }
    ROW[PURCHASE_DATE_INDEX] = GetDateWhenCellEmpty(ROW[PURCHASE_DATE_INDEX])
    
  }


  if (found_change) { SHEET.getDataRange().setValues(DATA) }
}

function AddToFixed(num: number, add_val: number) {
  return ~~((num + add_val) * 100) / 100;
}

function SetDateToNextWeds(date: Date) {
  while (date.getUTCDay() !== 3) {
    date.setUTCDate(date.getUTCDate() + 1);
  }
  return date;
}

function SetDateToNextFri(year: number) {
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

function ComputeMonthlyIncome() {
  let cur_year = new Date().getUTCFullYear();
  const TAB_NAME = "Household Budget";
  let tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    `${TAB_NAME} ${cur_year}`
  );
  let start_date = new Date(`12/28/${cur_year - 1}`);
  const MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  const ROS_PAY_DAY = new PayDay(350.95, start_date, (_, __, ___) => {
    return true;
  });

  const MY_PAY_DAY = new PayDay(
    880.78,
    start_date,
    (_, total_days, day_inc) => {
      const SHOULD_PAY = total_days % (day_inc * 2) === 0;
      return SHOULD_PAY;
    }
  );

  ROS_PAY_DAY.SetPayoutDate(SetDateToNextWeds);
  MY_PAY_DAY.SetPayoutDate(SetDateToNextFri(cur_year - 1));

  while (Boolean(tab)) {
    let total = 0;
    const BUDGET_DATA = tab!
      .getRange(1, 1, 2, tab!.getLastColumn())
      .getValues();
    const BUDGET_MONTHS = BUDGET_DATA[1];
    const PRESENT_MONTHS = BUDGET_MONTHS.filter((x) => MONTHS.includes(x));

    let my_pay_day_month = PRESENT_MONTHS.includes(MY_PAY_DAY.PayMonth());
    let ros_pay_day_month = PRESENT_MONTHS.includes(ROS_PAY_DAY.PayMonth());
    while (!my_pay_day_month || !ros_pay_day_month) {
      if (!ros_pay_day_month) {
        ROS_PAY_DAY.PayOut();
      }
      if (!my_pay_day_month) {
        MY_PAY_DAY.PayOut();
      }
      ros_pay_day_month = PRESENT_MONTHS.includes(ROS_PAY_DAY.PayMonth());
      my_pay_day_month = PRESENT_MONTHS.includes(MY_PAY_DAY.PayMonth());
    }

    for (const MONTH of PRESENT_MONTHS) {
      while (
        ROS_PAY_DAY.PayMonth() === MONTH ||
        MY_PAY_DAY.PayMonth() === MONTH
      ) {
        if (ROS_PAY_DAY.PayMonth() === MONTH) {
          total += ROS_PAY_DAY.PayOut();
        }
        if (MY_PAY_DAY.PayMonth() === MONTH) {
          total += MY_PAY_DAY.PayOut();
        }
      }
      total = AddToFixed(total, 0);
      const MONTH_INDEX = BUDGET_MONTHS.indexOf(MONTH);
      BUDGET_DATA[0][MONTH_INDEX] = total;
      total = 0;
    }

    tab!.getRange(1, 1, 2, tab!.getLastColumn()).setValues(BUDGET_DATA);
    tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      `${TAB_NAME} ${++cur_year}`
    );
  }
}

function CreateDateString(date: Date | string, local: boolean = false) {
  if (typeof date === "string") {
    date = new Date(date);
  }

  let date_str = `${
    date.getUTCMonth() + 1
  }/${date.getUTCDate()}/${date.getUTCFullYear()}`;

  if (local) {
    date_str = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  }

  return date_str;
}

function GetFirstNonEmptyCellInCol(
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

function GenerateRepaymentSchedule() {
  const TAB_NAME = "Multi Week Loans";
  const TAB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAME);

  if (!TAB) {
    return;
  }

  const HEADERS = TAB.getRange(1, 1, 1, TAB.getLastColumn()).getValues()[0];
  const HEADER_COL_CACHE = new Map<string, string>();

  const GetCachedCol = (header: string) => {
    if (!HEADER_COL_CACHE.has(header)) {
      HEADER_COL_CACHE.set(header, IndexToColLetter(HEADERS.indexOf(header)));
    }
    return HEADER_COL_CACHE.get(header)!;
  }

  const RepaymentCol = (row: number) => `${GetCachedCol("Repayment Date")}${row}`;
  const AmountCol = (row: number) => `${GetCachedCol("Repayment Amount")}${row}`;
  const LoaneeCol = (row: number) => `${GetCachedCol("Loanee")}${row}`;
  const RoundUpCol = (row: number) => `${GetCachedCol("Round Up?")}${row}`;
  const NumOfPaymentsCol = (row: number) => `${GetCachedCol("Number of Repayments")}${row}`;
  const PurchaseDateCol = (row: number) => `${GetCachedCol("Purchase Date")}${row}`;

  const LAST_ROW = GetFirstNonEmptyCellInCol(TAB, () => HEADERS.indexOf("Number of Repayments"));
  const REPAYMENTS = Number(
    TAB.getRange(NumOfPaymentsCol(LAST_ROW)).getValue()
  );

  if (isNaN(REPAYMENTS) || REPAYMENTS === 0) {
    return;
  }

  const AMOUNT = TAB.getRange(AmountCol(LAST_ROW)).getValue() / REPAYMENTS;
  const LOANEE = TAB.getRange(LoaneeCol(LAST_ROW)).getValue();
  const SHOULD_ROUND_UP = TAB.getRange(RoundUpCol(LAST_ROW)).getValue();
  let repayment_date = new Date(
    TAB.getRange(RepaymentCol(LAST_ROW)).getValue()
  );

  let repayment_intraval = 7;
  if (LOANEE === "Dan") {
    repayment_intraval = 14;
  }

  TAB.getRange(NumOfPaymentsCol(LAST_ROW)).setValue("");
  TAB.getRange(RepaymentCol(LAST_ROW)).setValue("");
  const TODAY = new Date();

  for (let i = 0; i < REPAYMENTS; i++) {
    const PURCHASE_RANGE = TAB.getRange(PurchaseDateCol(LAST_ROW + i));
    TAB.getRange(RepaymentCol(LAST_ROW + i)).setValue(
      CreateDateString(repayment_date)
    );
    PURCHASE_RANGE.setValue(
      PURCHASE_RANGE.getValue() === "" ? CreateDateString(TODAY, true) : PURCHASE_RANGE.getValue()
    );
    TAB.getRange(AmountCol(LAST_ROW + i)).setValue(
      SHOULD_ROUND_UP === "Yes" ? Math.ceil(AMOUNT) : AddToFixed(AMOUNT, 0)
    );
    TAB.getRange(LoaneeCol(LAST_ROW + i)).setValue(LOANEE);
    repayment_date.setUTCDate(repayment_date.getUTCDate() + repayment_intraval);
  }
}

function CreateNewHouseholdBudgetTab() {
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
function CompareDates(date1: Date | string, date2: Date | string) {
  date1 = new Date(CreateDateString(date1))
  date2 = new Date(CreateDateString(date2))
  return date1.getTime() > date2.getTime()
}

function GroupByDate(date_header: string, tab_name: string, shade_red: boolean = true) {
    const TAB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab_name);

    if (!TAB) { return; } 

    const InsertGroupingRow = function(row: number, date: string, arr: any[][]) {
        const NEW_ROW = new Array(arr[0].length).fill("")
        NEW_ROW[PURCHASE_LOCATION_INDEX] = `${PURCHASE_HEADER} ${date}`
        NEW_ROW[CARD_INDEX] = " "
        arr.splice(row, 0, NEW_ROW)
    }

    let range_data = TAB.getDataRange().getValues();
    const DATE_COL_INDEX = range_data[0].indexOf(date_header);
    const PURCHASE_LOCATION_INDEX = range_data[0].indexOf("Purchase Location");
    const CARD_INDEX = range_data[0].indexOf("Card");

    const DATA = range_data.filter(row => !row[PURCHASE_LOCATION_INDEX].toString().includes(PURCHASE_HEADER))

    if (DATE_COL_INDEX === -1) { return; }

    InsertGroupingRow(1, CreateDateString(DATA[1][DATE_COL_INDEX]), DATA)


    const DATE_MAP = new Map<string, number[]>();
    let last_recorded_date = ""

    for (let col = 2; col < DATA.length; col++) {
        const row = DATA[col]
        const NEW_DATE = CreateDateString(row[DATE_COL_INDEX])

        if (row[DATE_COL_INDEX] === "") { continue }

        if (!DATE_MAP.has(NEW_DATE)) {
            DATE_MAP.set(NEW_DATE, [col+1])

            if (last_recorded_date === "") {
                last_recorded_date = NEW_DATE
            } else {
                DATE_MAP.get(last_recorded_date)!.push(col)
                InsertGroupingRow(col, CreateDateString(row[DATE_COL_INDEX]), DATA)
                last_recorded_date = NEW_DATE
            }
        }
    }
    TAB.getRange(1, 1, DATA.length, DATA[0].length).setValues(DATA)

    DATE_MAP.clear()
    let i = 0
    for (const ROW of DATA) {
        const DATE = new Date(ROW[DATE_COL_INDEX])
        if (DATE.toString() === "Invalid Date") { 
            i++
            continue 
        }

        const DATE_STR = CreateDateString(DATE)
        if (!DATE_MAP.has(DATE_STR)) {
            DATE_MAP.set(DATE_STR, [i, 0, DATE.getUTCFullYear(), DATE.getUTCMonth(), DATE.getUTCDate()])
            i++
            continue
        }

        DATE_MAP.get(DATE_STR)![1]++
        i++
    }

    i = 0
    const LIGHT_RED_SHADES = ["#FF7F7F", "#FF9F9F"]
    for(let [key, val] of DATE_MAP) {
        let start = val[0]+1
        let end = start + val[1] + 1
        const DUE_DATE = new Date(val[2], val[3], val[4])
        const CUR_DATE = new Date()

        const RANGE = TAB.getRange(start, 1, end-start, TAB.getLastColumn())
        const COLOR_RANGE = TAB.getRange(start-1, 1, end-start, TAB.getLastColumn())
        const DUE_DATE_PAST = CompareDates(CUR_DATE, DUE_DATE)
        
        if (DUE_DATE_PAST && shade_red) {
            COLOR_RANGE.setBackground(LIGHT_RED_SHADES[i++ % 2])
        }

        try {
            const GROUP = TAB.getRowGroup(start, 1)
            if (DUE_DATE_PAST) { 
                GROUP?.collapse() 
            }
            else {
                GROUP?.remove()
                RANGE.shiftRowGroupDepth(1)
            }
        }
        catch {
            RANGE.shiftRowGroupDepth(1)
        }
    }

}

function onEdit(_: unknown) {
  ComputeTotal();
  GenerateRepaymentSchedule();
}

function onOpen(_: unknown) {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addToUi();
}

function onOpenInstallable(_: unknown) {
  ComputeMonthlyIncome();
  GroupByDate("Due Date", "One Week Loans")
  GroupByDate("Purchase Date", "Multi Week Loans", false)
}
