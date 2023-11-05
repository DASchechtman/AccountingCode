function GroupByDate(
  date_header: string,
  tab_name: string,
  shade_red: boolean = true
) {
  const TAB = new GoogleSheetTabs(tab_name);
  const CURRENT_ONE_WEEK_TAB = new GoogleSheetTabs(ONE_WEEK_LOANS_TAB_NAME);
  const ROW_COMPARE = new Map<string, [number, number]>()

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

  const __CheckIfDateEntriesAdded = function (date: string) {
    const CACHED_DATA = __GetCachedOneWeekLoansData()
    const CACHE_INDEX = 0
    const CURRENT_INDEX = 1

    if (ROW_COMPARE.size > 0) {
      let comp = ROW_COMPARE.get(date)
      if (!comp) { return false }
      return comp[0] !== comp[1]
    }

    for (let i = 0; i < CACHED_DATA.length; i++) {
      __StoreCompResults(String(CACHED_DATA[i]![DATE_COL_INDEX]), CACHE_INDEX)
    }

    for (let i = 0; i < CURRENT_ONE_WEEK_TAB.NumberOfRows(); i++) {
      __StoreCompResults(String(CURRENT_ONE_WEEK_TAB.GetRow(i)![DATE_COL_INDEX]), CURRENT_INDEX)
    }

    let comp = ROW_COMPARE.get(date)
    if (!comp) { return false }
    return comp[0] !== comp[1]
  }

  const GenerateLoanGroupHeader = function () {
    let last_recorded_date = "";
    const FIRST_ROW_PAST_HEADERS = 1

    for (let i = FIRST_ROW_PAST_HEADERS; i < TAB.NumberOfRows(); i++) {
      const ROW = TAB.GetRow(i)!.map(x => String(x));

      if (ROW[PURCHASE_LOCATION_INDEX].includes(PURCHASE_HEADER)) {
        i = __GetNextDateGroup(ROW[PURCHASE_LOCATION_INDEX], i);
        continue;
      } else if (ROW[DATE_COL_INDEX] === "") {
        continue;
      }

      const NEW_DATE = __CreateDateString(ROW[DATE_COL_INDEX]);

      if (last_recorded_date === "" || last_recorded_date !== NEW_DATE) {
        last_recorded_date = NEW_DATE;
        TAB.InsertRow(i, [], { AlterRow: __InsertGroupingRow(NEW_DATE) });
      }
    }
  };

  const GetGroupBoundries = function () {
    const BOUNDRIES = new Map<string, any[]>();
    for (let i = 1; i < TAB.NumberOfRows(); i++) {
      const ROW = TAB.GetRow(i)
      if (!ROW) { continue }

      if (String(ROW[PURCHASE_LOCATION_INDEX]).includes(PURCHASE_HEADER)) {
        const DATE = String(ROW[PURCHASE_LOCATION_INDEX]).split(" ")[2];
        const ARR = [i + 1, 0, DATE]
        i = __GetNextDateGroup(String(ROW[PURCHASE_LOCATION_INDEX]), i);
        ARR[1] = i - (ARR[0] as number) + 1;
        BOUNDRIES.set(DATE, ARR);
        continue
      }
    }

    return BOUNDRIES
  }

  const GroupRowsInSheet = function () {
    const LIGHT_RED_SHADES = ["#FF7F7F", "#FF9F9F"]
    let i = 0

    BOUNDRIES.forEach((val, key) => {
      const DUE_DATE = new Date(val[2])
      const CUR_DATE = new Date()
      const DUE_DATE_HAS_PASSED = __CompareDates(CUR_DATE, DUE_DATE)
      const GROUP_RANGE = TAB.GetTab().getRange(val[0] + 1, 1, val[1], TAB.GetTab().getLastColumn())
      const COLOR_RANGE = TAB.GetTab().getRange(val[0], 1, val[1] + 1, TAB.GetTab().getLastColumn())

      if (DUE_DATE_HAS_PASSED && shade_red) {
        COLOR_RANGE.setBackground(LIGHT_RED_SHADES[i++ % LIGHT_RED_SHADES.length])
      }

      try {
        let GROUP = TAB.GetTab().getRowGroup(val[0], 1)
        if (__CheckIfDateEntriesAdded(__CreateDateString(DUE_DATE))) {
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
    })
  }


  const DATE_COL_INDEX = TAB.GetHeaderIndex(date_header);
  const PURCHASE_LOCATION_INDEX = TAB.GetHeaderIndex("Purchase Location");
  const CARD_INDEX = TAB.GetHeaderIndex("Card");

  if (DATE_COL_INDEX === -1) {
    return;
  }

  GenerateLoanGroupHeader()
  const BOUNDRIES = GetGroupBoundries()
  GroupRowsInSheet()

  TAB.SaveToTab();
}

function ComputeTotal() {
  const TAB_NAME = "One Week Loans";
  const SHEET = new GoogleSheetTabs(TAB_NAME);

  const PURCHASE_COL_HEADER = "Purchase Location";
  const DUE_DATE_COL_HEADER = "Due Date";
  const AMOUNT_COL_HEADER = "Amount";
  const TOTAL_COL_HEADER = "Total";
  const PURCHASE_DATE_COL_HEADER = "Purchase Date";

  const PURCHASE_LOCATION_INDEX = SHEET.GetCol(PURCHASE_COL_HEADER)?.map(x => String(x));
  const DUE_DATE_INDEX = SHEET.GetCol(DUE_DATE_COL_HEADER)?.map(x => String(x));
  const AMOUNT_INDEX = SHEET.GetCol(AMOUNT_COL_HEADER)
  const TOTAL_INDEX = SHEET.GetCol(TOTAL_COL_HEADER);
  const PURCHASE_DATE_INDEX = SHEET.GetCol(PURCHASE_DATE_COL_HEADER)?.map(x => String(x));

  if (!PURCHASE_LOCATION_INDEX || !DUE_DATE_INDEX || !AMOUNT_INDEX || !TOTAL_INDEX || !PURCHASE_DATE_INDEX) {
    return
  }

  let total = 0;
  let last_amt = 0
  let last_recorded_date = "";

  for (let i = 1; i < SHEET.NumberOfRows(); i++) {
    const PURCHASE_LOCATION = PURCHASE_LOCATION_INDEX[i];
    const DUE_DATE = DUE_DATE_INDEX[i];
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
        TOTAL_INDEX[last_amt] = __AddToFixed(total, 0, Math.ceil)
        TOTAL_INDEX[i] = AMOUNT
      }
      else {
        TOTAL_INDEX[i] = __AddToFixed(total, AMOUNT, Math.ceil)
      }
    }
    else if (last_recorded_date === "" || last_recorded_date !== DUE_DATE) {
      if (last_recorded_date !== "") {
        TOTAL_INDEX[last_amt] = __AddToFixed(total, 0, Math.ceil)
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

    PURCHASE_DATE_INDEX[i] = __GetDateWhenCellEmpty(PURCHASE_DATE_INDEX[i]);
  }

  SHEET.WriteCol(PURCHASE_COL_HEADER, PURCHASE_LOCATION_INDEX)
  SHEET.WriteCol(DUE_DATE_COL_HEADER, DUE_DATE_INDEX)
  SHEET.WriteCol(AMOUNT_COL_HEADER, AMOUNT_INDEX)
  SHEET.WriteCol(TOTAL_COL_HEADER, TOTAL_INDEX.map(x => typeof x === "number" ? __AddToFixed(x, 0, Math.ceil) : x))
  SHEET.WriteCol(PURCHASE_DATE_COL_HEADER, PURCHASE_DATE_INDEX)
  SHEET.SaveToTab();
}

function GenerateRepaymentSchedule() {
  let generated_repayment_schedule = false
  const TAB_NAME = "Multi Week Loans";
  const TAB = new GoogleSheetTabs(TAB_NAME);

  const NUM_OF_REPAYMENT_COL_NAME = "Number of Repayments";
  const LOANEE_COL_NAME = "Loanee";
  const REPAYMENT_AMT_COL_NAME = "Repayment Amount";
  const PURCHASE_COL_NAME = "Purchase Date";
  const ROUND_UP_COL_NAME = "Round Up?";
  const REPAYMENT_DATE_COL_NAME = "Repayment Date";

  const NUM_OF_PAYMENTS_COL = TAB.GetCol(NUM_OF_REPAYMENT_COL_NAME)
  const LOANEE_COL = TAB.GetCol(LOANEE_COL_NAME)
  const REPAYMENT_COL = TAB.GetCol(REPAYMENT_AMT_COL_NAME)
  const PURCHASE_COL = TAB.GetCol(PURCHASE_COL_NAME)
  const ROUND_UP_COL = TAB.GetCol(ROUND_UP_COL_NAME)
  const REPAYMENT_DATE_COL = TAB.GetCol(REPAYMENT_DATE_COL_NAME)
  let last_row_index = 0

  if (!NUM_OF_PAYMENTS_COL || !LOANEE_COL || !REPAYMENT_COL || !PURCHASE_COL || !ROUND_UP_COL || !REPAYMENT_DATE_COL) { return generated_repayment_schedule }

  const LAST_ROW = NUM_OF_PAYMENTS_COL.find(cell => {
    const cell_num = Number(cell)
    return !isNaN(cell_num) && cell_num > 0
  })

  if (LAST_ROW === undefined) { return generated_repayment_schedule }
  generated_repayment_schedule = true

  last_row_index = NUM_OF_PAYMENTS_COL.indexOf(LAST_ROW)

  const NUM_OF_REPAYMENTS = Number(NUM_OF_PAYMENTS_COL[last_row_index])
  const PURCHASE_DATE = new Date()
  const LOANEE = LOANEE_COL[last_row_index]
  let installment = Number(REPAYMENT_COL[last_row_index]) / NUM_OF_REPAYMENTS
  let payment_days = LOANEE === "Dan" ? 14 : 7
  let payment_start_date = new Date(REPAYMENT_DATE_COL[last_row_index])

  if (ROUND_UP_COL[last_row_index] === "Yes") {
    installment = Math.ceil(installment)
  }

  for (let i = last_row_index; i < last_row_index + NUM_OF_REPAYMENTS; i++) {
    if (i >= NUM_OF_PAYMENTS_COL.length) { NUM_OF_PAYMENTS_COL.push("") }
    if (i >= LOANEE_COL.length) { LOANEE_COL.push("") }
    if (i >= REPAYMENT_COL.length) { REPAYMENT_COL.push("") }
    if (i >= PURCHASE_COL.length) { PURCHASE_COL.push("") }
    if (i >= ROUND_UP_COL.length) { ROUND_UP_COL.push("") }
    if (i >= REPAYMENT_DATE_COL.length) { REPAYMENT_DATE_COL.push("") }

    NUM_OF_PAYMENTS_COL[i] = ""
    LOANEE_COL[i] = LOANEE
    REPAYMENT_COL[i] = installment
    PURCHASE_COL[i] = __CreateDateString(PURCHASE_DATE)
    REPAYMENT_DATE_COL[i] = __CreateDateString(payment_start_date)

    payment_start_date.setDate(payment_start_date.getDate() + payment_days)
  }

  TAB.WriteCol(NUM_OF_REPAYMENT_COL_NAME, NUM_OF_PAYMENTS_COL.map(cell => cell === "Number of Repayments" ? cell : ""))
  TAB.WriteCol(LOANEE_COL_NAME, LOANEE_COL)
  TAB.WriteCol(REPAYMENT_AMT_COL_NAME, REPAYMENT_COL)
  TAB.WriteCol(PURCHASE_COL_NAME, PURCHASE_COL)
  TAB.WriteCol(ROUND_UP_COL_NAME, ROUND_UP_COL)
  TAB.WriteCol(REPAYMENT_DATE_COL_NAME, REPAYMENT_DATE_COL)
  TAB.SaveToTab()

  AddMultiWeekLoanToRepayment(last_row_index)
  return generated_repayment_schedule
}

function AddMultiWeekLoanToRepayment(start_row: number) {
  const ONE_WEEK_TAB = new GoogleSheetTabs("One Week Loans");
  const MULTI_WEEK_TAB = new GoogleSheetTabs("Multi Week Loans");

  const MULTI_COL_INDEXES = [
    MULTI_WEEK_TAB.GetHeaderIndex("Repayment Date"),
    MULTI_WEEK_TAB.GetHeaderIndex("Purchase Date"),
    MULTI_WEEK_TAB.GetHeaderIndex("Repayment Amount"),
    MULTI_WEEK_TAB.GetHeaderIndex("Purchase Location"),
    MULTI_WEEK_TAB.GetHeaderIndex("Card")
  ]

  const [
    MULTI_TAB_DUE_DATE_COL_INDEX,
    MULTI_TAB_PURCHASE_DATE_COL_INDEX,
    MULTI_TAB_PAYMENT_AMT_COL_INDEX,
    MULTI_TAB_PURCHASE_LOCATION_COL_INDEX,
    MULTI_TAB_CARD_COL_INDEX
  ] = MULTI_COL_INDEXES
  if (MULTI_COL_INDEXES.includes(-1)) { return }

  const WEEKLY_COL_INDEXES = [
    ONE_WEEK_TAB.GetHeaderIndex("Due Date"),
    ONE_WEEK_TAB.GetHeaderIndex("Purchase Date"),
    ONE_WEEK_TAB.GetHeaderIndex("Amount"),
    ONE_WEEK_TAB.GetHeaderIndex("Purchase Location"),
    ONE_WEEK_TAB.GetHeaderIndex("Card")
  ]

  const [
    WEEKLY_TAB_DUE_DATE_COL_INDEX,
    WEEKLY_TAB_PURCHASE_DATE_COL_INDEX,
    WEEKLY_TAB_PAYMENT_AMT_COL_INDEX,
    WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX,
    WEEKLY_TAB_CARD_COL_INDEX
  ] = WEEKLY_COL_INDEXES
  if (WEEKLY_COL_INDEXES.includes(-1)) { return }

  const __GetDateIndexBoundries = function (date: string): [number, number] {
    let i = 0
    const ROW = ONE_WEEK_TAB.FindRow(row => {
      const FOUND = row[WEEKLY_TAB_DUE_DATE_COL_INDEX] === date
      i += Number(!FOUND)
      return FOUND
    })

    if (!ROW) { return [-1, -1] }

    let ret: [number, number] = [i, 0]

    while (true) {
      const ROW = ONE_WEEK_TAB.GetRow(i)
      if (!ROW) { break }
      if (ROW[WEEKLY_TAB_DUE_DATE_COL_INDEX] !== date) { break }
      i++
    }

    ret[1] = i - 1
    return ret
  }

  const __HasMultiWeekRepayment = function (begin: number, end: number, purchase_desc: string) {
    let has_repayment = false
    for (let i = begin; i <= end; i++) {
      const ROW = ONE_WEEK_TAB.GetRow(i)
      if (!ROW) { continue }
      if (ROW[WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX] === purchase_desc) {
        has_repayment = true
        break
      }
    }
    return has_repayment
  }

  let purchase_desc = ""
  let credit_card_name = ""

  for (let i = start_row; i < MULTI_WEEK_TAB.NumberOfRows(); i++) {
    const ROW = MULTI_WEEK_TAB.GetRow(i)
    if (!ROW) { continue }

    const DUE_DATE = String(ROW[MULTI_TAB_DUE_DATE_COL_INDEX])
    if (DUE_DATE === "") { continue }

    if (ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX] !== "") { purchase_desc = String(ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX]) }
    if (ROW[MULTI_TAB_CARD_COL_INDEX] !== "") { credit_card_name = String(ROW[MULTI_TAB_CARD_COL_INDEX]) }

    const [START_INDEX, END_INDEX] = __GetDateIndexBoundries(DUE_DATE)
    if (__HasMultiWeekRepayment(START_INDEX, END_INDEX, purchase_desc)) { continue }

    const NEW_ROW: DataArrayEntry = []
    NEW_ROW[WEEKLY_TAB_DUE_DATE_COL_INDEX] = DUE_DATE
    NEW_ROW[WEEKLY_TAB_PURCHASE_DATE_COL_INDEX] = ROW[MULTI_TAB_PURCHASE_DATE_COL_INDEX]
    NEW_ROW[WEEKLY_TAB_PAYMENT_AMT_COL_INDEX] = ROW[MULTI_TAB_PAYMENT_AMT_COL_INDEX]
    NEW_ROW[WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX] = purchase_desc
    NEW_ROW[WEEKLY_TAB_CARD_COL_INDEX] = credit_card_name

    if (END_INDEX > -1) {
      const IMMEDIATELY_AFTER_GROUP = END_INDEX + 1
      ONE_WEEK_TAB.InsertRow(IMMEDIATELY_AFTER_GROUP, NEW_ROW)
    } else {
      ONE_WEEK_TAB.AppendRow(NEW_ROW)
    }
  }

  ONE_WEEK_TAB.SaveToTab()
}

function ComputeMonthlyIncome() {
  const TAB_NAME = "Household Budget";
  const START_CELL_INDEX = 1
  let cur_year = new Date().getUTCFullYear();
  const LAST_YEAR = cur_year - 1
  let start_date = new Date(`12/28/${LAST_YEAR}`);

  const __RosPayDay = function () {
    return true
  }

  const __DansPayDay = function ({ total_days, inc }: PayOutParams) {
    const SHOULD_PAY = total_days % (inc * 2) === 0;
    return SHOULD_PAY;
  }

  const ROS_PAY_DAY = new PayDay(350.95, start_date, __RosPayDay);
  ROS_PAY_DAY.SetPayoutDate(__SetDateToNextWeds);

  const MY_PAY_DAY = new PayDay(880.78, start_date, __DansPayDay);
  MY_PAY_DAY.SetPayoutDate(__SetDateToNextFri(LAST_YEAR));

  const __SetCellToFormula = function (row: DataArrayEntry, cell_index: number, data: number) {
    row[cell_index] = `${data}`
    if (cell_index > 1) {
      const COL_LETTER = __IndexToColLetter(cell_index - 2)
      row[cell_index] += `+IF(${COL_LETTER}24>0,${COL_LETTER}24,0)`
    }
    else {
      try {
        const LAST_SHEET_NAME = `${TAB_NAME} ${cur_year - 1}`
        const PREV_TAB = new GoogleSheetTabs(LAST_SHEET_NAME);
        const REMAINING_BUDGET_ROW = PREV_TAB.FindRow(row => row[0] === "Estimated Savings:")
        if (REMAINING_BUDGET_ROW) {
          row[cell_index] += `+'${LAST_SHEET_NAME}'!${__IndexToColLetter(REMAINING_BUDGET_ROW.length - 2)}24`
        }
      } catch { }
    }

    row[cell_index] = `=MIN(${row[cell_index]},5000)`
  }

  while (true) {
    try {
      const TAB = new GoogleSheetTabs(`${TAB_NAME} ${cur_year}`);
      const INCOME_ROW = TAB.GetRow(0)?.map(x => {
        if (typeof x === "number") { return Number(0) }
        return x
      });
      const MONTH_ROW = TAB.GetRow(1)?.map(x => String(x));
      if (!INCOME_ROW || !MONTH_ROW) { break }

      while (MONTH_ROW[START_CELL_INDEX] !== ROS_PAY_DAY.PayMonth() || MONTH_ROW[START_CELL_INDEX] !== MY_PAY_DAY.PayMonth()) {
        const ROS_PAY_MONTH = ROS_PAY_DAY.PayMonth()
        const MY_PAY_MONTH = MY_PAY_DAY.PayMonth()
        if (MONTH_ROW[START_CELL_INDEX] !== ROS_PAY_MONTH) { ROS_PAY_DAY.PayOut() }
        if (MONTH_ROW[START_CELL_INDEX] !== MY_PAY_MONTH) { MY_PAY_DAY.PayOut() }
      }

      for (let i = START_CELL_INDEX; i < MONTH_ROW.length; i++) {
        if (MONTH_ROW[i] === "") { continue }

        let total = 0

        while (ROS_PAY_DAY.PayMonth() === MONTH_ROW[i] || MY_PAY_DAY.PayMonth() === MONTH_ROW[i]) {
          if (ROS_PAY_DAY.PayMonth() === MONTH_ROW[i]) {
            total = __AddToFixed(total, ROS_PAY_DAY.PayOut())
          }
          if (MY_PAY_DAY.PayMonth() === MONTH_ROW[i]) {
            total = __AddToFixed(total, MY_PAY_DAY.PayOut())
          }
        }

        __SetCellToFormula(INCOME_ROW, i, total)
      }

      TAB.WriteRow(0, INCOME_ROW)
      TAB.SaveToTab()
      cur_year++
    } catch {
      break
    }
  }
}

function AddIncomeRow() {
  const UI = SpreadsheetApp.getUi()
  const YEAR_INPUT = UI.prompt("What year is this income for?", UI.ButtonSet.OK_CANCEL).getResponseText()
  const INCOME_LABEL = UI.prompt("What is the income label?", UI.ButtonSet.OK_CANCEL).getResponseText()

  const BUDGET_PLANNER_TAB = new GoogleSheetTabs(BUDGET_PLANNER_TAB_NAME)
  const JAN_LABEL_COL = BUDGET_PLANNER_TAB.GetHeaderIndex("January")

  const INCOME_YEAR = Number(YEAR_INPUT)
  const INCOME_PER_MONTH_COL = 1
  const INCOME_STREAM_COL = 1
  const INCOME_YEAR_COL = 0

  const IncomePerMonthRow = () => BUDGET_PLANNER_TAB.IndexOfRow(row => row[INCOME_PER_MONTH_COL] === "Income Per Paycheck")

  const IncomeStreamRow = () => BUDGET_PLANNER_TAB.IndexOfRow(row => row[INCOME_STREAM_COL] === "Income Stream")

  const __AddValidationToRow = function () {
    const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    const VALIDATION_OPTIONS = [...WEEKDAYS.map(day => PAYMENT_SCHEDULE.map(pay => `${day} : ${pay}`)).flat(), "Custom", "N/A"]
    const VALIDATION = SpreadsheetApp
      .newDataValidation()
      .requireValueInList(VALIDATION_OPTIONS)
      .build()

    for (let i = IncomePerMonthRow(); i <= IncomeStreamRow(); i++) {
      BUDGET_PLANNER_TAB.GetRowRange(i)?.setDataValidation(null)
    }

    for (let j = IncomeStreamRow() + 1; j < BUDGET_PLANNER_TAB.NumberOfRows(); j++) {
      const ROW = BUDGET_PLANNER_TAB.GetRow(j)!
      for (let i = JAN_LABEL_COL - 1; i < ROW.length; i += 2) {
        BUDGET_PLANNER_TAB.GetTab().getRange(j + 1, i + 1).setDataValidation(VALIDATION)
        if (ROW[i] !== "") { continue }
        ROW[i] = VALIDATION_OPTIONS[VALIDATION_OPTIONS.length - 1]
        ROW[i + 1] = "-"
      }
      BUDGET_PLANNER_TAB.WriteRow(j, ROW)
    }
  }

  const __ValidateUserInput = function () {
    if (!isNaN(INCOME_YEAR) && INCOME_LABEL !== "") { return true }
    let alert = ""
    if (isNaN(INCOME_YEAR)) { alert = `Please enter a valid year instead of '${YEAR_INPUT}'` }
    else if (INCOME_LABEL === "") { alert = "Please enter a valid income label (anything but an empty string)" }
    UI.alert(alert)
    return false
  }

  const __InsertRow = function (row: DataArrayEntry, row_index: number) {
    if (row[INCOME_PER_MONTH_COL] === "") {
      BUDGET_PLANNER_TAB.InsertRow(row_index, [INCOME_YEAR, INCOME_LABEL], { should_fill: true })
      return true
    }
    if (Number(row[INCOME_YEAR_COL]) <= INCOME_YEAR) { return false }
    BUDGET_PLANNER_TAB.InsertRow(row_index, [INCOME_YEAR, INCOME_LABEL], { should_fill: true })
    return true
  }

  if (!__ValidateUserInput() || IncomePerMonthRow() === -1 || IncomeStreamRow() === -1) { return }

  const STOP = IncomeStreamRow()
  for (let i = IncomePerMonthRow() + 1; i < STOP; i++) {
    if (__InsertRow(BUDGET_PLANNER_TAB.GetRow(i)!, i)) { break }
  }

  const ShouldLoop = (i: number) => i < BUDGET_PLANNER_TAB.NumberOfRows()
  const START = IncomeStreamRow() + 1
  let inserted = false
  for (let i = START; ShouldLoop(i); i++) {
    inserted = __InsertRow(BUDGET_PLANNER_TAB.GetRow(i)!, i)
    if (inserted) { break }
  }

  if (!ShouldLoop(START) || !inserted) {
    BUDGET_PLANNER_TAB.AppendRow([INCOME_YEAR, INCOME_LABEL], true)
  }

  __AddValidationToRow()


  BUDGET_PLANNER_TAB.SaveToTab()
}

function ComputeIncomeForEachMonth() {
  const ALL_YEARS = 'ALL-YEARS'
  const BUDGET_TAB = new GoogleSheetTabs(BUDGET_PLANNER_TAB_NAME)
  const HEADERS = BUDGET_TAB.GetRow(0)!

  const MON = "Monday"
  const TUE = "Tuesday"
  const WED = "Wednesday"
  const THU = "Thursday"
  const FRI = "Friday"

  const WEEKLY = "Weekly"
  const BI_WEEKLY = "Bi-Weekly"
  const SEMI_MONTHLY = "Semi-Monthly"
  const MONTHLY = "Monthly"

  const INCOME_PER_MONTH_ROW = BUDGET_TAB.IndexOfRow(row => row[1] === "Income Per Paycheck")
  const INCOME_STREAM_ROW = BUDGET_TAB.IndexOfRow(row => row[1] === "Income Stream")
  const JAN_COL = BUDGET_TAB.GetHeaderIndex("January")

  let day = ""
  let pay_schedule = ""
  let pay = 0

  const __SetDate = function (day: string) {
    let day_code = -1
    switch (day) {
      case MON: { day_code = 1; break }
      case TUE: { day_code = 2; break }
      case WED: { day_code = 3; break }
      case THU: { day_code = 4; break }
      case FRI: { day_code = 5; break }
    }

    return function (date: Date) {
      while (date.getUTCDay() !== day_code) {
        date.setDate(date.getDate() + 1)
      }
      return date
    }
  }

  const __SetPayoutCheck = function (schedule: string) {
    switch (schedule) {
      case WEEKLY: { return () => true }
      case BI_WEEKLY: {
        return function (params: PayOutParams) {
          return params.total_days % (params.inc * 2) === 0
        }
      }
      case SEMI_MONTHLY: {
        let month = ""
        let count = 0
        return function (params: PayOutParams) {
          if (month !== params.pay_month) {
            month = params.pay_month
            count = 0
          }
          return count++ < 2
        }
      }
      case MONTHLY: {
        let month = ""
        let paid = false
        return function (params: PayOutParams) {
          if (month !== params.pay_month) {
            month = params.pay_month
            paid = false
          }
          const PAID = paid
          paid = true
          return !PAID
        }
      }
    }

    return () => true
  }

  for (let i = INCOME_STREAM_ROW + 1; i < BUDGET_TAB.NumberOfRows(); i++) {
    const ROW = BUDGET_TAB.GetRow(i)!
    const PAY_ROW_INDEX = BUDGET_TAB.IndexOfRow(row => row[1] === ROW[1] && row[0] === ROW[0], INCOME_PER_MONTH_ROW)
    const PAY_ROW = BUDGET_TAB.GetRow(PAY_ROW_INDEX)!
    const PAY = new PayDay(0, new Date(`1/1/${ROW[0]}`), () => true)
    for (let j = JAN_COL - 1; j < ROW.length; j += 2) {
      if (pay === 0 || PAY_ROW[j + 1] !== "") {
        pay = Number(PAY_ROW[j + 1])
        PAY.SetPayoutAmount(pay)
      }

      if (ROW[j] === "Custom") {
        PAY.SetPayoutCheck(__SetPayoutCheck(MONTHLY))
      }
      else if (ROW[j] !== "N/A") {
        [day, pay_schedule] = String(ROW[j]).split(" : ")
        PAY.SetPayoutDate(__SetDate(day))
        PAY.SetPayoutCheck(__SetPayoutCheck(pay_schedule))
      }

      let total = 0
      while (PAY.PayMonth() === HEADERS[j + 1]) {
        total = __AddToFixed(total, PAY.PayOut())
      }
      ROW[j + 1] = total
    }
    BUDGET_TAB.WriteRow(i, ROW)
  }

  BUDGET_TAB.SaveToTab()
}

function onEdit(e: unknown) {
  if (!__EventObjectIsEditEventObject(e)) { return }
  const TAB_NAME = e.range.getSheet().getName();

  switch (TAB_NAME) {
    case MULTI_WEEK_LOANS_TAB_NAME: {
      const GENERATED = GenerateRepaymentSchedule()
      if (GENERATED) {
        GroupByDate("Purchase Date", MULTI_WEEK_LOANS_TAB_NAME, false);
        GroupByDate("Due Date", ONE_WEEK_LOANS_TAB_NAME);
        ComputeTotal();
      }
      break
    }
    case ONE_WEEK_LOANS_TAB_NAME: {
      ComputeTotal()
      break
    }
    case BUDGET_PLANNER_TAB_NAME: {
      break
    }
  }
}

function onOpen(_: SpreadSheetOpenEventObject) {
  ComputeMonthlyIncome();
  GroupByDate("Due Date", ONE_WEEK_LOANS_TAB_NAME);
  const x = new Map()
  Console.Log(x.size, 1, 2, 3)

  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "__CreateNewHouseholdBudgetTab")
    .addItem("Compute One Week Loans", "ComputeTotal")
    .addItem("Add Income to Planner", "AddIncomeRow")
    .addItem("Generate Income Schedule", "ComputeIncomeForEachMonth")
    .addToUi();
  __CacheOneWeekLoans()
}