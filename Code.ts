function GroupByDate(
  date_header: string,
  tab_name: string,
  shade_red: boolean = true
) {
  const TAB = new GoogleSheetTabs(tab_name);

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
        TAB.InsertRow(i, [], __InsertGroupingRow(NEW_DATE));
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
    console.log(BOUNDRIES.size)
    BOUNDRIES.forEach((val, key) => {
      const DUE_DATE = new Date(val[2])
      const CUR_DATE = new Date()
      const DUE_DATE_HAS_PASSED = __CompareDates(CUR_DATE, DUE_DATE)
      const GROUP_RANGE = TAB.GetTab().getRange(val[0]+1, 1, val[1], TAB.GetTab().getLastColumn())
      const COLOR_RANGE = TAB.GetTab().getRange(val[0], 1, val[1]+1, TAB.GetTab().getLastColumn())

      if (DUE_DATE_HAS_PASSED && shade_red) {
        COLOR_RANGE.setBackground(LIGHT_RED_SHADES[i++ % LIGHT_RED_SHADES.length])
      }

      try {
        let GROUP = TAB.GetTab().getRowGroup(val[0], 1)
        GROUP?.remove()
        GROUP_RANGE.shiftRowGroupDepth(1)
        GROUP = TAB.GetTab().getRowGroup(val[0], 1)
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
  const TAB_NAME = "Multi Week Loans";
  const TAB = new GoogleSheetTabs(TAB_NAME);
  const NUM_OF_PAYMENTS_COL = TAB.GetCol("Number of Repayments")
  const LOANEE_COL = TAB.GetCol("Loanee")
  const REPAYMENT_COL = TAB.GetCol("Repayment Amount")
  const PURCHASE_COL = TAB.GetCol("Purchase Date")
  const ROUND_UP_COL = TAB.GetCol("Round Up?")
  const REPAYMENT_DATE_COL = TAB.GetCol("Repayment Date")
  let last_row_index = 0

  if (!NUM_OF_PAYMENTS_COL || !LOANEE_COL || !REPAYMENT_COL || !PURCHASE_COL || !ROUND_UP_COL || !REPAYMENT_DATE_COL) { return }

  const LAST_ROW = NUM_OF_PAYMENTS_COL.find(cell => {
    const cell_num = Number(cell)
    return !isNaN(cell_num) && cell_num > 0
  })

  if (LAST_ROW === undefined) { return }

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

  TAB.WriteCol("Number of Repayments", NUM_OF_PAYMENTS_COL.map(cell => cell === "Number of Repayments" ? cell : ""))
  TAB.WriteCol("Loanee", LOANEE_COL)
  TAB.WriteCol("Repayment Amount", REPAYMENT_COL)
  TAB.WriteCol("Purchase Date", PURCHASE_COL)
  TAB.WriteCol("Round Up?", ROUND_UP_COL)
  TAB.WriteCol("Repayment Date", REPAYMENT_DATE_COL)
  TAB.SaveToTab()

  AddMultiWeekLoanToRepayment()
}

function AddMultiWeekLoanToRepayment() {
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

    while(true) {
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

  for(let i = 1; i < MULTI_WEEK_TAB.NumberOfRows(); i++) {
    const ROW = MULTI_WEEK_TAB.GetRow(i)
    if (!ROW) { continue }

    const DUE_DATE = String(ROW[MULTI_TAB_DUE_DATE_COL_INDEX])
    if (DUE_DATE === "") { continue }

    if (ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX] !== "" ) { purchase_desc = String(ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX]) }
    if (ROW[MULTI_TAB_CARD_COL_INDEX] !== "" ) { credit_card_name = String(ROW[MULTI_TAB_CARD_COL_INDEX]) }

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
  let cur_year = new Date().getUTCFullYear();
  const TAB_NAME = "Household Budget";
  let start_date = new Date(`12/28/${cur_year - 1}`);

  const __RosPayDay = function (date: Date, total_days: number, inc: number) {
    return true
  }

  const __DansPayDay = function (date: Date, total_days: number, inc: number) {
    const SHOULD_PAY =  total_days % (inc * 2) === 0;
    return SHOULD_PAY;
  }

  const ROS_PAY_DAY = new PayDay(350.95, start_date, __RosPayDay);
  ROS_PAY_DAY.SetPayoutDate(__SetDateToNextWeds);

  const MY_PAY_DAY = new PayDay(880.78, start_date, __DansPayDay);
  MY_PAY_DAY.SetPayoutDate(__SetDateToNextFri(cur_year - 1));

  while (true) {
    try {
      const TAB = new GoogleSheetTabs(`${TAB_NAME} ${cur_year}`);
      const INCOME_ROW = TAB.GetRow(0)?.map(x => {
        if (typeof x === "number") { return Number(0) }
        return x
      });
      const MONTH_ROW = TAB.GetRow(1)?.map(x => String(x));
      if (!INCOME_ROW || !MONTH_ROW) { break }

      while (MONTH_ROW[1] !== ROS_PAY_DAY.PayMonth() || MONTH_ROW[1] !== MY_PAY_DAY.PayMonth()) {
        const MONTH_1 = ROS_PAY_DAY.PayMonth()
        const MONTH_2 = MY_PAY_DAY.PayMonth()
        if (MONTH_ROW[1] !== MONTH_1) { ROS_PAY_DAY.PayOut() }
        if (MONTH_ROW[1] !== MONTH_2) { MY_PAY_DAY.PayOut() }
      }

      for (let i = 1; i < MONTH_ROW.length; i++) {
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

        INCOME_ROW[i] = total
      }

      TAB.WriteRow(0, INCOME_ROW)
      TAB.SaveToTab()
      cur_year++
    } catch {
      break
    }
  }
}

function onEdit(_: unknown) {
  GenerateRepaymentSchedule();
  ComputeTotal();
}

function onOpen(_: unknown) {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Compute One Week Loans", "ComputeTotal")
    .addToUi();

  AddMultiWeekLoanToRepayment()
  ComputeMonthlyIncome();
  GroupByDate("Due Date", "One Week Loans");
  GroupByDate("Purchase Date", "Multi Week Loans", false);
  ComputeTotal();
}