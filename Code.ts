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
        const GROUP = TAB.GetTab().getRowGroup(val[0], 1)
        if (DUE_DATE_HAS_PASSED) {
          GROUP?.collapse()
        }
        else {
          GROUP?.remove()
          GROUP_RANGE.shiftRowGroupDepth(1)
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
        TOTAL_INDEX[last_amt] = __AddToFixed(total, __FindMultiWeekRepayment(last_recorded_date), Math.ceil)
        TOTAL_INDEX[i] = AMOUNT
      }
      else {
        TOTAL_INDEX[i] = __AddToFixed(__AddToFixed(total, AMOUNT), __FindMultiWeekRepayment(DUE_DATE), Math.ceil)
      }
    }
    else if (last_recorded_date === "" || last_recorded_date !== DUE_DATE) {
      if (last_recorded_date !== "") { 
        TOTAL_INDEX[last_amt] = __AddToFixed(total, __FindMultiWeekRepayment(last_recorded_date), Math.ceil)
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

  ROS_PAY_DAY.SetPayoutDate(__SetDateToNextWeds);
  MY_PAY_DAY.SetPayoutDate(__SetDateToNextFri(cur_year - 1));

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
      total = __AddToFixed(total, 0);
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

function AddMultiWeekLoanToRepayment() {
  const ONE_WEEK_TAB = new GoogleSheetTabs("One Week Loans");
  const MULTI_WEEK_TAB = new GoogleSheetTabs("Multi Week Loans");
  const DATE_MAP = new Map<string, number>();

  const MULTI_DUE_DATE_COL = MULTI_WEEK_TAB.GetHeaderIndex("Repayment Date");
  const MULTI_PURCHASE_DATE_COL = MULTI_WEEK_TAB.GetHeaderIndex("Purchase Date");
  const MULTI_PAYMENT_AMT_COL = MULTI_WEEK_TAB.GetHeaderIndex("Repayment Amount");
  if(!MULTI_DUE_DATE_COL || !MULTI_PURCHASE_DATE_COL || !MULTI_PAYMENT_AMT_COL) { return }

  const ONE_DUE_DATE_COL = ONE_WEEK_TAB.GetHeaderIndex("Due Date");
  const ONE_PURCHASE_DATE_COL = ONE_WEEK_TAB.GetHeaderIndex("Purchase Date");
  const ONE_PAYMENT_AMT_COL = ONE_WEEK_TAB.GetHeaderIndex("Amount");
  if(!ONE_DUE_DATE_COL || !ONE_PURCHASE_DATE_COL || !ONE_PAYMENT_AMT_COL) { return }

  const __GetDateIndex = function (date: string) {
    let i = 0
    const ROW = ONE_WEEK_TAB.FindRow(row => {
      i++
      return row[ONE_DUE_DATE_COL] === date
    })

    if (!ROW) { return -1 }
    return i
  }

  for(let i = 1; i < MULTI_WEEK_TAB.NumberOfRows(); i++) {
    const ROW = MULTI_WEEK_TAB.GetRow(i)
    if (!ROW) { continue }

    const DUE_DATE = String(ROW[MULTI_DUE_DATE_COL])
    const INDEX = __GetDateIndex(DUE_DATE)
    if (INDEX === -1) { continue }
    if (ROW[MULTI_DUE_DATE_COL] === "") { continue }

    const NEW_ROW = new Array<string | number>()
    NEW_ROW[ONE_DUE_DATE_COL] = DUE_DATE
    NEW_ROW[ONE_PURCHASE_DATE_COL] = ROW[MULTI_PURCHASE_DATE_COL]
    NEW_ROW[ONE_PAYMENT_AMT_COL] = ROW[MULTI_PAYMENT_AMT_COL]
    ONE_WEEK_TAB.InsertRow(INDEX, NEW_ROW)
  }

  ONE_WEEK_TAB.SaveToTab()
}

function onEdit(_: unknown) {
  ComputeTotal();
  GenerateRepaymentSchedule();
}

function onOpen(_: unknown) {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Compute One Week Loans", "ComputeTotal")
    .addToUi();

  //AddMultiWeekLoanToRepayment()
  ComputeMonthlyIncome();
  GroupByDate("Due Date", "One Week Loans");
  GroupByDate("Purchase Date", "Multi Week Loans", false);
}