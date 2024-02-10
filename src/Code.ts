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

  const WEEKLY_COL_INDEXES = [
    ONE_WEEK_TAB.GetHeaderIndex("Due Date"),
    ONE_WEEK_TAB.GetHeaderIndex("Purchase Date"),
    ONE_WEEK_TAB.GetHeaderIndex("Amount"),
    ONE_WEEK_TAB.GetHeaderIndex("Purchase Location"),
    ONE_WEEK_TAB.GetHeaderIndex("Card")
  ]

  if (
    !__CheckAllAreNotInvalidIndex(WEEKLY_COL_INDEXES)
    || !__CheckAllAreNotInvalidIndex(MULTI_COL_INDEXES)
  ) { return }

  const [
    MULTI_TAB_DUE_DATE_COL_INDEX,
    MULTI_TAB_PURCHASE_DATE_COL_INDEX,
    MULTI_TAB_PAYMENT_AMT_COL_INDEX,
    MULTI_TAB_PURCHASE_LOCATION_COL_INDEX,
    MULTI_TAB_CARD_COL_INDEX
  ] = MULTI_COL_INDEXES

  const [
    WEEKLY_TAB_DUE_DATE_COL_INDEX,
    WEEKLY_TAB_PURCHASE_DATE_COL_INDEX,
    WEEKLY_TAB_PAYMENT_AMT_COL_INDEX,
    WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX,
    WEEKLY_TAB_CARD_COL_INDEX
  ] = WEEKLY_COL_INDEXES

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


function onEdit(e: unknown) {
  if (!__EventObjectIsEditEventObject(e)) { return }
  const TAB_NAME = e.range.getSheet().getName();

  switch (TAB_NAME) {
    case MULTI_WEEK_LOANS_TAB_NAME: {
      break
    }
    case ONE_WEEK_LOANS_TAB_NAME: {
      __ComputeTotal()
      break
    }
    case BUDGET_PLANNER_TAB_NAME: {
      break
    }
  }
}

function onOpen(_: SpreadSheetOpenEventObject) {
 //ComputeMonthlyIncome();

  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Compute One Week Loans", "ComputeOneWeekLoans")
    .addItem("Add Income to Planner", "AddIncometoPlanner")
    .addItem("Generate Income Schedule", "GenerateIncomeSchedule")
    .addItem("Create Multi Week Repayment Schedule", "CreateMultiWeekRepaymentSchedule")
    .addItem("Group One Week Loans", "GroupOneWeekLoans")
    .addToUi();
  __CacheSheets()
}