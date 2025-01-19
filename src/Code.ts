function onEdit(e: unknown) {
  if (!__Util_EventObjectIsEditEventObject(e)) { return }
  const TAB_NAME = e.range.getSheet().getName();

  if (TAB_NAME === WEEKLY_CREDIT_CHARGES_TAB_NAME) { 
    WeeklyCreditChargesOnEdit() 
  }
  else if (TAB_NAME === HOUSE_SAVINGS_TAB_NAME) { 
    HouseSavingsOnEdit() 
  }
}

function onOpen(_: SpreadSheetOpenEventObject) {
  const UI = SpreadsheetApp.getUi();

  const INCOME_FUNCS = UI.createMenu("Income Features")
    .addItem("Add Income to Planner", "AddIncomeToPlanner")
    .addItem("Generate Income Schedule", "GenerateIncomeSchedule")
    .addItem("Figure Out Paymet Planner", "FigureOutPaymentPlanner")
  
  const LOAN_FUNCS = UI.createMenu("Loan Features")
    .addItem("Compute One Week Loans", "ComputeOneWeekLoans")
    .addItem("Create Multi Week Repayment Schedule", "CreateMultiWeekRepaymentSchedule")

  UI.createMenu("Budgeting")
    .addSubMenu(INCOME_FUNCS)
    .addSubMenu(LOAN_FUNCS)
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Group One Week Loans", "GroupOneWeekLoans")
    .addItem("Break Down Repayments", "BreakDownRepayment")
    .addItem("Tally Personal Ledger Expenses", "TallyPersonalLedgerExpenses")
    .addToUi();
  __Util_CacheSheets()
}

function onDailyTrigger() {
  GroupAndCollapseBills()
  ComputeTotalMonthly()
}