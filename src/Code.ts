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

function onOpen() {
  const UI = SpreadsheetApp.getUi();

  UI.createMenu("Budgeting")
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Group One Week Loans", "GroupOneWeekLoans")
    .addItem("Break Down Repayments", "BreakDownRepayment")
    .addItem("Tally Personal Ledger Expenses", "TallyPersonalLedgerExpenses")
    .addToUi();
}

function onDailyTrigger() {
  GroupAndCollapseBills()
  //ComputeTotalMonthly()
}