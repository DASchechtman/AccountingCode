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
  const UI = SpreadsheetApp.getUi();

  const INCOME_FUNCS = UI.createMenu("Income Features")
    .addItem("Add Income to Planner", "AddIncomeToPlanner")
    .addItem("Generate Income Schedule", "GenerateIncomeSchedule")
  
  const LOAN_FUNCS = UI.createMenu("Loan Features")
    .addItem("Compute One Week Loans", "ComputeOneWeekLoans")
    .addItem("Create Multi Week Repayment Schedule", "CreateMultiWeekRepaymentSchedule")

  UI.createMenu("Budgeting")
    .addSubMenu(INCOME_FUNCS)
    .addSubMenu(LOAN_FUNCS)
    .addItem("Create New Household Budget Tab", "CreateNewHouseholdBudgetTab")
    .addItem("Group One Week Loans", "GroupOneWeekLoans")
    .addToUi();
  __CacheSheets()
}