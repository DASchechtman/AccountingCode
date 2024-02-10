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