function __CreateBudgetTab() {
    const TEMPLATE_TAB = new GoogleSheetTabs("Household Budget Template");
    const SHEET = SpreadsheetApp.getActiveSpreadsheet();
    let year = new Date().getUTCFullYear();
    let budget_tab: GoogleSheetTabs;
  
    while (true) {
      try{
        budget_tab = new GoogleSheetTabs(`Household Budget ${year}`);
        year++;
      }
      catch(err) {
        break;
      }
    }
  
    SHEET.insertSheet(`Household Budget ${year}`);
    budget_tab = new GoogleSheetTabs(`Household Budget ${year}`);
  
    TEMPLATE_TAB.CopyTo(budget_tab);
  
    budget_tab.GetTab().setFrozenRows(3);
    budget_tab.GetTab().setFrozenColumns(1);
    
    const MISC_HOUSEHOLD_PURCHASES_INDEX = budget_tab.IndexOfRow(budget_tab.FindRow(row => row[0] === "Misc Household Purchases"))
  
    if (MISC_HOUSEHOLD_PURCHASES_INDEX !== -1) {
      budget_tab.GetTab().getRange(`B${MISC_HOUSEHOLD_PURCHASES_INDEX+1}:B`).setNumberFormat("mm/dd/yyyy")
      budget_tab.GetTab().getRange(`C${MISC_HOUSEHOLD_PURCHASES_INDEX+1}:C`).setNumberFormat("$#,##0.00")
    }
  
    return budget_tab
  }

function CreateNewHouseholdBudgetTab() {
    const TAB = __CreateBudgetTab()
    __SetMonthDates(TAB)
    ComputeMonthlyIncome()
  }