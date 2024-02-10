function __ComputeMonthlyIncome() {
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
    __ComputeMonthlyIncome()
  }