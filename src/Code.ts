function SafelyCreateMenu(CreateMenu: () => void) {
  try {
    CreateMenu();
  } catch {}
}

function onEdit(e: unknown) {
  if (!__Util_EventObjectIsEditEventObject(e)) {
    return;
  }
  const TAB_NAME = e.range.getSheet().getName();

  if (TAB_NAME === WEEKLY_CREDIT_CHARGES_TAB_NAME) {
    WeeklyCreditChargesOnEdit();
  } else if (TAB_NAME === HOUSE_SAVINGS_TAB_NAME) {
    HouseSavingsOnEdit();
  } else if (TAB_NAME === INVESTMENT_ALLOC_TAB) {
    InvestmentAllocationCalcOnEdit();
  }
}

function onOpen() {
  const UI = SpreadsheetApp.getUi();

  SafelyCreateMenu(() => {
    UI.createMenu("Budgeting")
      .addToUi();
  });

  SafelyCreateMenu(() => {
    UI.createMenu("Debug")
      .addItem("Test Daily Trigger", "onDailyTrigger")
      .addToUi();
  });
}

function onDailyTrigger() {
  const TODAY = new Date();
  if (TODAY.getDate() !== 28) { return }
  StoreSpendDataAndReset()
}

function onHourlyTrigger() {
  ScanEmailForCharges();
}
