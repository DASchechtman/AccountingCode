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
