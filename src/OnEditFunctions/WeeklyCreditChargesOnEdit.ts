const PAY_AMT = [126, 125];

function __WCCOE_GetSumFormula(start_range: string, end_range: string) {
  return `=SUM(ARRAYFORMULA(ROUNDUP(${start_range}:${end_range})))`;
}

function __WCCOE_SetLastRowToHaveSum(
  sheet: GoogleSheetTabs,
  start_range: string,
  amt_col_index: number,
  total_col_index: number,
) {
  const LAST_ROW = sheet.GetRow(sheet.NumberOfRows() - 1)!;
  const END_RANGE = `${__Util_IndexToColLetter(amt_col_index)}${sheet.NumberOfRows()}`;
  LAST_ROW[total_col_index] = __WCCOE_GetSumFormula(start_range, END_RANGE);
  sheet.OverWriteRow(LAST_ROW);
  return END_RANGE;
}

function __WCCOE_GetWeeklyCharges(
  sheet: GoogleSheetTabs,
  start: number,
  purchase_col: number,
  amt_col: number,
) {
  let cells = new Array<string>();
  for (let i = start; i < sheet.NumberOfRows(); i++) {
    const ROW = sheet.GetRow(i)!;
    const PURCHASES = ROW[purchase_col].toString();

    if (PURCHASES.startsWith("Purchases for")) {
      break;
    }

    cells.push(`${__Util_IndexToColLetter(amt_col)}${i + 1}`);
  }

  return `${cells[0]}:${cells.at(-1)}`;
}

function WeeklyCreditChargesOnEdit() {
  const WEEKLY_CHARGES_SHEET = new GoogleSheetTabs(
    WEEKLY_CREDIT_CHARGES_TAB_NAME,
  );
  const TOTAL_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Total");
  const TIPS_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Tips");
  const MONEY_LEFT_COL_INDEX =
    WEEKLY_CHARGES_SHEET.GetHeaderIndex("Money Left");
  const AMT_COL_INDEX = WEEKLY_CHARGES_SHEET.GetHeaderIndex("Amount");
  const PURCHASE_LOC_COL_INDEX =
    WEEKLY_CHARGES_SHEET.GetHeaderIndex("Purchase Location");
  const START = __Util_GetRowThatStartsTheMonth();
  const END = __Util_GetRowThatEndsTheMonth();
  const MONTHLY_ALLOWANCE = new Array<number>();
  const WEEKLY_CHARGES = new Map<number, number>();

  let index = -1;

  WEEKLY_CHARGES_SHEET.ForEachRow((row, i) => {
    const PURCHASE = String(row[PURCHASE_LOC_COL_INDEX]);
    if (__Util_IsHeader(PURCHASE)) {
      index = i;
      WEEKLY_CHARGES.set(i, 0);
      MONTHLY_ALLOWANCE.push(PAY_AMT.at(-1)!);
    } else if (index !== -1) {
      const AMT = WEEKLY_CHARGES.get(index)!;
      WEEKLY_CHARGES.set(index, AMT + 1);
    }
  }, START);

  const FULL_MONTHLY_ALLOWANCE = Math.min(
    MONTHLY_ALLOWANCE.reduce((a, b) => a + b, 0),
    600,
  );
  const SUM_RANGES = new Array<string>();
  const TIP_CELLS = new Array<string>();
  let SaveMoneyLeft: (() => void) | null = null;

  for (let [week_index, charge_count] of WEEKLY_CHARGES) {
    const AMT_COL_LETTER = __Util_IndexToColLetter(AMT_COL_INDEX);
    let end_row = week_index + charge_count + 1;
    const SUM_RANGE = `${AMT_COL_LETTER}${week_index + 2}:${AMT_COL_LETTER}${end_row}`;
    const ROW = WEEKLY_CHARGES_SHEET.GetRow(week_index + charge_count)!;
    ROW[TOTAL_COL_INDEX] = `=SUM(ARRAYFORMULA(ROUNDUP(${SUM_RANGE})))`;
    ROW[MONEY_LEFT_COL_INDEX] = "";
    if (!SaveMoneyLeft) {
      SaveMoneyLeft = () => {
        ROW[MONEY_LEFT_COL_INDEX] =
          `= ${FULL_MONTHLY_ALLOWANCE} + SUM(${TIP_CELLS.join()}) - SUM(${SUM_RANGES})`;
        WEEKLY_CHARGES_SHEET.OverWriteRow(ROW);
      };
    }

    TIP_CELLS.push(`${__Util_IndexToColLetter(TIPS_INDEX)}${week_index + 1}`);
    WEEKLY_CHARGES_SHEET.OverWriteRow(ROW);
    SUM_RANGES.push(SUM_RANGE);
  }

  SaveMoneyLeft?.call(null);
  WEEKLY_CHARGES_SHEET.SaveToTab();
}
