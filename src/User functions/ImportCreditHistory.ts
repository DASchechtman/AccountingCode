type ImportedData = Array<{ date: string; name: string; amt: number }>;

function __ICH_IsCorrectInput(data: any): data is ImportedData {
  if (!(data instanceof Array)) {
    return false;
  }

  for (let el of data) {
    if (!(data instanceof Object)) {
      return false;
    }
    if (!("date" in el) || !("name" in el) || !("amt" in el)) {
      return false;
    }
    if (
      typeof el.date !== "string" ||
      typeof el.name !== "string" ||
      typeof el.amt !== "number"
    ) {
      return false;
    }
  }

  return true;
}

function __ICH_CreateRowObject(insert_row: number) {
  return {
    insert_row: insert_row,
    existing_rows: [],
    new_rows: [],
  };
}

function __ICH_RemoveAllGroups(sheet: GoogleSheetTabs) {
  const TAB = sheet.GetTab();
  const START = __Util_GetRowThatStartsTheMonth();
  const END = __Util_GetRowThatEndsTheMonth();
  for (let i = START; i < END; i++) {
    try {
      TAB.getRowGroup(i + 1, 1)?.remove();
    } catch {}
  }
}

function __ICH_FindInsertIndex(
  sheet: GoogleSheetTabs,
  start_row: number,
  date: string,
) {
  const PURCHASE_LOC_INDEX = sheet.GetHeaderIndex("Purchase Location");
  return sheet.FindRowIndex((row) =>
    String(row[PURCHASE_LOC_INDEX]).includes(date),
  );
}

function __ICH_RecordExistingPurchases(
  sheet: GoogleSheetTabs,
  purchase_loc_index: number,
): [DataArrayEntry[], string[]] {
  const ROWS_TO_COMPARE = new Array<DataArrayEntry>();
  const DATES = new Array<string>();
  const START = __Util_GetRowThatStartsTheMonth();
  const END = __Util_GetRowThatEndsTheMonth();
  let found_current_pay_period_rows = false;

  sheet.ForEachRow(
    (row, i) => {
      if (i > END) {
        return "break";
      }
      const STARTS_WITH_HEADER = String(row[purchase_loc_index]).startsWith(
        PURCHASE_HEADER,
      );

      let date = "";
      if (STARTS_WITH_HEADER) {
        date = __Util_GetDateFromDateHeader(row[purchase_loc_index] as string);
      }

      if (date !== "" && __Util_DateInCurrentPayPeriod(date)) {
        found_current_pay_period_rows = true;
      } else if (
        STARTS_WITH_HEADER &&
        found_current_pay_period_rows &&
        !__Util_DateInCurrentPayPeriod(date)
      ) {
        found_current_pay_period_rows = false;
      }

      if (found_current_pay_period_rows) {
        if (STARTS_WITH_HEADER) {
          DATES.push(`${date}:${i}`);
        } else {
          ROWS_TO_COMPARE.push(row);
        }
      }
    },
    START,
    END + 1,
  );

  return [ROWS_TO_COMPARE, DATES];
}

function __ICH_FilterNewPurchases(
  imported_data: ImportedData,
  ROWS_TO_COMPARE: DataArray,
) {
  const ROWS_TO_ADD = new Array<DataArrayEntry>();

  for (let el of imported_data) {
    let index = ROWS_TO_COMPARE.findIndex((x, i) => {
      return x.includes(el.amt);
    });

    if (index > -1) {
      ROWS_TO_COMPARE.splice(index, 1);
    } else {
      ROWS_TO_ADD.push(["Chase", "Card", el.name, el.amt, "", el.date]);
    }
  }

  return ROWS_TO_ADD;
}

function __ICH_RecordNewPurchases(
  ROWS_TO_ADD: DataArray,
  DATES: Array<string>,
  PURCHASE_DATE_INDEX: number,
  SHEET_TRACKER: GoogleSheetTabs,
  DUE_DATE_INDEX: number,
) {
  for (let el of ROWS_TO_ADD) {
    let arr = String(DATES[0]).split(":");
    let last_date = arr[0];
    let last_insert_index = Number(arr[1]);

    for (let data of DATES) {
      let [group_date, group_index] = data.split(":");

      let date_1 = new Date(el[PURCHASE_DATE_INDEX] as string);
      let date_2 = new Date(group_date);

      if (date_1 >= date_2) {
        last_date = group_date;
        last_insert_index = __ICH_FindInsertIndex(
          SHEET_TRACKER,
          last_insert_index,
          last_date,
        );
      }
    }

    el[DUE_DATE_INDEX] = last_date;
    SHEET_TRACKER.InsertRow(last_insert_index + 1, el);
  }
}

function __ICH_IsBeforeThe26th(date: string) {
  const DATE_STR = __Util_CreateDateString(date).split("/");
  const END_DATE = new Date(`${DATE_STR[0]}/26/${DATE_STR[2]}`);
  const DATE1 = new Date(date);
  return DATE1 <= END_DATE;
}

function __ICH_AddToSheet(imported_data: any) {
  if (imported_data === undefined) {
    imported_data = [
      {
        date: __Util_CreateDateString(new Date()),
        name: "Example Purchase",
        amt: 5,
      },
    ];
  }
  if (!__ICH_IsCorrectInput(imported_data)) {
    throw new Error("Wrong Input!");
  }
  console.log(JSON.stringify(imported_data));
  imported_data = imported_data.filter((x) => __ICH_IsBeforeThe26th(x.date));

  const SHEET_TRACKER = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME);
  const CARD_INDEX = SHEET_TRACKER.GetHeaderIndex("Card");
  const PAY_WHERE_INDEX = SHEET_TRACKER.GetHeaderIndex("Pay Where?");
  const PURCHASE_LOC_INDEX = SHEET_TRACKER.GetHeaderIndex("Purchase Location");
  const AMT_INDEX = SHEET_TRACKER.GetHeaderIndex("Amount");
  const DUE_DATE_INDEX = SHEET_TRACKER.GetHeaderIndex("Due Date");
  const PURCHASE_DATE_INDEX = SHEET_TRACKER.GetHeaderIndex("Purchase Date");
  const MAP = new Array<{
    date: string;
    data: { date: string; name: string; amt: number }[];
  }>();
  const START = __Util_GetRowThatStartsTheMonth();
  const END = __Util_GetRowThatEndsTheMonth();

  SHEET_TRACKER.ForEachRow(
    (row, i) => {
      const PURCHASE = String(row[PURCHASE_LOC_INDEX]);
      if (__Util_IsHeader(PURCHASE)) {
        const DATE = __Util_GetDateFromDateHeader(PURCHASE);
        MAP.push({
          date: DATE,
          data: [],
        });
      }
    },
    START,
    END + 1,
  );

  for (let el of imported_data) {
    for (let i = 0; i < MAP.length; i++) {
      const CUR = MAP.at(i)!;
      const NEXT = MAP.at(i + 1);
      if (
        NEXT &&
        new Date(el.date) >= new Date(CUR.date) &&
        new Date(el.date) < new Date(NEXT.date)
      ) {
        CUR.data.push(el);
        break;
      } else if (NEXT === undefined) {
        CUR.data.push(el);
      }
    }
  }

  for (let entry of MAP) {
    for (let el of entry.data) {
      const ARR = new Array<string | number>();
      ARR[CARD_INDEX] = "Chase";
      ARR[PAY_WHERE_INDEX] = "Card";
      ARR[PURCHASE_LOC_INDEX] = el.name;
      ARR[PURCHASE_DATE_INDEX] = el.date;
      ARR[DUE_DATE_INDEX] = el.date;
      ARR[AMT_INDEX] = el.amt;
      const insert_row =
        SHEET_TRACKER.FindRowIndex((row) =>
          String(row[PURCHASE_LOC_INDEX]).includes(entry.date),
        ) + 1;
      const ROW = SHEET_TRACKER.GetRow(insert_row);
      if (ROW && ROW[PURCHASE_DATE_INDEX] === "") {
        ROW[CARD_INDEX] = ARR[CARD_INDEX];
        ROW[PAY_WHERE_INDEX] = ARR[PAY_WHERE_INDEX];
        ROW[PURCHASE_LOC_INDEX] = ARR[PURCHASE_LOC_INDEX];
        ROW[PURCHASE_DATE_INDEX] = ARR[PURCHASE_DATE_INDEX];
        ROW[DUE_DATE_INDEX] = ARR[DUE_DATE_INDEX];
        ROW[AMT_INDEX] = ARR[AMT_INDEX];
        SHEET_TRACKER.OverWriteRow(ROW);
      } else {
        SHEET_TRACKER.InsertRow(insert_row + 1, ARR);
      }
    }
  }

  SHEET_TRACKER.SaveToTab();
}

function ImportCreditHistory() {
  const HTML = HtmlService.createHtmlOutputFromFile("ImportUi")
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(HTML, "Importer");
}
