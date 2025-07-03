type ImportedData = Array<{ date: string, name: string, amt: number }>

function __ICH_IsCorrectInput(data: any): data is ImportedData {
    if (!(data instanceof Array)) {
        return false
    }

    for (let el of data) {
        if (!(data instanceof Object)) { return false }
        if (!('date' in el) || !('name' in el) || !('amt' in el)) { return false }
        if (typeof el.date !== 'string' || typeof el.name !== 'string' || typeof el.amt !== 'number') { return false }
    }

    return true
}

function __ICH_CreateRowObject(insert_row: number) {
    return {
        insert_row: insert_row,
        existing_rows: [],
        new_rows: []
    }
}

function __ICH_RemoveAllGroups(sheet: GoogleSheetTabs) {
    const TAB = sheet.GetTab()
    const START = Number(__Cache_Utils_QueryFirstWeek('start'))
    const END = Number(__Cache_Utils_QueryLastWeek('end'))
    for (let i = START; i < END; i++) {
        try {
            TAB.getRowGroup(i + 1, 1)?.remove()
        } catch {}
    }
}

function __ICH_FindInsertIndex(sheet: GoogleSheetTabs, start_row: number, date: string) {
    const PURCHASE_LOC_INDEX = sheet.GetHeaderIndex("Purchase Location")

    for (let i = start_row; i < sheet.NumberOfRows(); i++) {
        const ROW = sheet.GetRow(i)!
        if (String(ROW[PURCHASE_LOC_INDEX]).includes(date)) { return i }
    }
    return -1
}

function __ICH_RecordExistingPurchases(sheet: GoogleSheetTabs, purchase_loc_index: number): [DataArrayEntry[], string[]] {
    const ROWS_TO_COMPARE = new Array<DataArrayEntry>()
    const DATES = new Array<string>()
    const START = Number(__Cache_Utils_QueryFirstWeek('start'))
    const END = Number(__Cache_Utils_QueryLastWeek('end'))
    let found_current_pay_period_rows = false

    sheet.ForEachRow((row, i) => {
        if (i > END) { return 'break' }
        const STARTS_WITH_HEADER = String(row[purchase_loc_index]).startsWith(PURCHASE_HEADER)

        let date = ""
        if (STARTS_WITH_HEADER) {
            date = __Util_GetDateFromDateHeader(row[purchase_loc_index] as string)
        }

        if (date !== "" && __Util_DateInCurrentPayPeriod(date)) {
            found_current_pay_period_rows = true
        }
        else if (STARTS_WITH_HEADER && found_current_pay_period_rows && !__Util_DateInCurrentPayPeriod(date)) {
            found_current_pay_period_rows = false
        }

        if (found_current_pay_period_rows) {
            if (STARTS_WITH_HEADER) {
                DATES.push(`${date}:${i}`)
            }
            else {
                ROWS_TO_COMPARE.push(row)
            }
        }
    }, START)

    return [ROWS_TO_COMPARE, DATES]
}

function __ICH_FilterNewPurchases(imported_data: ImportedData, ROWS_TO_COMPARE: DataArray) {
    const ROWS_TO_ADD = new Array<DataArrayEntry>()

    for (let el of imported_data) {
        let index = ROWS_TO_COMPARE.findIndex((x, i)=> {
            return x.includes(el.amt)
        })

        if (index > -1) {
            ROWS_TO_COMPARE.splice(index, 1)
        }
        else {
            ROWS_TO_ADD.push(["Chase", "Card", el.name, el.amt, "", el.date])
        }
        
    }

    return ROWS_TO_ADD
}

function __ICH_RecordNewPurchases(
    ROWS_TO_ADD: DataArray, 
    DATES: Array<string>, 
    PURCHASE_DATE_INDEX: number, 
    SHEET_TRACKER: GoogleSheetTabs,
    DUE_DATE_INDEX: number
) {
    for (let el of ROWS_TO_ADD) {
        let arr = String(DATES[0]).split(":")
        let last_date = arr[0]
        let last_insert_index = Number(arr[1])
   
        for (let data of DATES) {
            let [group_date, group_index] = data.split(':')

            let date_1 = new Date(el[PURCHASE_DATE_INDEX] as string)
            let date_2 = new Date(group_date)

            if (date_1 >= date_2) {
                last_date = group_date
                last_insert_index = __ICH_FindInsertIndex(SHEET_TRACKER, last_insert_index, last_date)
            }
        }

        el[DUE_DATE_INDEX] = last_date
        SHEET_TRACKER.InsertRow(last_insert_index + 1, el)
    }
}

function __ICH_AddToSheet(imported_data: any) {
    if (!__ICH_IsCorrectInput(imported_data)) { throw new Error("Wrong Input!") }

    const SHEET_TRACKER = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    const CARD_INDEX = SHEET_TRACKER.GetHeaderIndex("Card")
    const PAY_INDEX = SHEET_TRACKER.GetHeaderIndex("Pay Where?")
    const PURCHASE_LOC_INDEX = SHEET_TRACKER.GetHeaderIndex("Purchase Location")
    const AMOUNT_INDEX = SHEET_TRACKER.GetHeaderIndex("Amount")
    const DUE_DATE_INDEX = SHEET_TRACKER.GetHeaderIndex('Due Date')
    const PURCHASE_DATE_INDEX = SHEET_TRACKER.GetHeaderIndex("Purchase Date")

    let [ROWS_TO_COMPARE, DATES] = __ICH_RecordExistingPurchases(SHEET_TRACKER, PURCHASE_LOC_INDEX)
    let ROWS_TO_ADD = __ICH_FilterNewPurchases(imported_data, ROWS_TO_COMPARE)
    __ICH_RecordNewPurchases(ROWS_TO_ADD, DATES, PURCHASE_DATE_INDEX, SHEET_TRACKER, DUE_DATE_INDEX)
    SHEET_TRACKER.SaveToTab()

    __Cache_Utils_StoreOneWeekLoanCurrentMonthInfo()
    __ICH_RemoveAllGroups(SHEET_TRACKER)
    __Util_GroupCurrentMonthCharges()
}

function ImportCreditHistory() {
    const HTML = HtmlService.createHtmlOutputFromFile("ImportUi")
        .setWidth(700)
        .setHeight(600)
    SpreadsheetApp.getUi().showModalDialog(HTML, "Importer")
}