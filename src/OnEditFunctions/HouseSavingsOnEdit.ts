function __HSOE_FindBucketRowIndex(label: string, sheet: GoogleSheetTabs) {
    for (let i = 0; i < sheet.NumberOfRows(); i++) {
        let row = sheet.GetRow(i)
        if (row && row[0] === label) {
            return i
        }
    }
    return -1
}

function __HSOE_GetSavingBucketRange(bucket: string, sheet: GoogleSheetTabs, start: number) {
    let start_range = ''
    let end_range = ''
    let found_bucket = false
    let start_of_bucket = 0

    const IsNotANumber = (val: any)=> {
        const IS_A_NUM = !isNaN(Number(val))
        const IS_TICKER_NUM = String(val).startsWith("=GOOGLEFINANCE")

        return !IS_A_NUM && !IS_TICKER_NUM
    } 

    sheet.ForEachRow((row, i) => {

        const FOUND_END_OF_BUCKET = (
            found_bucket
            && IsNotANumber(row[0])
            && found_bucket
        )

        const FOUND_START_OF_BUCKET = row[0] === bucket

        if (FOUND_START_OF_BUCKET) {
            found_bucket = true
            start_of_bucket = i + 2
        }
        else if (start_range === '' && found_bucket && start_of_bucket === i) {
            start_range = `A${i + 1}`
        }
        else if (FOUND_END_OF_BUCKET && i > start_of_bucket) {
            end_range = `A${i}`
            return 'break'
        }

    }, start)

    if (start_range !== '' && end_range === '') {
        end_range = `A${sheet.NumberOfRows()}`
    }

    return `${start_range}:${end_range}`
}

function HouseSavingsOnEdit() {
    const HOUSE_SAVINGS_SHEET = new GoogleSheetTabs(HOUSE_SAVINGS_TAB_NAME)

    const BUCKETS_ROW = __HSOE_FindBucketRowIndex("Savings Buckets", HOUSE_SAVINGS_SHEET)

    HOUSE_SAVINGS_SHEET.ForEachRow((row) => {
        if (row[0] === 'Total') { return 'break' }

        const RANGE = __HSOE_GetSavingBucketRange(row[0] as string, HOUSE_SAVINGS_SHEET, BUCKETS_ROW)
        
        const SPLIT_RANGE = RANGE
            .split(':')
            .map(s => s.trim())
            .filter(s => s !== '')

        if (SPLIT_RANGE.length !== 2) {
            row[1] = 0
        }
        else {
            row[1] = `=SUM(${RANGE})`
        }

        return row
    }, true)

    HOUSE_SAVINGS_SHEET.SaveToTab()
}