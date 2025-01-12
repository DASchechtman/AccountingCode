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

    for (let i = start; i < sheet.NumberOfRows(); i++) {
        const ROW = sheet.GetRow(i)
        const DEPOSIT_AMOUNT = Number(ROW?.at(0))

        const FOUND_END_OF_BUCKET = (
            found_bucket
            && isNaN(DEPOSIT_AMOUNT)
        )

        const FOUND_START_OF_BUCKET = ROW?.at(0) === bucket

        if (FOUND_END_OF_BUCKET) {
            end_range = `A${i}`
            break
        }
        else if (FOUND_START_OF_BUCKET) {
            found_bucket = true
            i++
        }
        else if (start_range === '' && found_bucket) {
            start_range = `A${i + 1}`
        }
    }

    if (found_bucket && end_range === '') {
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