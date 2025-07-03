
function __SEFC_FindCharges() {
    const TODAY = __Util_CreateDateString(new Date())
    const NOW = new Date(TODAY)
    const YEASTERDAY = new Date(NOW.getTime() - 48 * 60 * 60 * 1000)
    const QUERY = `in:inbox after:${Math.floor(YEASTERDAY.getTime() / 1000)}`
    const LABEL_TESTER = /You made a \$([0-9]+(\.[0-9]{2})?) transaction with (.+)/

    const THREADS = GmailApp.search(QUERY).filter((val) => {
        const IS_FROM_CHASE = val.getMessages()[0].getFrom().includes("no.reply.alerts@chase.com")
        if (!IS_FROM_CHASE) { return false }

        const THREAD_LABEL = val.getFirstMessageSubject()
        const RET = LABEL_TESTER.test(THREAD_LABEL)

        if (RET) {
            const LABEL = GmailApp.getUserLabelByName("Chase Charges")
            val.moveToArchive().addLabel(LABEL)
        }

        return RET
    }).map(val => {
        const RES = LABEL_TESTER.exec(val.getFirstMessageSubject())
        const DATE = new Date(val.getMessages()[0].getDate().toString())
        return {
            amt: Number(RES!.at(1)),
            where: RES!.at(3),
            when: __Util_CreateDateString(DATE)
        }
    })

    return THREADS
}

function ScanEmailForCharges() {
    const CHARGES = __SEFC_FindCharges()
    if (CHARGES.length === 0) { return false }

    const SHEET = new GoogleSheetTabs(WEEKLY_CREDIT_CHARGES_TAB_NAME)
    const INSERT_DATA = new Array<{i: number, date: string}>()
    let in_current_month = false
    let last_date = ""

    const PURCHASE_LOC_INDEX = SHEET.GetHeaderIndex("Purchase Location")
    const START = Number(__Cache_Utils_QueryFirstWeek('start'))
    const END = Number(__Cache_Utils_QueryLastWeek('end'))


    SHEET.ForEachRow((row, i) => {
        if (i > END) { return 'break' }
        
        const IS_HEADER = String(row[PURCHASE_LOC_INDEX]).startsWith(PURCHASE_HEADER)

        if (IS_HEADER) {
            last_date = __Util_GetDateFromDateHeader(String(row[PURCHASE_LOC_INDEX]))
            in_current_month = __Util_DateInCurrentPayPeriod(last_date)
        }

        if (in_current_month && IS_HEADER) {
            INSERT_DATA.push({
                i: i,
                date: last_date
            })
        }
        
    }, START)

    for (let charge of CHARGES) {
        let insert_row = INSERT_DATA[0]

        for (let i = 0; i < INSERT_DATA.length; i++) {
            let insert = INSERT_DATA[i]
            if (new Date(charge.when) >= new Date(insert.date)) {
                insert_row = INSERT_DATA[i]
            }
        }

        SHEET.InsertRow(insert_row.i+1, ["Chase", "Card", charge.where!, charge.amt!, insert_row.date, charge.when!])
    }

    SHEET.SaveToTab()

    return true
}