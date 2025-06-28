
function __DRR_CreateReminder(index: number, sheet: GoogleSheetTabs) {
    const WHO_INDEX = index + 1


    const REPAY_DATE_INDEX = index + 6
    const AMOUNT_OWED_INDEX = index + 4

    sheet.ForEachRow((row, i)  => {
        if (['ro', 'dan', ''].includes(String(row[WHO_INDEX]).toLowerCase())) { return 'break' }

        const WHO = row[WHO_INDEX]
        const AMT_OWED = row[AMOUNT_OWED_INDEX]
        const REPAY_DATE = new Date(String(row[REPAY_DATE_INDEX]))

        if (REPAY_DATE.toString() === 'Invalid Date') { return 'continue' }

        const EVENT_TITLE = `Reminder: ${WHO} owes $${AMT_OWED} on ${__Util_CreateDateString(REPAY_DATE)}`

        const ALL_EVENTS = CalendarApp.getEventsForDay(REPAY_DATE)
        const INDEX = ALL_EVENTS.findIndex(e => e.getTitle() === EVENT_TITLE)

        if (INDEX === -1) {
            let event = CalendarApp.getDefaultCalendar().createAllDayEvent(EVENT_TITLE, REPAY_DATE)
            event.addGuest('rbautista969@yahoo.com')
            event.addEmailReminder(1440)
        }
    }, 1)

}

function RepayReminder() {
    const REMINDER_PAY_TAB = new GoogleSheetTabs(FRIEND_FAMILY_LOANS_TAB_NAME)
    const REMINDER_GROUP_SIZE = 8
    const REPAYMENT_ROW = REMINDER_PAY_TAB.GetRow(0)!

    for (let i = 0; i < REPAYMENT_ROW.length; i += REMINDER_GROUP_SIZE + 1) {
        __DRR_CreateReminder(i, REMINDER_PAY_TAB)
    }
}