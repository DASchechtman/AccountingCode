let __FOPP_semi_monthly_counter = 0
let __FOPP_monthly_counter = 0
let __FOPP_semi_monthly_cur_month = ""
let __FOPP_monthly_cur_month = ""
let __FOPP_days_inc = 0

function __FOPP_WeeklyPayOut(_: PayOutParams) {
    return true
}

function __FOPP_BiWeeklyPayOut(x: PayOutParams) {
    let payout = __FOPP_days_inc % 14 === 0
    __FOPP_days_inc += x.inc
    return payout
}

function __FOPP_SemiMonthlyPayOut(x: PayOutParams) {
    if (x.pay_month !== __FOPP_semi_monthly_cur_month) {
        __FOPP_semi_monthly_cur_month = x.pay_month
        __FOPP_semi_monthly_counter = 0
    }
    return __FOPP_semi_monthly_counter++ < 2
}

function __FOPP_MonthlyPayout(x: PayOutParams) {
    if (x.pay_month !== __FOPP_monthly_cur_month) {
        __FOPP_monthly_cur_month = x.pay_month
        __FOPP_monthly_counter = 0
    }
    return __FOPP_monthly_counter++ < 1
}

function __FOPP_GetPayoutFunc(frequency: string) {
    let PayoutFunc: CheckPayOut = () => false

    if (frequency === "Weekly") {
        PayoutFunc = __FOPP_WeeklyPayOut
    } else if (frequency === "Bi-Weekly") {
        PayoutFunc = __FOPP_BiWeeklyPayOut
    }
    else if (frequency === "Semi-Monthly") {
        PayoutFunc = __FOPP_SemiMonthlyPayOut
    }
    else if (frequency === "Monthly") {
        PayoutFunc = __FOPP_MonthlyPayout
    }

    return PayoutFunc
}

function __FOPP_ConvertToPayPercentage(take_home: number, percent: string) {
    const PERCENT_SIGN_TEST = /^\d{1,3}\%$/
    const PERCENT_DEC_TEST = /^(1\.0)|(0\.0)|(0\.\d{1,2})$/
    
    if (PERCENT_SIGN_TEST.test(percent)) {
        const PERCENT = percent.split("%")[0].toNumber()
        return take_home * (PERCENT / 100)
    }

    if (PERCENT_DEC_TEST.test(percent)) {
        return take_home * Number(percent)
    }

    return 0
}


function FigureOutPaymentPlanner() {
    const SHEET = new GoogleSheetTabs("Settings");
    const USER_ROW = SHEET.GetRow(5)!.map(x => String(x))
    const LAST_USER_DATES = new Map<string, [Date, number]>()

    for (let row_index = 6; row_index < SHEET.NumberOfRows(); row_index++) {
        let row = SHEET.GetRow(row_index)!
        for (let cell_index = 0; cell_index < row.length; cell_index += 8) {

            if (!LAST_USER_DATES.has(USER_ROW[cell_index])) {
                LAST_USER_DATES.set(USER_ROW[cell_index], [new Date(), 0])
            }

            const USER_DATA = LAST_USER_DATES.get(USER_ROW[cell_index])!
            const PAY_DATE_INDEX = cell_index + 1
            const FREQ_INDEX = cell_index + 2
            const TAKE_HOME_INDEX = cell_index + 3
            const SAVING_INDEX = cell_index + 4
            const LE_INDEX = cell_index + 5
            const PERSONAL_INDEX = cell_index + 6
            const FREQUENCY = String(row[FREQ_INDEX])

            let pay_sched_start_date: string | Date = String(row[PAY_DATE_INDEX])
            let take_home_pay = USER_DATA[1]

            if (String(row[TAKE_HOME_INDEX]).charAt(0) === "=") {
                let val = Number(String(row[TAKE_HOME_INDEX]).substring(1))
                if (!isNaN(val)) { take_home_pay = val }
            }

            if (pay_sched_start_date === "") { pay_sched_start_date = USER_DATA[0] }
            if (!pay_sched_start_date || !FREQUENCY || !take_home_pay) { continue }

            const PAY_DAY = new PayDay(Number(take_home_pay), new Date(pay_sched_start_date), __FOPP_GetPayoutFunc(FREQUENCY))
            const CELL_MONTH = String(row[cell_index]).toUpperCase()
            let num_of_payouts = 0
            const CurMonth = () => { return PAY_DAY.PayMonth().slice(0, 3).toUpperCase() === CELL_MONTH }

            while (CurMonth()) { num_of_payouts += Number(PAY_DAY.PayOut() > 0) }

            row[PAY_DATE_INDEX] = __Util_CreateDateString(pay_sched_start_date)
            row[TAKE_HOME_INDEX] = `=MULTIPLY(${take_home_pay}, ${num_of_payouts})`
            row[SAVING_INDEX] = `=${__Util_IndexToColLetter(TAKE_HOME_INDEX)}${row_index+1}*${__Util_IndexToColLetter(SAVING_INDEX)}6`
            row[LE_INDEX] = `=${__Util_IndexToColLetter(TAKE_HOME_INDEX)}${row_index+1}*${__Util_IndexToColLetter(LE_INDEX)}6`
            row[PERSONAL_INDEX] = `=${__Util_IndexToColLetter(TAKE_HOME_INDEX)}${row_index+1}*${__Util_IndexToColLetter(PERSONAL_INDEX)}6`

            USER_DATA[0] = PAY_DAY.GetDate()
            USER_DATA[1] = take_home_pay

            LAST_USER_DATES.set(USER_ROW[cell_index], USER_DATA)
        }
        SHEET.WriteRow(row_index, row)
    }

    SHEET.SaveToTab()
}