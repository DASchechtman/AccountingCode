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


function FigureOutPaymentPlanner() {
    const SHEET = new GoogleSheetTabs("Settings");
    const SHORTEN_MONTHS = MONTHS.map(month => month.slice(0, 3));
    let last_date = new Date()

    for (let row_index = 6; row_index < SHEET.NumberOfRows(); row_index++) {
        let row = SHEET.GetRow(row_index)!
        for (let cell_index = 0; cell_index < row.length; cell_index++) {
            if (!SHORTEN_MONTHS.includes(String(row[cell_index]))) { continue }
            let pay_sched_start_date: string | Date = String(row[cell_index + 1])
            const FREQUENCY = String(row[cell_index + 2])
            const TAKE_HOME_PAY = Number(row[cell_index + 3])

            if (pay_sched_start_date === "") {
                pay_sched_start_date = last_date
            }

            if (!pay_sched_start_date || !FREQUENCY || !TAKE_HOME_PAY) {
                continue
            }

            pay_sched_start_date = new Date(pay_sched_start_date)
            const PAY_DAY = new PayDay(Number(TAKE_HOME_PAY), pay_sched_start_date, __FOPP_GetPayoutFunc(FREQUENCY))
            let amt = 0
            let cur_month = PAY_DAY.PayMonth().slice(0, 3).toUpperCase()
            const CELL_MONTH = String(row[cell_index]).toUpperCase()

            while (cur_month === CELL_MONTH) {
                amt += PAY_DAY.PayOut()
                cur_month = PAY_DAY.PayMonth().slice(0, 3).toUpperCase()
            }

            last_date = PAY_DAY.GetDate()

            row[cell_index + 1] = __Util_CreateDateString(pay_sched_start_date)
            row[cell_index + 3] = amt
        }
        SHEET.WriteRow(row_index, row)
    }

    SHEET.SaveToTab()
}