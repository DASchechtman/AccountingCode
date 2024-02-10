function __ComputeIncomeForEachMonth() {
    const ALL_YEARS = 'ALL-YEARS'
    const BUDGET_TAB = new GoogleSheetTabs(BUDGET_PLANNER_TAB_NAME)
    const HEADERS = BUDGET_TAB.GetRow(0)!
  
    const MON = "Monday"
    const TUE = "Tuesday"
    const WED = "Wednesday"
    const THU = "Thursday"
    const FRI = "Friday"
  
    const WEEKLY = "Weekly"
    const BI_WEEKLY = "Bi-Weekly"
    const SEMI_MONTHLY = "Semi-Monthly"
    const MONTHLY = "Monthly"
  
    const INCOME_PER_MONTH_ROW = BUDGET_TAB.IndexOfRow(row => row[1] === "Income Per Paycheck")
    const INCOME_STREAM_ROW = BUDGET_TAB.IndexOfRow(row => row[1] === "Income Stream")
    const JAN_COL = BUDGET_TAB.GetHeaderIndex("January")
  
    let day = ""
    let pay_schedule = ""
    let pay = 0
  
    const __SetDate = function (day: string) {
      let day_code = -1
      switch (day) {
        case MON: { day_code = 1; break }
        case TUE: { day_code = 2; break }
        case WED: { day_code = 3; break }
        case THU: { day_code = 4; break }
        case FRI: { day_code = 5; break }
      }
  
      return function (date: Date) {
        while (date.getUTCDay() !== day_code) {
          date.setDate(date.getDate() + 1)
        }
        return date
      }
    }
  
    const __PayBiWeekly = function (params: PayOutParams) {
      return params.total_days % (params.inc * 2) === 0
    }
  
    const __PaySemiMonthly = function () {
      let month = ""
      let count = 0
      return function(params: PayOutParams) {
        if (month !== params.pay_month) {
          month = params.pay_month
          count = 0
        }
        return count++ < 2
      }
    }
  
    const __PayMonthly = function () {
      let month = ""
      let paid = false
      return function(params: PayOutParams) {
        if (month !== params.pay_month) {
          month = params.pay_month
          paid = false
        }
        const PAID = paid
        paid = true
        return !PAID
      }
    }
  
  
    const __SetPayoutCheck = function (schedule: string) {
      let func: ((params: PayOutParams) => boolean) = (_) => true
  
      if (schedule === BI_WEEKLY) {
        func = __PayBiWeekly
      }
      else if (schedule === SEMI_MONTHLY) {
        func = __PaySemiMonthly()
      }
      else if (schedule === MONTHLY) {
        func = __PayMonthly()
      }
  
      return func
    }
  
    for (let i = INCOME_STREAM_ROW + 1; i < BUDGET_TAB.NumberOfRows(); i++) {
      const ROW = BUDGET_TAB.GetRow(i)!
      const PAY_ROW_INDEX = BUDGET_TAB.IndexOfRow(row => row[1] === ROW[1] && row[0] === ROW[0], INCOME_PER_MONTH_ROW)
      const PAY_ROW = BUDGET_TAB.GetRow(PAY_ROW_INDEX)!
      const PAY = new PayDay(0, new Date(`1/1/${ROW[0]}`), () => true)
      for (let j = JAN_COL - 1; j < ROW.length; j += 2) {
        if (pay === 0 || PAY_ROW[j + 1] !== "") {
          pay = Number(PAY_ROW[j + 1])
          PAY.SetPayoutAmount(pay)
        }
  
        if (ROW[j] === "Custom") {
          PAY.SetPayoutCheck(__SetPayoutCheck(MONTHLY))
        }
        else if (ROW[j] !== "N/A") {
          [day, pay_schedule] = String(ROW[j]).split(" : ")
          PAY.SetPayoutDate(__SetDate(day))
          PAY.SetPayoutCheck(__SetPayoutCheck(pay_schedule))
        }
  
        let total = 0
        while (PAY.PayMonth() === HEADERS[j + 1]) {
          total = __AddToFixed(total, PAY.PayOut())
        }
        ROW[j + 1] = total
      }
      BUDGET_TAB.WriteRow(i, ROW)
    }
  
    BUDGET_TAB.SaveToTab()
  }

function GenerateIncomeSchedule() {
    __ComputeIncomeForEachMonth()
}