
type PayOutParams = {date: Date, total_days: number, inc: number, pay_month: string};
type CheckPayOut = ({date, total_days, inc, pay_month}: PayOutParams) => boolean;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Tab = GoogleAppsScript.Spreadsheet.Sheet;
type DataArray = Array<DataArrayEntry>;
type DataArrayItem = string | number | boolean
type DataArrayEntry = Array<DataArrayItem>;
type DataArrayElement = string | number;
type Some = {type: "Some", val: NonNullable<unknown>}
type None = {type: "None"}
type Maybe = Some | None
type SpreadSheetEditEventObject = {
    authMode: GoogleAppsScript.Script.AuthMode;
    triggerUid: string;
    user: GoogleAppsScript.Base.User;
    source: Spreadsheet;
    range: GoogleAppsScript.Spreadsheet.Range;
    value: string;
    oldValue: string;
    changeType: string;
}
type SpreadSheetOpenEventObject = {
    authMode: GoogleAppsScript.Script.AuthMode;
    source: Spreadsheet;
    triggerUid: string;
    user: GoogleAppsScript.Base.User;
}

const PURCHASE_HEADER = "Purchases for"
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
const PAYMENT_SCHEDULE = ["Weekly", "Bi-Weekly", "Semi-Monthly", "Monthly"]

const WEEKLY_CREDIT_CHARGES_TAB_NAME = "One Week Loans"
const HOUSE_BUDGET_DASHBOARD_TAB_NAME = "Household Budget Dashboard"
const MULTI_WEEK_LOANS_TAB_NAME = "Multi Week Loans"
const PERSONAL_SPEND_TRACKER_TAB_NAME = "Personal Spend Tracker"
const BUDGET_PLANNER_TAB_NAME = "Budget Planner"
const PAYMENT_SCHEDULE_TAB_NAME = "Settings"
const WEEKLY_PAYMENT_BREAK_DOWN_TAB_NAME = "One Week Loans Breakdown"
const PERSONAL_LEDGER_TAB_NAME = "Personal Spend Ledger"
const HOUSE_SAVINGS_TAB_NAME = "Household Savings"
const FRIEND_FAMILY_LOANS_TAB_NAME = "Family/Friend Loans"
const INVESTMENT_DATA_TAB_NAME = "Investment Data"
const INVESTMENT_ALLOC_TAB = "Robinhood Investment Allocation"

const INVESTMENT_TRACKER_RANGE = "G2:Q13"