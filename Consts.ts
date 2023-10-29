type CheckPayOut = (date: Date, dir: number, inc: number) => boolean;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Tab = GoogleAppsScript.Spreadsheet.Sheet;
type DataArray = Array<DataArrayEntry>;
type DataArrayEntry = Array<string | number>;

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
const PAYMENT_SCHEDULE = ["Weekly", "Bi-Weekly", "Semi-Monthly", "Monthly", "Custom"]

const ONE_WEEK_LOANS_TAB_NAME = "One Week Loans"
const HOUSE_BUDGET_DASHBOARD_TAB_NAME = "Household Budget Dashboard"
const MULTI_WEEK_LOANS_TAB_NAME = "Multi Week Loans"
const PERSONAL_SPEND_TRACKER_TAB_NAME = "Personal Spend Tracker"
const BUDGET_PLANNER_TAB_NAME = "Budget Planner"