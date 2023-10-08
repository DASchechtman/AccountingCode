type CheckPayOut = (date: Date, dir: number, inc: number) => boolean;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Tab = GoogleAppsScript.Spreadsheet.Sheet;
type DataArray = Array<Array<string | number>>;
type DataArrayEntry = Array<string | number>;

const PURCHASE_HEADER = "Purchases for"
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]