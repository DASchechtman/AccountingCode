class PayDay {
    private pay_out_amt: number;
    private pay_date: Date;
    private ShouldPayOut: CheckPayOut;
    private months: string[]
    private day_inc = 7;
    private total_days = 0;

    constructor(pay_out_amt: number, pay_date: Date, ShouldPayOut: CheckPayOut) {
        this.pay_out_amt = pay_out_amt;
        this.ShouldPayOut = ShouldPayOut;
        this.pay_date = new Date(pay_date);
        this.months = MONTHS
    }

    public SetPayoutDate(PayOutDate: (date: Date) => Date) {
        this.pay_date = PayOutDate(this.pay_date);
    }

    public SetPayoutAmount(pay_out_amt: number) {
        this.pay_out_amt = pay_out_amt;
    }

    public SetPayoutCheck(ShouldPayOut: CheckPayOut) {
        this.ShouldPayOut = ShouldPayOut;
    }

    public PayOut() {
        let pay_amt = this.pay_out_amt;
        const SHOULD_PAY = this.ShouldPayOut({
            date: this.pay_date,
            total_days: this.total_days,
            inc: this.day_inc,
            pay_month: this.PayMonth()
        })

        if (!SHOULD_PAY) {
            pay_amt = 0;
        }

        this.pay_date.setUTCDate(this.pay_date.getUTCDate() + this.day_inc);
        this.total_days += this.day_inc;
        return pay_amt;
    }

    public PayMonth() {
        return this.months[this.GetMonthIndex()];
    }

    public GetDate() {
        return new Date(this.pay_date.getTime())
    }

    private GetMonthIndex() {
        let month = this.pay_date.getUTCMonth();
        const MONTH_DAY = this.pay_date.getUTCDate();
        if (MONTH_DAY >= 28) {
            month = (month + 1) % this.months.length;
        }
        return month
    }
}

class GoogleSheetTabs {
    private tab: Tab
    private headers: Map<string, number>
    private data: DataArray

    constructor(tab: Tab | string) {
        if (typeof tab === "string") {
            const SHEET_TAB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab)
            if (SHEET_TAB === null) { throw new Error("Tab does not exist") }
            tab = SHEET_TAB
        }

        this.tab = tab
        this.data = []
        this.InitSheetData()

        const HEADERS = this.data[0]

        this.headers = new Map<string, number>()
        for (let i = 0; i < HEADERS.length; i++) {
            const HEADER = HEADERS[i]
            if (typeof HEADER !== "string") { continue }
            this.headers.set(HEADER, i)
        }
    }

    public GetHeaderIndex(header_name: string) {
        return this.headers.get(header_name) === undefined ? -1 : this.headers.get(header_name)!
    }

    public GetHeaderNames() {
        return Array.from(this.headers.keys())
    }

    public GetCol(header_name: string) {
        const COL: DataArrayEntry = []
        const COL_INDEX = this.headers.get(header_name)

        if (COL_INDEX === undefined) { return undefined }

        for (let i = 0; i < this.data.length; i++) {
            COL.push(this.data[i][COL_INDEX])
        }

        return COL
    }

    public GetColByIndex(col_index: number) {
        if (col_index < 0 || col_index >= this.data[0].length) { return undefined }
        const COL: DataArrayEntry = []
        for (let i = 0; i < this.data.length; i++) {
            COL.push(this.data[i][col_index])
        }
        return COL
    }

    public WriteCol(header_name: string, col: DataArrayEntry) {
        const COL_INDEX = this.headers.get(header_name)
        if (COL_INDEX === undefined) { return }
        const LONGEST_ROW = this.FindLongestRowLength()

        for (let i = col.length - 1; i >= 0; i--) {
            if (this.data[i] === undefined) { this.data[i] = new Array(LONGEST_ROW).fill("") }
            this.data[i][COL_INDEX] = col[i]
        }
    }

    public GetRow(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        return this.CreateRowCopy(this.data[row_index])
    }

    public WriteRow(row_index: number, row: DataArrayEntry) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        this.data[row_index] = this.CreateRowCopy(row)
    }

    public WriteRowAt(row_index: number, start: number, row: DataArrayEntry) {
        if (row_index < 0 || row_index >= this.data.length) { return }
        if (start < 0) { start = 0 }
        while (start + row.length >= this.data[row_index].length) { this.data[row_index].push("") }

        for (let i = 0; i < row.length; i++) {
            this.data[row_index][start + i] = row[i]
        }
    }

    public AppendRow(row: DataArrayEntry, should_fill: boolean = false) {
        row = this.CreateRowCopy(row)
        this.data.push(row)
        if (should_fill) {
            const LONGEST_ROW = this.FindLongestRowLength()
            while (row.length < LONGEST_ROW) {
                row.push("")
            }
        }
        return row
    }

    public InsertRow(row_index: number, row: DataArrayEntry, { AlterRow, should_fill }: {
        AlterRow?: (row: DataArrayEntry) => DataArrayEntry,
        should_fill?: boolean
    } = {}) {
        if (row_index < 0) { row_index = 0 }
        row = this.CreateRowCopy(row)
        if (AlterRow) { row = AlterRow(row) }

        const LONGEST_ROW = this.FindLongestRowLength()
        while (row.length < LONGEST_ROW && should_fill) {
            row.push("")
        }

        if (row_index >= this.data.length) { return this.AppendRow(row) }
        this.data.splice(row_index, 0, row)

        return row
    }

    public AppendToRow(row_index: number, ...row: DataArrayElement[]) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        this.data[row_index].push(...row.map(__Util_ConvertToStrOrNum))
        return row
    }

    public FindRow(func: (row: DataArrayEntry) => boolean) {
        return this.data.find(func)
    }

    public IndexOfRow(row?: DataArrayEntry | ((row: DataArrayEntry) => boolean), index_from?: number) {
        let search_row = row
        if (typeof search_row === "function") { search_row = this.FindRow(search_row) }
        if (search_row === undefined) { return -1 }
        return this.data.indexOf(search_row, index_from)
    }

    public GetRowRange(row_index: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }
        const RANGE_NOTATION = `A${row_index + 1}:${__Util_IndexToColLetter(this.data[row_index].length)}${row_index + 1}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public GetRowSubRange(row_index: number, start: number, end: number) {
        if (row_index < 0 || row_index >= this.data.length) { return undefined }

        if (start > end || end < start) { start = end }
        if (start < 0) { start = 0 }
        if (end < 0) { end = 0 }

        const RANGE_NOTATION = `${__Util_IndexToColLetter(start)}${row_index + 1}:${__Util_IndexToColLetter(end)}${row_index + 1}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public GetRange(start_row: number, end_row: number, start_col: number, end_col: number) {
        const RANGE_1 = this.GetRowSubRange(start_row, start_col, end_col)
        const RANGE_2 = this.GetRowSubRange(end_row, start_col, end_col)

        if (RANGE_1 === undefined || RANGE_2 === undefined) { return undefined }
        const FIRST_NOTATION_PART = RANGE_1.getA1Notation().split(":")[0]
        const SECOND_NOTATION_PART = RANGE_2.getA1Notation().split(":")[1]
        const RANGE_NOTATION = `${FIRST_NOTATION_PART}:${SECOND_NOTATION_PART}`
        return this.tab.getRange(RANGE_NOTATION)
    }

    public NumberOfRows() {
        return this.data.length
    }

    public SaveToTab() {
        this.SetAllRowsToSameLength()
        const WRITE_RANGE = this.tab.getRange(1, 1, this.data.length, this.data[0].length)
        WRITE_RANGE.setValues(this.data)
    }

    public GetTab() {
        return this.tab
    }

    public CopyTo(tab: GoogleSheetTabs) {
        for (let i = 0; i < this.data.length; i++) {
            if (i >= tab.NumberOfRows()) {
                tab.AppendRow(this.data[i])
            }
            else {
                tab.WriteRow(i, this.data[i])
            }

            const ROW_RANGE = this.GetRowRange(i)
            const TAB_ROW_RANGE = tab.GetRowRange(i)
            if (ROW_RANGE === undefined || TAB_ROW_RANGE === undefined) { continue }
            ROW_RANGE.copyTo(TAB_ROW_RANGE, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
        }

        for (let i = 0; i < tab.data[0].length; i++) {
            tab.GetTab().autoResizeColumn(i + 1)
            const width = tab.GetTab().getColumnWidth(i + 1)
            tab.GetTab().setColumnWidth(i + 1, width + 25)
        }
    }

    public ClearTab() {
        this.data.map(row => row.fill(""))
    }

    public EraseTab() {
        this.data = []
    }

    private FindLongestRowLength() {
        let longest_row = -1
        for (let i = 0; i < this.data.length; i++) {
            if (this.data[i].length > longest_row) {
                longest_row = this.data[i].length
            }
        }
        return longest_row
    }

    private SetAllRowsToSameLength() {
        const LONGEST_ROW = this.FindLongestRowLength()
        for (let i = 0; i < this.data.length; i++) {
            while (this.data[i].length < LONGEST_ROW) {
                this.data[i].push("")
            }
            this.data[i] = this.data[i].map(__Util_ConvertToStrOrNum)
        }
    }

    private CreateRowCopy(row: any[]) {
        return [...row].map(__Util_ConvertToStrOrNum)
    }

    private InitSheetData() {
        const RANGE_DATA = this.tab.getDataRange().getValues().map(row => row.map(__Util_ConvertToStrOrNum))
        this.data = this.tab.getDataRange().getFormulas()

        for (let row = 0; row < RANGE_DATA.length; row++) {
            for (let col = 0; col < RANGE_DATA[row].length; col++) {
                if (this.data[row][col] !== "") { continue }
                this.data[row][col] = RANGE_DATA[row][col]
            }
        }

    }
}

class FormulaInterpreter {
    private parser: Parser

    constructor() {
        const CELL_PARSER = new Parser(__SFI_SequenceOf(__SFI_Letters, __SFI_Int)).Map(state => {
            const [letters, int] = [state.child_nodes[0].result.res, state.child_nodes[1].result.res]
            return {
                res: `${letters}${int}`.toUpperCase(),
                extras: [],
                child_nodes: []
            }
        })
        const CELL_RANGE_PARSER = new Parser(__SFI_SequenceOf(CELL_PARSER, __SFI_Str(":"), CELL_PARSER)).Map(state => {
            const Res = (index: number) => state.child_nodes[index].result.res
            return {
                res: `${Res(0)}:${Res(2)}`,
                extras: [],
                child_nodes: []
            }
        })
        const DATE_SEG_PARSER = __SFI_Regex(/[0-9]{2,2}/)
        const DATE_PARSER = new Parser(__SFI_SequenceOf(DATE_SEG_PARSER, __SFI_Str("/"), DATE_SEG_PARSER, __SFI_Str("/"), DATE_SEG_PARSER, DATE_SEG_PARSER)).Map(state => {
            const Res = (index: number) => state.child_nodes[index].result.res

            return {
                res: `${Res(0)}/${Res(2)}/${Res(4)}${Res(5)}`,
                extras: [],
                child_nodes: []
            }
        })
        const DATA_PARSER = new Parser(__SFI_Choice(DATE_PARSER, __SFI_Bool, __SFI_Float, __SFI_Int, __SFI_Letters))
        const CELL_DATA_PARSER = new Parser(__SFI_Choice(CELL_RANGE_PARSER, CELL_PARSER, DATA_PARSER))
        const OPER_PARSER = new Parser(__SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/")))
        const SPACES = __SFI_ManyZero(__SFI_Str(" "))

        const FORMULA_CHUNK_PARSER = __SFI_ManyOne(__SFI_SequenceOf(CELL_DATA_PARSER, OPER_PARSER))
        const FORMULA_PARSER = new Parser(__SFI_SequenceOf(__SFI_Str("="), FORMULA_CHUNK_PARSER, CELL_DATA_PARSER)).Map(state => {
            const OPERS = state.child_nodes[1].result.child_nodes
            const OPER_ARR = new Array<ParserState>()

            for (const CHILD of OPERS) {
                OPER_ARR.push(CHILD.result.child_nodes[0], CHILD.result.child_nodes[1])
            }

            OPER_ARR.push(state.child_nodes[2])

            let str_oper = OPER_ARR.map(x => x.result.res)

            for (let i = 1; i < str_oper.length; i++) {
                let cur_char = str_oper[i]
                let prev_char = str_oper[i - 1]
                let res1 = /^[0-9]+$/.test(cur_char)
                let res2 = /^[+\-*/]$/.test(prev_char)

                if (res1 && res2) {
                    let tmp = str_oper[i]
                    str_oper[i] = str_oper[i - 1]
                    str_oper[i - 1] = tmp
                }
            }

            return {
                res: str_oper.join(""),
                extras: [],
                child_nodes: OPER_ARR
            }
        })

        const FUNC_PARAMS_PARSER = __SFI_ManyOne(__SFI_SequenceOf(CELL_DATA_PARSER, SPACES, __SFI_Str(","), SPACES))
        const FUNC_FORMULA_PARSER = __SFI_SequenceOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), FUNC_PARAMS_PARSER,  CELL_DATA_PARSER, __SFI_Str(")"))

        this.parser = new Parser(__SFI_Choice(FUNC_FORMULA_PARSER, FORMULA_PARSER))
    }

    public ParseInput(input: string) {
        let x = this.parser.Run(input)
        return x
    }
}

function UtilsClassMain() {
    let y = new FormulaInterpreter().ParseInput("=5-5")
    console.log(y.toString())
}