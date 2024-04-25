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
    private readonly PARSER: Parser
    private readonly TAB: GoogleSheetTabs
    private readonly INTERPRET_ACTION: Map<__SFI_ParserType, (state: ParserState) => Maybe>
    private readonly None: None = { type: "None" }

    constructor(tab: string | GoogleSheetTabs) {
        this.PARSER = new Parser(__SFI_CreateFormulaParser())
        this.INTERPRET_ACTION = new Map()
        this.InitInterpretActions()

        if (typeof tab === "string") {
            this.TAB = new GoogleSheetTabs(tab)
        }
        else {
            this.TAB = tab
        }
    }

    public ParseInput(input: string) {
        let x = this.PARSER.Run(input)
        return x
    }

    private InitInterpretActions() {
        const Default = (state: ParserState) => { return state }

        this.INTERPRET_ACTION.set('OP_ADD', (state) => {
            const ALLOWED_TYPES = ['NUMBER', 'FLOAT', 'INT']
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[0].type)) { return this.None }
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[1].type)) { return this.None }

            const LEFT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[0]), 0)
            const RIGHT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[1]), 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }

            return this.WrapValue(LEFT + RIGHT)
        })

        this.INTERPRET_ACTION.set('OP_SUB', (state) => {
            const ALLOWED_TYPES = ['NUMBER', 'FLOAT', 'INT']
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[0].type)) { return this.None }
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[1].type)) { return this.None }

            const LEFT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[0]), 0)
            const RIGHT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[1]), 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }
            return this.WrapValue(LEFT - RIGHT)
        })

        this.INTERPRET_ACTION.set('OP_MUL', (state) => {
            const ALLOWED_TYPES = ['NUMBER', 'FLOAT', 'INT']
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[0].type)) { return this.None }
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[1].type)) { return this.None }

            const LEFT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[0]), 0)
            const RIGHT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[1]), 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }
            return this.WrapValue(LEFT * RIGHT)
        })

        this.INTERPRET_ACTION.set('OP_DIV', (state) => {
            const ALLOWED_TYPES = ['NUMBER', 'FLOAT', 'INT']
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[0].type)) { return this.None }
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[1].type)) { return this.None }

            const LEFT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[0]), 1)
            const RIGHT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[1]), 1)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }
            if (RIGHT === 0) { return this.None }

            return this.WrapValue(LEFT / RIGHT)
        })

        this.INTERPRET_ACTION.set('OP_POW', (state) => {
            const ALLOWED_TYPES = ['NUMBER', 'FLOAT', 'INT']
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[0].type)) { return this.None }
            if (!ALLOWED_TYPES.includes(state.result.child_nodes[1].type)) { return this.None }

            const LEFT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[0]), 0)
            const RIGHT = this.UnwrapValueOrDefault(this.InterpretNode(state.result.child_nodes[1]), 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }
            return this.WrapValue(Math.pow(LEFT, RIGHT))
        })

        this.INTERPRET_ACTION.set('OP_PAREN', (state) => {
            return this.InterpretNode(state.result.child_nodes[0])
        })

        this.INTERPRET_ACTION.set('FUNCTION', (state) => {
            const RunFunction = (func: __SFI_ParserType) => {
                return this.INTERPRET_ACTION.get(func)?.call(this, state) || this.None
            }

            switch (state.result.res) {
                case 'MULTIPLY': { return RunFunction('FUNC_MUL') }
                case 'DIVIDE':   { return RunFunction('FUNC_DIV') }
                case 'ADD':      { return RunFunction('FUNC_ADD') }
                case 'SUBTRACT': { return RunFunction('FUNC_SUB') }
                case 'POWER':    { return RunFunction('FUNC_POW') }
                case 'SUM':      { return RunFunction('FUNC_SUM') }
            }
            return this.None
        })

        this.INTERPRET_ACTION.set('FUNC_MUL', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                vals.push(val)
            }

            if (vals.length !== 2) { return this.None }
            const LEFT = this.UnwrapValueOrDefault(vals[0], 0)
            const RIGHT = this.UnwrapValueOrDefault(vals[1], 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }

            return this.WrapValue(LEFT * RIGHT)
        })

        this.INTERPRET_ACTION.set('FUNC_DIV', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                vals.push(val)
            }

            if (vals.length !== 2) { return this.None }
            const LEFT = this.UnwrapValueOrDefault(vals[0], 1)
            const RIGHT = this.UnwrapValueOrDefault(vals[1], 1)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }
            if (RIGHT === 0) { return this.None }

            return this.WrapValue(LEFT / RIGHT)
        })

        this.INTERPRET_ACTION.set('FUNC_ADD', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                vals.push(val)
            }

            if (vals.length !== 2) { return this.None }
            const LEFT = this.UnwrapValueOrDefault(vals[0], 0)
            const RIGHT = this.UnwrapValueOrDefault(vals[1], 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }

            return this.WrapValue(LEFT + RIGHT)
        })

        this.INTERPRET_ACTION.set('FUNC_SUB', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                vals.push(val)
            }

            if (vals.length !== 2) { return this.None }
            const LEFT = this.UnwrapValueOrDefault(vals[0], 0)
            const RIGHT = this.UnwrapValueOrDefault(vals[1], 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }

            return this.WrapValue(LEFT - RIGHT)
        })

        this.INTERPRET_ACTION.set('FUNC_POW', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                vals.push(val)
            }

            if (vals.length !== 2) { return this.None }
            const LEFT = this.UnwrapValueOrDefault(vals[0], 0)
            const RIGHT = this.UnwrapValueOrDefault(vals[1], 0)

            if (typeof LEFT !== "number" || typeof RIGHT !== "number") { return this.None }

            return this.WrapValue(Math.pow(LEFT, RIGHT))
        })

        this.INTERPRET_ACTION.set('FUNC_SUM', (state) => {
            let vals = new Array<Maybe>()
            for (let child of state.result.child_nodes) {
                let val = this.InterpretNode(child)
                if (val.type === "None") { return this.None }
                if (val.val instanceof Array) {
                    vals.push(...val.val.map(v => this.WrapValue(v)))
                }
                else {
                    vals.push(val)
                }
            }

            return this.WrapValue(vals.reduce((acc, val) => {
                let unwrapped = this.UnwrapValueOrDefault(val, 0)
                if (typeof unwrapped !== "number") { return acc }
                return acc + unwrapped
            }, 0))
        })

        this.INTERPRET_ACTION.set('INT', state => {
            let val = parseInt(state.result.res)
            if (isNaN(val)) { return this.None }
            return this.WrapValue(val)
        })

        this.INTERPRET_ACTION.set('FLOAT', state => {
            let val = parseFloat(state.result.res)
            if (isNaN(val)) { return this.None }
            return this.WrapValue(val)
        })

        this.INTERPRET_ACTION.set('NUMBER', state => {
            let val = Number(state.result.res)
            if (isNaN(val)) { return this.None }
            return this.WrapValue(val)
        
        })

        this.INTERPRET_ACTION.set('STRING', state => this.WrapValue(state.result.res))

        this.INTERPRET_ACTION.set('BOOLEAN', state => {
            let val = state.result.res.toLowerCase()
            if (val === "true") { return this.WrapValue(true) }
            if (val === "false") { return this.WrapValue(false) }
            return this.None
        })

        this.INTERPRET_ACTION.set('DATE', state => {
            let val = new Date(state.result.res)
            if (isNaN(val.getTime())) { return this.None }
            return this.WrapValue(val)
        })
        
        this.INTERPRET_ACTION.set('SPREADSHEET_CELL', (state) => {
            let col = state.result.res.match(/[A-Za-z]+/g)?.[0]
            let row = state.result.res.match(/\d+/g)?.[0]

            if (col === undefined) { return this.None }
            if (row === undefined) { return this.None }

            let col_index = __Util_ColLetterToIndex(col)
            let row_index = parseInt(row) - 1
            let cell_val = this.TAB.GetRow(row_index)?.[col_index]

            if (cell_val === undefined) { return this.None }

            if (typeof cell_val === "string") {
                let parse_attempt = this.PARSER.Run(cell_val)
                if (parse_attempt.is_error) { return this.WrapValue(cell_val) }
                return this.InterpretNode(parse_attempt)
            }

            return this.WrapValue(cell_val)
        })

        this.INTERPRET_ACTION.set('SPREADSHEET_RANGE', (state) => {
            let [start_cell, end_cell] = state.result.res.split(":")

            const COL = /[A-Za-z]/g
            const ROW = /\d+/g

            let start_col = start_cell.match(COL)?.[0]
            let start_row = start_cell.match(ROW)?.[0]
            let end_col = end_cell.match(COL)?.[0]
            let end_row = end_cell.match(ROW)?.[0]

            if (
                start_col === undefined 
                || start_row === undefined 
                || end_col === undefined 
                || end_row === undefined
            ) { return this.None }
            
            let start_col_index = __Util_ColLetterToIndex(start_col)
            let start_row_index = parseInt(start_row) - 1
            let end_col_index = __Util_ColLetterToIndex(end_col)
            let end_row_index = parseInt(end_row) - 1

            let range_vals = new Array<Array<unknown>>()

            for (let i = start_row_index; i <= end_row_index; i++) {
                let vals = new Array<unknown>()
                for (let j = start_col_index; j <= end_col_index; j++) {
                    const VAL = this.TAB.GetRow(i)?.[j]
                    if (VAL === undefined) { return this.None }

                    if (typeof VAL === "string") {
                        let parse_attempt = this.PARSER.Run(VAL)
                        if (parse_attempt.is_error) { 
                            range_vals.push([VAL])
                            continue
                        }
                        let interp = this.UnwrapValueOrNone(this.InterpretNode(parse_attempt))
                        if (interp.type === "None") { return this.None }
                        vals.push(interp.val)
                    }
                    else {
                        vals.push(VAL)
                    }
                }
                range_vals.push(vals)
            }

            return this.WrapValue(range_vals)
        })
    }

    private InterpretNode(node: ParserState) {
        return this.INTERPRET_ACTION.get(node.result.type)?.call(this, node) || this.None
    }

    private WrapValue(value: any): Some {
        return { type: "Some", val: value }
    }

    private UnwrapValueOrDefault(value: Maybe, default_val: unknown) {
        if (value.type === "None") { return default_val }
        return value.val
    }

    private UnwrapValueOrNone(value: Maybe) {
        if (value.type === "None") { return this.None }
        return value
    }
}

function UtilsClassMain() {
    let y = new FormulaInterpreter(PAYMENT_SCHEDULE_TAB_NAME).ParseInput("=5true-5")
    console.log(y.toString())
}