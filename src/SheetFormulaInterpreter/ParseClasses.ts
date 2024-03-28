class ErrorMessage {
    private error: string = ""
    public get message() { return this.error }
    public set message(value: string) { this.error = value }
    public toString() { return this.error }
}

class ParserState {
    private parse_error_msg: string = ""
    private static to_str_indent: number = 0

    private target_str: string
    public index: number
    public result: ParserStateResults
    public type: __SFI_ParserType
    
    public get target() { return this.target_str.slice(this.index) }
    public get full_target() { return this.target_str }
    public get parser_error() { return this.parse_error_msg }
    public get is_error() { return this.parser_error.length > 0 }

    public set parser_error(value: string) { this.parse_error_msg = value }
    public set is_error(value: boolean) {
        if (value && this.parser_error.length === 0) {
            this.parser_error = "An unknown error occurred."
        }
        else if (!value) {
            this.parser_error = ""
        }
    }

    public constructor();
    public constructor(target: string);
    public constructor(target: ParserState);
    public constructor(...args: any[]) {
        this.target_str = ""
        this.index = 0
        this.result = {res: "", extras: [], child_nodes: []}
        this.type = ""
        this.parse_error_msg = ""
        
        if (args.length === 0) {
            return
        }
        else if (args.length === 1) {
            if (typeof args[0] === "string") {
                this.target_str = args[0]
            }
            else if (args[0] instanceof ParserState) {
                this.target_str = args[0].target_str
                this.index = args[0].index
                this.result = this.CloneResults(args[0].result)
                this.type = args[0].type
                this.parse_error_msg = args[0].parse_error_msg
            }
        }
    }

    public Clone() {  
        return new ParserState(this)
    }

    public CreateEmptyState() {
        return new ParserState(this.target_str)
    }

    public CreatePartialClone({index, result, parse_error_msg, type}: {index?: number, result?: ParserStateResults, parse_error_msg?: string, type?: __SFI_ParserType}) {
        const NEW_STATE = this.CreateEmptyState()
        if (index) { NEW_STATE.index = index }
        if (result) { NEW_STATE.result = this.CloneResults(result) }
        if (parse_error_msg) { this.parse_error_msg = parse_error_msg }
        if (type) { NEW_STATE.type = type }
        return NEW_STATE
    }

    public CloneIndexOnly() {
        return this.CreatePartialClone({index: this.index})
    }

    public toString() {
        const INDENT = this.CreateIndent()
        const INDENT_OFFSET = this.CreateIndent(2)
        const STR_ARR = new Array<string>()
        const CHILD_STR = this.StringifyChildNodes()

        STR_ARR.push(`${INDENT}{\n`)
        STR_ARR.push(`${INDENT_OFFSET}target: ${this.target_str},\n`)
        STR_ARR.push(`${INDENT_OFFSET}index: ${this.index},\n`)
        STR_ARR.push(`${INDENT_OFFSET}result: "${this.result.res}",\n`)
        STR_ARR.push(`${INDENT_OFFSET}error: "${this.parse_error_msg}",\n`)
        STR_ARR.push(`${INDENT_OFFSET}extras: ${this.result.extras.length > 0 ? this.result.extras : '[]'},\n`)
        STR_ARR.push(`${INDENT_OFFSET}child_nodes: [${CHILD_STR.length > 0 ? '\n'+CHILD_STR+`${INDENT_OFFSET}]` : ']'}\n`)
        STR_ARR.push(`${INDENT}}\n`)

        return STR_ARR.join("")
    }

    private StringifyChildNodes() {
        ParserState.to_str_indent += 4
        const STR_ARR = new Array<string>()
        let i = 0
        let len = this.result.child_nodes.length

        for (const node of this.result.child_nodes) {
            STR_ARR.push(`${node.toString().slice(0, -1)}${++i === len ? '' : ','}\n`)
        }

        ParserState.to_str_indent -= 4
        return STR_ARR.join("")
    }

    private CreateIndent(offset: number = 0) {
        return " ".repeat(ParserState.to_str_indent + offset)
    }

    private CloneResults(results: ParserStateResults) {
        return {
            res: results.res,
            extras: results.extras.slice(),
            child_nodes: results.child_nodes.map((node) => node.Clone())
        }
    }
}

class Parser {
    private ParserFunc: __SFI_ParserFunc

    public get func() { return this.ParserFunc }

    constructor(ParserFunc: __SFI_ParserFunc) {
        this.ParserFunc = ParserFunc
    }

    public Run(target: string | ParserState) {
        const STATE = typeof target === "string" ? new ParserState(target) : target
        return this.ParserFunc(STATE)
    }

    public Map(transform: (state: ParserStateResults) => ParserStateResults) {
        return new Parser((state: ParserState) => {
            const NEW_STATE = this.ParserFunc(state)
            if (NEW_STATE.is_error) { return NEW_STATE }
            NEW_STATE.result = transform(NEW_STATE.result)
            return NEW_STATE
        })
    }
}