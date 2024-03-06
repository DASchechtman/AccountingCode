class ParserState {
    private target_str: string
    private parse_error_msg: string
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
        this.parse_error_msg = ""
        this.index = 0
        this.result = {res: "", extras: [], child_nodes: []}
        this.type = ""
        
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
                this.parse_error_msg = args[0].parse_error_msg
                this.type = args[0].type
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
        if (parse_error_msg) { NEW_STATE.parse_error_msg = parse_error_msg }
        if (type) { NEW_STATE.type = type }
        return NEW_STATE
    }

    public CloneIndexOnly() {
        return this.CreatePartialClone({index: this.index})
    }

    public Map(func: (state: ParserStateResults) => ParserStateResults) {
        
    }

    public toString() {
        return JSON.stringify(this, null, 2)
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

    public Map(func: (state: ParserStateResults) => ParserStateResults) {
        return new Parser((state: ParserState) => {
            const NEW_STATE = this.ParserFunc(state)
            if (NEW_STATE.is_error) { return NEW_STATE }
            NEW_STATE.result = func(NEW_STATE.result)
            return NEW_STATE
        })
    }
}