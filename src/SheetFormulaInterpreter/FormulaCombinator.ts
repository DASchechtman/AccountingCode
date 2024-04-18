function __SFI_CreateFormulaTree(formula: Array<ParserState>) {
    __SFI_ConvertFormulaToTree(formula, ["^"])
    __SFI_ConvertFormulaToTree(formula, ["*", "/"])
    __SFI_ConvertFormulaToTree(formula, ["+", "-"])
    __SFI_ConvertFormulaToTree(formula, ["(", ")"])
}

function __SFI_ConvertFormulaToTree(formula: Array<ParserState>, opers: Array<string>) {
    for (let i = 0; i < formula.length; i++) { 
        const EL = formula[i]
        
        const IS_PAREN_WRAPPED = EL.result.res === '(' && opers.length == 2 && opers[0] === '(' && opers[1] === ')'
        if (IS_PAREN_WRAPPED) {
            let j = i
            while (formula[j].result.res !== ')') { j++ }
            const SUB_ARR = formula.slice(i+1, j)
            const PAREN_EL = SUB_ARR[1]
            PAREN_EL.result.res = '()'
            PAREN_EL.result.child_nodes = SUB_ARR
            formula.splice(i, j-i+1, PAREN_EL)
        }
        else if (opers.includes(EL.result.res)) {
            const LEFT = formula[i-1]
            const RIGHT = formula[i+1]
            EL.result.child_nodes.push(LEFT, RIGHT)
            formula.splice(formula.indexOf(LEFT), 1)
            formula.splice(formula.indexOf(RIGHT), 1)
            i = formula.indexOf(EL)
        }
    }

    return formula
}

function __SFI_CreateFormulaParser() {
    const DateSegParser = __SFI_Regex(/[0-9]{1,2}/)
    const WhiteSpaceParser = __SFI_ManyZero(__SFI_Str(" "))

    const DateParser = new Parser(__SFI_SequenceOf(DateSegParser, __SFI_Str("/"), DateSegParser, __SFI_Str("/"), DateSegParser, DateSegParser)).Map(state => {
        const DATE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: DATE,
            extras: ["TYPE", "date"],
            child_nodes: [],
        }
    })

    const CellDataParser = __SFI_Choice(DateParser, __SFI_Int, __SFI_Float, __SFI_Bool, __SFI_Letters)
    const CellParser = new Parser(__SFI_SequenceOf(__SFI_Letters, __SFI_Int)).Map(state => {
        const CELL = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL,
            extras: ["TYPE", "cell"],
            child_nodes: [],
        }
    })
    const CellRangeParser = new Parser(__SFI_SequenceOf(CellParser, __SFI_Str(":"), CellParser)).Map(state => {
        const CELL_RANGE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL_RANGE,
            extras: ["TYPE", "cell_range", "FROM", state.child_nodes[0].result.res, "TO", state.child_nodes[2].result.res],
            child_nodes: [],
        }
    })


    const FormulaArg = __SFI_Choice(CellParser, CellDataParser)
    const OperParser = __SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/"))
    const FormulaParam = new Parser(__SFI_ManyOne(__SFI_SequenceOf(FormulaArg, WhiteSpaceParser, OperParser))).Map(state => {
        let children = new Array<ParserState>()

        for (const CHILD of state.child_nodes) {
            children.push(CHILD.result.child_nodes[0], CHILD.result.child_nodes[2])
        }

        return {
            res: "",
            extras: [],
            child_nodes: children
        }
    })
    const FormulaParser = new Parser(__SFI_SequenceOf(__SFI_Str("="), FormulaParam, FormulaArg)).Map(state => {
        let children = [...state.child_nodes[1].result.child_nodes, state.child_nodes[2]]
        __SFI_CreateFormulaTree(children)
        return {
            res: "",
            extras: ["TYPE", "formula"],
            child_nodes: children
        }
    })

    const FuncArgParser = __SFI_Choice(CellRangeParser, CellParser, CellDataParser)
    const FuncFormulaParam = __SFI_ManyOne(__SFI_SequenceOf(FuncArgParser, __SFI_Str(","), WhiteSpaceParser))
    const FuncFormulaParamList = __SFI_SequenceOf(FuncFormulaParam, FuncArgParser)
    const FuncFormula = __SFI_SequenceOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), __SFI_Choice(FuncFormulaParamList, FuncArgParser), __SFI_Str(")"))

    return __SFI_Choice(FuncFormula, FormulaParser)
}

function __SFI_ParseFormulaMain() {
    const NumberParser = new Parser(__SFI_Int).Map(state => {
        return {
            res: state.res,
            extras: ["TYPE", "number", "VALUE", state.res],
            child_nodes: []
        }
    
    })
    const LetterParser = new Parser(__SFI_Letters).Map(state => {
        return {
            res: state.res,
            extras: ["TYPE", "string", "VALUE", state.res],
            child_nodes: [state.ReconstructState()],
        }
    })

    const DiceRollParser = new Parser(__SFI_SequenceOf(NumberParser, __SFI_Str("d"), NumberParser)).Map(state => {
        return {
            res: state.child_nodes.map(node => node.result.res).join(""),
            extras: ["TYPE", "dice_roll", "DICE", state.child_nodes[0].result.res, "SIDES", state.child_nodes[2].result.res],
            child_nodes: state.child_nodes,
        }
    })

    const MainParser = new Parser(__SFI_SequenceOf(LetterParser, __SFI_Str(":"))).Map(state => {
            return { 
                res: state.child_nodes[0].result.res, 
                extras: [], 
                child_nodes: [],
            }
        })
        .Chain(res_state => {
            return new Parser(state => {
                if (res_state.res === 'string') { return LetterParser.Run(state) }
                else if (res_state.res === 'number') { return NumberParser.Run(state) }
                return DiceRollParser.Run(state)
            }).Map(state => {
                return {
                    ...res_state,
                    child_nodes: [state.ReconstructState()],
                }
            })
        })
    
    console.log(MainParser.Run("string:hello").toString())
}