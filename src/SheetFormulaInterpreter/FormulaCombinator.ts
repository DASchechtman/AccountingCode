function __SFI_ConvertFormulaToTree(formula: Array<ParserState>, opers: Array<string>) {
    for (let i = 0; i < formula.length; i++) { 
        const EL = formula[i]

        if (opers.includes(EL.result.res)) {
            const LEFT = formula[i-1]
            const RIGHT = formula[i+1]
            EL.result.child_nodes.push(LEFT, RIGHT)
            formula.splice(formula.indexOf(LEFT), 1)
            formula.splice(formula.indexOf(RIGHT), 1)
            i = formula.indexOf(EL)
        }
    }
}

function __SFI_CreateFormulaParser() {
    const __SFI_DateSegParser = __SFI_Regex(/[0-9]{1,2}/)
    const __SFI_WhiteSpaceParser = __SFI_ManyZero(__SFI_Str(" "))

    const __SFI_DateParser = new Parser(__SFI_SequenceOf(__SFI_DateSegParser, __SFI_Str("/"), __SFI_DateSegParser, __SFI_Str("/"), __SFI_DateSegParser, __SFI_DateSegParser)).Map(state => {
        const DATE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: DATE,
            extras: ["TYPE", "date"],
            child_nodes: [],
        }
    })

    const __SFI_CellDataParser = __SFI_Choice(__SFI_DateParser, __SFI_Int, __SFI_Float, __SFI_Bool, __SFI_Letters)
    const __SFI_CellParser = new Parser(__SFI_SequenceOf(__SFI_Letters, __SFI_Int)).Map(state => {
        const CELL = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL,
            extras: ["TYPE", "cell"],
            child_nodes: [],
        }
    })
    const __SFI_CellRangeParser = new Parser(__SFI_SequenceOf(__SFI_CellParser, __SFI_Str(":"), __SFI_CellParser)).Map(state => {
        const CELL_RANGE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL_RANGE,
            extras: ["TYPE", "cell_range", "FROM", state.child_nodes[0].result.res, "TO", state.child_nodes[2].result.res],
            child_nodes: [],
        }
    })


    const __SFI_FormulaArg = __SFI_Choice(__SFI_CellParser, __SFI_CellDataParser)
    const __SFI_OperParser = __SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/"))
    const __SFI_FormulaParam = new Parser(__SFI_ManyOne(__SFI_SequenceOf(__SFI_FormulaArg, __SFI_WhiteSpaceParser, __SFI_OperParser))).Map(state => {
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
    const __SFI_FormulaParser = new Parser(__SFI_SequenceOf(__SFI_Str("="), __SFI_FormulaParam, __SFI_FormulaArg)).Map(state => {
        let children = [...state.child_nodes[1].result.child_nodes, state.child_nodes[2]]
        __SFI_ConvertFormulaToTree(children, ["*", "/"])
        __SFI_ConvertFormulaToTree(children, ["+", "-"])
        return {
            res: "",
            extras: ["TYPE", "formula"],
            child_nodes: children
        }
    })

    const __SFI_FuncArgParser = __SFI_Choice(__SFI_CellRangeParser, __SFI_CellParser, __SFI_CellDataParser)
    const __SFI_FuncFormulaParam = __SFI_ManyOne(__SFI_SequenceOf(__SFI_FuncArgParser, __SFI_Str(","), __SFI_WhiteSpaceParser))
    const __SFI_FuncFormulaParamList = __SFI_SequenceOf(__SFI_FuncFormulaParam, __SFI_FuncArgParser)
    const __SFI_FuncFormula = __SFI_SequenceOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), __SFI_Choice(__SFI_FuncFormulaParamList, __SFI_FuncArgParser), __SFI_Str(")"))

    return __SFI_Choice(__SFI_FuncFormula, __SFI_FormulaParser)
}