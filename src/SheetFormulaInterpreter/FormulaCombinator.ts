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
            const SUB_ARR = formula.slice(i + 1, j)
            const PAREN_EL = SUB_ARR[1]
            PAREN_EL.result.res = '()'
            PAREN_EL.result.child_nodes = SUB_ARR
            formula.splice(i, j - i + 1, PAREN_EL)
        }
        else if (opers.includes(EL.result.res)) {
            const LEFT = formula[i - 1]
            const RIGHT = formula[i + 1]
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

    const DateParser = new Parser(__SFI_SeqOf(DateSegParser, __SFI_Str("/"), DateSegParser, __SFI_Str("/"), DateSegParser, DateSegParser)).Map(state => {
        const DATE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: DATE,
            extras: ["TYPE", "date"],
            child_nodes: [],
        }
    })

    const CellDataParser = __SFI_Choice(DateParser, __SFI_Int, __SFI_Float, __SFI_Bool, __SFI_Letters)
    const CellParser = new Parser(__SFI_SeqOf(__SFI_Letters, __SFI_Int)).Map(state => {
        const CELL = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL,
            extras: ["TYPE", "cell"],
            child_nodes: [],
        }
    })
    const CellRangeParser = new Parser(__SFI_SeqOf(CellParser, __SFI_Str(":"), CellParser)).Map(state => {
        const CELL_RANGE = state.child_nodes.map(node => node.result.res).join("")
        return {
            res: CELL_RANGE,
            extras: ["TYPE", "cell_range", "FROM", state.child_nodes[0].result.res, "TO", state.child_nodes[2].result.res],
            child_nodes: [],
        }
    })


    const FormulaArg = __SFI_Choice(CellParser, CellDataParser)
    const OperParser = __SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/"))
    const FormulaParam = new Parser(__SFI_ManyOne(__SFI_SeqOf(FormulaArg, WhiteSpaceParser, OperParser))).Map(state => {
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
    const FormulaParser = new Parser(__SFI_SeqOf(__SFI_Str("="), FormulaParam, FormulaArg)).Map(state => {
        let children = [...state.child_nodes[1].result.child_nodes, state.child_nodes[2]]
        __SFI_CreateFormulaTree(children)
        return {
            res: "",
            extras: ["TYPE", "formula"],
            child_nodes: children
        }
    })

    const FuncArgParser = __SFI_Choice(CellRangeParser, CellParser, CellDataParser)
    const FuncFormulaParam = __SFI_ManyOne(__SFI_SeqOf(FuncArgParser, __SFI_Str(","), WhiteSpaceParser))
    const FuncFormulaParamList = __SFI_SeqOf(FuncFormulaParam, FuncArgParser)
    const FuncFormula = __SFI_SeqOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), __SFI_Choice(FuncFormulaParamList, FuncArgParser), __SFI_Str(")"))

    return __SFI_Choice(FuncFormula, FormulaParser)
}

// all the main functions below are just for me to keep recrod of my learning of parser combinators
// as they relate to this project

// this function will not be used in the project, this is just me
// playing around with code and learning more about parser combinators
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

    const DiceRollParser = new Parser(__SFI_SeqOf(NumberParser, __SFI_Str("d"), NumberParser)).Map(state => {
        return {
            res: state.child_nodes.map(node => node.result.res).join(""),
            extras: ["TYPE", "dice_roll", "DICE", state.child_nodes[0].result.res, "SIDES", state.child_nodes[2].result.res],
            child_nodes: state.child_nodes,
        }
    })

    const MainParser = new Parser(__SFI_SeqOf(LetterParser, __SFI_Str(":"))).Map(state => {
        return {
            res: state.child_nodes[0].result.res,
            extras: [],
            child_nodes: [],
        }
    }).Chain(res_state => {
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

// this function will not be used in the project, this is just me
// playing around with code and learning more about parser combinators
function __SFI_ParseFormulaMain2() {
    const Item = __SFI_LazyEval(() => {
        const Num = __SFI_Choice(__SFI_Float, __SFI_Int)
        return __SFI_Choice(BetweenSquareBrackets, Num)
    })
    const OB = __SFI_Str("[")
    const CB = __SFI_Str("]")
    const Sep = __SFI_Str(",")

    const BetweenSquareBrackets = new Parser(__SFI_SeqOf(OB, __SFI_SepBy(Item, Sep), CB)).Map(state => {
        const CHILDREN = state.child_nodes[1].result.child_nodes
        const LEFT = state.child_nodes[0].result.res
        const RIGHT = state.child_nodes[2].result.res
        const NUMBER_CHILDREN = CHILDREN.filter(node => node.result.res !== ',')
        const NUM_TO_STR_LIST = NUMBER_CHILDREN.map(node => node.result.res)
        return {
            res: `${LEFT}${NUM_TO_STR_LIST.join(", ")}${RIGHT}`,
            extras: NUM_TO_STR_LIST,
            child_nodes: [...NUMBER_CHILDREN]
        }
    })


    const MainParser = new Parser(state => BetweenSquareBrackets.Run(state))

    console.log(MainParser.Run("[1,2.73,[3,4,5],6]").toString())
}

// this function will not be used in the project, this is just me
// playing around with code and learning more about parser combinators
function __SFI_ParseFormulaMain3() {
    const Num = __SFI_Choice(__SFI_Float, __SFI_Int)
    const Data = __SFI_Choice(Num, __SFI_Letters)
    const OP = __SFI_Str("(")
    const CP = __SFI_Str(")")
    const OB = __SFI_Str("[")
    const CB = __SFI_Str("]")
    const WS = __SFI_ManyOne(__SFI_Str(" "))
    const OpWs = __SFI_ManyZero(__SFI_Str(" "))
    const Opers = __SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/"))
    const Assign = __SFI_Str("=")

    const ExpVal = __SFI_LazyEval(() => {
        return __SFI_Choice(Exp, AutoArrExp, ArrExp, Num, __SFI_Letters)
    })

    const ArrExp = new Parser(__SFI_SeqOf(OB, Opers, WS, __SFI_SepBy(Data, WS), CB)).Map(state => {
        const OPER = state.child_nodes[1].result.res
        const VALS = state.child_nodes[3].result.child_nodes.filter(node => node.result.res !== ' ' && node.result.res !== '').map(node => node.result.res)
        return {
            res: "[]",
            extras: [OPER, ...VALS],
            child_nodes: []
        }
    })

    const AutoArrExp = new Parser(__SFI_SeqOf(OB, Opers, WS, Data, __SFI_Str(" to "), Data, OpWs, CB)).Map(state => {
        const OPER = state.child_nodes[1].result.res
        const X = state.child_nodes[3]
        const Y = state.child_nodes[5]
        return {
            res: "[...]",
            extras: [OPER],
            child_nodes: [X, Y]
        }
    })

    const Exp = new Parser(__SFI_SeqOf(OP, Opers, WS, __SFI_SepBy(ExpVal, WS), CP)).Map(state => {
        const ALLOWED_OPERS = ['+', '-', '*', '/']
        const OPER = state.child_nodes[1].result.res
        const CHILDREN = state.child_nodes[3].result.child_nodes.filter(node => node.result.res !== ' ' && node.result.res !== '')
        return {
            res: OPER,
            extras: [],
            child_nodes: CHILDREN
        }
    })

    const VarExp = new Parser(__SFI_ManyZero(__SFI_SeqOf(__SFI_Letters, Assign, ExpVal, __SFI_Str(";"), OpWs))).Map(state => {
        let x = new Array<ParserState>()
        for (const CHILD of state.child_nodes) {
            const VAR_NAME = CHILD.result.child_nodes[0]
            const VAR_VAL = CHILD.result.child_nodes[2]
            x.push(VAR_NAME, VAR_VAL)
        }

        if (x.length === 0) { return { res: "~" } }


        return {
            res: "=",
            extras: [],
            child_nodes: x
        }
    })

    const mem = new Map<string, number>()

    const Interpreter = (node: ParserState) => {
        let val = 0
        const OperMap = new Map<string, (a: number, b: number) => number>()
        OperMap.set('+', (a, b) => a + b)
        OperMap.set('-', (a, b) => a - b)
        OperMap.set('*', (a, b) => a * b)
        OperMap.set('/', (a, b) => a / b)
        switch (node.result.res) {
            case '?': {
                Interpreter(node.result.child_nodes[0])
                val = Interpreter(node.result.child_nodes[1])
                break
            }
            case '+': {
                if (node.result.child_nodes.length !== 0 && node.result.child_nodes.length !== 2) { throw new Error("Invalid Expression") }
                val = OperMap.get('+')!(Interpreter(node.result.child_nodes[0]), Interpreter(node.result.child_nodes[1]))
                break
            }
            case '-': {
                if (node.result.child_nodes.length !== 0 && node.result.child_nodes.length !== 2) { throw new Error("Invalid Expression") }
                val = OperMap.get('-')!(Interpreter(node.result.child_nodes[0]), Interpreter(node.result.child_nodes[1]))
                break
            }
            case '*': {
                if (node.result.child_nodes.length !== 0 && node.result.child_nodes.length !== 2) { throw new Error("Invalid Expression") }
                val = OperMap.get('*')!(Interpreter(node.result.child_nodes[0]), Interpreter(node.result.child_nodes[1]))
                break
            }
            case '/': {
                if (node.result.child_nodes.length !== 0 && node.result.child_nodes.length !== 2) { throw new Error("Invalid Expression") }
                let right_val = Interpreter(node.result.child_nodes[1])
                if (right_val === 0) { right_val = 1 }
                val = OperMap.get('/')!(Interpreter(node.result.child_nodes[0]), right_val)
                break
            }
            case '=': {
                for (let i = 0; i < node.result.child_nodes.length; i += 2) {
                    const VAR_NAME = node.result.child_nodes[i].result.res
                    const VAR_VAL = Interpreter(node.result.child_nodes[i + 1])
                    mem.set(VAR_NAME, VAR_VAL)
                }
                break
            }
            case '[]': {
                const OPER = node.result.extras.shift()!

                node.result.extras = node.result.extras.map(val => {
                    if (Number(val) === 0 && OPER === "/") { return "1" }
                    return val
                })
                
                let init_val = Number(node.result.extras.shift())
                while (node.result.extras.length !== 0) {
                    init_val = OperMap.get(OPER)!(init_val, Number(node.result.extras.shift()))
                }
                val = init_val
                break
            }
            case '[...]': {
                const OPER = node.result.extras.shift()!
                const X = Interpreter(node.result.child_nodes[0])
                const Y = Interpreter(node.result.child_nodes[1])

                let start = Math.min(X, Y)
                let end = Math.max(X, Y)
                let init_val = 0
                if (OPER === '*' || OPER === '/') { init_val = 1 }

                for (let i = start; i <= end; i++) {
                    let j = i
                    if (j === 0 && OPER === '/') { j = 1 }
                    init_val = OperMap.get(OPER)!(init_val, j)
                }
                val = init_val
                break
            }
            case '~': { break }
            default: { 
                val = parseFloat(node.result.res)
                if (isNaN(val) && mem.has(node.result.res)) {
                    val = mem.get(node.result.res)!
                }
                else if (isNaN(val)) {
                    val = 0
                }
            }
        }

        return val
    }

    let res = new Parser(__SFI_SeqOf(VarExp, Exp))
        .Map(state => {
            return {
                res: "?",
                extras: [],
                child_nodes: [state.child_nodes[0], state.child_nodes[1]]
            }
        })
        .Run("a=100; b=(+ [/ 10 5 2 0] 0); c=[+ 1032 to 23]; (* (+ a b) c)")

    if (res.is_error) {
        console.log(res.parser_error)
    }
    else {
        console.log(Interpreter(res).toString())
    }
}