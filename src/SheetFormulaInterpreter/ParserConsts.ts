interface ParserStateResults {
    res: string
    extras: string[]
    child_nodes: ParserState[]
    type: __SFI_ParserType
    ReconstructState: () => ParserState
}

interface console {
    log: (...params: any[]) => undefined
}

type __SFI_ParserFunc = (state: ParserState) => ParserState
type __SFI_ParserType = (
    'INT' 
    | 'FLOAT'
    | 'NUMBER'
    | 'STRING' 
    | 'BOOLEAN'
    | 'DATE'
    | 'PARENS'
    | 'NULL'
    | 'DIGITS' 
    | 'OPERATOR'
    | 'OP_ADD'
    | 'OP_SUB'
    | 'OP_MUL'
    | 'OP_DIV'
    | 'OP_POW'
    | 'OP_PAREN'
    | 'FUNCTION'
    | 'FUNC_MUL'
    | 'FUNC_DIV'
    | 'FUNC_ADD'
    | 'FUNC_SUB'
    | 'FUNC_POW'
    | 'FUNC_SUM'
    | 'NODE' 
    | 'REGEX' 
    | 'KEYWORD'
    | 'END_OF_INPUT' 
    | 'SPREADSHEET_CELL'
    | 'SPREADSHEET_RANGE'
    | 'SPREADSHEET_FUNCTION_FORMULA'
    | 'SPREADSHEET_MATH_FORMULA'
)