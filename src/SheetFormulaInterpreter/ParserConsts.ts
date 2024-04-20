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
    | 'KEYWORD' 
    | 'OPERATOR' 
    | 'NODE' 
    | 'REGEX' 
    | 'END_OF_INPUT' 
    | 'NULL'
    | 'SPREADSHEET_CELL'
    | 'SPREADSHEET_RANGE'
    | 'SPREADSHEET_FUNCTION_FORMULA'
    | 'SPREADSHEET_MATH_FORMULA'
    | 'PARENS'
)