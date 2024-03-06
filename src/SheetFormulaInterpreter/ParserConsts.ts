interface ParserStateResults {
    res: string
    extras: string[]
    child_nodes: ParserState[]
}

type __SFI_ParserFunc = (state: ParserState) => ParserState
type __SFI_ParserType = (
    'INT' |
    'FLOAT' |
    'STRING' |
    'BOOLEAN' |
    'KEYWORD' |
    'OPERATOR' |
    'NODE' |
    ''
)