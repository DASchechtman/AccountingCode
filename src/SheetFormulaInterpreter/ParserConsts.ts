interface ParserStateResults {
    res: string
    extras: string[]
    child_nodes: ParserState[]
    ReconstructState: () => ParserState
}

interface console {
    log: (...params: any[]) => undefined
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
    'REGEX' |
    'END_OF_INPUT' |
    ''
)