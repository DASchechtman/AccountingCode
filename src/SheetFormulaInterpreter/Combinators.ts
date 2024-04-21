function __SFI_MutateParserState(mutator: __SFI_ParserFunc | Parser, state: ParserState, orig_state?: ParserState) {
    const NEW_STATE = state.CloneIndexOnly()
    const NEXT = typeof mutator === 'function' ? mutator(NEW_STATE) : mutator.Run(NEW_STATE)
    if (orig_state && !NEXT.is_error) {
        orig_state.index = NEXT.index
    }
    return NEXT
}

const __SFI_Str = (str: string) => (state: ParserState) => {

    if (state.is_error) { return state }

    const TARGET_STR = state.target

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }
    else if (TARGET_STR.startsWith(str)) {
        state.result.res = str
        state.index += str.length
        state.type = "KEYWORD"
    }
    else {
        state.parser_error = `Expected '${TARGET_STR.slice(0, str.length)}' to start with '${str}' but it did not.`
    }

    return state
}

const __SFI_Letters = (state: ParserState) => {
    if (state.is_error) { return state }

    const TARGET_STR = state.target
    const LETTERS_REGEX = /^[a-zA-Z]+/
    const LETTERS_MATCH = LETTERS_REGEX.exec(TARGET_STR)

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }


    if (!LETTERS_MATCH) {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to be a letter but it was not.`
        return state
    }

    state.result.res = LETTERS_MATCH[0]
    state.index += LETTERS_MATCH[0].length
    state.type = "STRING"

    return state
}

const __SFI_Int = (state: ParserState) => {

    if (state.is_error) { return state }

    const TARGET_STR = state.target
    const INT_REGEX = /^-?\d+/
    const INT_MATCH = INT_REGEX.exec(TARGET_STR)

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }


    if (!INT_MATCH) {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to be an integer but it was not.`
        return state
    }

    state.result.res = INT_MATCH[0]
    state.index += INT_MATCH[0].length
    state.type = "INT"

    return state
}

const __SFI_Digits = (state: ParserState) => {
    if (state.is_error) { return state }

    const TARGET_STR = state.target
    const DIGITS_REGEX = /^\d+/
    const DIGITS_MATCH = DIGITS_REGEX.exec(TARGET_STR)

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }

    if (!DIGITS_MATCH) {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to be a digit but it was not.`
        return state
    }

    state.result.res = DIGITS_MATCH[0]
    state.index += DIGITS_MATCH[0].length
    state.type = "INT"

    return state
}

const __SFI_Regex = (regex: RegExp | string) => (state: ParserState) => {
    if (state.is_error) { return state }

    let match_regex: RegExp
    if (typeof regex === 'string') {
        match_regex = new RegExp(`^${regex}`)
    }
    else {
        match_regex = new RegExp(`^${regex.source}`)
    }

    const TARGET_STR = state.target
    const MATCH = match_regex.exec(TARGET_STR)

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }

    if (!MATCH) {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to match the regex '${regex}' but it did not.`
        return state
    }

    state.result.res = MATCH[0]
    state.index += MATCH[0].length
    state.type = "REGEX"

    return state
}

const __SFI_Float = (state: ParserState) => {

    if (state.is_error) { return state }

    const TARGET_STR = state.target
    const FLOAT_REGEX = /^(-?\d+\.\d+)|(-?\d+f)/
    const FLOAT_MATCH = FLOAT_REGEX.exec(TARGET_STR)

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }


    if (!FLOAT_MATCH) {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to be a float but it was not.`
        return state
    }

    state.result.res = FLOAT_MATCH[0]
    state.index += FLOAT_MATCH[0].length
    state.type = "FLOAT"

    return state
}

const __SFI_Bool = (state: ParserState) => {

    if (state.is_error) { return state }

    const TARGET_STR = state.target.toLowerCase()

    if (TARGET_STR.length === 0) {
        state.parser_error = "Unexpected End of Input."
    }
    else if (TARGET_STR.startsWith("true")) {
        state.result.res = "true"
        state.index += 4
        state.type = "BOOLEAN"
    }
    else if (TARGET_STR.startsWith("false")) {
        state.result.res = "false"
        state.index += 5
        state.type = "BOOLEAN"
    }
    else {
        state.parser_error = `Expected '${TARGET_STR.slice(0, 1)}' at ${state.index} to be a boolean but it was not.`
    }

    return state
}

const __SFI_SeqOf = (...parsers: (__SFI_ParserFunc | Parser)[]) => (state: ParserState) => {

    if (state.is_error) { return state }

    let next_state = state
    for (let parser of parsers) {
        next_state = __SFI_MutateParserState(parser, next_state, state)
        if (next_state.is_error) { return next_state }
        state.result.child_nodes.push(next_state)
    }

    state.type = "NODE"
    return state
}

const __SFI_ManyZero = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }

    let next_state = state
    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)
    }

    state.type = "NODE"
    return state
}

const __SFI_ManyOne = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }

    let next_state = state
    let first_state = __SFI_MutateParserState(parser, next_state, state)

    if (first_state.is_error) {
        first_state.parser_error = "Expected at least one match for ManyOne but got none."
        return first_state
    }

    state.result.child_nodes.push(first_state)

    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)
    }

    state.type = "NODE"
    return state
}

const __SFI_Choice = (...parsers: (__SFI_ParserFunc | Parser)[]) => (state: ParserState) => {

    if (state.is_error) { return state }

    for (let parser of parsers) {
        let next_state = __SFI_MutateParserState(parser, state)
        if (!next_state.is_error) {
            return next_state
        }
    }

    state.parser_error = "All parsers failed."
    return state
}

const __SFI_Optional = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }

    let next_state = __SFI_MutateParserState(parser, state)
    if (next_state.is_error) {
        return state
    }

    return next_state
}

const __SFI_EndOfInput = (state: ParserState) => {
    if (state.is_error) { return state }

    if (state.target.length === 0) {
        state.type = "END_OF_INPUT"
    }
    else {
        state.parser_error = "Expected End of Input."
    }

    return state
}

const __SFI_LazyEval = (parser: () => __SFI_ParserFunc | Parser) => (state: ParserState) => {
    return __SFI_MutateParserState(parser(), state)
}

const __SFI_SepBy = (parser: __SFI_ParserFunc | Parser, separator: __SFI_ParserFunc | Parser) => (state: ParserState) => {
    if (state.is_error) { return state }
    let next_state = state

    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)

        next_state = __SFI_MutateParserState(separator, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)
    }

    state.type = "NODE"
    return state
}

const __SFI_SepByOne = (parser: __SFI_ParserFunc | Parser, separator: __SFI_ParserFunc | Parser) => (state: ParserState) => {
    if (state.is_error) { return state }
    let next_state = state

    next_state = __SFI_MutateParserState(parser, next_state, state)
    if (next_state.is_error) {
        next_state.parser_error = "Expected at least one value to be  seperated."
        return next_state
    }
    state.result.child_nodes.push(next_state)

    next_state = __SFI_MutateParserState(separator, next_state, state)
    if (next_state.is_error) {
        next_state.parser_error = "Expected at least one separator after the first value."
        return next_state
    }
    state.result.child_nodes.push(next_state)

    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)

        next_state = __SFI_MutateParserState(separator, next_state, state)
        if (next_state.is_error) { break }
        state.result.child_nodes.push(next_state)
    }

    state.type = "NODE"
    return state
}