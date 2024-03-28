function __SFI_MutateParserState(mutator: __SFI_ParserFunc | Parser, state: ParserState) {
    return typeof mutator === 'function' ? mutator(state) : mutator.Run(state)
}

const __SFI_Str = (str: string) => (state: ParserState) => {

    if (state.is_error) { return state }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }
    else if (target_str.startsWith(str)) {
        NEW_STATE.result.res = str
        NEW_STATE.index += str.length
        NEW_STATE.type = "KEYWORD"
    }
    else {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, str.length)}' to start with '${str}' but it did not.`
    }

    return NEW_STATE
}

const __SFI_Letters = (state: ParserState) => {
    if (state.is_error) { return state }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target
    const letters_regex = /^[a-zA-Z]+/
    const letters_match = letters_regex.exec(target_str)

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }


    if (!letters_match) {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to be a letter but it was not.`
        return NEW_STATE
    }

    NEW_STATE.result.res = letters_match[0]
    NEW_STATE.index += letters_match[0].length
    NEW_STATE.type = "STRING"

    return NEW_STATE
}

const __SFI_Int = (state: ParserState) => {

    if (state.is_error) { return state }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target
    const int_regex = /^-?\d+/
    const int_match = int_regex.exec(target_str)

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }


    if (!int_match) {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to be an integer but it was not.`
        return NEW_STATE
    }

    NEW_STATE.result.res = int_match[0]
    NEW_STATE.index += int_match[0].length
    NEW_STATE.type = "INT"

    return NEW_STATE
}

const __SFI_Regex= (regex: RegExp | string) => (state: ParserState) => {
    if (state.is_error) { return state }
    let match_regex: RegExp
    if (typeof regex === 'string') {
        match_regex = new RegExp(`^${regex}`)
    }
    else {
        match_regex = new RegExp(`^${regex.source}`)
    }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target
    const match = match_regex.exec(target_str)

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }

    if (!match) {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to match the regex '${regex}' but it did not.`
        return NEW_STATE
    }

    NEW_STATE.result.res = match[0]
    NEW_STATE.index += match[0].length
    NEW_STATE.type = "REGEX"

    return NEW_STATE
}

const __SFI_Float = (state: ParserState) => {

    if (state.is_error) { return state }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target
    const float_regex = /^(-?\d+\.\d+)|(-?\d+f)/
    const float_match = float_regex.exec(target_str)

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }


    if (!float_match) {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to be a float but it was not.`
        return NEW_STATE
    }

    NEW_STATE.result.res = float_match[0]
    NEW_STATE.index += float_match[0].length
    NEW_STATE.type = "FLOAT"

    return NEW_STATE
}

const __SFI_Bool = (state: ParserState) => {

    if (state.is_error) { return state }

    const NEW_STATE = state.CloneIndexOnly()
    const target_str = NEW_STATE.target.toLowerCase()

    if (target_str.length === 0) {
        NEW_STATE.parser_error = "Unexpected End of Input."
    }
    else if (target_str.startsWith("true")) {
        NEW_STATE.result.res = "true"
        NEW_STATE.index += 4
        NEW_STATE.type = "BOOLEAN"
    }
    else if (target_str.startsWith("false")) {
        NEW_STATE.result.res = "false"
        NEW_STATE.index += 5
        NEW_STATE.type = "BOOLEAN"
    }
    else {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to be a boolean but it was not.`
    }

    return NEW_STATE
}

const __SFI_SequenceOf = (...parsers: (__SFI_ParserFunc | Parser)[]) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    for (let parser of parsers) {
        next_state = __SFI_MutateParserState(parser, next_state)
        if (next_state.is_error) { return next_state }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_ManyZero = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state)
        if (next_state.is_error) { break }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_ManyOne = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    let first_state = __SFI_MutateParserState(parser, next_state)

    if (first_state.is_error) {
        first_state.parser_error = "Expected at least one match for ManyOne but got none."
        return first_state
    }

    NEW_STATE.result.child_nodes.push(first_state)
    NEW_STATE.index = first_state.index

    while (true) {
        next_state = __SFI_MutateParserState(parser, next_state)
        if (next_state.is_error) { break }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_Choice = (...parsers: (__SFI_ParserFunc | Parser)[]) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    for (let parser of parsers) {
        let next_state = __SFI_MutateParserState(parser, NEW_STATE)
        if (!next_state.is_error) {
            return next_state
        }
    }

    NEW_STATE.parser_error = "All parsers failed."
    return NEW_STATE
}

const __SFI_Optional = (parser: __SFI_ParserFunc | Parser) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = __SFI_MutateParserState(parser, NEW_STATE)
    if (next_state.is_error) {
        return NEW_STATE
    }

    return next_state
}