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
        NEW_STATE.type = "STR-LITERAL"
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
    NEW_STATE.type = "STR"

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
        NEW_STATE.type = "BOOL"
    }
    else if (target_str.startsWith("false")) {
        NEW_STATE.result.res = "false"
        NEW_STATE.index += 5
        NEW_STATE.type = "BOOL"
    }
    else {
        NEW_STATE.parser_error = `Expected '${target_str.slice(0, 1)}' at ${state.index} to be a boolean but it was not.`
    }

    return NEW_STATE
}

const __SFI_SequenceOf = (...parsers: __SFI_ParserFunc[]) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    for (let parser of parsers) {
        next_state = parser(next_state)
        if (next_state.is_error) { return next_state }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_ManyZero = (parser: __SFI_ParserFunc) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    while (true) {
        next_state = parser(next_state)
        if (next_state.is_error) { break }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_ManyOne = (parser: __SFI_ParserFunc) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = NEW_STATE
    let first_state = parser(next_state)

    if (first_state.is_error) {
        first_state.parser_error = "Expected at least one match for ManyOne but got none."
        return first_state
    }

    NEW_STATE.result.child_nodes.push(first_state)
    NEW_STATE.index = first_state.index

    while (true) {
        next_state = parser(next_state)
        if (next_state.is_error) { break }
        NEW_STATE.result.child_nodes.push(next_state)
        NEW_STATE.index = next_state.index
    }

    NEW_STATE.type = "NODE"
    return NEW_STATE
}

const __SFI_Choice = (...parsers: __SFI_ParserFunc[]) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    for (let parser of parsers) {
        let next_state = parser(NEW_STATE)
        if (!next_state.is_error) {
            return next_state
        }
    }

    NEW_STATE.parser_error = "All parsers failed."
    return NEW_STATE
}

const __SFI_Optional = (parser: __SFI_ParserFunc) => (state: ParserState) => {

    if (state.is_error) { return state }
    const NEW_STATE = state.CloneIndexOnly()

    let next_state = parser(NEW_STATE)
    if (next_state.is_error) {
        return NEW_STATE
    }

    return next_state
}


const data_type = __SFI_Choice(__SFI_Bool, __SFI_Float, __SFI_Int, __SFI_Letters)
const cell = __SFI_SequenceOf(__SFI_Letters, __SFI_Int)
const cell_range = __SFI_SequenceOf(cell, __SFI_Str(":"), cell)
const all_data = __SFI_Choice(cell_range, cell, data_type)

const data_list_element = __SFI_SequenceOf(all_data, __SFI_Str(","), __SFI_ManyZero(__SFI_Str(" ")))

const func_params = __SFI_SequenceOf(__SFI_ManyOne(data_list_element), all_data)
const func_call = __SFI_SequenceOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), func_params, __SFI_Str(")"))

function ComboMain() {
    let x = new Parser(func_call).Run("=Multiply(A1, 2, B77:D18)")
    const TAB = new GoogleSheetTabs("Settings")
    let row = TAB.GetRow(0)!
    row[1] = x.toString()
    TAB.WriteRow(0, row)
    TAB.SaveToTab()
}