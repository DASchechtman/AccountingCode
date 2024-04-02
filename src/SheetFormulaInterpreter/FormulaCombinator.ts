const __SFI_DateSegParser = __SFI_Regex(/([0-9]{2,2})/)
const __SFI_WhiteSpaceParser = __SFI_ManyZero(__SFI_Str(" "))

const __SFI_DateParser = __SFI_SequenceOf(__SFI_DateSegParser, __SFI_Str("/"), __SFI_DateSegParser, __SFI_Str("/"), __SFI_DateSegParser, __SFI_DateSegParser)

const __SFI_CellDataParser = __SFI_Choice(__SFI_DateParser, __SFI_Int, __SFI_Float, __SFI_Bool, __SFI_Letters)
const __SFI_CellParser = __SFI_SequenceOf(__SFI_Letters, __SFI_Int)
const __SFI_CellRangeParser = __SFI_SequenceOf(__SFI_CellParser, __SFI_Str(":"), __SFI_CellParser)


const __SFI_FormulaArg = __SFI_Choice(__SFI_CellParser, __SFI_CellDataParser)
const __SFI_OperParser = __SFI_Choice(__SFI_Str("+"), __SFI_Str("-"), __SFI_Str("*"), __SFI_Str("/"))
const __SFI_FormulaParam = __SFI_ManyOne(__SFI_SequenceOf(__SFI_FormulaArg, __SFI_WhiteSpaceParser, __SFI_OperParser))
const __SFI_FormulaParser = __SFI_SequenceOf(__SFI_Str("="), __SFI_FormulaParam, __SFI_FormulaArg)

const __SFI_FuncArgParser = __SFI_Choice(__SFI_CellRangeParser, __SFI_CellParser, __SFI_CellDataParser)
const __SFI_FuncFormulaParam = __SFI_ManyOne(__SFI_SequenceOf(__SFI_FuncArgParser, __SFI_Str(","), __SFI_WhiteSpaceParser))
const __SFI_FuncFormulaParamList = __SFI_SequenceOf(__SFI_FuncFormulaParam, __SFI_FuncArgParser)
const __SFI_FuncFormula = __SFI_SequenceOf(__SFI_Str("="), __SFI_Letters, __SFI_Str("("), __SFI_Choice(__SFI_FuncFormulaParamList, __SFI_FuncArgParser), __SFI_Str(")"))

const __SFI_CellFormulaParser = __SFI_Choice(__SFI_FuncFormula, __SFI_FormulaParser)