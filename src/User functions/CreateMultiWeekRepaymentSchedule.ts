function __GenerateRepaymentSchedule() {
    let generated_repayment_schedule = false
    const TAB_NAME = "Multi Week Loans";
    const TAB = new GoogleSheetTabs(TAB_NAME);

    const NUM_OF_REPAYMENT_COL_NAME = "Number of Repayments";
    const LOANEE_COL_NAME = "Loanee";
    const REPAYMENT_AMT_COL_NAME = "Repayment Amount";
    const PURCHASE_COL_NAME = "Purchase Date";
    const ROUND_UP_COL_NAME = "Round Up?";
    const REPAYMENT_DATE_COL_NAME = "Repayment Date";

    const NUM_OF_PAYMENTS_COL = TAB.GetCol(NUM_OF_REPAYMENT_COL_NAME)
    const LOANEE_COL = TAB.GetCol(LOANEE_COL_NAME)
    const REPAYMENT_COL = TAB.GetCol(REPAYMENT_AMT_COL_NAME)
    const PURCHASE_COL = TAB.GetCol(PURCHASE_COL_NAME)
    const ROUND_UP_COL = TAB.GetCol(ROUND_UP_COL_NAME)
    const REPAYMENT_DATE_COL = TAB.GetCol(REPAYMENT_DATE_COL_NAME)
    let last_row_index = 0

    if (!NUM_OF_PAYMENTS_COL || !LOANEE_COL || !REPAYMENT_COL || !PURCHASE_COL || !ROUND_UP_COL || !REPAYMENT_DATE_COL) { return generated_repayment_schedule }

    const LAST_ROW = NUM_OF_PAYMENTS_COL.find(cell => {
        const cell_num = Number(cell)
        return !isNaN(cell_num) && cell_num > 0
    })

    if (LAST_ROW === undefined) { return generated_repayment_schedule }
    generated_repayment_schedule = true

    last_row_index = NUM_OF_PAYMENTS_COL.indexOf(LAST_ROW)

    const NUM_OF_REPAYMENTS = Number(NUM_OF_PAYMENTS_COL[last_row_index])
    const PURCHASE_DATE = new Date()
    const LOANEE = LOANEE_COL[last_row_index]
    let installment = Number(REPAYMENT_COL[last_row_index]) / NUM_OF_REPAYMENTS
    let payment_days = LOANEE === "Dan" ? 14 : 7
    let payment_start_date = new Date(REPAYMENT_DATE_COL[last_row_index])

    if (ROUND_UP_COL[last_row_index] === "Yes") {
        installment = Math.ceil(installment)
    }

    for (let i = last_row_index; i < last_row_index + NUM_OF_REPAYMENTS; i++) {
        if (i >= NUM_OF_PAYMENTS_COL.length) { NUM_OF_PAYMENTS_COL.push("") }
        if (i >= LOANEE_COL.length) { LOANEE_COL.push("") }
        if (i >= REPAYMENT_COL.length) { REPAYMENT_COL.push("") }
        if (i >= PURCHASE_COL.length) { PURCHASE_COL.push("") }
        if (i >= ROUND_UP_COL.length) { ROUND_UP_COL.push("") }
        if (i >= REPAYMENT_DATE_COL.length) { REPAYMENT_DATE_COL.push("") }

        NUM_OF_PAYMENTS_COL[i] = ""
        LOANEE_COL[i] = LOANEE
        REPAYMENT_COL[i] = installment
        PURCHASE_COL[i] = __CreateDateString(PURCHASE_DATE)
        REPAYMENT_DATE_COL[i] = __CreateDateString(payment_start_date)

        payment_start_date.setDate(payment_start_date.getDate() + payment_days)
    }

    TAB.WriteCol(NUM_OF_REPAYMENT_COL_NAME, NUM_OF_PAYMENTS_COL.map(cell => cell === "Number of Repayments" ? cell : ""))
    TAB.WriteCol(LOANEE_COL_NAME, LOANEE_COL)
    TAB.WriteCol(REPAYMENT_AMT_COL_NAME, REPAYMENT_COL)
    TAB.WriteCol(PURCHASE_COL_NAME, PURCHASE_COL)
    TAB.WriteCol(ROUND_UP_COL_NAME, ROUND_UP_COL)
    TAB.WriteCol(REPAYMENT_DATE_COL_NAME, REPAYMENT_DATE_COL)
    TAB.SaveToTab()

    __AddMultiWeekLoanToRepayment(last_row_index)
    return generated_repayment_schedule
}

function __AddMultiWeekLoanToRepayment(start_row: number) {
    const ONE_WEEK_TAB = new GoogleSheetTabs("One Week Loans");
    const MULTI_WEEK_TAB = new GoogleSheetTabs("Multi Week Loans");

    const MULTI_COL_INDEXES = [
        MULTI_WEEK_TAB.GetHeaderIndex("Repayment Date"),
        MULTI_WEEK_TAB.GetHeaderIndex("Purchase Date"),
        MULTI_WEEK_TAB.GetHeaderIndex("Repayment Amount"),
        MULTI_WEEK_TAB.GetHeaderIndex("Purchase Location"),
        MULTI_WEEK_TAB.GetHeaderIndex("Card")
    ]

    const WEEKLY_COL_INDEXES = [
        ONE_WEEK_TAB.GetHeaderIndex("Due Date"),
        ONE_WEEK_TAB.GetHeaderIndex("Purchase Date"),
        ONE_WEEK_TAB.GetHeaderIndex("Amount"),
        ONE_WEEK_TAB.GetHeaderIndex("Purchase Location"),
        ONE_WEEK_TAB.GetHeaderIndex("Card")
    ]

    if (
        !__CheckAllAreNotInvalidIndex(WEEKLY_COL_INDEXES)
        || !__CheckAllAreNotInvalidIndex(MULTI_COL_INDEXES)
    ) { return }

    const [
        MULTI_TAB_DUE_DATE_COL_INDEX,
        MULTI_TAB_PURCHASE_DATE_COL_INDEX,
        MULTI_TAB_PAYMENT_AMT_COL_INDEX,
        MULTI_TAB_PURCHASE_LOCATION_COL_INDEX,
        MULTI_TAB_CARD_COL_INDEX
    ] = MULTI_COL_INDEXES

    const [
        WEEKLY_TAB_DUE_DATE_COL_INDEX,
        WEEKLY_TAB_PURCHASE_DATE_COL_INDEX,
        WEEKLY_TAB_PAYMENT_AMT_COL_INDEX,
        WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX,
        WEEKLY_TAB_CARD_COL_INDEX
    ] = WEEKLY_COL_INDEXES

    const __GetDateIndexBoundries = function (date: string): [number, number] {
        let i = 0
        const ROW = ONE_WEEK_TAB.FindRow(row => {
            const FOUND = row[WEEKLY_TAB_DUE_DATE_COL_INDEX] === date
            i += Number(!FOUND)
            return FOUND
        })

        if (!ROW) { return [-1, -1] }

        let ret: [number, number] = [i, 0]

        while (true) {
            const ROW = ONE_WEEK_TAB.GetRow(i)
            if (!ROW) { break }
            if (ROW[WEEKLY_TAB_DUE_DATE_COL_INDEX] !== date) { break }
            i++
        }

        ret[1] = i - 1
        return ret
    }

    const __HasMultiWeekRepayment = function (begin: number, end: number, purchase_desc: string) {
        let has_repayment = false
        for (let i = begin; i <= end; i++) {
            const ROW = ONE_WEEK_TAB.GetRow(i)
            if (!ROW) { continue }
            if (ROW[WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX] === purchase_desc) {
                has_repayment = true
                break
            }
        }
        return has_repayment
    }

    let purchase_desc = ""
    let credit_card_name = ""

    for (let i = start_row; i < MULTI_WEEK_TAB.NumberOfRows(); i++) {
        const ROW = MULTI_WEEK_TAB.GetRow(i)
        if (!ROW) { continue }

        const DUE_DATE = String(ROW[MULTI_TAB_DUE_DATE_COL_INDEX])
        if (DUE_DATE === "") { continue }

        if (ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX] !== "") { purchase_desc = String(ROW[MULTI_TAB_PURCHASE_LOCATION_COL_INDEX]) }
        if (ROW[MULTI_TAB_CARD_COL_INDEX] !== "") { credit_card_name = String(ROW[MULTI_TAB_CARD_COL_INDEX]) }

        const [START_INDEX, END_INDEX] = __GetDateIndexBoundries(DUE_DATE)
        if (__HasMultiWeekRepayment(START_INDEX, END_INDEX, purchase_desc)) { continue }

        const NEW_ROW: DataArrayEntry = []
        NEW_ROW[WEEKLY_TAB_DUE_DATE_COL_INDEX] = DUE_DATE
        NEW_ROW[WEEKLY_TAB_PURCHASE_DATE_COL_INDEX] = ROW[MULTI_TAB_PURCHASE_DATE_COL_INDEX]
        NEW_ROW[WEEKLY_TAB_PAYMENT_AMT_COL_INDEX] = ROW[MULTI_TAB_PAYMENT_AMT_COL_INDEX]
        NEW_ROW[WEEKLY_TAB_PURCHASE_LOCATION_COL_INDEX] = purchase_desc
        NEW_ROW[WEEKLY_TAB_CARD_COL_INDEX] = credit_card_name

        if (END_INDEX > -1) {
            const IMMEDIATELY_AFTER_GROUP = END_INDEX + 1
            ONE_WEEK_TAB.InsertRow(IMMEDIATELY_AFTER_GROUP, NEW_ROW)
        } else {
            ONE_WEEK_TAB.AppendRow(NEW_ROW)
        }
    }

    ONE_WEEK_TAB.SaveToTab()
}

function CreateMultiWeekRepaymentSchedule() {
    const GENERATED = __GenerateRepaymentSchedule()
    if (GENERATED) {
        __GroupByDate("Purchase Date", MULTI_WEEK_LOANS_TAB_NAME, false);
        __GroupByDate("Due Date", ONE_WEEK_LOANS_TAB_NAME);
        __ComputeTotal();
    }
}