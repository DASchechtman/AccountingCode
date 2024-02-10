function CreateMultiWeekRepaymentSchedule() {
    const GENERATED = GenerateRepaymentSchedule()
    if (GENERATED) {
        __GroupByDate("Purchase Date", MULTI_WEEK_LOANS_TAB_NAME, false);
        __GroupByDate("Due Date", ONE_WEEK_LOANS_TAB_NAME);
        __ComputeTotal();
    }
}