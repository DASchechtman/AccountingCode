function GroupAndCollapseBills() {
    __Util_GroupByDate("Due Date", WEEKLY_CREDIT_CHARGES_TAB_NAME);
    BreakDownRepayment()
}