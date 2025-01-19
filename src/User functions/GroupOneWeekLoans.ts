function GroupOneWeekLoans() {
    __Util_GroupByDate("Due Date", WEEKLY_CREDIT_CHARGES_TAB_NAME);
    __Util_ComputeTotal()
}