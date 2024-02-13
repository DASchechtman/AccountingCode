function GroupOneWeekLoans() {
    __Util_GroupByDate("Due Date", ONE_WEEK_LOANS_TAB_NAME);
    __Util_ComputeTotal()
}