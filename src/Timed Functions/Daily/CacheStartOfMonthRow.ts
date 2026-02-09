function CacheStartOfMonthRow() {
    const TODAY = new Date()
    if (TODAY.getDate() === 28) {
        const START_OF_MONTH_ROW = __Util_GetRowThatStartsTheMonth(true)
        PropertiesService
            .getUserProperties()
            .setProperty(START_OF_MONTH_ROW_CACHE_KEY, String(START_OF_MONTH_ROW))
    }
}