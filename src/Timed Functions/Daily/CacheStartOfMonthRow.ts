function CacheStartOfMonthRow() {
    const TODAY = new Date()
    if (TODAY.getDate() === 28) {
        const START_OF_MONTH_ROW = __Util_GetRowThatStartsTheMonth(true)
        const END_OF_MONTH_ROW = __Util_GetRowThatEndsTheMonth(true)

        PropertiesService
            .getDocumentProperties()
            .setProperty(END_OF_MONTH_ROW_CACHE_KEY, String(END_OF_MONTH_ROW))

        PropertiesService
            .getDocumentProperties()
            .setProperty(START_OF_MONTH_ROW_CACHE_KEY, String(START_OF_MONTH_ROW))

        return true
    }
    return false
}

function UpdateStartOfMonthRowCache() {
    const START_OF_MONTH_ROW = __Util_GetRowThatStartsTheMonth(true)
    PropertiesService
        .getDocumentProperties()
        .setProperty(START_OF_MONTH_ROW_CACHE_KEY, String(START_OF_MONTH_ROW))
}

function UpdateEndOfMonthRowCache() {
    const END_OF_MONTH_ROW = __Util_GetRowThatEndsTheMonth(true)
    PropertiesService
        .getDocumentProperties()
        .setProperty(END_OF_MONTH_ROW_CACHE_KEY, String(END_OF_MONTH_ROW))
}