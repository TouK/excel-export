package pl.touk.excel.export.getters

class LongToDatePropertyGetter extends PropertyGetter<Long, Date> {
    LongToDatePropertyGetter(String propertyName) {
        super(propertyName)
    }

    Date format(Long timestamp) {
        return new Date(timestamp)
    }
}

