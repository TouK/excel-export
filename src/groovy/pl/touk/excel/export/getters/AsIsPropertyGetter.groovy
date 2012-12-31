package pl.touk.excel.export.getters

class AsIsPropertyGetter extends PropertyGetter<Object, Object> {
    AsIsPropertyGetter(String propertyName) {
        super(propertyName)
    }

    protected format(Object value) {
        return value
    }
}
