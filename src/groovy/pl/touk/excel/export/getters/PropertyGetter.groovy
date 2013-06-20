package pl.touk.excel.export.getters

abstract class PropertyGetter<From, To> implements Getter<To> {
    protected String propertyName

    PropertyGetter(String propertyName) {
        this.propertyName = propertyName
    }

    String getPropertyName() {
        return propertyName
    }

    To getFormattedValue(Object object) {
        if(propertyName == null) {
            return null
        }
        if(propertyName.contains('.')) {
            return getValueFromChildren(object)
        }
        return getFormattedPropertyValue(object)
    }

    private Object getFormattedPropertyValue(object) {
        if (!object.properties.containsKey(propertyName)) {
            return null
        }
        return format(object."$propertyName")
    }

    private Object getValueFromChildren(object) {
        def value = propertyName.tokenize('.').inject(object) { Object currentObject, propertyName ->
            if (!(currentObject instanceof Map) && (currentObject == null || !currentObject.properties.containsKey(propertyName))) {
                return null
            }
            return currentObject."$propertyName"
        }
        format(value)
    }

    def protected abstract format(From value)
}