package pl.touk.excel.export.getters

class AsClosurePropertyGetter extends PropertyGetter<Object, Object> {

	def closure

    AsClosurePropertyGetter(Closure closure) {
        super(null)
        this.closure = closure
    }

    def getFormattedValue(Object object) {
    	closure.call(object)
    }

    def format(Object value) {
        null
    }
}
