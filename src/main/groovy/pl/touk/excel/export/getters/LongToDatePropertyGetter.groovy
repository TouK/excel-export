package pl.touk.excel.export.getters

import groovy.transform.InheritConstructors

@InheritConstructors
class LongToDatePropertyGetter extends PropertyGetter<Long, Date> {

    Date format(Long timestamp) {
        return new Date(timestamp)
    }
}

