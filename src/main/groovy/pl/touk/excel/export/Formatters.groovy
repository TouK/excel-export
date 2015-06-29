package pl.touk.excel.export

import pl.touk.excel.export.getters.Getter
import pl.touk.excel.export.getters.PropertyGetter
import pl.touk.excel.export.getters.AsIsPropertyGetter
import pl.touk.excel.export.getters.LongToDatePropertyGetter

class Formatters {
    static PropertyGetter asDate(String propertyName) {
        return new LongToDatePropertyGetter(propertyName)
    }

    static PropertyGetter asIs(String propertyName) {
        return new AsIsPropertyGetter(propertyName)
    }

    static List<Getter> convertSafelyToGetters(List properties) {
        properties.collect {
            if (it instanceof Getter) {
                return it
            } else if(it instanceof String) {
                return asIs(it)
            } else {
                throw IllegalArgumentException('List of properties, which should be either String, a Getter. Found: ' +
                        it?.toString() + ' of class ' + it?.getClass())
            }
        }
    }

    static List<Object> convertSafelyFromGetters(List properties) {
        properties.collect {
            (it instanceof Getter) ? it.getPropertyName() : it
        }
    }
}
