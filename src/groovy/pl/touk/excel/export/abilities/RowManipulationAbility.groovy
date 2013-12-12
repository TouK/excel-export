package pl.touk.excel.export.abilities

import org.codehaus.groovy.runtime.NullObject
import pl.touk.excel.export.Formatters
import pl.touk.excel.export.getters.Getter

import java.sql.Timestamp
import pl.touk.excel.export.IPoiSheetManipulator

@Category(IPoiSheetManipulator)
class RowManipulationAbility {
    private static final handledPropertyTypes = [String, Getter, Date, Boolean, Timestamp, NullObject, Long, Integer, BigDecimal, BigInteger, Byte, Double, Float, Short]

    IPoiSheetManipulator fillHeader(List properties) {
        fillRow(Formatters.convertSafelyFromGetters(properties), 0)
    }

    IPoiSheetManipulator fillRow(List<Object> properties) {
        fillRow(properties, 1)
    }

    IPoiSheetManipulator fillRow(List<Object> properties, int rowNumber) {
        fillRowWithValues(properties, rowNumber)
    }

    IPoiSheetManipulator fillRowWithValues(List<Object> properties, int rowNumber) {
        properties.eachWithIndex { Object property, int index ->
            def propertyToBeInserted = property == null ? "" : property
            verifyPropertyTypeCanBeHandled(property)
            putCellValue(rowNumber, index, propertyToBeInserted)
        }
        this
    }

    IPoiSheetManipulator add(List<Object> objects, List<Object> selectedProperties) {
        add(objects, selectedProperties, 1)
    }

    IPoiSheetManipulator add(List<Object> objects, List<Object> selectedProperties, int rowNumber) {
        objects.eachWithIndex() { Object object, int index ->
            this.add(object, selectedProperties, rowNumber + index)
        }
        this
    }

    IPoiSheetManipulator add(Object object, List<Object> selectedProperties, int rowNumber) {
        List<Object> properties = getPropertiesFromObject(object, Formatters.convertSafelyToGetters(selectedProperties))
        fillRow(properties, rowNumber)
    }

    private static List<Object> getPropertiesFromObject(Object object, List<Getter> selectedProperties) {
        selectedProperties.collect { it.getFormattedValue(object) }
    }

    private static void verifyPropertyTypeCanBeHandled(Object property) {
        if(!(property.getClass() in handledPropertyTypes)) {
            throw new IllegalArgumentException("Properties should by of types: " + handledPropertyTypes + ". Found " + property.getClass())
        }
    }
}
