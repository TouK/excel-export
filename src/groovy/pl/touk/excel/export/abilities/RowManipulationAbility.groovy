package pl.touk.excel.export.abilities

import org.codehaus.groovy.runtime.NullObject
import pl.touk.excel.export.Formatters
import pl.touk.excel.export.getters.Getter

import java.sql.Timestamp
import pl.touk.excel.export.multisheet.SheetManipulator

@Category(SheetManipulator)
class RowManipulationAbility {
    private static final handledPropertyTypes = [String, Getter, Date, Boolean, Timestamp, NullObject, Long, Integer, BigDecimal, BigInteger, Byte, Double, Float, Short]

    SheetManipulator fillHeader(List properties) {
        fillRow(Formatters.convertSafelyFromGetters(properties), 0)
    }

    SheetManipulator fillRow(List<Object> properties) {
        fillRow(properties, 1)
    }

    SheetManipulator fillRow(List<Object> properties, int rowNumber) {
        fillRowWithValues(properties, rowNumber)
    }

    SheetManipulator fillRowWithValues(List<Object> properties, int rowNumber) {
        properties.eachWithIndex { Object property, int index ->
            def propertyToBeInserted = RowManipulationAbility.getPropertyToBeInserted(property)
            CellManipulationAbility.putCellValue(this, rowNumber, index, propertyToBeInserted)
        }
        return this
    }

    SheetManipulator add(List<Object> objects, List<Object> selectedProperties, boolean evaluateAsExpressions = false) {
        add(objects, selectedProperties, 1, evaluateAsExpressions)
    }

    SheetManipulator add(List<Object> objects, List<Object> selectedProperties, int rowNumber, boolean evaluateAsExpressions = false) {
        objects.eachWithIndex() { Object object, int index ->
            RowManipulationAbility.add(this, object, selectedProperties, rowNumber + index, evaluateAsExpressions)
        }
        return this
    }

    SheetManipulator add(Object object, List<Object> selectedProperties, int rowNumber, boolean evaluateAsExpressions = false) {
        List<Object> properties = null
        if(evaluateAsExpressions){
            properties = RowManipulationAbility.evaluatePropertiesFromObject(object, selectedProperties)
        } else {
            properties = RowManipulationAbility.getPropertiesFromObject(object, Formatters.convertSafelyToGetters(selectedProperties))
        }
        fillRow(properties, rowNumber)
    }

    private static Object getPropertyToBeInserted(Object property){
        property = property == null ? "" : property
        if(!RowManipulationAbility.verifyPropertyTypeCanBeHandled(property)){
            property = property.toString()
        }
        return property
    }

    private static List<Object> getPropertiesFromObject(Object object, List<Getter> selectedProperties) {
        selectedProperties.collect { it.getFormattedValue(object) }
    }

    private static List<Object> evaluatePropertiesFromObject(Object object, List selectedProperties) {
        selectedProperties.collect { expression -> 
            Eval.me('it', object, expression)
        }
    }

    private static boolean verifyPropertyTypeCanBeHandled(Object property) {
        if(!(handledPropertyTypes.find {it.isAssignableFrom(property.getClass())} )) {
            return false
        } else {
            return true
        }
    }
}
