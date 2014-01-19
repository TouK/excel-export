
package pl.touk.excel.export

import pl.touk.excel.export.multisheet.AdditionalSheet

import static pl.touk.excel.export.Formatters.asDate

class SampleObject {
    static final List<String> propertyNames = ["stringValue", "dateValue", "longValue", "booleanValue", asDate("dateAsLong"), "notExistingValue", "child.stringValue",
            "bigDecimalValue", "bigIntegerValue", "byteValue", "doubleValue", "floatValue", "integerValue", "shortValue",
            "bytePrimitiveValue", "doublePrimitiveValue", "floatPrimitiveValue", "integerPrimitiveValue", "shortPrimitiveValue", "booleanPrimitiveValue",
            "simpleMap.simpleMapKey1", "simpleMap.simpleMapKey2", "nestedMap.nestedMapKey.childMapKey"]

    String stringValue = UUID.randomUUID().toString()
    Date dateValue = new Date()
    Long longValue = 654L
    Boolean booleanValue = true
    Set setValue =  [1, 2, 3, 4]
    Long dateAsLong = 1234567890L
    ChildObject child = new ChildObject()
    BigDecimal bigDecimalValue = BigDecimal.ONE
    BigInteger bigIntegerValue = BigInteger.TEN
    Byte byteValue = 100
    Double doubleValue = 123.45
    Float floatValue = 123.45f
    Integer integerValue = 99
    Short shortValue = 256
    byte bytePrimitiveValue = 100
    double doublePrimitiveValue = 123.45
    float floatPrimitiveValue = 123.45f
    int integerPrimitiveValue = 99
    short shortPrimitiveValue = 256
    boolean booleanPrimitiveValue = 256
    def simpleMap = [simpleMapKey1: 'simpleMapValue1', simpleMapKey2: 'simpleMapValue2']
    def nestedMap = [nestedMapKey: [childMapKey: 'childMapValue']]

    private void verifyRowHasSelectedProperties(AdditionalSheet additionalSheet, int rowNumber) {
        assert additionalSheet.getCellAt(rowNumber, 0)?.getStringCellValue() == stringValue
        assert additionalSheet.getCellAt(rowNumber, 1)?.getDateCellValue() == dateValue
        assert additionalSheet.getCellAt(rowNumber, 2)?.getNumericCellValue() == longValue
        assert additionalSheet.getCellAt(rowNumber, 3)?.getBooleanCellValue() == booleanValue
        assert additionalSheet.getCellAt(rowNumber, 4)?.getDateCellValue() == new Date(dateAsLong)
        assert additionalSheet.getCellAt(rowNumber, 5)?.getStringCellValue() == ''
        assert additionalSheet.getCellAt(rowNumber, 6)?.getStringCellValue() == child.stringValue
        assert additionalSheet.getCellAt(rowNumber, 7)?.getNumericCellValue() == bigDecimalValue
        assert additionalSheet.getCellAt(rowNumber, 8)?.getNumericCellValue() == bigIntegerValue
        assert additionalSheet.getCellAt(rowNumber, 9)?.getNumericCellValue() == byteValue
        assert additionalSheet.getCellAt(rowNumber, 10)?.getNumericCellValue() == doubleValue
        assert additionalSheet.getCellAt(rowNumber, 11)?.getNumericCellValue() == floatValue.toDouble()
        assert additionalSheet.getCellAt(rowNumber, 12)?.getNumericCellValue() == integerValue
        assert additionalSheet.getCellAt(rowNumber, 13)?.getNumericCellValue() == shortValue
        assert additionalSheet.getCellAt(rowNumber, 14)?.getNumericCellValue() == bytePrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 15)?.getNumericCellValue() == doublePrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 16)?.getNumericCellValue() == floatPrimitiveValue.toDouble()
        assert additionalSheet.getCellAt(rowNumber, 17)?.getNumericCellValue() == integerPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 18)?.getNumericCellValue() == shortPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 19)?.getBooleanCellValue() == booleanPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 20)?.getStringCellValue() == simpleMap.simpleMapKey1
        assert additionalSheet.getCellAt(rowNumber, 21)?.getStringCellValue() == simpleMap."simpleMapKey2"
        assert additionalSheet.getCellAt(rowNumber, 22)?.getStringCellValue() == nestedMap.nestedMapKey.childMapKey
    }
}

class ChildObject {
    String stringValue = "childName"
}
