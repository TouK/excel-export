package pl.touk.excel.export

import pl.touk.excel.export.multisheet.AdditionalSheet

import static pl.touk.excel.export.Formatters.asDate

class SampleObject {
    static final List<String> propertyNames = ["stringValue", "dateValue", "longValue", "booleanValue", asDate("dateAsLong"), "notExistingValue", "child.stringValue",
            "bigDecimalValue", "bigIntegerValue", "byteValue", "doubleValue", "floatValue", "integerValue", "shortValue",
            "bytePrimitiveValue", "doublePrimitiveValue", "floatPrimitiveValue", "integerPrimitiveValue", "shortPrimitiveValue", "booleanPrimitiveValue",
            "simpleMap.simpleMapKey1", "simpleMap.simpleMapKey2", "nestedMap.nestedMapKey.childMapKey", "child", "enumValue"]

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
    EnumObject enumValue = EnumObject.SECOND_VALUE

    private void verifyRowHasSelectedProperties(AdditionalSheet additionalSheet, int rowNumber) {
        assert additionalSheet.getCellAt(rowNumber, 0)?.stringCellValue == stringValue
        assert additionalSheet.getCellAt(rowNumber, 1)?.dateCellValue == dateValue
        assert additionalSheet.getCellAt(rowNumber, 2)?.numericCellValue == longValue
        assert additionalSheet.getCellAt(rowNumber, 3)?.booleanCellValue == booleanValue
        assert additionalSheet.getCellAt(rowNumber, 4)?.dateCellValue == new Date(dateAsLong)
        assert additionalSheet.getCellAt(rowNumber, 5)?.stringCellValue == ''
        assert additionalSheet.getCellAt(rowNumber, 6)?.stringCellValue == child.stringValue
        assert additionalSheet.getCellAt(rowNumber, 7)?.numericCellValue == bigDecimalValue
        assert additionalSheet.getCellAt(rowNumber, 8)?.numericCellValue == bigIntegerValue
        assert additionalSheet.getCellAt(rowNumber, 9)?.numericCellValue == byteValue
        assert additionalSheet.getCellAt(rowNumber, 10)?.numericCellValue == doubleValue
        assert additionalSheet.getCellAt(rowNumber, 11)?.numericCellValue == floatValue.toDouble()
        assert additionalSheet.getCellAt(rowNumber, 12)?.numericCellValue == integerValue
        assert additionalSheet.getCellAt(rowNumber, 13)?.numericCellValue == shortValue
        assert additionalSheet.getCellAt(rowNumber, 14)?.numericCellValue == bytePrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 15)?.numericCellValue == doublePrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 16)?.numericCellValue == floatPrimitiveValue.toDouble()
        assert additionalSheet.getCellAt(rowNumber, 17)?.numericCellValue == integerPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 18)?.numericCellValue == shortPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 19)?.booleanCellValue == booleanPrimitiveValue
        assert additionalSheet.getCellAt(rowNumber, 20)?.stringCellValue == simpleMap.simpleMapKey1
        assert additionalSheet.getCellAt(rowNumber, 21)?.stringCellValue == simpleMap."simpleMapKey2"
        assert additionalSheet.getCellAt(rowNumber, 22)?.stringCellValue == nestedMap.nestedMapKey.childMapKey
        assert additionalSheet.getCellAt(rowNumber, 23)?.stringCellValue == child.toString()
        assert additionalSheet.getCellAt(rowNumber, 24)?.stringCellValue == EnumObject.SECOND_VALUE.toString()
    }
}

class ChildObject {
    String stringValue = "childName"

    String toString(){
        'String representation of this object'
    }
}

enum EnumObject {
    FIRST_VALUE, SECOND_VALUE
}
