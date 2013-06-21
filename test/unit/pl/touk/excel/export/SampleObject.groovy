
package pl.touk.excel.export

class SampleObject {
    String stringValue = "string"
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
}

class ChildObject {
    String stringValue = "childName"
}
