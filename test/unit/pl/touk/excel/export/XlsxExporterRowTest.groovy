package pl.touk.excel.export

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Test

import static pl.touk.excel.export.Formatters.asDate

class XlsxExporterRowTest extends XlsxExporterTest {
    List<String> allPropertiesToBeAdded = ["stringValue", "dateValue", "longValue", "booleanValue", asDate("dateAsLong"), "notExistingValue", "child.stringValue",
            "bigDecimalValue", "bigIntegerValue", "byteValue", "doubleValue", "floatValue", "integerValue", "shortValue",
            "bytePrimitiveValue", "doublePrimitiveValue", "floatPrimitiveValue", "integerPrimitiveValue", "shortPrimitiveValue", "booleanPrimitiveValue",
            "simpleMap.simpleMapKey1", "simpleMap.simpleMapKey2", "nestedMap.nestedMapKey.childMapKey"]

    @Test
    void shouldFillRowAtFirstPosition() throws IOException {
        //given
        List rowValues = ["First", "Second", "Third", " "]

        //when
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(rowValues, 1)
    }

    @Test
    void shouldFillRowAtPosition() throws IOException {
        //given
        List rowValues = ["First", "Second", "Third", " "]
        int rowNumber = 15

        //when
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(rowValues, rowNumber)
    }

    @Test
    void shouldFillRowAtPositionWithZero() throws IOException {
        //given
        List rowValues = [0]
        int rowNumber = 25

        //when
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        //then
        assert getCell(new XSSFWorkbook(getFilePath()), rowNumber, 0).getNumericCellValue() == 0
    }

    @Test
    void shouldHandleSubclassesOfValidTypesInProperties() {
        //given
        java.util.Date handledType = new java.util.Date(123123)
        java.sql.Date subclassOfHandledType = new java.sql.Date(321321)
        List rowValues = [handledType, subclassOfHandledType ]

        //when
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        //then
        verifyDateAt(handledType, 1, 0)
        verifyDateAt(subclassOfHandledType, 1, 1)
    }

    @Test
    void shouldFillRowFromPropertyList() {
        //given
        SampleObject testObject = new SampleObject()

        //when
        xlsxReporter.add(testObject, allPropertiesToBeAdded, 3)

        //then
        verifyRowHasSelectedProperties(3, testObject)
    }

    @Test
    void shouldFillRows() {
        //given
        List objects = [new SampleObject(), new SampleObject(), new SampleObject()]

        //when
        xlsxReporter.add(objects, allPropertiesToBeAdded, 0)

        //then
        verifyRowHasSelectedProperties(0, objects.get(0))
        verifyRowHasSelectedProperties(1, objects.get(1))
        verifyRowHasSelectedProperties(2, objects.get(2))
    }

    @Test
    void shouldFillRowsFromListsOfMaps() {
        //given

        // Make sure we can handle nulls
        List stones = [
            [
                first: 'Keith',
                middle: null,
                last: 'Richards'
            ],
            [
                first: 'Mick',
                middle: null,
                last: 'Jaggar'
            ],
            [
                first: 'Ronnie',
                middle: null,
                last: 'Wood'
            ],
            [
                first: 'Charlie',
                middle: null,
                last: 'Watts'
            ],
        ]

        //when
        xlsxReporter.add(stones, ['first','middle'], 0)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == 'Keith'
        assert xlsxReporter.getCellAt(0, 1).getStringCellValue() == ''
        assert xlsxReporter.getCellAt(0, 2) == null
        assert xlsxReporter.getCellAt(3, 0).getStringCellValue() == 'Charlie'
        assert xlsxReporter.getCellAt(3, 1).getStringCellValue() == ''
        assert xlsxReporter.getCellAt(3, 2) == null
    }


    @Test(expected=IllegalArgumentException.class)
    void shouldThrowExceptionWhenListIsPassedAsArgument() {
        //given
        List<ArrayList> objects = [new SampleObjectWithList(), new SampleObjectWithList()]

        //when
        xlsxReporter.add(objects, ['list'])
    }

    private void verifyRowHasSelectedProperties(int rowNumber, SampleObject sampleObject) {
        assert xlsxReporter.getCellAt(rowNumber, 0)?.getStringCellValue() == sampleObject.stringValue
        assert xlsxReporter.getCellAt(rowNumber, 1)?.getDateCellValue() == sampleObject.dateValue
        assert xlsxReporter.getCellAt(rowNumber, 2)?.getNumericCellValue() == sampleObject.longValue
        assert xlsxReporter.getCellAt(rowNumber, 3)?.getBooleanCellValue() == sampleObject.booleanValue
        assert xlsxReporter.getCellAt(rowNumber, 4)?.getDateCellValue() == new Date(sampleObject.dateAsLong)
        assert xlsxReporter.getCellAt(rowNumber, 5)?.getStringCellValue() == ''
        assert xlsxReporter.getCellAt(rowNumber, 6)?.getStringCellValue() == sampleObject.child.stringValue
        assert xlsxReporter.getCellAt(rowNumber, 7)?.getNumericCellValue() == sampleObject.bigDecimalValue
        assert xlsxReporter.getCellAt(rowNumber, 8)?.getNumericCellValue() == sampleObject.bigIntegerValue
        assert xlsxReporter.getCellAt(rowNumber, 9)?.getNumericCellValue() == sampleObject.byteValue
        assert xlsxReporter.getCellAt(rowNumber, 10)?.getNumericCellValue() == sampleObject.doubleValue
        assert xlsxReporter.getCellAt(rowNumber, 11)?.getNumericCellValue() == sampleObject.floatValue.toDouble()
        assert xlsxReporter.getCellAt(rowNumber, 12)?.getNumericCellValue() == sampleObject.integerValue
        assert xlsxReporter.getCellAt(rowNumber, 13)?.getNumericCellValue() == sampleObject.shortValue
        assert xlsxReporter.getCellAt(rowNumber, 14)?.getNumericCellValue() == sampleObject.bytePrimitiveValue
        assert xlsxReporter.getCellAt(rowNumber, 15)?.getNumericCellValue() == sampleObject.doublePrimitiveValue
        assert xlsxReporter.getCellAt(rowNumber, 16)?.getNumericCellValue() == sampleObject.floatPrimitiveValue.toDouble()
        assert xlsxReporter.getCellAt(rowNumber, 17)?.getNumericCellValue() == sampleObject.integerPrimitiveValue
        assert xlsxReporter.getCellAt(rowNumber, 18)?.getNumericCellValue() == sampleObject.shortPrimitiveValue
        assert xlsxReporter.getCellAt(rowNumber, 19)?.getBooleanCellValue() == sampleObject.booleanPrimitiveValue
        assert xlsxReporter.getCellAt(rowNumber, 20)?.getStringCellValue() == sampleObject.simpleMap.simpleMapKey1
        assert xlsxReporter.getCellAt(rowNumber, 21)?.getStringCellValue() == sampleObject.simpleMap."simpleMapKey2"
        assert xlsxReporter.getCellAt(rowNumber, 22)?.getStringCellValue() == sampleObject.nestedMap.nestedMapKey.childMapKey
    }
}
