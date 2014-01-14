package pl.touk.excel.export

import org.junit.Test
import pl.touk.excel.export.multisheet.SheetManipulator

import static pl.touk.excel.export.Formatters.asDate

class XlsxExporterMultipleSheetTest extends XlsxExporterTest {
    List<String> allPropertiesToBeAdded = ["stringValue", "dateValue", "longValue", "booleanValue", asDate("dateAsLong"), "notExistingValue", "child.stringValue",
            "bigDecimalValue", "bigIntegerValue", "byteValue", "doubleValue", "floatValue", "integerValue", "shortValue",
            "bytePrimitiveValue", "doublePrimitiveValue", "floatPrimitiveValue", "integerPrimitiveValue", "shortPrimitiveValue", "booleanPrimitiveValue"]

    @Test
    void shouldFillRowsInSeparateSheets() {
        //given
        List sheet1Objects = [new SampleObject(), new SampleObject(), new SampleObject()]

        // make sheet second sheet's objects differ
        List sheet2Objects = [new SampleObject(), new SampleObject(), new SampleObject()].collect{ it.stringValue = UUID.randomUUID(); it }

        //when
        xlsxReporter.add(sheet1Objects, allPropertiesToBeAdded, 0)
        xlsxReporter.withSheet('sheet2').add( sheet2Objects, allPropertiesToBeAdded, 0)

        //then
        sheet1Objects.eachWithIndex{ object, index ->
            verifyRowHasSelectedProperties( xlsxReporter, index, object )
        }
        sheet2Objects.eachWithIndex{ object, index ->
            verifyRowHasSelectedProperties( xlsxReporter.withSheet('sheet2'), index, object )
        }
    }


    private void verifyRowHasSelectedProperties(SheetManipulator sheet,  int rowNumber, SampleObject sampleObject) {
        assert sheet.getCellAt(rowNumber, 0)?.getStringCellValue() == sampleObject.stringValue
        assert sheet.getCellAt(rowNumber, 1)?.getDateCellValue() == sampleObject.dateValue
        assert sheet.getCellAt(rowNumber, 2)?.getNumericCellValue() == sampleObject.longValue
        assert sheet.getCellAt(rowNumber, 3)?.getBooleanCellValue() == sampleObject.booleanValue
        assert sheet.getCellAt(rowNumber, 4)?.getDateCellValue() == new Date(sampleObject.dateAsLong)
        assert sheet.getCellAt(rowNumber, 5)?.getStringCellValue() == ''
        assert sheet.getCellAt(rowNumber, 6)?.getStringCellValue() == sampleObject.child.stringValue
        assert sheet.getCellAt(rowNumber, 7)?.getNumericCellValue() == sampleObject.bigDecimalValue
        assert sheet.getCellAt(rowNumber, 8)?.getNumericCellValue() == sampleObject.bigIntegerValue
        assert sheet.getCellAt(rowNumber, 9)?.getNumericCellValue() == sampleObject.byteValue
        assert sheet.getCellAt(rowNumber, 10)?.getNumericCellValue() == sampleObject.doubleValue
        assert sheet.getCellAt(rowNumber, 11)?.getNumericCellValue() == sampleObject.floatValue.toDouble()
        assert sheet.getCellAt(rowNumber, 12)?.getNumericCellValue() == sampleObject.integerValue
        assert sheet.getCellAt(rowNumber, 13)?.getNumericCellValue() == sampleObject.shortValue
        assert sheet.getCellAt(rowNumber, 14)?.getNumericCellValue() == sampleObject.bytePrimitiveValue
        assert sheet.getCellAt(rowNumber, 15)?.getNumericCellValue() == sampleObject.doublePrimitiveValue
        assert sheet.getCellAt(rowNumber, 16)?.getNumericCellValue() == sampleObject.floatPrimitiveValue.toDouble()
        assert sheet.getCellAt(rowNumber, 17)?.getNumericCellValue() == sampleObject.integerPrimitiveValue
        assert sheet.getCellAt(rowNumber, 18)?.getNumericCellValue() == sampleObject.shortPrimitiveValue
        assert sheet.getCellAt(rowNumber, 19)?.getBooleanCellValue() == sampleObject.booleanPrimitiveValue
    }
}
