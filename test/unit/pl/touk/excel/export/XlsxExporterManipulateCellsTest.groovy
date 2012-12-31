
package pl.touk.excel.export

import org.junit.Test

class XlsxExporterManipulateCellsTest extends XlsxExporterTest {
    @Test
    void shouldPutAndReadCellValue() {
        //given
        String stringValue = "spomething"
        Date dateValue = new Date()
        Long longValue = 1L
        Boolean booleanValue = true

        //when
        xlsxReporter.putCellValue(0, 0, stringValue).
                     putCellValue(0, 1, dateValue).
                     putCellValue(0, 2, longValue).
                     putCellValue(0, 3, booleanValue)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == stringValue
        assert xlsxReporter.getCellAt(0, 1).getDateCellValue() == dateValue
        assert xlsxReporter.getCellAt(0, 2).getNumericCellValue() == longValue
        assert xlsxReporter.getCellAt(0, 3).getBooleanCellValue() == booleanValue
    }
}
