
package pl.touk.excel.export

import org.junit.Test

class XlsxExporterDateStyleTest extends XlsxExporterTest {
    @Test
    void shouldSetDefaultDateStyleForDateCells() {
        //given
        String stringValue = "spomething"
        Date dateValue = new Date()

        //when
        xlsxReporter.
                putCellValue(0, 0, stringValue).
                putCellValue(0, 1, dateValue)

        //then
        assert xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString() == "General"
        assert xlsxReporter.getCellAt(0, 1).getCellStyle().getDataFormatString() == XlsxExporter.defaultDateFormat
    }

    @Test
    void shouldSetDateStyleForDateCells() {
        //given
        Date dateValue = new Date()
        String expectedFormat = "yyyy-MM-dd"

        //when
        xlsxReporter.setDateCellFormat(expectedFormat)
        xlsxReporter.putCellValue(0, 0, dateValue)

        //then
        assert xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString() == expectedFormat
    }
}
