package pl.touk.excel.export

class XlsxExporterDateStyleTest extends XlsxExporterTest {

    void shouldSetDefaultDateStyleForDateCells() {
        given:
        String stringValue = "spomething"
        Date dateValue = new Date()

        when:
        xlsxReporter.
                putCellValue(0, 0, stringValue).
                putCellValue(0, 1, dateValue)

        then:
        xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString() == "General"
        xlsxReporter.getCellAt(0, 1).getCellStyle().getDataFormatString() == XlsxExporter.defaultDateFormat
    }

    void shouldSetDateStyleForDateCells() {
        given:
        Date dateValue = new Date()
        String expectedFormat = "yyyy-MM-dd"

        when:
        xlsxReporter.setDateCellFormat(expectedFormat)
        xlsxReporter.putCellValue(0, 0, dateValue)

        then:
        xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString() == expectedFormat
    }
}
