package pl.touk.excel.export

class XlsxExporterDateStyleSpec extends XlsxExporterSpec {

    void "should set default date style for date cells"() {
        given:
        String stringValue = "spomething"
        Date dateValue = new Date()

        when:
        xlsxReporter.
                putCellValue(0, 0, stringValue).
                putCellValue(0, 1, dateValue)

        then:
        xlsxReporter.getCellAt(0, 0).cellStyle.dataFormatString == "General"
        xlsxReporter.getCellAt(0, 1).cellStyle.dataFormatString == XlsxExporter.defaultDateFormat
    }

    void "should set date style for date cells"() {
        given:
        Date dateValue = new Date()
        String expectedFormat = "yyyy-MM-dd"

        when:
        xlsxReporter.setDateCellFormat(expectedFormat)
        xlsxReporter.putCellValue(0, 0, dateValue)

        then:
        xlsxReporter.getCellAt(0, 0).cellStyle.dataFormatString == expectedFormat
    }
}
