package pl.touk.excel.export

import static pl.touk.excel.export.Formatters.asDate

class XlsxExporterHeaderTest extends XlsxExporterTest {

    void shouldCreateHeader() {
        given:
        List headerProperties = ["First", "Second", "Third"]

        when:
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(headerProperties, 0)
    }

    void shouldAcceptDatesInHeader() {
        given:
        List headerProperties = ["First", asDate("MyDateAsLong"), "Third"]

        when:
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(headerProperties, 0)
    }
}
