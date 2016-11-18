package pl.touk.excel.export

import static pl.touk.excel.export.Formatters.asDate

class XlsxExporterHeaderSpec extends XlsxExporterSpec {

    void "should create header"() {
        given:
        List headerProperties = ["First", "Second", "Third"]

        when:
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(headerProperties, 0)
    }

    void "should accept dates in header"() {
        given:
        List headerProperties = ["First", asDate("MyDateAsLong"), "Third"]

        when:
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(headerProperties, 0)
    }
}
