package pl.touk.excel.export

import org.junit.Test
import pl.touk.excel.export.getters.Getter

import static pl.touk.excel.export.Formatters.asDate

abstract class XlsxExporterHeaderTest extends XlsxExporterTest {
    @Test
    void shouldCreateHeader() {
        //given
        List<Getter> headerProperties = ["First", "Second", "Third"]

        //when
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(headerProperties, 0)
    }

    @Test
    void shouldAcceptDatesInHeader() {
        //given
        List<Getter> headerProperties = ["First", asDate("MyDateAsLong"), "Third"]

        //when
        xlsxReporter.fillHeader(headerProperties)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(headerProperties, 0)
    }
}
