package pl.touk.excel.export

import org.apache.commons.io.output.NullOutputStream
import spock.lang.Ignore

class WebXlsxExporterTest extends XlsxTestOnTemporaryFolder {
    String originalValue = "1"
    int valueRow = 0
    int valueColumn = 0

    /*This test lies. It passes no matter if the real file is overwritten or not. For now, I am not sure why
            but it may be for caches. I leave this test here to make sure we do not fall into writing it again and
            feeling safe*/
    @Ignore
    void shouldNotOverwriteTemplate() {
        given:
        new XlsxExporter(getFilePath()).with {
            putCellValue(valueRow, valueColumn, originalValue)
            save()
        }

        when:
        new WebXlsxExporter(getFilePath()).with {
            putCellValue(0, 0, "2")
            save(new NullOutputStream())
        }

        then:
        getCellValue(new XlsxExporter(getFilePath()), valueRow, valueColumn) == originalValue
    }
}
