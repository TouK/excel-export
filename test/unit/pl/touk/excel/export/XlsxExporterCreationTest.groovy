
package pl.touk.excel.export

import org.junit.Test
import groovy.io.GroovyPrintStream

class XlsxExporterCreationTest extends XlsxExporterTest {
    @Test
    void shouldCreateDocumentIfDoesntExist() {
        //given
        assert !(new File(getFilePath()).exists())

        //when
        xlsxReporter.save()

        //then
        assert new File(getFilePath()).exists()
    }

    @Test
    void shouldSaveToStream() {
        //given
        GroovyPrintStream outputStream = new GroovyPrintStream(getFilePath())

        //when
        xlsxReporter.save(outputStream)

        //then
        assert new File(getFilePath()).exists()
    }

    @Test
    void shouldOpenExistingDocumentIfExists() {
        //given
        String myCustomValue = "myCustomValue"
        xlsxReporter.putCellValue(1, 1, myCustomValue)
        xlsxReporter.save()

        //when
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath())

        //then
        assert getCellValue(newXlsxReporter, 1, 1) == myCustomValue
    }

    @Test
    void shouldOverrideExistingFile() {
        //given
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        //when
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath())
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        //then no exceptions are thrown and
        assert getCellValue(newXlsxReporter, 1, 1) == "new"
    }

    @Test
    void shouldNotOverrideExistingFileWhenGivingNewPath() {
        //given
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        //when
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath(), testFolder.getAbsolutePath() + "/newTestReport.xlsx")
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        //then no exceptions are thrown and
        assert getCellValue(new XlsxExporter(getFilePath()), 1, 1) == "old" //otherwise this sucker keeps the value in cache
        assert getCellValue(newXlsxReporter, 1, 1) == "new"
    }
}
