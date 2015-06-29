package pl.touk.excel.export

import groovy.io.GroovyPrintStream

class XlsxExporterCreationTest extends XlsxExporterTest {

    void shouldCreateDocumentIfDoesntExist() {
        given:
        assert !(new File(getFilePath()).exists())

        when:
        xlsxReporter.save()

        then:
        new File(getFilePath()).exists()
    }

    void shouldSaveToStream() {
        given:
        GroovyPrintStream outputStream = new GroovyPrintStream(getFilePath())

        when:
        xlsxReporter.save(outputStream)

        then:
        new File(getFilePath()).exists()
    }

    void shouldOpenExistingDocumentIfExists() {
        given:
        String myCustomValue = "myCustomValue"
        xlsxReporter.putCellValue(1, 1, myCustomValue)
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath())

        then:
        getCellValue(newXlsxReporter, 1, 1) == myCustomValue
    }

    void shouldOverrideExistingFile() {
        given:
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath())
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        then: "no exceptions are thrown and"
        getCellValue(newXlsxReporter, 1, 1) == "new"
    }

    void shouldNotOverrideExistingFileWhenGivingNewPath() {
        given:
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(getFilePath(), testFolder.getAbsolutePath() + "/newTestReport.xlsx")
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        then: "no exceptions are thrown and"
        getCellValue(new XlsxExporter(getFilePath()), 1, 1) == "old"
        //otherwise this sucker keeps the value in cache
        getCellValue(newXlsxReporter, 1, 1) == "new"
    }

    void shouldBeAbleToRenameInitialSheet() {
        given:
        String otherSheetName = 'something else'

        when:
        XlsxExporter namedSheetExporter = new XlsxExporter()
        namedSheetExporter.setWorksheetName(otherSheetName)

        then:
        assert namedSheetExporter.sheet.sheetName == otherSheetName
    }

    void shouldHaveDefaultNameForInitialSheet() {
        expect:
        new XlsxExporter().sheet.sheetName == XlsxExporter.defaultSheetName
    }
}
