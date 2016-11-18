package pl.touk.excel.export

import groovy.io.GroovyPrintStream

class XlsxExporterCreationSpec extends XlsxExporterSpec {

    void "should create document if doesn't exist"() {
        expect:
        !file.exists()

        when:
        xlsxReporter.save()

        then:
        file.exists()
    }

    void "should save to stream"() {
        given:
        GroovyPrintStream outputStream = new GroovyPrintStream(filePath)

        when:
        xlsxReporter.save(outputStream)

        then:
        file.exists()
    }

    void "should open existing document if exists"() {
        given:
        String myCustomValue = "myCustomValue"
        xlsxReporter.putCellValue(1, 1, myCustomValue)
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(filePath)

        then:
        getCellValue(newXlsxReporter, 1, 1) == myCustomValue
    }

    void "should override existing file"() {
        given:
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(filePath)
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        then: "no exceptions are thrown and"
        getCellValue(newXlsxReporter, 1, 1) == "new"
    }

    void "should not override existing file when giving new path"() {
        given:
        xlsxReporter.putCellValue(1, 1, "old")
        xlsxReporter.save()

        when:
        XlsxExporter newXlsxReporter = new XlsxExporter(filePath, testFolder.absolutePath + "/newTestReport.xlsx")
        newXlsxReporter.putCellValue(1, 1, "new")
        newXlsxReporter.save()

        then: "no exceptions are thrown and"
        getCellValue(new XlsxExporter(filePath), 1, 1) == "old"
        //otherwise this sucker keeps the value in cache
        getCellValue(newXlsxReporter, 1, 1) == "new"
    }

    void "should be able to rename initial sheet"() {
        given:
        String otherSheetName = 'something else'

        when:
        XlsxExporter namedSheetExporter = new XlsxExporter()
        namedSheetExporter.setWorksheetName(otherSheetName)

        then:
        namedSheetExporter.sheet.sheetName == otherSheetName
    }

    void "should have default name for initial sheet"() {
        expect:
        new XlsxExporter().sheet.sheetName == XlsxExporter.defaultSheetName
    }
}
