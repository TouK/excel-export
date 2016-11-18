package pl.touk.excel.export

class XlsxExporterManipulateCellsSpec extends XlsxExporterSpec {

    void "should put and read cell value"() {
        given:
        String stringValue = "spomething"
        Date dateValue = new Date()
        Long longValue = 1L
        Boolean booleanValue = true

        when:
        xlsxReporter.putCellValue(0, 0, stringValue).
                     putCellValue(0, 1, dateValue).
                     putCellValue(0, 2, longValue).
                     putCellValue(0, 3, booleanValue)

        then:
        xlsxReporter.getCellAt(0, 0).stringCellValue == stringValue
        xlsxReporter.getCellAt(0, 1).dateCellValue == dateValue
        xlsxReporter.getCellAt(0, 2).numericCellValue == longValue
        xlsxReporter.getCellAt(0, 3).booleanCellValue == booleanValue
    }

    void "should change cell and columns style"() {
        given:
        String stringValue = """Lorem Ipsum is simply dummy text of the printing and \
            typesetting industry. Lorem Ipsum has been the industry's standard dummy \
            text ever since the 1500s, when an unknown printer took a galley of type \
            and scrambled it to make a type specimen book. It has survived not only five \
            centuries, but also the leap into electronic typesetting, remaining essentially"""
        xlsxReporter.putCellValue(0, 0, stringValue)

        when:
        xlsxReporter.getCellAt(0, 0).cellStyle.setShrinkToFit(true)
        then:
        xlsxReporter.getCellAt(0, 0)?.cellStyle?.shrinkToFit

        // column
        when:
        xlsxReporter.sheet.setColumnWidth(0, 70)
        then:
        xlsxReporter.sheet.getColumnWidth(0) == 70

        when:
        xlsxReporter.sheet.getColumnStyle(0).setShrinkToFit(true)
        then:
        xlsxReporter.sheet.getColumnStyle(0).shrinkToFit
    }
}
