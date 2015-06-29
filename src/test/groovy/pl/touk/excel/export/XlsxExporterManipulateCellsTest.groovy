package pl.touk.excel.export

class XlsxExporterManipulateCellsTest extends XlsxExporterTest {

    void shouldPutAndReadCellValue() {
        given :
        String stringValue = "spomething"
        Date dateValue = new Date()
        Long longValue = 1L
        Boolean booleanValue = true

        when :
        xlsxReporter.putCellValue(0, 0, stringValue).
                     putCellValue(0, 1, dateValue).
                     putCellValue(0, 2, longValue).
                     putCellValue(0, 3, booleanValue)

        then :
        xlsxReporter.getCellAt(0, 0).getStringCellValue() == stringValue
        xlsxReporter.getCellAt(0, 1).getDateCellValue() == dateValue
        xlsxReporter.getCellAt(0, 2).getNumericCellValue() == longValue
        xlsxReporter.getCellAt(0, 3).getBooleanCellValue() == booleanValue
    }
	
	void shouldChangeCellAndColumnsStyle() {
		given :
		String stringValue = """Lorem Ipsum is simply dummy text of the printing and \
			typesetting industry. Lorem Ipsum has been the industry's standard dummy \
			text ever since the 1500s, when an unknown printer took a galley of type \
			and scrambled it to make a type specimen book. It has survived not only five \
			centuries, but also the leap into electronic typesetting, remaining essentially"""
		xlsxReporter.putCellValue(0, 0, stringValue)
		
		when:
		xlsxReporter.getCellAt(0, 0).getCellStyle().setShrinkToFit(true)
		then:
        xlsxReporter.getCellAt(0, 0)?.getCellStyle()?.getShrinkToFit()
		
		// column
		when:
        xlsxReporter.getSheet().setColumnWidth(0, 70)
		then:
        xlsxReporter.getSheet().getColumnWidth(0) == 70

        when:
        xlsxReporter.getSheet().getColumnStyle(0).setShrinkToFit(true)
		then:
		xlsxReporter.getSheet().getColumnStyle(0).getShrinkToFit()
	}
}
