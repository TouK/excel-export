
package pl.touk.excel.export

import org.junit.Test

class XlsxExporterManipulateCellsTest extends XlsxExporterTest {
    @Test
    void shouldPutAndReadCellValue() {
        //given
        String stringValue = "spomething"
        Date dateValue = new Date()
        Long longValue = 1L
        Boolean booleanValue = true

        //when
        xlsxReporter.putCellValue(0, 0, stringValue).
                     putCellValue(0, 1, dateValue).
                     putCellValue(0, 2, longValue).
                     putCellValue(0, 3, booleanValue)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == stringValue
        assert xlsxReporter.getCellAt(0, 1).getDateCellValue() == dateValue
        assert xlsxReporter.getCellAt(0, 2).getNumericCellValue() == longValue
        assert xlsxReporter.getCellAt(0, 3).getBooleanCellValue() == booleanValue
    }
	
	@Test
	void shouldChangeCellAndColumnsStyle() {
		//given
		String stringValue = """Lorem Ipsum is simply dummy text of the printing and \
			typesetting industry. Lorem Ipsum has been the industry's standard dummy \
			text ever since the 1500s, when an unknown printer took a galley of type \
			and scrambled it to make a type specimen book. It has survived not only five \
			centuries, but also the leap into electronic typesetting, remaining essentially"""

		//when
		xlsxReporter.putCellValue(0, 0, stringValue)
		
		//then
		xlsxReporter.getCellAt(0, 0).getCellStyle().setShrinkToFit(true)
		assert xlsxReporter.getCellAt(0, 0)?.getCellStyle().getShrinkToFit()
		
		// column
		xlsxReporter.getSheet().setColumnWidth(0, 70)
		assert xlsxReporter.getSheet().getColumnWidth(0) == 70
		xlsxReporter.getSheet().getColumnStyle(0).setShrinkToFit(true)
		assert xlsxReporter.getSheet().getColumnStyle(0).getShrinkToFit()
		
	} 
}
