package pl.touk.excel.export
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Before
import org.junit.Rule
import org.junit.rules.TemporaryFolder
import pl.touk.excel.export.getters.Getter

abstract class XlsxTestOnTemporaryFolder {
    @Rule public TemporaryFolder temporaryFolder = new TemporaryFolder();
    File testFolder

    @Before
    void setUpTestFolder() {
        testFolder = temporaryFolder.newFolder("newFolder")
    }

    protected String getFilePath() {
        return testFolder.getAbsolutePath() + "/testReport.xlsx"
    }

    protected void verifyValuesAtRow(List<Object> values, int rowNumber) {
        XSSFWorkbook workbook = new XSSFWorkbook(getFilePath())
        values.eachWithIndex { Object value, int index ->
            String propertyName = (value instanceof Getter) ? value.propertyName : value
            assert getCellValue(workbook, rowNumber, index) == propertyName
        }
    }

    protected String getCellValue(XlsxExporter xlsxExporter, int rowNumber, int columnNumber) {
        return getCellValue(xlsxExporter.workbook, rowNumber, columnNumber)
    }

    protected String getCellValue(XSSFWorkbook workbook, int rowNumber, int columnNumber) {
        return getCell(workbook, rowNumber, columnNumber).getStringCellValue()
    }

    protected XSSFCell getCell(XSSFWorkbook workbook, int rowNumber, int columnNumber) {
        return workbook.getSheet(XlsxExporter.sheetName).getRow((Short)rowNumber).getCell((Short)columnNumber)
    }
}
