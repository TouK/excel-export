package pl.touk.excel.export

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Before
import org.junit.Rule
import org.junit.rules.TemporaryFolder
import pl.touk.excel.export.getters.Getter
import spock.lang.Specification

abstract class XlsxTestOnTemporaryFolder extends Specification {
    @Rule
    TemporaryFolder temporaryFolder = new TemporaryFolder()
    File testFolder

    @Before
    void setUpTestFolder() {
        testFolder = temporaryFolder.root
    }

    protected String getFilePath() {
        return testFolder.absolutePath + "/testReport.xlsx"
    }

    protected File getFile() {
        new File(filePath)
    }

    protected void verifyDateAt(Date date, int rowNumber, int columnNumber) {
        XSSFWorkbook workbook = new XSSFWorkbook(filePath)
        assert getCell(workbook, rowNumber, columnNumber).getDateCellValue() == date
    }

    protected void verifyValuesAtRow(List<Object> values, int rowNumber) {
        XSSFWorkbook workbook = new XSSFWorkbook(filePath)
        values.eachWithIndex { Object value, int index ->
            String propertyName = (value instanceof Getter) ? value.propertyName : value
            assert getCellValue(workbook, rowNumber, index) == propertyName
        }
    }

    protected String getCellValue(XlsxExporter xlsxExporter, int rowNumber, int columnNumber) {
        return getCellValue(xlsxExporter.workbook, rowNumber, columnNumber)
    }

    protected String getCellValue(XSSFWorkbook workbook, int rowNumber, int columnNumber) {
        return getCell(workbook, rowNumber, columnNumber).stringCellValue
    }

    protected XSSFCell getCell(XSSFWorkbook workbook, int rowNumber, int columnNumber) {
        return workbook.getSheet(XlsxExporter.defaultSheetName).getRow(rowNumber).getCell(columnNumber)
    }
}
