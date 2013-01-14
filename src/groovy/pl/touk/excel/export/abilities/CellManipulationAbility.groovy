package pl.touk.excel.export.abilities
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import pl.touk.excel.export.XlsxExporter
import pl.touk.excel.export.getters.Getter

@Category(XlsxExporter)
class CellManipulationAbility {
    XSSFCell getCellAt(int rowNumber, int columnNumber) {
        XSSFRow row = getOrCreateRow(rowNumber, sheet)
        row.getCell((Short) columnNumber)
    }

    XlsxExporter putCellValue(int rowNumber, int columnNumber, String value) {
        getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(creationHelper.createRichTextString(value))
        this
    }

    XlsxExporter putCellValue(int rowNumber, int columnNumber, Getter formatter) {
        putCellValue(rowNumber, columnNumber, formatter.propertyName)
        this
    }

    XlsxExporter putCellValue(int rowNumber, int columnNumber, Number value) {
        getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(value.toDouble())
        this
    }

    XlsxExporter putCellValue(int rowNumber, int columnNumber, Date value) {
        XSSFCell cell = getOrCreateCellAt(rowNumber, columnNumber, sheet)
        cell.setCellValue(value)
        cell.setCellStyle(dateCellStyle)
        this
    }

    XlsxExporter putCellValue(int rowNumber, int columnNumber, Boolean value) {
        getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(value)
        this
    }

    private static XSSFCell getOrCreateCellAt(int rowNumber, int columnNumber, Sheet sheet) {
        (XSSFCell) CellUtil.getCell(getOrCreateRow(rowNumber, sheet), columnNumber)
    }

    private static XSSFRow getOrCreateRow(int rowNumber, Sheet sheet) {
        (XSSFRow) CellUtil.getRow(rowNumber, sheet)
    }
}
