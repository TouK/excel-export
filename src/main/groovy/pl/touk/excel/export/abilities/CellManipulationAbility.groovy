package pl.touk.excel.export.abilities
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import pl.touk.excel.export.getters.Getter
import pl.touk.excel.export.multisheet.SheetManipulator

@Category(SheetManipulator)
class CellManipulationAbility {
    XSSFCell getCellAt(int rowNumber, int columnNumber) {
        XSSFRow row = CellManipulationAbility.getOrCreateRow(rowNumber, sheet)
        row.getCell((Short) columnNumber)
    }

    SheetManipulator putCellValue(int rowNumber, int columnNumber, String value) {
        CellManipulationAbility.getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(getCreationHelper().createRichTextString(value))
        return this
    }

    SheetManipulator putCellValue(int rowNumber, int columnNumber, Getter formatter) {
        CellManipulationAbility.putCellValue(rowNumber, columnNumber, formatter.propertyName)
        return this
    }

    SheetManipulator putCellValue(int rowNumber, int columnNumber, Number value) {
        CellManipulationAbility.getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(value.toDouble())
        return this
    }

    SheetManipulator putCellValue(int rowNumber, int columnNumber, Date value) {
        XSSFCell cell = CellManipulationAbility.getOrCreateCellAt(rowNumber, columnNumber, sheet)
        cell.setCellValue(value)
        cell.setCellStyle(dateCellStyle)
        return this
    }

    SheetManipulator putCellValue(int rowNumber, int columnNumber, Boolean value) {
        CellManipulationAbility.getOrCreateCellAt(rowNumber, columnNumber, sheet).setCellValue(value)
        return this
    }

    private static XSSFCell getOrCreateCellAt(int rowNumber, int columnNumber, Sheet sheet) {
        (XSSFCell) CellUtil.getCell(getOrCreateRow(rowNumber, sheet), columnNumber)
    }

    private static XSSFRow getOrCreateRow(int rowNumber, Sheet sheet) {
        (XSSFRow) CellUtil.getRow(rowNumber, sheet)
    }
}
