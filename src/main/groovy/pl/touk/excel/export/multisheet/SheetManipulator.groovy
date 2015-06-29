package pl.touk.excel.export.multisheet

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.CellStyle

interface SheetManipulator {

    Sheet getSheet()

    CreationHelper getCreationHelper()

    CellStyle getDateCellStyle()

}
