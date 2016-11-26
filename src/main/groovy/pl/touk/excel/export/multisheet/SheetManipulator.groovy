package pl.touk.excel.export.multisheet

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Sheet

interface SheetManipulator {

    Sheet getSheet()

    CreationHelper getCreationHelper()

    CellStyle getDateCellStyle()

}
