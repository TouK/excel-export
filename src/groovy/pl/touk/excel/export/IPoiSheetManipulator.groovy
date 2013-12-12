package pl.touk.excel.export

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.CellStyle

interface IPoiSheetManipulator {

    Sheet getSheet()

    CreationHelper getCreationHelper()

    CellStyle getDateCellStyle()

}
