package pl.touk.excel.export

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper

import pl.touk.excel.export.abilities.CellManipulationAbility
import pl.touk.excel.export.abilities.RowManipulationAbility

@Mixin([RowManipulationAbility, CellManipulationAbility])
class ExcelExportSheet implements IPoiSheetManipulator {

    protected Sheet sheet
    XlsxExporter exporter

    /* Support the getters in CellManipulationAbility */
    Sheet getSheet() {
        return sheet
    }

    CreationHelper getCreationHelper() {
        return exporter?.creationHelper
    }

    CellStyle getDateCellStyle() {
        return exporter?.dateCellStyle
    }

}
