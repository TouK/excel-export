package pl.touk.excel.export.multisheet

import groovy.transform.TypeChecked
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper

import pl.touk.excel.export.abilities.CellManipulationAbility
import pl.touk.excel.export.abilities.RowManipulationAbility

@Mixin([RowManipulationAbility, CellManipulationAbility])
@TypeChecked
class AdditionalSheet implements SheetManipulator {
    private Sheet sheet
    private CreationHelper creationHelper
    private CellStyle dateCellStyle

    AdditionalSheet(Sheet sheet, CreationHelper creationHelper, CellStyle dateCellStyle) {
        this.sheet = sheet
        this.creationHelper = creationHelper
        this.dateCellStyle = dateCellStyle
    }

    Sheet getSheet() {
        return sheet
    }

    CreationHelper getCreationHelper() {
        return creationHelper
    }

    CellStyle getDateCellStyle() {
        return dateCellStyle
    }

}
