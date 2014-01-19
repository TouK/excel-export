package pl.touk.excel.export
import groovy.transform.PackageScope
import groovy.transform.TypeChecked
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import pl.touk.excel.export.abilities.CellManipulationAbility
import pl.touk.excel.export.abilities.FileManipulationAbility
import pl.touk.excel.export.abilities.RowManipulationAbility
import pl.touk.excel.export.multisheet.AdditionalSheet
import pl.touk.excel.export.multisheet.SheetManipulator

@Mixin([RowManipulationAbility, CellManipulationAbility, FileManipulationAbility])
@TypeChecked
class XlsxExporter implements SheetManipulator {
    static final String defaultSheetName = "Report"
    static final String filenameSuffix = ".xlsx"
    @PackageScope static final String defaultDateFormat = "yyyy/mm/dd h:mm:ss"

    private Map<String, AdditionalSheet> sheets = [:]
    private String worksheetName
    private CellStyle dateCellStyle
    private CreationHelper creationHelper
    protected XSSFWorkbook workbook
    protected String fileNameWithPath
    protected OPCPackage zipPackage

    XlsxExporter() {
        this.workbook = new XSSFWorkbook()
        setUp(workbook)
    }

    XlsxExporter(String destinationFileNameWithPath) {
        this.fileNameWithPath = destinationFileNameWithPath
        this.workbook = createOrLoadWorkbook(destinationFileNameWithPath)
        setUp(workbook)
    }

    private XSSFWorkbook createOrLoadWorkbook(String fileNameWithPath) {
        if(new File(fileNameWithPath).exists()) {
            zipPackage = OPCPackage.open(fileNameWithPath);
            return new XSSFWorkbook(zipPackage)
        } else {
            return new XSSFWorkbook()
        }
    }

    XlsxExporter(String templateFileNameWithPath, String destinationFileNameWithPath) {
        this.fileNameWithPath = destinationFileNameWithPath
        this.workbook = copyAndLoad(templateFileNameWithPath, destinationFileNameWithPath)
        setUp(workbook)
    }

    private XSSFWorkbook copyAndLoad(String templateNameWithPath, String destinationNameWithPath) {
        if(!new File(templateNameWithPath).exists()) {
            throw new IOException("No template file under path: " + templateNameWithPath)
        }
        copy(templateNameWithPath, destinationNameWithPath)
        zipPackage = OPCPackage.open(destinationNameWithPath);
        return new XSSFWorkbook(zipPackage)
    }

    private setUp(XSSFWorkbook workbook) {
        this.creationHelper = workbook.getCreationHelper()
        this.dateCellStyle = createDateCellStyle(XlsxExporter.defaultDateFormat)
    }

    Sheet getSheet() {
        if(sheets.isEmpty()) {
            AdditionalSheet additionalSheet = withDefaultSheet()
            sheets.put(worksheetName, additionalSheet)
        }
        return sheets[worksheetName].sheet
    }

    AdditionalSheet withDefaultSheet() {
        worksheetName = worksheetName ?: defaultSheetName
        return sheet(worksheetName)
    }

    AdditionalSheet sheet(String sheetName) {
        if ( !sheets[sheetName] ) {
            Sheet workbookSheet = workbook.getSheet( sheetName ) ?: workbook.createSheet( sheetName )
            sheets[sheetName] = new AdditionalSheet(workbookSheet, workbook.creationHelper, dateCellStyle)
        }
        return sheets[sheetName]
    }

    private void copy(String templateNameWithPath, String destinationNameWithPath) {
        zipPackage = OPCPackage.open(templateNameWithPath);
        XSSFWorkbook originalWorkbook = new XSSFWorkbook(zipPackage)
        new FileOutputStream(destinationNameWithPath).with { OutputStream it ->
            originalWorkbook.write(it)
        }
    }

    XlsxExporter setDateCellFormat(String format) {
        this.dateCellStyle = createDateCellStyle(format)
        return this
    }

    private CellStyle createDateCellStyle(String expectedDateFormat) {
        CellStyle dateCellStyle = workbook.createCellStyle()
        XSSFDataFormat dateFormat = workbook.createDataFormat()
        dateCellStyle.dataFormat = dateFormat.getFormat(expectedDateFormat)
        return dateCellStyle
    }

    void setWorksheetName(String worksheetName) {
        this.worksheetName = worksheetName
    }

    XSSFWorkbook getWorkbook() {
        return workbook
    }

    CellStyle getDateCellStyle() {
        return dateCellStyle
    }

    CreationHelper getCreationHelper() {
        return creationHelper
    }

    //FIXME: nope, that doesn't work
    void finalize() {
        closeZipPackageIfPossible()
    }

    private void closeZipPackageIfPossible() {
        if(zipPackage) {
            try {
                zipPackage.close()
            } finally {
                zipPackage = null
            }
        }
    }
}