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

    private String worksheetName //TODO: remove it. Let's use a list for all sheets
    @PackageScope CellStyle dateCellStyle //TODO: make it private
    @PackageScope CreationHelper creationHelper //TODO: make it private

    protected XSSFWorkbook workbook

    protected Map<Sheet, AdditionalSheet> additionalSheets = [:]
    private Sheet defaultSheet
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

    private setUp(XSSFWorkbook workbook) {
        this.creationHelper = workbook.getCreationHelper()
        this.dateCellStyle = createDateCellStyle(XlsxExporter.defaultDateFormat)
    }

    // Moves creation of initial sheet away from constructor, allowing its name to change
    Sheet getSheet() {
        defaultSheet = (defaultSheet) ?: withSheet(worksheetName ?: defaultSheetName).sheet
        return defaultSheet
    }

    private CellStyle createDateCellStyle(String expectedDateFormat) {
        CellStyle dateCellStyle = workbook.createCellStyle()
        XSSFDataFormat dateFormat = workbook.createDataFormat()
        dateCellStyle.dataFormat = dateFormat.getFormat(expectedDateFormat)
        return dateCellStyle
    }

    public AdditionalSheet withSheet(String sheetName) {
        Sheet workbookSheet = workbook.getSheet( sheetName ) ?: workbook.createSheet( sheetName )

        // No local sheet representation for it?  Create it.
        if ( !additionalSheets[ workbookSheet ] ) {
            additionalSheets[ workbookSheet ] = new AdditionalSheet(workbookSheet, workbook.creationHelper, dateCellStyle)
        }

        return additionalSheets[ workbookSheet ]
    }

    private XSSFWorkbook copyAndLoad(String templateNameWithPath, String destinationNameWithPath) {
        if(!new File(templateNameWithPath).exists()) {
            throw new IOException("No template file under path: " + templateNameWithPath)
        }
        copy(templateNameWithPath, destinationNameWithPath)
        zipPackage = OPCPackage.open(destinationNameWithPath);
        return new XSSFWorkbook(zipPackage)
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

    //TODO doesn't work
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

    // For the sake of people getting to the underlying POI, workbook itself
    // is made public
    XSSFWorkbook getWorkbook() {
        return workbook
    }
}