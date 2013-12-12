package pl.touk.excel.export
import groovy.transform.PackageScope
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import pl.touk.excel.export.abilities.CellManipulationAbility
import pl.touk.excel.export.abilities.FileManipulationAbility
import pl.touk.excel.export.abilities.RowManipulationAbility

@Mixin([RowManipulationAbility, CellManipulationAbility, FileManipulationAbility])
class XlsxExporter implements IPoiSheetManipulator {
    protected static final String sheetName = "Report"
    @PackageScope static final String defaultDateFormat = "yyyy/mm/dd h:mm:ss"

    String worksheetName
    CellStyle dateCellStyle
    CreationHelper creationHelper

    // For the sake of people getting to the underlying POI, workbook itself
    // is made public
    XSSFWorkbook workbook

    protected Map sheets = [:]
    protected String fileNameWithPath
    protected OPCPackage zipPackage

    private Sheet sheet

    // Moves creation of initial sheet away from constructor, allowing its name to change
    public Sheet getSheet() {
        if ( !sheet ) {
            sheet = withSheet(worksheetName ?: sheetName).sheet
        }

        return sheet
    }

    XlsxExporter() {
        this.workbook = new XSSFWorkbook()
        setUp()
    }

    XlsxExporter(String fileNameWithPath) {
        this.fileNameWithPath = fileNameWithPath
        this.workbook = createOrLoadWorkbook(fileNameWithPath)
        setUp()
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
        setUp()
    }

    private setUp() {
        this.creationHelper = workbook.getCreationHelper()
        this.dateCellStyle = createDateCellStyle(XlsxExporter.defaultDateFormat)
    }

    private CellStyle createDateCellStyle(String expectedDateFormat) {
        CellStyle dateCellStyle = workbook.createCellStyle()
        XSSFDataFormat dateFormat = workbook.createDataFormat()
        dateCellStyle.dataFormat = dateFormat.getFormat(expectedDateFormat)
        dateCellStyle
    }

    public ExcelExportSheet withSheet(String sheetName) {
        Sheet workbookSheet = workbook.getSheet( sheetName ) ?: workbook.createSheet( sheetName )

        // No local sheet representation for it?  Create it.
        if ( !sheets[ workbookSheet ] ) {
            sheets[ workbookSheet ] = new ExcelExportSheet(
                exporter: this,
                sheet: workbookSheet
            )
        }

        return sheets[ workbookSheet ]
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
        new FileOutputStream(destinationNameWithPath).with {
            originalWorkbook.write(it)
        }
    }

    XlsxExporter setDateCellFormat(String format) {
        this.dateCellStyle = createDateCellStyle(format)
        this
    }

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