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
class XlsxExporter {
    static final String sheetName = "Report"
    protected static final String FILENAME_SUFFIX = ".xlsx"
    @PackageScope static final String defaultDateFormat = "yyyy/mm/dd h:mm:ss"

    protected CellStyle dateCellStyle
    protected String fileNameWithPath
    protected CreationHelper creationHelper
    protected XSSFWorkbook workbook
    protected Sheet sheet
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
        this.sheet = createOrLoadSheet(workbook, sheetName)
        this.dateCellStyle = createDateCellStyle(workbook, XlsxExporter.defaultDateFormat)
    }

    private CellStyle createDateCellStyle(XSSFWorkbook workbook, String expectedDateFormat) {
        CellStyle dateCellStyle = workbook.createCellStyle()
        XSSFDataFormat dateFormat = workbook.createDataFormat()
        dateCellStyle.dataFormat = dateFormat.getFormat(expectedDateFormat)
        dateCellStyle
    }

    private Sheet createOrLoadSheet(XSSFWorkbook workbook, String sheetName) {
        workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
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
        this.dateCellStyle = createDateCellStyle(workbook, format)
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