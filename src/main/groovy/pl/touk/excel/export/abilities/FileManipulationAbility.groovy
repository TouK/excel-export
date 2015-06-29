package pl.touk.excel.export.abilities

import pl.touk.excel.export.XlsxExporter

@Category(XlsxExporter)
class FileManipulationAbility {
    void save(OutputStream outputStream) {
        workbook.write(outputStream)
        outputStream.flush()
        closeZipPackageIfPossible()
    }

    void save() {
        if(fileNameWithPath == null) {
            throw new Exception("No filename given. You cannot create and save a report without giving filename or OutputStream")
        }
        deleteIfAlreadyExists()
        new FileOutputStream(fileNameWithPath).with {
            workbook.write(it)
        }
        closeZipPackageIfPossible()
    }

    void deleteIfAlreadyExists() {
        File existingFile = new File(fileNameWithPath)
        if (existingFile.exists()) {
            existingFile.delete()
        }
    }
}
