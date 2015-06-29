
package pl.touk.excel.export

import javax.servlet.http.HttpServletResponse

class WebXlsxExporter extends XlsxExporter {

    WebXlsxExporter() {
        super()
    }

    WebXlsxExporter(String templateFileNameWithPath) {
        File tmpFile = File.createTempFile('tmpWebXlsx', filenameSuffix)
        this.fileNameWithPath = tmpFile.getAbsolutePath()
        this.workbook = copyAndLoad(templateFileNameWithPath, fileNameWithPath)
        setUp(workbook)
    }

    WebXlsxExporter setResponseHeaders(HttpServletResponse response) {
        setHeaders(response, new Date().format('yyyy-MM-dd_hh-mm-ss') + filenameSuffix)
    }

    WebXlsxExporter setResponseHeaders(HttpServletResponse response, Closure filenameClosure) {
        setHeaders(response, filenameClosure)
    }

    WebXlsxExporter setResponseHeaders(HttpServletResponse response, String filename) {
        setHeaders(response, filename)
    }

    private WebXlsxExporter setHeaders(HttpServletResponse response, def filename) {
        response.setHeader("Content-disposition", "attachment; filename=$filename;")
        response.setHeader("Content-Type", "application/vnd.ms-excel")
        this
    }
}
