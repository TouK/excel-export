package pl.touk.excel.export

import org.junit.Before

abstract class XlsxExporterTest extends XlsxTestOnTemporaryFolder {
    XlsxExporter xlsxReporter

    @Before
    void setUpReporter() {
        xlsxReporter = new XlsxExporter(getFilePath())
    }

}

