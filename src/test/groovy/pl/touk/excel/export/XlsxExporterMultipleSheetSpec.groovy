package pl.touk.excel.export

class XlsxExporterMultipleSheetSpec extends XlsxExporterSpec {

    void "should fill rows in separate sheets"() {
        given:
        List objectsForDefaultSheet = [new SampleObject(), new SampleObject(), new SampleObject()]
        List objectsForNamedSheet = [new SampleObject(), new SampleObject(), new SampleObject()]
        String sheetName = "sheet2"

        when:
        xlsxReporter.add(objectsForDefaultSheet, SampleObject.propertyNames, 0)
        xlsxReporter.sheet(sheetName).add( objectsForNamedSheet, SampleObject.propertyNames, 0)

        then:
        verifySaved(objectsForDefaultSheet, XlsxExporter.defaultSheetName)
        verifySaved(objectsForNamedSheet, sheetName)
    }

    private void verifySaved(ArrayList<SampleObject> objects, String sheetName) {
        objects.eachWithIndex { SampleObject sampleObject, int i ->
            sampleObject.verifyRowHasSelectedProperties(xlsxReporter.sheet(sheetName), i)
        }
    }
}
