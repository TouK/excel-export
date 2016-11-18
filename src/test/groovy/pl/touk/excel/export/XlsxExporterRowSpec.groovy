package pl.touk.excel.export

import org.apache.poi.xssf.usermodel.XSSFWorkbook

class XlsxExporterRowSpec extends XlsxExporterSpec {

    void "should fill row at first position"() throws IOException {
        given:
        List rowValues = ["First", "Second", "Third", " "]

        when:
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(rowValues, 1)
    }

    void "should fill row at position"() throws IOException {
        given:
        List rowValues = ["First", "Second", "Third", " "]
        int rowNumber = 15

        when:
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        then:
        verifyValuesAtRow(rowValues, rowNumber)
    }

    void "should fill row at position with zero"() throws IOException {
        given:
        List rowValues = [0]
        int rowNumber = 25

        when:
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        then:
        getCell(new XSSFWorkbook(filePath), rowNumber, 0).numericCellValue == 0
    }

    void "should handle subclasses of valid types in properties"() {
        given:
        java.util.Date handledType = new java.util.Date(123123)
        java.sql.Date subclassOfHandledType = new java.sql.Date(321321)
        List rowValues = [handledType, subclassOfHandledType]

        when:
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        then:
        verifyDateAt(handledType, 1, 0)
        verifyDateAt(subclassOfHandledType, 1, 1)
    }

    void "should fill row from property list"() {
        given:
        SampleObject testObject = new SampleObject()

        when:
        xlsxReporter.add(testObject, SampleObject.propertyNames, 3)

        then:
        testObject.verifyRowHasSelectedProperties(xlsxReporter.withDefaultSheet(), 3)
    }

    void "should fill rows"() {
        given:
        List objects = [new SampleObject(), new SampleObject(), new SampleObject()]

        when:
        xlsxReporter.add(objects, SampleObject.propertyNames, 0)

        then:
        objects.eachWithIndex { SampleObject sampleObject, int i ->
            sampleObject.verifyRowHasSelectedProperties(xlsxReporter.withDefaultSheet(), i)
        }
    }

    void "should fill rows from lists of maps"() {
        given:
        List stones = [[first: 'Keith'], [first: 'Mick']]

        when:
        xlsxReporter.add(stones, ['first'], 0)

        then:
        xlsxReporter.getCellAt(0, 0).stringCellValue == stones[0].first
        xlsxReporter.getCellAt(1, 0).stringCellValue == stones[1].first
    }

    void "should handle nulls when found in map"() {
        given:
        List stones = [[middle: null]]

        when:
        xlsxReporter.add(stones, ['middle'], 0)

        then:
        xlsxReporter.getCellAt(0, 0).stringCellValue == ''
    }

    void "should handle list as to string"() {
        given:
        List<ArrayList> objects = [new SampleObjectWithList(), new SampleObjectWithList()]

        when:
        xlsxReporter.add(objects, ['list'], 0)

        then:
        xlsxReporter.getCellAt(0, 0).stringCellValue == objects[0].list.toString()
        xlsxReporter.getCellAt(1, 0).stringCellValue == objects[1].list.toString()
    }
}
