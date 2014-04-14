package pl.touk.excel.export
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Test

class XlsxExporterRowTest extends XlsxExporterTest {

    @Test
    void shouldFillRowAtFirstPosition() throws IOException {
        //given
        List rowValues = ["First", "Second", "Third", " "]

        //when
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(rowValues, 1)
    }

    @Test
    void shouldFillRowAtPosition() throws IOException {
        //given
        List rowValues = ["First", "Second", "Third", " "]
        int rowNumber = 15

        //when
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        //then
        verifyValuesAtRow(rowValues, rowNumber)
    }

    @Test
    void shouldFillRowAtPositionWithZero() throws IOException {
        //given
        List rowValues = [0]
        int rowNumber = 25

        //when
        xlsxReporter.fillRow(rowValues, rowNumber)
        xlsxReporter.save()

        //then
        assert getCell(new XSSFWorkbook(getFilePath()), rowNumber, 0).getNumericCellValue() == 0
    }

    @Test
    void shouldHandleSubclassesOfValidTypesInProperties() {
        //given
        java.util.Date handledType = new java.util.Date(123123)
        java.sql.Date subclassOfHandledType = new java.sql.Date(321321)
        List rowValues = [handledType, subclassOfHandledType ]

        //when
        xlsxReporter.fillRow(rowValues)
        xlsxReporter.save()

        //then
        verifyDateAt(handledType, 1, 0)
        verifyDateAt(subclassOfHandledType, 1, 1)
    }

    @Test
    void shouldFillRowFromPropertyList() {
        //given
        SampleObject testObject = new SampleObject()

        //when
        xlsxReporter.add(testObject, SampleObject.propertyNames, 3)

        //then
        testObject.verifyRowHasSelectedProperties(xlsxReporter.withDefaultSheet(), 3)
    }

    @Test
    void shouldFillRowEvaluatingPropertyList() {
        //given
        SampleObject testObject = new SampleObject()
        def propertyExpressions = [
           'stringValue',
           { it.stringValue },
           { it.child.stringValue.toUpperCase() },
           { it.booleanValue?" it is true ":"it is false " },
           { it.shortValue + it.integerValue },
           { Math.sqrt(it.doubleValue) }
        ]

        //when
        xlsxReporter.add(testObject, propertyExpressions, 0)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == testObject.stringValue
        assert xlsxReporter.getCellAt(0, 1).getStringCellValue() == testObject.stringValue
        assert xlsxReporter.getCellAt(0, 2).getStringCellValue() == testObject.child.stringValue.toUpperCase()
        assert xlsxReporter.getCellAt(0, 3).getStringCellValue() == testObject.booleanValue?" it is true ":"it is false "
        assert xlsxReporter.getCellAt(0, 4).getNumericCellValue() == testObject.integerValue + testObject.shortValue
        assert xlsxReporter.getCellAt(0, 5).getNumericCellValue() == Math.sqrt(testObject.doubleValue)
    }
    
    @Test
    void shouldFillRows() {
        //given
        List objects = [new SampleObject(), new SampleObject(), new SampleObject()]

        //when
        xlsxReporter.add(objects, SampleObject.propertyNames, 0)

        //then
        objects.eachWithIndex { SampleObject sampleObject, int i ->
            sampleObject.verifyRowHasSelectedProperties(xlsxReporter.withDefaultSheet(), i)
        }
    }

    @Test
    void shouldFillRowsFromListsOfMaps() {
        //given
        List stones = [ [first: 'Keith'], [first: 'Mick'] ]

        //when
        xlsxReporter.add(stones, ['first'], 0)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == stones[0].first
        assert xlsxReporter.getCellAt(1, 0).getStringCellValue() == stones[1].first
    }

    @Test
    void shoulHanldeNullsWhenFoundInMap() {
        //given
        List stones = [ [middle: null] ]

        //when
        xlsxReporter.add(stones, ['middle'], 0)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == ''
    }

    @Test
    void shouldHanldeListAsToString() {
        //given
        List<ArrayList> objects = [new SampleObjectWithList(), new SampleObjectWithList()]

        //when
        xlsxReporter.add(objects, ['list'], 0)

        //then
        assert xlsxReporter.getCellAt(0, 0).getStringCellValue() == objects[0].list.toString()
        assert xlsxReporter.getCellAt(1, 0).getStringCellValue() == objects[1].list.toString()
    }
}
