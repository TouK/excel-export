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


    @Test(expected=IllegalArgumentException.class)
    void shouldThrowExceptionWhenListIsPassedAsArgument() {
        //given
        List<ArrayList> objects = [new SampleObjectWithList(), new SampleObjectWithList()]

        //when
        xlsxReporter.add(objects, ['list'])
    }
}
