package pl.touk.excel.export.getters

import spock.lang.Specification
import spock.lang.Unroll

class PropertyGetterSpec extends Specification {

    @Unroll
    void "as-is property getter returns an unmodified value of '#raw'"() {
        given:
        Map object = [asIs: raw]
        PropertyGetter getter = new AsIsPropertyGetter('asIs')

        when:
        def formatted = getter.getFormattedValue(object)

        then:
        formatted == raw

        where:
        raw        | _
        0          | _
        ""         | _
        new Date() | _
        1L         | _
        true       | _
        null       | _
        1.0f       | _
        1.0d       | _
    }

    @Unroll
    void "long-to-date property getter returns a date value of '#dateVal'"() {
        given:
        Map object = [raw: longVal]
        PropertyGetter getter = new LongToDatePropertyGetter('raw')

        when:
        def formatted = getter.getFormattedValue(object)

        then:
        formatted == dateVal

        where:
        longVal   | dateVal
        100l      | new Date(100l)
        99999999l | new Date(99999999l)
    }
}
