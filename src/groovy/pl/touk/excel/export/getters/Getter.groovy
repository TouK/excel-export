package pl.touk.excel.export.getters

interface Getter<DestinationFormat> {
    String getPropertyName()
    DestinationFormat getFormattedValue(Object object)
}
