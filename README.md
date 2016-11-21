This is the excel-export Grails plugin using [Apache POI](https://poi.apache.org/)

[![Build Status](https://travis-ci.org/TouK/excel-export.svg?branch=master)](https://travis-ci.org/TouK/excel-export) [ ![Download](https://api.bintray.com/packages/grails-excel-export/plugins/jakubnabrdalik.plugins%3Aexcel-export/images/download.svg) ](https://bintray.com/grails-excel-export/plugins/jakubnabrdalik.plugins%3Aexcel-export/_latestVersion)

#What does it do?

This plugin allows for easy exporting of object lists to an Office Open XML workbook (Microsoft Excel 2007+ xlsx) file, while still allowing you to handle the export file on a cell-by-cell basis.

#When should I use it?

There are two scenarios for which this plugin was created:

1. When you want to export data from your controllers for download and want to maintain full control of how you handle this data.

2. When your customer says: 'I want 100 reports in this new project' and nobody has any clue what those reports look like, you can use this plugin as a DSL, i.e. tell your client 'Hey, I've got good news. We have a nice DSL for you, so that you can write all those reports yourself. And it's free!' (or charge them anyway).

In either case, you can export to either a file on disk or to the HTTP response output stream (download as xlsx).

This plugin has been used like this in commercial projects.

#How do I use it?

Say, in your controller you have a list of objects, like so:

```groovy
List<Product> products = productFactory.createProducts()
```

To export selected properties of those products to a file on disk, see the following example, where `withProperties` is a list of properties that are going to be exported to xlsx, in the given order:

```groovy
def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']
new XlsxExporter('/tmp/myReportFile.xlsx').
    add(products, withProperties).
    save()
```

Notice that you can use nested properties (e.g. `price.value`) of your objects.

To add a header row to the spreadshet and make the file downloadable from a controller, you do this:

```groovy
def headers = ['Name', 'Description', 'Valid Till', 'Product Number', 'Price']
def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

new WebXlsxExporter().with {
    setResponseHeaders(response)
    fillHeader(headers)
    add(products, withProperties)
    save(response.outputStream)
}
```

`WebXlsxExporter` is the same thing as `XlsxExporter`, just with the ability to handle HTTP response headers.

You can also manipulate the file on a cell-by-cell basis:

```groovy
new WebXlsxExporter().with {
    setResponseHeaders(response)
    fillRow(["aaa", "bbb", 13, new Date()], 1)
    fillRow(["ccc", "ddd", 87, new Date()], 2)
    putCellValue(3, 3, "Now I'm here")
    save(response.outputStream)
}
```

You can also mix approaches (cell-by-cell, and object list):

```groovy
def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

new WebXlsxExporter().with {
     setResponseHeaders(response)
     fillRow(["aaa", "bbb", 13, new Date()], 1)
     fillRow(["ccc", "ddd", 87, new Date()], 2)
     putCellValue(3, 3, "Now I'm here")
     add(products, withProperties, 4) //NOTICE: we are adding objects starting from line 4
     save(response.outputStream)
}
```

#What about multiple sheets?

If you'd like to work with multiple sheets, just call `sheet(sheetName)` on your exporter. It returns an instance
of `AdditionalSheet` that shares the same row/cell manipulation API as the exporter itself:

```groovy
    List<Product> products = productFactory.createProducts()
    def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

    new WebXlsxExporter().with {
        setResponseHeaders(response)                                                                                                                     print methods of controller
        sheet('second sheet').with {
            fillHeader(withProperties)
            add( products, withProperties )
        }
        save(response.outputStream)
    }
```

You can also mix using additional sheets with a default sheet:

```groovy
    List<Product> products = productFactory.createProducts()
    def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

    new WebXlsxExporter().with {
        setResponseHeaders(response)
        fillHeader(withProperties)
        add( products, withProperties )
        sheet('second sheet').with {
            fillHeader(withProperties)
            add( products, withProperties )
        }
        save(response.outputStream)
    }
```

And if you'd like to change the name of default sheet, just set it before first call:

```groovy
    WebXlsxExporter webXlsxExporter = new WebXlsxExporter()
    webXlsxExporter.setWorksheetName("products")
    webXlsxExporter.with {
        ...
    }
```

#How to export my own types?

This plugin handles basic property types pretty well (String, Date, Boolean, Timestamp, NullObject, Long, Integer, BigDecimal, BigInteger, Byte, Double, Float, Short). It also handles nested properties, and if everything fails, tries to call `toString()`. But sooner or later, you'll want to export a property of a different type the way you like it.
What you need to write, is a Getter. Or, better, a `PropertyGetter`. It's super easy, here is example of one that takes Currency and turns it into a String:

```groovy
class CurrencyGetter extends PropertyGetter<Currency, String> { // From Currency, to String
    CurrencyGetter(String propertyName) {
        super(propertyName)
    }

    @Override
    protected String format(Currency value) {
        return value.displayName // you can do anything you like in here
    }
}
```

The `format()` method allows you customize the value before the object is saved in an xlsx cell.

And, of course, to use it, just add it into the `withProperties` list, like this:

```groovy
def withProperties = ['name', new CurrencyGetter('price.currency'), 'price.value']

new WebXlsxExporter().with {
    setResponseHeaders(response)
    add(products, withProperties)
    save(response.outputStream)
}
```

Of course we could have just used `currency.displayName` in `withProperties`, but you get the idea.

There are two Getters ready for your convenience.

`LongToDatePropertyGetter` gets a long and saves it as a date in xlsx, while `MessageFromPropertyGetter` handles i18n. Speaking of which...

#How to i18n?

To get i18n of headers in your controller, just use controller's existing message method:

```groovy
def headers = [message(code: 'product.name.header'),
    message(code: 'product.description.header'),
    message(code: 'product.validTill.header'),
    message(code: 'product.productNumber.header'),
    message(code: 'price.value.header')]
```

You can do more though. To i18n values, use `MessageFromPropertyGetter`:

```groovy
MessageSource messageSource //injected in the controller automatically by Grails, just declare it

def export() {
    List<Product> products = productFactory.createProducts()

    def headers = ['name', 'type', 'value']
    def withProperties = ['name', new MessageFromPropertyGetter(messageSource, 'type'), 'price.value']

    new WebXlsxExporter().with {
        setResponseHeaders(response)
        fillHeader(headers)
        add(products, withProperties)
        save(response.outputStream)
    }
}
```

This will use grails i18n, based on the value of some property (`type` in the example above) of your objects.

#I want fancy diagrams, colours, and other stuff in my Excel!

Making xlsx files look really great with Apache POI is pretty fun but not very efficient. So we have found out that it's easier to create a template manually (in MS Excel or Open Office), load this template in your code, fill it up with data, and hand it back to the user.

For this scenario, every constructor takes a path to a template file (just a normal xlsx file).

After loading the template, fill the data, and save to the output stream

```groovy
new WebXlsxExporter('/tmp/myTemplate.xlsx').with {
    setResponseHeaders(response)
    add(products, withProperties)
    save(response.outputStream)
}
```

If you just want to save the file to disk instead of a stream, use a different constructor:

```groovy
new XlsxExporter('/tmp/myTemplate.xlsx', '/tmp/myReport.xlsx')
```

If you just open an existing file, and save it, like this:

```groovy
new XlsxExporter('/tmp/myReport.xlsx').with {
    add(products, withProperties)
    save()
}
```

you are going to overwrite it.

#But I don't want no template (for whatever reason)

Ok, so if you don't want to use a temple, but want to format a cell style directly in the code, you can still do that.

You can get the cell style like this:

```groovy
xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString()
```

Of course there is a corresponding `setCellStyle()` method, but this is a part of the [Apache POI API](https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFCell.html), and is outside the scope of this plugin.

#How to get it installed?

Like any other Grails plugin, just add to the `dependencies` block of your app's build.gradle file:
```groovy
dependencies {
    compile "org.grails.plugins:excel-export:2.0.1"
}
```

Excluding xerces may or may not be needed, depending on your setup. If you get

```
Error executing script RunApp: org/apache/xerces/dom/DeferredElementImpl (Use --stacktrace to see the full trace)
```

you NEED to exclude xercesImpl to use ApachePOI. Don't worry, it won't break anything.

If you have more strange problems with xml and you are using Java 7, exclude xml-apis as well.

To understand why you need to exclude anything, please take a look here: http://stackoverflow.com/questions/11677572/dealing-with-xerces-hell-in-java-maven

If you want a working example, clone this project: https://github.com/TouK/excel-export-samples

#Are there any other libs like this?

Yeah, there's plenty of them. But most were too simplistic or too 'automagical' for my needs. Apache POI is pretty simple to use itself (and has fantastic API) but we needed something even simpler for several projects. Also a bit DSL-like so our customers could write reports on their own. After preparing a few getters for our custom objects, this is what we ended up with:

```groovy
def withProperties = ["id", "name", "inProduction", "workLogCount", "createdAt", "createdAtDate", asDate("firstEventTime"),
    firstUnacknowledgedTime(), firstUnacknowledged(), firstTakeOwnershipTime(), firstTakeOwnership(),
    firstReleaseOwnershipTime(), firstReleaseOwnership(), firstClearedTime(), firstCleared(),
    firstDeferedTime(), firstDefered(), firstUndeferedTime(), firstUndefered(), childConnectedTime(), childConnected(),
    parentConnectedTime(), parentConnected(), parentDisconnectedTime(), parentDisconnected(),
    childDisconnectedTime(), childDisconnected(), childOrphanedTime(), childOrphaned(), createdTime(), created(),
    updatedTime(), updated(), workLogAddedTime(), workLogAdded()]

def reporter = new XlsxReporter("/tmp/sampleTemplate.xlsx")

reporter.with {
    fillHeader withProperties
    add events, withProperties
    save "/tmp/sampleReport1.xlsx"
}
```

All the methods in `withProperties` are static imports generating a new instance of a corresponding `PropertyGetter` implementation. To our surprise, this worked really well with some clients, who started writing their own reports instead of paying us for doing the boring work.

Hope it helps.

#License

Copyright 2012-2014 TouK

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

#Changes

2.0.1 upgrade poi to 3.12 (thanks to Sergio Maria Matone) and Grails to 3+ (thanks to mansiarora)

0.2.1 calling toString() on unhandled property types, instead of throwing IllegalArgumentException

0.2.0 working with multiple sheets and renaming default sheet

0.1.10 not exporting release plugin dependency anymore (Issue #14)

0.1.9 upgrade to release plugin 3.0.1 (run 'grails refresh-dependencies' if you have problems in grails 2.3.2)

0.1.8: fix for grails 2.3.1 (groovy changing how Mixins see private methods)

0.1.7: fixed Property Type Validation not accepting subclasses

0.1.6: handling maps in object properties
