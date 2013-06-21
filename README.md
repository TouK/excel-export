This is excel-export Grails plugin using Apache POI

#What does it do?

It exports your objects to an xlsx (MS Excel 2007+) file, while still allowing you to handle the export file on a cell-by-cell basis.

#When should I use it?

There are two scenarios on which this plugin was created:

1. When you want to export data from your controllers ('download to excel' button) and want to maintain full control of how you handle this data.

2. When your customer says: 'I want 100 reports in this new project' and nobody has any clue what those reports look like, you can use this plugin as a DSL, i.e. tell your client 'Hey, I've got good news. We have a nice DSL for you, so that you can write all those reports yourself. And it's free!' (or charge them anyway).

In both cases you can export either to a file on disk, or to outputStream (download as xlsx).

This plugin has been used like this in commercial projects.

#How do I use it?

Say, in your controller you have a list of objects, like so:

        List<Product> products = productFactory.createProducts()

To export chosen properties of those products to a file, you do this:

        def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']
        new XlsxExporter('/tmp/myReportFile.xlsx').
            add(products, withProperties).
            save()

withProperties is a list of properties that are going to be exported to xlsx, in the given order.

Notice, that you can use nested properties (price.value) of your objects.

To add a header, and make it downloadable from a controller, you do this:

      def headers = ['Name', 'Description', 'Valid Till', 'Product Number', 'Price']
      def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

      new WebXlsxExporter().with {
          setResponseHeaders(response)
          fillHeader(headers)
          add(products, withProperties)
          save(response.outputStream)
      }

WebXlsxExporter is the same thing as XlsxExporter, plus it handles HTTP response headers.

To manipulate the file on a cell-by-cell basis, you do this:

      new WebXlsxExporter().with {
          setResponseHeaders(response)
          fillRow(["aaa", "bbb", 13, new Date()], 1)
          fillRow(["ccc", "ddd", 87, new Date()], 2)
          putCellValue(3, 3, "Now I'm here")
          save(response.outputStream)
      }

You can mix playing on cell-by-cell approach with add method

    def withProperties = ['name', 'description', 'validTill', 'productNumber', 'price.value']

    new WebXlsxExporter().with {
        setResponseHeaders(response)
        fillRow(["aaa", "bbb", 13, new Date()], 1)
        fillRow(["ccc", "ddd", 87, new Date()], 2)
        putCellValue(3, 3, "Now I'm here")
        add(products, withProperties, 4) //NOTICE: we are adding objects starting from line 4
        save(response.outputStream)
    }

#How to export my own types?

This plugin handles basic property types pretty well (String, Date, Boolean, Timestamp, NullObject, Long, Integer, BigDecimal, BigInteger, Byte, Double, Float, Short), it also handles nested properties, but sooner or later, you'll want to export a property of a different type.
What you need to write, is a Getter. Or, better, a PropertyGetter. It's super easy, here is example of one that takes Currency and turns it into a String

    class CurrencyGetter extends PropertyGetter<Currency, String> { //From Currency, to String
        CurrencyGetter(String propertyName) {
            super(propertyName)
        }

        @Override
        protected String format(Currency value) {
            return value.displayName //you can do anything you like in here
        }
    }

The 'format' method, allows you to do anything, before the object is saved in an xlsx cell.

And, of course, to use it, just add it into withProperties list, like this:

    def withProperties = ['name', new CurrencyGetter('price.currency'), 'price.value']

    new WebXlsxExporter().with {
        setResponseHeaders(response)
        add(products, withProperties)
        save(response.outputStream)
    }

Of course we could have just used 'currency.displayName' in withProperties, but you get the idea.

There are two Getters ready for your convenience.

LongToDatePropertyGetter gets a long and saves it as a date in xlsx, while MessageFromPropertyGetter handles i18n. Speaking about which...



#How to i18n?

To get i18n of headers in your controller, just use controller's message  method:

        def headers = [message(code: 'product.name.header'),
                       message(code: 'product.description.header'),
                       message(code: 'product.validTill.header'),
                       message(code: 'product.productNumber.header'),
                       message(code: 'price.value.header')]

You can do more though. To i18n values, use MessageFromPropertyGetter:

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

This will use grails i18n, based on the value of some property ('type' in here) of your objects.


#I want fancy diagrams, colours, and other stuff in my Excel!

Making xlsx files look really great with Apache POI is pretty fun. But not very efficient. So we have found out, that it's easier to create a template manually (in MS Excel or Open Office), load this template in your code, fill it up with data, and handle back to the user.

For this scenario, every constructor takes a path to a template file (just normal xlsx file).

After loading the template, fill the data, and save to the output stream

         new WebXlsxExporter('/tmp/myTemplate.xlsx').with {
             setResponseHeaders(response)
             add(products, withProperties)
             save(response.outputStream)
         }

If you just want to save the file to disk instead of a stream, use a different constructor:

        new XlsxExporter('/tmp/myTemplate.xlsx', '/tmp/myReport.xlsx")

If you just open an existing file, and save it, like this:

       new XlsxExporter('/tmp/myReport.xlsx").with {
           add(products, withProperties)
           save()
       }

you are going to override it.

#But I don't want no template (for whatever reason)

Ok, so if you don't want to use a temple, but want to format a cell style directly in the code, you can still do that.

You can get the cell style like this:

        xlsxReporter.getCellAt(0, 0).getCellStyle().getDataFormatString()

and of course you have a corresponding setCellStyle method, but this is pure Apache POI, not my plugin, so you have the documentation here: http://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFCell.html


#How to get it installed?

The plugin is released like all other Grails plugins, to grailsPlugin repo.
Here is what you need to add to your BuildConfig.groovy

    grails.project.dependency.resolution = {
        inherits("global") {
            ...
            excludes 'xercesImpl'   //#1 important thing
        }

        plugins {
            runtime (":excel-export:0.1.6")     //#2 important thing
            ...
        }
    ...

Excluding xerces may or may not be needed, depending on your setup. If you get

    Error executing script RunApp: org/apache/xerces/dom/DeferredElementImpl (Use --stacktrace to see the full trace)

you NEED to exclude xercesImpl to use ApachePOI. Don't worry, it won't break anything.

If you have more strange problems with xml and you are using Java 7, exclude xml-apis as well:

    grails.project.dependency.resolution = {
        inherits("global") {
            ...
            excludes 'xercesImpl', 'xml-apis'
        }
    ...

To understand why you need to exclude anything, please take a look here: http://stackoverflow.com/questions/11677572/dealing-with-xerces-hell-in-java-maven

If you want a working example, clone this project: https://github.com/TouK/excel-export-samples

#Are there any other libs like this?

Yeah, there's plenty of them. But most were too simplistic or too 'automagical' for my needs. Apache POI is pretty simple to use itself (and has fantastic API) but we needed something even simpler for several projects. Also a bit DSL-like so our customers could write reports on their own. After preparing a few getters for our custom objects, this is what we ended up with:

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

All the methods in 'withProperties' are static imports generating a new instance of a corresponding PropertyGetter implementation. To our surprise, this worked really well with some clients, who started writing their own reports instead of paying us for doing the boring work.

Hope it helps.

#Licence

Apache Licence v2.0

#Changes

0.1.6: handling maps in object properties





