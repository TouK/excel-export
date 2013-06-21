class ExcelExportGrailsPlugin {
    // the plugin version
    def version = "0.1.6"
    // the version or versions of Grails the plugin is designed for
    def grailsVersion = "2.0 > *"
    // the other plugins this plugin depends on
    def dependsOn = [:]
    // resources that are excluded from plugin packaging
    def pluginExcludes = [
        "grails-app/views/error.gsp"
    ]

    // TODO Fill in these fields
    def title = "Excel Export Plugin" // Headline display name of the plugin
    def author = "Jakub Nabrdalik"
    def authorEmail = "jakubn@gmail.com"
    def description = 'This plugin helps you export data in Excel (xlsx) format, using Apache POI.'

    // URL to the plugin's documentation
    def documentation = "https://github.com/TouK/excel-export/blob/master/README.md"

    // Extra (optional) plugin metadata

    // License: one of 'APACHE', 'GPL2', 'GPL3'
    def license = "APACHE"

    // Details of company behind the plugin (if there is one)
    def organization = [ name: "TouK", url: "http://touk.pl/" ]

    // Any additional developers beyond the author specified above.
    def developers = [ [ name: "Jakub Nabrdalik", email: "jakubn@gmail.com" ]]

    // Location of the plugin's issue tracker.
    def issueManagement = [ system: "Github", url: "https://github.com/TouK/excel-export/issues" ]

    // Online location of the plugin's browseable source code.
    def scm = [ url: "https://github.com/TouK/excel-export" ]

    def doWithWebDescriptor = { xml ->
        // TODO Implement additions to web.xml (optional), this event occurs before
    }

    def doWithSpring = {
        // TODO Implement runtime spring config (optional)
    }

    def doWithDynamicMethods = { ctx ->
        // TODO Implement registering dynamic methods to classes (optional)
    }

    def doWithApplicationContext = { applicationContext ->
        // TODO Implement post initialization spring config (optional)
    }

    def onChange = { event ->
        // TODO Implement code that is executed when any artefact that this plugin is
        // watching is modified and reloaded. The event contains: event.source,
        // event.application, event.manager, event.ctx, and event.plugin.
    }

    def onConfigChange = { event ->
        // TODO Implement code that is executed when the project configuration changes.
        // The event is the same as for 'onChange'.
    }

    def onShutdown = { event ->
        // TODO Implement code that is executed when the application shuts down (optional)
    }
}
