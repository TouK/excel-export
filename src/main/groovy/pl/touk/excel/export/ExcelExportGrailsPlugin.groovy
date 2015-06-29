package pl.touk.excel.export

import grails.plugins.*

class ExcelExportGrailsPlugin extends Plugin {
    // the plugin version
    def version = "2.0.0-SNAPSHOT"
    // the version or versions of Grails the plugin is designed for
    def grailsVersion = "3.0.2 > *"

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
    def organization = [name: "TouK", url: "http://touk.pl/"]

    // Any additional developers beyond the author specified above.
    def developers = [[name: "Jakub Nabrdalik", email: "jakubn@gmail.com"], [name: "Mansi Arora", email: "mansi.arora@tothenew.com"]]

    // Location of the plugin's issue tracker.
    def issueManagement = [system: "Github", url: "https://github.com/TouK/excel-export/issues"]

    // Online location of the plugin's browseable source code.
    def scm = [url: "https://github.com/TouK/excel-export"]

    // resources that are excluded from plugin packaging
    def pluginExcludes = [
            "grails-app/views/error.gsp"
    ]
    def profiles = ['web']

    Closure doWithSpring() { {->
        // TODO Implement runtime spring config (optional)
    }
    }

    void doWithDynamicMethods() {
        // TODO Implement registering dynamic methods to classes (optional)
    }

    void doWithApplicationContext() {
        // TODO Implement post initialization spring config (optional)
    }

    void onChange(Map<String, Object> event) {
        // TODO Implement code that is executed when any artefact that this plugin is
        // watching is modified and reloaded. The event contains: event.source,
        // event.application, event.manager, event.ctx, and event.plugin.
    }

    void onConfigChange(Map<String, Object> event) {
        // TODO Implement code that is executed when the project configuration changes.
        // The event is the same as for 'onChange'.
    }

    void onShutdown(Map<String, Object> event) {
        // TODO Implement code that is executed when the application shuts down (optional)
    }
}
