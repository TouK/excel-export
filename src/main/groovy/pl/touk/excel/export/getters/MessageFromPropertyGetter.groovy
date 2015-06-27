
package pl.touk.excel.export.getters

import org.springframework.context.MessageSource
import org.springframework.context.i18n.LocaleContextHolder

class MessageFromPropertyGetter implements Getter {
    private MessageSource messageSource
    private String propertyName
    private Locale locale

    MessageFromPropertyGetter(MessageSource messageSource, String propertyName) {
        this.messageSource = messageSource
        this.propertyName = propertyName
        this.locale = LocaleContextHolder.getLocale()
    }

    MessageFromPropertyGetter(MessageSource messageSource, String propertyName, Locale locale) {
        this.messageSource = messageSource
        this.propertyName = propertyName
        this.locale = locale
    }

    String getPropertyName() {
        return propertyName
    }

    Object getFormattedValue(Object object) {
        return messageSource.getMessage(object.getProperties().get(propertyName), [].toArray(), object.getProperties().get(propertyName), locale)
    }
}
