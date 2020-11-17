import os
import uno
import traceback
import logging
from com.sun.star.beans import PropertyValue

class oodocument:
    def __init__(self, orig, host='0.0.0.0', port='8001'):
        self.orig = orig
        self.host = host
        self.port = port
        self.font_color = None
        self.font_back_color = None
        self.__set_document()

    def __str__(self):
        return f'{self.orig}'

    def __set_document(self):
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()
        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)
        try:
            # connect to the running office
            ctx = resolver.resolve(
                f'uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext')
        except Exception as e:
            logging.error(traceback.format_exc())
            raise

        # get the central desktop object
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        self.document = desktop.loadComponentFromURL(
            absoluteUrl(self.orig), "_blank", 0, ())

    def convert_to(self, dest, format):
        # Filters https://git.libreoffice.org/core/+/refs/heads/master/filter/source/config/fragments/filters
        formats = {
            'pdf': 'writer_pdf_Export',
            'txt': 'Text',
            'docx': 'MS Word 2007 XML',
            'odt': 'writer8',
            'doc': 'MS Word 97'
        }
        filter = PropertyValue()
        filter.Name = 'FilterName'
        filter.Value = formats[format]
        self.save(dest, filter)

    def replace_with(self, data={}, dest=None, format=None):
        search = self.document.createSearchDescriptor()
        for find, replace in data.items():
            self.__find_and_replace(self.document, search, find, replace)
        if dest is None and format is None:
            self.save()
        else:
            self.convert_to(dest, format)

    def save(self, dest=None, filter=None):
        if dest is None and filter is None:
            self.document.store()
        else:
            self.document.storeToURL( absoluteUrl(dest), (filter,))

    def dispose(self):
        self.document.dispose()

    def __find_and_replace(self, document, search, find=None, replace=None):
        """This function searches and replaces. Create search, call function findFirst, and finally replace what we found."""
        # What to search for
        search.SearchString = find
        search.SearchCaseSensitive = True
        search.SearchWords = True
        found = document.findFirst(search)
        while found:
            found.String = found.String.replace(find, replace)
            found.setPropertyValue( "CharColor", self.font_color) if self.font_color else ''
            found.setPropertyValue( "CharBackColor", self.font_back_color) if self.font_back_color else ''
            found = document.findNext(found.End, search)

    def set_font_color(self, r, g, b):
        self.font_color = self.__rgbToOOColor(r, g, b)

    def set_font_back_color(self, r, g, b):
        self.font_back_color = self.__rgbToOOColor(r, g, b)

    def __rgbToOOColor(self, r=0, g=0, b=0):
        return (r * 256 * 256 + g * 256 + b)

def absoluteUrl(relativeFile):
    """Constructs absolute path to the current dir in the format required by PyUNO that working with files"""
    mbPrefix = '' if relativeFile[0] == '/' else os.path.realpath('.') + '/'
    return 'file:///' + mbPrefix + relativeFile


if __name__ == "__main__":
    oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
    oo.convert_to('./output.txt', 'txt')
    oo.dispose()
