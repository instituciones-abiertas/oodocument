import os
import uno
import traceback
import logging
from com.sun.star.beans import PropertyValue

MAX_MOVES = 32767


class oodocument:
    def __init__(self, orig, host="0.0.0.0", port="8001"):
        self.orig = orig
        self.host = host
        self.port = port
        self.font_color = None
        self.font_back_color = None
        self.__set_document()

    def __str__(self):
        return f"{self.orig}"

    def __set_document(self):
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()
        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        try:
            # connect to the running office
            ctx = resolver.resolve(f"uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext")
        except Exception as e:
            logging.error(traceback.format_exc())
            raise

        # get the central desktop object
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self.document = desktop.loadComponentFromURL(absoluteUrl(self.orig), "_blank", 0, ())

    def convert_to(self, dest, format):
        # Filters https://git.libreoffice.org/core/+/refs/heads/master/filter/source/config/fragments/filters
        formats = {
            "pdf": "writer_pdf_Export",
            "txt": "Text",
            "docx": "MS Word 2007 XML",
            "odt": "writer8",
            "doc": "MS Word 97",
        }
        filter = PropertyValue()
        filter.Name = "FilterName"
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

    def replace_with_index(self, data=[], dest=None, format=None, offset=0, word_neighbors=0):
        data.sort(key=lambda x: x[0], reverse=True)
        for start_index, end_index, replace, word_check in data:
            self.__find_and_replace_index(
                self.document, start_index, end_index, replace, word_check, offset, word_neighbors
            )
        if dest is None and format is None:
            self.save()
        else:
            self.convert_to(dest, format)

    def replace_with_index_in_header(
        self,
        data=[],
        dest=None,
        format=None,
        offset=0,
        word_neighbors=0,
        style_name="Default Style",
    ):
        data.sort(key=lambda x: x[0], reverse=True)
        try:
            paragraph = self.document.getStyleFamilies().getByName("PageStyles").getByName(style_name).HeaderText
        except Exception as e:
            logging.error(traceback.format_exc())
            raise

        for start_index, end_index, replace, word_check in data:
            self.__find_and_replace_index(
                self.document, start_index, end_index, replace, word_check, offset, word_neighbors, paragraph
            )
        if dest is None and format is None:
            self.save()
        else:
            self.convert_to(dest, format)

    def save(self, dest=None, filter=None):
        if dest is None and filter is None:
            self.document.store()
        else:
            self.document.storeToURL(absoluteUrl(dest), (filter,))

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
            found.setPropertyValue("CharColor", self.font_color) if self.font_color else ""
            found.setPropertyValue("CharBackColor", self.font_back_color) if self.font_back_color else ""
            found = document.findNext(found.End, search)

    def safe_goRight(self, cursor, count, expand):
        if count > MAX_MOVES:
            jumps = count // MAX_MOVES
            for i in range(jumps):
                cursor.goRight(MAX_MOVES, False)
            padding = count - (MAX_MOVES * jumps)
            cursor.goRight(padding, expand)
        else:
            cursor.goRight(count, expand)

    def __find_and_replace_index(
        self,
        document,
        start_index=None,
        end_index=None,
        replace=None,
        word_check=None,
        offset=0,
        character_neighbors=20,
        paragraph_to_replace=None,
    ):
        """
        Searches the given word and replaces it's value with the given
        replacement string.
        """
        if paragraph_to_replace:
            text = paragraph_to_replace.Text
        else:
            text = document.Text

        cursor = text.createTextCursor()
        total_lenght = (end_index + offset) - ((start_index) + offset) - 1

        self.safe_goRight(cursor, start_index + offset, False)
        cursor.goRight(total_lenght, True)
        character_word = 0

        if word_check != cursor.String:
            character_word = 1
            while word_check != cursor.String and character_neighbors > character_word:
                cursor.gotoStart(False)
                self.safe_goRight(cursor, start_index + offset - character_word, False)
                cursor.goRight(total_lenght, True)
                character_word += 1

        if word_check != cursor.String:
            character_word = 1
            while word_check != cursor.String and character_neighbors > character_word:
                cursor.gotoStart(False)
                self.safe_goRight(cursor, start_index + offset + character_word, False)
                cursor.goRight(total_lenght, True)
                character_word += 1

        if word_check == cursor.String:
            cursor.String = replace
            cursor.setPropertyValue("CharColor", self.font_color) if self.font_color else ""
            cursor.setPropertyValue("CharBackColor", self.font_back_color) if self.font_back_color else ""
            cursor.gotoStart(False)

    def set_font_color(self, r, g, b):
        self.font_color = self.__rgbToOOColor(r, g, b)

    def set_font_back_color(self, r, g, b):
        self.font_back_color = self.__rgbToOOColor(r, g, b)

    def __rgbToOOColor(self, r=0, g=0, b=0):
        return r * 256 * 256 + g * 256 + b

    def __cursor_go_x_next_words(self, count_words, cursor):
        for x in range(count_words):
            cursor.gotoNextWord(True)


def absoluteUrl(relativeFile):
    """Constructs absolute path to the current dir in the format required by PyUNO that working with files"""
    mbPrefix = "" if relativeFile[0] == "/" else os.path.realpath(".") + "/"
    return "file:///" + mbPrefix + relativeFile


if __name__ == "__main__":
    oo = oodocument("./input.docx", host="0.0.0.0", port=8001)
    oo.convert_to("./output.txt", "txt")
    oo.dispose()
