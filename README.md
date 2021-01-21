# oodocument

Connects vía Uno bridge interface with libreoffice to perform simple format conversions and replace string over a document.
[Pypi Package link](https://pypi.org/project/oodocument/)

## Dependencies

 If you are **virtualenv** user, you need to create with option **--system-site-packages** as it needs to use system site-packages which has uno module installed by `$ sudo apt-get install python3-uno`. (Unfortunately, uno module is not available from pip).

```
 sudo apt-get install libreoffice-common python3-uno
```

You must run headless libreoffice service (libreoffice executable can be different on different gnu/linux distros, another options can be  `soffice` or `loffice`)

```
/usr/bin/libreoffice --headless --nologo --nofirststartwizard --accept="socket,host=0.0.0.0,port=8001;urp"
```

## Features

- Search and replace (with font colors and background)
- Support file Conversion to pdf, txt, odt and docx

## Install

`pip install oodcument`

## Examples

### Search and Replace

Will open **input.docx** file, search for **holamundo** String and replacing it by **XXX**, then would save the output to **output.pdf** with **pdf** format.

```
from oodocument import oodocument
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with(data, './output.pdf', 'pdf')
oo.dispose()
```

### Search and Replace with Indexes

Will open **input.docx** file with **holamundo** String. The package will found indexes and replace **mundo** String and replacing it by **XXX**, then would save the output to **output.pdf** with **pdf** format.

In this case, you must build a  3-component tuple with the following format:```(start_index,end_index,text_to_replace)``` 

```
from oodocument import oodocument
data = []
data.append((5,10,'XXX'))
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with_index(data, './output.pdf', 'pdf')
oo.dispose()
```

Another Feature is the possibility of adding an offset parameter in ```replace_with_index```

```
from oodocument import oodocument
data = []
data.append((4,9,'XXX'))
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with_index(data, './output.pdf', 'pdf',1)
oo.dispose()
```


Will open **input.docx** file, search for **holamundo** String and replacing it by **XXX** with yellow background and red color font, then would save the output to the same file.

```
from oodocument import oodocument
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.set_font_color(255, 255, 0)
oo.set_font_back_color(255, 0, 0)
oo.replace_with(data)
oo.dispose()
```

Will convert **input.docx** file to **output.txt** file with **txt** format

```
from oodocument import oodocument
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.convert_to('./output.txt', 'txt')
oo.dispose()
```
