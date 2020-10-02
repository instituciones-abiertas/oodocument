# oodocument
> Connects vía Uno bridge interface with libreoffice to perform simple format conversions and replace string over a document.

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

- Search and replace
- Support file Conversion to pdf, txt and docx

## Example

### Search and Replace

Will open **input.docx** file, search for **holamundo** String and replacing it by **XXX**, then would save the output to **output.pdf** with **pdf** format.

```
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with(data, './output.pdf', 'pdf')
oo.dispose()
```

Will open **input.docx** file, search for **holamundo** String and replacing it by **XXX**, then would save the output to the same file.

```
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with(data)
oo.dispose()
```

Will convert **input.docx** file to **output.txt** file with **txt** format

```
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.convert_to('./output.txt', 'txt')
oo.dispose()
```

## TODO
- Set global temoral directory
- Tutorial dockerizing libreoffice ( including binding directories, volume)
- Tests


