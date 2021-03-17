# IA² | oodocument

<p align="center">
  <a target="_blank" rel="noopener noreferrer">
    <img width="220px" src="https://www.ia2.coop/public/ia2-logo-blue.png" alt="IA²" />
  </a>
</p>

<h4 align="center">Document conversions for IA²</h4>

---

<p align="center" style="margin-top: 14px;">
  <a href="https://badge.fury.io/py/oodocument">
    <img src="https://badge.fury.io/py/oodocument.svg" alt="PyPI version" height="20">
  </a>
  <a href="https://pypi.org/project/oodocument/">
    <img src="https://img.shields.io/pypi/dm/oodocument.svg" alt="PyPI version" height="20">
  </a>
  <a
    href="https://github.com/instituciones-abiertas/oodocument/blob/main/LICENSE"
  >
    <img
      src="https://img.shields.io/badge/License-GPL%20v3-blue.svg"
      alt="License" height="20"
    >
  </a>
  <a
    href="https://github.com/instituciones-abiertas/oodocument/blob/main/CODE_OF_CONDUCT.md"
  >
    <img
      src="https://img.shields.io/badge/Contributor%20Covenant-v2.0%20adopted-ff69b4.svg"
      alt="Contributor Covenant" height="20"
    >
  </a>
</p>

## About

The `oodocument` tool connects vía the `Uno bridge interface` with `libreoffice` to perform simple format conversions and string replacement over different kind of documents.

## Pre-requisites and Dependencies

If you are using `virtualenv`, you need to create your environment using the `--system-site-packages` option, since we need to use system site-packages where the `uno` module is installed.

Install the uno module and the libreoffice utilities. Unfortunately, this module is not available from `pip`.

```bash
sudo apt-get install libreoffice-common python3-uno
```

You must run the headless libreoffice service. Note the `libreoffice` executable location can be different depending on the gnu/linux distro you use, another options can be `soffice` or `loffice`.

```
/usr/bin/libreoffice --headless --nologo --nofirststartwizard --accept="socket,host=0.0.0.0,port=8001;urp"
```

## Features

- Search and replace.
- Support font colors, background and clear hyperlinks.
- Support file Conversion to `.pdf`, `.txt`, `.odt` and `.docx`.

## Install package via pip

```bash
pip install oodocument
```

## Run Test

```
make test
```

## Examples

### Search and Replace

The next example will open an **input.docx** file, search for an **holamundo** string and replace it by **XXX**. Then the output is saved to **output.pdf** with **pdf** format.

```python
from oodocument import oodocument
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with(data, './output.pdf', 'pdf')
oo.dispose()
```

### Search and Replace with Indexes

This example opens an **input.docx** file with an **holamundo** string. The package will find indexes and replace the **mundo** substring with **XXX**. Then the output is saved to **output.pdf** with **pdf** format.

In this case, you must build a 3 components tuple with the following format: `(start_index,end_index,text_to_replace)
`

```python
from oodocument import oodocument
data = []
data.append((5,10,'XXX'))
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with_index(data, './output.pdf', 'pdf')
oo.dispose()
```

Another Feature is the possibility of adding an offset parameter in the `replace_with_index` function.

```python
from oodocument import oodocument
data = []
data.append((4,9,'XXX'))
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.replace_with_index(data, './output.pdf', 'pdf',1)
oo.dispose()
```

### Search and Replace with Indexes in Header

This example will open an **input.docx** file with an **holamundo** string in the document header and replace the **mundo** value in the header with the value **XXX**. The name of the style may vary, in that case the style name is `"Default Style"`.

```python
from oodocument import oodocument
data = []
header_style_name = "Default Style"
neighbor_character = 20
data.append((4, 10, "XXX", "mundo"))
oo = oodocument("./input.docx", host="0.0.0.0", port=8001)
oo.replace_with_index_in_header(data, "./output.pdf", "pdf", 0, neighbor_character, header_style_name)
oo.dispose()
```

The next example Will open an **input.docx** file, search for an **holamundo** string and replace it's value by **XXX** with a yellow background and red color font. Then the output is saved to the same file.

```python
from oodocument import oodocument
data = {}
data['holamundo'] = 'XXX'
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.set_font_color(255, 255, 0)
oo.set_font_back_color(255, 0, 0)
oo.set_clear_hyperlinks(False) # default value is True
oo.replace_with(data)
oo.dispose()
```

This one converts an **input.docx** file to **output.txt** file with a **txt** format.

```python
from oodocument import oodocument
oo = oodocument('./input.docx', host='0.0.0.0', port=8001)
oo.convert_to('./output.txt', 'txt')
oo.dispose()
```

## License

[**GNU General Public License version 3**](https://opensource.org/licenses/GPL-3.0)
