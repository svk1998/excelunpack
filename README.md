
# ExcelUnpack

![Build and Deploy Status](https://github.com/svk1998/excelunpack/actions/workflows/python-publish.yml/badge.svg)
[![PyPI Version](https://img.shields.io/pypi/v/excel-unpack)](https://pypi.org/project/excel-unpack/)
![PyPI Downloads](https://img.shields.io/pypi/dm/excel-unpack)

ExcelUnpack is a Python library and command-line tool to extract and organize attachments from Excel files.

## Installation

You can install ExcelUnpack using pip:

```
pip install excel-unpack
```

## Usage

### As a Command-Line Tool:

```
excel-unpack path/to/excel/file.xlsx path/to/output/folder
```

### As a Python Library:

```python
from excelunpack.excel_attachments import extract_attachments_from_xlsx

xlsx_file = "path/to/excel/file.xlsx"
output_dir = "path/to/output/folder"
extract_attachments_from_xlsx(xlsx_file, output_dir)
```
