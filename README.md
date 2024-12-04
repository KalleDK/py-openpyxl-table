# OpenPyxl Table
```python

from openpyxl_table import DictWriter
from openpyxl import Workbook
wb = Workbook()



if isinstance(ws := wb.active, Worksheet):
    with DictWriter(ws, "A1", ['first_name', 'middle', 'last_name'], auto_header=True) as writer:
        writer.writerow({'first_name': 'Baked', 'middle': 'lfds', 'last_name': 'Beans'})
        writer.writerow({'first_name': 'Lovely', 'middle': 'lfds','last_name': 'Spam'})
        writer.writerow({'first_name': 'Wonderful', 'middle': 'lfds','last_name': 'Spam'})


    with DictWriter(ws, "K1", ['first_name', 'middle', 'last_name'], auto_header=False, displayName="Table2") as writer:
        writer.writerow({'first_name': 'First', 'middle': 'Middle', 'last_name': 'Last'})
        writer.writerow({'first_name': 'Baked', 'middle': 'lfds', 'last_name': 'Beans'})
        writer.writerow({'first_name': 'Lovely', 'middle': 'lfds','last_name': 'Spam'})
        writer.writerow({'first_name': 'Wonderful', 'middle': 'lfds','last_name': 'Spam'})


wb.save("test.xlsx")

```
