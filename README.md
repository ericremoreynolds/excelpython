# ExcelPython v2

```python
from xlpython import *

@xlfunc
@xlarg("x", "nparray", 2)
@xlarg("y", "nparray", 2)
def matrixmult(x, y):
    return x.dot(y)
```

![image](https://cloud.githubusercontent.com/assets/5197585/3907706/6c3a2cea-22fd-11e4-812f-41c814d1cc54.png)


### Get started

* Download the [latest release](https://github.com/ericremoreynolds/excelpython/releases)
* Check out the [docs](docs/) for tutorial to help you get going
* The only prerequisites are Excel and Python with PyWin32 installed

### About ExcelPython

ExcelPython is a lightweight, easily distributable library for interfacing Excel and Python. It enables easy access to Python scripts from Excel VBA, allowing you to substitute VBA with Python for complex automation tasks which would be facilitated by Python's extensive standard library.

v2 is a major rewrite of the previous ExcelPython, moving over to a new approach which is more robust and configurable - and importantly should eliminate all those irritating DLL not found problems! The technical details: Python now runs out-of-process and communication happens over COM.

### Help me!

Check out the [docs](docs/) folder for tutorials to help you get started and links to other resources.

If you don't find your answer, need more help, find a bug, think of a useful new feature, or just want to give some feedback by letting me and everyone else know what you're doing with ExcelPython, please create an issue ticket!
