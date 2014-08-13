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

### About ExcelPython

ExcelPython is a lightweight, easily distributable library for interfacing Excel and Python. It enables easy access to Python scripts from Excel VBA, allowing you to substitute VBA with Python for complex automation tasks which would be facilitated by Python's extensive standard library, while allowing you to avoid the complexities of Python COM programming.

### Getting started

* Download the [latest release](https://github.com/ericremoreynolds/excelpython/releases)
* Check out the [docs](docs/) for tutorial to help you get going
* The only prerequisites are Excel and Python with PyWin32 installed

### Help me!

Check out the [docs](docs/) folder for tutorials to help you get started and links to other resources. Failing that, try the [issues section](https://github.com/ericremoreynolds/excelpython/issues?q=) or the [discussion forum on SourceForge](https://sourceforge.net/p/excelpython/discussion/general/).

If you still don't find your answer, need more help, find a bug, think of a useful new feature, or just want to give some feedback by letting us know what you're doing with ExcelPython, please go ahead and create an [issue ticket](https://github.com/ericremoreynolds/excelpython/issues/new)!
