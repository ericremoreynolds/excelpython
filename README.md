# ExcelPython v2

### Get started

* Download the [latest release](/releases)
* Unzip it into the folder containing your workbook (this will create the `xlpython` folder)
* Import the module `xlpython\xlpython.bas` into your workbook

and you're set! No registration or typelibs. All that is needed is Python with PyWin32 installed.

```vb.net
Sub test1()
    MsgBox Py.Str(Py.Eval("1+2"))
End Sub

Sub test2()
    Set vars = Py.Dict()
    Py.Exec "from datetime import date", vars
    MsgBox Py.Str(Py.Eval("date.today()", vars))
End Sub

Sub test3()
    ' Assumes there's a file called MyScript.py in the same folder as the workbook
    ' with a function called MyFunction taking 3 arguments
    Set m = Py.Module("MyScript")
    MsgBox Py.Str(Py.Call(m, "MyFunction", Py.Tuple(1, 2, 3)))
End Sub
```

Note that the latest release zip also contains an Excel add-in `xlpython.xlam`. We're still in the process of developing that, not to mention writing the documentation - but feel welcome to take a look, and see what it does! Sneak preview: automatic generation of VBA wrappers for Python functions.
 
### About ExcelPython

ExcelPython is a lightweight, easily distributable library for interfacing Excel and Python. It enables easy access to Python scripts from Excel VBA, allowing you to substitute VBA with Python for complex automation tasks which would be facilitated by Python's extensive standard library.

v2 is a major rewrite of the previous ExcelPython, moving over to a new approach which is more robust and configurable - and importantly should eliminate all those irritating DLL not found problems! The technical details: Python now runs out-of-process and communication happens over COM.

### Help me!

Check out the [docs](docs/) folder for tutorials to help you get started and links to other resources.

If you don't find your answer, need more help, find a bug, think of a useful new feature, or just want to give some feedback by letting me and everyone else know what you're doing with ExcelPython, please create an issue ticket!
