xlpython
========

To get started:
* Download the latest release zip
* Unzip it into the folder containing your workbook (this will create the `.xlpy` folder)
* Import the module `.xlpy\xlpython.bas` into your workbook
and you're set! No registration or typelibs. All that is needed is Python with PyWin32 installed.

```
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
