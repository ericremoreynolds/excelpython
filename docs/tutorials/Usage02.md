A more practical use of ExcelPython
---

The above example gives a light introduction to manipulating a few simple objects through VBA. In practice though, what's needed is a way to get a load of inputs from Excel, pass them to a method defined in a Python script somewhere, get the outputs back from Python and use them VBA or place them in the spreadsheet as required.

The key to doing this are the methods `Py.Module` and `Py.Call`.

PyModule returns a pointer to a Python module, much like the import statement. If no additional arguments are specified, the embedded Python interpreter will look for the requested module in the default search path.

    ?Py.Str(Py.Module("datetime"))
    <module 'datetime' (built-in)\>

By default the Python instance associated with a workbook runs in that workbook's folder (as determined by `ThisWorkbook.Path`), so if you place a Python script in that folder, you can load it as a module.

    ?Py.Str(Py.Module("MyScript"))
    <module 'MyScript' from 'C:\WorkbookFolder\MyScript.py'\>

Once you have access to the module you want, you can use `Py.Call` to invoke functions contained in the module (note that `Py.Call` actually calls any method of any object, not just module objects). This is done by explicitly passing the ordered and keyword arguments. For example calling

    ?Py.Str(Py.Call(Py.Module("datetime"), "date", Py.Tuple(2013, 8, 9)))
    2013-08-09

is equivalent to calling `datetime.date(2013, 8, 9)`, whereas

    ?Py.Str(Py.Call(Py.Module("datetime"), "timedelta", Py.Tuple(3), Py.Dict("milliseconds", 500)))
    3 days, 0:00:00.500000

is like calling `datetime.timedelta(3, milliseconds=500)`.
