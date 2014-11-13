A very simple usage example
---

You can try out ExcelPython from the VBA Immediate Window (Ctrl + G if it's not already visible in the VBA window).

Type the following

    ?Py.Str(Py.Eval("1+2"))

and press return. You should see 

    ?Py.Str(Py.Eval("1+2"))
    3

(It can happen that you get an `Object required` error - this is probably because when you press return in the Immediate Window, you must have your ExcelPython-enabled workbook selected from the VBA project list. This is because otherwise the Immediate Window statement will execute in the wrong VBA context and therefore will not be able to find the `Py` function.)

You can try evaluating any Python expression, and you'll get the result
printed in the Immediate Window.

    ?Py.Str(Py.Eval("[1,2,3]+[4,5,6]"))
    [1, 2, 3, 4, 5, 6]

Why is the PyStr function needed? If you try just calling

    ?Py.Eval("1+2")

you'll get a type mismatch error. This is because the `Py.Eval` function returns a handle to the Python object itself, and VBA doesn't know how to print it to the Immediate Window. `Py.Str` takes a Python object and calls the Python str function on it and returns a VBA string, which can be printed)

The `Py.Eval` function takes as an optional second argument a locals dictionary to be used when evaluating the expression. To build the dictionary, you can use the function `Py.Dict` which takes alternating keys and values as arguments.

    ?Py.Str(Py.Eval("x+y", Py.Dict("x", 3, "y", 4)))
    7
    ?Py.Str(Py.Eval("x+y", Py.Dict("x", "abc", "y", "def")))
    abcdef
