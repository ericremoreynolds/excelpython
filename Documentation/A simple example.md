A very simple usage example
---

First of all add a reference to the ExcelPython type library. To do this, open Excel, go to the VBA window (Alt-F11) and select Tools > References. Scroll down and select ExcelPythonXX where XX is 26 or 27 depending on which version of Python you have installed on your machine (note only Python 2.6 and 2.7 are supported currently). Click OK to get back to VBA.

You can try out ExcelPython from the VBA Immediate Window (Ctrl + G if it's not already visible in the VBA window).

Type the following

    ?PyStr(PyEval("1+2"))

and press return. You should see 

    ?PyStr(PyEval("1+2"))
    3

You can try evaluating any Python expression, and you'll get the result
printed in the Immediate Window.

    ?PyStr(PyEval("[1,2,3]+[4,5,6]"))
    [1, 2, 3, 4, 5, 6]

(As a side note: why is the PyStr function needed? If you try just calling

    ?PyEval("1+2")

you'll get a type mismatch error. This is because the PyEval function
returns a handle to the Python object itself, and VBA doesn't know how
to print it to the Immediate Window. PyStr takes a Python object and
calls the Python str function on it and returns a VBA string, which can
be printed)

The PyEval function takes as an optional second argument a locals dictionary to be used when evaluating the expression. To build the dictionary, you can use the function PyDict which takes alternating keys and values as arguments.

    ?PyStr(PyEval("x+y", PyDict("x", 3, "y", 4)))
    7
    ?PyStr(PyEval("x+y", PyDict("x", "abc", "y", "def")))
    abcdef
