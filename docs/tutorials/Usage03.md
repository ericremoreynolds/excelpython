Putting it all together
---

Let's suppose we have the following script MyScript.py saved in the same folder as the workbook.

```python
def MyFunction(x, y):
  return {
    "sum": x + y,
    "sorted": sorted([x, y])
  }
```

The following VBA code can be used to call this function, taking the input parameters from a worksheet and pasting the outputs back to that sheet.

    Function PyPath()
        PyPath = ThisWorkbook.Path
    End Function
    
    Sub CallPythonCode()
        Set res = Py.Call( _
            Py.Module("MyScript", Py.AddPath(PyPath)), "MyFunction", _
            Py.Dict( _
                "x", Sheet1.Range("A1").Value2, _
                "y", Sheet1.Range("A2").Value2))
        Sheet1.Range("A3").Value2 = Py.Var(Py.GetItem(res, "sum"))
        Sheet1.Range("A4:B4").Value2 = Py.Var(Py.GetItem(res, "sorted"))
    End Sub

The function `PyPath` is a utility function which returns our additional module search path. It's set up so that to access Python scripts we just need to place them in the same folder as the workbook. It's worth defining this once in your workbook's VBA project for use wherever necessary.

The calls to `Py.Module` and `Py.Call` have been explained above. Once the result dictionary is obtained from the function, its elements can be accessed using `Py.GetItem`.

Finally, we use `Py.Var` to convert the Python objects into variants so that they can be pasted into worksheet ranges.
