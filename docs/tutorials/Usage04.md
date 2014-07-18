Ranges, lists and SAFEARRAYs
---

If you pass a VBA object to an ExcelPython function that expects a Python object, it will get converted automatically.

That's why you can do the following

```
?Py.Str(Range("A1:C1").Value2)
((3.0, 1.0, 2.0),)
```
   
The function `Py.Str` applies the Python `str` function to whatever object you pass it. If that object is not already a Python object, ExcelPython will convert it for you.

In this specific example, the object that gets passed is a VBA array:

```
?TypeName(Range("A1:C1").Value2)
Variant()
```

Specifically, it's a 1x3 two-dimensional array:

```
?UBound(Range("A1:C1").Value2,1) & " to " & LBound(Range("A1:C1").Value2, 1)
1 to 1
?UBound(Range("A1:C1").Value2,2) & " to " & LBound(Range("A1:C1").Value2, 2)
3 to 1
```
  
This is an 'oddity' of Excel's VBA model - the values of 1x1 ranges are scalars:

```
?TypeName(Range("A1").Value2)
Double
```
  
and all other size ranges have two-dimensional array values.

Since two-dimensional tuples do not exist in Python, it gets converted to a tuple-of-tuples. Often however, you will want to pass a range as a one-dimensional tuple. Consider for example the built-in Python function `sorted`. If you call this function on the range "A1:C1" it returns the same value as is passed in:

```
?Py.Str(Py.Call(Py.Builtins, "sorted", Py.Tuple(Range("A1:C1").Value2)))
[(3.0, 1.0, 2.0)]
```

This is because `sorted` only sorts the top-level sequence, which contains only one element, the tuple `(3.0, 1.0, 2.0)`.

The VBA code is equivalent to the following Python code:

```python
sorted(((3.0, 1.0, 2.0),))
```

but we want the equivalent to:

```python
sorted((3.0, 1.0, 2.0))
```

To do this, you can convert the 1x3 two-dimensional VBA array to 3-element one-dimensional array before passing it to the python function Python:

```
?Py.Str(NDims(Range("A1:C1").Value2, 1))
(3.0, 1.0, 2.0)
?Py.Str(Py.Call(Py.Builtins, "sorted", Py.Tuple(NDims(Range("A1:C1").Value2, 1))))
[1.0, 2.0, 3.0]
```

NumPy arrays
--

Conversion to/from NumPy types is a frequent necessity, and ExcelPython does not do this automatically.

For example, the function `scipy.stats.norm.cdf` returns a `numpy.float64` object:

```
Set vars = Py.Dict()
Py.Exec "from scipy.stats import norm", vars
?Py.Str(Py.Eval("norm.cdf(0.0)", vars))
0.5
?Py.Str(Py.Eval("type(norm.cdf(0.0))", vars))
<type 'numpy.float64'>
?Py.Var(Py.Eval("norm.cdf(0.0)", vars))
```

The last expression causes an error, because the `Py.Var` function (or, more generally PyWin32) doesn't know how to convert a `numpy.float64` object to a native VBA type.

So how do you get the value out of Python and into VBA? The key is to convert it to a type which `Py.Var` _does_ know how to convert, such as `float`. Note that this can be done either within your Python code:

```
?Py.Var(Py.Eval("float(norm.cdf(0.0))", vars))
 0.5 
```

or in the VBA code by calling the Python conversion function directly on the returned object:

```
?Py.Var(Py.Call(Py.Builtins, "float", Py.Tuple(Py.Eval("norm.cdf(0.0)", locals))))
 0.5
```

This second method, while a bit more convoluted, means your Python code does not have to be written to take into account that it could be called from Excel.

The same applies to passing in/out Numpy arrays. Suppose you have a function that takes NumPy arrays as inputs and returns them as outputs:

```python
# MatrixFunctions.py
import numpy as np

def xl_mmult(x, y):
    return x.dot(y)
```

This can be called from VBA code (without modifying the original Python code) like so:

```vb.net
Public Function PyXLMult(x As Range, y As Range)

On Error GoTo fail:

    ' Python equivalent: import numpy; numpyArray = numpy.array
    Set numpyArray = Py.GetAttr(Py.Module("numpy"), "array")

    ' Python equivalent: x_array = numpyArray(x)
    Set x_array = Py.Call(numpyArray, Py.Tuple(x.Value2))

    ' Python equivalent: y_array = numpyArray(y)
    Set y_array = Py.Call(numpyArray, Py.Tuple(y.Value2))

    ' Python equivalent: result_array = xl_mmult(x_array, y_array)
    Set result_array = Py.Call(Py.Module("MatrixFunctions"), "xl_mmult", Py.Tuple(x_array, y_array))

    ' Python equivalent: result_list = result_array.tolist()
    Set result_list = Py.Call(result_array, "tolist")

    PyXLMult = Py.Var(result_list)
    Exit Function

fail:
    PyXLMult = Err.Description

End Function
```
