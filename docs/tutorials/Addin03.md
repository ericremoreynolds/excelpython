## Something a bit more interesting - NumPy arrays

One of the attractions of using Python from Excel is to gain access to the vast range of publicly available Python libraries for numerical computing. Since [NumPy](http://www.numpy.org/) is the cornerstone of many of these libraries, the ExcelPython add-in makes it easy to pass function arguments as and convert return values from [numpy arrays](http://docs.scipy.org/doc/numpy/reference/generated/numpy.array.html).

We will now define a simple function for doing matrix multiplication using NumPy.

* Add the following code to `Book1.py` from the [first tutorial](./Addin01.md)

    ```python
    @xlfunc
    @xlarg("x", "nparray", dims=2)
    @xlarg("y", "nparray", dims=2)
    def MatrixMult(x, y):
        return x.dot(y)
    ```
    
* Click 'Import Python UDFs' to pick up the changes

* The function `MatrixMult` can now be used as an array function from Excel

    ![image](https://cloud.githubusercontent.com/assets/5197585/3918420/9ab74be2-2399-11e4-9b55-8a8005afeabc.png)

    To enter the above array formula in Excel
    * fill in the values in the ranges `D1:E2` and `G1:H2` 
    * select cells `A1:B2`
    * type in the formula `=MatrixMult(D1:E2, G1:H2)`
    * press `Ctrl+Shift+Enter`.
