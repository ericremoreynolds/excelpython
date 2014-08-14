## Loading the ExcelPython add-in

* Download the latest [release](https://github.com/ericremoreynolds/excelpython/releases) and unzip it somewhere.

* Open the add-in `xlpython.xlam` in Excel.

* If all goes well you should see the ExcelPython tab in Excel's toolbar.

    ![image](https://cloud.githubusercontent.com/assets/5197585/3917034/3623f40a-2385-11e4-9754-5e3b924e38a9.png)

Note that it is possible to [permanently install the add-in](#permanently-installing-the-excelpython-add-in) so you don't need to open it manually each time.

## Writing a user-defined function in Python

To interact with Python, a workbook must first be setup to use ExcelPython. To do this it is first necessary to save it as a macro-enabled workbook.

* Choose an empty folder and in it save an empty workbook as `Book1.xlsm`.

* From the ExcelPython tab in the toolbar click 'Setup ExcelPython'.

Next write your user-defined function in Python. In the previous step ExcelPython will have created a file called `Book1.py` in the same folder as `Book1.xlsm` in which the Python functions to be used in the workbook can be defined. 

* Edit `Book1.py` to contain the following code:

    ```python
    # Book1.py
    from xlpython import *
    
    @xlfunc
    def DoubleSum(x, y):
    	'''Returns twice the sum of the two arguments'''
    	return 2 * (x + y)
    ```

* Switch back to Excel and click 'Import Python UDFs' in the ExcelPython tab to pick up the changes made to `Book1.py`.

* Enter the formula `=DoubleSum(1, 2)` into a cell and you should get the correct result:

    ![image](https://cloud.githubusercontent.com/assets/5197585/3917596/e5365b3c-238e-11e4-8bce-0d97caceca2e.png)
    
* Note that the `DoubleSum` function is usable from VBA as well. Open the VBA window (`Alt+F11`), switch to the Immediate Window (`Ctrl+G`) and type

    ```
    ?DoubleSum(1, 2)
    ```
    
## Array arguments

You can pass a range as a function argument, as opposed to a single cell. Its value will be converted to a tuple of tuples.

* Add the following code to `Book1.py` from the previous section

    ```python
    @xlfunc
    def MyUDF(x):
        return repr(x)
    ```
    
    This function simply returns its argument converted to string representation. This will allow us to explore how formula arguments are converted into Python objects.
    
* Click 'Import Python UDFs' to pick up the changes

* Modify the workbook as below

    ![image](https://cloud.githubusercontent.com/assets/5197585/3918899/302a5790-23a0-11e4-80fe-7c75b63c4225.png)

As you can see the value of the 2x2 range `F1:G2` has been convert to a tuple containing tuples representing the range's two rows.

At this point it is worth talking about one of Excel's oddities, namely that the value of a 1x1 range is always a scalar, whereas the value of any range larger than 1x1 is represented by a two-dimensional array.

![image](https://cloud.githubusercontent.com/assets/5197585/3918954/f2a67fce-23a0-11e4-810d-52870204e77f.png)

![image](https://cloud.githubusercontent.com/assets/5197585/3918958/02575a2e-23a1-11e4-8a5f-3f9e21a64abe.png)

ExcelPython provides a mechanism for normalizing the input arguments so that your function can safely make assumptions about their dimensionality.

**work in progress...*

## Something a bit more interesting

One of the attractions of using Python from Excel is to gain access to the vast range of publicly available Python libraries for numerical computing. Since [NumPy](http://www.numpy.org/) is the cornerstone of many of these libraries, the ExcelPython add-in makes it easy to pass function arguments as and convert return values from [numpy arrays](http://docs.scipy.org/doc/numpy/reference/generated/numpy.array.html).

We will now define a simple function for doing matrix multiplication using NumPy.

* Add the following code to `Book1.py` from the previous section

    ```python
    @xlfunc
    @xlarg("x", "nparray", 2)
    @xlarg("y", "nparray", 2)
    def MatrixMult(x, y):
        return x.dot(y)
    ```
    
* Click 'Import Python UDFs' to pick up the changes

* The function `MatrixMult` can now be used as an array function from Excel

    ![image](https://cloud.githubusercontent.com/assets/5197585/3918420/9ab74be2-2399-11e4-9b55-8a8005afeabc.png)

    To enter the above array formula in Excel, fill in the values in the ranges `D1:E2` and `G1:H2` then select cells `A1:B2`, type in the formula `=MatrixMult(D1:E2, G1:H2)` and press `Ctrl+Shift+Enter`.

## Permanently installing the ExcelPython add-in

The best place to place the ExcelPython add-in the Excel startup folder. All Excel files placed in this folder are automatically opened by Excel on startup, so if you place ExcelPython there you do not need to manually open it each time you want to use it. The `xlpython.xlam` file must be placed in the `XLSTART` folder and the `xlpython` folder must be copied alongside it like so:

![image](https://cloud.githubusercontent.com/assets/5197585/3917303/0ef6b35e-238a-11e4-9017-ab8cdb74719d.png)

Unfortunately the folder's location varies depending on the version of Excel. Some candidates are

    Excel 2013 current user: %APPDATA%\Roaming\Microsoft\Excel\XLSTART
    Excel 2010 current user: %APPDATA%\Microsoft\Excel\XLSTART
    Excel 2007 current user: %APPDATA%\Microsoft\Excel\XLSTART
    Excel 2003 current user: %APPDATA%\Microsoft\Excel\XLSTART

    Excel 2013 all users: %PROGRAMFILES%\Microsoft Office 15\root\Office15\XLSTART
    Excel 2010 all users: %PROGRAMFILES%\Microsoft Office\Office14\XLSTART
    Excel 2007 all users: %PROGRAMFILES%\Microsoft Office\Office12\XLSTART
    Excel 2003 all users: %PROGRAMFILES%\Microsoft Office\Office11\XLSTART
    
Note also that `%PROGRAMFILES%` may need to be substituted with `%PROGRAMFILES(x86)%` for 32-bit Excel installed on a 64-bit machine.

If you are in doubt as to where the folder is located, you can also determine it by opening Excel, opening the VBA window (`Alt+F11`), switching to the Immediate Window (`Ctrl+G`) and typing

    ?Application.StartupPath

for the current user and

    ?Application.Path + "\XLSTART"
    
for all users.
