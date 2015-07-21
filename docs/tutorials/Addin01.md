## Loading the ExcelPython add-in

* Download the latest [release](https://github.com/ericremoreynolds/excelpython/releases) and unzip it somewhere.

* Open the add-in `xlpython.xlam` in Excel.

* If all goes well you should see the ExcelPython tab in Excel's toolbar.

    ![image](https://cloud.githubusercontent.com/assets/5197585/3917034/3623f40a-2385-11e4-9754-5e3b924e38a9.png)
    
* You may get an error saying `Programmatic access to Visual Basic Project is not trusted`. If so check out the [add-in troubleshooting guide](./AddinTrouble.md).

Note that it is possible to [permanently install the add-in](./Addin05.md) so you don't need to open it manually each time.

## Writing a user-defined function in Python

To interact with Python, a workbook must first be setup to use ExcelPython. To do this it is first necessary to save it as a macro-enabled workbook.

* Choose an empty folder and in it save an empty workbook as `Book1.xlsm`.

* From the ExcelPython tab in the toolbar click 'Setup ExcelPython'.

Next write your user-defined function in Python. In the previous step ExcelPython will have created a file called `Book1.py` in the same folder as `Book1.xlsm` in which the Python functions to be used in the workbook must be defined. 

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

* Enter the formula `=DoubleSum(1, 2)` or `=DoubleSum(1; 2)` depending on your Excel version's locale into a cell and you should get the correct result:

    ![image](https://cloud.githubusercontent.com/assets/5197585/3917596/e5365b3c-238e-11e4-8bce-0d97caceca2e.png)
    
* Note that the `DoubleSum` function is usable from VBA as well. Open the VBA window (`Alt+F11`), switch to the Immediate Window (`Ctrl+G`) and type

    ```
    ?DoubleSum(1, 2)
    ```

    or
    
    ```
    ?DoubleSum(1; 2)
    ```


To continue move onto the [next tutorial](./Addin02.md).
