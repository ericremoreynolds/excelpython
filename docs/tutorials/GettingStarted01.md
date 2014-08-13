**nb: work in progress**

# ExcelPython add-in quickstart

Download the latest [release](https://github.com/ericremoreynolds/excelpython/releases) and unzip it somewhere. Load the add-in by double-clicking `xlpython.xlam` or by opening it from Excel. If all goes well you should see the ExcelPython tab in Excel's toolbar.

Note that it is possible to permanently install the add-in so you don't need to open it each time, see below.

## Setting up a workbook to interact with Python

To interact with Python a workbook must first be enabled to use ExcelPython. To do this it is first necessary to save it as a macro-enabled workbook.
1. Choose an empty folder and in it save an empty workbook as `Book1.xlsm`
2. From the ExcelPython tab in the toolbar click 'Setup ExcelPython'

## Permanently installing the ExcelPython add-in

The best place is Excel's `XLSTART` folder which is located in Office's installation folder in Program Files. If you are 
