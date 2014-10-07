---
layout: page
title: "Quickstart"
---

## Quickstart

### Installation

The easiest way to install ExcelPython is using the Windows installer for the latest release on [GitHub][].

[GitHub]: http://github.com/ericremoreynolds/excelpython/releases

### Write a user-defined function in Python

* Save a your workbook as a macro-enabled workbook (*.xlsm).

* Click **Setup ExcelPython** from the ExcelPython tab in Excel.

* Open `YourWorkbook.py` in the same folder as your workbook, enter the following code

    ```python
    # YourWorkbook.py
    from xlpython import *

    @xlfunc
    def DoubleSum(x, y):
      '''Returns twice the sum of the two arguments'''
      return 2 * (x + y)
    ```
		
* Click **Import Python UDFs**.

* Now you can call the `DoubleSum` function from an Excel formula:

    ```
    =DoubleSum(B1, C1)
    ```
		
### What now?

ExcelPython has a lot of functionality and can be used in many different ways, to find out more check out the [detailed tutorials][] available on the GitHub repo.

[detailed tutorials]: https://github.com/ericremoreynolds/excelpython/tree/master/docs