---
layout: page
title: "Examples"
---

## Examples

**Note**: The lite examples now require the latest version of xlwings installed. Also, the only sample that currently works
on Windows and Mac is the Fibonacci Lite sample. The others use ActiveX controls that are not supported on Mac. This
will be changed at some point to make them all cross-platform. Standalone samples work on Windows only.

### Instructions

1. Download the zip-file
2. Windows: Right-Click > Extract All... > Extract / Mac: Double-click the zip-file
3. Open the Spreadsheet in the unzipped folder
4. "Protected View": click on "Enable Editing"
5. Optional: If Excel gives you an additional "Security Warning": click on "Enable Content", then
   please close and reopen the file
6. Run the examples by clicking on the "Run" button


### Downloads

**Example 1: Fibonacci Sequence**

This is the simplest possible example demonstrating the calculation of the Fibonacci sequence.

* **Lite (Win & Mac):** [fibonacci.zip][] (32.5 KB) - Dependencies: Python, pywin32, xlwings
* **Standalone (Win):** [fibonacci_standalone.zip][] (6.3 MB)

[fibonacci.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/fibonacci.zip
[fibonacci_standalone.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/fibonacci_standalone.zip


**Example 2: Database**

This example shows how easy it is to work with databases. It uses [Chinook][], a popular [SQLite][] sample
database.

* **Lite (Win):** [database.zip][] (484.4 KB) - Dependencies: Python, pywin32, xlwings
* **Standalone (Win):** [database_standalone.zip][] (7.1 MB)

[Chinook]: http://chinookdatabase.codeplex.com/
[SQLite]: http://sqlite.org/
[database.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/database.zip
[database_standalone.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/database_standalone.zip

**Example 3: Monte Carlo Simulation**

This example shows the computational power of Python by performing a Monte Carlo simulation of the price development of
a financial asset. Prices are assumed to follow a log-normal distribution.

* **Lite (Win):** [simulation.zip][] (151.1 KB) - Dependencies: Python, pywin32, xlwings, NumPy
* **Standalone (Win):** [simulation_standalone.zip][] (16.8 MB)



[simulation.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/simulation.zip
[simulation_standalone.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/simulation_standalone.zip


### Lite Versions

These versions are small in size but require an installation of Python with xlwings. It is highly recommended to install
one of the following scientific Python distributions as they already contain all of the necessary packages used in the
examples, most importantly pywin32, numpy, scipy and pandas.

* [Anaconda](https://store.continuum.io/cshop/anaconda/)
* [WinPython](http://winpython.sourceforge.net/) (see Notes below)
* [Canopy](https://www.enthought.com/products/canopy/)
* [Python(x,y)](https://code.google.com/p/pythonxy/)


### Standalone Versions

These versions run out-of-the-box after unzipping without any dependencies but are bigger in size.


### Notes

**WinPython**: Since WinPython doesn't change the PATH environment variables, you either have to add the location
  of the Python interpreter to the PATH manually or change the directory in the spreadsheet as follows:

* Press `Alt-F11` to fire up the VBA Editor
* Double-click the `xlwings` module
* In the `RunPython` function, change `PYTHON_DIR = ""` to the directory of where `python.exe` is, e.g.:
`PYTHON_DIR = "C:\WinPython-64bit-2.7.6.3\python-2.7.6"`
