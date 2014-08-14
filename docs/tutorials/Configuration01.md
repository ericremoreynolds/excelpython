**NB: work in progress**

# Configuring ExcelPython

On of the most important features of ExcelPython is the ability to specify exactly which Python installation you wish to use, and to define an isolated Python execution environment with which your workbook interacts.

In many cases there is no need to make any changes to the default configuration, which runs the PC's default Python installation, i.e. the one that appears by entering `python.exe` in the Start > Run box. In other cases however, you may be interested in targeting a specific copy of Python installed on your PC, or executing Python with specific environment variables.

Let's assume we're developing an Excel workbook called `MatrixAlgebra.xlsm` in the folder `%SOMEFOLDER%\MatrixAlgebra`, and that the Python code is in a file called `MatrixAlgebra.py` in the same folder. Let's suppose we want to distribute this workbook in a zip file, and we want to include a copy of a portable Python distribution in the folder `%SOMEFOLDER%\MatrixAlgebra\PortablePython` which we want to use to execute our Python code within the context of the workbook. The ExcelPython runtime has already been copied to `%SOMEFOLDER%\MatrixAlgebra\xlpython`.

Inside this last folder there is a file called `xlpython.cfg` which determines how the Python process is launched when some functionality within the workbook (e.g. a VBA function or a worksheet formula) tries to interact with Python.

To specify that the Python distribution in `%SOMEFOLDER%\MatrixAlgebra\PortablePython` must be used, make the following modification to `xlpython.cfg`

    Command = $(ConfigDir)\PortablePython\pythonw.exe -u "$(ConfigDir)\xlpyserver.py" $(CLSID)
