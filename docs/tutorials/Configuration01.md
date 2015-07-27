**NB: work in progress**

# Configuring ExcelPython

On of the most important features of ExcelPython is the ability to specify exactly which Python installation you wish to use, and to define an isolated Python execution environment with which your workbook interacts.

In many cases there is no need to make any changes to the default configuration, which runs the PC's default Python installation, i.e. the one that appears by entering `python.exe` in the Start > Run box. In other cases however, you may be interested in targeting a specific copy of Python installed on your PC, or executing Python with specific environment variables.

Let's assume we're developing an Excel workbook called `MatrixAlgebra.xlsm` in the folder `%SOMEFOLDER%\MatrixAlgebra`, and that the Python code is in a file called `MatrixAlgebra.py` in the same folder. Let's suppose we want to distribute this workbook in a zip file, and we want to include a copy of a portable Python distribution in the folder `%SOMEFOLDER%\MatrixAlgebra\PortablePython` which we want to use to execute our Python code within the context of the workbook. The ExcelPython runtime has already been copied to `%SOMEFOLDER%\MatrixAlgebra\xlpython`.

Inside this last folder there is a file called `xlpython.cfg` which determines how the Python process is launched when some functionality within the workbook (e.g. a VBA function or a worksheet formula) tries to interact with Python.

To specify that the Python distribution in `%SOMEFOLDER%\MatrixAlgebra\PortablePython` must be used, make the following modification to `xlpython.cfg`

    Command = $(WorkbookDir)\PortablePython\pythonw.exe -u "$(ConfigDir)\xlpyserver.py" $(CLSID)

## Using the registered default Python X.Y installation

Macros are available providing the folder of each Python version installed and registered as the default installation for that version, as stored in the Windows registry.

These take the form `$(Registry:PythonXYFolder)` where `X` is the major version and `Y` is the minor version (for example for Python 2.6 `$(Registry:Python26Folder)`). The installation path is read from `HKCU\Software\Python\PythonCore\X.Y\InstallPath`.

Minor version agnostic macros are also available: `$(Registry:Python2Folder)` and `$(Registry:Python3Folder)`. They are equal to the folder of the installation with the greatest minor version. For example if Python 2.6 and Python 2.7 are both installed and registered, `$(Registry:Python2Folder)` is equal to the installation folder of the default Python 2.7.

The obvious use case for these macros is the `Command` setting in `xlpython.cfg`. For example, to instruct ExcelPython to use Python 2 even if Python 3 takes precedence according to the system path, one can modify the `Command` setting as follows:

    # The command line used to launch the COM server
    Command = "$(Registry:Python2Folder)\pythonw.exe" -u "$(ConfigDir)\xlpyserver.py" $(CLSID)

