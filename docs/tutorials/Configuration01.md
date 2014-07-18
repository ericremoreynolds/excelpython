### Note: this documentation is out of date!

# Configuring ExcelPython

On of the most important features of ExcelPython is the ability to specify exactly which Python installation you wish to distribute, and to define an isolated Python execution environment with which your workbook interacts.

In tutorials, the default configuration is used, which runs the PC's default Python installation (i.e. the one that appears by running `python.exe` in the Start > Run box).

However when distributing a workbook which has been developed with xlpython, it is better to include a custom `.xlpy` configuration file which ensures that all the workbook's code will execute in a seperate Python process, thus not interfering with other xlpython workbooks which might have been opened by the user in Excel at the same time.

Let's assume we're developing an Excel workbook called `MatrixAlgebra.xlsm` in the folder `...\Desktop\MatrixAlgebra`, and that the Python code is in a file called `MatrixAlgebra.py` in the same folder. Let's suppose we want to distribute this workbook in a zip file, and we want to include a copy of a portable Python distribution in the folder `...\Desktop\MatrixAlgebra\PortablePython` which we want to use to execute our Python code within the context of the workbook.

To create a config file for this setup, copy `xlpython.xlpy` into `...\Desktop\MatrixAlgebra` and rename it `MatrixAlgebra.xlpy`, then edit it as follows:

    CLSID = {66496c37-eb73-4b42-baf6-fad4296e464c}

    Command = $(ConfigDir)\PortablePython\python.exe $(DllDir)\xlpython.py $(CLSID)
    
    WorkingDir = $(ConfigDir)
    
Then, in VBA, edit the xlpython module to use this new config file:

    Function Py()
        If 0 <> GetPythonInterface(Py, ThisWorbook.Path + "\MatrixAlgebra.xlpy") Then Err.Raise 1000, Description:=Py
    End Function
