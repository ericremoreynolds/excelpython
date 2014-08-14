**nb: work in progress**

## Loading the ExcelPython add-in

1. Download the latest [release](https://github.com/ericremoreynolds/excelpython/releases) and unzip it somewhere.
1. Load the add-in by double-clicking `xlpython.xlam` or by opening it from Excel.
1. If all goes well you should see the ExcelPython tab in Excel's toolbar.
![image](https://cloud.githubusercontent.com/assets/5197585/3917034/3623f40a-2385-11e4-9754-5e3b924e38a9.png)

Note that it is possible to [permanently install the add-in](#permanently-installing-the-excelpython-add-in) so you don't need to open it manually each time.

## Setting up a workbook to interact with Python

To interact with Python a workbook must first be enabled to use ExcelPython. To do this it is first necessary to save it as a macro-enabled workbook.
1. Choose an empty folder and in it save an empty workbook as `Book1.xlsm`
2. From the ExcelPython tab in the toolbar click 'Setup ExcelPython'

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

If you are in doubt as to where the folder is located, you can also determine it by opening Excel, opening the VBA window (Alt+F11), switching to the Immediate Window (Ctrl+G) and typing

    ?Application.StartupPath

for the current user and

    ?Application.Path + "\XLSTART"
    
for all users.
