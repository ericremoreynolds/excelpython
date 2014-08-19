## Permanently installing the ExcelPython add-in

The best place to put the ExcelPython add-in is in the Excel startup folder. All files placed in this folder are automatically opened by Excel on startup, so if you place ExcelPython there you do not need to manually open it each time you want to use it. The `xlpython.xlam` file must be placed in the `XLSTART` folder and the `xlpython` folder must be copied alongside it like so:

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
