## Programmatic access to Visual Basic Project is not trusted

![image](https://cloud.githubusercontent.com/assets/5197585/4011749/4c09f4ee-2a07-11e4-8882-c7f6c12b4096.png)

This appears because your Excel trust settings do not allow an add-in to manipulate another workbook's VBA code, which the ExcelPython add-in needs to be able to do to perform its tasks.

To change this trust setting, select File > Options > Trust Center > Trust Center Settings > Macro Settings and ensure the checkbox labeled 'Trust access to the VBA project object model' is checked.

![image](https://cloud.githubusercontent.com/assets/5197585/3921677/77648a04-23c3-11e4-9af4-2a14ca47787e.png)

## Could not create Python process - the system cannot find the file specified

![image](https://cloud.githubusercontent.com/assets/5197585/4011849/6e2e882c-2a08-11e4-88e2-93abc620d4e2.png)

This is probably because you don't have Python installed on your machine. If you are sure you do, then it may not be on the system path, or if you have just installed it maybe Excel just needs to be restarted.

## Python process exited before it was possible to create the interface object

![image](https://cloud.githubusercontent.com/assets/5197585/4011823/2abd178e-2a08-11e4-941e-61d4a669eb7f.png)

This means that there was an error in running the ExcelPython server script (`xlpyserver.py`) and it can happen for many reasons.

The most likely cause is that you haven't installed the [Python for Windows extensions (PyWin32)](http://sourceforge.net/projects/pywin32/) library into your Python distribution. Failing this you should consult the specified log file, which will show you the output of the Python interpreter including any errors that might have been raised during the script's execution.
