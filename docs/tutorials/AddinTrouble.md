## Programmatic access to Visual Basic Project is not trusted

This appears because your Excel trust settings do not allow an add-in to manipulate another workbook's VBA code, which the ExcelPython add-in needs to be able to do to perform its tasks.

To change this trust setting, select File > Options > Trust Center > Trust Center Settings > Macro Settings and ensure the checkbox labeled 'Trust access to the VBA project object model' is checked.

![image](https://cloud.githubusercontent.com/assets/5197585/3921677/77648a04-23c3-11e4-9af4-2a14ca47787e.png)
