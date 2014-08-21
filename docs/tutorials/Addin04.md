# Writing macros in Python

In addition to writing user-defined functions, VBA is typically used for defining macros to automate Excel.

* Add the following code to `Book1.py` from the previous tutorial.

    ```python
    @xlsub
    @xlarg("app", vba="Application")
    def my_macro(app):
      from datetime import datetime
      app.StatusBar = str(datetime.now())
    ```

* Click 'Import Python UDFs' to pick up the changes.

* From the Developer tab, select Insert > Form Controls > Button and create the button on Sheet2. You should be prompted to select the macro you want to run when the button is clicked, and should be able to select `my_macro`.

    Note that the Developer tab is usually not visible by default. You can enable it from the Excel Options panel.
    
* If all has gone according to plan, clicking the button will update Excel's status bar (at the bottom of the window) with the current date and time:

    ![image](https://cloud.githubusercontent.com/assets/5197585/3968943/17015b82-27bb-11e4-9ba9-b6b6026dc5d4.png)

To understand the above code, let's go through it line by line.

* `@xlsub` indicates that `my_macro` should be wrapped as a VBA subroutine rather than as a VBA function
* `@xlarg("app", vba="Application")` indicates that the `app` argument should be hard-coded to the VBA expression specified, in this case the `Application` object. Using the `vba` keyword results in the argument being removed from the argument list in Excel. As such our macro does not have any arguments in Excel - indeed, buttons can only be associated with macros with zero parameters.
* `app.StatusBar = str(datetime.now())`. Once inside the function, the `app` variable has been set to the Excel `Application` object, and in this line we set the status bar message to the current date and time.

Another example of how to use this functionality is to manipulate a worksheet. The following macro copies the contents of cell A1 into cell A2 in Sheet2.

```python
@xlsub
@xlarg("sheet", vba="Sheet2")
def my_other_macro(sheet):
    sheet.Range("A2").Value = sheet.Range("A1").Value
```

To continue move onto the [next tutorial](./Addin05.md).
