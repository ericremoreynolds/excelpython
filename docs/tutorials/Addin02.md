## Array arguments

You can pass a range as a function argument, as opposed to a single cell. Its value will be converted to a tuple of tuples.

* Add the following code to `Book1.py` from the [previous tutorial](./Addin01.md)

    ```python
    @xlfunc
    def MyUDF(x):
        return repr(x)
    ```
    
    This function simply returns its argument converted to string representation. This will allow us to explore how formula arguments are converted into Python objects.
    
* Click 'Import Python UDFs' to pick up the changes

* Modify the workbook as below

    ![image](https://cloud.githubusercontent.com/assets/5197585/3918899/302a5790-23a0-11e4-80fe-7c75b63c4225.png)

As you can see the value of the 2x2 range `F1:G2` has been convert to a tuple containing tuples representing the range's two rows.

At this point it is worth talking about one of Excel's oddities, namely that the value of a 1x1 range is always a scalar, whereas the value of any range larger than 1x1 is represented by a two-dimensional array.

![image](https://cloud.githubusercontent.com/assets/5197585/3918954/f2a67fce-23a0-11e4-810d-52870204e77f.png)

![image](https://cloud.githubusercontent.com/assets/5197585/3918991/7e2187ec-23a1-11e4-8fd8-6405c3bbc7b7.png)

ExcelPython provides a mechanism for normalizing the input arguments so that your function can safely make assumptions about their dimensionality.

* Modify `Book1.py` as follows

    ```python
    @xlfunc
    @xlarg("x", dims=2)         # add this line
    def MyUDF(x):
    	return str(x)
    ```
    
* Click 'Import Python UDFs' to pick up the changes.

* Now 1x1 ranges are passed as two-dimensional

    ![image](https://cloud.githubusercontent.com/assets/5197585/3919574/8c257714-23aa-11e4-82d5-97da8b5a5fb2.png)

At other times it you may want to assume that an argument that is one-dimensional

* Modify `Book1.py` as follows

    ```python
    @xlfunc
    @xlarg("x", dims=1)         # modify this line
    def MyUDF(x):
    	return str(x)
    ```
    
* Click 'Import Python UDFs' to pick up the changes.

    ![image](https://cloud.githubusercontent.com/assets/5197585/3919614/54cedaa2-23ab-11e4-8ba9-56dcd86815ad.png)

    ![image](https://cloud.githubusercontent.com/assets/5197585/3919622/6b5ed790-23ab-11e4-9dce-52b45bb72717.png)

    ![image](https://cloud.githubusercontent.com/assets/5197585/3919656/00f96f9a-23ac-11e4-8d7c-c1ae1896002e.png)

* Clearly having specified the argument as one-dimensional, an error is raised if a two-dimensional range is passed

    ![image](https://cloud.githubusercontent.com/assets/5197585/3919669/379cdd66-23ac-11e4-8e47-dfe333143a45.png)
