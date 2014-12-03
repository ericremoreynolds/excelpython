## Debugging code with ExcelPython

ExcelPython does not have any special provision for debugging Python code because it is fully compatible with existing Python debugging tools. Here we will see how to set up a couple of the most widely-used ones, [Python Tools for Visual Studio](#debugging-with-python-tools-for-visual-studio) and [PyDev](#debugging-with-eclipse-and-pydev).

## Debugging with Python Tools for Visual Studio

On Windows it is possible to debug Python scripts in Microsoft Visual Studio using an add-in called [Python Tools for Visual Studio](http://pytools.codeplex.com/) (PTVS). To get set up (full instructions can be found on the PTVS website) you need to

1. Install a recent version of Visual Studio, either Express or a commercial license.
2. Download and install the relevant distribution of the PTVS add-in
3. Install the PTVS debug package into your Python distribution
4. Prepare your Python script to enable debugging in Visual Studio.

Assuming you've done steps 1 and 2, step 3 can be achieved using [PIP](https://pypi.python.org/pypi/pip).

```
C:\>pip install ptvsd
Downloading/unpacking ptvsd
  Running setup.py (path:c:\docume~1\user\locals~1\temp\pip_build_user\ptvsd\setup.py) egg_info for package ptvsd

Installing collected packages: ptvsd
  Running setup.py install for ptvsd

Successfully installed ptvsd
Cleaning up...
```

Now you should be able to `import ptvsd` in your Python interpreter.

Let's see how to prepare a [basic script](./Addin01.md) for debugging with PTVS

```python
# Book1.py
from xlpython import *

# Add these two lines
import ptvsd
try:
	ptvsd.enable_attach(secret='cows')
except ptvsd.AttachAlreadyEnabledError:
	pass


@xlfunc
def DoubleSum(x, y):
	return 2 * (x + y)
```

These to lines enable us to attach the Visual Studio debugger to the Python script, so that we can debug it. Note that the Python process itself must be running for this to work, and ExcelPython does not launch it until it is needed - so just to make sure that it is running and that the `Book1.py` script is loaded, click 'Import Python UDFs'.

(Note: the `try ... except` clause is needed because if `enable_attach` is called twice PTVS will complain that it's already been called. Since we have no way of controlling if it's already been called (for example if the script is reloaded, or if it is called in a different script) we catch the error and ignore it, as recommended in the PTVS documentation.)

At this point you may need to unblock a port on the Windows firewall:

![image](https://cloud.githubusercontent.com/assets/5197585/4387988/f02fbbe8-43e4-11e4-997e-31f12adbdf98.png)

Having done this you should now be in a position to attach the Visual Studio debugger to the Python script. To do this:
* open the Debug > Attach to Process dialog
* select 'Python remote debugging' as the transport
* specify the relevant qualifier in the 'Qualifier' box, for example `cows@localhost`
* click refresh and select the Python process

![image](https://cloud.githubusercontent.com/assets/5197585/4388048/e5ae4c06-43e5-11e4-8b24-300aa0ab3d4e.png)

Now you can set breakpoints in the code and the next time your function is invoked from Excel, the execution will be interrupted when the breakpoint is hit.

![image](https://cloud.githubusercontent.com/assets/5197585/4388088/59fefbaa-43e6-11e4-945a-31d66d3e2730.png)

## Debugging with Eclipse and PyDev

Coming soon!
