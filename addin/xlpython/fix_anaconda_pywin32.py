import sys
import _win32sysloader
import os

# Anaconda and possibly other distributions have a bug in PyWin32,
# the `pythoncom` module can't be loaded because some required
# DLLs are located in the wrong place, so they aren't found

# force windows to load them with the full path so that subsequent
# imports work correctly

filename = "pywintypes%d%d.dll" % (sys.version_info[0], sys.version_info[1])

found = _win32sysloader.GetModuleFilename(filename)

if not found:
	found = _win32sysloader.LoadModule(os.path.join(sys.prefix, 'lib', 'site-packages', 'win32', filename))



    