#include "ExcelPython.h"

PythonException::PythonException()
	: Exception()
{
	PyNewRef type(NULL, false);
	PyNewRef exc(NULL, false);
	PyNewRef traceback(NULL, false);
	PyErr_Fetch(&type.ptr, &exc.ptr, &traceback.ptr);
	if(type != NULL)
	{
		if(exc != NULL)
		{
			PyNewRef repr(PyObject_Str(exc), false);
			*this << PyString_AS_STRING(repr.ptr);
		}

		if(traceback != NULL)
		{
			try
			{
				PyBorrowedRef modtb(PyImport_ImportModule("traceback"), false);
				if(modtb == NULL)
				{
					PyErr_Clear();
					throw Exception() << "Could not get 'traceback' module";
				}

				PyNewRef format_tb(PyString_FromString("format_tb"), false);
				PyNewRef list(PyObject_CallMethodObjArgs(modtb, format_tb, traceback, NULL), false);
				if(list == NULL)
				{
					PyErr_Clear();
					throw Exception() << "Call to 'traceback.format_tb' failed";
				}

				Py_ssize_t n = PyList_Size(list);
				for(Py_ssize_t k=0; k<n; k++)
				{
					PyObject* str = PyList_GET_ITEM(list.ptr, k);
					if(str == NULL || 0 == PyString_Check(str))
						throw Exception() << "Line " << k << " of formatted traceback is missing or is not a string";

					if(k == 0)
						*this << "\n";

					*this << PyString_AS_STRING(str);
				}
			}
			catch(Exception& ex)
			{
				throw ex.context() << "while formatting python error";
			}
		}
	}
	else
		*this << "Unidentified Python error";
}