#include "ExcelPython.h"

struct PyConsoleReadWrite
{
	PyObject_HEAD;
};

PyObject* PyConsoleReadWrite_Write(PyObject* self, PyObject* args)
{
	char* msg;
	if(!PyArg_ParseTuple(args, "s", &msg))
		return NULL;

	fputs(msg, stdout);

	Py_RETURN_NONE;
}

PyObject* PyConsoleReadWrite_Flush(PyObject* self, PyObject* args)
{
	fflush(stdout);

	Py_RETURN_NONE;
}

PyObject* PyConsoleReadWrite_ReadLine(PyObject* self, PyObject* args)
{
	static char buf[4096];
	fgets(buf, 4096, stdin);
	
	return PyString_FromString(buf);
}

//int GetSoftspace()
//{
//	return _soft;
//}
//void SetSoftspace(int s)
//{
//	_soft = s;
//}

	//attrs["write"] = &ConsoleReadWrite::Write;
	//attrs["flush"] = &ConsoleReadWrite::Flush;
	//attrs["readline"] = &ConsoleReadWrite::ReadLine;
	////attrs["softspace"] = Property(&ConsoleReadWrite::GetSoftspace, &ConsoleReadWrite::SetSoftspace);


static PyMethodDef PyConsoleReadWrite_Methods[] = {
	{ "write", (PyCFunction) PyConsoleReadWrite_Write, METH_VARARGS, "Writes to the console." },
	{ "flush", (PyCFunction) PyConsoleReadWrite_Flush, METH_NOARGS, "Flushes the console." },
	{ "readline", (PyCFunction) PyConsoleReadWrite_ReadLine, METH_NOARGS, "Reads a line from the console." },
	{NULL}  /* Sentinel */
};

PyTypeObject* PyConsoleReadWrite_Type()
{
	static bool ready = false;
	static PyTypeObject type;
	if(!ready)
	{
		memset(&type, 0, sizeof(PyTypeObject));
		type.ob_type = &PyType_Type;
		type.ob_refcnt = 1;
		type.tp_name = "Windows Console";
		type.tp_basicsize = sizeof(PyConsoleReadWrite);
		type.tp_flags = Py_TPFLAGS_DEFAULT;
		type.tp_doc = "Windows Console";
		type.tp_methods = PyConsoleReadWrite_Methods;

		if(0 != PyType_Ready(&type))
			throw PythonException();

		ready = true;
	}

	return &type;
}

PyObject* PyConsoleReadWrite_New()
{
	PyConsoleReadWrite* obj = PyObject_New(PyConsoleReadWrite, PyConsoleReadWrite_Type());
	return (PyObject*) obj;	
}


//#include <windows.h>
//
//#include "resource.h"
//
//const char g_szClassName[] = "myWindowClass";
//
//// Step 4: the Window Procedure
//LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
//{
//	switch(msg)
//	{
//	case WM_CLOSE:
//		DestroyWindow(hwnd);
//		break;
//
//	case WM_DESTROY:
//		PostQuitMessage(0);
//		break;
//
//	default:
//		return DefWindowProc(hwnd, msg, wParam, lParam);
//	}
//	return 0;
//}
//
//int __stdcall DialogFunc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
//{
//	switch(uMsg)
//	{
//	case WM_COMMAND:
//		{
//			WORD id = LOWORD(wParam);
//			switch(id)
//			{
//			case IDOK:
//				break;
//
//			case IDCANCEL:
//				EndDialog(hWnd, 1);
//				break;
//			}
//
//		}
//		return TRUE;
//
//	default:
//		return FALSE;
//	}
//}
//
//void ShowConsoleDialog(int hWndParent)
//{
//	WNDCLASSEXA wc;
//	wc.cbSize = sizeof(WNDCLASSEX);
//	wc.style = 0;
//	wc.lpfnWndProc = WndProc;
//	wc.cbClsExtra = 0;
//	wc.cbWndExtra = 0;
//	wc.hInstance = GetModuleHandle(NULL);
//	wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
//	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
//	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
//	wc.lpszMenuName = NULL;
//	wc.lpszClassName = "ExcelPythonConsoleWindow";
//	wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
//	RegisterClassExA(&wc);
//
//	// Step 2: Creating the Window
//	HWND hwnd = CreateWindowExA(
//		WS_EX_CLIENTEDGE,
//		"ExcelPythonConsoleWindow",
//		"The title of my window",
//		WS_OVERLAPPEDWINDOW,
//		CW_USEDEFAULT,
//		CW_USEDEFAULT,
//		240,
//		120,
//		(HWND) hWndParent,
//		NULL,
//		GetModuleHandle(NULL),
//		NULL);
//
//	if(hwnd == NULL)
//	{
//		MessageBoxA((HWND) hWndParent, "Window Creation Failed!", "Error!", MB_ICONEXCLAMATION | MB_OK);
//		return;
//	}
//
//	EnableWindow((HWND) hWndParent, FALSE);
//	ShowWindow(hwnd, SW_SHOWNORMAL);
//	UpdateWindow(hwnd);
//
//	MSG msg;
//	while(GetMessage(&msg, NULL, 0, 0) > 0)
//	{
//		TranslateMessage(&msg);
//
//		if(msg.hwnd != hwnd && msg.message == WM_ACTIVATE)
//			SetActiveWindow(hwnd);
//		else
//			DispatchMessage(&msg);
//	}
//
//	EnableWindow((HWND) hWndParent, TRUE);
//}