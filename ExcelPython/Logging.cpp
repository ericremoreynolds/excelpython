#include "ExcelPython.h"

#include <ctime>

Log& Log::GetInstance()
{
	static Log L;
	return L;
}

Log::Log()
{
	hLogFile = INVALID_HANDLE_VALUE;
	this->Indent = 0;

	TCHAR moduleName[MAX_PATH+10];
	GetModuleFileName(hModuleDLL, moduleName, MAX_PATH);
	_tcscat_s(moduleName, MAX_PATH+10, _T(".log"));

	hLogFile = CreateFile(
		moduleName, 
		FILE_APPEND_DATA,
		FILE_SHARE_READ,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		NULL);

	if(hLogFile != INVALID_HANDLE_VALUE)
	{
		__time64_t rawtime;
		struct tm timeinfo;
		char buffer[120];

		_time64(&rawtime);
		errno_t err = localtime_s(&timeinfo, &rawtime);
		if(err)
		{
			Write("\n---------- <error getting time> ----------\n");
		}
		else
		{
			strftime(buffer, 120, "\n---------- %Y-%m-%d %H:%M:%S ----------\n", &timeinfo);
			Write(buffer);
		}
	}
}

void Log::Write(const char* msg)
{
	if(hLogFile != INVALID_HANDLE_VALUE)
	{
		DWORD nBytesWritten;
		BOOL bErrorWrite = WriteFile(hLogFile, msg, strlen(msg), &nBytesWritten, NULL);
		BOOL bErrorFlush = FlushFileBuffers(hLogFile);
	}
}

void Log::WriteIndent()
{
	static const char spaces[] = "                                                                                                                                                                                                                                                                                                            ";
	if(hLogFile != INVALID_HANDLE_VALUE)
	{
		DWORD nBytesWritten;
		BOOL bErrorWrite = WriteFile(hLogFile, spaces, Indent, &nBytesWritten, NULL);
	}
}