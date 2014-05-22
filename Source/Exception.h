#pragma once

#include <sstream>
#include <string>

class Exception
{
protected:
	std::string _message;

public:
	Exception()
		: _message("")
	{
	}

	Exception(const Exception& inner)
		: _message("")
	{
		(*this) << inner.message();
	}

	Exception(const std::exception& inner)
		: _message("")
	{
		(*this) << "C++ exception: " << inner.what();
	}

	template<class T>
	Exception& operator<< (const T& t)
	{
		std::ostringstream oss;
		oss << t;
		_message = _message + oss.str();
		return *this;
	}

	//template<>
	Exception& operator<< (const std::string& t)
	{
		_message = _message + t;
		return *this;
	}

	//template<>
	Exception& operator<< (const char * t)
	{
		_message = _message + t;
		return *this;
	}

	Exception& context()
	{
		_message = _message + "\n  ";
		return *this;
	}

	const std::string& message() const
	{
		return _message;
	}
};

class PythonException : public Exception
{
public:
	PythonException();
};

class ComException : public Exception
{
protected:
	_com_error _err;

public:
	ComException(HRESULT hr)
		: _err(hr)
	{
		_bstr_t desc = _err.Description();
		boost::scoped_array<char> narrow(new char[desc.length()+1]);
		WideCharToMultiByte(CP_ACP, 0, desc, desc.length(), narrow.get(), desc.length() + 1, "?", NULL);
		narrow[desc.length()] = 0;
		*this << narrow;
	}
};
