
class Log
{
protected:
	HANDLE hLogFile;

public:
	Log();

	int Indent;
	void WriteIndent();
	void Write(const char* msg);

	static Log& GetInstance();
};

class LogCtx
{
protected:
	const char* fname;

public:
	LogCtx(const char* fname)
	{
		this->fname = fname;
		Log::GetInstance().WriteIndent();
		Log::GetInstance().Write(">");
		Log::GetInstance().Write(this->fname);
		Log::GetInstance().Write("\n");
		Log::GetInstance().Indent++;
	}

	~LogCtx()
	{
		Log::GetInstance().Indent--;
		Log::GetInstance().WriteIndent();
		Log::GetInstance().Write("<");
		Log::GetInstance().Write(this->fname);
		Log::GetInstance().Write("\n");
	}

	template<class T>
	LogCtx& operator% (const T& t)
	{
		Log::GetInstance().WriteIndent();
		return (*this) << t;
	}

	template<class T>
	LogCtx& operator<< (const T& t)
	{
		std::ostringstream oss;
		oss << t;
		Log::GetInstance().Write(oss.str().c_str());
		return *this;
	}

	//template<>
	LogCtx& operator<< (const std::string& t)
	{
		Log::GetInstance().Write(t.c_str());
		return *this;
	}

	//template<>
	LogCtx& operator<< (char * t)
	{
		Log::GetInstance().Write(t == NULL ? "<null string>" : t);
		return *this;
	}
};