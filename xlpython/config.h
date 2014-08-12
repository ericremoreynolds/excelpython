class Config
{
public:
	typedef std::unordered_map<std::string, Config*> ConfigMap;
	typedef std::unordered_map<std::string, std::string> ValueMap;

protected:
	static ConfigMap configs;

	Config(const std::string& filename);
	~Config();

	FILETIME ftLastModify;

public:
	static Config* GetConfig(const std::string& filename);

	std::string GetValue(const std::string& key);
	std::string GetValue(const std::string& key, const std::string& dfault);
	bool TryGetValue(const std::string& key, std::string& value);

	int GetValueAsInt(const std::string& key);
	int GetValueAsInt(const std::string& key, int dfault);

	bool HasValue(const std::string& key);
	ValueMap::iterator GetIterator();

	bool CheckRPCServer();
	void ActivateRPCServer();

	IDispatch* pInterface;
	ValueMap values;
	HANDLE hJob;
};
