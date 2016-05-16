#pragma once
#include <string>
#include "windows.h"
#include "Wbemcli.h"

class AdapterConfig
{
public:
	AdapterConfig();
	~AdapterConfig();

	bool select_adapter(LPCTSTR name);

	bool init();
	void finalize();
	bool set_static_dns(LPCTSTR lpszDns1, LPCTSTR lpszDns2);
	bool set_static_ip(LPCTSTR lpszIP, LPCTSTR lpszNetmask, LPCTSTR lpszGateway);
	bool set_auto_dns();
	bool set_auto_ip();
	HRESULT result();
	void print_wmi_error(HRESULT hr);
	void  CreateOneElementBstrArray(VARIANT *v, LPCWSTR s);
	wchar_t* init_error_str() const;

private:
	bool adapter_selected() const;

private:
	bool _init;
	HRESULT _last_hres;
	int _failed_step;
	int _adapter_index;
	IWbemLocator* _locator;
	IWbemServices* _service;
};

