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
	char* init_error_str() const;

	HRESULT call_result() const;
	int wmi_result() const;

private:
	void print_wmi_error(HRESULT hr);
	void CreateOneElementBstrArray(VARIANT *v, LPCWSTR s);
	bool adapter_selected() const;

private:
	bool _init;
	HRESULT _last_hres;
	int _wmi_ret;
	int _failed_step;
	int _adapter_index;
	IWbemLocator* _locator;
	IWbemServices* _service;
};

