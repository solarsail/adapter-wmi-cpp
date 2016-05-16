#include <sstream>
#include <iostream>
#include "AdapterConfig.h"
#include "comutil.h"
#pragma comment(lib, "wbemuuid.lib")//wmi
#pragma comment(lib, "comsuppw.lib")


AdapterConfig::AdapterConfig() :
	_init(false),
	_last_hres(0),
	_failed_step(0),
	_adapter_index(-1)
{
	_init = init();
	if (!_init) {
		std::wcerr << L"初始化错误：" << init_error_str() << std::endl;
	}
}


AdapterConfig::~AdapterConfig()
{
	finalize();
}

bool AdapterConfig::select_adapter(LPCTSTR name)
{
	bool ret = false;

	std::wstringstream ws;
	ws << L"SELECT * From Win32_NetworkAdapter WHERE NetConnectionID='" << name << L"'";
	_bstr_t query(ws.str().c_str());

	IEnumWbemClassObject *e = nullptr;
	_service->ExecQuery(L"WQL", query, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 0, &e);

	IWbemClassObject *adapter = nullptr;
	ULONG u;
	VARIANT data;
	if (e) {
		e->Next(WBEM_INFINITE, 1, &adapter, &u);
		if (u > 0) {
			adapter->Get(L"Index", 0, &data, 0, 0);
			_adapter_index = data.intVal;
			adapter->Release();
			ret = true;
		}
		e->Release();
	}

	return ret;
}

HRESULT AdapterConfig::result()
{
	
	return _last_hres;
}

bool AdapterConfig::init()
{
	if (_init)
		return true;

	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------
	_last_hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(_last_hres)) {
		_failed_step = 1;
		return false;
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------
	_last_hres = CoInitializeSecurity(
		NULL,
		-1,                          // COM negotiates service
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		NULL                         // Reserved
	);

	if (FAILED(_last_hres)) {
		_failed_step = 2;
		CoUninitialize();
		return false;
	}

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI -------------------------
	_last_hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&_locator);

	if (FAILED(_last_hres)) {
		_failed_step = 3;
		CoUninitialize();
		return false;
	}

	// Step 4: ---------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	// Connect to the local root\cimv2 namespace
	// and obtain pointer pSvc to make IWbemServices calls.
	_last_hres = _locator->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		NULL,                    // User name. NULL = current user
		NULL,                    // User password. NULL = current
		0,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		0,                       // Authority (for example, Kerberos)
		0,                       // Context object 
		&_service                // pointer to IWbemServices proxy
	);

	if (FAILED(_last_hres)) {
		_failed_step = 4;
		_locator->Release();
		CoUninitialize();
		return false;
	}

	// Step 5: --------------------------------------------------
	// Set security levels for the proxy ------------------------
	_last_hres = CoSetProxyBlanket(
		_service,                    // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
		NULL,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(_last_hres)) {
		_failed_step = 5;
		_service->Release();
		_locator->Release();
		CoUninitialize();
		return false;
	}

	return true;
}

void AdapterConfig::finalize()
{
	if (_init) {
		if (_service)
			_service->Release();
		if (_locator)
			_locator->Release();

		CoUninitialize();
		_init = false;
	}
}

bool AdapterConfig::set_static_dns(LPCTSTR lpszDns1, LPCTSTR lpszDns2)
{
	if (!adapter_selected())
		return false;

	ULONG size = 0;
	size += (lpszDns1 != NULL);
	size += (lpszDns2 != NULL);
	//if (size == 0)
		//return false;

	bool ret = true;

	std::wstringstream ws;
	ws << L"Win32_NetworkAdapterConfiguration.Index='" << _adapter_index << L"'";
	_bstr_t instance_path(ws.str().c_str()); //index为网卡号
	_bstr_t class_path(L"Win32_NetworkAdapterConfiguration");
	_bstr_t set_dns_method(L"SetDNSServerSearchOrder");
	LPCTSTR set_dns_method_arg = L"DNSServerSearchOrder";
	_bstr_t dns1(lpszDns1);
	_bstr_t dns2(lpszDns2);

	long DnsIndex1[] = { 0 };
	long DnsIndex2[] = { 1 };

	SAFEARRAY *ip_list = SafeArrayCreateVector(VT_BSTR, 0, size);
	if (size > 0) {
		SafeArrayPutElement(ip_list, DnsIndex1, dns1.GetBSTR());
	}
	if (size > 1) {
		SafeArrayPutElement(ip_list, DnsIndex2, dns2.GetBSTR());
	}
	
	VARIANT dns;
	VARIANT out_value;
	dns.vt = VT_ARRAY | VT_BSTR;
	dns.parray = ip_list;

	IWbemClassObject *class_obj = nullptr;
	IWbemClassObject *arg_class_obj = nullptr;
	IWbemClassObject *arg_inst = nullptr;
	IWbemClassObject *out_obj = nullptr;

	// 获取 Win32_NetworkAdapterConfiguration 类对象
	_last_hres = _service->GetObject(class_path, 0, NULL, &class_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 获取方法参数的类对象
	_last_hres = class_obj->GetMethod(set_dns_method, 0, &arg_class_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 实例化参数对象
	_last_hres = arg_class_obj->SpawnInstance(0, &arg_inst);
	if (FAILED(_last_hres)) goto err;
	// 填充参数内容
	_last_hres = arg_inst->Put(set_dns_method_arg, 0, &dns, 0);
	if (FAILED(_last_hres)) goto err;
	// 在 Win32_NetworkAdapterConfiguration.Index='_adapter_index' 的实例上调用方法
	_last_hres = _service->ExecMethod(instance_path, set_dns_method, 0, NULL,
		arg_inst, &out_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 获取返回值
	out_obj->Get(L"ReturnValue", 0, &out_value, 0, 0);
	int code = out_value.intVal;

out:
	if (out_obj)
		out_obj->Release();
	if (arg_inst)
		arg_inst->Release();
	if (arg_class_obj)
		arg_class_obj->Release();
	if (class_obj)
		class_obj->Release();

	return ret;

err:
	ret = false;
	print_wmi_error(_last_hres);
	goto out;
}

bool AdapterConfig::set_static_ip(LPCTSTR lpszIP, LPCTSTR lpszNetmask, LPCTSTR lpszGateway)
{
	if (!adapter_selected())
		return false;

	bool ret = true;

	std::wstringstream ws;
	ws << L"Win32_NetworkAdapterConfiguration.Index='" << _adapter_index << L"'";
	_bstr_t instance_path(ws.str().c_str());
	_bstr_t class_path(L"Win32_NetworkAdapterConfiguration");
	
	_bstr_t enable_static_method(L"EnableStatic");
	_bstr_t set_gateway_method(L"SetGateways");

	VARIANT address_obj;
	VARIANT netmask_obj;
	VARIANT gateway_obj;
	VARIANT out_value;

	CreateOneElementBstrArray(&address_obj, lpszIP);
	CreateOneElementBstrArray(&netmask_obj, lpszNetmask);
	CreateOneElementBstrArray(&gateway_obj, lpszGateway);
	//VARIANT  metric_obj;
	//CreateOneElementBstrArray(&metric_obj, L"1");

	IWbemClassObject *class_obj = nullptr;
	IWbemClassObject *arg_class_obj = nullptr;
	IWbemClassObject *arg_inst = nullptr;
	IWbemClassObject *out_obj = nullptr;

	int code = 0;

	// 获取 Win32_NetworkAdapterConfiguration 类对象
	_last_hres = _service->GetObject(class_path, 0, NULL, &class_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 获取方法参数的类对象
	_last_hres = class_obj->GetMethod(enable_static_method, 0, &arg_class_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 实例化参数对象
	_last_hres = arg_class_obj->SpawnInstance(0, &arg_inst);
	if (FAILED(_last_hres)) goto err;
	// 填充参数内容
	_last_hres = arg_inst->Put(L"IPAddress", 0, &address_obj, 0);
	_last_hres = arg_inst->Put(L"SubnetMask", 0, &netmask_obj, 0);
	// 在 Win32_NetworkAdapterConfiguration.Index='_adapter_index' 的实例上调用方法
	_last_hres = _service->ExecMethod(instance_path, enable_static_method, 0, NULL,
		arg_inst, &out_obj, NULL);
	if (FAILED(_last_hres)) goto err;

	out_obj->Get(L"ReturnValue", 0, &out_value, 0, 0);
	code = out_value.intVal;
	// TODO: check return code

	// 清理中间对象
	arg_inst->Release();
	arg_inst = nullptr;
	arg_class_obj->Release();
	arg_class_obj = nullptr;
	out_obj->Release();
	out_obj = nullptr;

	// 修改网关

	// 获取方法参数的类对象
	_last_hres = class_obj->GetMethod(set_gateway_method, 0, &arg_inst, NULL);
	if (FAILED(_last_hres)) goto err;
	// 实例化参数对象
	_last_hres = arg_inst->SpawnInstance(0, &arg_class_obj);
	if (FAILED(_last_hres)) goto err;
	// 填充参数内容
	_last_hres = arg_class_obj->Put(L"DefaultIPGateway", 0, &gateway_obj, 0);
	//_last_hres = arg_class_obj->Put(L"GatewayCostMetric", 0, &metric_obj, 0);
	// 在 Win32_NetworkAdapterConfiguration.Index='_adapter_index' 的实例上调用方法
	_last_hres = _service->ExecMethod(instance_path, set_gateway_method, 0, NULL,
		arg_class_obj, &out_obj, NULL);
	if (FAILED(_last_hres)) goto err;

	out_obj->Get(L"ReturnValue", 0, &out_value, 0, 0);
	code = out_value.intVal;
	// TODO: check return code

out:
	// 清理
	if (out_obj)
		out_obj->Release();
	if (arg_inst)
		arg_inst->Release();
	if (arg_class_obj)
		arg_class_obj->Release();
	if (class_obj)
		class_obj->Release();

	return ret;

err:
	// 错误处理
	ret = false;
	print_wmi_error(_last_hres);
	goto out;
}

bool AdapterConfig::set_auto_dns()
{
	return set_static_dns(0, 0);
}

bool AdapterConfig::set_auto_ip()
{
	if (!adapter_selected())
		return false;

	bool ret = true;

	std::wstringstream ws;
	ws << L"Win32_NetworkAdapterConfiguration.Index='" << _adapter_index << L"'";
	_bstr_t instance_path(ws.str().c_str());
	_bstr_t class_path(L"Win32_NetworkAdapterConfiguration");
	_bstr_t enable_dhcp_method(L"EnableDHCP");

	IWbemClassObject *class_obj = nullptr;
	IWbemClassObject *out_obj = nullptr;

	// 获取 Win32_NetworkAdapterConfiguration 类对象
	_last_hres = _service->GetObject(class_path, 0, NULL, &class_obj, NULL);
	if (FAILED(_last_hres)) goto err;
	// 在 Win32_NetworkAdapterConfiguration.Index='_adapter_index' 的实例上调用方法
	_last_hres = _service->ExecMethod(instance_path, enable_dhcp_method, 0, NULL,
		NULL, &out_obj, NULL);
	if (FAILED(_last_hres)) goto err;

	VARIANT out_value;
	out_obj->Get(L"ReturnValue", 0, &out_value, 0, 0);
	int code = out_value.intVal;

out:
	if (out_obj)
		out_obj->Release();
	if (class_obj)
		class_obj->Release();

	return ret;

err:
	ret = false;
	print_wmi_error(_last_hres);
	goto out;
}

void AdapterConfig::CreateOneElementBstrArray(VARIANT*  v, LPCWSTR  s)
{
	SAFEARRAYBOUND  bound[1];
	SAFEARRAY*  array;
	bound[0].lLbound = 0;
	bound[0].cElements = 1;
	array = SafeArrayCreate(VT_BSTR, 1, bound);
	long  index = 0;
	BSTR  bstr = SysAllocString(s);
	SafeArrayPutElement(array, &index, bstr);
	SysFreeString(bstr);
	VariantInit(v);
	v->vt = VT_BSTR | VT_ARRAY;
	v->parray = array;
}

wchar_t * AdapterConfig::init_error_str() const
{
	static wchar_t* estr[] = {
		L"No error",
		L"Failed to initialize COM library",
		L"Failed to initialize security",
		L"Failed to create IWbemLocator object",
		L"Could not connect",
		L"Could not set proxy blanket"
	};
	return estr[_failed_step];
}

bool AdapterConfig::adapter_selected() const
{
	return _adapter_index >= 0;
}

void AdapterConfig::print_wmi_error(HRESULT hr) {
	IWbemStatusCodeText * pStatus = NULL;
	HRESULT hres = CoCreateInstance(CLSID_WbemStatusCodeText, 0,
		CLSCTX_INPROC_SERVER, IID_IWbemStatusCodeText, (LPVOID *)&pStatus);
	if (S_OK == hres) {
		BSTR bstrError;
		hres = pStatus->GetErrorCodeText(hr, 0, 0, &bstrError);
		if (S_OK != hres)
			bstrError = SysAllocString(L"Get last error failed");
		std::wcerr << L"WMI Error: " << bstrError << std::endl;
		SysFreeString(bstrError);
		pStatus->Release();
	}
}