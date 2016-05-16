#include "AdapterConfig.h"

int main()
{
	AdapterConfig conf;
	conf.select_adapter(L"vEthernet (wlan-vswitch)");
	//conf.set_auto_dns();
	//conf.set_static_ip(L"192.168.1.233", L"255.255.255.0", L"192.168.1.1");
	//conf.set_static_dns(L"124.16.136.254", nullptr);
	conf.set_auto_ip();

	return 0;
}