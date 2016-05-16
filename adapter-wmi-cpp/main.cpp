#include "AdapterConfig.h"

int main()
{
	AdapterConfig conf;
	conf.select_adapter(L"vEthernet (wlan-vswitch)");
	conf.set_static_dns(L"124.16.136.254", nullptr);

	return 0;
}