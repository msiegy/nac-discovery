# nac-discovery
A discovery and collection tool which pulls information from an inventory of switches and provides useful data points and recommendations on which ports to apply or exclude 802.1x NAC configurations. Information is exported to an Excel workbook so that any team can work with the data.

Ideally inventory files should be built dynamically using input from external sources and network tools like NETMRI, DNA Center, SolarWinds etc.

Built using the opensource Nornir and Napalm Python libraries.

*Local modification of mac_vendor_lookup required to handle cisco MAC formatting*
