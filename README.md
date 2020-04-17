# collectswitchfacts.py
A discovery and collection tool which pulls information from an inventory of switches and provides useful data points and recommendations on which ports to apply or exclude 802.1x NAC configurations. Information is exported to an Excel workbook so that any team can work with the data to then build appropriate configs. Today this information gathering and port recommendation relies on LLDP, Mac Tables, Port Descriptions and MAC Vendor OUI lookup. 

Ideally inventory files should be built dynamically using input from external sources and network tools like NETMRI, DNA Center, SolarWinds etc.

Built using the opensource Nornir and Napalm Python libraries. Nornir inventory files not included in repo, please consult https://nornir.readthedocs.io/en/latest/tutorials/intro/inventory.html

*Local modification of mac_vendor_lookup library required to handle cisco MAC formatting. PR planned*


# iosnacconfparser.py
A Separate project which uses a switch configuration file parser and the python library ciscoconfparse for identifying where to apply appropriate NAC configs and then generates those configs for deployment. 

iosparser.py """ Iterates Cisco IOS configuration files in a directory to return a list of interfaces that contain relevant children statements. From this list generate configuration files that include NAC changes. One file for complete configuration and one file for new changes only. Today the script considers the following: Switchport mode access, shutdown status and description keywords as whether to apply NAC commands."""


Both scrips were tested with Catalyst 9300/500 switches running 16.6 and 16.8, but should run on most ios versions.
