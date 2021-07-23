# PCInfoHunter

SYNTAX
-----------
pcinfohunter [-search] [-update] [[-computer] <String>] [[-file] <String>] [[-save] <String>] [[-log] <String>] [-i]

DESCRIPTION
-----------
This is a tool with the purpose of creating, updating, and querying a database of system info, AD info, and custom info for hosts in the domain's network.
It's composed of two main functions (search and update) which call other functions to proccess information and increase functionality. The search function 
queries the data from the table for the specified hostnames. The update function modifies a table depending on what's already in it and changes in AD. 
If the host info is not currently on the table, a new row will be added. If host is online its fields will be updated. If the host was deleted from AD,
all of its info will be deleted from the table.
  
OPTIONS
----------
-search         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Search data for specified hostnames (mandatory)<br>
-update         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Changes table values by updating fields for each host in AD (mandatory)<br>
-computer       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Used to input a string or text file with a list of hosts to be searched (mandatory)<br>
-file           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Takes the location of import table<br>
-save           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Saves table to specified csv or txt file<br>
-savelog        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Saves information about script and table statistics such as script duration, number of updated computers, name of computers with new info, etc.<br>
-i              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Activates user interaction meaning it will ask user for input not provided such as list of items, strings to search for, saving location, etc.<br>
  
PROPERTIES
----------
HostName, ModelNumber, OSversionModel, ImageVersion, Department, RAM, HDmodel, HDsize, CPU, LastDomainJoinDate, LastLogonDate, OSversionNum, IPaddress, AD_IPaddress,
Location, AD_Path, Manufacturer, SerialNumber, BIOSversion, BIOSdate, BaseBoard, VideoCard, FreeSpace, MemoryNumber, PatchVersion, LastUpdate
  
EXAMPLES
-----------
pcinfohunter -search -computer computer1,computer2,computer3... -save C:\Users\Current_User\Desktop\results.csv
pcinfohunter -search -computer computer_list.txt 
pcinfohunter -search *
pcinfohunter -update -file C:\Users\Current_User\Desktop\results.csv -save C:\Users\Current_User\Desktop\results.csv -log C:\Users\Current_User\Desktop\results_log.txt
pcinfohunter -i

NOTES
-----
For INTERACTIVE mode use -i
