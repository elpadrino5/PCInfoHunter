#!/usr/bin/env powershell

#parameters
param([Switch]$update, [Switch]$search, [Parameter(ValueFromPipeline=$true)][String[]]$computer, [String]$file, [String]$save, [String]$log, [Switch]$i)

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region Data Processing Functions

function AD($pc)
{
    $ad = Get-ADComputer -Identity $pc -Properties *
    return $ad
}

#gets department from name of OU in AD
function Department($string) 
{
    #splits the path string to match with possible department
    $department = $string -split '/' | Select-String -Pattern 'finance','Government','Human','Office','Procurement','Risk','Access','Cranes','Gates', 'Marketing', 'Facilities', 'Information', 'Public'
    if ($department -eq $null)
        {
        $depstring = $stringAD -split '/' -split'-'| Select-String -Pattern 'BIMT', 'PCOB', 'SOC','TMT','BI','BMT'
        $department = $depstring | Select -First 1
        }
    if ($department -match "Procurement and Contracting Services")
        {$department = 'Procurement'}
    if ($department -match "Facilities Development")
        {$department = 'Engineering'}
    if ($department -match "Government & External Affairs")
        {$department = 'Government'}
    if ($department -match "Office of Chief Executive Officer")
        {$department = 'Office of CEO'}
    if ($department -match "Office of Executive Vice President")
        {$department = 'Office of EVP'}
    if ($department -match "Risk Management and Safety")
        {$department = 'Risk Management'}
    if ($department -match "Human Resources")
        {$department = 'HR'} 
    if ($department -match "Information Technology")
        {$department = 'IT'}
    if ($department -match "Public_Safety")
        {$department = 'Public Safety'}  

    return $department
} 

function FreeSpace($freeSpace)
{
    $freespace = $freeSpace | select -First 1
    if ($freespace -eq $null)
    {
        Write-Error "freespace property is null"
    }
    else
    {
        $freespace = $freespace / (1024*1024*1024)
        $freespace = [math]::Round($freespace,0)
    }
    return $freespace;
}

function HDsize($diskDrive)
{
    $intType = $diskDrive.interfacetype
    $size = $diskDrive.size
    $model = $diskDrive.model
    $HDsize = '' 
    #get-wmiobject win32_diskdrive -ComputerName $pc | select -ExpandProperty interfacetype -Skip $i -First 1

    for($i = 0; $i -lt $model.Count; $i++)
    {
        if (($intType | select -Skip $i -First 1) -ne 'USB')
        {
            $HDsize = $size | select -Skip $i -First 1
            break;            
        }
    }
    if ($HDsize -eq $null)
    {
        $HDsize = '';
    }
    else
    {
        $HDsize = $HDsize / (1024*1024*1024)
        $HDsize = [math]::Round($HDsize,0)
    }
    return $HDsize;
}

function HDmodel($diskDrive)
{
    $intType = $diskDrive.interfacetype
    $model = $diskDrive.model
    $hdmodel = ''

    for($i = 0; $i -lt $model.Count; $i++)
    {
        if (($intType | select -Skip $i -First 1) -ne 'USB')
        {
            $hdmodel = $model | select -Skip $i -First 1
            break;            
        }
    }
    return $hdmodel;
}

function ImageVersion($pc)
{    
        if (Test-Path \\$pc\C$\Custom\Docs\Version.txt)
        {
         $imgVersion = Get-Content -Path \\$pc\C$\Custom\Docs\Version.txt    
        }
        else
        { 
          $imgVersion =  "Not Found"
        }            
    return $imgVersion;    
}

function IP($ipAddress)
{
    $IP = $ipAddress | Where { $_ -like '172.*' -or $_ -like '10.*'} | Select -First 1 
    return $IP;
}  

#gets location from name of OU in AD
function Location($string) 
{
    #splits the path string to match with possible department
    $locstring = $string -split '/' -split '-'| Select-String -Pattern 'BIMT', 'PCOB', 'SOC', 'TMT','BI','BMT'
    $location = $locstring | Select -First 1
    if ($location -eq $null)
        {$location = ''}
    return $location
}

 function MemoryNumber($partNumber)
{
    for($i = 0; $i -lt $partNumber.Count; $i++)
    {
        if (($partNumber | select -Skip $i -First 1) -ne $null)
        {
            $mn = $partNumber | select -Skip $i -First 1
            break;            
        }
    }
    return $mn;
}

function OSversionModel($osvermodel)
{
    $osvm = $osvermodel -replace 'Windows','Win' -replace 'Enterprise','Ent' -replace 'Professional','Pro' 
    return $osvm
}

function PatchVersion($pc)
{    
    if (Test-Path \\$pc\C$\Custom\Docs\patch_version.txt)
    {
        $PatchVersion = Get-Content -Path \\$pc\C$\Custom\Docs\patch_version.txt    
    }
    else
    { 
        $PatchVersion = "Not Found"
    }            
    return $PatchVersion;    
}

function RAM($RAM)
{
    if ($RAM -eq $null)
    {
        $RAM = '';
    }
    else
    {
        $RAM = $RAM / (1024*1024)
        $RAM = [math]::Round($RAM,0)
    }
    return $RAM
}

function VideoCard($videoController)
{ 
    $vc = '';
    $adapterDacType = $videoController.adapterDacType
    $name = $videoController.name

    for($i = 0; $i -lt $name.Count; $i++)
    {
        $adt = $adapterDacType | select -Skip $i -First 1
        if ($adt -ne $null)
        {
            $vc = $videoController | Where-Object {$_.adapterDacType -eq $adt} | select -expandproperty name -Skip $i -First 1
            break;            
        }
    }
    return $vc;
}

#endregion

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region Utility Functions

 function CreateLog($log)
{
    $date = GetDate
    if ($log) { $p = $log } 
    else      { $p = "PCInfo_$date.log" }
    Set-Content -Value "" -Path $p
    #print -string "log path at createLog --> $p"
    return $p
}

 function CreateTable()
{
    $tbl = "Import Table";
    $table = New-Object System.Data.DataTable "$tbl";

    #Columns
    $col1  = New-Object System.Data.DataColumn HostName,([string]); 
    $col2  = New-Object System.Data.DataColumn ModelNumber,([string]);
    $col3  = New-Object System.Data.DataColumn OSversionModel,([string]);
    $col4  = New-Object System.Data.DataColumn ImageVersion,([string]);
    $col5  = New-Object System.Data.DataColumn Department,([string]);
    $col6  = New-Object System.Data.DataColumn RAM,([string]);
    $col7  = New-Object System.Data.DataColumn HDmodel,([string]);
    $col8  = New-Object System.Data.DataColumn HDsize,([string]);
    $col9  = New-Object System.Data.DataColumn CPU,([string]);
    $col10 = New-Object System.Data.DataColumn LastDomainJoinDate,([string]); #-----------------------
    $col11 = New-Object System.Data.DataColumn LastLogonDate,([string]);
    $col12 = New-Object System.Data.DataColumn OSversionNum,([string]);
    $col13 = New-Object System.Data.DataColumn IPaddress,([string]);
    $col14 = New-Object System.Data.DataColumn AD_IPaddress,([string]);
    $col15 = New-Object System.Data.DataColumn Location,([string]);
    $col16 = New-Object System.Data.DataColumn AD_Path,([string]); #----------------------------
    $col17 = New-Object System.Data.DataColumn Manufacturer,([string]);
    $col18 = New-Object System.Data.DataColumn SerialNumber,([string]);
    $col19 = New-Object System.Data.DataColumn BIOSversion,([string]);
    $col20 = New-Object System.Data.DataColumn BIOSdate,([string]);
    $col21 = New-Object System.Data.DataColumn BaseBoard,([string]);
    $col22 = New-Object System.Data.DataColumn VideoCard,([string]); 
    $col23 = New-Object System.Data.DataColumn FreeSpace,([string]);
    $col24 = New-Object System.Data.DataColumn MemoryNumber,([string]);
    $col25 = New-Object System.Data.DataColumn PatchVersion,([string]);
    $col26 = New-Object System.Data.DataColumn LastUpdate,([string]);

    #Add the Columns
    $table.columns.add($col1);
    $table.columns.add($col2);
    $table.columns.add($col3);
    $table.columns.add($col4);
    $table.columns.add($col5);
    $table.columns.add($col6);
    $table.columns.add($col7);
    $table.columns.add($col8);
    $table.columns.add($col9);
    $table.columns.add($col10);
    $table.columns.add($col11);
    $table.columns.add($col12);
    $table.columns.add($col13);
    $table.columns.add($col14);
    $table.columns.add($col15);
    $table.columns.add($col16);
    $table.columns.add($col17);
    $table.columns.add($col18);
    $table.columns.add($col19);
    $table.columns.add($col20);
    $table.columns.add($col21);
    $table.columns.add($col22);
    $table.columns.add($col23);
    $table.columns.add($col24);
    $table.columns.add($col25);
    $table.columns.add($col26);

    return ,$table;
}

function GetHostname($computer)
{     
    if ($i)
    {
        $userInput = Read-host "`nType hostnames or path of text file (separate hostnames by a ',' or single space)"
        if ($userInput -like '*,*')
        {
            $userInput = $userInput -split ','
            $userInput = $userInput -replace '\s',''
        }
        else { $userInput = $userInput -split '\s' }

        $computer = $userInput
        return $computer
    }
      
    if ($computer -like "*.txt")
    {
        Write-Output "Input hostlist received: $computer"
        $computer = Get-Content -Path $computer;
    }
    elseif ($computer -eq "*")
    {
        $computer = FilterHostnames
    }
    else
    {
        $computer = $computer.Trim();
        $computer = $computer -replace "\n"," "
        $computer = $computer -replace "\s+"," "
    }                   

    return $computer
}

 function Getdate
{
    $datesec = get-date -Format "yyyy-MM-dd_HH-mm-ss" | Out-String
    $datesec = $datesec -replace '\s','';
    return $datesec
}

function GetDuration($stopwatch)
{
    $stopwatch.stop();
    $min = $stopwatch.Elapsed.Minutes
    $swstr = $min -as [string]
    $swstr+= " minutes "
    $sec = $stopwatch.Elapsed.Seconds 
    $swstr += $sec -as [string]
    $swstr+= " seconds "
    return $swstr
}

function FilterHostnames
{
    $hts = Get-ADComputer -Filter 'Name -like "*"' -Properties Name,OperatingSystem | Where-Object {($_.operatingSystem -notlike '*server*') -and ($_.operatingSystem -ne $null)} | Sort-Object | foreach{$_.name} 
    $servers = "dsc-nvr-2 fake-f9kl10s20x graviton HEIMDALL1 NAGGER photon photon1 photon2 photon3"
    #[System.Collections.ArrayList]$allhosts = [System.Collections.ArrayList]::new();
    $allhosts = @()

    foreach($ea in $hts)
    {
        $each = $ea | Out-String 
        $each = $each -replace '\s',''
        $each = $each -replace '\n',''
          
        if ($each -ne '')
        {
            #$os = Get-ADComputer $each  -Properties OperatingSystem | select -ExpandProperty OperatingSystem -First 1 
            if ((-not($each  -like '*dsc-nvr-2*')) -and (-not($each  -like '*NAGGER*')) -and (-not($each  -like '*fake-f9kl10s20x*')) -and (-not($each  -like '*graviton*')) -and (-not($each  -like '*HEIMDALL1*')) -and (-not($each  -like '*photon*')))
            {
                $allhosts += $each
            } 
        }
    }
    return $allhosts
}

function Help
{
        Write-Host 
("SYNTAX
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
-search         Search data for specified hostnames (mandatory)
-update         Changes table values by updating fields for each host in AD (mandatory)
-computer       Used to input a string or text file with a list of hosts to be searched (mandatory)
-file           Takes the location of import table
-save           Saves table to specified csv or txt file
-savelog        Saves information about script and table statistics such as script duration, number of updated computers, name of computers with new info, etc.
-i              Activates user interaction meaning it will ask user for input not provided such as list of items, strings to search for, saving location, etc.

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

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!                               !!
!!  For INTERACTIVE mode use -i  !!
!!                               !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

}

function Import_Csv($file)
{
    $default_path = "C:\Users\amartinez\Desktop\repos\helpdesk\InfoHunterData\PCInfo\PCInfo.csv"

    if (!($file))
        {
            if (Test-Path -Path $default_path)
                {$csv = Import-Csv -Path $default_path}
            else 
                {Write-Error "Default path ($default_path) for -input (input csv file) does not exist!"; break}            
        }
    else
        {
            if (Test-Path -Path $file)
                {$csv = Import-Csv -Path $file}
            else 
                {Write-Error "The given path for -file '$file' does not exist!"; break}
        }
    return $csv
}

function Save ($output, $path, $mode, $i)
{
    $date = GetDate

    if ($mode -eq 'search')
    {
        if ($path)
        {
            if ($path -like "*.csv") 
                { $output | Sort-Object -Property HostName | export-csv -Path $path -noType }
            else                     
                { $output | Out-File -FilePath $path }
        }   
        elseif ($i) #interactive mode
        {            
            $svinput = Read-Host "Do you wish to save these results to a csv, txt, or other file format?(y/n)"
            ""
            if ($svinput -like "y*")
            {
            $path = Read-Host "Please specify file path you wish to save to (e.g. C:\Users\$env:UserName\Desktop\PCInfo_$date.csv)" 
                if ($path -like "*.csv")
                    {
                        $output | Sort-Object -Property HostName | export-csv -Path $path -noType
                    }
                else
                    {   $output | Sort-Object -Property HostName | Format-List | Out-File -FilePath $path}
            }
        }
    }
    elseif ($mode -eq 'update')
    {
        if ($path -eq '')
        {
            $default_path = "PCInfo_$date.csv"
            $path = $default_path
        }

        if ($path -like "*.csv") 
            {$output | Sort-Object -Property HostName | export-csv -Path $path -noType}
        else                     
            {$output | Sort-Object -Property HostName | Format-List | Out-File -FilePath $path}
    }

    if ($path -or ($mode -eq 'update'))
        { Write-Output "`nSaving file to $path..." }
}

#endregion

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region Printing & Logging Functions

function Print($string,$q,$qq,$qqq)
{
    if (!($q) -and !($qq) -and !($qqq)) 
    {
        Write-Host ($string)
    }
}

function Print_Error($string)
{
    Print -string "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!                                                                                               !!
!!  $string
!!                                                                                               !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -qqq $qqq 
}

function Print_Log($string, $path, [switch]$noNewline)
{   
    #print -string "log path at printlog --> $path"       
    if ($noNewline) { Add-Content -Path $path -Value $string -NoNewline}
    else            { Add-Content -Path $path -Value $string}        
}

function Statistics($added, $newent, $omit, $updated, $execTime, $log)
{
        #print -string "log path at statistics --> $log"
        #print newlines before showing statistics
        $str = "`n`n"
        Print -string $str

        #log the date and time 
        $str = Get-Date
        Print_log -string "When: $str" -path $log 

        #print and log the duration of update function
        if ($execTime -ne '')
        {
            $str = "Execution Time: $execTime"
            Print -string $str
            Print_log -string $str -path $log
        }

        #print and log list of hostnames that no longer have empty info
        if ($newent -eq '')
        {
            $str = "New info found for: NONE"
        }
        else
        {
            $str = "New info found for: $newent"
        }
        Print -string $str
        Print_log -string $str -path $log

        #print and log list of hostnames that were just added to the table
        if ($added -ne '')
        {
            $str = "Added to table: $added"
        }
        else
        {
            $str = "Added to table: NONE"
        }
        Print -string $str
        Print_log -string $str -path $log

        #print and log list of hostnames that were excluded from output table due they no longer exist in AD
        if ($omit -ne '')
        {
            $str= "Excluded from table (was deleted from AD): $omit"
        } 
        else
        {
            $str= "Excluded from table (was deleted from AD): NONE"
        } 
        Print -string $str
        Print_log -string $str -path $log

        #print and log count of computers that were updated
        if ($updated -ne '')
        {
            $str = "Updated: $updated records"
            Print -string $str
            Print_log -string $str -path $log
        } 
}

#endregion

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region WMI classes 
# Objective: return all properties info from the class

function Win32_BaseBoard($pc)
{
    $bb = Get-WmiObject win32_BaseBoard -ComputerName $pc
    return $bb
}

function Win32_Bios($pc)
{
    $bios = Get-WmiObject win32_Bios -ComputerName $pc 
    return $bios
}

function Win32_ComputerSystem($pc)
{
    $compSys = Get-WmiObject win32_ComputerSystem -ComputerName $pc
    return $compSys
}

function Win32_DiskDrive($pc)
{
    $diskdrive = get-wmiobject win32_DiskDrive -ComputerName $pc
    return $diskdrive
}

function Win32_LogicalDisk ($pc)
{
    $ld = Get-WmiObject Win32_LogicalDisk -ComputerName $pc 
    return $ld
}

function Win32_VideoController($pc)
{
    $vc = get-wmiobject win32_VideoController -ComputerName $pc 
    return $vc
}

function Win32_PhysicalMemory($pc)
{
    $pm = get-wmiobject Win32_PhysicalMemory -ComputerName $pc
    return $pm 
}

function Win32_Processor($pc)
{
    $cpu = Get-WmiObject win32_Processor -ComputerName $pc 
    return $cpu
}

function Win32_NetworkAdapterConfiguration($pc)
{
    $nac = get-wmiobject Win32_NetworkAdapterConfiguration -ComputerName $pc
    return $nac
}

#endregion

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region Main Functions

function Search ($computer,$csvFile,$table)
{
    $nf= '';
    $i = 0;
    $found = $false
    $computer = $computer -split '\s'
    $computer = $computer.Trim()
    $csvFile = $csvFile | where {($_.Hostname -ne "") -and ($_.Hostname -ne $null)}

    foreach($name in $computer)
    {
        $found = $false
        foreach($line in $csvfile)
        {
            $pc = $line.HostName;
            if($name -like $pc)
            {
                    $found = $true
                    $row = $table.NewRow();
                    $row.HostName = $pc;   
                    $row.ModelNumber = $line.ModelNumber;    
                    $row.OSversionModel = $line.OSversionModel; 
                    $row.ImageVersion = $line.ImageVersion; 
                    $row.Department = $line.Department;       
                    $row.RAM = $line.RAM;
                    $row.HDmodel = $line.HDmodel;
                    $row.HDsize = $line.HDsize;
                    $row.CPU = $line.CPU;
                    $row.LastLogonDate = $line.LastLogonDate;       
                    $row.OSversionNum = $line.OSversionNum;
                    $row.IPaddress = $line.IPaddress;
                    $row.AD_IPaddress = $line.AD_IPaddress;
                    $row.Location = $line.Location;
                    $row.Manufacturer = $line.Manufacturer;       
                    $row.SerialNumber = $line.SerialNumber;  
                    $row.BIOSversion = $line.BIOSversion;
                    $row.BIOSdate = $line.BIOSdate;
                    $row.BaseBoard = $line.BaseBoard;
                    $row.VideoCard = $line.VideoCard;
                    $row.FreeSpace = $line.FreeSpace;
                    $row.MemoryNumber = $line.MemoryNumber;
                    $row.PatchVersion = $line.PatchVersion;
                    $row.LastUpdate = $line.LastUpdate;        
                    $table.Rows.add($row);              
            }            
        }
        if (!($found))
        {
            Print -string "`n$name NOT FOUND!"
            $nf += "$name "
        }   
        $i++;
    }    

    if($nf -ne '')
    {  Print -string "The following hosts are not in Active Directory: $nf `n" }

    return $table
}

function Update($csvFile,$table,$log)
{   
    #start stopwatch
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()    

    #initialize variables used for statistics
    $newent = ''; $added = ''; $omit = ''; $i = 0; $updated = 1;

    #initial message
    Print -string "`nThis update may take more than 30 minutes to complete (around 12 minutes with fast connectivity) `nPress Ctrl + C to quit update...`n"
    
    #create new hosts list     
    [System.Collections.ArrayList]$ADhosts = FilterHostnames
    $str = $ADhosts.GetType().Name; print -string $str
    $cnt = $ADhosts.Count; Print -string "AD hostlist count: $cnt"


    #remove blank lines from csv file
    #$csvcnt = 0; foreach($m in $csvFile) {$csvcnt++}; print -string $csvcnt
    $csvFile = $csvFile | where {($_.Hostname -ne '') -and ($_.Hostname -ne $null)}
    #$csvcnt = 0; foreach($m in $csvFile) {$csvcnt++}; print -string $csvcnt

    #counter for loop
    $i = 0

    #loop through each row in the imported table
    foreach($line in $csvFile)
    {
        #initialize vars outside 2nd loop
        $connected = $false 
        $matchedName = ""
        $name = ''            
        $empty = 0;   
        $inAD = 0;    

        #store row property 'hostname' in var
        $pc = $line.Hostname;

        #increment count at the begining so we can start with 1 and print hostname for the row we're trying to update
        $i++;
        $str = "`n$i $pc" 
        Print -string $str            
                
        #loop through each name in AD hosts list
        foreach($name in $ADhosts)
        { 
            #compare that hostname in imported row matches hostname in AD hosts list
            if ($pc.Equals($name))
            {
                #save name if it matches
                $matchedName = $name
                #signals that hostname was found in list
                $inAD = 1;
                #create new row
                $row = $table.NewRow();

                #gather AD info
                $ad = AD $pc
                #assign AD info to fields
                $row.HostName           = $pc;   
                $row.AD_IPaddress       = $ad.IPv4Address
                $row.AD_Path            = $ad.CanonicalName                  
                $row.LastLogonDate      = $ad.LastLogonDate
                $row.LastDomainJoinDate = $ad.Created
                $row.OSversionModel     = $ad.OperatingSystem
                $row.Department         = Department     -string     $ad.CanonicalName #gets the path in AD where host is located                 
                $row.OSversionNum       = OSversionModel -osvermodel $ad.OperatingSystemVersion         
                $row.Location           = Location       -string     $ad.CanonicalName
                         
                
                #see if computer is online so it can gather system data
                if (test-Connection -ComputerName $pc -count 1 -quiet)
                {
                    #signal that row fields for system info were previously empty 
                    if([string]::IsNullOrEmpty($line.ModelNumber))
                        {$empty = 1;}                                              
                    
                    Print -string "Trying to find new info for $pc..."

                    $Error.Clear() #clear $error variable to output last error
                    try
                        {$row.ModelNumber = Get-WmiObject win32_computerSystem -ComputerName $pc |select -ExpandProperty Model -First 1} #attempt to get a wmi object
                    catch
                        {
                            print -string "no es facil";
                            print -string $Error; #print last error
                        } 
                        
                    if ($Error.Count -lt 1)
                    {
                        #get info by wmi class
                        $bios             = Win32_Bios $pc
                        $computerSystem   = Win32_ComputerSystem $pc
                        $diskDrive        = Win32_DiskDrive $pc
                        $baseBoard        = Win32_BaseBoard $pc
                        $videoController  = Win32_VideoController $pc
                        $logicalDisk      = Win32_LogicalDisk $pc
                        $physicalMemory   = Win32_PhysicalMemory $pc
                        $netAdaConfig     = Win32_NetworkAdapterConfiguration $pc
                        $processor        = win32_Processor $pc

                        #gather info to input in each attribute
                        $row.ImageVersion = ImageVersion $pc #looks in file directory   
                        $row.RAM          = RAM -RAM $computerSystem.TotalPhysicalMemory                   
                        $row.HDmodel      = HDmodel -diskDrive $diskDrive     
                        $row.HDsize       = HDsize -diskDrive $diskDrive     
                        $row.CPU          = $processor.name
                        $row.IPaddress    = IP -ipAddress $netAdaConfig.ipAddress
                        $row.Manufacturer = $computerSystem.Manufacturer
                        $row.SerialNumber = $bios.serialnumber
                        $row.BIOSversion  = $bios.SMBIOSBIOSVersion  
                        $row.BIOSdate     = $bios.releasedate.Substring(0,8) 
                        $row.BaseBoard    = $baseBoard.product 
                        $row.VideoCard    = VideoCard -videoController $videoController
                        $row.FreeSpace    = FreeSpace -freeSpace $logicalDisk.freeSpace 
                        $row.MemoryNumber = MemoryNumber -partNumber $physicalMemory.partNumber
                        $row.PatchVersion = PatchVersion $pc;
                        $row.LastUpdate   = Get-Date -Format "MM/dd/yyyy hh:mm" 
                    }
                }
                else
                    {Print -string "$pc not online"}

                if([string]::IsNullOrEmpty($row.ModelNumber))
                {
                    "Could not find new info.`n"
                    $row.HostName        = $pc;   
                    $row.ModelNumber     = $line.ModelNumber;    
                    $row.ImageVersion    = $line.ImageVersion;  
                    $row.RAM             = $line.RAM;
                    $row.HDmodel         = $line.HDmodel;
                    $row.HDsize          = $line.HDsize;
                    $row.CPU             = $line.CPU;       
                    $row.IPaddress       = $line.IPaddress;
                    $row.Manufacturer    = $line.Manufacturer;       
                    $row.SerialNumber    = $line.SerialNumber;
                    $row.BIOSversion     = $line.BIOSversion;
                    $row.BIOSdate        = $line.BIOSdate;
                    $row.VideoCard       = $line.VideoCard;
                    $row.BaseBoard       = $line.BaseBoard;            
                    $row.FreeSpace       = $line.FreeSpace;
                    $row.MemoryNumber    = $line.MemoryNumber;
                    $row.PatchVersion    = $line.PatchVersion;
                    $row.LastUpdate      = $line.LastUpdate;
                }
                else
                {
                    $updated += 1;                    
                    if($empty -eq 1)
                    {                        
                        $newent += "$pc "
                        Print -string "FOUND NEW INFO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                    }
                    else
                        {Print -string "INFO UPDATED!!!"}
                }
                #add row
                $table.Rows.add($row)                   
            }      
        }
        #print -string "name to remove: $matchedName"
        #remove host from list because it matched a host in imported table
        $ADhosts.Remove($matchedName);         

        #If host wasn't found in AD do not include in output        
        if($inAD -eq 0) 
        {
            $omit += "$pc ";
            Print -string "Not including hostname '$pc' in table because it does not exist in AD."
        }
    }
    
    $cnt = $ADhosts.Count; Print -string "`nadhost cnt: $cnt"

    # check if there's a host in AD that is not on the table. If so add a row to the table.
    foreach($pc in $ADhosts) 
    {
        print -string "`nAdding new host '$pc'..."   

        # keep count of added fields            
        $added += "$pc ";

        #create a new row
        $row = $table.NewRow()
                
        #gather AD info
        $ad = AD $pc

        #assign AD info to fields
        $row.HostName           = $pc;   
        $row.AD_IPaddress       = $ad.IPv4Address
        $row.AD_Path            = $ad.CanonicalName                  
        $row.LastLogonDate      = $ad.LastLogonDate
        $row.LastDomainJoinDate = $ad.Created
        $row.OSversionModel     = $ad.OperatingSystem
        $row.Department         = Department     -string     $ad.CanonicalName #gets the path in AD where host is located                 
        $row.OSversionNum       = OSversionModel -osvermodel $ad.OperatingSystemVersion         
        $row.Location           = Location       -string     $ad.CanonicalName     

        if (test-Connection -ComputerName $pc -count 1 -quiet) #check that host is online
        {
            $Error.Clear() #clear $error variable to output last error

            try  #attempt to get a wmi object
                {$row.ModelNumber = Get-WmiObject win32_computerSystem -ComputerName $pc |select -ExpandProperty Model -First 1} 
            catch
                {print -string "no es facil"
                    print -string "$Error"} #print last error 
                        
            if ($Error.Count -lt 1)
            {
                #get info by wmi class
                $bios             = Win32_Bios $pc
                $computerSystem   = Win32_ComputerSystem $pc
                $diskDrive        = Win32_DiskDrive $pc
                $baseBoard        = Win32_BaseBoard $pc
                $videoController  = Win32_VideoController $pc
                $logicalDisk      = Win32_LogicalDisk $pc
                $physicalMemory   = Win32_PhysicalMemory $pc
                $netAdaConfig     = Win32_NetworkAdapterConfiguration $pc
                $processor        = win32_Processor $pc

                #gather info to input in each attribute
                $row.ImageVersion = ImageVersion $pc #looks in file directory   
                $row.RAM          = RAM -RAM $computerSystem.TotalPhysicalMemory                   
                $row.HDmodel      = HDmodel -diskDrive $diskDrive     
                $row.HDsize       = HDsize -diskDrive $diskDrive     
                $row.CPU          = $processor.name
                $row.IPaddress    = IP -ipAddress $netAdaConfig.ipAddress
                $row.Manufacturer = $computerSystem.Manufacturer
                $row.SerialNumber = $bios.serialnumber
                $row.BIOSversion  = $bios.SMBIOSBIOSVersion  
                $row.BIOSdate     = $bios.releasedate.Substring(0,8) 
                $row.BaseBoard    = $baseBoard.product 
                $row.VideoCard    = VideoCard -videoController $videoController
                $row.FreeSpace    = FreeSpace -freeSpace $logicalDisk.freeSpace 
                $row.MemoryNumber = MemoryNumber -partNumber $physicalMemory.partNumber
                $row.PatchVersion = PatchVersion $pc;
                $row.LastUpdate   = Get-Date -Format "MM/dd/yyyy HH:mm"
            }
        }            
        $table.Rows.add($row)            
    }
    #do not include any empty lines and sort table by hostname
    $table = $table | Sort-Object -Property hostname

    #display the output in table format                 
    $table | Format-Table -AutoSize

    #Get the time it took update to finish
    $execTime = GetDuration($stopwatch)

    #Call Statistics to print and log stats of the update
    Statistics -added $added -newent $newent -omit $omit -updated $updated -execTime $execTime -log $log

    #return output table
    return $table 
}

#endregion

#________________________________________________________________________________________________________________________________________________________________________________________________________
#region Control Section

$again = 0;
do 
{
    $table = CreateTable 
    if ($search)
    {    
        $mode = "search";
        $computer = GetHostname -computer $computer -i $i
        $csvFile = Import_Csv -file $file
        $table = Search -csvFile $csvFile -computer $computer -table $table  
        $table
        Save -output $table -path $save -mode $mode -i $i
    }
    elseif ($update)
    {      
        $log = CreateLog -log $log
        $mode = "update";         
        $csvFile = Import_Csv -file $file
        $table = Update -csvFile $csvFile -table $table -log $log  
        #$table | Measure-Object -Line
        $table = $table | where {($_.Hostname -ne '') -and ($_.Hostname -ne $null)} 
        #$table | Measure-Object -Line
        Save -output $table -path $save -mode $mode -i $i      
    }
    else
        {Help; break}

    if (($mode -eq "search") -and ($i))
    {   
        $again = Read-Host "`nWould you like to perform another $action ?(y/n)"
        ""
        if  ($again -like "y*") {$again = 1}
    }
}
while (($again -eq 1) -and ($mode -eq "search") -and ($i))

#endregion


