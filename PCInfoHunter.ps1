#!/usr/bin/env powershell

#PCInfoHunter
#######################

<#

.SYNOPSIS
This is a tool with the purpose of creating, updating, and querying a database of system info, AD info, 
and custom info for hosts in the domain's network.

.DESCRIPTION
    This script is composed of two main functions: update and search which call other function to gather information.
The update function updates the fields of a table or adds such fields to the output table for each host in Active Directory.
The search function pulls information from the input database.
    The following are some important remarks about the switches:
        -update and -search update or query content from a csv file
        -save saves the full results to a user given csv or text file
        -save is optional
        -hostlist is used to pass a list of hostnames that will be searched
        -hostlist is optional but must be used in combination with -search        
        -hostlist requires a text file containing hostnames separated by new lines

Options:
            -search         Search data for specified hostnames
            -update         Changes table values by updating fields for each host in AD 
            -hostlist       To input a text file with a list of hosts to be searched.
                            It must be used in combination with search function.
            -save           Save table to specified csv or txt file

Usage:  pcinfohunter [-search] [-update] [-hostlist example_list.txt] [-save results.csv]
         

.EXAMPLE
.\PCInfoHunter.ps1 -update -save C:\Users\Current_User\Desktop\results.csv

.EXAMPLE
.\PCInfoHunter.ps1 -search -hostlist hosts.txt -save C:\Users\Current_User\Desktop\results.csv

.NOTES

.LINK

#>

#script parameters
param([Switch]$update, [Switch]$search, [String]$hostlist, [String]$save)

#global variables
$h = 0

# Functions to gather info
function location #function to get location from name of OU in AD
{
    param([string]$computer)
    #gets the path in AD where host is located
    $stringAD = Get-ADComputer $computer -Properties CanonicalName | FT CanonicalName -HideTableHeaders | Out-String
    #splits the path string to match with possible department
    $locstring = $stringAD -split '/' -split'-'| Select-String -Pattern #name of OUs in AD
    $location= $locstring | Select -First 1
    if ($location -eq $null)
        {$location = ''}
    return $location
}

function department #function to get department from name of OU in AD
{
    param([string]$computer)
    #gets the path in AD where host is located
    $stringAD = Get-ADComputer $computer -Properties CanonicalName | FT CanonicalName -HideTableHeaders | Out-String
    #splits the path string to match with possible department
    $department = $stringAD -split '/' | Select-String -Pattern 'Finance', 'Marketing', #etc
    if ($department -eq $null)
        {
        $depstring = $stringAD -split '/' -split'-'| Select-String -Pattern #name of OUs in AD
        $department = $depstring | Select -First 1
        }

    return $department
} 
 
function imageVersion
{
    param([string]$computer)  
    
        if (Test-Path #path of file with version number)
        {
         $imgVersion = Get-Content -Path #path of file with version number
        }
        else
        { 
          $imgVersion =  "Not Found"
        }            
    return $imgVersion;    
}

 function PatchVersion
{
    param([string]$computer)  
    
    if (Test-Path #path of file with path version)
    {
        $PatchVersion = Get-Content -Path \\$computer\C$\Custom\Docs\patch_version.txt    
    }
    else
    { 
        $PatchVersion = "Not Found"
    }            
    return $PatchVersion;    
}

function OSversionNum
{
     param([string]$computer)
     $version = Get-ADComputer $computer -Properties OperatingSystemVersion | select -ExpandProperty OperatingSystemVersion -First 1   
     return $version;
}

function IP
{
    param([string]$computer)
    $IP = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computer |
    Where { $_.IPAddress } |Select -Expand IPAddress -First 1 | Where { $_ -like '133.*' -or $_ -like '33.*'}
    return $IP;
}

function AD_IP
{
    param([string]$computer)
    $IP = Get-ADComputer $computer -Properties IPv4address | select -ExpandProperty IPv4address -First 1 
    return $IP
}

 function OSversionModel
 {
     param([String]$computer)
     $versionModel = Get-ADComputer $computer -Properties OperatingSystem | select -ExpandProperty OperatingSystem -First 1 
     $OSversionModel= $versionModel -replace 'Windows','Win' -replace 'Enterprise','Ent' -replace 'Professional','Pro' 
     return $OSversionModel
 }

 function BIOSdate 
 {
    param([String]$computer)
    $biosdate = Get-WmiObject win32_bios -ComputerName $computer | select -ExpandProperty ReleaseDate
    $biosdate = $biosdate.substring(0,8)
    return $biosdate
 }

function RAM
{
    param([string]$computer)
    $RAM = Get-WmiObject win32_computerSystem -ComputerName $computer | select -ExpandProperty TotalPhysicalMemory -First 1
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


function FreeSpace
{
    param([string]$computer)
    $freespace = Get-WmiObject win32_logicalDisk -ComputerName $computer | select -ExpandProperty FreeSpace -First 1
    if ($freespace -eq $null)
    {
        $freespace = '';
    }
    else
    {
        $freespace = $freespace / (1024*1024*1024)
        $freespace = [math]::Round($freespace,0)
    }
    return $freespace;
}

function formatnumby6
{
    param([string]$number)
    $number = $number / (1024*1024)
    $number = [math]::Round($number,0)
    return $number;
}

function formatnumby9
{
    param([string]$number)
    $number = $number / (1024*1024*1024)
    $number = [math]::Round($number,0)
    return $number;
}

function VideoCard
{ 
   # param([string]$computer)
  #  $videocard = Get-WmiObject win32_videocontroller -ComputerName $computer | select -ExpandProperty name -Skip 1 -First 1
   # if ($videocard -eq $null) 
  #  {
   #     $videocard = Get-WmiObject win32_videocontroller -ComputerName $computer | select -ExpandProperty name -First 1
   # }
    $videocard = $null;
    for($i = 0; $i -le 8; $i++)
    {
        if ((get-wmiobject win32_videocontroller -ComputerName $computer | select -ExpandProperty adapterdactype -Skip $i -First 1) -ne $null)
        {
            $videocard = Get-WmiObject win32_videocontroller -ComputerName $computer | select -ExpandProperty name -Skip $i -First 1
            break;            
        }
    }
    return $videocard;
}

function HDsize
{
    param([string]$computer)
    $hdsize = "Not Found"
    for($i = 0; $i -le 5; $i++)
    {
        if ((get-wmiobject win32_diskdrive -ComputerName $computer | select -ExpandProperty interfacetype -Skip $i -First 1) -ne 'USB')
        {
            $HDsize = Get-WmiObject win32_DiskDrive -ComputerName $computer | select -ExpandProperty Size -Skip $i -First 1
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

function HDmodel
{
    param([string]$computer)
    $hdmodel = "Not Found"
    for($i = 0; $i -le 5; $i++)
    {
        if ((get-wmiobject win32_diskdrive -ComputerName $computer | select -ExpandProperty interfacetype -Skip $i -First 1) -ne 'USB')
        {
            $hdmodel = Get-WmiObject win32_DiskDrive -ComputerName $computer | select -ExpandProperty Model -Skip $i -First 1
            break;            
        }
    }
    return $hdmodel;
}

function MemoryNumber
{
    param([string]$computer)
    for($i = 0; $i -le 4; $i++)
    {
        if ((Get-WmiObject win32_PhysicalMemory -ComputerName $computer | select -ExpandProperty PartNumber -Skip $i -First 1) -ne $null)
        {
            $memNum = Get-WmiObject win32_PhysicalMemory -ComputerName $computer | select -ExpandProperty PartNumber -Skip $i -First 1 
            break;            
        }
    }
    return $memNum;
}



#----------------------------------------------------#
# Creating a table for output
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
$col10 = New-Object System.Data.DataColumn LastLogonDate,([string]);
$col11 = New-Object System.Data.DataColumn OSversionNum,([string]);
$col12 = New-Object System.Data.DataColumn IPaddress,([string]);
$col13 = New-Object System.Data.DataColumn AD_IPaddress,([string]);
$col14 = New-Object System.Data.DataColumn Location,([string]);
$col15 = New-Object System.Data.DataColumn Manufacturer,([string]);
$col16 = New-Object System.Data.DataColumn SerialNumber,([string]);
$col17 = New-Object System.Data.DataColumn BIOSversion,([string]);
$col18 = New-Object System.Data.DataColumn BIOSdate,([string]);
$col19 = New-Object System.Data.DataColumn BaseBoard,([string]);
$col20 = New-Object System.Data.DataColumn VideoCard,([string]); 
$col21 = New-Object System.Data.DataColumn FreeSpace,([string]);
$col22 = New-Object System.Data.DataColumn MemoryNumber,([string]);
$col23 = New-Object System.Data.DataColumn PatchVersion,([string]);
$col24 = New-Object System.Data.DataColumn LastUpdate,([string]);

#Add the Columns
$table.columns.add($col1);
$table.Columns.add($col2);
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

#------------------------------------------------#
# Function to update given database (also removes and add users based on AD)
 function update
{
    param($csvFile)
    #stat stopwatch
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()
    $hosts = Get-ADComputer -Filter 'Name -like "*"' | Sort-Object | FT Name -HideTableHeaders
    
    $strhosts = $hosts | Out-String
    $strhostsC = $strhosts -split "\n"
    $allhosts = @();
    foreach($each in $strhostsC)
    {
        $nameNS = $each -replace '\s',''     
        if ($nameNS -ne '')
        {
            $os = Get-ADComputer $nameNS -Properties OperatingSystem | select -ExpandProperty OperatingSystem -First 1 
            $allhosts += $nameNS
	}
    }
    $newent = ''
    $added = ''
    $omit = ''
    $i = 0; 
    $updated = 1;
    foreach($line in $csvFile)
    {
        $connected = $false     
        $empty = 0;   
        $inAD = 0;
        $i = $i + 1;
        $computer = $line.Hostname;
        write-output "$i $computer"

        foreach($name in $allhosts)
        {
            
           if($computer.Equals($name))
           {
                $inAD = 1;
           }
        }
        if ($inAD -eq 1)
        {
            $row = $table.NewRow();
            $row.HostName = $computer;     
            $row.OSversionModel = OSversionModel $computer;  
            $row.Department = department $computer;          
            $row.LastLogonDate = Get-ADComputer $computer -Properties LastLogonDate | select -ExpandProperty LastLogonDate -First 1         
            $row.OSversionNum = OSversionNum $computer;          
            $row.AD_IPaddress = AD_IP $computer;
            $row.Location = location $computer;
           
                
            if (test-Connection -ComputerName $computer -count 1 -quiet)
            {
                if([string]::IsNullOrEmpty($line.ModelNumber))
                {  
                    $empty = 1;
                }                      
                #gather info to input in each attribute
                Write-Output "Trying to find new info for $computer..."
                $row.ModelNumber = Get-WmiObject win32_computerSystem -ComputerName $computer |select -ExpandProperty Model -First 1   
                $row.ImageVersion = imageVersion $computer;     
                $row.RAM = RAM $computer;                
                $row.HDmodel = HDmodel $computer;       
                $row.HDsize = HDsize $computer $index;
                $row.CPU = Get-WmiObject win32_Processor -ComputerName $computer |select -ExpandProperty Name -First 1
                $row.IPaddress = IP $computer; 
                $row.Manufacturer = Get-WmiObject win32_computerSystem -ComputerName $computer |select -ExpandProperty Manufacturer -First 1     
                $row.SerialNumber = Get-WmiObject win32_bios -ComputerName $computer |select -ExpandProperty SerialNumber -First 1 
                $row.BIOSversion = Get-WmiObject win32_bios -ComputerName $computer | select -ExpandProperty SMBIOSBIOSVersion 
                $row.BIOSdate = BIOSdate $computer; 
                $row.BaseBoard = Get-WmiObject win32_BaseBoard -ComputerName $computer | select -ExpandProperty Product 
                $row.VideoCard = VideoCard $computer; 
                $row.FreeSpace = FreeSpace $computer; 
                $row.MemoryNumber = MemoryNumber $computer;
                $row.PatchVersion = PatchVersion $computer;
            }
            else
            {
                write-output "$computer not online"
            }

             if([string]::IsNullOrEmpty($row.ModelNumber))
             {
                "Could not find new info.`n"
                $row.HostName = $computer;   
                $row.ModelNumber = $line.ModelNumber;    
                $row.ImageVersion = $line.ImageVersion;  
                #$number = $line.RAM;
                #if ($number -ne '')    
                #{$row.RAM = formatnumby6 $number;}
                $row.RAM = $line.RAM;
                $row.HDmodel = $line.HDmodel;
                #$number = $line.HDsize;
               # if ($number -ne '')
               # {$row.HDsize = formatnumby9 $number;}
                $row.HDsize = $line.HDsize;
                $row.CPU = $line.CPU;       
                $row.IPaddress = $line.IPaddress;
                $row.Manufacturer = $line.Manufacturer;       
                $row.SerialNumber = $line.SerialNumber;
                $row.BIOSversion = $line.BIOSversion;
                $row.BIOSdate = $line.BIOSdate;
                $row.VideoCard = $line.VideoCard;
                $row.BaseBoard = $line.BaseBoard;            
                #$number = $line.FreeSpace;
                #if ($number -ne '')
               # {$row.FreeSpace = formatnumby9 $number;}
                $row.FreeSpace = $line.FreeSpace;
                $row.MemoryNumber = $line.MemoryNumber;
                $row.PatchVersion = $line.PatchVersion;
                $row.LastUpdate = $line.LastUpdate;
             }
             else
             {
             `  $updated += 1;
                $row.LastUpdate = Get-Date -Format "MM/dd/yyyy hh:mm" 
                if($empty -eq 1)
                {                        
                    $newent += "$computer "
                    "FOUND NEW INFO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!`n"
                }
                else
                {
                "INFO UPDATED!!!`n"
                }
             }                     
        }        
        else
        {
            $omit += "$computer ";
            Write-Output "Not including hostname '$computer' in table because it does not exist in AD.`n"
        }
        $table.Rows.add($row);
    }
        #$table | Format-Table -AutoSize
        $name = '';
        foreach($name in $allhosts)
        {
            $intable = 0;
            $computer = $name
           # write-output "Checking that $computer is in table"
            for ($i=0; $i -lt $table.Rows.Count; $i++)
            {                
               # $place = $table.Rows[$i][0] | Out-String
                if($computer -eq $table.Rows[$i][0])
                {
                     $place = $table.Rows[$i][0] | Out-String                     
                    # Write-Output "$computer already in table`n"
                     $intable = 1;
                }
            }

            if ($intable -eq 0)
            {                
                $added += "$computer ";
                $row = $table.NewRow();

                if (test-Connection -ComputerName $computer -count 1 -quiet)
                {    
                    #gather info to input in each attribute
                    $row.HostName = $computer; 
                    $row.ModelNumber = Get-WmiObject win32_computerSystem -ComputerName $computer |select -ExpandProperty Model -First 1   
                    $row.OSversionModel = OSversionModel $computer;
                    $row.ImageVersion = imageVersion $computer;
                    $row.Department = department $computer;      
                    $row.RAM = RAM $computer;
                    $row.HDmodel = HDmodel $computer;                    
                    $row.HDsize = HDsize $computer;
                    $row.CPU = Get-WmiObject win32_Processor -ComputerName $computer |select -ExpandProperty Name -First 1
                    $row.LastLogonDate = Get-ADComputer $computer -Properties LastLogonDate | select -ExpandProperty LastLogonDate -First 1        
                    $row.OSversionNum = OSversionNum $computer;
                    $row.IPaddress = gwmi Win32_NetworkAdapterConfiguration -ComputerName $computer |Where { $_.IPAddress } |Select -Expand IPAddress -First 1 | Where { $_ -like '172.*' -or $_ -like '10.*'} 
                    $row.AD_IPaddress = AD_IP $computer;
                    $row.Location = location $computer;
                    $row.Manufacturer = Get-WmiObject win32_computerSystem -ComputerName $computer |select -ExpandProperty Manufacturer -First 1     
                    $row.SerialNumber = Get-WmiObject win32_bios -ComputerName $computer |select -ExpandProperty SerialNumber -First 1   
                    $row.BIOSversion = Get-WmiObject win32_bios -ComputerName $computer | select -ExpandProperty SMBIOSBIOSVersion 
                    $row.BIOSdate = Get-WmiObject win32_bios -ComputerName $computer | select -ExpandProperty ReleaseDate 
                    $row.BaseBoard = Get-WmiObject win32_BaseBoard -ComputerName $computer | select -ExpandProperty Product 
                    $row.VideoCard = Get-WmiObject win32_videocontroller -ComputerName $computer | select -ExpandProperty name -Skip 1 -First 1 
                    $row.FreeSpace = FreeSpace $computer;
                    $row.MemoryNumber = MemoryNumber $computer;
                    $row.PatchVersion = PatchVersion $computer;
                    if([string]::IsNullOrEmpty($line.ModelNumber))
                    {  
                        $row.LastUpdate = ''; 
                    } 
                    else
                    {
                        $row.LastUpdate = Get-Date -Format "MM/dd/yyyy hh:mm"                                         
                    }
                }
            else
                {                
                $row.HostName = $computer;          
                $row.OSversionModel = OSversionModel $computer;  
                $row.Department = department $computer; 
                $row.LastLogonDate = Get-ADComputer $computer -Properties LastLogonDate | select -ExpandProperty LastLogonDate -First 1         
                $row.OSversionNum = OSversionNum $computer;
                $row.AD_IPaddress = AD_IP;
                $row.Location = location $computer;
                }
                $table.Rows.add($row);
            }               
        }         
    $table | Format-Table -AutoSize
    $stopwatch.stop();
    $dwatch = $stopwatch.Elapsed.Minutes | Out-String 
    $dwatch += " minutes " 
    $sec = $stopwatch.Elapsed.Seconds | Out-String
    $dwatch += $sec
    $dwatch += " seconds" 
    $dwatch -replace "\n", ""
    ""
    if ($added -ne '')
    {
        Write-Output "The following host were added to the table: $added`n"
    }
    if ($newent -eq '')
    {
        "No new info was found"
    }
    else
    {
        Write-Output "New info found for hosts: $newent`n"
    }
    write-output "Updated: $updated`n"
    if ($omit -ne '')
    {
        Write-Output "The following host were not included in the database because they're no longer on AD: $omit`n"
    }    
}

#---------------------------------------------#
# Function to search certain users in given database
function search ($csvFile, $hostlist)
{
    #start stopwatch
    $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
    $nf= '';
    $i = 0;
    $found = $false
    $hostlist = $hostlist -split '\s'
    foreach($name in $hostlist)
    {
        $found = $false
        foreach($line in $csvfile)
        {
            $computer = $line.HostName;
            if($name -like $computer)
            {
                    $found = $true
                    write-output "$i $computer"
                    $row = $table.NewRow();
                    $row.HostName = $computer;   
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
          Write-Output "$i $name NOT FOUND!"
          $nf += "$name "
      }   
     $i++;
   }

#$table | Format-Table -AutoSize
$table
$stopwatch.stop();
$dwatch = $stopwatch.Elapsed.Minutes | Out-String 
$dwatch += " minutes " 
$sec = $stopwatch.Elapsed.Seconds | Out-String
$dwatch += $sec
$dwatch += " seconds" 
$dwatch -replace "\n", ""
""
if($nf -ne '')
{
    Write-Output "The following hosts are not in Active Directory: $nf `n"
}
}

#-------------------------------------------------#
# Handling user arguments
$action = "none";
$again = 0;
do {
    if ($update)
    {
        $action = "update";
        $csvFile = Import-Csv Path
        update $csvFile
    }
    elseif ($search)
    {    
        $action = "search";
        if ($hostlist -ne '')
        {      
            if ($hostlist -like "*.txt")
            {
                Write-Output "Input host list received: $hostlist"
                $hostlist = Get-Content -Path $hostlist;
            }
            else
            {
                Write-Output "ERROR! the input for 'hostlist' must be a txt file"
            }
        }
        else
        {
             $stxWrong = 0;
             $userInput = Read-host "Enter hostname of computers"
             if ($userInput -match ',' -or $userInput -match '  ' -or $userInput -match ';' -or $userInput -match ':')
             {
                $stxWrong = 1;
             }
             while($stxWrong -eq 1)
             {
                 Write-Host "Please divide with single spaces between hosts. Do not use any other delimiter such as commas, colons, semicolons, etc."
                 $userInput = Read-host "Enter hostname of computers"
                 if ($userInput -match ',' -or $userInput -match '  ' -or $userInput -match ';' -or $userInput -match ':')
                     {$stxWrong = 1;}
                 else
                     {$stxWrong =e 0;}
             }
             $hostlist = $userInput 
        }
        $csvFile = Import-Csv Path
        search $csvFile $hostlist
        $hostlist = $null;
    }
    else
    {
        $h = 1
        Get-Help .\PCInfoHunter.ps1
    }

    #------------------------------------------------------------#
    #Saving Section
    if ($save -ne '')
    {    
        if ($save -like "*.csv")
        {    
            $table | Sort-Object -Property HostName | export-csv -LiteralPath $save -noType 
        }
        else
        {    
            $table | Sort-Object -Property HostName | Out-File -FilePath $save 
        }
    }
    else
    {
        if($h -ne 1)
        {
            $svinput = Read-Host "Do you wish to save these results to a csv, txt, or other file format?(y/n)"
            ""
            if ($svinput -like "y*")
            {
            $save = Read-Host "Please specify file path you wish to save to (e.g. C:\Users\$env:UserName\Desktop\results.csv)" 
                if ($save -like "*.csv")
                    {
                        $table | Sort-Object -Property HostName | export-csv -LiteralPath $save -noType
                    }
                else
                    {
                        $table | Sort-Object -Property HostName | Format-Table | Out-File -FilePath $save
                    }
            }
        }
    }
$table.Clear();
if ($action -eq "search")
{
    ""
    $again = Read-Host "Would you like to perform another $action ?(y/n)"
    ""
}
if  ($again -like "y*")
    {
        $again = 1;
    }
else
    {
        $again = 0;
    }

}
while (($again -eq 1) -and ($action -eq "search"));
#EOF
$action = $null;