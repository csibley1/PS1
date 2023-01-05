Import-Module activedirectory
function convertFrom-base36 # Source: Mirko Schnellbach, reposted on https://ss64.com/ps/syntax-base36.html
{
    [CmdletBinding()]
    param ([parameter(valuefrompipeline=$true, HelpMessage="Alphadecimal string to convert")][string]$base36Num="")
    $alphabet = "0123456789abcdefghijklmnopqrstuvwxyz"
    $inputarray = $base36Num.tolower().tochararray()
    [array]::reverse($inputarray)
    [long]$decNum=0
    $pos=0

    foreach ($c in $inputarray)
    {
        $decNum += $alphabet.IndexOf($c) * [long][Math]::Pow(36, $pos)
        $pos++
    }
    $decNum
}

function get-Computerlist
{
# Check if Excel is already running
write-host "Debug 1 - check if Excel is running"
$doNotKill = (get-process excel -erroraction SilentlyContinue).id

# Open the workbook, extract tabs to CSV. "Debug" lines show where in the process it is while loading.
write-host "Debug 2 - get computer list path"
$excelFile = (get-item 'C:\Users\csibley\CONTOSO.onmicrosoft.com\IT - Inventory\Computer List.xlsx' | Select-Object fullname).fullname
write-host "Debug 3 - get desktop path"
$csvPath = (get-item ~\desktop | Select-Object fullname).fullname+'\'
write-host "Debug 4 - start Excel"
$Excel = New-Object -ComObject Excel.Application
write-host "Debug 5 - Excel.visible = false"
$Excel.Visible = $false
write-host "Debug 6 - Excel.displayalerts = false"
$Excel.DisplayAlerts = $false
write-host "Debug 7 - load computer list into Excel"
$wb = $Excel.Workbooks.Open($excelFile)
write-host "Debug 8 - begin extracting worksheets"
foreach ($ws in $wb.Worksheets)
{
    $sname = $ws.name
    $ws.SaveAs("$csvpath$sname.csv", 6)
    $script:csvfile += [System.Collections.ArrayList]@("$csvpath$sname.csv")
}

# Close Excel
write-host "Debug 9 - close Excel"
$Excel.Workbooks.Close()
$Excel.Quit()
(get-process excel).id | foreach-object{if ($_ -ne $doNotKill){$expid = $_}}

# Load CSVs into variables, then clean up
$script:retired = import-csv '~\desktop\retired list.csv'
$script:active = import-csv '~\desktop\computers.csv'
foreach ($csv in $csvfile){remove-item $csv}
stop-process -id $expid
remove-variable -force -name csvfile -scope script
}

get-Computerlist

#Main menu startup
do 
{
    $host.ui.RawUI.WindowTitle = “IT Toolbox”
    Clear-Host
    write-host @"
 __        _____  ____  _  __
 \ \      / / _ \|  _ \| |/ /
  \ \ /\ / / | | | |_) | ' / 
   \ V  V /| |_| |  _ <| . \ 
    \_/\_/  \___/|_| \_\_|\_\
                             
       R  E  L  A  T  E  D         
.....................................

  1. PC Lookup - Computer list
  2. Pw stats
  3. DNS info
  4. PC Lookup - AD
  5. Refresh computer list
  6. Passwords expiring soon
  7. User lookup
  8. Systems below 21H2
  9. Convert svc tag to express code
  10. Get termed users
  11. Kill mmc.exe
  X. Exit
.....................................


"@
    $menu = read-host "Enter menu number"
    switch ($menu){
        1 #### PC Lookup
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - PC Lookup”
            write-host "............. PC Lookup ..............."
            do
            {
                $srch = (read-host "Computer or user name (N to Quit)")+'*'
                if ($srch -eq 'n*' -or $srch -eq '*'){break}
                # Find computers & their users in the Active list
                foreach ($actv in $active){
                    if ($actv.'Computer Name - Orange= Check O365 ID - no UserName' -like $srch -or $actv.'User Name' -like $srch)
                    {
                        $act = (get-date).ToString()+'  ---ACTIVE---'+'      '+$actv.'Computer Name - Orange= Check O365 ID - no UserName'+'  '+$actv.'Purchase Date'+'  '+$actv.'User Name'
                        $act | out-file ~\desktop\pclookup.txt -append
                        $act
                    }
                }
                # Find computers & their users in the Retired list
                foreach ($retpc in $retired){
                    if ($retpc.'Computer Name' -like $srch -or $retpc.'User Name' -like $srch)
                    {
                        $ret = (get-date).ToString()+'  ***RETIRED**'+'      '+$retpc.'Computer Name'+'  '+$retpc.'Date'+'  '+$retpc.'User Name'
                        $ret | out-file ~\desktop\pclookup.txt -append
                        write-host $ret -ForegroundColor Red
                    }
                }
            }until('')
        }
        2 #### Password stats
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - Password stats”
            write-host "............ PW Stats ................"
            do
            {
                write-host "Search by partial username or last four phone digits."
                $srch = (read-host "Username (Press Enter to quit)")+'*'
                if ($srch.Length -eq '1'){
                break
                }else{
                    Try{
                        if ([int]$srch.Substring(0,4)){$srch = "*"+$srch}
                    }catch{
                        # Not searching by phone, continue...
                    }
                }
                $mobile_slice = 3,3
                $stats = get-aduser -filter {samaccountname -like $srch 
                                             -or mobile -like $srch 
                                             -or telephonenumber -like $srch} `
                                             -properties distinguishedname,samaccountname,title,manager,mobile,telephonenumber,pwdlastset,badpasswordtime,badpwdcount,lockouttime,lastlogontimestamp | `
                                             Select-Object distinguishedname,samaccountname,title,manager,mobile,telephonenumber,`
                                             @{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}},`
                                             @{name ="badPasswordTime";expression={[datetime]::FromFileTime($_.badPasswordTime)}},badpwdcount,`
                                             @{name ="lockoutTime";expression={[datetime]::FromFileTime($_.lockoutTime)}},`
                                             @{name ="lastlogontimestamp";expression={[datetime]::FromFileTime($_.lastlogontimestamp)}}
                "`n"
                (get-date).DateTime
                foreach ($usrstat in $stats)
                {
                    if ($usrstat)
                    {
                        if ($usrstat.distinguishedname -like "*exiting accounts*"){write-host "Terminated" -backgroundcolor Red}
                        $usrstat
                        if ($usrstat.mobile){
                            # Switch statement determines the approximate region of the user. Searches for numbers in the format "+1 6128675309".
                            # If there is no space after +1, $mobile_slice is adjusted accordingly.
                            if ($usrstat.mobile[2] -ne " "){$mobile_slice = 2,3}
                            switch($usrstat.mobile.substring(($mobile_slice -split ",")[0],($mobile_slice -split ",")[1])){
                            {$_ -eq 612 -or $_ -eq 651}
                            {
                                write-host "MSP" -ForegroundColor Green
                            }

                            218
                            {
                                write-host "Duluth" -ForegroundColor Green
                            }

                            507
                            {
                                write-host "Rochester" -ForegroundColor Green
                            }

                            default
                            {
                                write-host "ERROR" -ForegroundColor Yellow
                            }
                            }
                        }else{
                            write-host "Mobile number is blank!" -foregroundcolor Yellow
                        }
                        'PW expires on '+($usrstat.pwdLastSet).adddays(180)+"`n`n"
                        if ($usrstat.lockoutTime -gt (get-date).adddays(-1)){Unlock-ADAccount -Identity $usrstat.samaccountname -confirm}
                    }else{write-host "No user found" -backgroundcolor DarkRed}
                }
            }until('')
        }

        3 #### DNS info
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - DNS info”
            write-host "............. DNS Info ................"
            do
            {
                $srch = (read-host "Computer name (Press Enter to quit)")+'*'
                if ($srch.Length -eq '1'){break}
                Get-DnsServerResourceRecord -zonename "CONTOSO.com" -ComputerName "DNS-SERVER" | Where-Object HostName -like "$srch" | Format-Table
            }until('')
        }

        4 #### Get AD computer & group membership
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - Active Directory PC Lookup”
            write-host ".......... Active Directory PC Lookup ............"
            do
            {
                $srch = (read-host "Computer name")+'*'
                if ($srch.Length -eq 1){
                    break
                }else{    
                    $adcom = get-adcomputer -filter {name -like $srch} -properties name,whencreated,lastlogondate,operatingsystem,operatingsystemversion,distinguishedname | Select-Object name,whencreated,lastlogondate,operatingsystem,operatingsystemversion,distinguishedname
                    if (!$adcom){
                        "Nothing found"
                    }else{
                        $adcom
                        write-host "Groups for "($adcom).name -ForegroundColor Cyan
                        (get-adprincipalgroupmembership ($adcom).distinguishedname).name
                        
                    }
                }
            }until('')
        }


        5 #### Refresh computer list
        {
            $host.ui.RawUI.WindowTitle = “Refreshing...”
            get-Computerlist
        }
        
        6 ### Get passwords expiring within a week
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - Passwords expiring”
            write-host "............. Passwords expiring soon ................"
            $ous = 'OU=BUSINESSUNIT HQ,DC=CONTOSO,DC=com', 'OU=BUSINESSUNIT2,DC=CONTOSO,DC=com', 'OU=BUSINESSUNIT3,DC=CONTOSO,DC=com'
            $ous | ForEach-Object {
                Get-ADUser -searchbase $_ -Filter * -Properties Mobile,PasswordLastSet | Where-Object {$_.PasswordLastSet -lt (Get-Date).AddDays(-174) -and $_.PasswordLastSet -gt (Get-Date).AddDays(-190) } | Select-Object -Property Name, Mobile, PasswordLastSet,@{label="PasswordExpires";Expression={($_.passwordLastSet).adddays(180)}}
                } | Sort-Object PasswordExpires | out-string
            read-host "Press Enter to continue..."
        }

        7 ### Look up a user if all you have is their first name
        {
            $host.ui.RawUI.WindowTitle = “User lookup”
            write-host "............ Full-name user lookup ..........."
            do
            {
                $srch = (read-host "Name: ")+'*'
                if ($srch.Length -eq 1){break}
                else{    
                    $namesearch = get-aduser -filter {name -like $srch} -properties name,samaccountname,distinguishedname,mobile,telephonenumber,passwordlastset | Select-Object distinguishedname,samaccountname,mobile,telephonenumber,passwordlastset
                    if (!$namesearch){
                        "Nothing found"
                    }else{
                    $namesearch
                    }
                }
            }until('')
        }

        8 ### Check for computers older than the specified patch level
        {
            $host.ui.RawUI.WindowTitle = “IT Toolbox - Check Windows patch level”
            $patch = "10.0 (19044)"
            $pcfilter = get-adcomputer -filter {operatingsystemversion -le "10.0 (23000)" -and name -notlike "FAKE*" -and name -notlike "FAKE2*" -and name -notlike "FAKE3*"} -properties name, lastlogondate, operatingsystemversion | Select-Object name,lastlogondate,operatingsystemversion | Sort-Object operatingsystemversion
            $pcfilter | ForEach-Object{
            foreach ($actv in $active)
            {
                # Check the Active computer list for Win10 patch level
                if ($actv.'Computer Name - Orange= Check O365 ID - no UserName' -like $_.name+"*" -and $_.operatingsystemversion -le $patch){
                    if ($actv.'Computer Name - Orange= Check O365 ID - no UserName'.length -gt 7){
                        $actv.'Computer Name - Orange= Check O365 ID - no UserName'.Substring(0,7).insert(7,'... ')+$_.operatingsystemversion+"    "+$actv.'Purchase Date'+"`t"+$actv.'User Name'
                    }else{
                        $actv.'Computer Name - Orange= Check O365 ID - no UserName'+"    "+$_.operatingsystemversion+"    "+$actv.'Purchase Date'+"`t"+$actv.'User Name'
                    }
                }
            }
           
           foreach ($retpc in $retired)
            {
                # Outdated computers that are on the Retired list, but have not been deleted from AD
                if ($retpc.'Computer Name' -like $_.name+"*" -and $_.operatingsystemversion -le $patch){
                    if ($retpc.'Computer Name'.length -gt 7){
                        $retpc.'Computer Name'.substring(0,7).insert(7,'... ')+$_.operatingsystemversion+"    "+$retpc.'Date'+"`t"+$retpc.'User Name'+"  ***RETIRED***"
                    }else{
                        $retpc.'Computer Name'+"    "+$_.operatingsystemversion+"    "+$retpc.'Date'+"`t"+$retpc.'User Name'+"  ***RETIRED***"
                    }
                }
            }
            }
            write-output $pcfilter | Group-Object -property operatingsystemversion | Select-Object name,count | out-host
            write-host "TOTAL ($patch or under) " -foregroundcolor White -nonewline; write-host @($pcfilter | Where-Object {$_.operatingsystemversion -le $patch}).count -ForegroundColor Cyan
            read-host "Press Enter to continue..."
        }
        
        9 ### Convert service tag to express svc code
        {
            do
            {
                $text = read-host "Enter tag"
                convertFrom-base36 $text
            }until ($text -eq '')
        }

        10 ### Get all terminated users in AD
        {
            get-aduser -filter * -searchbase "OU=TERMINATIONS,DC=CONTOSO,DC=com" | Select-Object name | Sort-Object name | Out-String
            read-host "Press Enter to continue..."
        }

        11 ### Kill mmc.exe after accidentally launching it while working remotely
        {
            taskkill /f /im mmc.exe
            write-host "MMC closed"
            start-sleep 1
        }
        
        x ### Exit
        {

        }
    }
}until($menu -eq 'x')
