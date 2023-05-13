function install-Drivers
{
    # Get computer model
    $fullmodel = (Get-ciminstance -Classname Win32_ComputerSystem).Model
    switch ((Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber){
        {$_ -ge 22000} {
            $opsystem = "11"
        }

        {$_ -ge 19041 -and $_ -lt 22000} {
            $opsystem = "10"
        }

        default {
            throw "Supported versions of Windows are 10 (21H2) and 11"
        }
    }
    $opsystem
    write-host "Computer model appears to be: " $fullmodel
    $mdl1 = ($fullmodel -split " ",2)[0]
    $mdl2 = ($fullmodel -split " ",2)[1]
    $cab = (Get-ChildItem E:\drivers\$opsystem\$mdl1\$mdl2 -filter *.cab).fullname
    $drivers = (Get-ChildItem E:\drivers\$opsystem\$mdl1\$mdl2 -filter *.exe).fullname

    # Create local folder for drivers
    mkdir C:\Dell\Drivers\cabdriver
    # Copy drivers, then install
    if ($cab){
        write-host "Installing drivers from " $cab
        expand -F:* $cab C:\Dell\Drivers\cabdriver
        pnputil.exe /add-driver C:\Dell\Drivers\cabdriver\*.inf /subdirs /install
    }else{
        write-host "**** NO CAB FOUND ****"
    }
    write-host "-----   Copying drivers...   -----"
    if ($drivers){
        $drivers | copy-item -destination "C:\Dell\Drivers" -Verbose
        $main_drivers = Get-ChildItem -filter *.exe C:\Dell\Drivers | Where-Object length -gt 1GB
        $single_drivers = Get-ChildItem -filter *.exe C:\Dell\Drivers | Where-Object length -lt 1GB
        $main_drivers | ForEach-Object{
            (get-date).datetime; write-host "Extracting " $_.Name;Start-Process $_.FullName -ArgumentList '/s /e=C:\Dell\Drivers\cabdriver' -Wait
            (get-date).datetime; write-host "Installing..."
            pnputil.exe /add-driver C:\Dell\Drivers\cabdriver\*.inf /subdirs /install
        }
        $single_drivers | ForEach-Object{
            (get-date).datetime; write-host "Installing " $_.Name;Start-Process $_.FullName -argumentlist '/s' -Wait
        }
        Start-Sleep 10
    }else{
        write-host "**** NO DRIVERS ****"
    }
}

# Add registry key to disable Fast Boot
"Disabling Fast Boot..."
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t REG_DWORD /d 0 /f

# Install drivers. Alienware models do not have driver CAB packages
install-Drivers

# Rename computer to service tag
write-host "Renaming computer..."
$newname = (get-ciminstance -classname win32_bios).serialnumber

# Authenticating with old Account01 pw for this step as it is easier to type than the new one.
Rename-Computer -newname $newname -localcredential Account01 -passthru

# Set admin password to new format, as well as setting it to never expire
"Setting admin password..."
$secPass = ConvertTo-SecureString # A VERY SECURE AND UNIQUE PASSWORD IS GENERATED HERE
Set-LocalUser -Name "Account01" -password $secPass -PasswordNeverExpires:$true
shutdown -r -t 5
