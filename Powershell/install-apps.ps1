# Software install Script
#
# Applications to install:
#
# Devolutions Remote Desktop Manager
#  
# Microsoft OneDrive Machine-install

#region Set logging 
$logFile = "c:\temp\" + (get-date -format 'yyyyMMdd') + '_softwareinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}
#endregion

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\49C7497\Setup.RemoteDesktopManager.2021.1.44.0.msi', 'TRANSFORMS="C:\temp\49C7497\RDMSettings.mst"', '/quiet'
    if (Test-Path "C:\Program Files (x86)\Devolutions\Remote Desktop Manager\RemoteDesktopManager64.exe") {
        Write-Log "Remote Desktop Manager has been installed"
    }
    else {
        write-log "Error locating the Remote Desktop Manager executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Remote Desktop Manager: $ErrorMessage"
}
#endregion

#region OneDrive Machine Install
<#
try {
    REG ADD "HKLM\Software\Microsoft\OneDrive" /v "AllUsersInstall" /t REG_DWORD /d 1 /reg:64
    Start-Process -filepath 'C:\temp\OneDrive\OneDriveSetup.exe' -Wait -ErrorAction Stop -ArgumentList '/silent /allusers'
    REG ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive /t REG_SZ /d "C:\Program Files\Microsoft OneDrive\OneDrive.exe /background" /f
    REG ADD "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" /v "SilentAccountConfig" /t REG_DWORD /d 1 /f
    REG ADD "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" /v "KFMSilentOptIn" /t REG_SZ /d "<your-AzureAdTenantId>" /f
    if (Test-Path "C:\Program Files\Notepad++\notepad++.exe") {
        Write-Log "Notepad++ has been installed"
    }
    else {
        write-log "Error locating the Notepad++ executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Notepad++: $ErrorMessage"
}
#endregion
#>
