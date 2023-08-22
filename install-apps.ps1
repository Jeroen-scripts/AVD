# Software install Script
#
# Applications to install:
# Devolutions Remote Desktop Manager
# Omnitracker Client 

#region Set logging 
$logFile = "c:\installlogs\" + (get-date -format 'yyyyMMdd') + '_softwareinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}
#endregion

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\RDM\Setup.RemoteDesktopManager.2021.1.44.0.msi', 'TRANSFORMS="C:\temp\software\RDM\RDMSettings.mst"', '/quiet'
    if (Test-Path "C:\Program Files (x86)\Devolutions\Remote Desktop Manager\RemoteDesktopManager64.exe") {
        Write-Log "Remote Desktop Manager has been installed"
        $TargetFile   = "C:\Program Files (x86)\Devolutions\Remote Desktop Manager\RemoteDesktopManager64.exe"
        $ShortcutFile = "C:\Users\Public\Desktop\rdm.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut     = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
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

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\OT\OT_12_0_0_x64.msi', 'TRANSFORMS="C:\temp\software\OT\Omninet Omnitracker 12.0.0.10344 x64 EN R01.mst"', '/quiet'
    if (Test-Path "C:\Program Files\OMNITRACKER\OMNINET.OMNITRACKER.Client.exe") {
        Write-Log "Omnitracker Client has been installed. Creating desktop shortcut"
        $TargetFile   = "C:\Program Files\OMNITRACKER\OMNINET.OMNITRACKER.Client.exe"
        $ShortcutFile = "C:\Users\Public\Desktop\omnitracker.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut     = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
    }
    else {
        write-log "Error locating the Omnitracker Client executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Omnitracker Client: $ErrorMessage"
}
#endregion

#region cleanup
try {
    Remove-Item -Path "C:\temp\software" -Recurse -Force
    if (Test-Path "C:\temp\software") {
        Write-Log "Software source folder failed to be removed"
    }
    else {
        write-log "Error removing software source folder"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error removing software source folder: $ErrorMessage"
}

try {
    Remove-Item -Path "C:\temp\software.zip" -Force
    if (Test-Path "C:\temp\software.zip") {
        Write-Log "Software zip file failed to be removed"
    }
    else {
        write-log "Error removing software zip file"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error removing software zip file: $ErrorMessage"
}
