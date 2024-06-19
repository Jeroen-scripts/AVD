# Software install Script
#
# Applications to install:
# .NET 8 Desktop Runtime
# Devolutions Remote Desktop Manager
# Omnitracker Client 

#region Set logging 
$logFile = "c:\installlogs\" + (get-date -format 'yyyyMMdd') + '_softwareinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}
#endregion

#region .NET 8.0.6 Desktop Runtime (prereq RDM)
try {
    Start-Process -filepath C:\temp\Software\NET8DesktopRuntime\windowsdesktop-runtime-8.0.6-win-x64.exe -Wait -ErrorAction Stop -ArgumentList '/install /quiet /norestart /log "c:\installlogs\DotNet80x64-Install.log"'
    Write-Log "Installing .NET 8 Desktop Runtime"
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error occured installing .NET desktop runtime: $ErrorMessage"
}
#endregion

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\RDM\rdms24prd.RemoteDesktopManager.2024.1.21.0.msi', '/quiet'
    if (Test-Path "C:\Program Files\Devolutions\Remote Desktop Manager\RemoteDesktopManager.exe") {
        Write-Log "Remote Desktop Manager has been installed. Creating desktop shortcut"
        $TargetFile   = "C:\Program Files\Devolutions\Remote Desktop Manager\RemoteDesktopManager.exe"
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
    write-log "Error occured installing Remote Desktop Manager: $ErrorMessage"
}
#endregion

#region Omnitracker
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\OT\OT.msi', 'TRANSFORMS="C:\temp\software\OT\OT.mst"', '/quiet'
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
        $ErrorMessage = $_.Exception.message
        write-log "Error locating the Omnitracker Client executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error occured installing Omnitracker Client: $ErrorMessage"
}
#endregion

#region cleanup
try {
    Remove-Item -Path "C:\temp\*" -Recurse -Force
    if (!(Test-Path "C:\temp\software")) {
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
#endregion
