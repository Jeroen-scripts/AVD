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

#download software repo
#c:\\temp\\azcopy.exe copy 'https://aibsoftwarerepository.blob.core.windows.net/softwaresource/software.zip?sp=r&st=2023-08-16T14:22:43Z&se=2023-08-30T22:22:43Z&spr=https&sv=2022-11-02&sr=b&sig=aB8m1nc%2Bzz1sPGk%2FKOHZ8DMkcGPcykWIs6%2B6ph5q8mw%3D' c:\\temp\\software.zip
#Expand-Archive 'c:\\temp\\software.zip' c:\\temp
#endregion

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\RDM\Setup.RemoteDesktopManager.2021.1.44.0.msi', 'TRANSFORMS="C:\temp\software\RDM\RDMSettings.mst"', '/quiet'
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

#region Remote Desktop Manager
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\temp\software\OT\OT_12_0_0_x64.msi', 'TRANSFORMS="C:\temp\software\OT\Omninet Omnitracker 12.0.0.10344 x64 EN R01.mst"', '/quiet'
    if (Test-Path "C:\Program Files\OMNITRACKER\OMNINET.OMNITRACKER.Client.exe") {
        Write-Log "Omnitracker Client has been installed"
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
    Remove-Item -Path "C:\temp" -Recurse -Force
    if (Test-Path "C:\temp") {
        Write-Log "Temporary folder removed"
    }
    else {
        write-log "Error removing temporary folder"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error removing temporary source folder: $ErrorMessage"
}
#endregion
