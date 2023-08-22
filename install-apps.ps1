# Software install Script
#
# Applications to install:
# Devolutions Remote Desktop Manager
# Omnitracker Client
# 7-zip file manager
# Adobe Acrobat Reader DC

#region Set logging 
$logFile = "c:\installlogs\" + (get-date -format 'yyyyMMdd') + '_softwareinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}
#endregion

#region Adobe Reader
try {
    $installer = Get-ChildItem -Path "c:\temp\software\ADOBE"
    Start-Process -filepath "c:\temp\software\ADOBE\$($installer)" -Wait -ErrorAction Stop -ArgumentList '/sAll', '/rs', '/msi', 'EULA_ACCEPT=YES'
    if (Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe") {

        #Stop autoupdate
        Stop-Process -processname armsvc -Force
        Set-Service "AdobeARMservice" -startuptype Disabled
        Disable-ScheduledTask "Adobe Acrobat Update Task"

        Write-Log "Adobe Reader has been installed"
        <#
        $TargetFile   = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
        $ShortcutFile = "C:\Users\Public\Desktop\Arcobat Reader.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut     = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
        #>
    }
    else {
        write-log "Error locating the Adobe Reader executable"
    }
}
catch {
    write-log "Error installing Adobe Reader"
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

#region 7zip
$tempFolder = ("C:\temp\")
$Domain = "https://www.7-zip.org/download.html"
$temp   = (Invoke-WebRequest -uri $Domain)
$table  = $temp.ParsedHtml.getElementsByTagName('table')[0]
$ver    = $table.cells[1].getElementsByTagName('p')[0].innertext.split()[2,3]
$latest = $ver[0].replace(".","")
Invoke-WebRequest -Uri "https://www.7-zip.org/a/7z$latest-x64.exe" -OutFile "$tempFolder\7z$latest-x64.exe"

try {
    Start-Process -filepath "$tempFolder\7z$latest-x64.exe" -Wait -ErrorAction Stop -ArgumentList '/S', '/D="C:\Program Files\7-Zip"'
    if (Test-Path "C:\Program Files\7-Zip\7zFM.exe") {
        Write-Log "7-zip has been installed"
        $TargetFile   = "C:\Program Files\7-Zip\7zFM.exe"
        $ShortcutFile = "C:\Users\Public\Desktop\7-Zip File Manager.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut     = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
    }
    else {
        write-log "Error locating the 7-zip executable"
    }
}
catch {
    write-log "Error installing latest version of 7-zip"
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
