
######################
# Install Office 365 #
#                    #
#                    #
######################


$logFile = "c:\installlogs\" + (get-date -format 'yyyyMMdd') + 'officeinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}


$ODTDownloadLinkRegex = '/officedeploymenttool[a-z0-9_-]*\.exe$'
$guid = [guid]::NewGuid().Guid
$tempFolder = (Join-Path -Path "C:\tempOffice\" -ChildPath $guid)
$ODTDownloadUrl = 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

if (!(Test-Path -Path $tempFolder)) {
    New-Item -Path $tempFolder -ItemType Directory
}

# Creating configuration xml
$xmlWriter = New-Object System.XMl.XmlTextWriter("$tempFolder\installOfficeAVD.xml", $Null)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"

$xmlWriter.WriteStartDocument()

$xmlWriter.WriteStartElement("Configuration")  # Configuration Start Node

$xmlWriter.WriteStartElement("Add") # Starting Add node
$xmlWriter.WriteAttributeString("OfficeClientEdition","64")
$xmlWriter.WriteAttributeString("Channel","MonthlyEnterprise")

$xmlWriter.WriteStartElement("Product") # Starting Product node
$xmlWriter.WriteAttributeString("ID","O365ProPlusRetail")

$xmlWriter.WriteStartElement("Language") # Starting Language node
$xmlWriter.WriteAttributeString("ID","en-US")
$xmlWriter.WriteEndElement() # Ending Language node
$xmlWriter.WriteStartElement("Language") # Starting Language ID node
$xmlWriter.WriteAttributeString("ID","nl-NL")
$xmlWriter.WriteEndElement() # Ending Language node
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Groove")
$xmlWriter.WriteEndElement() # Ending Language node
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Lync")
$xmlWriter.WriteEndElement() # Ending Language node
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","OneDrive")
$xmlWriter.WriteEndElement() # Ending Language node
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Teams")
$xmlWriter.WriteEndElement() # Ending Language node

$xmlWriter.WriteEndElement() # Ending Product node.

$xmlWriter.WriteEndElement() # Ending Add node.

$xmlWriter.WriteStartElement("RemoveMSI") # Starting RemoveMSI node
$xmlWriter.WriteEndElement() # Ending RemoveMSI node
$xmlWriter.WriteStartElement("Updates") # Starting Updates node
$xmlWriter.WriteAttributeString("Enabled","FALSE")
$xmlWriter.WriteEndElement() # Ending Updates node
$xmlWriter.WriteStartElement("Display") # Starting Display node
$xmlWriter.WriteAttributeString("Level","None")
$xmlWriter.WriteAttributeString("AcceptEULA","TRUE")
$xmlWriter.WriteEndElement() # Ending Display node
$xmlWriter.WriteStartElement("Logging") # Starting Logging node
$xmlWriter.WriteAttributeString("Level","Standard")
$xmlWriter.WriteAttributeString("Path","%temp%\AVDOfficeInstall")
$xmlWriter.WriteEndElement() # Ending Logging node
$xmlWriter.WriteStartElement("Property") # Starting Property node
$xmlWriter.WriteAttributeString("Name","FORCEAPPSHUTDOWN")
$xmlWriter.WriteAttributeString("Value","TRUE")
$xmlWriter.WriteEndElement() # Ending Property node
$xmlWriter.WriteStartElement("Property") # Starting Property node
$xmlWriter.WriteAttributeString("Name","SharedComputerLicensing")
$xmlWriter.WriteAttributeString("Value","1")
$xmlWriter.WriteEndElement() # Ending Property node

$xmlWriter.WriteEndElement()  # Configuration end node.

$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()

Write-log "Compiled configuration xml"

try {
    
    $HttpContent = Invoke-WebRequest -Uri $ODTDownloadUrl -UseBasicParsing
    
    if ($HttpContent.StatusCode -ne 200) { 
        throw "Office Installation script failed to find Office deployment tool link -- Response $($Response.StatusCode) ($($Response.StatusDescription))"
    }

    $ODTDownloadLinks = $HttpContent.Links | Where-Object { $_.href -Match $ODTDownloadLinkRegex }

    #pick the first link in case there are multiple
    $ODTToolLink = $ODTDownloadLinks[0].href
    Write-Log "AVD AIB Customization Office Apps : Office deployment tool link is $ODTToolLink"

    $ODTexePath = Join-Path -Path $tempFolder -ChildPath "officedeploymenttool.exe"

    #download office deployment tool

    Write-Log "AVD AIB Customization Office Apps : Downloading ODT tool into folder $ODTexePath"
    $ODTResponse = Invoke-WebRequest -Uri "$ODTToolLink" -UseBasicParsing -UseDefaultCredentials -OutFile $ODTexePath -PassThru

    if ($ODTResponse.StatusCode -ne 200) { 
        throw "Office Installation script failed to download Office deployment tool -- Response $($ODTResponse.StatusCode) ($($ODTResponse.StatusDescription))"
    }

    Write-Log "AVD AIB Customization Office Apps : Extracting setup.exe into $tempFolder"
    # extract setup.exe
    Start-Process -FilePath $ODTexePath -ArgumentList "/extract:`"$($tempFolder)`" /quiet" -PassThru -Wait -NoNewWindow

    $setupExePath = Join-Path -Path $tempFolder -ChildPath 'setup.exe'
            
    Write-Log "AVD AIB Customization Office Apps : Running setup.exe to download Office"
    $ODTRunSetupExe = Start-Process -FilePath $setupExePath -ArgumentList "/download installOfficeAVD.xml" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden

    if (!$ODTRunSetupExe) {
        Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to download Office"
    }

    if ( $ODTRunSetupExe.ExitCode) {
        Throw "AVD AIB Customization Office Apps : Exit code $($ODTRunSetupExe.ExitCode) returned from `"$setupExePath`" to download Office"
    }

    Write-Log "AVD AIB Customization Office Apps : Running setup.exe to Install Office"
    $InstallOffice = Start-Process -FilePath $setupExePath -ArgumentList "/configure installOfficeAVD.xml" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden

    if (!$InstallOffice) {
        Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to install Office"
    }

    if ( $ODTRunSetupExe.ExitCode ) {
        Throw "AVD AIB Customization Office Apps : Exit code $($ODTRunSetupExe.ExitCode) returned from `"$setupExePath`" to download Office"
    }
    
}
catch {
    $PSCmdlet.ThrowTerminatingError($PSitem)
}
