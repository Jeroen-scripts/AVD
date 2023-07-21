$xmlWriter = New-Object System.XMl.XmlTextWriter("C:\Temp\installOfficeAVD.xml", $Null)
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






$xmlWriter.WriteStartElement("Product") # Starting Product node
$xmlWriter.WriteAttributeString("ID","O365ProPlusRetail")

$xmlWriter.WriteStartElement("Language") # Starting Language node
$xmlWriter.WriteAttributeString("ID","en-US")
$xmlWriter.WriteStartElement("Language") # Starting Language ID node
$xmlWriter.WriteAttributeString("ID","nl-NL")
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Groove")
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Lync")
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","OneDrive")
$xmlWriter.WriteStartElement("ExcludeApp") # Starting Exclude ID node
$xmlWriter.WriteAttributeString("ID","Teams")















$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()