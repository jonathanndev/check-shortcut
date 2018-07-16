Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = ""
Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = "C:\emporio\ORCAMENT.exe"
 '  oLink.Arguments = ""
 '  oLink.IconLocation = ""
 '  oLink.WindowStyle = "1"   
 '  oLink.WorkingDirectory = "C:\emporio\"
    oLink.Save