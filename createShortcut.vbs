
Dim ws, link

Set ws = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

pathDesktop = ws.SpecialFolders("Desktop")

If (NOT fs.FileExists(pathDesktop & "\PINBALL.lnk")) Then
    Set link = ws.CreateShortcut(pathDesktop & "\pinballAtalho.lnk")
    link.TargetPath = "C:\Arquivos de programas\Windows NT\Pinball\PINBALL.EXE"
    link.Save
End If

