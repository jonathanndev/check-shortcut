
Dim ws 
Dim link

Set ws = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

pathDesktop = ws.SpecialFolders("Desktop")

Main




Function Main
    Dim n
    For n = 1 To 10
        If (NOT fs.FileExists(pathDesktop & "\PINBALL.lnk")) Then
            CreateShortcut
        End If
        WScript.Sleep(3600000) '1 hour
    Next
End Function

Function CreateShortcut
    Set link = ws.CreateShortcut(pathDesktop & "\pinballAtalho.lnk")
    link.TargetPath = "C:\Arquivos de programas\Windows NT\Pinball\PINBALL.EXE"
    link.Save
End Function
