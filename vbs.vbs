dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject") 
intAnswer = _
    Msgbox("Do you want the hosts file to be unloked?", _
        vbYesNo, "Click 2 Trick")

If intAnswer = vbYes Then
    If filesys.FileExists("C:\Windows\System32\drivers\etc\hosts") Then 
		filesys.MoveFile "C:\Windows\System32\drivers\etc\hosts", "C:\" 
	End If 
    Msgbox "Hack complete"
Else
	If filesys.FileExists("C:\hosts") Then
		filesys.MoveFile "C:\hosts", "C:\Windows\System32\drivers\etc\"  
    End if	
    Msgbox "host file restored !"
End If