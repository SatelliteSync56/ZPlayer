Dim s
Dim mediainit
Set Sound = CreateObject("WMPlayer.OCX.7")
Set wShell=CreateObject("WScript.Shell")
s = MsgBox("ZPlayer by Trangle, 2019, support development at trangleproject.wordpress.com",1,"ZPlayer 1.08")
If s = 2 Then
Wscript.Quit
End If
'mediainit = MsgBox("Please select an audio file.",1,"ZPlayer 1.08")
'If mediainit = 2 Then
'Wscript.Quit
'End If
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
'stpswrd = InputBox("Password:","Create Account")
Sound.URL = sFileSelected
mediainit = MsgBox("Now playing: " & sFileSelected,1,"ZPlayer 1.08")
If mediainit = 2 Then
Sound.close
Wscript.Quit
End If
Sound.Controls.play
While Sound.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
Sound.close
