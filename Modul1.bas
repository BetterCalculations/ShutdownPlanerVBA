Attribute VB_Name = "Modul1"
Public Static Sub Start_App()
    UserForm1.Show
    
End Sub


Public Static Sub ShutDown()
    If UserForm1.TextBox1.Value = CStr(Time) Then
        OpenTask = CreateObject("WScript.Shell").Run("shutdown -s -t 10")
        MsgBox ("Ihr PC wird in 10 Sekunden heruntergefahren" + vbNewLine + "Your PC is Shutdown in 10 seconds")
        Exit Sub
    Else
        Application.OnTime Now() + TimeValue("00:0:01"), "ShutDown"
    End If
End Sub
