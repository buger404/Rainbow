Attribute VB_Name = "Core"
Public Type Account
    QQ As String
    Money As Double
End Type
Public Type AccountList
    AC() As Account
End Type
Public ACL As AccountList
Public Sub NewLog(Str As String)
    Open App.Path & "\Logs\" & Year(Now) & "." & Month(Now) & "." & Day(Now) & "_" & Hour(Now) & ".log" For Append As #1
    Print #1, Now, Str
    Close #1
    
    LogWindow.LogBox.Text = Now & "     " & Str & vbCrLf & LogWindow.LogBox.Text
    If Len(LogWindow.LogBox.Text) > 10000 Then LogWindow.LogBox.Text = Left(LogWindow.LogBox.Text, 10000)
End Sub
