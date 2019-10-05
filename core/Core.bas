Attribute VB_Name = "Core"
Public SE As GMusicList
Public MousePage As MousePage
Public NotifyPage As NotifyPage
Public MouseLocked As Boolean
Public Sock As Winsock
Public AC As String, ACN As String, Money As Double
Public AnyMouseTouch As Boolean
Dim TapTime As Long
Public Sub DoTap()
    ECore.Display: DoEvents
End Sub
Public Sub FakeSleep(time As Long)
    Dim NTime As Long
    NTime = GetTickCount
    Do While GetTickCount - NTime < time
        Call DoTap
    Loop
End Sub
Public Sub RemoteSend(Cmd As String)
    Sock.SendData Cmd & Chr(-4046)
End Sub
