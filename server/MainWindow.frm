VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rainbow Server"
   ClientHeight    =   6000
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8808
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainWindow.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   734
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer StateTimer 
      Interval        =   1000
      Left            =   8088
      Top             =   264
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   408
      Top             =   288
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Label sockState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4212
      TabIndex        =   0
      Top             =   3480
      Width           =   336
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then MkDir App.Path & "\Logs"
    LogWindow.Show
    
    ReDim ACL.AC(0)
    If Dir(App.Path & "\user") <> "" Then
        Open App.Path & "\user" For Binary As #1
        Get #1, , ACL
        Close #1
        NewLog "User information has been created , " & UBound(ACL.AC) & " users in total ."
    End If
    
    NewLog "Sock 0 starts listening ..."
    
    Sock(0).Bind 64046: Sock(0).Listen
    
    Me.Width = 734 * Screen.TwipsPerPixelX
    Me.Height = 530 * Screen.TwipsPerPixelY
    
    sockState.Left = Me.Width / Screen.TwipsPerPixelX / 2
    sockState.Top = Me.Height / Screen.TwipsPerPixelY / 2 + 20
    
    NewLog "Local IP : " & Sock(0).LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = (MsgBox("所有连接都将会丢失，无论如何也要关闭吗？", 48 + vbYesNo, "虹光服务端") = vbNo)
    If Cancel = 0 Then Unload LogWindow
End Sub

Private Sub Sock_Close(Index As Integer)
    Unload Sock(Index)
    NewLog "Sock " & Index & " disconnected ."
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Sock(Index).Close
    Load Sock(Sock.UBound + 1)
    With Sock(Sock.UBound)
        .Accept requestID
        NewLog "Sock " & Sock.UBound & " connected , remote IP : " & Sock(Index).RemoteHostIP
    End With
    Sock(Index).Bind 64046: Sock(Index).Listen
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim temp As String, Cmds() As String, Args() As String
    Sock(Index).GetData temp
    Cmds = Split(temp, Chr(-4046))
    
    For i = 0 To UBound(Cmds) - 1
        Args = Split(Cmds(i), "*")
        Select Case Args(0)
            Case "login"
                If Args(1) = "0" Then
                    NewLog "Sock " & Index & " : Logined as a visitor ."
                    RemoteSend Index, "loginback*0"
                Else
                    Dim Find As Boolean
                    For s = 1 To UBound(ACL.AC)
                        If ACL.AC(s).QQ = Args(1) Then Sock(Index).Tag = s: Find = True: Exit For
                    Next
                    If Find Then
                        NewLog "Sock " & Index & " : QQ " & Args(1) & " logined ."
                        RemoteSend Index, "loginback*" & ACL.AC(Sock(Index).Tag).Money
                    Else
                        ReDim Preserve ACL.AC(UBound(ACL.AC) + 1)
                        ACL.AC(UBound(ACL.AC)).QQ = Args(1)
                        Open App.Path & "\user" For Binary As #1
                        Put #1, , ACL
                        Close #1
                        NewLog "Sock " & Index & " : QQ " & Args(1) & " joined the game ."
                        RemoteSend Index, "loginback*0"
                        RemoteSend Index, "newborn"
                    End If
                End If
        End Select
    Next
End Sub

Private Sub StateTimer_Timer()
    Dim Ret As String
    
    For i = 0 To Sock.Count - 1
        Ret = Ret & "Sock " & Sock.Item(i).Index & " : " & Sock.Item(i).State & " (" & Sock.Item(i).RemoteHostIP & ",QQ:" & Sock.Item(i).Tag & ")" & vbCrLf
    Next
    
    sockState.Caption = Ret
End Sub

Public Sub RemoteSend(Index As Integer, Cmd As String)
    Sock(Index).SendData Cmd & Chr(-4046)
End Sub
