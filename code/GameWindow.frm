VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "虹光"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   Icon            =   "GameWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   336
      Top             =   336
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面类模块声明
    Dim LoginPage As LoginPage
    Dim StartupPage As StartupPage
    Dim EndMark As Boolean
'==================================================

Private Sub LockMouse()
    Dim r As RECT
    GetClientRect Me.Hwnd, r
    r.Left = Me.Left / Screen.TwipsPerPixelX + 3
    r.top = Me.top / Screen.TwipsPerPixelY + 27 / (Screen.TwipsPerPixelY / 15)
    r.Bottom = r.Bottom + r.top - 3
    r.Right = r.Right + r.Left - 3
    ClipCursor r
    ShowCursor False
    MouseLocked = True
End Sub
Private Sub UnLockMouse()
    ClipCursor ByVal 0
    ShowCursor True
    MouseLocked = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
    If MouseLocked And KeyAscii = vbKeyEscape Then Call UnLockMouse
End Sub

Private Sub Form_Load()
    Set Sock = Me.Winsock

    On Error GoTo SkipFaceCreator
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Dir(App.path & "\temp", vbDirectory) <> "" Then FSO.DeleteFolder App.path & "\temp"
    FSO.CreateFolder App.path & "\temp"
SkipFaceCreator:

    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, 1152 * 0.8, 786 * 0.8
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan
    
    '创建存档（可选）
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表
    Set SE = New GMusicList
    SE.Create App.path & "\music\se"

    '开始显示
    Set StartupPage = New StartupPage
    
    EC.ActivePage = "StartupPage"
    
    Me.Show
    Call LockMouse
    
    Dim time As Long
    time = GetTickCount
    StartupPage.OpenTime = time
    Do While GetTickCount - time <= 1500
        DoTap
    Loop
    
    '在此处初始化你的页面
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set LoginPage = New LoginPage
        Set MousePage = New MousePage
    '=============================================
    
    StartupPage.FinishTime = GetTickCount
    Do While Not EndMark
        EC.Display: DoEvents
    Loop
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Not MouseLocked Then Call LockMouse
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Not MouseLocked Then Exit Sub
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    EndMark = True
    '释放Emerald资源
    EndEmerald
    End
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim temp As String, Cmds() As String, Args() As String
    Winsock.GetData temp
    Cmds = Split(temp, Chr(-4046))
    
    For i = 0 To UBound(Cmds) - 1
        Args = Split(Cmds(i), "*")
        Select Case Args(0)
            Case "loginback"
                Money = Val(Args(1))
        End Select
    Next
End Sub

