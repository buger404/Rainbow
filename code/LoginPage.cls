VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Touched As Boolean, TouchIndex As Integer, CurrentAccount As Long
    Dim Account() As String, AccountN() As String, AccountMode As Long
    Dim AccountTime As Long
    Dim StayText As String
'======================================
Public Sub Login(ByVal index As Long)
    CurrentAccount = index

    If index = 0 Then
        AC = "0": ACN = "游客"
    Else
        AC = Account(index): ACN = AccountN(index)
    End If
    
    AccountMode = 2
    ECore.NewTransform transFadeIn, 700
    
    StayText = "正在连接服务器"
    FakeSleep 1000
    'Sock.Connect "172.17.10.36", "64046"
    Sock.Connect "47.100.137.52", "64046"
    'Sock.Connect "127.0.0.1", "6404"
    Do While Sock.state <> 7
        Select Case Sock.state
            Case 3: NotifyPage.Popup "连接出错", "服务器的连接被挂起", 2, "server.png", "好的": GoTo CancelLogin
            Case 4: StayText = "正在解析服务器域名"
            Case 5: StayText = "已识别服务器主机"
            Case 6: StayText = "正在连接服务器"
            Case 8: NotifyPage.Popup "连接出错", "连接被服务器关闭", 2, "server.png", "好的": GoTo CancelLogin
            Case 9: NotifyPage.Popup "连接出错", "请检查您的网络连接后重试", 1, "server.png", "好的": GoTo CancelLogin
        End Select
        DoTap
    Loop
    Money = -1
    StayText = "正在取得玩家信息"
    FakeSleep 500
    RemoteSend "login*" & AC
    Do While Money = -1
        DoTap
    Loop
    StayText = "登录成功"
    MousePage.Login
    SE.Play "login.wav"
    NotifyPage.Popup "登录成功", "后续功能敬请期待", 3, "server.png", "有生之年系列"
    Exit Sub
CancelLogin:
    AccountMode = 0
    ECore.NewTransform transFadeIn, 700
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Dim X As Long, y As Long
    
    Page.Clear '清空画布

    Page.DrawImage "background" & IIf(Touched Or AccountMode = 2, "2", "") & ".png", Mouse.X / GW * -8, Mouse.y / GH * -8
    If Touched Then Page.Paint 0, 0, 0, GW, GH, argb(0.3 * 255, 0, 0, 0)
    
    Dim TouchFace As Integer
    
    If AccountMode = 0 Then
        Page.DrawImage "face.png", GW / 2, GH / 2 - 30 - 5, alpha:=0, pos:=posOnCenter
        TouchFace = CheckMouse2
        If TouchFace Then
            If Not Touched Then Touched = True: ECore.NewTransform transFadeIn, 1000
            Page.DrawImage "faceback2.png", GW / 2, GH / 2 - 30, alpha:=1, pos:=posOnCenter
            If TouchFace = 3 Then Login 0
        Else
            If Touched Then Touched = False: ECore.NewTransform transFadeIn, 600
            Page.DrawImage "faceback.png", GW / 2, GH / 2 - 30, alpha:=0.5, pos:=posOnCenter
        End If
    End If
    
    Select Case AccountMode
        Case 0
            Page.DrawImage "face.png", GW / 2, GH / 2 - 30 - 5, alpha:=IIf(Touched, 1, 0.5), pos:=posOnCenter
            Page.Writes "游客登录", GW / 2, GH / 2 + 42, 18, argb(IIf(Touched, 1, 0.6) * 255, 255, 255, 255), align:=StringAlignmentCenter
        Case 1
            Dim pro As Single
            If AccountTime <> 0 Then
                pro = Cubic((GetTickCount - AccountTime) / 600, 0, 1, 1, 1)
                If pro > 1 Then pro = 1: AccountTime = 0
            End If
            For i = 1 To UBound(Account)
                X = GW / 2 - UBound(Account) * 200 / 2 + i * 200 - 100
                If pro <> 0 Then X = GW / 2 + (X - GW / 2) * pro
                Page.DrawImage "face.png", X, GH / 2 - 30 - 5, alpha:=0, pos:=posOnCenter
                If pro = 0 Then
                    TouchFace = CheckMouse2
                    If TouchFace Then
                        If (Not Touched) And (TouchIndex <> i) Then Touched = True: TouchIndex = i: ECore.NewTransform transFadeIn, 1000
                        If TouchFace = 3 Then Login i
                    Else
                        If Touched And TouchIndex = i Then Touched = False: TouchIndex = 0: ECore.NewTransform transFadeIn, 600
                    End If
                End If
                Page.DrawImage "faceback" & IIf(Touched And TouchIndex = i, "2", "") & ".png", X, GH / 2 - 30, alpha:=IIf(Touched And TouchIndex = i, 1, 0.5), pos:=posOnCenter
                Page.DrawImage Account(i) & ".bmp", X, GH / 2 - 30 - 5, alpha:=IIf(Touched And TouchIndex = i, 1, 0.7), pos:=posOnCenter
                Page.Writes AccountN(i), X, GH / 2 + 42, 18, argb(IIf(Touched And TouchIndex = i, 1, 0.5) * 255, 255, 255, 255), align:=StringAlignmentCenter
            Next
        Case 2
            If CurrentAccount >= 0 Then Page.Paint 0, 0, 0, GW, GH, argb(0.3 * 255, 255, 255, 255)
            Page.DrawImage "faceback.png", GW / 2, GH / 2 - 30 + FaceY, alpha:=1, pos:=posOnCenter
            If CurrentAccount <= 0 Then
                Page.DrawImage "face.png", GW / 2, GH / 2 - 30 - 5 + 1, alpha:=1, pos:=posOnCenter
            Else
                Page.DrawImage Account(CurrentAccount) & ".bmp", GW / 2, GH / 2 - 30 - 5 + 1, alpha:=1, pos:=posOnCenter
            End If
            Page.Writes StayText, GW / 2, GH / 2 + 42 + FaceY, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End Select
    
    Dim TouchQQ As Integer
    TouchQQ = CheckMouse(GW / 2 - 141 / 2, GH - 40 - 37 / 2, 141, 37)
    If AccountMode = 2 Then
        Page.ShowLoading GW / 2 - 20, GH - 60 - 20, 40, 40, 3, argb(255, 255, 170, 186), argb(255, 255, 230, 150), argb(255, 132, 230, 255)
    Else
        Page.DrawImage "loginway" & AccountMode & ".png", GW / 2, GH - 40, alpha:=IIf(TouchQQ, 1, 0.7), pos:=posOnCenter
        If TouchQQ = 3 Then
            CurrentAccount = -1
            If AccountMode = 0 Then
                StayText = "请稍候"
                AccountMode = 2: ECore.NewTransform transFadeIn, 700
                Dim QQ() As String, temp As String, Args() As String
                StayText = "正在检测登录的QQ"
                QQ = GetLoginQQ
                If UBound(QQ) = 0 Then
                    NotifyPage.Popup "使用QQ登录", "您没有在电脑上登录QQ", 1, "login.png", "我这就去登录QQ"
                    AccountMode = 0: ECore.NewTransform transFadeIn, 700
                    Exit Sub
                End If
                ReDim Account(UBound(QQ)): ReDim AccountN(UBound(QQ))
                For i = 1 To UBound(QQ)
                    StayText = "正在连接QQ" & QQ(i)
                    Call DoTap
                    temp = NetContent("https://users.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & QQ(i))
                    Args = Split(temp, """")
                    If Dir(App.path & "\temp\" & QQ(i) & ".bmp") = "" Then
                        StayText = "正在获取头像：QQ" & QQ(i)
                        Call DoTap
                        Download Args(3), App.path & "\temp\" & QQ(i) & ".bmp"
                        Page.Res.newImage App.path & "\temp\" & QQ(i) & ".bmp", 148 * 0.8 + 3, 148 * 0.8 + 3
                        Page.Res.ClipCircle QQ(i) & ".bmp"
                    End If
                    Account(i) = QQ(i): AccountN(i) = Args(5)
                Next
                AccountMode = 1
                AccountTime = GetTickCount
                ECore.NewTransform transFadeIn, 700
            Else
                AccountMode = 0
                ECore.NewTransform transFadeIn, 700
            End If
        End If
    End If
    
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets\login", 0.8
    '创建页面
    ECore.Add Page, "LoginPage"
    ReDim Account(0)
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
