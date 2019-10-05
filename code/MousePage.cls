VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MousePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
    Private Type MouseWater
        X As Long
        y As Long
        time As Long
        r As Long
        g As Long
        b As Long
    End Type
    Dim Water() As MouseWater
    Dim TL As Long, TR As Long, TabIndex As Integer
    Dim OTL As Long, OTR As Long, TabTime As Long
    Dim LoginTime As Long, Logined As Boolean
    Dim MouseTime As Long, MouseTouched As Boolean
    
'======================================
Public Sub Login()
    Page.Res.newImage App.path & "\assets\login\faceback.png", 148 * 0.4, 148 * 0.4
    Call DoTap
    If AC = "0" Then
        Page.Res.newImage App.path & "\assets\login\face.png", 148 * 0.4, 148 * 0.4
        Call DoTap
    Else
        Page.Res.newImage App.path & "\temp\" & AC & ".bmp", 148 * 0.4 + 3, 148 * 0.4 + 3
        Call DoTap
        Page.Res.ClipCircle AC & ".bmp"
        Call DoTap
    End If
    LoginTime = GetTickCount
    Logined = True
End Sub
Public Sub SwitchTab(index As Integer)
    If index > 0 Then NotifyPage.Popup "即将推出", "我将在第一时间编写代码", 2, "system.png", "好的": Exit Sub
    TabTime = GetTickCount
    OTL = TL: OTR = TR
    
    TL = 60 + 100 * index - 40
    TR = 60 + 100 * index + 40
    TabIndex = index
End Sub
Public Sub NewWater(ByVal X As Long, ByVal y As Long)
    Dim Co(3) As Byte, Color As Long
    Color = GetPixel(ECore.CDC, X, y)
    CopyMemory Co(0), Color, 4
    
    ReDim Preserve Water(UBound(Water) + 1)
    With Water(UBound(Water))
        .X = X
        .y = y
        .time = GetTickCount
        .r = 255 'Co(2) + 40
        .g = 255 'Co(1) + 40
        .b = 255 'Co(0) + 40
    End With
End Sub

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    Dim pro As Single

    If ECore.ActivePage <> "StartupPage" Then
        Page.Paint 0, 0, 0, GW, 55, argb(255, 38, 38, 38)
        Dim M As Integer, Tabs(2) As String, i As Integer
        Tabs(0) = "登录": Tabs(1) = "地图": Tabs(2) = "商店"
        For i = 0 To 2
            M = CheckMouse(20 + 100 * i, 0, 80, 55)
            If M Then
                Page.Paint 0, 20 + 100 * i, 0, 80, 55, argb(50, 255, 255, 255)
                If M = 3 Then Call SwitchTab(i)
            End If
            Page.Writes Tabs(i), 60 + 100 * i, 15, 18, argb(255, 255, 255, 255), style:=IIf(M = 0, FontStyleRegular, FontStyleBold), align:=StringAlignmentCenter
        Next
        Dim NTL As Long, NTR As Long
        If TabTime <> 0 Then
            pro = Cubic((GetTickCount - TabTime) / 500, 0, 1, 1, 1)
            If pro > 1 Then pro = 1: TabTime = 0
            If TR > OTR Then
                If pro < 0.5 Then
                    NTL = OTL: NTR = OTR + (TR - OTR) * (pro / 0.5)
                Else
                    NTR = TR: NTL = OTL + (TL - OTL) * ((pro - 0.5) / 0.5)
                End If
            Else
                If pro < 0.5 Then
                    NTR = OTR: NTL = OTL + (TL - OTL) * (pro / 0.5)
                Else
                    NTL = TL: NTR = OTR + (TR - OTR) * ((pro - 0.5) / 0.5)
                End If
            End If
        Else
            NTL = TL: NTR = TR
        End If
        Page.Paint 0, NTL, 55 - 6, NTR - NTL, 6, argb(255, 50, 202, 185)
    End If

    If Logined Then
        pro = 1
        If LoginTime <> 0 Then pro = Cubic((GetTickCount - LoginTime) / 1000, 0, 1, 1, 1)
        If pro > 1 Then pro = 1: LoginTime = 0
        Page.DrawImage "faceback.png", GW - 70, 55, alpha:=pro, pos:=posOnCenter
        Page.DrawImage IIf(AC = "0", "face.png", AC & ".bmp"), GW - 70, 55, alpha:=pro, pos:=posOnCenter
    End If

    For i = 1 To UBound(Water)
        If i > UBound(Water) Then Exit For
        pro = Cubic((GetTickCount - Water(i).time) / 1000, 0, 1, 1, 1)
        If pro > 1 Then
            Water(i) = Water(UBound(Water)): ReDim Preserve Water(UBound(Water) - 1)
            i = i - 1
        Else
            Page.Paint 1, Water(i).X, Water(i).y, 20 + pro * 200, 20 + pro * 200, argb(150 - 150 * pro, Water(i).r, Water(i).g, Water(i).b), pos:=posOnCenter
        End If
    Next
    
    If AnyMouseTouch <> MouseTouched Then MouseTouched = AnyMouseTouch: MouseTime = GetTickCount
    AnyMouseTouch = False
    
    pro = 1
    If MouseTime <> 0 Then pro = Cubic((GetTickCount - MouseTime) / 300, 0, 1, 1, 1)
    If pro > 1 Then pro = 1: MouseTime = 0
    If Not MouseTouched Then pro = 1 - pro
    
    Page.Paint 1, Mouse.X, Mouse.y, 24 - 4 * pro, 24 - 4 * pro, argb(255, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, Mouse.X, Mouse.y, 18 - 4 * pro, 18 - 4 * pro, argb(255, 96 + (16 - 96) * pro, 96 + (229 - 96) * pro, 96 + (255 - 96) * pro), pos:=posOnCenter
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
    ReDim Water(0)
    '导入游戏资源
    'Page.Res.NewImages App.path & "\assets\login"
    '创建页面
    'ECore.Add Page, "MousePage"
    Page.TopPage = True
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
