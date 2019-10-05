VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotifyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Public Title As String, MsgType As Integer, Content As String, BtnText As String, Pic As String
    Public MsgTime As Long, MsgCheck As Integer, LastPage As String, DC As Long, ODC As Long
    Dim Touched As Boolean
'======================================
Public Sub Popup(Title As String, Content As String, MsgType As Integer, Pic As String, BtnText As String)
    Me.Title = Title: Me.Content = Content: Me.MsgType = MsgType: Me.Pic = Pic: Me.BtnText = BtnText
    PaintDC GDC, ODC: BlurTo DC, GDC, GameWindow, 20
    If ECore.ActivePage <> "NotifyPage" Then LastPage = ECore.ActivePage
    ECore.ActivePage = "NotifyPage"
    
    MsgTime = GetTickCount: MsgCheck = 0
    Do While MsgCheck = 0
        Call DoTap
    Loop
    
    ECore.NewTransform transFadeIn, 600, LastPage
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Dim pro As Single
    pro = 1
    If MsgTime <> 0 Then pro = Cubic((GetTickCount - MsgTime) / 1000, 0, 1, 1, 1)
    If pro > 1 Then pro = 1: MsgTime = 0
    
    If pro < 0.5 Then
        PaintDC ODC, Page.CDC, 0, 0, alpha:=1
        PaintDC DC, Page.CDC, 0, 0, alpha:=pro / 0.5
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(80 * pro / 0.5, 255, 255, 255)
    Else
        PaintDC DC, Page.CDC, 0, 0, alpha:=1
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(80, 255, 255, 255)
        pro = (pro - 0.5) / 0.5
        Page.DrawImage "dialog.png", GW / 2, GH / 2, alpha:=pro, pos:=posOnCenter
        Dim W As Long, H As Long
        W = Page.Res.ImgSize(Pic, imgGetWidth): H = Page.Res.ImgSize(Pic, imgGetHeight)
        Page.DrawImage Pic, GW / 2, GH / 2 - 122, alpha:=pro, pos:=posOnCenter
        If MsgType <> 0 Then Page.DrawImage "icon" & MsgType & ".png", GW / 2 + W * 0.4, GH / 2 - 122 + H * 0.25, alpha:=pro, pos:=posOnCenter
        Page.Writes Title, GW / 2, GH / 2 - 65, 18, argb(pro * 255, 38, 38, 38), style:=FontStyleBold, align:=StringAlignmentCenter
        Page.Writes Content, GW / 2, GH / 2 - 28, 16, argb(pro * 255, 127, 127, 127), align:=StringAlignmentCenter
        Dim M As Integer
        Page.DrawImage "button.png", GW / 2, GH / 2 + 80, alpha:=0, pos:=posOnCenter
        M = CheckMouse2
        If M Then
            If Not Touched Then Touched = True: ECore.NewTransform
        Else
            If Touched Then Touched = False: ECore.NewTransform
        End If
        Page.DrawImage "button.png", GW / 2, GH / 2 + 80, alpha:=pro * IIf(Touched, 1, 0.5), pos:=posOnCenter
        If Touched Then
            Page.Writes BtnText, GW / 2, GH / 2 + 69, 16, argb(pro * 255, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
        Else
            Page.Writes BtnText, GW / 2, GH / 2 + 68, 16, argb(pro * 255, 127, 127, 127), align:=StringAlignmentCenter
        End If
        If M = 3 Then MsgCheck = 1: Touched = False
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
    Page.Res.NewImages App.path & "\assets\notify", 0.8
    '创建页面
    ECore.Add Page, "NotifyPage"
    DC = CreateCDC(GW, GH): ODC = CreateCDC(GW, GH)
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
