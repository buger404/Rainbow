VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartupPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Public FinishTime As Long, OpenTime As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Page.DrawImage "background.png", -5, -5
    
    Dim pro As Single
    pro = 1
    If OpenTime <> 0 Then
        pro = (GetTickCount - OpenTime) / 1000
        If pro > 1 Then pro = 1: OpenTime = 0
    End If
    If FinishTime <> 0 Then
        pro = 1 - (GetTickCount - (FinishTime + 2000)) / 1000
        If pro < 0 Then pro = 0: FinishTime = 0: MousePage.SwitchTab 0: ECore.NewTransform transFadeIn, 1000, "LoginPage"
    End If
    
    Page.DrawImage "logo.png", GW / 2, GH / 2 - 30, alpha:=pro, pos:=posOnCenter
    
    Page.Writes "Copyright 2019 Error404 all rights reserved .", GW / 2, GH - 40, 16, argb(pro * 200, 255, 255, 255), align:=StringAlignmentCenter
    Page.ShowLoading GW / 2 - 10, GH / 2 + 40, 30, 30, 3, argb(pro * 255, 255, 170, 186), argb(pro * 255, 255, 230, 150), argb(pro * 255, 132, 230, 255)
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
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
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.newImage App.path & "\assets\login\background.png", 0.8
    Page.Res.newImage App.path & "\assets\login\logo.png", 0.8
    '创建页面
    ECore.Add Page, "StartupPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
