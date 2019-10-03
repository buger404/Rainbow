VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContinuePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Dim pro As Single
    
    If GetTickCount Mod 600 < 300 Then
        pro = 0.6 + (GetTickCount Mod 300) / 300 * 0.4
    Else
        pro = 1 - (GetTickCount Mod 300) / 300 * 0.4
    End If
    
    Page.DrawImage "background.png", -5, -5
    Page.DrawImage "light.png", GW / 2, GH / 2 - 30, alpha:=1, pos:=posOnCenter
    Page.Paint 0, GW / 2, GH / 2 - 35, 121, 120, argb(255, 89, 89, 89), pos:=posOnCenter
    Page.Writes "点击继续", GW / 2, GH - 140, 18, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter
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
    ECore.Add Page, "ContinuePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
