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
'   ҳ�������
    Dim Page As GPage
    Public FinishTime As Long, OpenTime As Long
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear '��ջ���

    Page.DrawImage "background.png", -5, -5
    
    Dim pro As Single
    pro = 1
    If OpenTime <> 0 Then
        pro = (GetTickCount - OpenTime) / 1000
        If pro > 1 Then pro = 1: OpenTime = 0
    End If
    If FinishTime <> 0 Then
        pro = 1 - (GetTickCount - (FinishTime + 2000)) / 1000
        If pro < 0 Then pro = 0: FinishTime = 0: ECore.NewTransform transFadeIn, 1000, "LoginPage"
    End If
    
    Page.DrawImage "logo.png", GW / 2, GH / 2 - 30, alpha:=pro, pos:=posOnCenter
    
    Page.Writes "Copyright 2019 Error404 all rights reserved .", GW / 2, GH - 40, 16, argb(pro * 200, 255, 255, 255), align:=StringAlignmentCenter
    Page.ShowLoading GW / 2 - 10, GH / 2 + 40, 30, 30, 3, argb(pro * 255, 255, 170, 186), argb(pro * 255, 255, 230, 150), argb(pro * 255, 132, 230, 255)
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.newImage App.path & "\assets\login\background.png", 0.8
    Page.Res.newImage App.path & "\assets\login\logo.png", 0.8
    '����ҳ��
    ECore.Add Page, "StartupPage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub