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
'   ҳ�������
    Dim Page As GPage
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear '��ջ���

    Dim pro As Single
    
    If GetTickCount Mod 600 < 300 Then
        pro = 0.6 + (GetTickCount Mod 300) / 300 * 0.4
    Else
        pro = 1 - (GetTickCount Mod 300) / 300 * 0.4
    End If
    
    Page.DrawImage "background.png", -5, -5
    Page.DrawImage "light.png", GW / 2, GH / 2 - 30, alpha:=1, pos:=posOnCenter
    Page.Paint 0, GW / 2, GH / 2 - 35, 121, 120, argb(255, 89, 89, 89), pos:=posOnCenter
    Page.Writes "�������", GW / 2, GH - 140, 18, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter
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
    Page.Res.NewImages App.path & "\assets\login", 0.8
    '����ҳ��
    ECore.Add Page, "ContinuePage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub