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
'   ҳ�������
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
'======================================

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
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear 0 '��ջ���

    Dim pro As Single
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
    
    Page.Paint 1, Mouse.X, Mouse.y, 20, 20, argb(255, 96, 96, 96), pos:=posOnCenter
    Page.Paint 1, Mouse.X, Mouse.y, 20, 20, argb(255, 255, 255, 255), size:=3, style:=1, pos:=posOnCenter
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
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
Public Sub AnimationDone(ID As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    ReDim Water(0)
    '������Ϸ��Դ
    'Page.Res.NewImages App.path & "\assets\login"
    '����ҳ��
    'ECore.Add Page, "MousePage"
    Page.TopPage = True
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub