VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������
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
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
    If MouseLocked And KeyAscii = vbKeyEscape Then Call UnLockMouse
End Sub

Private Sub Form_Load()
    On Error GoTo SkipFaceCreator
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Dir(App.path & "\temp", vbDirectory) <> "" Then FSO.DeleteFolder App.path & "\temp"
    FSO.CreateFolder App.path & "\temp"
SkipFaceCreator:

    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, 1152 * 0.8, 786 * 0.8
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"

    '��ʼ��ʾ
    Set StartupPage = New StartupPage
    
    EC.ActivePage = "StartupPage"
    
    Me.Show
    Call LockMouse
    
    Dim Time As Long
    Time = GetTickCount
    StartupPage.OpenTime = Time
    Do While GetTickCount - Time <= 1500
        DoTap
    Loop
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set LoginPage = New LoginPage
        Set MousePage = New MousePage
    '=============================================
    
    StartupPage.FinishTime = GetTickCount
    Do While Not EndMark
        EC.Display: DoEvents
    Loop
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Not MouseLocked Then Call LockMouse
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Not MouseLocked Then Exit Sub
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    EndMark = True
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
