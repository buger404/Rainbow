Attribute VB_Name = "QQCore"
Public Sub Download(ByVal nUrl As String, ByVal nFile As String)
     Dim XmlHttp, b() As Byte
     Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
     XmlHttp.Open "GET", nUrl, True
     XmlHttp.Send
     Do While XmlHttp.readyState <> 4
        DoTap
     Loop
     If XmlHttp.readyState = 4 Then
         b() = XmlHttp.responseBody
         Open nFile For Binary As #1
         Put #1, , b()
         Close #1
     End If
     Set XmlHttp = Nothing
End Sub
Public Function NetContent(ByVal nUrl As String) As String
     Dim XmlHttp, b() As Byte
     Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
     XmlHttp.Open "GET", nUrl, True
     XmlHttp.Send
     Do While XmlHttp.readyState <> 4
        DoTap
     Loop
     NetContent = StrConvEx(XmlHttp.responseBody)
     Set XmlHttp = Nothing
End Function
Function StrConvEx(b, Optional Charset As String = "GB2312")
    Dim o As Object
    Set o = CreateObject("Adodb.Stream")
    With o
        .Type = 1: .mode = 3
        .Open: .Write b
        .position = 0: .Type = 2
        .Charset = Charset
    End With
    StrConvEx = o.ReadText: o.close
    Set o = Nothing
End Function
Public Function GetLoginQQ() As String()
    Dim Hwnd As Long, QQ As String, Size As Integer, Class As String
    Dim Ret() As String
    ReDim Ret(0)
    Hwnd = FindWindowA("CTXOPConntion_Class", vbNullString)
    If Hwnd = 0 Then Exit Function
    
    Do While Hwnd <> 0
        QQ = String(255, vbNullChar)
        GetWindowTextA Hwnd, QQ, Len(QQ)
        QQ = Left(QQ, InStr(QQ, vbNullChar) - 1)
        If InStr(QQ, "OP_") = 1 Then
            QQ = Mid(QQ, 4)
            ReDim Preserve Ret(UBound(Ret) + 1)
            Ret(UBound(Ret)) = QQ
        End If
        Hwnd = GetWindow(Hwnd, GW_HWNDNEXT)
        DoTap
    Loop
    GetLoginQQ = Ret
End Function
