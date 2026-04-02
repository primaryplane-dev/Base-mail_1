Option Explicit

Public Sub subBeforeEdit()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub subAfterEdit()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' AS400格納時のバイト数を求める(半角 1 バイト、全角 2 バイト)
'  ２バイト文字の前後に制御文字が入ることを考慮する
Public Function LenAS(ByVal i_文字列 As String) As Integer
    Dim i       As Integer
    Dim moji_z  As Integer     '前回の文字のバイト数
    Dim ret     As Integer     '長さ

    For i = 0 To (Len(i_文字列) - 1)
        Select Case LenB(StrConv(Mid(i_文字列, i + 1, 1), vbFromUnicode))
        Case 1
            If moji_z = 2 Then ret = ret + 1   '制御文字(2byte文字の終了)
            ret = ret + 1
            moji_z = 1
        Case 2
            If moji_z <> 2 Then ret = ret + 1  '制御文字(2byte文字の開始)
            ret = ret + 2
            moji_z = 2
        End Select
    Next
    If moji_z = 2 Then ret = ret + 1

    LenAS = ret
End Function

Public Function fncGetKJNM(ByVal i_BUTUKCD As String) As String
    fncGetKJNM = ""
    Select Case i_BUTUKCD
    Case "001": fncGetKJNM = "武蔵工場"
    Case "002": fncGetKJNM = "枚方工場"
    Case "102": fncGetKJNM = "タカラ食品"
    Case "103": fncGetKJNM = "福岡工場"
    End Select
End Function

Public Function fncGetBUTUKCD(ByVal i_SVpath As String) As String
    fncGetBUTUKCD = ""
    Select Case i_SVpath
    Case P_SVPathM: fncGetBUTUKCD = "001"
    Case P_SVPathH: fncGetBUTUKCD = "002"
    Case P_SVPathT: fncGetBUTUKCD = "102"
    Case P_SVPathF: fncGetBUTUKCD = "103"
    End Select
End Function

Public Function fncGetSVPath(ByVal i_BUTUKCD As String) As String
    fncGetSVPath = ""
    Select Case i_BUTUKCD
    Case "001": fncGetSVPath = P_SVPathM
    Case "002": fncGetSVPath = P_SVPathH
    Case "102": fncGetSVPath = P_SVPathT
    Case "103": fncGetSVPath = P_SVPathF
    End Select
End Function

Public Function fncGetIPAdd(ByVal i_BUTUKCD As String) As String
    fncGetIPAdd = ""
    Select Case i_BUTUKCD
    Case "001": fncGetIPAdd = P_IPAddM
    Case "002": fncGetIPAdd = P_IPAddH
    Case "102": fncGetIPAdd = P_IPAddT
    Case "103": fncGetIPAdd = P_IPAddF
    End Select
End Function

Public Function fncGetKCD(ByVal i_BUTUKCD As String) As String
    fncGetKCD = ""
    Select Case i_BUTUKCD
    Case "001": fncGetKCD = "09"
    Case "002": fncGetKCD = "06"
    Case "102": fncGetKCD = "36"
    Case "103": fncGetKCD = "39"
    End Select
End Function

Public Function fncGetKJKey(ByVal i_BUTUKCD As String) As String
    fncGetKJKey = ""
    Select Case i_BUTUKCD
    Case "001": fncGetKJKey = "M"
    Case "002": fncGetKJKey = "H"
    Case "102": fncGetKJKey = "T"
    Case "103": fncGetKJKey = "F"
    End Select
End Function

Public Function fncGetKJMark(ByVal i_BUTUKCD As String) As String
    fncGetKJMark = ""
    Select Case i_BUTUKCD
    Case "001": fncGetKJMark = "+FMU"
    Case "002": fncGetKJMark = "+FHI"
    Case "102": fncGetKJMark = "+TF"
    Case "103": fncGetKJMark = "+FFU"
    End Select
End Function

Public Function fncGetLINO(ByVal i_BUTUKCD As String, ByVal i_LINM As String) As String
    fncGetLINO = ""
    Select Case i_BUTUKCD
    Case "001"
        If InStr(i_LINM, "デバイダー") > 0 Or InStr(i_LINM, "火星人") > 0 Then fncGetLINO = "027"
        If InStr(i_LINM, "HM") > 0 Or InStr(i_LINM, "メロンビス") > 0 Then fncGetLINO = "028"
    Case "002"
        If InStr(i_LINM, "デバイダー") > 0 Then fncGetLINO = "114"
        If InStr(i_LINM, "HM") > 0 Or InStr(i_LINM, "ﾌﾘｯﾂﾗｲﾝ") > 0 Then fncGetLINO = "214"
    Case "103"
        If InStr(i_LINM, "HM") > 0 Or InStr(i_LINM, "ﾌﾘｯﾂﾗｲﾝ") > 0 Then fncGetLINO = "925"
        If InStr(i_LINM, "デバイダー") > 0 Or InStr(i_LINM, "FE") > 0 Then fncGetLINO = "935"
    Case "102"
        If InStr(i_LINM, "AD") > 0 Then fncGetLINO = "915"
        If InStr(i_LINM, "バラエティーライン") > 0 Or InStr(i_LINM, "デバイダー") > 0 Or InStr(i_LINM, "HM") > 0 Or InStr(i_LINM, "焼成冷凍") > 0 Then fncGetLINO = "925"
    End Select
End Function

Public Function fncGetKLINO(ByVal i_BUTUKCD As String, ByVal i_LINO As String) As String
    fncGetKLINO = ""
    Select Case i_BUTUKCD
    Case "001"
        If i_LINO = "027" Or i_LINO = "227" Then fncGetKLINO = "127": Exit Function
        If i_LINO = "028" Or i_LINO = "228" Then fncGetKLINO = "128": Exit Function
    Case Else
        fncGetKLINO = i_LINO: Exit Function
    End Select
End Function

Public Sub subGetAdd(ByVal i_BUTUKCD As String, ByRef o_Add1 As String, ByRef o_Add2 As String)
    Select Case i_BUTUKCD
    Case "001": o_Add1 = "フジパン㈱　武蔵工場": o_Add2 = "〒358-0032　埼玉県入間市狭山ヶ原松原108-5"
    Case "002": o_Add1 = "フジパン㈱　枚方工場": o_Add2 = "〒573-0014　大阪府枚方市村野高見台1-40"
    Case "102": o_Add1 = "タカラ食品㈱": o_Add2 = "〒492-8279　愛知県稲沢市天池遠松町10"
    Case "103": o_Add1 = "㈱九州フジパン　福岡工場": o_Add2 = "〒811-0123　福岡県糟屋郡新宮町上府北3丁目1-1"
    End Select
End Sub

Public Function fncGetBUTUKCD2(ByVal i_KNM As String) As String
    fncGetBUTUKCD2 = ""
    If InStr(i_KNM, "武蔵") > 0 Then
        fncGetBUTUKCD2 = "001"
    ElseIf InStr(i_KNM, "枚方") > 0 Then
        fncGetBUTUKCD2 = "002"
    ElseIf InStr(i_KNM, "福岡") > 0 Then
        fncGetBUTUKCD2 = "103"
    ElseIf InStr(i_KNM, "タカラ") > 0 Then
        fncGetBUTUKCD2 = "102"
    End If
End Function

'Public Sub subSetSVPath()
'    ReDim P_SVPath(0)
'    ReDim Preserve P_SVPath(UBound(P_SVPath) + 1)
'    P_SVPath(UBound(P_SVPath)) = P_SVPathH
'
'    ReDim Preserve P_SVPath(UBound(P_SVPath) + 1)
'    P_SVPath(UBound(P_SVPath)) = P_SVPathM
'
'    ReDim Preserve P_SVPath(UBound(P_SVPath) + 1)
'    P_SVPath(UBound(P_SVPath)) = P_SVPathT
'
'    ReDim Preserve P_SVPath(UBound(P_SVPath) + 1)
'    P_SVPath(UBound(P_SVPath)) = P_SVPathF
'
'    ReDim P_IPAdd(0)
'    ReDim Preserve P_IPAdd(UBound(P_IPAdd) + 1)
'    P_IPAdd(UBound(P_IPAdd)) = P_IPAddH
'
'    ReDim Preserve P_IPAdd(UBound(P_IPAdd) + 1)
'    P_IPAdd(UBound(P_IPAdd)) = P_IPAddM
'
'    ReDim Preserve P_IPAdd(UBound(P_IPAdd) + 1)
'    P_IPAdd(UBound(P_IPAdd)) = P_IPAddT
'
'    ReDim Preserve P_IPAdd(UBound(P_IPAdd) + 1)
'    P_IPAdd(UBound(P_IPAdd)) = P_IPAddF
'
'End Sub

Public Function fncEditChk(ByVal i_BUTUKCD As String) As Long
    fncEditChk = 0
    
    If Not Dir(fncGetSVPath(i_BUTUKCD) & "EditB.txt") = "" Then
        fncEditChk = 1
    End If
    If Not Dir(fncGetSVPath(i_BUTUKCD) & "EditS.txt") = "" Then
        fncEditChk = 2
    End If
End Function

'Public Function fncEditChk(ByRef o_blH As Boolean, ByRef o_blM As Boolean, ByRef o_blT As Boolean, ByRef o_blF As Boolean) As Long
'    Dim i       As Integer
'    o_blH = False: o_blM = False: o_blT = False: o_blF = False
'    fncEditChk = 0
'    For i = 1 To UBound(P_SVPath)
'        If Not Dir(P_SVPath(i) & "EditB.txt") = "" Then
'            If P_SVPath(i) = P_SVPathH Then o_blH = True
'            If P_SVPath(i) = P_SVPathM Then o_blM = True
'            If P_SVPath(i) = P_SVPathT Then o_blT = True
'            If P_SVPath(i) = P_SVPathF Then o_blF = True
'            fncEditChk = 1
'        End If
'        If Not Dir(P_SVPath(i) & "EditS.txt") = "" Then
'            If P_SVPath(i) = P_SVPathH Then o_blH = True
'            If P_SVPath(i) = P_SVPathM Then o_blM = True
'            If P_SVPath(i) = P_SVPathT Then o_blT = True
'            If P_SVPath(i) = P_SVPathF Then o_blF = True
'            fncEditChk = 2
'        End If
'    Next
'End Function

Public Sub subMakeEditFile(ByVal i_BUTUKCD As String)
    Open fncGetSVPath(i_BUTUKCD) & "EditB.txt" For Output As #1
    Close #1
End Sub

'Public Sub subMakeEditFile()
'    Dim i       As Integer
'    For i = 1 To UBound(P_SVPath)
'        Open P_SVPath(i) & "EditB.txt" For Output As #1
'        Close #1
'    Next
'End Sub

Public Sub subDeleteEditFile(ByVal i_BUTUKCD As String)
    If Not Dir(fncGetSVPath(i_BUTUKCD) & "EditB.txt") = "" Then
        Kill fncGetSVPath(i_BUTUKCD) & "EditB.txt"
    End If
End Sub

'Public Sub subDeleteEditFile()
'    Dim i       As Integer
'    For i = 1 To UBound(P_SVPath)
'        If Not Dir(P_SVPath(i) & "EditB.txt") = "" Then
'            Kill P_SVPath(i) & "EditB.txt"
'        End If
'    Next
'End Sub

Public Function fncExistData() As Boolean
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    fncExistData = False
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & " FROM LIBWMF.WGIP01"
    strSQL = strSQL & " WHERE GIDELT <> 'X'"
    strSQL = strSQL & "   AND GIKBN = '2' "
    strSQL = strSQL & "   AND NOT GIPROJ = '9' "
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then fncExistData = True
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Function

