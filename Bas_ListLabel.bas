Option Explicit

Public Sub subMainLabel()
    Call subBeforeEdit
    Call subEditListLabel
    Call subAfterEdit
End Sub

Private Sub subDeleteLabel()
    Dim ST      As Worksheet
    Dim bExist  As Boolean
    Application.DisplayAlerts = False
    bExist = False
    For Each ST In Worksheets
        If ST.Name = "Label" Then bExist = True
    Next ST
    If bExist Then ThisWorkbook.Sheets("Label").Delete
    Application.DisplayAlerts = True
End Sub

Private Sub subInitializeLabel()
    'ラベルの雛型の内容をクリアする
    
    'PROJ:1の雛型
    stHinaLabel.Range(stHinaLabel.Cells(2, 2), stHinaLabel.Cells(3, 2)).ClearContents
    stHinaLabel.Range(stHinaLabel.Cells(6, 2), stHinaLabel.Cells(11, 2)).ClearContents
    
    'PROJ:5の雛型
    stHinaLabel.Range(stHinaLabel.Cells(16, 2), stHinaLabel.Cells(17, 2)).ClearContents
    stHinaLabel.Range(stHinaLabel.Cells(20, 2), stHinaLabel.Cells(28, 2)).ClearContents

    'PROJ:3の雛型
    stHinaLabel.Range(stHinaLabel.Cells(33, 2), stHinaLabel.Cells(34, 2)).ClearContents
    stHinaLabel.Range(stHinaLabel.Cells(37, 2), stHinaLabel.Cells(45, 2)).ClearContents

End Sub

Private Sub subEditListLabel()
    Dim ST          As Worksheet
    Dim STHina      As Worksheet: Set STHina = stHinaLabel
    Dim ST2         As Worksheet: Set ST2 = stList
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim lMaxRow     As Long: lMaxRow = stList.Cells(stList.Rows.Count, 1).End(xlUp).Row
    Dim i           As Long
    Dim errFLG      As Boolean: errFLG = False
    Dim Cnt         As Long: Cnt = 0
    Dim r           As Range
    Dim lWk         As Long
    Dim arrRow()    As Long
    Dim bFind       As Boolean
    Dim bHUp        As Boolean
    
    ReDim arrRow(0)
'    Call subSetSVPath
    
    'ラベルイメージをクリア
    Call subDeleteLabel
    stHinaLabel2.Copy After:=Worksheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Visible = True
    Set ST = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ST.Name = "Label"

    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & " FROM LIBWMF.WGIP01"
    strSQL = strSQL & " WHERE GIDELT <> 'X'"
    strSQL = strSQL & "   AND GIKBN = '2' "
    strSQL = strSQL & "   AND NOT GIPROJ = '9' "
    strSQL = strSQL & " ORDER BY GIKJCD, GIHNO "
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lRow = 1
    Do While Not RS.EOF
        Do While (True)
            If errFLG Then Exit Do
            bFind = False
            For lRow2 = 1 To lMaxRow
                If RS("GIHNO") = ST2.Cells(lRow2, 1) And fncGetKJNM(RS("GIKJCD")) = ST2.Cells(lRow2, 2) And ST2.Cells(lRow2, 5) = "○" Then bFind = True: Exit For
            Next
            If Not bFind Then Exit Do
            bHUp = False
            Select Case RS("GIPROJ")
            Case "1"
                bHUp = True
                STHina.Cells(2, 2) = RS("GIHINM")
                STHina.Cells(3, 2) = RS("GIJTI2")
                For i = 1 To 5
                    STHina.Cells(5 + i, 2) = RS("GIGNM" & i)
                Next
                STHina.Cells(11, 2) = "㈱ベーカリーシステム研究所　" & fncGetKJMark(RS("GIKJCD"))
                Set r = STHina.Range(STHina.Cells(1, 1), STHina.Cells(13, 4))
                lWk = 14
            Case "3"
                STHina.Cells(33, 2) = RS("GIHINM")
                STHina.Cells(34, 2) = RS("GIJTI2")
                For i = 1 To 8
                    STHina.Cells(36 + i, 2) = RS("GIGNM" & i)
                Next
                STHina.Cells(45, 2) = "㈱ベーカリーシステム研究所　" & fncGetKJMark(RS("GIKJCD"))
                Set r = STHina.Range(STHina.Cells(32, 1), STHina.Cells(47, 4))
                lWk = 17
            Case "5", "10"
                bHUp = True
                STHina.Cells(16, 2) = RS("GIHINM")
                STHina.Cells(17, 2) = RS("GIJTI2")
                For i = 1 To 8
                    STHina.Cells(19 + i, 2) = RS("GIGNM" & i)
                Next
                STHina.Cells(28, 2) = "㈱ベーカリーシステム研究所　" & fncGetKJMark(RS("GIKJCD"))
                Set r = STHina.Range(STHina.Cells(15, 1), STHina.Cells(30, 4))
                lWk = 17
            Case Else: Exit Do
            End Select
            Err.Clear
            On Error Resume Next
            STHina.Range(r.Address).Copy ST.Cells(lRow, 1)
            If bHUp Then
                ST.Rows(lRow + 3).RowHeight = 23.5
            End If
            For i = 1 To 3
                Err.Number = 0
                Application.Wait [now()] + 150 / 86400000
                STHina.Range(r.Address).CopyPicture
                If Err.Number = 0 Then Exit For
            Next
            If Err.Number <> 0 Then
                errFLG = True: Exit Do
            End If
            For i = 1 To 3
                Err.Number = 0
                Application.Wait [now()] + 150 / 86400000
                ST.Cells(lRow, 6).PasteSpecial
                If Err.Number = 0 Then Exit For
            Next
            If Err.Number <> 0 Then
                errFLG = True: Exit Do
            End If
            On Error GoTo 0
            ST.Shapes(ST.Shapes.Count).Height = ST.Range(ST.Cells(lRow, 6), ST.Cells(lRow + lWk, 6)).Height
            lRow = lRow + lWk + 2
            Cnt = Cnt + 1
            If RS.RecordCount = Cnt Then Exit Do
            If Cnt Mod 3 = 0 Then
                ReDim Preserve arrRow(UBound(arrRow) + 1)
                arrRow(UBound(arrRow)) = lRow
            End If
            Exit Do
        Loop
        On Error GoTo 0
        Application.CutCopyMode = False
        If errFLG Then Exit Do
        RS.MoveNext
    Loop

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    If Not errFLG Then
        '対象シートを表示していないとページブレイクで落ちる
        Application.ScreenUpdating = True
        ST.Select
        Application.ScreenUpdating = False
        ActiveWindow.View = xlPageBreakPreview
        ST.PageSetup.PrintArea = ST.Range(ST.Cells(1, 1), ST.Cells(lRow - 1, 16)).Address
        ST.DisplayPageBreaks = True
        ST.DisplayAutomaticPageBreaks = False
        For i = 1 To UBound(arrRow)
            If lRow - 1 > arrRow(i) Then
                ST.Rows(arrRow(i)).PageBreak = xlPageBreakManual
            End If
        Next
        ST.Activate
        
        ST.Cells(1, 1).Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    
    Set ST = Nothing: Set STHina = Nothing
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    If errFLG Then MsgBox "ラベルの作成に失敗した製品があります。やり直してください"

End Sub



