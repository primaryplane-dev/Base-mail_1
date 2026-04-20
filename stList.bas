Option Explicit

Private Sub cmdBack_Click()
    stMenu.Select
End Sub

Private Sub cmdLayout9_Click()
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    Call subMainLay9
    stLay9.Select
End Sub

Private Sub cmdUpdate_Click()
    Dim strWk           As String
    Dim lWk             As Long
    Dim lWk2            As Long
    Dim arrChangeKJ()   As String
    Dim i               As Long
    Dim j               As Long
    Dim delRec()        As GIrec
    Dim strBody         As String
    Dim fromAddress     As String
    Dim toAddress       As String
    
    '入力チェック,変更チェック
    ReDim arrChangeKJ(0): ReDim delRec(0)
    If Not fncChk(arrChangeKJ) Then Exit Sub
    If UBound(arrChangeKJ) = 0 Then Exit Sub
    
    '排他ファイルの存在確認
    strWk = ""
    For i = 1 To UBound(arrChangeKJ)
        lWk = 0
        lWk = fncEditChk(arrChangeKJ(i))
        If lWk2 < lWk Then lWk2 = lWk
        If lWk > 0 Then
            If Not strWk = "" Then strWk = strWk & ","
            strWk = strWk & fncGetKJNM(arrChangeKJ(i))
        End If
    Next
    If lWk2 > 0 Then
        MsgBox (strWk & "でCSVファイルを作成中です。")
        'BASEエクセルで作成された排他ファイルの場合、強制削除ができる
        If lWk2 = 1 Then
            If MsgBox("排他ファイルを強制的に削除してCSVファイルを作成しますか？", vbYesNo) = vbNo Then
                Exit Sub
            End If
            '排他ファイル強制削除
            For i = 1 To UBound(arrChangeKJ)
                Call subDeleteEditFile(arrChangeKJ(i))
            Next
        Else
            Exit Sub
        End If
    End If
    
    '他PCで排他ロックがかかる前に先に排他ファイルを作成しておく
    For i = 1 To UBound(arrChangeKJ)
        '排他ファイルを作成
        Call subMakeEditFile(arrChangeKJ(i))
    Next
    'ASへの更新
    Call subSetData(delRec)
    
    Dim FSO     As New Scripting.FileSystemObject
    For i = 1 To UBound(arrChangeKJ)
        'CSVを作成
        Call subWriteCSV(arrChangeKJ(i), FSO)
        'CSVをファイルサーバと工場の共有フォルダに送る
        Call subSendCSV(arrChangeKJ(i), FSO)
        '排他ファイルを削除
        Call subDeleteEditFile(arrChangeKJ(i))
    Next
    Set FSO = Nothing
    
    
'削除があった工場にメール送信　20260407 修正(送信元選択フォームの表示)
For i = 1 To UBound(arrChangeKJ)
    Debug.Print "arrChangeKJ(" & i & ")=" & arrChangeKJ(i)

    strBody = ""
    For j = 1 To UBound(delRec)
        If delRec(j).KCD = arrChangeKJ(i) Then
            If Not strBody = "" Then strBody = strBody & vbLf
            strBody = strBody & fncMakebody2(1, delRec(j).KCD, delRec(j).HINM)
        End If
    Next
    If Not strBody = "" Then
        fromAddress = "": toAddress = ""
        ' 送信元選択フォームを表示
        Dim frm As New frmFromAddress
        frm.KCD = arrChangeKJ(i)   ' ← 工場コードをセット
        frm.Show vbModal
        If frm.Tag <> "" Then
            fromAddress = frm.Tag
        Else
            ' キャンセル時はメール送信しない
            Set frm = Nothing
            GoTo ContinueNextFactory
        End If
        Unload frm
        Set frm = Nothing
        Call subGetMailAdd(arrChangeKJ(i), fromAddress, toAddress)
        Call subSendMail(fromAddress, toAddress, "", "IJPデータ削除通知", strBody)
    End If
ContinueNextFactory:
Next
    
       
    'データ再表示
    Call subMain
    

End Sub

Private Sub cmdLabel_Click()
'    Call subSetSVPath
    If Not fncExistData Then MsgBox "データが存在しません": Exit Sub
    Call subMainLabel
End Sub

'入力チェックと変更チェック
Private Function fncChk(ByRef arrChangeKJ() As String) As Boolean
    Dim lMaxRow     As Long: lMaxRow = Me.Cells(stList.Rows.Count, 1).End(xlUp).Row
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim lCol        As Long
    Dim bEdit       As Boolean
    Dim bExist      As Boolean
    Dim i           As Long
    Dim CN          As New ADODB.Connection
    fncChk = False
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    '入力チェック
    For lRow = 3 To lMaxRow
        If Not Trim(Me.Cells(lRow, 1)) = "" Then
            If Not IsNumeric(Me.Cells(lRow, 1)) Then MsgBox "コードは数値のみ": GoTo Exit_
            If Not fncFindCode(CN, Trim(Me.Cells(lRow, 1))) Then MsgBox "コード'" & Trim(Me.Cells(lRow, 1)) & "'は存在しません": GoTo Exit_
        End If
    Next
    
    '変更チェック(ワークシートと比較)
    ReDim arrChangeKJ(0)
    lRow = 3: lRow2 = 3
    Do While Not stWork.Cells(lRow, 1) = ""
        bEdit = True
        'ワークシートの商品が削除されているか
        For lRow2 = 3 To lMaxRow
            If stWork.Cells(lRow, 1) = Me.Cells(lRow2, 1) And fncGetBUTUKCD2(stWork.Cells(lRow, 2)) = fncGetBUTUKCD2(Me.Cells(lRow2, 2)) Then bEdit = False: Exit For
        Next
        If bEdit Then
            bExist = False
            For i = 1 To UBound(arrChangeKJ)
                If arrChangeKJ(i) = fncGetBUTUKCD2(stWork.Cells(lRow, 2)) Then
                    bExist = True
                    Exit For
                End If
            Next
            If Not bExist Then
                ReDim Preserve arrChangeKJ(UBound(arrChangeKJ) + 1)
                arrChangeKJ(UBound(arrChangeKJ)) = fncGetBUTUKCD2(stWork.Cells(lRow, 2))
            End If
        Else
            '変更されているか
            For lCol = 1 To 4
                If Not stWork.Cells(lRow, lCol) = Me.Cells(lRow2, lCol) Then
                    bExist = False
                    For i = 1 To UBound(arrChangeKJ)
                        If arrChangeKJ(i) = fncGetBUTUKCD2(stWork.Cells(lRow, 2)) Then
                            bExist = True
                            Exit For
                        End If
                    Next
                    If Not bExist Then
                        ReDim Preserve arrChangeKJ(UBound(arrChangeKJ) + 1)
                        arrChangeKJ(UBound(arrChangeKJ)) = fncGetBUTUKCD2(stWork.Cells(lRow, 2))
                    End If
                    Exit For
                End If
            Next
        End If
        lRow = lRow + 1
    Loop
    
    '商品データが削除、変更された工場があるか
    If UBound(arrChangeKJ) = 0 Then GoTo Exit_

    fncChk = True
    
Exit_:
    'ＤＢ切断
    CN.Close: Set CN = Nothing

End Function

Private Function fncFindCode(ByRef CN As ADODB.Connection, ByVal i_CODE As String) As Boolean
    Dim RS      As New ADODB.Recordset
    Dim strSQL  As String
    fncFindCode = False
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM LIBWMF.WGIP01"
    strSQL = strSQL & "  WHERE GIDELT = ''"
    strSQL = strSQL & "    AND (GIHNO = '" & Format(i_CODE, "00000") & "' OR GIKHN1 = '" & Format(i_CODE, "00000") & "')"
    strSQL = strSQL & "    AND GIKBN = '2' "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then fncFindCode = True
    RS.Close: Set RS = Nothing
End Function

Private Sub cmdLabelAll_Click()
    Dim lMaxRow     As Long: lMaxRow = stList.Cells(stList.Rows.Count, 1).End(xlUp).Row
    Dim lRow        As Long
    If Not Me.Cells(2, 5) = "○" Then
        Me.Cells(2, 5) = "○"
        For lRow = 3 To lMaxRow
            If Not Trim(Me.Cells(lRow, 1)) = "" Then
                Me.Cells(lRow, 5) = "○"
            End If
        Next
    Else
        Me.Range(Me.Cells(3, 5), Me.Cells(lMaxRow, 5)) = ""
        Me.Cells(2, 5) = ""
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Row < 3 Then Exit Sub
    If Target.Column > 5 Then Exit Sub
    If Target.Count > 1 Then Exit Sub
    If Me.Cells(Target.Row, 1) = "" Then GoTo Exit_WorkSheet_SelectionChange
    If Target.Column = 4 Then
        P_KCD = fncGetBUTUKCD2(Me.Cells(Target.Row, 2))
        Call subOpen
        If P_Regist Then
            Me.Cells(Target.Row, Target.Column) = P_LINM
        End If
    ElseIf Target.Column = 5 Then
        If Me.Cells(Target.Row, Target.Column) = "○" Then
            Me.Cells(Target.Row, Target.Column) = ""
        Else
            Me.Cells(Target.Row, Target.Column) = "○"
        End If
    End If
Exit_WorkSheet_SelectionChange:
    Me.Cells(2, 1).Select
End Sub

Private Sub subOpen()
    Dim obj As New frmSelectLINM
    obj.Show
    Set obj = Nothing
End Sub

