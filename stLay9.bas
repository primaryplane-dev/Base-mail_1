Option Explicit

Private Sub cmdBack_Click()
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    stList.Select
End Sub

Private Sub cmdEnd_Click()
    Dim arrKJ()     As Variant: arrKJ = Array("001", "002", "102", "103")
    Dim i           As Long
    For i = 0 To UBound(arrKJ)
        Call subDeleteEditFile(arrKJ(i))
    Next
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    If Workbooks.Count = 1 Then Application.Quit
    ThisWorkbook.Close False
End Sub

Private Sub cmdUpdate_Click()
'    Dim blH             As Boolean
'    Dim blM             As Boolean
'    Dim blT             As Boolean
'    Dim blF             As Boolean
    Dim strWk           As String
    Dim lWk             As Long
    Dim lWk2            As Long
    Dim arrChangeKJ()   As String
    Dim i               As Long
    Dim j               As Long
    Dim delRec()        As GIrec
    Dim addRec()        As GIrec
    Dim strBody         As String
    Dim fromAddress     As String
    Dim toAddress       As String
    
    '入力チェック、変更チェック
    ReDim arrChangeKJ(0): ReDim delRec(0): ReDim addRec(0)
    If Not fncChk(arrChangeKJ) Then Exit Sub
    If UBound(arrChangeKJ) = 0 Then Exit Sub
    
    '排他ファイルの存在確認
    strWk = "": lWk = 0: lWk2 = 0
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
    
'    Call subSetSVPath
'    intWk = 0
'    intWk = fncEditChk(blH, blM, blT, blF)
'    If intWk > 0 Then
'        If blH Then strWk = "枚方工場"
'        If blM Then
'            If Not strWk = "" Then strWk = strWk & ","
'            strWk = strWk & "武蔵工場"
'        End If
'        If blT Then
'            If Not strWk = "" Then strWk = strWk & ","
'            strWk = strWk & "タカラ食品"
'        End If
'        If blF Then
'            If Not strWk = "" Then strWk = strWk & ","
'            strWk = strWk & "福岡工場"
'        End If
'        MsgBox (strWk & "でCSVファイルを作成中です。")
'        If intWk = 1 Then
'            If MsgBox("排他ファイルを強制的に削除してCSVファイルを作成しますか？", vbYesNo) = vbYes Then
'                Call subDeleteEditFile
'                GoTo Update
'            End If
'        End If
'        Exit Sub
'    End If
    
    '他PCで排他ロックがかかる前に先に排他ファイルを作成しておく
    For i = 1 To UBound(arrChangeKJ)
        '排他ファイルを作成
        Call subMakeEditFile(arrChangeKJ(i))
    Next
    'ASへの更新
    Call subLay9Update(delRec, addRec)
    
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
    
    
    '削除があった工場にメール送信
    For i = 1 To UBound(arrChangeKJ)
        '削除通知メール
        strBody = ""
        For j = 1 To UBound(delRec)
            If delRec(j).KCD = arrChangeKJ(i) Then
                If Not strBody = "" Then strBody = strBody & vbLf
                strBody = strBody & fncMakebody2(1, delRec(j).KCD, delRec(j).HINM)
            End If
        Next
        If Not strBody = "" Then
            fromAddress = "": toAddress = ""
            Call subGetMailAdd(arrChangeKJ(i), fromAddress, toAddress)
            Call subSendMail(fromAddress, toAddress, "", "IJPデータ削除通知", strBody)
        End If
        '追加通知メール
        strBody = ""
        For j = 1 To UBound(addRec)
            If addRec(j).KCD = arrChangeKJ(i) Then
                If Not strBody = "" Then strBody = strBody & vbLf
                strBody = strBody & fncMakebody2(2, addRec(j).KCD, addRec(j).HINM)
            End If
        Next
        If Not strBody = "" Then
            fromAddress = "": toAddress = ""
            Call subGetMailAdd(arrChangeKJ(i), fromAddress, toAddress)
            Call subSendMail(fromAddress, toAddress, "", "IJPデータ追加通知", strBody)
        End If
    Next
    
    'データ再表示
    Call subMainLay9
    
'Update:
'    Call subMakeEditFile
'    If Not fncChk Then: GoTo Exit_Update
'    Call subLay9Update
'    Dim FSO     As New Scripting.FileSystemObject
'    'CSVを作成する
'    Call subWriteCSV(FSO)
'    'CSVをファイルサーバと工場の共有フォルダに送る
'    Call subSendCSV(FSO)
'    Set FSO = Nothing
'    Call subMainLay9
'    'エラーになった製品があればシートを表示する
'    If Not stError.Cells(3, 1) = "" Then stError.Select
'
'Exit_Update:
'    Call subDeleteEditFile
End Sub

'入力チェックと変更チェック
Private Function fncChk(ByRef arrChangeKJ() As String) As Boolean
    Dim lMaxRow     As Long: lMaxRow = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim lCol        As Long
    Dim strWk       As String
    Dim strWk2      As String
    Dim sSp()       As String
    Dim sSp2()      As String
    Dim i           As Long
    Dim bEdit       As Boolean
    Dim bExist      As Boolean

    fncChk = False
    
    '入力チェック
    For lRow = 3 To lMaxRow
        If Not Trim(Me.Cells(lRow, 1)) = "" Then
            If LenAS(Trim(Me.Cells(lRow, 1))) > 100 Then MsgBox "品名は全角49文字以内": Exit Function
            If Trim(Me.Cells(lRow, 2)) = "" Then MsgBox "入数を入力してください": Exit Function
            strWk = StrConv(Trim(Me.Cells(lRow, 2)), vbNarrow)
            strWk = StrConv(strWk, vbLowerCase)
            strWk = Replace(strWk, "kg", "")
            strWk = Replace(strWk, "㎏", "")
            If InStr(strWk, "/") > 0 Then
                sSp = Split(strWk, "/")
                strWk2 = ""
                For i = 0 To UBound(sSp)
                    sSp2 = Split(sSp(i), ".")
                    If UBound(sSp2) = 0 Then
                        If Len(sSp2(0)) > 3 Then MsgBox "入数は整数部分3桁、小数点以下2桁以内": Exit Function
                    Else
                        If Len(sSp2(0)) > 3 Or Len(sSp2(1)) > 2 Then MsgBox "入数は整数部分3桁、小数点以下2桁以内": Exit Function
                    End If
                    If strWk2 = "" Then
                        strWk2 = sSp(i)
                    Else
                        strWk2 = Val(strWk2) * Val(sSp(i))
                    End If
                Next
                sSp2 = Split(strWk2, ".")
                If UBound(sSp2) = 0 Then
                    If Len(sSp2(0)) > 3 Then MsgBox "入数の総数は整数部分3桁、小数点以下2桁以内": Exit Function
                Else
                    If Len(sSp2(0)) > 3 Or Len(sSp2(1)) > 2 Then MsgBox "入数の総数は整数部分3桁、小数点以下2桁以内": Exit Function
                End If
            Else
                sSp = Split(strWk, ".")
                If UBound(sSp) = 0 Then
                    If Len(sSp(0)) > 3 Then MsgBox "入数は整数部分3桁、小数点以下2桁以内": Exit Function
                Else
                    If Len(sSp(0)) > 3 Or Len(sSp(1)) > 2 Then MsgBox "入数は整数部分3桁、小数点以下2桁以内": Exit Function
                End If
            End If
            If Trim(Me.Cells(lRow, 3)) = "" Then MsgBox "賞味期限日数を入力してください": Exit Function
            If Not IsNumeric(StrConv(Me.Cells(lRow, 3), vbNarrow)) Then MsgBox "賞味期限日数は数値のみ": Exit Function
            If Len(StrConv(Me.Cells(lRow, 3), vbNarrow)) > 3 Then MsgBox "賞味期限日数は3桁以内の数値のみ": Exit Function
            If Trim(Me.Cells(lRow, 4)) = "" Then MsgBox "工場を入力してください": Exit Function
            If Trim(Me.Cells(lRow, 5)) = "" Then MsgBox "ライン名を入力してください": Exit Function
            If Trim(Me.Cells(lRow, 6)) = "" Then MsgBox "製造品番(ロジ番)を入力してください": Exit Function
            If Not IsNumeric(StrConv(Me.Cells(lRow, 6), vbNarrow)) Or Len(Me.Cells(lRow, 6)) > 5 Then MsgBox "製造品番(ロジ番)は5桁以内の数値のみ": Exit Function
            If Trim(Me.Cells(lRow, 7)) = "" Then MsgBox "JANコードを入力してください": Exit Function
            If Not IsNumeric(StrConv(Me.Cells(lRow, 7), vbNarrow)) Or Not Len(Me.Cells(lRow, 7)) = 13 Then MsgBox "JANは13桁の数値のみ": Exit Function
            
        End If
    Next
    
    '変更チェック(ワークシートと比較)
    ReDim arrChangeKJ(0)
    lRow = 3: lRow2 = 3
    Do While Not stWorkLay9.Cells(lRow, 1) = ""
        bEdit = True
        'ワークシートの商品が削除されているか
        For lRow2 = 3 To lMaxRow
            If Trim(stWorkLay9.Cells(lRow, 6)) = Trim(Me.Cells(lRow2, 6)) And fncGetBUTUKCD2(Trim(stWorkLay9.Cells(lRow, 4))) = fncGetBUTUKCD2(Trim(Me.Cells(lRow2, 4))) Then bEdit = False: Exit For
        Next
        If bEdit Then
            bExist = False
            For i = 1 To UBound(arrChangeKJ)
                If arrChangeKJ(i) = fncGetBUTUKCD2(Trim(stWorkLay9.Cells(lRow, 4))) Then
                    bExist = True
                    Exit For
                End If
            Next
            If Not bExist Then
                ReDim Preserve arrChangeKJ(UBound(arrChangeKJ) + 1)
                arrChangeKJ(UBound(arrChangeKJ)) = fncGetBUTUKCD2(Trim(stWorkLay9.Cells(lRow, 4)))
            End If
        Else
            '変更されているか
            For lCol = 1 To 7
                If Not Trim(stWorkLay9.Cells(lRow, lCol)) = Trim(Me.Cells(lRow2, lCol)) Then
                    bExist = False
                    For i = 1 To UBound(arrChangeKJ)
                        If arrChangeKJ(i) = fncGetBUTUKCD2(Trim(stWorkLay9.Cells(lRow, 4))) Then
                            bExist = True
                            Exit For
                        End If
                    Next
                    If Not bExist Then
                        ReDim Preserve arrChangeKJ(UBound(arrChangeKJ) + 1)
                        arrChangeKJ(UBound(arrChangeKJ)) = fncGetBUTUKCD2(Trim(stWorkLay9.Cells(lRow, 4)))
                    End If
                    Exit For
                End If
            Next
        End If
        lRow = lRow + 1
    Loop
    
    '商品データが削除、変更された工場があるか
    If UBound(arrChangeKJ) = 0 Then Exit Function

    fncChk = True
End Function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Row < 3 Then Exit Sub
    If Target.Count > 1 Then Exit Sub
    If Me.Cells(Target.Row, 1) = "" Then Exit Sub
    If Target.Column = 4 Then   '工場選択
        Call subOpenKJ
        If P_Regist2 Then
            Me.Cells(Target.Row, Target.Column) = P_KJNM
        End If
        Me.Cells(Target.Row, 3).Select
    ElseIf Target.Column = 5 Then   'ライン選択
        If Me.Cells(Target.Row, 4) = "" Then Exit Sub
        P_KCD = fncGetBUTUKCD2(Me.Cells(Target.Row, 4))
        Call subOpenLI
        If P_Regist Then
            Me.Cells(Target.Row, Target.Column) = P_LINM
        End If
        Me.Cells(Target.Row, 3).Select
    End If
End Sub

Private Sub subOpenKJ()
    Dim obj As New frmSelectKJNM
    obj.Show
    Set obj = Nothing
End Sub

Private Sub subOpenLI()
    Dim obj As New frmSelectLINM
    obj.Show
    Set obj = Nothing
End Sub


