Option Explicit

Private Sub subIniError()
    stErrorHina.Cells.Copy stError.Cells
    '書式を登録(excel2010だと17行以降の書式がリセットされるため)
    stError.Range(stError.Cells(3, 1), stError.Cells(stError.Rows.Count, 2)).NumberFormatLocal = "@"
    stError.Cells(3, 1).HorizontalAlignment = xlCenter
'    stError.Range(stError.Cells(3, 1), stError.Cells(stError.Rows.Count, 2)).HorizontalAlignment = xlCenter
    stError.Range(stError.Cells(3, 1), stError.Cells(stError.Rows.Count, 2)).Font.Size = 12
End Sub

Public Function fncGetUpdData(ByRef CN As ADODB.Connection, ByRef arrChangeKJ() As String, ByRef rec() As GIrec) As Boolean
    Dim BK          As Workbook
    Dim ST          As Worksheet
    Dim BK2         As Workbook
    Dim ST2         As Worksheet
    Dim RS          As ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim strWk       As String
    Dim strWk2      As String
    Dim i           As Integer
    Dim j           As Integer
    Dim buf         As String
    Dim lWk         As Long
    Dim strPath     As String
    Dim sSp()       As String
    Dim bExist      As Boolean
    Dim bExist2     As Boolean
    
    fncGetUpdData = False
    Call subBeforeEdit
    On Error GoTo ErrorHandler

    buf = Dir(ThisWorkbook.Path & "\工場送付用(JAN抽出)*.xlsx")
    strWk = ""
    lWk = 0
    Do While buf <> ""
        strWk = Replace(buf, "工場送付用(JAN抽出)", "")
        If lWk < Val(strWk) Then lWk = Val(strWk)
        buf = Dir()
    Loop
    
    On Error GoTo ErrorHandler
    Set BK = Workbooks.Open(ThisWorkbook.Path & "\工場送付用(JAN抽出)" & lWk & ".xlsx")
    Set ST = BK.Sheets(1)
    
    'エラー出力シートの初期化
    Call subIniError
    
    ReDim rec(0)
    ReDim arrChangeKJ(0)
    
    lRow = 2
    lRow2 = 3
    strWk = ""
    'カルテの存在チェック
    Do While Not ST.Cells(lRow, 1) = ""
        buf = ""
        strPath = P_KarutePath
        On Error GoTo ErrorHandler
        buf = Dir(strPath & StrConv(ST.Cells(lRow, 1), vbNarrow) & "*.xls")
        If buf = "" Then
            MsgBox ("品番:" & StrConv(ST.Cells(lRow, 1), vbNarrow) & "カルテ用データが見つかりません")
            GoTo Exit_
        End If
        lRow = lRow + 1
    Loop
    
    strSQL = ""
    strSQL = strSQL & "SELECT GIHNO "
    strSQL = strSQL & "  FROM LIBWMF.WGIP01 "
    strSQL = strSQL & " WHERE GIDELT = '' "
    strSQL = strSQL & "   AND GIKBN = '2' "
    strSQL = strSQL & "   AND NOT GIPROJ = '9' "
    strSQL = strSQL & " GROUP BY GIHNO "
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lRow = 2
    Do While Not ST.Cells(lRow, 1) = ""
        bExist = False
        RS.MoveFirst
        Do While Not RS.EOF
            If Val(StrConv(ST.Cells(lRow, 1), vbNarrow)) = Val(RS("GIHNO")) Then
                bExist = True
                Exit Do
            End If
            RS.MoveNext
        Loop
        If Not bExist Then
            strPath = P_KarutePath
            buf = Dir(strPath & StrConv(ST.Cells(lRow, 1), vbNarrow) & "*.xls")
            Do While buf <> ""
                On Error GoTo ErrorHandler
                Set BK2 = Workbooks.Open(strPath & buf)
                Set ST2 = BK2.Sheets("商品カルテ用データ")
                bExist2 = False
                For i = 1 To UBound(arrChangeKJ)
                    If arrChangeKJ(i) = fncGetBUTUKCD2(Trim(ST2.Cells(17, 2))) Then
                        bExist2 = True
                        Exit For
                    End If
                Next
                If Not bExist2 Then
                    ReDim Preserve arrChangeKJ(UBound(arrChangeKJ) + 1)
                    arrChangeKJ(UBound(arrChangeKJ)) = fncGetBUTUKCD2(Trim(ST2.Cells(17, 2)))
                End If
                ReDim Preserve rec(UBound(rec) + 1)
                rec(UBound(rec)).HNO = Format(Trim(ST.Cells(lRow, 1)), "00000")
                rec(UBound(rec)).JAN = Trim(ST.Cells(lRow, 3))
                rec(UBound(rec)).KHN1 = Format(Trim(ST.Cells(lRow, 4)), "00000")
                rec(UBound(rec)).HINM = Trim(ST2.Cells(1, 3).Text)
                strWk = Trim(ST2.Cells(45, 2).Text)
                Call subChkMojibake(strWk)
                strWk = StrConv(strWk, vbWide)      '2024/04/18 Hirata Add
                '----------------------------------------------------------2024/05/28 Add
                sSp = Split(strWk, vbLf & vbLf)
                j = 0
                For i = 0 To UBound(sSp)
                    strWk = sSp(i)
                    strWk = Replace(strWk, vbLf, "")
                    For j = j + 1 To 8
                        rec(UBound(rec)).GNM(j) = Left(strWk, 35)
                        strWk = Replace(strWk, rec(UBound(rec)).GNM(j), "")
                        If strWk = "" Then Exit For
                    Next
                Next
                i = j
                '----------------------------------------------------------
    '            For i = 1 To 8
    '                rec(UBound(rec)).GNM(i) = Left(strWk, 35)
    '                strWk = Replace(strWk, rec(UBound(rec)).GNM(i), "")
    '                If strWk = "" Then Exit For
    '            Next
                If Trim(ST2.Cells(8, 2).Text) = "焼成冷凍" Then
                    rec(UBound(rec)).PROJ = "3"
                Else
                    rec(UBound(rec)).PROJ = "1"
                    If i >= 6 Then rec(UBound(rec)).PROJ = "5"
                End If
                If ST2.Cells(15, 2).Text = "Amazon" Then
                    rec(UBound(rec)).PROJ = "10"
                End If
                If InStr(Trim(ST2.Cells(4, 3)), "kg") > 0 Or InStr(Trim(ST2.Cells(4, 3)), "Kg") > 0 Or InStr(Trim(ST2.Cells(4, 3)), "KG") > 0 Or InStr(Trim(ST2.Cells(4, 3)), "㎏") > 0 Then
                    rec(UBound(rec)).TNI = "1"
                End If
                rec(UBound(rec)).JTI(1) = Val(Trim(ST2.Cells(4, 3)))
                strWk = Trim(ST2.Cells(7, 3))
                If strWk = "対応なし" Then strWk = ""
                If InStr(strWk, ",") > 0 Or InStr(strWk, "、") > 0 Then strWk = ""
                If strWk = "" Then
                    strWk = rec(UBound(rec)).JTI(1) & IIf(rec(UBound(rec)).TNI = "1", "kg", "個")
                Else
                    strWk = rec(UBound(rec)).JTI(1) & IIf(rec(UBound(rec)).TNI = "1", "kg", "個") & "(" & strWk & ")"
                End If
                rec(UBound(rec)).JTI(2) = strWk
                rec(UBound(rec)).HHI = Trim(ST2.Cells(11, 2).Text)
                rec(UBound(rec)).KCD = fncGetBUTUKCD2(Trim(ST2.Cells(17, 2)))
                rec(UBound(rec)).LINO = fncGetLINO(rec(UBound(rec)).KCD, Trim(ST2.Cells(22, 2)))
                rec(UBound(rec)).LINM = IIf(rec(UBound(rec)).LINO = "", "", Trim(ST2.Cells(22, 2)))
                strWk = Trim(ST2.Cells(55, 2).Text)
                strWk = StrConv(strWk, vbWide)      '2024/04/19 Hayashi Add
                For i = 1 To 2
                    rec(UBound(rec)).YOB(i) = Left(strWk, 50)
                    strWk = Replace(strWk, rec(UBound(rec)).YOB(i), "")
                    If strWk = "" Then Exit For
                Next
                If Not strWk = "" Then
                    stError.Cells(lRow2, 1) = rec(UBound(rec)).HNO
                    stError.Cells(lRow2, 2) = "特記事項を100文字以内に変更してください"
                    stError.Range(stError.Cells(1, 1), stError.Cells(lRow2, 2)).Borders.LineStyle = xlContinuous
                    lRow2 = lRow2 + 1
                End If
                If Not Trim(ST.Cells(lRow, 2)) = rec(UBound(rec)).HINM Then
                    stError.Cells(lRow2, 1) = rec(UBound(rec)).HNO
                    stError.Cells(lRow2, 2) = "冷凍生地名が異なります"
                    stError.Range(stError.Cells(1, 1), stError.Cells(lRow2, 2)).Borders.LineStyle = xlContinuous
                    lRow2 = lRow2 + 1
                End If
                BK2.Saved = True: BK2.Close
                buf = Dir()
            Loop
        End If
        lRow = lRow + 1
    Loop
    Set ST2 = Nothing: Set BK2 = Nothing
    
    'ＤＢ接続
    'CN.CursorLocation = adUseClient
    'CN.Open P_ConnectString
    
    'Delete
    'Deleteは行わない、追加のみに変更 20250827
    'Call subDelUpdate(CN)
    
    'For i = 1 To UBound(rec)
    '    Call subUpdate(CN, rec(i))
    'Next
    'ＤＢ切断
    'CN.Close: Set CN = Nothing
    fncGetUpdData = True
    
Exit_:
    '早期リターン時もクリーンアップ
    GoTo CleanUp

CleanUp:
    On Error Resume Next
    If Not RS Is Nothing Then If RS.State = 1 Then RS.Close: Set RS = Nothing
    If Not BK2 Is Nothing Then BK2.Saved = True: BK2.Close: Set BK2 = Nothing
    If Not BK Is Nothing Then BK.Saved = True: BK.Close: Set BK = Nothing
    Set ST2 = Nothing: Set ST = Nothing
    Call subAfterEdit
    Exit Function

ErrorHandler:
    MsgBox "fncGetUpdDataでエラー発生: " & Err.Description, vbCritical
    Resume CleanUp
End Function

Private Sub subChkMojibake(ByRef strGNM As String)
    Dim lRow        As Long
    lRow = 2
    Do While Not stMojibake.Cells(lRow, 1) = ""
        strGNM = Replace(strGNM, stMojibake.Cells(lRow, 1), stMojibake.Cells(lRow, 2))
        lRow = lRow + 1
    Loop
End Sub

Public Sub subSetData(ByRef delRec() As GIrec)
    Dim lMaxRow     As Long: lMaxRow = stList.Cells(stList.Rows.Count, 1).End(xlUp).Row
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As New ADODB.Connection
    Dim strSQL      As String
    Dim i           As Long
    Dim bdltFLG     As Boolean
    Dim strLINO     As String
    Dim lResult     As Long
    
    Call subBeforeEdit
    
'    Call subSetSVPath
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    lRow = 3
    ReDim delRec(0)
    'ワークシートと比較
    Do While Not stWork.Cells(lRow, 1) = ""
        bdltFLG = True
        For i = 1 To lMaxRow
            If ST.Cells(i, 1) = stWork.Cells(lRow, 1) And fncGetBUTUKCD2(ST.Cells(i, 2)) = fncGetBUTUKCD2(stWork.Cells(lRow, 2)) Then bdltFLG = False: Exit For
        Next
        If bdltFLG Then
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBWMF.WGIP01 SET"
            strSQL = strSQL & "     GIUPHZ =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,GIUPTM =TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "    ,GIDELT = 'X'"
            strSQL = strSQL & " WHERE GIDELT = ''"
            strSQL = strSQL & "   AND GIHNO = '" & stWork.Cells(lRow, 1) & "' "
            strSQL = strSQL & "   AND GIKJCD = '" & fncGetBUTUKCD2(stWork.Cells(lRow, 2)) & "' "
            strSQL = strSQL & "   AND GIKBN = '2' "
            strSQL = strSQL & "   AND NOT GIPROJ = '9' "
            CN.Execute strSQL, lResult, &H80
            ReDim Preserve delRec(UBound(delRec) + 1)
            delRec(UBound(delRec)).KCD = fncGetBUTUKCD2(stWork.Cells(lRow, 2))
            delRec(UBound(delRec)).HINM = stWork.Cells(lRow, 3)
        End If
        lRow = lRow + 1
    Loop
    
    For lRow = 3 To lMaxRow
        If Not Trim(ST.Cells(lRow, 1)) = "" Then
            strLINO = fncGetLINO(fncGetBUTUKCD2(ST.Cells(lRow, 2)), ST.Cells(lRow, 4))
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBWMF.WGIP01 SET"
            strSQL = strSQL & "     GIUPHZ =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,GIUPTM =TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "    ,GILINO = '" & strLINO & "'"
            strSQL = strSQL & "    ,GILINM = '" & IIf(strLINO = "", "", Trim(ST.Cells(lRow, 4))) & "'"
            strSQL = strSQL & " WHERE GIDELT = ''"
            strSQL = strSQL & "   AND GIHNO = '" & ST.Cells(lRow, 1) & "' "
            strSQL = strSQL & "   AND GIKJCD = '" & fncGetBUTUKCD2(ST.Cells(lRow, 2)) & "' "
            strSQL = strSQL & "   AND GIKBN = '2' "
            strSQL = strSQL & "   AND NOT GIPROJ = '9' "
            CN.Execute strSQL, lResult, &H80
            '新規登録の場合
            If lResult = 0 Then
            End If
        End If
    Next
    
    CN.Close: Set CN = Nothing

    Set ST = Nothing
    Call subAfterEdit
    
End Sub

Public Sub subUpdate(ByRef CN As ADODB.Connection, ByRef rec As GIrec)
    Dim strSQL      As String
    Dim i           As Integer
    Dim lResult     As Long
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO LIBWMF.WGIP01 "
    strSQL = strSQL & "         ("
    strSQL = strSQL & "              GIDELT"
    strSQL = strSQL & "             ,GICRHZ"
    strSQL = strSQL & "             ,GICRTM"
    strSQL = strSQL & "             ,GIKJCD"
    strSQL = strSQL & "             ,GIKJNM"
    strSQL = strSQL & "             ,GIHNO"
    strSQL = strSQL & "             ,GIKHN1"
    strSQL = strSQL & "             ,GIHINM"
    For i = 1 To 8
        strSQL = strSQL & "             ,GIGNM" & i
    Next
    strSQL = strSQL & "             ,GIHYNO"
    For i = 1 To 2
        strSQL = strSQL & "             ,GIJTI" & i
    Next
    strSQL = strSQL & "             ,GIJTKB"
    strSQL = strSQL & "             ,GIHHI"
    strSQL = strSQL & "             ,GISHI"
    strSQL = strSQL & "             ,GIJAN"
    strSQL = strSQL & "             ,GILINO"
    strSQL = strSQL & "             ,GILINM"
    strSQL = strSQL & "             ,GIPROJ"
    For i = 1 To 20
        strSQL = strSQL & "             ,GIYB" & Format(i, "00")
    Next
    strSQL = strSQL & "             ,GIKBN"
    strSQL = strSQL & "         ) VALUES ("
    strSQL = strSQL & "              ''"
    strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'YYYYMMDD')"
    strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'HH24MISS')"
    strSQL = strSQL & "             , '" & rec.KCD & "'"
    strSQL = strSQL & "             , '" & fncGetKJNM(rec.KCD) & "'"
    strSQL = strSQL & "             , '" & rec.HNO & "'"
    strSQL = strSQL & "             , '" & rec.KHN1 & "'"
    strSQL = strSQL & "             , '" & rec.HINM & "'"
    For i = 1 To 8
        strSQL = strSQL & "             , '" & rec.GNM(i) & "'"
    Next
    strSQL = strSQL & "             , '" & rec.HYNO & "'"
    strSQL = strSQL & "             , " & Val(rec.JTI(1))
    strSQL = strSQL & "             , '" & rec.JTI(2) & "' "
    strSQL = strSQL & "             , '" & rec.TNI & "'"
    strSQL = strSQL & "             , " & Val(rec.HHI)
    strSQL = strSQL & "             , " & IIf(Val(rec.HHI) >= 90, 30, 0)
    strSQL = strSQL & "             , '" & rec.JAN & "'"
    strSQL = strSQL & "             , '" & rec.LINO & "'"
    strSQL = strSQL & "             , '" & rec.LINM & "'"
    strSQL = strSQL & "             , '" & rec.PROJ & "'"
    For i = 1 To 20
        strSQL = strSQL & "             , '" & rec.YOB(i) & "'"
    Next
    strSQL = strSQL & "             , '2'"
    strSQL = strSQL & "         )"
    
    CN.Execute strSQL, lResult, &H80
End Sub

'Public Sub subDelUpdate(ByRef CN As ADODB.Connection)
'    Dim lResult     As Long: lResult = 0
'    Dim strSQL      As String
'
'    strSQL = ""
'    strSQL = strSQL & "UPDATE LIBWMF.WGIP01 SET"
'    strSQL = strSQL & "     GIUPHZ =TO_CHAR(current timestamp, 'YYYYMMDD')"
'    strSQL = strSQL & "    ,GIUPTM =TO_CHAR(current timestamp, 'HH24MISS')"
'    strSQL = strSQL & "    ,GIDELT = 'X'"
'    strSQL = strSQL & " WHERE GIDELT = ''"
'    strSQL = strSQL & "   AND GIKBN = '2' "
'    strSQL = strSQL & "   AND NOT GIPROJ = '9' "
'
'    CN.Execute strSQL, lResult, &H80
'End Sub

Public Sub subLay9Update(ByRef delRec() As GIrec, ByRef addRec() As GIrec)
    Dim lMaxRow     As Long: lMaxRow = stLay9.Cells(stLay9.Rows.Count, 1).End(xlUp).Row
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stLay9
    Dim STW         As Worksheet: Set STW = stWorkLay9
    Dim CN          As New ADODB.Connection
    Dim strSQL      As String
    Dim i           As Long
    Dim bdltFLG     As Boolean
    Dim lResult     As Long
    Dim strWk       As String
    Dim strWk2      As String
    Dim sSp()       As String
    Dim strJTKB     As String
    Dim strLINO     As String
    
    Call subBeforeEdit
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    ReDim delRec(0)
    ReDim addRec(0)
    lRow = 3
    Do While Not STW.Cells(lRow, 1) = ""
        bdltFLG = True
        For i = 3 To lMaxRow
            'ワークシートの品番と工場が一致するものがなければ削除
            If Trim(ST.Cells(i, 6)) = Trim(STW.Cells(lRow, 6)) And Trim(ST.Cells(i, 4)) = Trim(STW.Cells(lRow, 4)) Then bdltFLG = False: Exit For
        Next
        If bdltFLG Then
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBWMF.WGIP01 SET"
            strSQL = strSQL & "     GIUPHZ =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,GIUPTM =TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "    ,GIDELT = 'X'"
            strSQL = strSQL & " WHERE GIDELT = ''"
            strSQL = strSQL & "   AND GIKHN1 = '" & Trim(STW.Cells(lRow, 6)) & "' "
            strSQL = strSQL & "   AND GIKJCD = '" & fncGetBUTUKCD2(Trim(STW.Cells(lRow, 4))) & "' "
            strSQL = strSQL & "   AND GIKBN = '2' "
            strSQL = strSQL & "   AND GIPROJ = '9'"
            CN.Execute strSQL, lResult, &H80
            ReDim Preserve delRec(UBound(delRec) + 1)
            delRec(UBound(delRec)).HINM = Trim(STW.Cells(lRow, 1))
            delRec(UBound(delRec)).KCD = fncGetBUTUKCD2(Trim(STW.Cells(lRow, 4)))
        End If
        lRow = lRow + 1
    Loop
    
    For lRow = 3 To lMaxRow
        strWk = "": strWk2 = "": strJTKB = "": strLINO = ""
        If Not Trim(ST.Cells(lRow, 1)) = "" Then
            strWk = StrConv(Trim(ST.Cells(lRow, 2)), vbNarrow)
            strWk = StrConv(strWk, vbLowerCase)
            If InStr(strWk, "/") > 0 Then                   '2024/12/10
                strWk2 = strWk
                strWk2 = Replace(strWk2, "㎏", "kg")
                If InStr(strWk, "kg") > 0 Or InStr(strWk, "㎏") > 0 Then
                    strWk = Replace(strWk, "kg", "")
                    strWk = Replace(strWk, "㎏", "")
                    strJTKB = "1"
                End If
                sSp = Split(strWk, "/")
                For i = 0 To UBound(sSp)
                    If i = 0 Then
                        strWk = Val(sSp(i))
                    Else
                        strWk = Val(strWk) * Val(sSp(i))
                    End If
                Next
            Else
                If InStr(strWk, "kg") > 0 Or InStr(strWk, "㎏") > 0 Then
                    strWk2 = strWk
                    strWk2 = Replace(strWk2, "㎏", "kg")
                    strWk = Replace(strWk, "kg", "")
                    strWk = Replace(strWk, "㎏", "")
                    strJTKB = "1"
                Else
                    strWk2 = Val(strWk)
'                    If ST.Cells(lRow, 5) = "〇" Then       '2024/12/10
'                        strWk2 = strWk2 & "コ"
'                    Else
'                        strWk2 = strWk2 & "個"
'                    End If
                    strJTKB = ""
                End If
            End If
            strLINO = fncGetLINO(fncGetBUTUKCD2(Trim(ST.Cells(lRow, 4))), Trim(ST.Cells(lRow, 5)))
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBWMF.WGIP01 SET"
            strSQL = strSQL & "     GIUPHZ =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,GIUPTM =TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "    ,GIHINM = '" & Trim(ST.Cells(lRow, 1)) & "' "
            strSQL = strSQL & "    ,GIJTI1 = '" & Val(strWk) & "' "
            strSQL = strSQL & "    ,GIJTI2 = '" & strWk2 & "' "
            strSQL = strSQL & "    ,GIJTKB = '" & strJTKB & "'"
            strSQL = strSQL & "    ,GIHHI = " & Val(StrConv(Trim(ST.Cells(lRow, 3)), vbNarrow))
            strSQL = strSQL & "    ,GILINO = '" & strLINO & "'"
            strSQL = strSQL & "    ,GILINM = '" & IIf(strLINO = "", "", Trim(ST.Cells(lRow, 5))) & "'"
            strSQL = strSQL & "    ,GIJAN = " & Val(StrConv(Trim(ST.Cells(lRow, 7)), vbNarrow))
            strSQL = strSQL & " WHERE GIDELT = ''"
            strSQL = strSQL & "   AND GIKHN1 = '" & StrConv(Trim(ST.Cells(lRow, 6)), vbNarrow) & "' "
            strSQL = strSQL & "   AND GIKJCD = '" & fncGetBUTUKCD2(Trim(ST.Cells(lRow, 4))) & "' "
            strSQL = strSQL & "   AND GIKBN = '2' "
            strSQL = strSQL & "   AND GIPROJ = '9' "
            CN.Execute strSQL, lResult, &H80
            If lResult = 0 Then
                strSQL = ""
                strSQL = strSQL & " INSERT INTO LIBWMF.WGIP01 "
                strSQL = strSQL & "         ("
                strSQL = strSQL & "              GIDELT"
                strSQL = strSQL & "             ,GICRHZ"
                strSQL = strSQL & "             ,GICRTM"
                strSQL = strSQL & "             ,GIKJCD"
                strSQL = strSQL & "             ,GIKJNM"
                strSQL = strSQL & "             ,GIHNO"
                strSQL = strSQL & "             ,GIKHN1"
                strSQL = strSQL & "             ,GIHINM"
                For i = 1 To 8
                    strSQL = strSQL & "             ,GIGNM" & i
                Next
                strSQL = strSQL & "             ,GIHYNO"
                For i = 1 To 2
                    strSQL = strSQL & "             ,GIJTI" & i
                Next
                strSQL = strSQL & "             ,GIJTKB"
                strSQL = strSQL & "             ,GIHHI"
                strSQL = strSQL & "             ,GISHI"
                strSQL = strSQL & "             ,GIJAN"
                strSQL = strSQL & "             ,GILINO"
                strSQL = strSQL & "             ,GILINM"
                strSQL = strSQL & "             ,GIPROJ"
                For i = 1 To 20
                    strSQL = strSQL & "             ,GIYB" & Format(i, "00")
                Next
                strSQL = strSQL & "             ,GIKBN"
                strSQL = strSQL & "             ,GIQBL"
                strSQL = strSQL & "         ) VALUES ("
                strSQL = strSQL & "              ''"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'YYYYMMDD')"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'HH24MISS')"
                strSQL = strSQL & "             , '" & fncGetBUTUKCD2(Trim(ST.Cells(lRow, 4))) & "'"
                strSQL = strSQL & "             , '" & Trim(ST.Cells(lRow, 4)) & "'"
                strSQL = strSQL & "             , ''"
                strSQL = strSQL & "             , '" & StrConv(Trim(ST.Cells(lRow, 6)), vbNarrow) & "'"
                strSQL = strSQL & "             , '" & Trim(ST.Cells(lRow, 1)) & "'"
                For i = 1 To 8
                    strSQL = strSQL & "             , ''"
                Next
                strSQL = strSQL & "             , ''"
                strSQL = strSQL & "             , " & Val(strWk)
                strSQL = strSQL & "             , '" & strWk2 & "'"
                strSQL = strSQL & "             , '" & strJTKB & "'"
                strSQL = strSQL & "             , " & Val(StrConv(Trim(ST.Cells(lRow, 3)), vbNarrow))
                strSQL = strSQL & "             , 0"
                strSQL = strSQL & "             , " & Val(StrConv(Trim(ST.Cells(lRow, 7)), vbNarrow))
                strSQL = strSQL & "             , '" & strLINO & "'"
                strSQL = strSQL & "             , '" & IIf(strLINO = "", "", Trim(ST.Cells(lRow, 5))) & "'"
                strSQL = strSQL & "             , '9'"
                For i = 1 To 20
                    strSQL = strSQL & "             , ''"
                Next
                strSQL = strSQL & "             , '2'"
                strSQL = strSQL & "             , ''"
                strSQL = strSQL & "         )"
                
                CN.Execute strSQL, lResult, &H80
                ReDim Preserve addRec(UBound(addRec) + 1)
                addRec(UBound(addRec)).HINM = Trim(ST.Cells(lRow, 1))
                addRec(UBound(addRec)).KCD = fncGetBUTUKCD2(Trim(ST.Cells(lRow, 4)))
            End If
        End If
    Next

    CN.Close: Set CN = Nothing

    Set ST = Nothing
    Set STW = Nothing
    
    Call subAfterEdit
    
End Sub
