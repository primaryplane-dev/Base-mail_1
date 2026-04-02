Option Explicit

Private Sub cmdEnd_Click()
    Dim arrKJ()     As Variant: arrKJ = Array("001", "002", "102", "103")
    Dim i           As Long
'    Call subSetSVPath
    For i = 0 To UBound(arrKJ)
        Call subDeleteEditFile(arrKJ(i))
    Next
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    Application.ThisWorkbook.Close False
End Sub

Private Sub cmdNew_Click()
'    Dim blH             As Boolean
'    Dim blM             As Boolean
'    Dim blT             As Boolean
'    Dim blF             As Boolean
    Dim strWk           As String
    Dim lWk             As Long
    Dim lWk2            As Long
    Dim CN              As New ADODB.Connection
    Dim arrChangeKJ()   As String
    Dim rec()           As GIrec
    Dim i               As Long
    Dim j               As Long
    Dim strBody         As String
    Dim fromAddress     As String
    Dim toAddress       As String
    Dim sSp()           As String
    Dim sSp2()          As String
    
    If Dir(ThisWorkbook.Path & "\工場送付用(JAN抽出)*.xlsx") = "" Then MsgBox ("工場送付用(JAN抽出)が見つかりません"): Exit Sub
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    ReDim rec(0)
    ReDim arrChangeKJ(0)
    '更新データの取得
    If Not fncGetUpdData(CN, arrChangeKJ, rec) Then GoTo Exit_Update
    If UBound(arrChangeKJ) = 0 Or UBound(rec) = 0 Then GoTo Exit_Update
    
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
    
    'ASに更新
    strBody = ""
    For i = 1 To UBound(rec)
        If Not strWk = "" Then strWk = strWk & "@@@"
        strWk = strWk & fncMakebody(rec(i).KCD, rec(i).HINM)
        Call subUpdate(CN, rec(i))
    Next
    
    Dim FSO     As New Scripting.FileSystemObject
    For i = 1 To UBound(arrChangeKJ)
        'CSVを作成する
        Call subWriteCSV(arrChangeKJ(i), FSO)
        'CSVをファイルサーバと工場の共有フォルダに送る
        Call subSendCSV(arrChangeKJ(i), FSO)
    Next
    Set FSO = Nothing
    Call subMain
    'エラーになった製品があればシートを表示する
    If Not stError.Cells(3, 1) = "" Then stError.Select
    
    '排他ファイルを削除
    For i = 1 To UBound(arrChangeKJ)
        Call subDeleteEditFile(arrChangeKJ(i))
    Next
    
    '追加があった工場にメール送信
    sSp = Split(strWk, "@@@")
    For i = 1 To UBound(arrChangeKJ)
        strBody = ""
        For j = 0 To UBound(sSp)
            sSp2 = Split(sSp(j), "@")
            If sSp2(0) = arrChangeKJ(i) Then
                If Not strBody = "" Then strBody = strBody & vbLf
                strBody = strBody & sSp2(1)
            End If
        Next
        If Not strBody = "" Then
            fromAddress = "": toAddress = ""
            Call subGetMailAdd(arrChangeKJ(i), fromAddress, toAddress)
            Call subSendMail(fromAddress, toAddress, "", "IJPデータ追加通知", strBody)
        End If
    Next
'Update:
'
'    Call subMakeEditFile
'    If Not fncReplace Then: GoTo Exit_Update
'    Dim FSO     As New Scripting.FileSystemObject
'    'CSVを作成する
'    Call subWriteCSV(FSO)
'    'CSVをファイルサーバと工場の共有フォルダに送る
'    Call subSendCSV(FSO)
'    Set FSO = Nothing
'    Call subMain
'    'エラーになった製品があればシートを表示する
'    If Not stError.Cells(3, 1) = "" Then stError.Select
Exit_Update:
    CN.Close: Set CN = Nothing
End Sub

Private Sub cmdUpdate_Click()
'    Call subSetSVPath
    Call subMain
End Sub
