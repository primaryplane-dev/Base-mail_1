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
    
' ---ＤＢ接続   ←テスト時はコメントにする---
 CN.CursorLocation = adUseClient
 CN.Open P_ConnectString

 ReDim rec(0)
 ReDim arrChangeKJ(0)
 '更新データの取得
 If Not fncGetUpdData(CN, arrChangeKJ, rec) Then GoTo Exit_Update
 If UBound(arrChangeKJ) = 0 Or UBound(rec) = 0 Then GoTo Exit_Update
'-----------------------------------------
    

' ----テスト用ダミーデータを直接セット----
'ReDim arrChangeKJ(1)
'arrChangeKJ(1) = "103"  ' テスト用工場コード
'
'ReDim rec(1)
'rec(1).KCD = "103"
'rec(1).HINM = "テスト商品"
'-----------------------------------------

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
         Call subUpdate(CN, rec(i))     '←テスト時はコメントにする
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
    
    '追加があった工場にメール送信　20260407 修正(送信元選択フォームの表示)
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
            ' 送信元選択フォームを表示
            Dim frm As New frmFromAddress
            frm.KCD = arrChangeKJ(i)   ' 工場コードをセット
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
            Call subSendMail(fromAddress, toAddress, "", "IJPデータ追加通知", strBody)
        End If
    ContinueNextFactory:
    Next

Exit_Update:
    CN.Close: Set CN = Nothing  '←テスト時はコメントにする
End Sub

Private Sub cmdUpdate_Click()
'    Call subSetSVPath
    Call subMain
End Sub
