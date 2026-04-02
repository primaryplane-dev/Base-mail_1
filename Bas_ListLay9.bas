Option Explicit

Public Sub subMainLay9()
    Call subBeforeEdit
    Call subEditListLay9
    Call subAfterEdit
End Sub

Private Sub subInitializeLay9()
    stHinaLay9.Cells.Copy stLay9.Cells
    stHinaLay9.Cells.Copy stWorkLay9.Cells
    stLay9.Select
    '書式を登録(excel2010だと17行以降の書式がリセットされるため)
    stLay9.Range(stLay9.Cells(3, 1), stLay9.Cells(stLay9.Rows.Count, 7)).NumberFormatLocal = "@"
    stLay9.Range(stLay9.Cells(3, 1), stLay9.Cells(stLay9.Rows.Count, 7)).HorizontalAlignment = xlCenter
    stLay9.Range(stLay9.Cells(3, 1), stLay9.Cells(stLay9.Rows.Count, 7)).Font.Size = 12
    
    stWorkLay9.Range(stWorkLay9.Cells(3, 1), stWorkLay9.Cells(stWorkLay9.Rows.Count, 7)).NumberFormatLocal = "@"
    stWorkLay9.Range(stWorkLay9.Cells(3, 1), stWorkLay9.Cells(stWorkLay9.Rows.Count, 7)).HorizontalAlignment = xlCenter
    stWorkLay9.Range(stWorkLay9.Cells(3, 1), stWorkLay9.Cells(stWorkLay9.Rows.Count, 7)).Font.Size = 12
End Sub

Private Sub subEditListLay9()
    Dim ST          As Worksheet: Set ST = stLay9
    Dim STW         As Worksheet: Set STW = stWorkLay9
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim strPath     As String
    Dim strWk       As String
    
'    Call subSetSVPath
    
    '雛型シート→編集シート
    Call subInitializeLay9
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & "SELECT GIHINM, GIJTI2, GIHHI, GIJTKB, GIKJCD, GILINM, GIKHN1, GIJAN"
    strSQL = strSQL & " FROM LIBWMF.WGIP01"
    strSQL = strSQL & " WHERE GIDELT <> 'X'"
    strSQL = strSQL & "   AND GIKBN = '2' "
    strSQL = strSQL & "   AND GIPROJ = '9' "
    strSQL = strSQL & " ORDER BY GIKHN1 "
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lRow = 3
    'stListからstWorkにコピーすると17行問題起こるため、値ひとつひとつ入れてます。
    Do While Not RS.EOF
        ST.Cells(lRow, 1) = RS("GIHINM"): STW.Cells(lRow, 1) = ST.Cells(lRow, 1)
        ST.Cells(lRow, 2) = RS("GIJTI2"): STW.Cells(lRow, 2) = ST.Cells(lRow, 2)
        ST.Cells(lRow, 3) = RS("GIHHI"): STW.Cells(lRow, 3) = ST.Cells(lRow, 3)
        ST.Cells(lRow, 4) = fncGetKJNM(RS("GIKJCD")): STW.Cells(lRow, 4) = ST.Cells(lRow, 4)
        ST.Cells(lRow, 5) = RS("GILINM"): STW.Cells(lRow, 5) = ST.Cells(lRow, 5)
        ST.Cells(lRow, 6) = RS("GIKHN1"): STW.Cells(lRow, 6) = ST.Cells(lRow, 6)
        ST.Cells(lRow, 7) = RS("GIJAN"): STW.Cells(lRow, 7) = ST.Cells(lRow, 7)
        lRow = lRow + 1
        RS.MoveNext
    Loop
    ST.Range(ST.Cells(3, 1), ST.Cells(lRow + 20 - 1, 7)).Borders.LineStyle = xlContinuous
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    ST.Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
    Set STW = Nothing
End Sub

