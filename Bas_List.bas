Option Explicit

Public Sub subMain()
    Call subBeforeEdit
    Call subEditList
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    STHina.Cells.Copy stList.Cells
    STHina.Cells.Copy stWork.Cells
    stList.Select
    '書式を登録(excel2010だと17行以降の書式がリセットされるため)
    stList.Range(stList.Cells(3, 1), stList.Cells(stList.Rows.Count, 4)).NumberFormatLocal = "@"
    stList.Range(stList.Cells(3, 1), stList.Cells(stList.Rows.Count, 4)).HorizontalAlignment = xlCenter
    stList.Range(stList.Cells(3, 1), stList.Cells(stList.Rows.Count, 4)).Font.Size = 12
    stWork.Range(stWork.Cells(3, 1), stWork.Cells(stWork.Rows.Count, 4)).NumberFormatLocal = "@"
    stWork.Range(stWork.Cells(3, 1), stWork.Cells(stWork.Rows.Count, 4)).HorizontalAlignment = xlCenter
    stWork.Range(stWork.Cells(3, 1), stWork.Cells(stWork.Rows.Count, 4)).Font.Size = 12
End Sub

Private Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim strPath     As String
    Dim strWk       As String
    
'    Call subSetSVPath
    
    '雛型シート→編集シート
    Call subInitialize
    
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
    lRow = 3
    'stListからstWorkにコピーすると17行問題起こるため、値ひとつひとつ入れてます。
    Do While Not RS.EOF
        ST.Cells(lRow, 1) = RS("GIHNO"): stWork.Cells(lRow, 1) = ST.Cells(lRow, 1)
        ST.Cells(lRow, 2) = fncGetKJNM(RS("GIKJCD")): stWork.Cells(lRow, 2) = ST.Cells(lRow, 2)
        ST.Cells(lRow, 3) = RS("GIHINM"): stWork.Cells(lRow, 3) = ST.Cells(lRow, 3)
        ST.Cells(lRow, 4) = RS("GILINM"): stWork.Cells(lRow, 4) = ST.Cells(lRow, 4)
        lRow = lRow + 1
        RS.MoveNext
    Loop
    ST.Range(ST.Cells(3, 1), ST.Cells(lRow - 1, 5)).Borders.LineStyle = xlContinuous
 
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    ST.Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
End Sub
