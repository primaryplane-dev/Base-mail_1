Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim arrKJ()     As Variant: arrKJ = Array("001", "002", "102", "103")
    Dim i           As Long
'    Call subSetSVPath
    For i = 0 To UBound(arrKJ)
        Call subDeleteEditFile(arrKJ(i))
    Next
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    ThisWorkbook.Saved = True
    If Application.Workbooks.Count = 1 Then Application.Quit
End Sub

Private Sub Workbook_Open()
'    Call subSetSVPath
End Sub


