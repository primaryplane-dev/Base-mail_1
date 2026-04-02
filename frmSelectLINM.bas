Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstLINM_Click()
    P_LINM = lstLINM.List(lstLINM.ListIndex, 0)
    P_Regist = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Call subMakeList
    P_Regist = False
End Sub

Private Sub subMakeList()
    lstLINM.Clear
    lstLINM.AddItem
    lstLINM.List(lstLINM.ListCount - 1, 0) = ""
    Select Case P_KCD
    Case "001"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "デバイダー"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "火星人"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "HM"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "メロンビス"
    Case "002"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "デバイダー"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "HM"
    Case "102"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "AD"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "デバイダー（HM）"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "デバイダー（焼成冷凍）"
    Case "103"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "HM"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "FE"
        lstLINM.AddItem
        lstLINM.List(lstLINM.ListCount - 1, 0) = "デバイダー"
    End Select
End Sub
