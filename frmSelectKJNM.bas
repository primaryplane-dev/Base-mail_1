Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstKJNM_Click()
    P_KCD2 = lstKJNM.List(lstKJNM.ListIndex, 0)
    P_KJNM = lstKJNM.List(lstKJNM.ListIndex, 1)
    P_Regist2 = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Call subMakeList
    P_Regist2 = False
End Sub

Private Sub subMakeList()
    lstKJNM.Clear
    lstKJNM.AddItem
    lstKJNM.List(lstKJNM.ListCount - 1, 0) = ""
    lstKJNM.List(lstKJNM.ListCount - 1, 1) = ""
    
    lstKJNM.AddItem
    lstKJNM.List(lstKJNM.ListCount - 1, 0) = "001"
    lstKJNM.List(lstKJNM.ListCount - 1, 1) = "武蔵工場"
    
    lstKJNM.AddItem
    lstKJNM.List(lstKJNM.ListCount - 1, 0) = "002"
    lstKJNM.List(lstKJNM.ListCount - 1, 1) = "枚方工場"
    
    lstKJNM.AddItem
    lstKJNM.List(lstKJNM.ListCount - 1, 0) = "103"
    lstKJNM.List(lstKJNM.ListCount - 1, 1) = "福岡工場"
    
    lstKJNM.AddItem
    lstKJNM.List(lstKJNM.ListCount - 1, 0) = "102"
    lstKJNM.List(lstKJNM.ListCount - 1, 1) = "タカラ工場"
End Sub