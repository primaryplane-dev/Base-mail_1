Option Explicit

Private Sub cmdCancel_Click()
    ' キャンセル時のリソース解放・変数初期化
    P_KCD2 = ""
    P_KJNM = ""
    P_Regist2 = False
    Unload Me
End Sub

Private Sub lstKJNM_Click()
    Dim kjcd As String, kjnm As String
    kjcd = lstKJNM.List(lstKJNM.ListIndex, 0)
    kjnm = lstKJNM.List(lstKJNM.ListIndex, 1)
    If kjcd = "" Or kjnm = "" Then
        MsgBox "工場コード・工場名を選択してください。", vbExclamation
        Exit Sub
    End If
    If Not kjcd Like "###" And Not kjcd Like "#####" Then
        MsgBox "工場コードは3桁または5桁の数字で入力してください。", vbExclamation
        Exit Sub
    End If
    If Len(kjnm) > 20 Then
        MsgBox "工場名は20文字以内で入力してください。", vbExclamation
        Exit Sub
    End If
    P_KCD2 = kjcd
    P_KJNM = kjnm
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