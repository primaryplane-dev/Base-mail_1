Option Explicit


Public KCD As String  ' ← 追加（工場コードを呼び出し元から受け取る）

Private Sub UserForm_Activate()
    Static initialized As Boolean
    If initialized Then Exit Sub
    initialized = True

    Dim BK As Workbook
    Dim ST As Worksheet
    Dim lRow As Long
    Dim topPos As Integer
    Dim opt As MSForms.OptionButton

    Debug.Print "UserForm_Activate KCD=" & KCD

' メーリングリスト.xlsxを開く（工場コードごとにパスを切り替え）
    Set BK = Workbooks.Open(fncGetMailListPathByKCD(Me.KCD))
    Set ST = BK.Sheets(1)
    lRow = 2
    topPos = 10

    Do While ST.Cells(lRow, 1).Value <> ""
        If UCase(StrConv(ST.Cells(lRow, 1).Value, vbNarrow)) = "BASE" Then
            Set opt = Me.fraFromAddress.Controls.Add("Forms.OptionButton.1")
            opt.Caption = ST.Cells(lRow, 2).Value
            opt.Left = 10
            opt.Top = topPos
            opt.Width = 200
            topPos = topPos + 20
        End If
        lRow = lRow + 1
    Loop

    BK.Close SaveChanges:=False
    Set BK = Nothing
    Set ST = Nothing
    ' フリー入力用OptionButtonとTextBoxを追加
    Set opt = Me.fraFromAddress.Controls.Add("Forms.OptionButton.1", "optFreeInput")
    opt.Caption = "フリー入力"
    opt.Left = 10
    opt.Top = topPos
    opt.Width = 80

    Dim txt As MSForms.TextBox
    Set txt = Me.fraFromAddress.Controls.Add("Forms.TextBox.1", "txtFreeInput")
    txt.Left = 100
    txt.Top = topPos - 2
    txt.Width = 180
    txt.Height = 18

End Sub

'OKボタン
Private Sub cmdOK_Click()
    Dim ctrl As Control
    For Each ctrl In Me.fraFromAddress.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Value = True Then
                If ctrl.Name = "optFreeInput" Then
                    ' フリー入力が選択された場合はTextBoxの値をチェック
                    Dim mailInput As String
                    mailInput = Trim(Me.fraFromAddress.Controls("txtFreeInput").Text)
                    If mailInput = "" Then
                        MsgBox "フリー入力欄に送信元アドレスを入力してください。", vbExclamation
                        Exit Sub
                    End If
                    If Len(mailInput) > 100 Then
                        MsgBox "メールアドレスは100文字以内で入力してください。", vbExclamation
                        Exit Sub
                    End If
                    If Not fncIsValidMailAddress(mailInput) Then
                        MsgBox "メールアドレスの形式が正しくありません。", vbExclamation
                        Exit Sub
                    End If
                    Me.Tag = mailInput
                Else
                    Me.Tag = ctrl.Caption
                End If
                Me.Hide
                Exit Sub
            End If
        End If
    Next
    MsgBox "送信元を選択してください。", vbExclamation
End Sub

'キャンセルボタン
Private Sub cmdCancel_Click()
    Unload Me
End Sub