Option Explicit
Private Const URI = "http://schemas.microsoft.com/cdo/configuration/"

Public Sub subSendMail(ByVal fromAddress As String, ByVal toAddress As String, ByVal ccAddress As String, ByVal subject As String, ByVal Body As String)
    '送信する
    Dim CDO     As Object
    Dim CONF    As Object
    Dim FLD     As Variant
    
    Set CONF = CreateObject("CDO.Configuration")
    CONF.Load -1    ' CDO Source Defaults
    Set FLD = CONF.Fields
    With FLD
        .Item(URI & "sendusing").Value = 2
        .Item(URI & "smtpserver").Value = "10.2.1.12"
        .Item(URI & "smtpserverport").Value = 25
        .Item(URI & "smtpauthenticate").Value = 1           'cdoBasic
        .Item(URI & "smtpusessl").Value = False
        .Item(URI & "sendusername").Value = fromAddress
'        .Item(URI & "sendpassword").Value = ACCOUNT_PASSWORD
        .Update
    End With

    Set CDO = CreateObject("CDO.Message")
    With CDO
        Set .Configuration = CONF
        .From = fromAddress
        .To = toAddress
        If Not ccAddress = "" Then .cc = ccAddress
        .subject = subject
        .TextBody = Body
        .TextBodyPart.Charset = "shift-jis"
        
        On Error Resume Next
        .Send
        If Err.Number <> 0 Then Debug.Print "(" & Err.Number & ")" & Err.Description
        Err.Clear
        On Error GoTo 0
    End With

    Set CONF = Nothing
    Set CDO = Nothing
End Sub

Public Sub subGetMailAdd(ByVal KCD As String, ByRef fromAddress, ByRef toAddress As String)
    Dim BK      As Workbook
    Dim ST      As Worksheet
    Dim lRow    As Long
    Dim strWk   As String
    
    Set BK = Workbooks.Open(fncGetSVPath(KCD) & "メーリングリスト.xlsx")
    Set ST = BK.Sheets(1)
    
    lRow = 2
    Do While Not ST.Cells(lRow, 1) = ""
        If Not ST.Cells(lRow, 2) = "" Then
            If Not toAddress = "" Then toAddress = toAddress & ","
            toAddress = toAddress & ST.Cells(lRow, 2)
            strWk = StrConv(ST.Cells(lRow, 1), vbNarrow)
            strWk = StrConv(strWk, vbUpperCase)
            If strWk = "BASE" And fromAddress = "" Then
                fromAddress = ST.Cells(lRow, 2)
            End If
        End If
        lRow = lRow + 1
    Loop
    BK.Saved = True
    BK.Close
    Set ST = Nothing
    Set BK = Nothing
    
End Sub

Public Function fncMakebody(ByVal KCD As String, ByVal HINM As String) As String
    fncMakebody = ""
    fncMakebody = KCD & "@" & fncGetKJNM(KCD) & " " & HINM & " IJP印字用データが追加されました"
End Function

Public Function fncMakebody2(ByVal mode As Long, ByVal KCD As String, ByVal HINM As String) As String
    fncMakebody2 = ""
    fncMakebody2 = fncGetKJNM(KCD) & " " & HINM & IIf(mode = 1, " IJP印字用データが削除されました", " IJP印字用データが追加されました")
End Function

