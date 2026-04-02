Option Explicit

Public Const P_Header = "製品番号,冷凍生地名,原材料名1,原材料名2,原材料名3,原材料名4,原材料名5,原材料名6,原材料名7,原材料名8,内容量1,内容量2,賞味期限,使用期限,JAN,検索用ラインNO.,プロジェクト" & _
                        ",予備1,予備2,予備3,予備4,予備5,予備6,予備7,予備8,予備9,予備10,予備11,予備12,予備13,予備14,予備15,予備16,予備17,予備18,予備19,予備20"

Public Sub subWriteCSV(ByVal i_BUTUKCD As String, ByRef FSO As Scripting.FileSystemObject)
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim TS          As Object
    Dim strWk       As String
    Dim strWk2      As String
    Dim i           As Long
    Dim j           As Long
    Dim strPath     As String
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strPath = fncGetSVPath(i_BUTUKCD)
    Set TS = FSO.OpenTextFile(Filename:=strPath & P_CsvName, IOMode:=2, Create:=True)   'ForWriting
    
    strSQL = ""
    strSQL = strSQL & "SELECT GIHNO "
    strSQL = strSQL & "     , GIKHN1 "
    strSQL = strSQL & "     , GIHINM "
    strSQL = strSQL & "     , GIJTI1 "
    strSQL = strSQL & "     , GIJTI2 "
    strSQL = strSQL & "     , COALESCE(GIJTKB,'') AS GIJTKB "
    strSQL = strSQL & "     , GIHHI "
    strSQL = strSQL & "     , COALESCE(GISHI,0) AS GISHI "
    strSQL = strSQL & "     , GIJAN "
    strSQL = strSQL & "     , GIKJCD "
    strSQL = strSQL & "     , GILINO "
    strSQL = strSQL & "     , GIPROJ "
    For j = 1 To 8
        strSQL = strSQL & " , GIGNM" & j
    Next
    For j = 1 To 20
        strSQL = strSQL & " , GIYB" & Format(j, "00")
    Next
    strSQL = strSQL & " FROM LIBWMF.WGIP01"
    strSQL = strSQL & " WHERE GIDELT <> 'X'"
    strSQL = strSQL & "   AND GIKJCD = '" & i_BUTUKCD & "'"
    strSQL = strSQL & " ORDER BY GIHNO "
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    TS.WriteLine (P_Header)
    Do While Not RS.EOF
        If i_BUTUKCD = "103" Then
            strWk = "": strWk2 = ""
            strWk2 = RS("GIKHN1")
            If Trim(strWk2) = "" Then strWk2 = RS("GIHNO")
            strWk = strWk & strWk2 & ","
            strWk = strWk & RS("GIHINM") & ","
            strWk = strWk & RS("GIGNM1") & ","
            strWk = strWk & RS("GIGNM2") & ","
            strWk = strWk & RS("GIGNM3") & ","
            strWk = strWk & RS("GIGNM4") & ","
            strWk = strWk & RS("GIGNM5") & ","
            strWk = strWk & RS("GIGNM6") & ","
            strWk = strWk & RS("GIGNM7") & ","
            strWk = strWk & RS("GIGNM8") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", "個") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 8, "コ", "個")) & ","                              '2024/12/10
            strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 9, "", IIf(RS("GIPROJ") = 8, "コ", "個"))) & ","    '2024/12/10
            strWk = strWk & RS("GIJTI2") & ","
            strWk = strWk & RS("GIHHI") & ","
            strWk = strWk & RS("GISHI") & ","
            strWk = strWk & Left(RS("GIJAN"), 12) & ","
            strWk = strWk & "925,"
            strWk = strWk & RS("GIPROJ") & ","
            For j = 1 To 19
                strWk = strWk & RS("GIYB" & Format(j, "00")) & ","
            Next
            strWk = strWk & RS("GIYB20")
            TS.WriteLine (strWk)
            
            strWk = "": strWk2 = ""
            strWk2 = RS("GIKHN1")
            If Trim(strWk2) = "" Then strWk2 = RS("GIHNO")
            strWk = strWk & strWk2 & ","
            strWk = strWk & RS("GIHINM") & ","
            strWk = strWk & RS("GIGNM1") & ","
            strWk = strWk & RS("GIGNM2") & ","
            strWk = strWk & RS("GIGNM3") & ","
            strWk = strWk & RS("GIGNM4") & ","
            strWk = strWk & RS("GIGNM5") & ","
            strWk = strWk & RS("GIGNM6") & ","
            strWk = strWk & RS("GIGNM7") & ","
            strWk = strWk & RS("GIGNM8") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", "個") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 8, "コ", "個")) & ","                              '2024/12/10
            strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 9, "", IIf(RS("GIPROJ") = 8, "コ", "個"))) & ","    '2024/12/10
            strWk = strWk & RS("GIJTI2") & ","
            strWk = strWk & RS("GIHHI") & ","
            strWk = strWk & RS("GISHI") & ","
            strWk = strWk & Left(RS("GIJAN"), 12) & ","
            strWk = strWk & "935,"
            strWk = strWk & RS("GIPROJ") & ","
            For j = 1 To 19
                strWk = strWk & RS("GIYB" & Format(j, "00")) & ","
            Next
            strWk = strWk & RS("GIYB20")
            TS.WriteLine (strWk)
        Else
            strWk = "": strWk2 = ""
            strWk2 = RS("GIKHN1")
            If Trim(strWk2) = "" Then strWk2 = RS("GIHNO")
            strWk = strWk & strWk2 & ","
            strWk = strWk & RS("GIHINM") & ","
            strWk = strWk & RS("GIGNM1") & ","
            strWk = strWk & RS("GIGNM2") & ","
            strWk = strWk & RS("GIGNM3") & ","
            strWk = strWk & RS("GIGNM4") & ","
            strWk = strWk & RS("GIGNM5") & ","
            strWk = strWk & RS("GIGNM6") & ","
            strWk = strWk & RS("GIGNM7") & ","
            strWk = strWk & RS("GIGNM8") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", "個") & ","
'                strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 8, "コ", "個")) & ","                              '2024/12/10
            strWk = strWk & RS("GIJTI1") & IIf(RS("GIJTKB") = "1", "kg", IIf(RS("GIPROJ") = 9, "", IIf(RS("GIPROJ") = 8, "コ", "個"))) & ","    '2024/12/10
            strWk = strWk & RS("GIJTI2") & ","
            strWk = strWk & RS("GIHHI") & ","
            strWk = strWk & RS("GISHI") & ","
            strWk = strWk & Left(RS("GIJAN"), 12) & ","
            strWk = strWk & fncGetKLINO(RS("GIKJCD"), RS("GILINO")) & ","
            strWk = strWk & RS("GIPROJ") & ","
            For j = 1 To 19
                strWk = strWk & RS("GIYB" & Format(j, "00")) & ","
            Next
            strWk = strWk & RS("GIYB20")
            TS.WriteLine (strWk)
        End If
     RS.MoveNext
    Loop
    
    TS.Close: Set TS = Nothing
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
End Sub

'制御PCのIPアドレスわかってから実装
Public Sub subSendCSV(ByVal i_BUTUKCD As String, ByRef FSO As Scripting.FileSystemObject)
    Dim strSVPath   As String       '工場ファイルサーバーパス
    Dim strIpAdd    As String       'IJPのPCIPアドレス
    Dim strPath     As String
    Dim sSp()       As String
    Dim sSp2()      As String
    Dim i           As Long
    Dim j           As Long
    Dim ID          As String
    Dim Pass        As String
    
    '-----テストPCへのCSV転送テストを行う場合に使用-------
    '上記テスト時このコメント解除して下のIDPassをコメントアウトする
'    ID = "administrator"
'    Pass = "24802480"
    '-----------------------------------------------------
    
    strSVPath = fncGetSVPath(i_BUTUKCD)
    strIpAdd = fncGetIPAdd(i_BUTUKCD)
    
    ID = "kgk"
    Pass = "kgk"
'    Call subSetSVPath
    
    'IJPのPC台数分送信処理を行う
    sSp = Split(strIpAdd, ",")
    
    'CSVの保存先パスがないPCの場合MkDirするために分割
    sSp2 = Split(P_SavePath & "BK", "\")
    
    For i = 0 To UBound(sSp)
        'テスト時(自分PCにCSV作成)ここはコメントアウトしなくても勝手に分岐される
        If sSp(i) = "C:\" Then
            '保存先パスが存在しなければMkDirする
            For j = 0 To UBound(sSp2)
                If j = 0 Then strPath = sSp(i)
                If j > 0 Then If Not strPath = "" Then strPath = strPath & "\"
                strPath = strPath & sSp2(j)
                If InStr(strPath, "C:\") > 0 Then
                    If Dir(strPath, vbDirectory) = "" Then MkDir strPath
                End If
            Next
            
            strPath = sSp(i) & P_SavePath
            '保存先フォルダにCSVがすでにあれば日付時間をファイル名に追記してBKフォルダに入れる
            If FSO.FileExists(strPath & P_CsvName) Then
                FSO.CopyFile strPath & P_CsvName, strPath & "BK\" & Format(Now, "yyyymmddhhmmss_") & P_CsvName, True
            End If
            If FSO.FileExists(strSVPath & P_CsvName) Then
                FSO.CopyFile strSVPath & P_CsvName, strPath & P_CsvName, True
            End If
            
        '本番時(IJPのPCにCSV送信)ここはコメントアウトしなくても勝手に分岐される
        Else
            Shell ("net use " & Replace(sSp(i), "\c$\", "") & " " & Pass & " /user:" & ID)
            Application.Wait (Now + TimeValue("0:00:03"))
            
            '-----テストPCにCSV転送テストを行う場合に使用-----
            '上記テスト時コメント解除する
'                For j = 0 To UBound(sSp2)
'                    If j = 0 Then strPath = sSp(i)
'                    If j > 0 Then If Not strPath = "" Then strPath = strPath & "\"
'                    strPath = strPath & sSp2(j)
'                    If InStr(strPath, sSp(i)) > 0 Then
'                        If Dir(strPath, vbDirectory) = "" Then MkDir strPath
'                    End If
'                Next
            '-------------------------------------------------
            
            strPath = sSp(i) & P_SavePath
            '保存先フォルダにCSVがすでにあれば日付時間をファイル名に追記してBKフォルダに入れる
            If FSO.FileExists(strPath & P_CsvName) Then
                FSO.CopyFile strPath & P_CsvName, strPath & "BK\" & Format(Now, "yyyymmddhhmmss_") & P_CsvName, True
            End If
            If FSO.FileExists(strSVPath & P_CsvName) Then
                FSO.CopyFile strSVPath & P_CsvName, strPath & P_CsvName, True
            End If
            Shell ("net use " & Replace(sSp(i), "\c$\", "") & " /delete")
        End If
    Next

End Sub

''制御PCのIPアドレスわかってから実装
'Public Sub subSendCSV(ByRef FSO As Scripting.FileSystemObject)
'    Dim strIpAdd    As String
'    Dim strPath     As String
'    Dim sSp()       As String
'    Dim sSp2()      As String
'    Dim i           As Long
'    Dim j           As Long
'    Dim k           As Long
'    Dim ID          As String
'    Dim Pass        As String
'
'    '-----テストPCにCSV転送テストを行う場合に使用-----
'    '上記テスト時このコメント解除して下のIDPassをコメントアウトする
''    ID = "administrator"
''    Pass = "24802480"
'    '-------------------------------------------------
'    ID = "kgk"
'    Pass = "kgk"
'    Call subSetSVPath
'    sSp2 = Split(P_SavePath & "BK", "\")
'
'    For i = 1 To UBound(P_IPAdd)
'        sSp = Split(P_IPAdd(i), ",")
'        For j = 0 To UBound(sSp)
'            If sSp(j) = "C:\" Then
'                For k = 0 To UBound(sSp2)
'                    If k = 0 Then strPath = sSp(j)
'                    If k > 0 Then If Not strPath = "" Then strPath = strPath & "\"
'                    strPath = strPath & sSp2(k)
'                    If InStr(strPath, "C:\") > 0 Then
'                        If Dir(strPath, vbDirectory) = "" Then MkDir strPath
'                    End If
'                Next
'                strPath = sSp(j) & P_SavePath
'                '保存先フォルダにCSVがすでにあれば日付時間をファイル名に追記してBKフォルダに入れる
'                If FSO.FileExists(strPath & P_CsvName) Then
'                    FSO.CopyFile strPath & P_CsvName, strPath & "BK\" & Format(Now, "yyyymmddhhmmss_") & P_CsvName, True
'                End If
'                If FSO.FileExists(P_SVPath(i) & P_CsvName) Then
'                    FSO.CopyFile P_SVPath(i) & P_CsvName, strPath & P_CsvName, True
'                End If
'            Else
'                Shell ("net use " & Replace(sSp(j), "\c$\", "") & " " & Pass & " /user:" & ID)
'                Application.Wait (Now + TimeValue("0:00:03"))
'
'                '-----テストPCにCSV転送テストを行う場合に使用-----
'                '上記テスト時コメント解除する
''                For k = 0 To UBound(sSp2)
''                    If k = 0 Then strPath = sSp(j)
''                    If k > 0 Then If Not strPath = "" Then strPath = strPath & "\"
''                    strPath = strPath & sSp2(k)
''                    If InStr(strPath, sSp(j)) > 0 Then
''                        If Dir(strPath, vbDirectory) = "" Then MkDir strPath
''                    End If
''                Next
'                '-------------------------------------------------
'
'                strPath = sSp(j) & P_SavePath
'                '保存先フォルダにCSVがすでにあれば日付時間をファイル名に追記してBKフォルダに入れる
'                If FSO.FileExists(strPath & P_CsvName) Then
'                    FSO.CopyFile strPath & P_CsvName, strPath & "BK\" & Format(Now, "yyyymmddhhmmss_") & P_CsvName, True
'                End If
'                If FSO.FileExists(P_SVPath(i) & P_CsvName) Then
'                    FSO.CopyFile P_SVPath(i) & P_CsvName, strPath & P_CsvName, True
'                End If
'                Shell ("net use " & Replace(sSp(j), "\c$\", "") & " /delete")
'            End If
'        Next
'    Next
'
'End Sub


'制御PCのIPアドレスわかるまでこっちを使用
'Public Sub subSendCSV(ByRef FSO As Scripting.FileSystemObject)
'    Dim i           As Long
'    Dim j           As Long
'    Dim k           As Long
'    Dim strPath     As String
'    Dim sSp()       As String
'    Dim sSp2()      As String
'    Call subSetSVPath
'    sSp2 = Split(P_SavePath & "BK", "\")
'    For i = 1 To UBound(P_IPAdd)
'        sSp = Split(P_IPAdd(i), ",")
'        For j = 0 To UBound(sSp)
'            For k = 0 To UBound(sSp2)
'                If k = 0 Then strPath = sSp(j)
'                If k > 0 Then If Not strPath = "" Then strPath = strPath & "\"
'                strPath = strPath & sSp2(k)
'                If InStr(strPath, "C:\") > 0 Then
'                    If Dir(strPath, vbDirectory) = "" Then MkDir strPath
'                End If
'            Next
'            strPath = sSp(j) & P_SavePath
'            '保存先フォルダにCSVがすでにあれば日付時間をファイル名に追記してBKフォルダに入れる
'            If FSO.FileExists(strPath & P_CsvName) Then
'                FSO.CopyFile strPath & P_CsvName, strPath & "BK\" & Format(Now, "yyyymmddhhmmss_") & P_CsvName, True
'            End If
'            If FSO.FileExists(P_SVPath(i) & P_CsvName) Then
'                FSO.CopyFile P_SVPath(i) & P_CsvName, strPath & P_CsvName, True
'            End If
'        Next
'    Next
'
'End Sub
