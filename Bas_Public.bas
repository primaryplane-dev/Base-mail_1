Option Explicit

'------------------------------------------------------本番----------------------------------------------------------
'Public Const P_ConnectString    As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=ODBC001;Password=FJPN2480;"
'
'Public Const P_SVPathM          As String = "\\srv0901\public1\IJP\"
'Public Const P_SVPathH          As String = "\\srv0601\public1\IJP\"
'Public Const P_SVPathF          As String = "\\srv3701\public1\IJP\"
'Public Const P_SVPathT          As String = "\\srv3601\public\IJP\"
'
'Public Const P_IPAddM           As String = "\\10.2.9.222\,\\10.2.9.224\,\\10.2.9.226\,\\10.2.9.228\"
'Public Const P_IPAddH           As String = "\\10.4.6.155\,\\10.4.6.163\"                                      'CHG 2024.03.31
'Public Const P_IPAddF           As String = "\\10.1.37.219\,\\10.1.37.230\"
'Public Const P_IPAddT           As String = "\\10.1.36.211\,\\10.1.36.212\"
'
'Public Const P_KarutePath       As String = "\\srv2401\public1\添加物関連\フェア商品　　　表示一覧\【最新】冷凍生地表示\"
'------------------------------------------------------テスト--------------------------------------------------------
Public Const P_ConnectString  As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=ODBC001;Password=FJPN2480;"

Public Const P_SVPathM          As String = "\\srv0290\System\システム部\bool\24 BASEインクジェットプリンタ\BASE版\TEST\musashi\"
Public Const P_SVPathH          As String = "\\srv0290\System\システム部\bool\24 BASEインクジェットプリンタ\BASE版\TEST\hirakata\"
Public Const P_SVPathF          As String = "\\srv0290\System\システム部\bool\24 BASEインクジェットプリンタ\BASE版\TEST\fukuoka\"
Public Const P_SVPathT          As String = "\\srv0290\System\システム部\bool\24 BASEインクジェットプリンタ\BASE版\TEST\takara\"

Public Const P_IPAddM           As String = "C:\"
Public Const P_IPAddH           As String = "C:\"
Public Const P_IPAddF           As String = "C:\"
Public Const P_IPAddT           As String = "C:\"

Public Const P_KarutePath       As String = "\\srv0290\System\システム部\bool\24 BASEインクジェットプリンタ\BASE版\カルテ\"
'--------------------------------------------------------------------------------------------------------------------

Public Const P_SavePath         As String = "bin\IJPControl\InjiData\"
Public Const P_CsvName          As String = "ProductData.csv"

'Public P_SVPath()               As String
'Public P_IPAdd()                As String

'frmSelectLINM
Public P_LINM                   As String
Public P_KCD                    As String
Public P_Regist                 As Boolean
'frmSelectKJNM
Public P_KJNM                   As String
Public P_KCD2                   As String
Public P_Regist2                As String

Public Type GIrec
    HNO                         As String
    KHN1                        As String
    HINM                        As String
    GNM(8)                      As String
    HYNO                        As String
    JTI(2)                      As String
    TNI                         As String
    HHI                         As String
    JAN                         As String
    LINO                        As String
    LINM                        As String
    PROJ                        As String
    YOB(20)                     As String
    KCD                         As String
End Type
