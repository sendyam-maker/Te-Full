VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030409_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文業績明細查詢"
   ClientHeight    =   5712
   ClientLeft      =   2520
   ClientTop       =   2568
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9324
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG1 
      Height          =   5208
      Left            =   36
      TabIndex        =   1
      Top             =   468
      Width           =   9252
      _ExtentX        =   16341
      _ExtentY        =   9165
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   8040
      TabIndex        =   0
      Top             =   20
      Width           =   1200
   End
   Begin VB.Label lbl 
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frm030409_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/25 Form2.0已修改(MFG1改Fonts)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 0
        'edit by nickc 2005/03/31
        'Me.Hide
        frm030409.Show
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Select Case frm030409.txt1(7).Text
Case "3"
    Me.Caption = "發文業績明細查詢"
Case "4"
    Me.Caption = "發文業績總計查詢"
End Select
'Modify By Sindy 98/03/12 增加依請款日查詢
If Trim(frm030409.txt1(11)) = "1" Then '1.請款
   Me.Lbl(0).Caption = "請款日：" & frm030409.txt1(1).Text & "－" & frm030409.txt1(2).Text
Else
   Me.Lbl(0).Caption = "發文日：" & frm030409.txt1(1).Text & "－" & frm030409.txt1(2).Text
End If
End Sub
    
Sub StrMenu()
Dim ii As Integer
Dim intNowRow As Integer
Dim strSaleZone As String '業務區
Dim intCnt As Integer
Dim dblPoint As Double
Dim intTCnt As Integer
Dim dblTPoint As Double
'Add By Sindy 98/03/12
Dim dblWork As Double
Dim dblTWork As Double
'98/03/12 End

Select Case Val(frm030409.txt1(7).Text)
Case 3 '查詢明細
    'edit by nickc 2007/12/17 改以選擇的功能
    'If Len(frm030409.txt1(6).Text) <> 0 Then
    If frm030409.opt1(1).Value = True Then
        'edit by nickc 2007/12/17 改以選擇的功能
        'strSQL = "SELECT ST02 As 智權人員, A0902 As 業務區, R098003 As 承辦人, R098004 As 發文日, R098005 As 本所案號, R098006 As 案件性質, R098007 As 點數, R098001, R098002 FROM R030409, STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002, R098005 "
        'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
'        strSQL = "SELECT ST02 As 承辦人, A0902 As 業務區, R098003 As 智權人員, R098004 As 發文日, R098005 As 本所案號, R098006 As 案件性質, R098007 As 點數, R098001, R098002 FROM R030409, STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002, R098005 "
        If Trim(frm030409.txt1(11)) = "1" Then '1.請款
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT ST02 As 承辦人, A0902 As 業務區, R098003 As 智權人員, R098004 As 請款日, R098005 As 本所案號, R098006 As 案件性質, R098007||'' As 點數, R098001, R098002, R098008 As 工作時數 FROM R030409, STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002, R098005 "
        Else
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT ST02 As 承辦人, A0902 As 業務區, R098003 As 智權人員, R098004 As 發文日, R098005 As 本所案號, R098006 As 案件性質, R098007||'' As 點數, R098001, R098002, R098008 As 工作時數 FROM R030409, STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002, R098005 "
        End If
    Else
        'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
'        strSQL = "SELECT A0902 As 業務區, ST02 As 智權人員, R098003 As 承辦人, R098004 As 發文日, R098005 As 本所案號, R098006 As 案件性質, R098007 As 點數, R098002, R098001 FROM R030409, STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002 , R098005 "
        If Trim(frm030409.txt1(11)) = "1" Then '1.請款
             'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT A0902 As 業務區, ST02 As 智權人員, R098003 As 承辦人, R098004 As 請款日, R098005 As 本所案號, R098006 As 案件性質, R098007||'' As 點數, R098002, R098001, R098008 As 工作時數 FROM R030409, STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002 , R098005 "
        Else
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT A0902 As 業務區, ST02 As 智權人員, R098003 As 承辦人, R098004 As 發文日, R098005 As 本所案號, R098006 As 案件性質, R098007||'' As 點數, R098002, R098001, R098008 As 工作時數 FROM R030409, STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002 , R098005 "
        End If
    End If
Case 4 '查詢總計
    'edit by nickc 2007/12/17 改以選擇的功能
    'If Len(frm030409.txt1(6).Text) <> 0 Then
    '2009/2/3 MODIFY BY SONIA 陳經理說改以人統計
    'If frm030409.opt1(1).Value = True Then
    '    strSQL = "SELECT '總計' As 智權人員, '' As 業務區, '' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, '', '' FROM R030409 WHERE ID='" & strUserNum & "' "
    'Else
    '    strSQL = "SELECT '總計' As 業務區, '' As 智權人員, '' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, '', '' FROM R030409 WHERE ID='" & strUserNum & "' "
    'End If
    'Modify By Sindy 98/03/12 加工作時數
'    If frm030409.opt1(1).Value = True Then
'        strSQL = "SELECT '小計' As 智權人員, '' As 業務區, ST02 As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, R098001 編號, '' FROM R030409,STAFF WHERE ID='" & strUserNum & "' AND R098001=ST01(+) group by R098001,ST02 "
'        strSQL = strSQL & "UNION SELECT '總計' As 智權人員, '' As 業務區, '' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, '' 編號, '' FROM R030409 WHERE ID='" & strUserNum & "' "
'        strSQL = strSQL & "ORDER BY 編號"
'    Else
'        strSQL = "SELECT A0902 As 業務區, ST02 As 智權人員, '小計' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, R098001||R098002 編號, '' FROM R030409,STAFF,ACC090 WHERE ID='" & strUserNum & "' AND R098002=ST01(+) AND R098001=A0901(+) group by R098001,R098002,ST02,A0902 "
'        strSQL = strSQL & "UNION SELECT '' As 業務區, '' As 智權人員, '總計' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0)) As 點數, '' 編號, '' FROM R030409 WHERE ID='" & strUserNum & "' "
'        strSQL = strSQL & "ORDER BY 編號"
'    End If
    If frm030409.opt1(1).Value = True Then
        'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
        If Trim(frm030409.txt1(11)) = "1" Then '1.請款
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT '小計' As 智權人員, '' As 業務區, ST02 As 承辦人, '' As 請款日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, R098001 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409,STAFF WHERE ID='" & strUserNum & "' AND R098001=ST01(+) group by R098001,ST02 "
            strSql = strSql & "UNION SELECT '總計' As 智權人員, '' As 業務區, '' As 承辦人, '' As 請款日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, '' 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409 WHERE ID='" & strUserNum & "' "
            strSql = strSql & "ORDER BY 編號"
        Else
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT '小計' As 智權人員, '' As 業務區, ST02 As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, R098001 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409,STAFF WHERE ID='" & strUserNum & "' AND R098001=ST01(+) group by R098001,ST02 "
            strSql = strSql & "UNION SELECT '總計' As 智權人員, '' As 業務區, '' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, '' 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409 WHERE ID='" & strUserNum & "' "
            strSql = strSql & "ORDER BY 編號"
        End If
    Else
        'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
        If Trim(frm030409.txt1(11)) = "1" Then '1.請款
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT A0902 As 業務區, ST02 As 智權人員, '小計' As 承辦人, '' As 請款日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, R098001||R098002 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409,STAFF,ACC090 WHERE ID='" & strUserNum & "' AND R098002=ST01(+) AND R098001=A0901(+) group by R098001,R098002,ST02,A0902 "
            strSql = strSql & "UNION SELECT '' As 業務區, '' As 智權人員, '總計' As 承辦人, '' As 請款日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, '' 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409 WHERE ID='" & strUserNum & "' "
            strSql = strSql & "ORDER BY 編號"
        Else
            'Modified by Lydia 2018/03/16 把點數結果改成文字方式=> ||''
            strSql = "SELECT A0902 As 業務區, ST02 As 智權人員, '小計' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, R098001||R098002 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409,STAFF,ACC090 WHERE ID='" & strUserNum & "' AND R098002=ST01(+) AND R098001=A0901(+) group by R098001,R098002,ST02,A0902 "
            strSql = strSql & "UNION SELECT '' As 業務區, '' As 智權人員, '總計' As 承辦人, '' As 發文日, '' As 本所案號, '件數：'||Count(*) As 案件性質, '點數：'||Sum(Nvl(R098007,0))||'' As 點數, '' 編號, '', '工作時數：'||Sum(Nvl(R098008,0)) As 工作時數 FROM R030409 WHERE ID='" & strUserNum & "' "
            strSql = strSql & "ORDER BY 編號"
        End If
    End If
End Select
CheckOC
Screen.MousePointer = vbHourglass
MFG1.MousePointer = flexHourglass
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    Set Me.MFG1.Recordset = adoRecordset
Else
    CheckOC
    ShowNoData
    Me.Enabled = True
    'edit by nickc 2005/03/31
    'Me.Hide
    frm030409.Show
    'add by nickc 2005/03/31
    Unload Me
    MFG1.MousePointer = flexDefault
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Me.MFG1.Visible = False
If Val(frm030409.txt1(7).Text) = 3 Then
    'edit by nickc 2007/12/17 改以選擇的功能
    'If Len(frm030409.txt1(6).Text) <> 0 Then
    If frm030409.opt1(1).Value = True Then
        strSaleZone = Me.MFG1.TextMatrix(1, 0)
    Else
        '2009/2/3 MODIFY BY SONIA 改以業務區+智權人員小計
        'strSaleZone = Me.MFG1.TextMatrix(1, 0)
        strSaleZone = Me.MFG1.TextMatrix(1, 0) & Me.MFG1.TextMatrix(1, 1)
        '2009/2/3 END
    End If
    intNowRow = 1
    intCnt = 0: dblPoint = 0
    dblWork = 0 'Add By Sindy 98/03/12
    intTCnt = 0: dblTPoint = 0
    dblTWork = 0 'Add By Sindy 98/03/12
ReDoFor:
    For ii = intNowRow To Me.MFG1.Rows - 1
         
        'edit by nickc 2007/12/17 改以選擇的功能
        'If Len(frm030409.txt1(6).Text) <> 0 Then
'        If frm030409.opt1(1).Value = True Then
'            If strSaleZone <> Me.MFG1.TextMatrix(ii, 1) Then
'                strSaleZone = Me.MFG1.TextMatrix(ii, 1)
'                Me.MFG1.AddItem "", ii
'                Me.MFG1.TextMatrix(ii, 0) = "小計"
'                Me.MFG1.TextMatrix(ii, 5) = "件數：" & intCnt
'                Me.MFG1.TextMatrix(ii, 6) = "點數：" & dblPoint
'                Me.MFG1.Refresh
'                intNowRow = ii + 1
'                intCnt = 0: dblPoint = 0
'                GoTo ReDoFor
'            Else
'                intCnt = intCnt + 1
'                dblPoint = dblPoint + Val(Me.MFG1.TextMatrix(ii, 6))
'                intTCnt = intTCnt + 1
'                dblTPoint = dblTPoint + Val(Me.MFG1.TextMatrix(ii, 6))
'            End If
'        Else
        If frm030409.opt1(1).Value = True Then
            If strSaleZone <> Me.MFG1.TextMatrix(ii, 0) Then
                strSaleZone = Me.MFG1.TextMatrix(ii, 0)
                Me.MFG1.AddItem "", ii
                Me.MFG1.TextMatrix(ii, 0) = "小計"
                Me.MFG1.TextMatrix(ii, 5) = "件數：" & intCnt
                Me.MFG1.TextMatrix(ii, 6) = "點數：" & dblPoint
                Me.MFG1.TextMatrix(ii, 9) = "工作時數：" & dblWork 'Add By Sindy 98/03/12
                Me.MFG1.Refresh
                intNowRow = ii + 1
                intCnt = 0: dblPoint = 0
                dblWork = 0 'Add By Sindy 98/03/12
                GoTo ReDoFor
            Else
                intCnt = intCnt + 1
                dblPoint = dblPoint + Val(Me.MFG1.TextMatrix(ii, 6))
                dblWork = dblWork + Val(Me.MFG1.TextMatrix(ii, 9)) 'Add By Sindy 98/03/12
                intTCnt = intTCnt + 1
                dblTPoint = dblTPoint + Val(Me.MFG1.TextMatrix(ii, 6))
                dblTWork = dblTWork + Val(Me.MFG1.TextMatrix(ii, 9)) 'Add By Sindy 98/03/12
            End If
        Else
            '2009/2/3 MODIFY BY SONIA 改以業務區+智權人員小計
            'If strSaleZone <> Me.MFG1.TextMatrix(ii, 0) Then
            '    strSaleZone = Me.MFG1.TextMatrix(ii, 0)
            If strSaleZone <> Me.MFG1.TextMatrix(ii, 0) & Me.MFG1.TextMatrix(ii, 1) Then
                strSaleZone = Me.MFG1.TextMatrix(ii, 0) & Me.MFG1.TextMatrix(ii, 1)
            '2009/2/3 END
                Me.MFG1.AddItem "", ii
                Me.MFG1.TextMatrix(ii, 0) = "小計"
                Me.MFG1.TextMatrix(ii, 5) = "件數：" & intCnt
                Me.MFG1.TextMatrix(ii, 6) = "點數：" & dblPoint
                Me.MFG1.TextMatrix(ii, 9) = "工作時數：" & dblWork 'Add By Sindy 98/03/12
                Me.MFG1.Refresh
                intNowRow = ii + 1
                intCnt = 0: dblPoint = 0
                dblWork = 0 'Add By Sindy 98/03/12
                GoTo ReDoFor
            Else
                intCnt = intCnt + 1
                dblPoint = dblPoint + Val(Me.MFG1.TextMatrix(ii, 6))
                dblWork = dblWork + Val(Me.MFG1.TextMatrix(ii, 9)) 'Add By Sindy 98/03/12
                intTCnt = intTCnt + 1
                dblTPoint = dblTPoint + Val(Me.MFG1.TextMatrix(ii, 6))
                dblTWork = dblTWork + Val(Me.MFG1.TextMatrix(ii, 9)) 'Add By Sindy 98/03/12
            End If
        End If
    Next ii
    Me.MFG1.AddItem "", ii
    Me.MFG1.TextMatrix(ii, 0) = "小計"
    Me.MFG1.TextMatrix(ii, 5) = "件數：" & intCnt
    Me.MFG1.TextMatrix(ii, 6) = "點數：" & dblPoint
    Me.MFG1.TextMatrix(ii, 9) = "工作時數：" & dblWork 'Add By Sindy 98/03/12
    ii = ii + 1
    Me.MFG1.AddItem "", ii
    Me.MFG1.TextMatrix(ii, 0) = "總計"
    Me.MFG1.TextMatrix(ii, 5) = "件數：" & intTCnt
    Me.MFG1.TextMatrix(ii, 6) = "點數：" & dblTPoint
    Me.MFG1.TextMatrix(ii, 9) = "工作時數：" & dblTWork 'Add By Sindy 98/03/12
End If
InitGrid
Me.MFG1.Visible = True
Me.Enabled = True
'add by nickc 2005/03/31
MFG1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm030409_1 = Nothing
End Sub

Private Sub InitGrid()
    With Me.MFG1
        .ColWidth(4) = 1600
        .ColWidth(5) = 1600
        .ColWidth(6) = 1600
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 1600 'Add By Sindy 98/03/12
    End With
End Sub
