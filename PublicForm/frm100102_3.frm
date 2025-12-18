VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100102_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "以申請人查詢"
   ClientHeight    =   5715
   ClientLeft      =   90
   ClientTop       =   990
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7572
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   10
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6348
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   4848
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3975
      Left            =   30
      TabIndex        =   11
      Top             =   1725
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   1500
      TabIndex        =   16
      Top             =   1440
      Width           =   7785
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13732;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1500
      TabIndex        =   15
      Top             =   1176
      Width           =   7785
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13732;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   1500
      TabIndex        =   14
      Top             =   915
      Width           =   7785
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13732;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   255
      Left            =   2250
      TabIndex        =   13
      Top             =   390
      Width           =   6915
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "12197;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊ 閉卷；△非申請人案；●銷卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1170
      TabIndex        =   12
      Top             =   180
      Width           =   3360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（日）："
      Height          =   255
      Left            =   45
      TabIndex        =   10
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（英）："
      Height          =   255
      Left            =   45
      TabIndex        =   9
      Top             =   1176
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（中）："
      Height          =   255
      Left            =   45
      TabIndex        =   8
      Top             =   915
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "母案本所案號："
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   652
      Width           =   1260
   End
   Begin VB.Label Label5 
      Caption         =   " "
      Height          =   255
      Left            =   1464
      TabIndex        =   2
      Top             =   652
      Width           =   8268
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   390
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   " "
      Height          =   255
      Left            =   1248
      TabIndex        =   0
      Top             =   390
      Width           =   972
   End
End
Attribute VB_Name = "frm100102_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/16 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Label2、Label7、Label8、Label10
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim strSQL1 As String
Dim StrTag As String, i As Integer, j As Integer, s As Integer, strSql As String
Dim Str01 As String, Str02 As String, strTemp As Variant, StrText4 As String, intK As Integer
Dim Str03 As String, Str04 As String, Str05 As String, Str06 As String, Str07 As String
Dim StrTest4 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "申請國家"
grdDataList.ColWidth(2) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "商品類別"
grdDataList.ColWidth(3) = 4500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "目前准駁"
grdDataList.ColWidth(4) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "排序案號"
grdDataList.ColWidth(5) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
           Dim Str01 As String
           grdDataList.col = 1
           Str01 = SystemNumber(grdDataList, 1)
           If Mid(UCase(Str01), 1, 1) = "N" Then
               Str01 = Mid(Str01, 2, 3)
           End If
           If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Select Case Pub_RplStr(Str01)
                   Case "CFP", "FCP", "P"   '專利
                         Screen.MousePointer = vbHourglass
                         frm100101_3.Show
                         frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_3.StrMenu
                         Screen.MousePointer = vbDefault
                   Case "CFT", "FCT", "T", "TF"   '商標
                         Screen.MousePointer = vbHourglass
                         frm100101_4.Show
                         frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_4.StrMenu
                         Screen.MousePointer = vbDefault
                   'Modify By Sindy 2009/07/24 增加LIN系統類別
                   'modify by sonia 2019/7/29 +ACS系統類別
                   Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
                         Screen.MousePointer = vbHourglass
                         frm100101_5.Show
                         frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_5.StrMenu
                         Screen.MousePointer = vbDefault
                   Case "LA"            '顧問
                         Screen.MousePointer = vbHourglass
                         frm100101_6.Show
                         frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_6.StrMenu
                         Screen.MousePointer = vbDefault
                   Case Else                  '服務
                        Select Case Pub_RplStr(Str01)
                            Case "TB"    '條碼
                                Screen.MousePointer = vbHourglass
                               frm100101_7.Show
                                frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                                frm100101_7.StrMenu
                                Screen.MousePointer = vbDefault
                            Case "TM"
                                Screen.MousePointer = vbHourglass
                               frm100101_8.Show
                                frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                                frm100101_8.StrMenu
                                Screen.MousePointer = vbDefault
                            Case "TD"
                                Screen.MousePointer = vbHourglass
                               frm100101_9.Show
                                frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                                frm100101_9.StrMenu
                                Screen.MousePointer = vbDefault
                            Case "TC", "CFC"
                                Screen.MousePointer = vbHourglass
                               frm100101_A.Show
                                frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                                frm100101_A.StrMenu
                                Screen.MousePointer = vbDefault
                            Case Else
                                Screen.MousePointer = vbHourglass
                               frm100101_B.Show
                                frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                                frm100101_B.StrMenu
                                Screen.MousePointer = vbDefault
                         End Select
               End Select
               Me.Enabled = True
               Exit Sub
               End If
        End If
     Next i
     Me.Enabled = True
Case 1
     Me.Enabled = False
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
        Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            'edit by nickc 2005/05/10
            'strSQL = "SELECT * FROM CASEPROGRESS WHERE CP01='" & SystemNumber(grdDataList.Text, 1) & "' and CP02='" & SystemNumber(grdDataList.Text, 2) & "' AND CP03='" & SystemNumber(grdDataList.Text, 3) & "' AND CP04='" & SystemNumber(grdDataList.Text, 4) & "' "
            'edit by nickc 2007/09/27
            'strSQL = "SELECT * FROM CASEPROGRESS WHERE CP01='" & Replace(SystemNumber(grdDataList.Text, 1), "N", "") & "' and CP02='" & SystemNumber(grdDataList.Text, 2) & "' AND CP03='" & SystemNumber(grdDataList.Text, 3) & "' AND CP04='" & SystemNumber(grdDataList.Text, 4) & "' "
            strSql = "SELECT * FROM CASEPROGRESS WHERE CP01='" & SystemNumber(Pub_RplStr(grdDataList.Text), 1) & "' and CP02='" & SystemNumber(Pub_RplStr(grdDataList.Text), 2) & "' AND CP03='" & SystemNumber(Pub_RplStr(grdDataList.Text), 3) & "' AND CP04='" & SystemNumber(Pub_RplStr(grdDataList.Text), 4) & "' "
            If Len(Trim(Str03)) <> 0 Then
                strSql = strSql + " AND CP05>=" & Val(ChangeTStringToWString(Str03)) & " "
            End If
            If Len(Trim(Str04)) <> 0 Then
                strSql = strSql + " AND CP05<=" & Val(ChangeTStringToWString(Str04)) & " "
            End If
            If Len(Trim(Str05)) <> 0 Then
                strSql = strSql + " AND CP10>='" & Str05 & "' "
            End If
            If Len(Trim(Str06)) <> 0 Then
                strSql = strSql + " AND CP10<='" & Str06 & "' "
            End If
            If Len(Trim(Str07)) <> 0 Then
                If Str07 = "N" Then
                     strSql = strSql + " AND CP09 < 'C'"
                End If
            End If
            'edit by nickc 2007/09/27
            'strSQL = Replace(strSQL, "＊", "")
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                frm100101_2.Show
                frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
                frm100101_2.StrMenu
                frm100101_2.cmdOK(0).Enabled = False
                frm100101_2.cmdOK(1).Enabled = False
                Screen.MousePointer = vbDefault
            Else
                s = MsgBox("此本所案號  " & grdDataList.Text & "找不到或輸入之條件沒有符合的資料", , "警告")
                Me.Enabled = True
                Exit Sub
            End If
            CheckOC
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
Screen.MousePointer = vbHourglass
Me.Enabled = False
Str01 = ""    '申請人編號
Str02 = ""    '本所案號
Str03 = ""    '收文日期(起)
Str04 = ""    '收文日期(迄)
Str05 = ""    '案件性質(起)
Str06 = ""    '案件性質(迄)
Str07 = ""    '是否含來函資料
Str03 = frm100102_1.Text4
Str04 = frm100102_1.Text5
Str05 = frm100102_1.Text6
Str06 = frm100102_1.Text7
Str07 = frm100102_1.Text8
Str01 = Me.Label3.Caption
Str02 = Me.Tag
Label5.Caption = Me.Tag
'組字串
strSQL1 = ""
If Len(Str03) <> 0 Then
    strSQL1 = strSQL1 + " and cp05>=" & Val(ChangeTStringToWString(Str03))
End If
If Len(Str04) <> 0 Then
    strSQL1 = strSQL1 + " and cp05<=" & Val(ChangeTStringToWString(Str04))
End If
If Len(Str05) <> 0 Then
    strSQL1 = strSQL1 + " and cp10>='" & Str04 & "' "
End If
If Len(Str06) <> 0 Then
    strSQL1 = strSQL1 + " and cp10<='" & Str05 & "' "
End If
If UCase(Str07) = "N" Then
    strSQL1 = strSQL1 + " and cp09 < 'C' "
End If

'顯示表單上面的值
'edit by nickc 2006/12/11
'strSQL = "SELECT TM05,TM06,TM07 FROM TRADEMARK WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' AND TM04='" & SystemNumber(Me.Tag, 4) & "' AND TM23='" & Label3.Caption & "' "
strSql = "SELECT TM05,TM06,TM07 FROM TRADEMARK WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' AND TM04='" & SystemNumber(Me.Tag, 4) & "' AND (TM23='" & Label3.Caption & "' or TM78='" & Label3.Caption & "' or TM79='" & Label3.Caption & "' or TM80='" & Label3.Caption & "' or TM81='" & Label3.Caption & "') "
strSql = strSql + "union all select PA05,PA06,PA07 FROM PATENT WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' AND PA04='" & SystemNumber(Me.Tag, 4) & "' AND (PA26='" & Label3.Caption & "' OR PA27='" & Label3.Caption & "' OR PA28='" & Label3.Caption & "' OR PA29='" & Label3.Caption & "' OR PA30='" & Label3.Caption & "') "
'edit by nickc 2006/12/11
'strSQL = strSQL + "union all select SP05,SP06,SP07 FROM SERVICEPRACTICE WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' AND SP04='" & SystemNumber(Me.Tag, 4) & "' AND (SP08='" & Label3.Caption & "' OR SP58='" & Label3.Caption & "' OR SP59='" & Label3.Caption & "') "
strSql = strSql + "union all select SP05,SP06,SP07 FROM SERVICEPRACTICE WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' AND SP04='" & SystemNumber(Me.Tag, 4) & "' AND (SP08='" & Label3.Caption & "' OR SP58='" & Label3.Caption & "' OR SP59='" & Label3.Caption & "' or SP65='" & Label3.Caption & "' or SP66='" & Label3.Caption & "') "
strSql = strSql + "union all select LC05,LC06,LC07 FROM LAWCASE WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' AND LC04='" & SystemNumber(Me.Tag, 4) & "' AND (LC11='" & Label3.Caption & "' OR LC43='" & Label3.Caption & "' OR LC44='" & Label3.Caption & "' OR LC45='" & Label3.Caption & "' OR LC46='" & Label3.Caption & "') "
strSql = strSql + "union all select HC06,'','' FROM HIRECASE WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' AND HC04='" & SystemNumber(Me.Tag, 4) & "' AND (HC05='" & Label3.Caption & "' OR HC24='" & Label3.Caption & "' OR HC25='" & Label3.Caption & "' OR HC26='" & Label3.Caption & "' OR HC27='" & Label3.Caption & "') "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If IsNull(adoRecordset.Fields(0)) Then
        Label2.Caption = ""
    Else
        Label2.Caption = adoRecordset.Fields(0)
    End If
    If IsNull(adoRecordset.Fields(1)) Then
        Label8.Caption = ""
    Else
        Label8.Caption = adoRecordset.Fields(1)
    End If
    If IsNull(adoRecordset.Fields(2)) Then
        Label10.Caption = ""
    Else
        Label10.Caption = adoRecordset.Fields(2)
    End If
End If
CheckOC

'欲搜尋的SQL字串
'Modify By Cheng 2002/04/25
'若已閉卷, 則在本所案號後加"＊"號
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM23='" & Label3.Caption & "' and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1
'strSQL = strSQL + "union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,na03 AS 申請國家,' ' AS 商品類別 FROM PATENT,nation WHERE pa09=na01(+) and PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND (PA26='" & Label3.Caption & "' OR PA27='" & Label3.Caption & "' OR PA28='" & Label3.Caption & "' OR PA29='" & Label3.Caption & "' OR PA30='" & Label3.Caption & "') and pa01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 1) & ") " & strSQL1
'strSQL = strSQL + "union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,na03 AS 申請國家,' ' AS 商品類別 FROM SERVICEPRACTICE,nation WHERE sp09=na01(+) and SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND (SP08='" & Label3.Caption & "' OR SP58='" & Label3.Caption & "' OR SP59='" & Label3.Caption & "') and sp01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 5) & ") " & strSQL1
'strSQL = strSQL + "union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04 AS 本所案號,na03 AS 申請國家,' ' AS 商品類別 FROM LAWCASE,nation WHERE lc15=na01(+) and LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC11='" & Label3.Caption & "' and lc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 3) & ") " & strSQL1
'strSQL = strSQL + "union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04 AS 本所案號,' ' AS 申請國家,' ' AS 商品類別 FROM HIRECASE WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC05='" & Label3.Caption & "' and hc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 4) & ") " & strSQL1 & "  ORDER BY 本所案號"
'Modify By Cheng 2004/04/12
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,TM01||TM02||TM04||TM03 排序案號 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM23='" & Label3.Caption & "' and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1
'edit by nickc 2005/05/10
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,TM01||TM02||TM04||TM03 排序案號 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND Decode(TM01, 'TF', substr(TM02,1,5), TM02)=Decode(TM01, 'TF', '" & Left(SystemNumber(Me.Tag, 2), 5) & "', '" & SystemNumber(Me.Tag, 2) & "') AND TM23='" & Label3.Caption & "' and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1
'edit by nickc 2006/12/11
'strSQL = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,TM01||TM02||TM04||TM03 排序案號 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND Decode(TM01, 'TF', substr(TM02,1,5), TM02)=Decode(TM01, 'TF', '" & Left(SystemNumber(Me.Tag, 2), 5) & "', '" & SystemNumber(Me.Tag, 2) & "') AND TM23='" & Label3.Caption & "' and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1
'edit by nickc 2006/12/25 改不控制申請人，因為有讓與的會不同申請人會查不到 cfp-14589
strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,TM01||TM02||TM04||TM03 排序案號 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND Decode(TM01, 'TF', substr(TM02,1,5), TM02)=Decode(TM01, 'TF', '" & Left(SystemNumber(Me.Tag, 2), 5) & "', '" & SystemNumber(Me.Tag, 2) & "') and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1

'End
'edit by nickc 2005/05/10
'strSQL = strSQL + "union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,PA01||PA02||PA03||PA04 排序案號 FROM PATENT,nation WHERE pa09=na01(+) and PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND (PA26='" & Label3.Caption & "' OR PA27='" & Label3.Caption & "' OR PA28='" & Label3.Caption & "' OR PA29='" & Label3.Caption & "' OR PA30='" & Label3.Caption & "') and pa01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 1) & ") " & strSQL1
'edit by nickc 2006/12/25 改不控制申請人，因為有讓與的會不同申請人會查不到 cfp-14589
'strSQL = strSQL + "union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,PA01||PA02||PA03||PA04 排序案號 FROM PATENT,nation WHERE pa09=na01(+) and PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND (PA26='" & Label3.Caption & "' OR PA27='" & Label3.Caption & "' OR PA28='" & Label3.Caption & "' OR PA29='" & Label3.Caption & "' OR PA30='" & Label3.Caption & "') and pa01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 1) & ") " & strSQL1
strSql = strSql + "union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,PA01||PA02||PA03||PA04 排序案號 FROM PATENT,nation WHERE pa09=na01(+) and PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' and pa01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 1) & ") " & strSQL1

'edit by nickc 2006/12/12
'strSQL = strSQL + "union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,SP01||SP02||SP03||SP04 排序案號 FROM SERVICEPRACTICE,nation WHERE sp09=na01(+) and SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND (SP08='" & Label3.Caption & "' OR SP58='" & Label3.Caption & "' OR SP59='" & Label3.Caption & "') and sp01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 5) & ") " & strSQL1
'edit by nickc 2006/12/25 改不控制申請人，因為有讓與的會不同申請人會查不到 cfp-14589
'strSQL = strSQL + "union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,SP01||SP02||SP03||SP04 排序案號 FROM SERVICEPRACTICE,nation WHERE sp09=na01(+) and SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND (SP08='" & Label3.Caption & "' OR SP58='" & Label3.Caption & "' OR SP59='" & Label3.Caption & "' OR SP65='" & Label3.Caption & "' OR SP66='" & Label3.Caption & "') and sp01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 5) & ") " & strSQL1
'strSQL = strSQL + "union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,LC01||LC02||LC03||LC04 排序案號 FROM LAWCASE,nation WHERE lc15=na01(+) and LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC11='" & Label3.Caption & "' and lc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 3) & ") " & strSQL1
strSql = strSql + "union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,SP01||SP02||SP03||SP04 排序案號 FROM SERVICEPRACTICE,nation WHERE sp09=na01(+) and SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' and sp01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 5) & ") " & strSQL1
strSql = strSql + "union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,na03 AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,LC01||LC02||LC03||LC04 排序案號 FROM LAWCASE,nation WHERE lc15=na01(+) and LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' and lc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 3) & ") " & strSQL1

'92.10.6 MODIFY BY SONIA
'strSQL = strSQL + "union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,' ' AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,HC01||HC02||HC03||HC04 排序案號 FROM HIRECASE WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC05='" & Label3.Caption & "' and hc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 4) & ") " & strSQL1 & "  ORDER BY 排序案號"
'edit by nickc 2006/12/25 改不控制申請人，因為有讓與的會不同申請人會查不到 cfp-14589
'strSQL = strSQL + "union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,' ' AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,HC01||HC02||HC03||HC04 排序案號 FROM HIRECASE WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC05='" & Label3.Caption & "' and hc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 4) & ") " & strSQL1
strSql = strSql + "union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,' ' AS 申請國家,' ' AS 商品類別,'' AS 目前准駁,HC01||HC02||HC03||HC04 排序案號 FROM HIRECASE WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' and hc01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 4) & ") " & strSQL1

'92.10.6 ADD BY SONIA
'Marked By Cheng 2004/04/12
'If SystemNumber(Me.Tag, 1) = "TF" Then
'   strSQL = strSQL + "union all SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,na03 AS 申請國家,TM09 AS 商品類別,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,TM01||TM02||TM04||TM03 排序案號 FROM TRADEMARK,nation WHERE tm10=na01(+) and TM01='" & SystemNumber(Me.Tag, 1) & "' AND SUBSTR(TM02,1,5)='" & Left(SystemNumber(Me.Tag, 2), 5) & "' AND TM23='" & Label3.Caption & "' and tm01 in (" & SQLGrpStr(SystemNumber(Str02, 1), 2) & ") " & strSQL1
'End If
'End
strSql = strSql & "  ORDER BY 排序案號"
'92.10.6 END
CheckOC
If Len(Trim(frm100102_1.Text3)) <> 0 Then
    strTemp = Split(frm100102_1.Text3, ",")
End If
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
     cmdOK(0).Enabled = True
     cmdOK(1).Enabled = True
     cmdOK(2).Enabled = True
Else
    ShowNoData
    Me.Enabled = True
    Screen.MousePointer = vbDefault
        '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)

    Exit Sub
End If
intK = adoRecordset.RecordCount
Set grdDataList.Recordset = adoRecordset
CheckOC
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100102_3 = Nothing
End Sub

Private Sub grdDataList_SelChange()
Screen.MousePointer = vbHourglass
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
grdDataList.Visible = True
Screen.MousePointer = vbDefault
End Sub
