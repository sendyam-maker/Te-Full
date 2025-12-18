VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020201 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標處期限通知"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   400
      Left            =   4500
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   60
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail"
      Height          =   400
      Index           =   1
      Left            =   3384
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   60
      Width           =   1080
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "事件說明"
      Height          =   400
      Left            =   2271
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   60
      Width           =   1080
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隱藏白色"
      Enabled         =   0   'False
      Height          =   400
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   1080
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有未發文"
      Height          =   400
      Index           =   1
      Left            =   1158
      TabIndex        =   8
      Top             =   60
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      Height          =   400
      Index           =   3
      Left            =   7855
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   2
      Left            =   6545
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   1260
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm020201.frx":0000
      Left            =   6120
      List            =   "frm020201.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   510
      Width           =   3195
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   5790
      TabIndex        =   3
      Top             =   60
      Width           =   705
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   400
      Left            =   8610
      TabIndex        =   1
      Top             =   60
      Width           =   705
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4875
      Left            =   30
      TabIndex        =   2
      Top             =   840
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8608
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorSel    =   -2147483639
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限|法定期限|承辦期限|承辦人|事件　|本所案號　　　|申請國家　　|案件性質|備註　　　　|案件名稱　　　"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   14
      Top             =   5160
      Width           =   1710
   End
   Begin MSForms.Label lblUserName 
      Height          =   180
      Left            =   2010
      TabIndex        =   11
      Top             =   570
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4830
      TabIndex        =   5
      Top             =   570
      Width           =   1260
   End
End
Attribute VB_Name = "frm020201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 lblUserName/grdDataList
'Create By Sindy 2020/1/13
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim bLvl5 As Boolean
Dim stNumList1(1 To 5) As String
Dim StrToMail(7) As String


Private Function GetValue(pRow As Integer, pCaseNo As String) As String
   Dim ii As Integer
   With Me.grdDataList
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pCaseNo) Then
         GetValue = .TextMatrix(pRow, ii)
         Exit For
      End If
   Next
   End With
End Function

'Add By Sindy 2020/2/11
Private Sub cmdExcel_Click()
Dim strFileName As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim i As Integer, j As Integer
Dim strColVal(1 To 20) As String
   
On Error GoTo flgErr
   
   strFileName = PUB_Getdesktop & "\期限提醒" & strSrvDate(2) & Right("000000" & ServerTime, 6) & ".xls"
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   'xlsSalesPoint.SheetsInNewWorkbook = 4 '3 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   wksaccrpt114.Cells.NumberFormatLocal = "@" '文字
   'wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Range("a1").Value = "本所期限"
   wksaccrpt114.Range("b1").Value = "法定期限"
   wksaccrpt114.Range("c1").Value = "承辦期限"
   wksaccrpt114.Range("d1").Value = "承辦人"
   wksaccrpt114.Range("e1").Value = "智權人員"
   wksaccrpt114.Range("f1").Value = "事件"
   wksaccrpt114.Range("g1").Value = "本所案號"
   wksaccrpt114.Range("h1").Value = "申請國家"
   wksaccrpt114.Range("i1").Value = "案件性質"
   wksaccrpt114.Range("j1").Value = "備註"
   wksaccrpt114.Range("k1").Value = "案件名稱"
   wksaccrpt114.Range("l1").Value = "申請人名稱"
   wksaccrpt114.Range("m1").Value = "註冊號數/申請案號"
   wksaccrpt114.Range("n1").Value = "FC代理人"
   wksaccrpt114.Range("o1").Value = "CF代理人"
   wksaccrpt114.Range("p1").Value = "發文日"
   wksaccrpt114.Range("q1").Value = "收達日"
'   wksaccrpt114.Columns("A:Q").Select
'   wksaccrpt114.Selection.NumberFormatLocal = "@" '文字
   '欄位值
   For i = 1 To grdDataList.Rows - 1
      If grdDataList.RowHeight(i) > 0 Then '有顯示出來的資料列,才要產生在Excel檔案裡
         For j = 1 To 20
            strColVal(j) = ""
         Next
         For j = 1 To 11
            strColVal(j) = grdDataList.TextMatrix(i, j)
         Next j
         strColVal(12) = GetValue(i, "申請人名稱") '申請人名稱
         '註冊號數/申請案號
         If GetValue(i, "TM15") <> "" Then
            strColVal(13) = GetValue(i, "TM15")
         Else
            strColVal(13) = GetValue(i, "TM12")
         End If
         strColVal(14) = GetValue(i, "FC代理人名稱") 'FC代理人
         strColVal(15) = GetValue(i, "CP代理人名稱") 'CF代理人
         strColVal(16) = ChangeWStringToTDateString(GetValue(i, "CP27")) '發文日
         strColVal(17) = ChangeWStringToTDateString(GetValue(i, "CP46")) '收達日
         For j = 1 To 17
            wksaccrpt114.Range(Chr(64 + j) & CStr(i + 1)).Value = strColVal(j)
         Next j
      End If
   Next i
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   wksaccrpt114.Range("A1:Q1").HorizontalAlignment = xlCenter '置中
   xlsSalesPoint.ActiveWindow.FreezePanes = False '取消凍結窗格
   xlsSalesPoint.ActiveWindow.SplitColumn = 0
   xlsSalesPoint.ActiveWindow.SplitRow = 1
'   wksaccrpt114.Range("A1:Q1").Select '凍結窗格-位置
   xlsSalesPoint.ActiveWindow.FreezePanes = True '凍結窗格
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   MsgBox "檔案已產生！電子檔位置：" & strFileName
   
   Exit Sub

flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdHelp_Click()
   frm020201_2.Show vbModal
End Sub

Private Sub cmdHide_Click()
   SetRst2Grid 'Add
   SetColor cmdHide.Tag
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
Dim Str01 As String
Dim StrTmpCp01020304 As String, StrTmpCp09 As String, StrTmpNp22 As String
Dim j As Integer, tmpArr As Variant, TmpArrNp22 As Variant
On Error GoTo ErrorHandler

   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         bolRefresh = False
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next

         StrTag = grdDataList.TextMatrix(i, 7)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If

         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show

         Select Case cmdState
            Case 1 '發E-Mail
               If Len(Trim(StrTag)) <> 0 Then
                  StrToMail(1) = Trim(StrTag) '本所案號
                  If Not IsNull(StrTag) Then
                     Me.Enabled = False
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     StrToMail(2) = Trim(grdDataList.TextMatrix(i, 11)) '案件名稱
                     StrToMail(3) = Trim(grdDataList.TextMatrix(i, 24)) '總收文號
                     StrToMail(4) = Mid(grdDataList.TextMatrix(i, 9), 2, Len(grdDataList.TextMatrix(i, 9))) '案件性質
                     StrToMail(5) = Trim(grdDataList.TextMatrix(i, 1)) '本所期限
                     StrToMail(6) = Trim(grdDataList.TextMatrix(i, 2)) '法定期限
                     StrToMail(7) = Trim(grdDataList.TextMatrix(i, 3)) '承辦期限
                     Screen.MousePointer = vbHourglass
                     '期限
                     If Trim(grdDataList.TextMatrix(i, 12)) = "A" Then
                        frm030301_1.strLimitKind = "本所"
                     ElseIf Trim(grdDataList.TextMatrix(i, 12)) = "B" Then
                        frm030301_1.strLimitKind = "承辦"
                     ElseIf Trim(grdDataList.TextMatrix(i, 12)) = "H" Then
                        frm030301_1.strLimitKind = "法定"
                     ElseIf Trim(grdDataList.TextMatrix(i, 12)) = "D" Then
                        frm030301_1.strLimitKind = "本所"
                        If Left(StrTag, 3) = "CFT" Or Left(StrTag, 3) = "CFC" Or _
                           (Left(StrTag, 1) = "S" And Trim(grdDataList.TextMatrix(i, 25)) <> "000") Then
                           frm030301_1.strLimitKind = "法定"
                        End If
                        '智權人員
                        frm030301_1.StrMailNum2 = Trim(grdDataList.TextMatrix(i, 16))
                        frm030301_1.lbl1(1).Caption = Trim(grdDataList.TextMatrix(i, 5))
                        frm030301_1.strNP22 = Trim(grdDataList.TextMatrix(i, 27)) 'NP22
                     End If
                     frm030301_1.txt1(1) = "                本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
                                                          "                案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf + _
                                                          "                總收文號：" + StrToMail(3) + vbCrLf + vbCrLf + _
                                                          "                案件性質：" + StrToMail(4) + vbCrLf + vbCrLf + _
                                                          "                本所期限：" + StrToMail(5) + vbCrLf + vbCrLf + _
                                                          "                法定期限：" + StrToMail(6) + vbCrLf + vbCrLf + _
                                                          "                承辦期限：" + StrToMail(7)
                     frm030301_1.Show
                     frm030301_1.Tag = StrTag
                     frm030301_1.strCP09 = Trim(StrToMail(3)) '總收文號
                     frm030301_1.strEvents = Trim(grdDataList.TextMatrix(i, 6)) '事件
                     frm030301_1.StrMenu
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If

            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P" '專利
                  Case "CFL", "FCL", "L", "LIN", "ACS" '法務
                  Case "LA" '顧問
                  Case "CFT", "FCT", "T", "TF" '商標
                        Screen.MousePointer = vbHourglass
                        frm100101_4.Show
                        frm100101_4.Tag = StrTag
                        frm100101_4.StrMenu
                        Screen.MousePointer = vbDefault
                  Case Else '服務
                       Select Case Pub_RplStr(Str01)
                           Case "TB"    '條碼
                              Screen.MousePointer = vbHourglass
                              frm100101_7.Show
                              frm100101_7.Tag = StrTag
                              frm100101_7.StrMenu
                              Screen.MousePointer = vbDefault
                           Case "TM"
                              Screen.MousePointer = vbHourglass
                              frm100101_8.Show
                              frm100101_8.Tag = StrTag
                              frm100101_8.StrMenu
                              Screen.MousePointer = vbDefault
                           Case "TD"
                              Screen.MousePointer = vbHourglass
                              frm100101_9.Show
                              frm100101_9.Tag = StrTag
                              frm100101_9.StrMenu
                              Screen.MousePointer = vbDefault
                           Case "TC", "CFC"
                              Screen.MousePointer = vbHourglass
                              frm100101_A.Show
                              frm100101_A.Tag = StrTag
                              frm100101_A.StrMenu
                              Screen.MousePointer = vbDefault
                           Case Else
                              Screen.MousePointer = vbHourglass
                              frm100101_B.Show
                              frm100101_B.Tag = StrTag
                              frm100101_B.StrMenu
                              Screen.MousePointer = vbDefault
                        End Select
               End Select

            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         Exit For
      End If
   Next i

   If bolRefresh = True Then
      cmdQuery_Click 0
   End If

ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmdQuery_Click(Index As Integer)

   Screen.MousePointer = vbHourglass
   If pub_bolInformCheck = False Then
      If MsgBox("是否確定要查詢？", vbYesNo + vbDefaultButton2) = vbNo Then
         GoTo SubOut
      End If
   End If

   Me.Enabled = False
   doQuery Index
   Me.Enabled = True

SubOut:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String

   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery(idx As Integer)
Dim stVTB As String
Dim stDate0 As String, stDate1 As String, stDate2 As String, stDate_3 As String
Dim stDate4 As String
Dim stDate5 As String, stDate6 As String, stDate7 As String, stDate_7 As String
Dim stNumList As String, stDept As String, stDeptST03 As String
Dim stNumList_2 As String '離職人員
Dim stNumList_all As String
Dim ii As Integer, stIdList
Dim stUserID As String
Dim strOtherUser As String
Dim txtData As Variant, strWhSql As String, strUser As String
Dim iRow As Long
Dim rsTmp As New ADODB.Recordset
Dim strSysCode As String

   stVTB = ""
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass

   stUserID = txtUsernum
   '使用者收文智權人員所屬部門
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
      'stDeptST03 = Pub_StrUserSt03
   Else
      stDept = GetST15(stUserID)
      'stDeptST03 = GetStaffDepartment(stUserID)
   End If

   '抓員工外譯對照資料
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   '期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next

   stDate0 = strSrvDate(1) - 10000 '系統日-1年
   stDate1 = CompWorkDay(2, strSrvDate(1)) '減1個工作天(不含當天)
   stDate2 = CompWorkDay(3, strSrvDate(1)) '減2個工作天(不含當天)
   stDate7 = CompWorkDay(8, strSrvDate(1)) '減7個工作天(不含當天)
   stDate_7 = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -7) + 19110000
   strSysCode = "'FCT','T','TB','TC','TD','TF','TM','TR','TS','TT'"
   
   '特殊權限
   bLvl5 = CheckLevel(stUserID, "V2") '商標處第五級管制人
   'Modify By Sindy 2020/2/11 V2:加所有程序組及承辦組(P2X)之管制人(含離職人員的期限資料且不管是不是過期)
   If bLvl5 = True Then
      '在職人員
      strSql = "SELECT st01,st02 FROM staff WHERE substr(st03,1,2)='P2' and st04='1' and st01<>'" & stUserID & "'"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If stNumList <> "" Then
               stNumList = stNumList & ",'" & rsTmp.Fields("st01") & "'"
            Else
               stNumList = "'" & rsTmp.Fields("st01") & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      '離職人員
      stNumList_2 = ""
      strSql = "SELECT st01,st02 FROM staff WHERE substr(st03,1,2)='P2' and st04='2'"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If stNumList_2 <> "" Then
               stNumList_2 = stNumList_2 & ",'" & rsTmp.Fields("st01") & "'"
            Else
               stNumList_2 = "'" & rsTmp.Fields("st01") & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      stNumList_all = stNumList & "," & stNumList_2 '全部人員(離職人員)
   Else
      stNumList_all = stNumList
   End If
   '2020/2/11 END
   
   '清除暫存檔
   strSql = "delete R030301 where ID='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   '代碼1:A=達本所,B=達承辦,C=待提申,D=未收文,E=未發文,F=待完成,G=待收達,H=達法定
   '      I=達指會,J=可送件
   '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   '承辦人為程序人員(部門P22)時
   If Trim(stDept) = "P22" Or bLvl5 = True Then
      '已收文未發文案件之期限提醒
      '2個工作天後達本所期限者(不含當日)未發文的案件。
      '事件顯示為A達本所。
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate2 & _
               " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND cp14=st01(+) and st03='P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate2 & _
               " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND cp14=st01(+) and st03='P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND cp14=st01(+) and st03='P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND cp14=st01(+) and st03='P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
      
      '下一程序未完成之期限提醒
      '針對下一程序且管制人員NP10為程序人員(部門P22)時，
      '2個工作天後達本所期限者(不含當日)未完成的案件。
      '事件顯示為F待完成。
      'AND NP08>=" & stDate0 & " AND NP08<=" & stDate2
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,TRADEMARK,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08>0 AND NP08<=" & stDate2 & _
               " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND NP06 is null" & _
               " AND NP10 IN(" & stNumList & ") AND NP10=st01(+) and st03='P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      'AND NP08>=" & stDate0 & " AND NP08<=" & stDate2
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08>0 AND NP08<=" & stDate2 & _
               " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND NP06 is null" & _
               " AND NP10 IN(" & stNumList & ") AND NP10=st01(+) and st03='P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND NP06 is null" & _
                  " AND NP10 IN(" & stNumList_2 & ") AND NP10=st01(+) and st03='P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND NP06 is null" & _
                  " AND NP10 IN(" & stNumList_2 & ") AND NP10=st01(+) and st03='P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
   End If
   
   '承辦組
   If Trim(stDept) = "P20" Or Trim(stDept) = "P21" Or bLvl5 = True Then
      '已收文未發文案件之期限提醒
      '7個工作天後達本所期限者(不含當日)未發文的案件。事件顯示為達本所。
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate7 & _
               " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate7 & _
               " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
      
      '已收文未發文案件之期限提醒
      '2個工作天後達承辦期限者(不含當日)未發文的案件。事件顯示為達承辦。
      'AND CP48>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP48>0 AND CP48<=" & stDate2 & _
               " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      'AND CP48>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP48>0 AND CP48<=" & stDate2 & _
               " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
      
      '已收文未發文案件之延展可辦提醒
      '延展達可辦期限前一個工作天未發文的案件。事件顯示為J可送件。
      '台灣案可辦期限為法定期限前六個月、大陸案為法定期限前一年。(大陸案自2014/05/01起改為期滿前一年即可辦理)
      ' AND cp07 > 20000000
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22,CP79)" & _
               " SELECT '" & strUserNum & "','J' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),nvl(CP14,'69008'),CP13,0,DECODE(CP141,'2',NVL(CP79,0),0) typ2" & _
               " FROM CaseProgress,TradeMark,STAFF" & _
               " WHERE substr(cp01,1,1)='T' and cp01<>'TF' and CP10='102'" & _
               " and CP158=0 AND CP159=0" & _
               " AND TM29||TM57 is null" & _
               " AND workdayadd(-2," & PUB_Get102DeadLine("1", "CP07") & ") < " & stDate1 & _
               " AND cp05<=workdayadd(-2," & PUB_Get102DeadLine("1", "CP07") & ")" & _
               " AND cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
               " AND tm10<>'020' AND tm10 is not null" & _
               " AND nvl(CP14,'69008') IN(" & stNumList_all & ") AND CP14=st01(+) and st03<>'P22'" & _
               " union SELECT '" & strUserNum & "','J' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),nvl(CP14,'69008'),CP13,0,DECODE(CP141,'2',NVL(CP79,0),0) typ2" & _
               " FROM CaseProgress,TradeMark,STAFF" & _
               " WHERE substr(cp01,1,1)='T' and cp01<>'TF' and CP10='102'" & _
               " and CP158=0 AND CP159=0" & _
               " AND TM29||TM57 is null" & _
               " AND workdayadd(-2," & PUB_Get102DeadLine("2", "CP07") & ") < " & stDate1 & _
               " AND cp05<=workdayadd(-2," & PUB_Get102DeadLine("2", "CP07") & ")" & _
               " AND cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
               " AND tm10='020'" & _
               " AND nvl(CP14,'69008') IN(" & stNumList_all & ") AND CP14=st01(+) and st03<>'P22'"
      cnnConnection.Execute strSql, intI
      '馬德里是3個月
      ' AND cp07 > 20000000
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','J' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),nvl(CP14,'69008'),CP13,0" & _
               " FROM CaseProgress,TradeMark,STAFF" & _
               " WHERE cp01='TF' and CP10='102'" & _
               " and CP158=0 AND CP159=0" & _
               " AND TM29||TM57 is null" & _
               " AND TO_CHAR(ADD_MONTHS(TO_date(cp07,'YYYYMMDD'),-3),'YYYYMMDD') < " & stDate1 & _
               " AND cp05<=TO_CHAR(ADD_MONTHS(TO_date(cp07,'YYYYMMDD'),-3),'YYYYMMDD')" & _
               " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null" & _
               " AND nvl(CP14,'69008') IN(" & stNumList_all & ") AND CP14=st01(+) and st03<>'P22'"
      cnnConnection.Execute strSql, intI
      
      '已發文未提申案件之期限提醒
      '非台灣案已有代理人收達日逾七個工作天仍未提申的案件。事件顯示為C待提申。
      '1.下一程序有提申期限：以提申期限檢查條件
      'AND np08>=20140101
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','C' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,CP14,cp83,NP22" & _
               " FROM caseprogress,trademark,STAFF,nextprogress" & _
               " WHERE np02 IN (" & strSysCode & ") and np07='998' and np06 is null and np01=cp09" & _
               " AND CP159=0 and CP158>0" & _
               " AND np08>0 and np08<=" & stDate_7 & _
               " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null" & _
               " AND TM29||TM57 is null" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','C' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,CP14,cp83,NP22" & _
               " FROM caseprogress,SERVICEPRACTICE,STAFF,nextprogress" & _
               " WHERE np02 IN (" & strSysCode & ") and np07='998' and np06 is null and np01=cp09" & _
               " AND CP159=0 and CP158>0" & _
               " AND np08>0 and np08<=" & stDate_7 & _
               " AND cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null" & _
               " AND SP15||SP61 is null" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','C' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,CP14,cp83,NP22" & _
                  " FROM caseprogress,trademark,STAFF,nextprogress" & _
                  " WHERE np02 IN (" & strSysCode & ") and np07='998' and np06 is null and np01=cp09" & _
                  " AND CP159=0 and CP158>0" & _
                  " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null" & _
                  " AND TM29||TM57 is null" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','C' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,CP14,cp83,NP22" & _
                  " FROM caseprogress,SERVICEPRACTICE,STAFF,nextprogress" & _
                  " WHERE np02 IN (" & strSysCode & ") and np07='998' and np06 is null and np01=cp09" & _
                  " AND CP159=0 and CP158>0" & _
                  " AND cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null" & _
                  " AND SP15||SP61 is null" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
      '2.下一程序無提申期限(不管提申期限是否已上續辦)：以代理人收達日+七個工作天檢查資料，
      '  代理人收達日+七個工作天顯示在本所期限及法定期限欄。
      'AND cp46>=20140101
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','C' EV1,'1' EV2,CP09,nvl(WORKDAYADD(-7,cp46),0),nvl(WORKDAYADD(-7,cp46),0),0,CP14,cp83,0" & _
               " FROM caseprogress,trademark,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP159=0 and CP158>0" & _
               " AND cp46>0 and cp46<=" & stDate_7 & _
               " AND nvl(cp47,0)=0" & _
               " AND cp24 is null" & _
               " AND cp10 not in ('001','108','201','203','206','302','612','706','711','714','308','701')" & _
               " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null" & _
               " AND TM29||TM57 is null" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND exists (select * from nextprogress where np01=cp09 and np07='998' and np06 is null)"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','C' EV1,'1' EV2,CP09,nvl(WORKDAYADD(-7,cp46),0),nvl(WORKDAYADD(-7,cp46),0),0,CP14,cp83,0" & _
               " FROM caseprogress,SERVICEPRACTICE,STAFF" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP159=0 and CP158>0" & _
               " AND cp46>0 and cp46<=" & stDate_7 & _
               " AND nvl(cp47,0)=0" & _
               " AND cp24 is null" & _
               " AND cp10 not in ('001','108','201','203','206','302','612','706','711','714','308','701')" & _
               " AND cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null" & _
               " AND SP15||SP61 is null" & _
               " AND CP14 IN(" & stNumList & ") AND CP14=st01(+) and st03<>'P22'" & _
               " AND exists (select * from nextprogress where np01=cp09 and np07='998' and np06 is null)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','C' EV1,'1' EV2,CP09,nvl(WORKDAYADD(-7,cp46),0),nvl(WORKDAYADD(-7,cp46),0),0,CP14,cp83,0" & _
                  " FROM caseprogress,trademark,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP159=0 and CP158>0" & _
                  " AND cp46>0" & _
                  " AND nvl(cp47,0)=0" & _
                  " AND cp24 is null" & _
                  " AND cp10 not in ('001','108','201','203','206','302','612','706','711','714','308','701')" & _
                  " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null" & _
                  " AND TM29||TM57 is null" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND exists (select * from nextprogress where np01=cp09 and np07='998' and np06 is null)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','C' EV1,'1' EV2,CP09,nvl(WORKDAYADD(-7,cp46),0),nvl(WORKDAYADD(-7,cp46),0),0,CP14,cp83,0" & _
                  " FROM caseprogress,SERVICEPRACTICE,STAFF" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP159=0 and CP158>0" & _
                  " AND cp46>0" & _
                  " AND nvl(cp47,0)=0" & _
                  " AND cp24 is null" & _
                  " AND cp10 not in ('001','108','201','203','206','302','612','706','711','714','308','701')" & _
                  " AND cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null" & _
                  " AND SP15||SP61 is null" & _
                  " AND CP14 IN(" & stNumList_2 & ") AND CP14=st01(+) and st03<>'P22'" & _
                  " AND exists (select * from nextprogress where np01=cp09 and np07='998' and np06 is null)"
         cnnConnection.Execute strSql, intI
      End If
      '已發文未收達之期限提醒
      '下一程序收達期限管制人NP10為程序(部門P22)時，針對"達"收達期限的案件，
      '事件顯示為G待收達。
      '本所期限及法定期限顯示下一程序資料。
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','G' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,TRADEMARK,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND NP06 is null AND NP07='997'" & _
               " AND NP10=st01(+) and st03='P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='G' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','G' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND NP06 is null AND NP07='997'" & _
               " AND NP10=st01(+) and st03='P22'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='G' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','G' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND NP06 is null AND NP07='997'" & _
                  " AND NP10=st01(+) and st03='P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='G' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','G' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND NP06 is null AND NP07='997'" & _
                  " AND NP10=st01(+) and st03='P22'" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='G' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
      '依發文進度之承辦人CP14顯示案件
      strSql = "UPDATE R030301 Rt Set Rt.CP14=(select c1.cp14 from caseprogress c1 where c1.cp09=Rt.cp09) where Rt.ID='" & strUserNum & "' and Rt.EV1='G'"
      cnnConnection.Execute strSql, intI
      '依承辦人CP14過濾資料
      strSql = "Delete R030301 where ID='" & strUserNum & "' and EV1='G' AND CP14 not IN(" & stNumList_all & ")"
      cnnConnection.Execute strSql, intI
      
      '下一程序未完成之期限提醒
      '下一程序非催審305期限且管制人員非程序人員(部門P22)時，
      '2個工作天後達本所期限者(不含當日)未完成的案件顯示案件。
      '事件顯示為F待完成。
      'AND NP08>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,TRADEMARK,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08>0 AND NP08<=" & stDate2 & _
               " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND NP06 is null AND NP07<>'305'" & _
               " AND NP10 IN(" & stNumList & ") AND NP10=st01(+) and st03<>'P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      'AND NP08>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
               " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
               " WHERE NP02 IN (" & strSysCode & ")" & _
               " AND NP08>0 AND NP08<=" & stDate2 & _
               " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND NP06 is null AND NP07<>'305'" & _
               " AND NP10 IN(" & stNumList & ") AND NP10=st01(+) and st03<>'P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,TRADEMARK,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = TM01(+) AND NP03 = TM02(+) AND NP04 = TM03(+) AND NP05 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND NP06 is null AND NP07<>'305'" & _
                  " AND NP10 IN(" & stNumList_2 & ") AND NP10=st01(+) and st03<>'P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
         'AND NP08>=" & stDate0 & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'1' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0,'',NP10,NP22" & _
                  " FROM NextPROGRESS,SERVICEPRACTICE,STAFF" & _
                  " WHERE NP02 IN (" & strSysCode & ")" & _
                  " AND NP02 = SP01(+) AND NP03 = SP02(+) AND NP04 = SP03(+) AND NP05 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND NP06 is null AND NP07<>'305'" & _
                  " AND NP10 IN(" & stNumList_2 & ") AND NP10=st01(+) and st03<>'P22' AND not(" & Mid(Trim(strNpSqlOfNoSalesDuty), 4) & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='1' and R030301.CP09=NP01 and R030301.np22=NP22)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
   End If
   
'***********************
''A' EV1,'0' EV2
'***********************
   '未分案-0
   If bLvl5 = True Then
      '已收文未發,7個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,TRADEMARK,STAFF_GROUP" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate7 & _
               " AND 'C1' = SG01(+) AND CP01 = SG02(+) AND CP10 = SG03(+)" & _
               " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
               " AND CP09 < 'C'  AND (CP14 IS NULL OR CP14 = '' )" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      'AND CP06>=" & stDate0 & "
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF_GROUP" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP06>0 AND CP06<=" & stDate7 & _
               " AND 'C1' = SG01(+) AND CP01 = SG02(+) AND CP10 = SG03(+)" & _
               " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
               " AND CP09 < 'C'  AND (CP14 IS NULL OR CP14 = '' )" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      
'***********************
''E' EV1,'0' EV2
'***********************
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         'AND CP05>20030000
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,TRADEMARK,STAFF_GROUP" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND 'C1' = SG01(+) AND CP01 = SG02(+) AND CP10 = SG03(+)" & _
                  " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
                  " AND CP09 < 'C'  AND (CP14 IS NULL OR CP14 = '' )" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,STAFF_GROUP" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND 'C1' = SG01(+) AND CP01 = SG02(+) AND CP10 = SG03(+)" & _
                  " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
                  " AND CP09 < 'C'  AND (CP14 IS NULL OR CP14 = '' )" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
'***********************
''E' EV1,'1' EV2
'***********************
   '所有未發文--承辦人-E(未發文)
   If idx = 1 Then
      'AND CP05>20030000
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,TRADEMARK" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
               " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
               " FROM CASEPROGRESS,SERVICEPRACTICE" & _
               " WHERE CP01 IN (" & strSysCode & ")" & _
               " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
               " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
               " AND CP14 IN(" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
      cnnConnection.Execute strSql, intI
      '***** 離職人員 *****
      If bLvl5 = True Then
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,TRADEMARK" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = TM01(+) AND CP02 = TM02(+) AND CP03 = TM03(+) AND CP04 = TM04(+) AND TM01 is not null" & _
                  " AND TM29||TM57 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),CP14,CP13,0" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE" & _
                  " WHERE CP01 IN (" & strSysCode & ")" & _
                  " AND CP01 = SP01(+) AND CP02 = SP02(+) AND CP03 = SP03(+) AND CP04 = SP04(+) AND SP01 is not null" & _
                  " AND SP15||SP61 is null AND CP158=0 AND CP159=0" & _
                  " AND CP14 IN(" & stNumList_2 & ")" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CP09 and R030301.np22=0)"
         cnnConnection.Execute strSql, intI
      End If
      '******************** END
   End If
   
   'F待完成,依發文進度之承辦人CP14顯示案件
   strSql = "UPDATE R030301 Rt Set Rt.CP14=(select c1.cp14 from caseprogress c1 where c1.cp09=Rt.cp09) where Rt.ID='" & strUserNum & "' and Rt.EV1='F'"
   cnnConnection.Execute strSql, intI
   'PS：承辦組上述期限若已達本所(a)時，則達承辦(b)或可送件(c)就不會重覆顯示。
   '(1)若同案有 'A達本所'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'A'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='A' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '(2)若同案有 'B達承辦'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'B'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='B' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '(3)若同案有 'G待收達'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'G'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='G' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   
'*****************
' 案件進度
'*****************
   '事件顯示為達承辦:但已完稿案件在承辦期限日期欄右邊加「完」字樣。
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   strExc(0) = "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),9,' '),'2')||decode(EV1,'B',decode(nvl(ep09,0),0,'','完'),'') 承辦期限," & _
      "S2.ST02 承辦人,S3.ST02 智權人員," & _
      "decode(EV2,'0','未分案',DECODE(EV1,'A','達本所','B','達承辦','C','待提申','D','未收文','E','未發文','F','待完成','G','待收達','H','達法定','I','達指會','J','可送件')) 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號,NA03 申請國家," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質,CP64 備註,TM05 案件名稱," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,CP10,CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04,R030301.CP79 typ2,CP44,CP46,TM23,TM15,TM12,TM44,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) 申請人名稱,NVL(F1.FA04,Decode(F1.FA05,null,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) FC代理人名稱,NVL(F2.FA04,Decode(F2.FA05,null,F2.FA06,F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)) CP代理人名稱" & _
      " FROM R030301,trademark,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm,nation,engineerprogress,customer,fagent f1,fagent f2" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16 AND S2.ST01(+)=R030301.CP14 AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND R030301.CP09=ep02(+) AND R030301.CP09=ti02(+) AND tm10=na01(+)" & _
      " AND substr(tm23,1,8)=cu01(+) AND substr(tm23,9,1)=cu02(+)" & _
      " AND substr(tm44,1,8)=f1.fa01(+) AND substr(tm44,9,1)=f1.fa02(+) AND substr(cp44,1,8)=f2.fa01(+) AND substr(cp44,9,1)=f2.fa02(+)"
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),9,' '),'2')||decode(EV1,'B',decode(nvl(ep09,0),0,'','完'),'') 承辦期限," & _
      "S2.ST02 承辦人,S3.ST02 智權人員," & _
      "decode(EV2,'0','未分案',DECODE(EV1,'A','達本所','B','達承辦','C','待提申','D','未收文','E','未發文','F','待完成','G','待收達','H','達法定','I','達指會','J','可送件')) 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號,NA03 申請國家," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,CP10,CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04,R030301.CP79 typ2,CP44,CP46,SP08 TM23,SP14 TM15,SP11 TM12,SP26 TM44,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) 申請人名稱,NVL(F1.FA04,Decode(F1.FA05,null,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) FC代理人名稱,NVL(F2.FA04,Decode(F2.FA05,null,F2.FA06,F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)) CP代理人名稱" & _
      " FROM R030301,SERVICEPRACTICE,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm,nation,engineerprogress,customer,fagent f1,fagent f2" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16 AND S2.ST01(+)=R030301.CP14 AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND R030301.CP09=ep02(+) AND R030301.CP09=ti02(+) AND sp09=na01(+)" & _
      " AND substr(sp08,1,8)=cu01(+) AND substr(sp08,9,1)=cu02(+)" & _
      " AND substr(sp26,1,8)=f1.fa01(+) AND substr(sp26,9,1)=f1.fa02(+) AND substr(cp44,1,8)=f2.fa01(+) AND substr(cp44,9,1)=f2.fa02(+)"
'*****************
' 下一程序
'*****************
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),9,' '),'2') 承辦期限," & _
      "S2.ST02 承辦人,S3.ST02 智權人員," & _
      "decode(EV2,'0','未分案',DECODE(EV1,'A','達本所','B','達承辦','C','待提申','D','未收文','E','未發文','F','待完成','G','待收達','H','達法定','I','達指會','J','可送件')) 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號,NA03 申請國家," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質,NP15 備註,TM05 案件名稱," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,NP07,CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04,R030301.CP79 typ2,CP44,CP46,TM23,TM15,TM12,TM44,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) 申請人名稱,NVL(F1.FA04,Decode(F1.FA05,null,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) FC代理人名稱,NVL(F2.FA04,Decode(F2.FA05,null,F2.FA06,F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)) CP代理人名稱" & _
      " FROM R030301,trademark,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm,nation,caseprogress,customer,fagent f1,fagent f2" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0 AND caseprogress.CP09(+)=R030301.CP09" & _
      " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16 AND S2.ST01(+)=R030301.CP14 AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) AND tm10=na01(+) AND substr(tm23,1,8)=cu01(+) AND substr(tm23,9,1)=cu02(+)" & _
      " AND substr(tm44,1,8)=f1.fa01(+) AND substr(tm44,9,1)=f1.fa02(+)" & _
      " AND substr(cp44,1,8)=f2.fa01(+) AND substr(cp44,9,1)=f2.fa02(+)"
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),9,' '),'2') 承辦期限," & _
      "S2.ST02 承辦人,S3.ST02 智權人員," & _
      "decode(EV2,'0','未分案',DECODE(EV1,'A','達本所','B','達承辦','C','待提申','D','未收文','E','未發文','F','待完成','G','待收達','H','達法定','I','達指會','J','可送件')) 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號,NA03 申請國家," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質,NP15 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,NP07,CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04,R030301.CP79 typ2,CP44,CP46,SP08 TM23,SP14 TM15,SP11 TM12,SP26 TM44,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) 申請人名稱,NVL(F1.FA04,Decode(F1.FA05,null,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) FC代理人名稱,NVL(F2.FA04,Decode(F2.FA05,null,F2.FA06,F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)) CP代理人名稱" & _
      " FROM R030301,SERVICEPRACTICE,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm,nation,caseprogress,customer,fagent f1,fagent f2" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0 AND caseprogress.CP09(+)=R030301.CP09" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16 AND S2.ST01(+)=R030301.CP14 AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) AND sp09=na01(+) AND substr(sp08,1,8)=cu01(+) AND substr(sp08,9,1)=cu02(+)" & _
      " AND substr(sp26,1,8)=f1.fa01(+) AND substr(sp26,9,1)=f1.fa02(+)" & _
      " AND substr(cp44,1,8)=f2.fa01(+) AND substr(cp44,9,1)=f2.fa02(+)"
   'strExc(0) = strExc(0) & " order by sort asc,承辦人 asc,本所案號 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'   Set grdDataList.Recordset = rsTmp
'   SetGrid
'   RecordShow
   Set m_adoRst = rsTmp.Clone 'Add
   If rsTmp.RecordCount > 0 Then
      'Set grdDataList.Recordset = rsTmp
      'Set m_adoRst = PUB_CreateRecordset(rsTmp, , , 300, Me.Name) 'Add
      'm_stSort = "sort asc,承辦人 asc,本所案號 asc" 'Add
      m_stSort = "TM01 asc,TM02 asc,TM03 asc,TM04 asc"
      m_adoRst.Sort = m_stSort 'Add
      SetRst2Grid 'Add
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   Else
      Screen.MousePointer = vbDefault
      SetRst2Grid 'Add
      MsgBox "查無資料！", vbInformation
      rsTmp.Close
      Set rsTmp = Nothing
      cmdHide.Enabled = False
      lblCnt.Caption = "共 0 筆"
      Exit Sub
   End If
   rsTmp.Close

   For iRow = 1 To grdDataList.Rows - 1
      '案件性質+相關總收文號的案件性質
      grdDataList.TextMatrix(iRow, 9) = grdDataList.TextMatrix(iRow, 9) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(iRow, 24), grdDataList.TextMatrix(iRow, 27), "1")
      '申請國家
      'A08 = "" & convForm(convForm(CheckStr(.Fields(8).Value), 8) & IIf(Val("" & .Fields("typ2")) > 2, "★收款後送件", ""), 20)
      If grdDataList.TextMatrix(iRow, 22) = "102" Then '延展
         grdDataList.TextMatrix(iRow, 8) = grdDataList.TextMatrix(iRow, 8) & IIf(Val(grdDataList.TextMatrix(iRow, 30)) > 2, "★收款後送件", "")
      End If
   Next iRow
End Sub

Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   'grdDataList.FixedCols = 3
   grdDataList.FixedCols = 0
End Sub

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 9 '8
      .Font.Size = 9
      '                0 1          2          3            4       5        6       7            8         9           10       11
      .FormatString = "V|本所期限  |法定期限  |承辦期限    |承辦人 |智權人員|事件　 |本所案號　　|申請國家 |案件性質   |備註    |案件名稱　　　　　　　　　　　　"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         If intI > 11 Then
            .ColWidth(intI) = 0
         End If
      Next
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(9) = flexAlignLeftCenter
      .ColAlignment(10) = flexAlignLeftCenter
      .Visible = True
   End With
End Sub

Private Sub SetColor(Optional sHide As String = "N")
Dim lngToday As Long, lngCP06 As Long, lngCP48 As Long, stType As String
Dim lngCP07 As Long
Dim ii As Integer, jj As Integer, dblCnt As Double

   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(2))
      For ii = 1 To .Rows - 1
         lngCP06 = Val(Replace(.TextMatrix(ii, 1), "/", "")) '本所期限
         lngCP07 = Val(Replace(.TextMatrix(ii, 2), "/", "")) '法定期限
         lngCP48 = Val(Replace(.TextMatrix(ii, 3), "/", "")) '承辦期限
         stType = .TextMatrix(ii, 12)
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
'            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
         
         '代碼1:A=達本所,B=達承辦,C=待提申,D=未收文,E=未發文,F=待完成,G=待收達,H=達法定
         '      I=達指會,J=可送件
         '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
         
         '逾管控期限
         If ((stType = "A" Or stType = "C" Or stType = "F" Or stType = "G") And lngCP06 > 0 And lngCP06 < lngToday) Or _
            ((stType = "B") And lngCP48 > 0 And lngCP48 < lngToday) Or _
            (stType = "J" And lngCP07 > 0 And lngCP07 < lngToday) Then
            .TextMatrix(ii, 7) = "*" & Trim(.TextMatrix(ii, 7))
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         ElseIf ((stType = "A" Or stType = "C" Or stType = "F" Or stType = "G") And lngCP06 = lngToday) Or _
            ((stType = "B") And lngCP48 = lngToday) Or _
            (stType = "J" And lngCP07 = lngToday) Then
            .TextMatrix(ii, 7) = "v" & Trim(.TextMatrix(ii, 7))
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         '可送件
         ElseIf stType = "J" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '紫
               .CellBackColor = &HFF80FF   '&HFF8080
            Next
         '未分案
         ElseIf .TextMatrix(ii, 13) = "0" Then
            .TextMatrix(ii, 7) = "#" & Trim(.TextMatrix(ii, 7))
'            '第五級不看
'            If bLvl5 = True Then
'               .RowHeight(ii) = 0
'            Else
               For jj = 1 To .Cols - 1
                  .col = jj
                  '黃
                  .CellBackColor = &HFFFF&
               Next
'            End If
         ElseIf sHide <> "N" Then
            .RowHeight(ii) = 0
         Else
            strExc(1) = .TextMatrix(ii, 13)
            Select Case strExc(1)
               '承辦人,核稿人
               Case "1", "4"
                  strExc(2) = .TextMatrix(ii, 15)
               Case "2" '管制人
                  strExc(2) = .TextMatrix(ii, 14)
               Case "3" '智權人員
                  strExc(2) = .TextMatrix(ii, 16)
               Case Else
                  strExc(2) = ""
            End Select

            If strExc(2) <> "" Then
               '例外情況
               If (Trim(txtUsernum) = "78011" Or Trim(txtUsernum) = "80030") And strExc(2) = "F4103" Then
                  '78011及80030為F4103的第二級主管
               Else
                  '本人或第二級才看
                  If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
                     .RowHeight(ii) = 0
                  End If
               End If
            End If
         End If
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With
   lblCnt.Caption = "共 " & dblCnt & " 筆"
   If sHide = "N" Then
      cmdHide.Tag = "Y"
      cmdHide.Caption = "隱藏白色"
   Else
      cmdHide.Tag = "N"
      cmdHide.Caption = "顯示白色"
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me

   Combo1.Clear
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v)：表示當日期限"
   Combo1.AddItem "黃色(#)：表示未分案"
   Combo1.AddItem "紫色：表示可送件"
   Combo1.AddItem "藍色：表示點選資料"
   Combo1.ListIndex = 0

   txtUsernum = strUserNum
   If Pub_StrUserSt03 = "M51" Then
      txtUsernum.Enabled = True
   End If
   'Added by Lydia 2023/05/10 開放輸入「員工編號」欄：總經理
   If InStr("01,08,", Pub_strUserST05 & ",") > 0 Then
      txtUsernum.Enabled = True
   End If
   'end 2023/05/10

   '從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020201 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 4
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

'Private Sub GrdDataList_Click()
'Dim nCol As Integer, nRow As Integer
'
'   With grdDataList
'      .Visible = False
'      nCol = .MouseCol
'      If nCol = 7 Then nCol = 29 '本所案號
'      nRow = .MouseRow
'      If nRow = 0 Then
'         .col = nCol
'         If m_blnColOrderAsc = False Then '字串降冪
'            .Sort = 5 '字串昇冪
'            m_blnColOrderAsc = True
'         Else
'            .Sort = 6 '字串降冪
'            m_blnColOrderAsc = False
'         End If
'      End If
'      .Visible = True
'   End With
'End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow grdDataList, X, Y, nCol, nRow
   grdDataList.col = IIf(nCol < 0, 0, nCol) 'nCol
   grdDataList.row = IIf(nRow < 0, 0, nRow) 'nRow
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iCol As Integer
   iCol = grdDataList.col
   If grdDataList.row < 1 Then
     grdDataList.Visible = False
     ChgEmptyDate True
     Set grdDataList.Recordset = Nothing
     If m_blnColOrderAsc = True Then
        m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc," & m_stSort
        m_blnColOrderAsc = False
     Else
        m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc," & m_stSort
        m_blnColOrderAsc = True
     End If
     SetRst2Grid
     SetGrid
     SetColor
     grdDataList.Visible = True
   End If
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To 0
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
