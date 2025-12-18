VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040210 
   BorderStyle     =   1  '單線固定
   Caption         =   "未發文案件查詢"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9390
   Begin VB.CommandButton Command1 
      Caption         =   "取消紀錄缺文件(&D)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   5895
      TabIndex        =   26
      Top             =   1320
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "紀錄缺文件(&R)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   7875
      TabIndex        =   25
      Top             =   1320
      Width           =   1425
   End
   Begin VB.TextBox txtDoc 
      Height          =   264
      Left            =   1755
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1818
      Width           =   330
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1035
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1500
      Width           =   555
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   3090
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1500
      Width           =   405
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   2730
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1500
      Width           =   330
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1500
      Width           =   990
   End
   Begin VB.TextBox txtCP10 
      Height          =   285
      Left            =   1035
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1170
      Width           =   735
   End
   Begin VB.TextBox txtCust 
      Height          =   285
      Left            =   1035
      MaxLength       =   9
      TabIndex        =   3
      Top             =   840
      Width           =   1185
   End
   Begin VB.TextBox txtDate 
      Height          =   270
      Index           =   1
      Left            =   3105
      MaxLength       =   7
      TabIndex        =   2
      Top             =   525
      Width           =   870
   End
   Begin VB.TextBox txtDate 
      Height          =   270
      Index           =   0
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   1
      Top             =   525
      Width           =   870
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5310
      TabIndex        =   10
      Top             =   1770
      Width           =   3990
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   4455
      TabIndex        =   12
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtCP14 
      Height          =   285
      Left            =   1035
      MaxLength       =   6
      TabIndex        =   0
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5355
      TabIndex        =   11
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   15
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2745
      Left            =   45
      TabIndex        =   16
      Top             =   2070
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   4842
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "V|收文日 |本所案號 |案件名稱 |申請國家 |專利種類 |案件性質|承辦期限 |本所期限 |未收金額 |缺文件 |指定日期|備註 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin MSForms.Label lblCust 
      Height          =   255
      Left            =   2250
      TabIndex        =   35
      Top             =   870
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1830
      TabIndex        =   34
      Top             =   1200
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   1980
      TabIndex        =   33
      Top             =   210
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAutoColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000C000&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   1
      Left            =   2385
      TabIndex        =   32
      Top             =   5310
      Width           =   150
   End
   Begin VB.Label lblAutoColor 
      Appearance      =   0  '平面
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   2
      Left            =   4140
      TabIndex        =   31
      Top             =   5310
      Width           =   150
   End
   Begin VB.Label lblAutoColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000FF7F&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   720
      TabIndex        =   30
      Top             =   5310
      Width           =   150
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "　　　3.    底色表示自動收文      底色表示當日期限      底色表示逾期限"
      Height          =   180
      Left            =   45
      TabIndex        =   29
      Top             =   5310
      Width           =   5535
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "　　　2.本查詢同時會列出期限區間迄日後2個工作天達指定送件日期的程序"
      Height          =   180
      Left            =   45
      TabIndex        =   28
      Top             =   5100
      Width           =   5985
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "備註：1.本查詢同時會列出已收款且有紀錄未收款無法發文程序"
      Height          =   180
      Left            =   48
      TabIndex        =   27
      Top             =   4896
      Width           =   4992
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含缺文件案件：         ( Y : 含)"
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   1860
      Width           =   2625
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   90
      TabIndex        =   23
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   90
      TabIndex        =   22
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   90
      TabIndex        =   21
      Top             =   892
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   2790
      X2              =   3120
      Y1              =   653
      Y2              =   653
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦或本所期限區間："
      Height          =   180
      Left            =   90
      TabIndex        =   20
      Top             =   570
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4500
      TabIndex        =   19
      Top             =   1830
      Width           =   780
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   18
      Top             =   4890
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   90
      TabIndex        =   17
      Top             =   262
      Width           =   720
   End
End
Attribute VB_Name = "frm040210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; lblUserName、lblCust、Label8 ; Printer列印未改
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Add by Morgan 2009/12/3
Option Explicit
Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrint As String
Dim prnPrint As Printer
Dim m_iCols As Integer
Dim m_iLstRow As Integer

Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   grdDataList.FixedCols = 3
   SetGrid
   SetColor
End Sub
Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
On Error GoTo ErrorHandler
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         
         Dim Str01 As String
         
         StrTag = grdDataList.TextMatrix(i, 2)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Mid(StrTag, 2)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         
         
         Select Case cmdState
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P"   '專利
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     
                  Case "FG"
                     frm100101_B.Show
                     frm100101_3.Tag = StrTag
                     frm100101_B.StrMenu
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         
         RowSelect True
         Exit For
      End If
   Next i
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
      MsgBox "(" & Err.NUMBER & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   PUB_RestorePrinter cboPrinter.Text
   DoPrint
   PUB_RestorePrinter strPrint
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
   doQuery
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim iRow As Integer
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   With grdDataList
   For iRow = 1 To .Rows - 1
      .col = 0
      .row = iRow
      If .Text = "V" Then
         If .TextMatrix(.row, 15) <> "" Then
            If Index = 0 Then
               If .TextMatrix(.row, 10) = "Y" Then
                  MsgBox .TextMatrix(.row, 2) & " 已設定缺文件，不必再紀錄！"
                  Exit For
               Else
                  If UpdateUD(.TextMatrix(.row, 14), .TextMatrix(.row, 15), "2") Then
                     .TextMatrix(.row, 10) = "Y"
                     RowSelect True
                  End If
               End If
            Else
               If .TextMatrix(.row, 10) = "Y" Then
                  If UpdateUD(.TextMatrix(.row, 14), .TextMatrix(.row, 15), "1") Then
                     .TextMatrix(.row, 10) = ""
                     RowSelect True
                  End If
               Else
                  MsgBox .TextMatrix(.row, 2) & " 未設定缺文件，無法取消！"
                  Exit For
               End If
            End If
         Else
            MsgBox .TextMatrix(.row, 2) & " 未設定收款後送件，不可設定缺文件！"
            Exit For
         End If
         
      End If
   Next
   End With
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Function UpdateUD(stUD01 As String, stUD02 As String, stUD03 As String) As Boolean
   strSql = "update undeliveredrec set ud03='" & stUD03 & "' where ud01='" & stUD01 & "' and ud02=" & stUD02
   cnnConnection.Execute strSql, intI
   If intI = 1 Then
      UpdateUD = True
   Else
      MsgBox "更新失敗！"
   End If
End Function

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

Private Sub doQuery()
   Dim stConCP As String, stConPA As String, stConUD As String
   Dim stConCP48 As String, stConCP06 As String, stConCP142 As String
   
   stConCP = ""
   stConPA = ""
   stConUD = ""
   stConCP48 = ""
   stConCP06 = ""
   If txtCP14 <> "" Then
      stConCP = stConCP & " and cp14='" & txtCP14 & "'"
      pub_QL05 = pub_QL05 & ";" & Label4 & txtCP14 & lblUserName 'Add By Sindy 2010/12/3
   End If
   
   If txtDate(0) <> "" Then
      stConCP48 = stConCP48 & " and cp48>=" & DBDATE(txtDate(0))
      stConCP06 = stConCP06 & " and cp06>=" & DBDATE(txtDate(0))
      
   '若無起始日則抓收文日2年內的以便加速
   Else
      stConCP = stConCP & " and cp05>" & (strSrvDate(1) - 20000)
   End If
   
   If txtDate(1) <> "" Then
      stConCP48 = stConCP48 & " and cp48<=" & DBDATE(txtDate(1))
      stConCP06 = stConCP06 & " and cp06<=" & DBDATE(txtDate(1))
      stConCP142 = stConCP142 & " and cp142<=" & CompWorkDay(3, DBDATE(txtDate(1))) 'Add by Morgan 2011/3/9
   End If
   
   If txtDate(0) <> "" Or txtDate(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & txtDate(0) & "-" & txtDate(1) 'Add By Sindy 2010/12/3
   End If

   If txtCust <> "" Then
      stConPA = stConPA & " and instr(pa26||pa27||pa28||pa29||pa30,'" & txtCust & "')>0"
      pub_QL05 = pub_QL05 & ";" & Label3 & txtCust & lblCust 'Add By Sindy 2010/12/3
   End If
   
   If txtCP10 <> "" Then
      stConCP = stConCP & " and cp10='" & txtCP10 & "'"
      pub_QL05 = pub_QL05 & ";" & Label5 & txtCP10 & Label8 'Add By Sindy 2010/12/3
   End If
   
   If txtSystem <> "" Then
      stConCP = stConCP & " and cp01='" & txtSystem & "'"
      pub_QL05 = pub_QL05 & ";" & Label6 & txtSystem 'Add By Sindy 2010/12/3
   End If
   
   If txtCode(0) <> "" Then
      stConCP = stConCP & " and cp02='" & txtCode(0) & "'"
      pub_QL05 = pub_QL05 & "-" & txtCode(0) 'Add By Sindy 2010/12/3
      If txtCode(1) <> "" Then
         stConCP = stConCP & " and cp03='" & txtCode(1) & "'"
         pub_QL05 = pub_QL05 & "-" & txtCode(1) 'Add By Sindy 2010/12/3
      End If
      If txtCode(2) <> "" Then
         stConCP = stConCP & " and cp04='" & txtCode(2) & "'"
         pub_QL05 = pub_QL05 & "-" & txtCode(2) 'Add By Sindy 2010/12/3
      End If
   End If
   
   '是否含缺文件案件
   If txtDoc <> "Y" Then
      stConUD = stConUD & " having substrb(max(ud02||ud03),-1)='1'"
      pub_QL05 = pub_QL05 & ";" & Left(Label9, 9) & txtDoc 'Add By Sindy 2010/12/3
   End If
   '1.達承辦期限
   '2.達本所期限
   '3.已收款(有無法發文紀錄)
   '4.達指定發文日(前2個工作天小於等於區間迄日)
   'Modify by Morgan 2011/3/9 +達指定發文日
   'Modify by Morgan 2011/10/13 +CP140
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strExc(0) = "select '' V,substrb(sqldatet(cp05),1,10) 收文日" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",pa05 案件名稱,na03 申請國家,decode(pa09,'020',PTM04,PTM03) 專利種類" & _
      ",decode(pa09,'000',CPM03,CPM04) 案件性質,sqldatet(cp48) 承辦期限,sqldatet(cp06) 本所期限,nvl(cp79,0) 未收金額" & _
      ",decode(UD03,'2','Y') 缺文件,substrb(sqldatet(cp142),1,9) 指定日期,cp64 進度備註,NVL(CP48,CP06) SRT,CP09,UD02,CP140" & _
      " from (select cp09 X01 from caseprogress where cp27 is null and cp57 is null" & stConCP & stConCP48 & _
      " union select cp09 from caseprogress where cp27 is null and cp57 is null" & stConCP & stConCP06 & _
      " union select cp09 from caseprogress where cp27 is null and cp57 is null and cp142>0" & stConCP & stConCP142 & _
      " union select cp09 from caseprogress,UndeliveredRec where cp27 is null and cp57 is null and cp16>0 and cp79=0" & stConCP & _
      " and Ud01(+)=cp09 and Ud02>0 group by cp09" & stConUD & _
      " ) X,caseprogress,patent,nation,PatentTrademarkMap,UndeliveredRec,casepropertymap" & _
      " where cp09(+)=X01 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & stConPA & _
      " and pa57 is null and na01(+)=pa09 and PTM01(+)='1' and PTM02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and UD01(+)=cp09 and (UD02 is null or UD02=(select max(b.ud02) from UndeliveredRec b where b.UD01=cp09))"
  
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   
   If RsTemp Is Nothing Then Exit Sub
   If RsTemp.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/3
      Set m_adoRst = RsTemp.Clone
      SetRst2Grid
      MsgBox "查無資料！", vbInformation
      lblCnt.Caption = "共 0 筆"
   Else
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/3
      'Modify by Amy 2014/06/09 +FormName 改存暫存TB
      'Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300)
      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
      m_stSort = "SRT asc,本所案號 asc"
      m_adoRst.Sort = m_stSort
      SetRst2Grid
      RecordShow
      m_blnColOrderAsc = True
   End If
End Sub

Private Sub SetGrid()
   m_iCols = 12
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      .FormatString = "V|收文日 　|本所案號　　　　 |案件名稱　|申請國家|專利種類|案件性質|承辦期限|本所期限|未收金額|缺文件|指定日期|備註　　　　　　　　"
      .ColWidth(0) = 225
      .ColWidth(1) = 810
      .ColWidth(2) = 1330
      .ColWidth(3) = 975
      .ColWidth(4) = 735
      .ColWidth(5) = 500
      .ColWidth(6) = 735
      .ColWidth(7) = 810
      .ColWidth(8) = 810
      .ColWidth(9) = 750
      .ColWidth(10) = 550
      .ColWidth(11) = 810
      .ColWidth(12) = 2000
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = flexAlignLeftCenter
         If intI > m_iCols Then
            .ColWidth(intI) = 0
         End If
      Next
      '置中
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(10) = flexAlignCenterCenter
      '靠右
      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(7) = flexAlignRightCenter
      .ColAlignment(8) = flexAlignRightCenter
      .ColAlignment(9) = flexAlignRightCenter
      .ColAlignment(11) = flexAlignRightCenter
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtCP14 = strUserNum
   If Pub_StrUserSt03 = "M51" Then
      txtCP14.Enabled = True
   Else
      txtCP14.Enabled = False
   End If
   txtDate(1) = strSrvDate(2)
   PUB_SetPrinter Me.Name, cboPrinter, strPrint
   Forms(0).StatusBar1.Panels(1).Text = ""
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If cboPrinter.Text <> cboPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
   End If
   Set frm040210 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer, iRow As Integer
    iCol = grdDataList.MouseCol
    iRow = grdDataList.MouseRow
    If iRow < 1 Then
      grdDataList.Visible = False
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc," & m_stSort
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc," & m_stSort
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      
      grdDataList.Visible = True
    End If
End Sub

Private Sub RowSelect(Optional bolUnSel As Boolean = False)
   Dim ii As Integer
   With grdDataList
   If .row > 0 Then
      .col = 0
      If .Text = "V" Or bolUnSel Then
         .Text = ""
         .col = 3
         For ii = 0 To 2
            .col = ii
            .CellBackColor = &HFFFFFF
         Next
      Else
         .Text = "V"
         For ii = 0 To 2
            .col = ii
            .CellBackColor = &HFFC0C0
         Next
      End If
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long, iRow As Integer
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         RowSelect
         iRow = m_iLstRow
         m_iLstRow = .row
         If m_iLstRow <> iRow Then
            If iRow > 0 And iRow < .Rows Then
               .row = iRow
               RowSelect True
            End If
         End If
         .row = m_iLstRow
         .Visible = True
      End If
   End With
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
   CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP10_GotFocus()
   TextInverse txtCP10
   CloseIme
End Sub

Private Sub txtCP14_Change()
   If Len(txtCP14) >= 5 Then
      lblUserName = GetStaffName(txtCP14, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub SetColor()
   Dim ii As Integer, jj As Integer, dblCnt As Double
   
   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      For ii = 1 To .Rows - 1
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellFontSize = 9
         Next
         
         'Add by Morgan 2011/10/13 區分自動收文
         For jj = 3 To .Cols - 1
            If .TextMatrix(ii, 16) <> "" Then
               .col = jj
               .CellBackColor = lblAutoColor(0).BackColor
            End If
               
            'Added by Morgan 2018/12/18
            '+當日期限,逾期限也變色( Ex:P088661 年費已過指定日期而未注意 )
            If jj = 7 And .TextMatrix(ii, jj) <> "" Then
               If DBDATE(.TextMatrix(ii, jj)) <= strSrvDate(1) Then
                  .col = jj
                  If DBDATE(.TextMatrix(ii, jj)) = strSrvDate(1) Then
                     .CellBackColor = lblAutoColor(1).BackColor
                  Else
                     .CellBackColor = lblAutoColor(2).BackColor
                  End If
               End If
            ElseIf jj = 8 And .TextMatrix(ii, jj) <> "" Then
               If DBDATE(.TextMatrix(ii, jj)) <= strSrvDate(1) Then
                  .col = jj
                  If DBDATE(.TextMatrix(ii, jj)) = strSrvDate(1) Then
                     .CellBackColor = lblAutoColor(1).BackColor
                  Else
                     .CellBackColor = lblAutoColor(2).BackColor
                  End If
               End If
            ElseIf jj = 11 And .TextMatrix(ii, jj) <> "" Then
               If DBDATE(.TextMatrix(ii, jj)) <= strSrvDate(1) Then
                  .col = jj
                  If DBDATE(.TextMatrix(ii, jj)) = strSrvDate(1) Then
                     .CellBackColor = lblAutoColor(1).BackColor
                  Else
                     .CellBackColor = lblAutoColor(2).BackColor
                  End If
               End If
            End If
            'end 2018/12/18
         Next
         
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With

   lblCnt.Caption = "共 " & dblCnt & " 筆"

End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      ReDim strTemp(1 To m_iCols)
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            Select Case iCol
               Case 3, 4, 5, 6, 10, 12
                  strTemp(iCol) = StrToStr(.TextMatrix(iRow, iCol), Len(.TextMatrix(0, iCol)))
               Case Else
                  strTemp(iCol) = .TextMatrix(iRow, iCol)
            End Select
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To m_iCols)
   PLeft(1) = ciStartX
   For intI = 2 To m_iCols
      PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1)) + ciColGap
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(130, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Select Case iCol
         Case 1, 7, 8, 9
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
         Case Else
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
      End Select
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Me.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "員工編號：" & txtCP14 & " 姓名：" & lblUserName
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(130, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To m_iCols
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print grdDataList.TextMatrix(0, intI)
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCust_GotFocus()
   TextInverse txtCust
   CloseIme
End Sub

Private Sub txtCust_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCust_Validate(Cancel As Boolean)
   If txtCust <> "" Then
      If ClsPDGetCustomer(txtCust, strExc(1)) Then
         lblCust = strExc(1)
      Else
         lblCust = ""
      End If
   Else
      lblCust = ""
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
   CloseIme
End Sub

Private Sub txtDoc_GotFocus()
   TextInverse txtDoc
   CloseIme
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSystem_GotFocus()
   TextInverse txtSystem
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
