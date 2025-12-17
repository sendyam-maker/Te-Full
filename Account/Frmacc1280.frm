VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1280 
   AutoRedraw      =   -1  'True
   Caption         =   "收文與收據資料檢核查詢"
   ClientHeight    =   5240
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5240
   ScaleWidth      =   9270
   Begin VB.TextBox txtExcept 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "N"
      Top             =   150
      Width           =   435
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1280.frx":0000
      Height          =   3945
      Left            =   50
      TabIndex        =   3
      Top             =   510
      Width           =   8900
      _ExtentX        =   15699
      _ExtentY        =   6967
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "收文與收據資料檢核查詢"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "conf"
         Caption         =   "確"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "status"
         Caption         =   "狀"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "rDate"
         Caption         =   "收文日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cp09"
         Caption         =   "總收文號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "caseNo"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "custNo"
         Caption         =   "客戶代號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "custName"
         Caption         =   "申請人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "salesname"
         Caption         =   "智權人員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "caseProp"
         Caption         =   "案件性質"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "naName"
         Caption         =   "申請國家"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "cp16"
         Caption         =   "費用"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "cp17"
         Caption         =   "規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "cp18"
         Caption         =   "點數"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "CP27"
         Caption         =   "發文日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   349.795
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1670.173
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1310.173
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1069.795
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
         EndProperty
         BeginProperty Column13 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   60
      Top             =   900
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1275
      TabIndex        =   0
      Top             =   157
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3195
      TabIndex        =   1
      Top             =   157
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "2.'確'欄：Y：可開收據，N：專業部未列印定稿，W：專業部未修改金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   900
      TabIndex        =   9
      Top             =   4780
      Width           =   6580
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "注意：1.非接洽記錄單收文,  7 天後開收據                         '狀'欄：紙：紙本接洽單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   320
      TabIndex        =   8
      Top             =   4520
      Width           =   7240
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "含例外處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5355
      TabIndex        =   7
      Top             =   180
      Width           =   1245
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7275
      TabIndex        =   6
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   5
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收文日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc1280"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 H5500
   Me.Width = 9390
   Me.Height = 5700
   'end 2023/10/06
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   FormClear
   'add by sonia 2019/7/15
   MaskEdBox1 = CFDate(ChangeWStringToTString(CompDate(1, -1, (Left(strSrvDate(1), 6) & "01"))))   '預設上月1日
   MaskEdBox2 = CFDate(ChangeWStringToTString(strSrvDate(1)))
   'end 2019/7/15
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1280 = Nothing
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   Screen.MousePointer = vbHourglass

   Dim stCon As String, stVTable As String, stCon1 As String
   Dim stConPA As String, stConTM As String, stConSP As String, stConLC As String
   
On Error GoTo Checking

   stCon = "": stConPA = "": stConTM = "": stConSP = "": stConLC = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stCon = stCon & " and cp05 >= " & Val(CADate(FCDate(MaskEdBox1.Text))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stCon = stCon & " and cp05 <= " & Val(CADate(FCDate(MaskEdBox2.Text))) & ""
   End If
   
   If txtExcept.Text = "N" Then
      'Modify by Morgan 2011/3/24 改智權人員收文或沒有FC代理人的案件都算非例外
      'stCon = stCon & " and not ( cp05 >= 20030101 and substr(cp01, 1, 2) = 'FC' and cp10 not in ('907') )"
      ''Modify by Morgan 2010/12/7 不必再限制客戶國籍
      ''stCon1 = stCon1 & " and cu10<010"
      ''Add by Morgan 2010/12/9
      'stConPA = stConPA & " and pa75 is null"
      'stConTM = stConTM & " and tm44 is null"
      'stConSP = stConSP & " and sp26 is null"
      'stConLC = stConLC & " and lc22 is null"
      stConPA = stConPA & " and (pa75 is null or substr(cp12,1,1)='S')"
      stConTM = stConTM & " and (tm44 is null or substr(cp12,1,1)='S')"
      stConSP = stConSP & " and (sp26 is null or substr(cp12,1,1)='S')"
      stConLC = stConLC & " and (lc22 is null or substr(cp12,1,1)='S')"
   Else
      'Modify by Morgan 2011/3/24 改智權人員收文或沒有FC代理人的案件都算非例外
      'stConPA = stConPA & " and (substr(cp01, 1, 2) = 'FC' or pa75 is null)"
      'stConTM = stConTM & " and (substr(cp01, 1, 2) = 'FC' or tm44 is null)"
      'stConSP = stConSP & " and (substr(cp01, 1, 2) = 'FC' or sp26 is null)"
      'stConLC = stConLC & " and (substr(cp01, 1, 2) = 'FC' or lc22 is null)"
     'modify by sonia 2023/5/5 substr(cp12,1,1)='S'改為substr(cp12,1,1)<>'F',否則唐韻如收文國外客戶就不會帶出來
     stCon1 = stCon1 & " and (FNo is null or substr(cp12,1,1)<>'F' or (substr(cp01, 1, 2) = 'FC' and nvl(cu10,' ')<'011' ))"
   End If
   
   stCon = stCon & " And CP20 Is Null "
   '無收據編號(cp60=null)或無收據資料(a0k01=null)
   'FCP, FCT, FCL未發文的不抓
   '2008/10/21 modify by sonia CP13改抓智權人員姓名
   'Modify by Morgan 2011/3/24 +FNo,CP12
   'modify by sonia 2015/10/30 +cp140
   'Modified by Morgan 2023/1/7 +cp31
   'modify by sonia 2023/5/17 加發文日CP27
   stVTable = " select cp05, cp09, cp01, cp02, cp03, cp04, nvl(pa26, nvl(pa27, nvl(pa28, nvl(pa29, pa30)))) as CustNo, st02 salesname, cp10, pa09 as naNo, cp16, cp17, cp18, cp27" & _
            ",pa75 FNo,cp12,cp140,cp31 from caseprogress, acc0k0, patent, staff " & _
            " where (cp60 is null or substr(cp60,1,1)='E') and (cp16 is not null and cp16 <> 0) and cp57 is null And 1=decode(cp01,'FCP',DECODE(cp27,null,1,0),1) " & stCon & _
            " and a0k01(+)=cp60 and a0k01 is null" & stConPA & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null and cp13=st01(+) " & _
            " union all" & _
            " select cp05, cp09, cp01, cp02, cp03, cp04, tm23 as CustNo, ST02 salesname, cp10, tm10 as naNo, cp16, cp17, cp18, cp27" & _
            ",tm44 FNo,cp12,cp140,cp31 from caseprogress, acc0k0, trademark, staff " & _
            " where (cp60 is null or substr(cp60,1,1)='E') and  (cp16 is not null and cp16 <> 0) and cp57 is null And 1=decode(cp01,'FCT',DECODE(cp27,null,1,0),1) " & stCon & _
            " and a0k01(+)=cp60 and a0k01 is null" & stConTM & _
            " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null and cp13=st01(+) " & _
            " union all" & _
            " select cp05, cp09, cp01, cp02, cp03, cp04, lc11 as CustNo, ST02 salesname, cp10, lc15 as naNo, cp16, cp17, cp18, cp27" & _
            ",lc22 FNo,cp12,cp140,cp31 from caseprogress, acc0k0, lawcase, staff " & _
            " where (cp60 is null or substr(cp60,1,1)='E') and  (cp16 is not null and cp16 <> 0) and cp57 is null" & stCon & _
            " and a0k01(+)=cp60 and a0k01 is null" & stConLC & _
            " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null and cp13=st01(+) " & _
            " union all" & _
            " select cp05, cp09, cp01, cp02, cp03, cp04, nvl(sp08, nvl(sp58, sp59)) as CustNo, ST02 salesname, cp10, sp09 as naNo, cp16, cp17, cp18, cp27" & _
            ",sp26 FNo,cp12,cp140,cp31 from caseprogress, acc0k0, servicepractice, staff " & _
            " where (cp60 is null or substr(cp60,1,1)='E') and  (cp16 is not null and cp16 <> 0) and cp57 is null" & stCon & _
            " and a0k01(+)=cp60 and a0k01 is null" & stConSP & _
            " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null and cp13=st01(+) "
            
   'Add by Morgan 2005/1/7 加顧問案件
   'modify by sonia 2015/10/30 +cp140
   'Modified by Morgan 2023/1/7 +cp31
   stVTable = stVTable & " union all" & _
            " select cp05, cp09, cp01, cp02, cp03, cp04, hc05 as CustNo, ST02 salesname, cp10, '000' as naNo, cp16, cp17, cp18, cp27" & _
            ",'' FNo,cp12,cp140,cp31 from caseprogress, acc0k0, hirecase, staff " & _
            " where cp60 is null and (cp16 is not null and cp16 <> 0) and cp57 is null" & stCon & _
            " and a0k01(+)=cp60 and a0k01 is null" & _
            " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04 and hc01 is not null and cp13=st01(+) "
            
   'Modify by Morgan 2005/1/12 加控制申請人非台灣的要例外處理
   'Modify by Morgan 2004/12/1 排除無申請人的資料
   '2011/4/27 modify by sonia 加智權部是否已確認碼
   'strSql = " select cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(cpm03,cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18" & _
            " from (" & stVTable & ") X,Customer,casepropertymap,Nation" & _
            " where custNo is not null and cu01(+)=substr(custNo,1,8) and cu02(+)=substr(custNo,9,1)" & stCon1 & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=naNo order by rDate asc, cp09 asc "
   '2011/8/26 modify by sonia 智權部已確認但專業部未改cp16則顯示W
   'strSql = " select decode(lc08||lc13,null,decode(lc01,null,null,'N'),'Y') conf,cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(cpm03,cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18" & _
            " from (" & stVTable & ") X,Customer,casepropertymap,Nation,LetterCache" & _
            " where custNo is not null and cu01(+)=substr(custNo,1,8) and cu02(+)=substr(custNo,9,1)" & stCon1 & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=naNo and cp09=lc01(+) order by rDate asc, cp09 asc "
   '2014/10/28 modify by sonia 改為專業部列印後才能印,否則可能收據先開但定稿還沒印, 2014/11/6專業部未改cp16仍顯示W
   'strSql = " select decode(lc08||lc13,null,decode(lc01,null,null,'N'),decode(lcv06,cp16,'Y',null,'Y','W')) conf,cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(cpm03,cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18" & _
            " from (" & stVTable & ") X,Customer,casepropertymap,Nation,LetterCache,LetterCachevar" & _
            " where custNo is not null and cu01(+)=substr(custNo,1,8) and cu02(+)=substr(custNo,9,1)" & stCon1 & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=naNo and cp09=lc01(+) and cp09=lcv01(+) and 'Y'=lcv05(+) order by rDate asc, cp09 asc "
   '2015/1/20 modify by sonia 將nvl(cpm03,cpm10) caseProp 改為nvl(decode(cpm03,'（無）',cpm04,cpm03),cpm10) caseProp (P-109072代理人請款)
   'modify by sonia 2015/10/30 自動收文狀態放"自"
   'Modified by Lydia 2017/02/21 抓確定已收文並且有費用,排除下一程序管控而產生的報表(Ex CFP-26318同時有專利證書和期限管制表的定稿,會產生4筆明細)
   'strSql = " select decode(lc01,null,null,decode(lc13,null,'N',decode(lcv06,cp16,'Y',null,'Y','W'))) conf,decode(cp140,'','','自') Status,cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(decode(cpm03,'（無）',cpm04,cpm03),cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18" & _
            " from (" & stVTable & ") X,Customer,casepropertymap,Nation,LetterCache,LetterCachevar" & _
            " where custNo is not null and cu01(+)=substr(custNo,1,8) and cu02(+)=substr(custNo,9,1)" & stCon1 & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=naNo and cp09=lc01(+) and cp09=lcv01(+) and 'Y'=lcv05(+) order by rDate asc, cp09 asc "
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。 2023/10/25取消新紙，法律所案件一律為紙
   'strSql = " select decode(lc01,null,null,decode(lc13,null,'N',decode(lcv06,cp16,'Y',null,'Y','W'))) conf,decode(cp140,'','','自') Status,cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(decode(cpm03,'（無）',cpm04,cpm03),cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18"
   'modify by sonia 2023/5/17 加發文日CP27
   strSql = " select decode(lc01,null,null,decode(lc13,null,'N',decode(lcv06,cp16,'Y',null,'Y','W'))) conf,decode(instr(cp01,'L'),0,'','紙') Status,cp05-19110000 as rDate,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseNo,custNo,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) custName, salesname, nvl(decode(cpm03,'（無）',cpm04,cpm03),cpm10) caseProp, nvl(na03, na04) naName,cp16,cp17,cp18,sqldatet(cp27) as CP27 " & _
            " from (" & stVTable & ") X,Customer,casepropertymap,Nation,LetterCache,LetterCachevar" & _
            " where custNo is not null and cu01(+)=substr(custNo,1,8) and cu02(+)=substr(custNo,9,1)" & stCon1 & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=naNo and cp09=lc01(+) and '0'=lc02(+) and cp09=lcv01(+) and '0'=lcv02(+) and 'Y'=lcv05(+) order by rDate asc, cp09 asc "
   'end 2017/02/21
   'END 2014/10/28
   '2011/8/26 END
            
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoRecordset
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
   End If
   
Checking:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   
End Sub
'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
End Sub
'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function
'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub txtExcept_GotFocus()
   TextInverse txtExcept
End Sub

Private Sub txtExcept_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = Asc("N")
   End If
End Sub
