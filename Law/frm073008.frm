VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm073008 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問到期明細表"
   ClientHeight    =   5610
   ClientLeft      =   1635
   ClientTop       =   2100
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdSend 
      Caption         =   "寄E-mail"
      Default         =   -1  'True
      Height          =   400
      Left            =   1440
      TabIndex        =   6
      Top             =   70
      Width           =   915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm073008.frx":0000
      Height          =   2055
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "a01"
         Caption         =   "部門"
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
      BeginProperty Column01 
         DataField       =   "a02"
         Caption         =   "E-mail收件人"
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
         DataField       =   "a03"
         Caption         =   "姓名"
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2640
      Width           =   345
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2970
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2212
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2212
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   4125
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   960
         Style           =   2  '單純下拉式
         TabIndex        =   12
         Top             =   200
         Width           =   2955
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1837
      Width           =   345
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3570
      TabIndex        =   8
      Top             =   70
      Width           =   760
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2970
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1387
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1387
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   2730
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2640
      Top             =   4320
      Visible         =   0   'False
      Width           =   1200
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
   Begin VB.Line Line2 
      X1              =   2610
      X2              =   2850
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "是否以E-mail寄發各區主管：            (Y:是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   450
      TabIndex        =   15
      Top             =   2663
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   14
      Top             =   2235
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "是否依業務區不同跳頁：            (Y:是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   10
      Top             =   1860
      Width           =   3555
   End
   Begin VB.Line Line1 
      X1              =   2610
      X2              =   2850
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "到期日期 ："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   9
      Top             =   1410
      Width           =   1215
   End
End
Attribute VB_Name = "frm073008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/01 已與莊敏惠確認，邱素蓮當時已不在定時跑程式通知主管；所以現在都是每月批次自動通知StrMenu13
'Memo by Lydia 2021/09/01 Form2.0已修改; DataGrid1改字型新細明體=ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 10) As Integer
Dim m_print As Integer
'Add By Cheng 2002/09/09
Dim blnClkSure As Boolean
'Add By Cheng 2003/03/31
Dim iPrint As Integer
'Added by Lydia 2015/03/11 列印用
Dim prnPrint As Printer
Dim strPrint As String
'Added by Lydia 2015/03/13 Grid
Dim AdoEmail As New ADODB.Recordset
Dim mESeqNo As String '暫存TB編號
Dim pST01 As String
Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
'Modified by Lydia 2015/03/16 改funtion
'    'Add By Cheng 2002/09/09
'    blnClkSure = False
'    m_print = 0
'    If ChkRange(Text1(0), Text1(1), "到期日期") = False Then
'       blnClkSure = True
'       Exit Sub
'    End If
'    'Add By Cheng 2002/03/22
'    If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
'       Me.Text1(0).SetFocus
'       Text1_GotFocus 0
'       Exit Sub
'    End If
'    If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
'       Me.Text1(1).SetFocus
'       Text1_GotFocus 1
'       Exit Sub
'    End If
    If CheckCase = False Then
       Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    GetPrintLeft
    PrintCase
    If m_print = 0 Then
        MsgBox "列印結束!", vbInformation
        'Add By Cheng 2003/03/31
        '列印接洽結案單
        PUB_PrintCaseCloseSheet strUserNum, "1"
        '刪除接洽結案單暫存資料
        PUB_DeleteCaseCloseSheet strUserNum
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintCase()
Dim i As Integer, St As String, Page As Integer ', iPrint As Integer
Dim TmpArea As String
'Add By Cheng 2003/03/31
Dim strSaleZone As String '業務區

On Error GoTo ErrHand
    'Modify By Cheng 2003/03/31
'    strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02)," & _
'                        "HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04)," & _
'                        "SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2)," & _
'                        "SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)," & _
'                        "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
'                        "CP16,HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),CU16,CU79 " & _
'                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER,ACC090 WHERE " & _
'                        "CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP13=ST01(+) AND " & _
'                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+))" & _
'                        strGetcdnSQL & " ORDER BY CP12,CP13,CP01||CP02||CP03||CP04"
'edit by nick 2004/10/22
'    strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02)," & _
'                        "HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04)," & _
'                        "SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2)," & _
'                        "SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)," & _
'                        "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
'                        "CP16,HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),CU16,CU79, HC01, HC02, HC03, HC04 " & _
'                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER,ACC090 WHERE " & _
'                        "CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP13=ST01(+) AND " & _
'                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+))" & _
'                        strGetcdnSQL & " ORDER BY CP12,CP13,CP01||CP02||CP03||CP04"
'Modified by Lydia 2015/03/20 baseTable改caseprogress
    'Modified by Lydia 2021/09/01 區分案源和非案源
    'strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02)," & _
                        "HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04)," & _
                        "SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2)," & _
                        "SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)," & _
                        "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
                        "CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))),C1.CU16,C1.CU79, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))),C2.CU16,C2.CU79 " & _
                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER C1,CUSTOMER C2,ACC090 WHERE " & _
                        "HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=ST01(+) AND " & _
                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
                        strGetcdnSQL & "  and cp27 is null ORDER BY CP12,CP13,CP01||CP02||CP03||CP04"
    strExc(0) = strGetcdnSQL(strExc(5))
    strExc(0) = "SELECT DECODE(CP12,A0901,A0902) as B01,DECODE(CP13,ST01,ST02) as B02," & _
                        "HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) as B03," & _
                        "SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) as B04," & _
                        "SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) as B05," & _
                        "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) as B06," & _
                        "CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) as B07,C1.CU16 as C1CU16,C1.CU79 AS C1CU79," & _
                        "HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) as B08,C2.CU16 AS C2CU16,C2.CU79 AS C2CU79, CP12, CP13 " & _
                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER C1,CUSTOMER C2,ACC090 WHERE " & _
                        "HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=ST01(+) AND " & _
                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
                        Replace(UCase(strExc(0)), "S1.", "") & "  and cp27 is null "
    strExc(5) = "SELECT A0902 AS B01,ST02 AS B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) AS B03," & _
                        "SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) as B04," & _
                        "SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) as B05," & _
                        "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) as B06," & _
                        "CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) as B07,C1.CU16 as C1CU16,C1.CU79 AS C1CU79," & _
                        "HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) as B08,C2.CU16 AS C2CU16,C2.CU79 AS C2CU79, ST15 AS CP12, ST01 AS CP13 " & _
                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER C1,CUSTOMER C2,ACC090,LAWOFFICESOURCE WHERE " & _
                        "HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP09=LOS06(+) AND SUBSTR(LOS04,1,5)=ST01(+) AND ST15=A0901(+) " & _
                        "AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
                        Replace(UCase(strExc(5)), "S1.", "") & "  and cp27 is null "
    strExc(0) = strExc(0) & " Union " & strExc(5) & " Order by CP12, CP13, B03 "
    'end 2021/09/01
    If RsTemp.State = adStateOpen Then RsTemp.Close
    RsTemp.Open strExc(0), cnnConnection
    If RsTemp.EOF And RsTemp.BOF Then
        MsgBox "資料庫內無資料 !", vbInformation
        m_print = 1
        Exit Sub
    End If
    i = 1
    If IsNull(RsTemp.Fields(0).Value) = False Then
        St = RsTemp.Fields(0).Value
    Else
        St = ""
    End If
    strSaleZone = "" & RsTemp.Fields(0).Value
    Page = 1
    CaseTitle TmpArea, Page
    CaseTitle1 strSaleZone, iPrint
'   iPrint = 2700
    With RsTemp
        Do While Not .EOF
            '若業務區不同
            If strSaleZone <> RsTemp.Fields(0).Value Then
                strSaleZone = RsTemp.Fields(0).Value
                '依業務區不同跳頁
                If Me.Text1(2).Text = "Y" Then
                   Printer.NewPage
                    Page = Page + 1
                    CaseTitle TmpArea, Page
                    CaseTitle1 strSaleZone, iPrint
                '不依業務區不同跳頁
                Else
                    iPrint = iPrint + 300
                    If iPrint > 8800 Then
                       Printer.NewPage
                       Page = Page + 1
                       CaseTitle TmpArea, Page
                       i = 0
                    End If
                    CaseTitle1 strSaleZone, iPrint
                End If
            End If
            If IsNull(.Fields(0).Value) = False Then
               St = .Fields(0).Value
            Else
               St = ""
            End If
            If iPrint > 9700 Then
               Printer.NewPage
               Page = Page + 1
               CaseTitle TmpArea, Page
               CaseTitle1 strSaleZone, iPrint
               i = 0
            End If
'            Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
'            If St <> TmpArea Then Printer.Print St
            '智權人員
            Printer.CurrentX = PLeft(0):    Printer.CurrentY = iPrint
            Printer.Print .Fields(1)
            '顧問案號
            Printer.CurrentX = PLeft(1):    Printer.CurrentY = iPrint
            '2011/8/8 modify by sonia
            'Printer.Print .Fields(2)
            If Left(Trim("" & .Fields(7)), 6) = "X65299" Then
               Printer.Print .Fields(2) & "（謝）"
            Else
               Printer.Print .Fields(2)
            End If
            '2011/8/8 end
            '顧問期間
            Printer.CurrentX = PLeft(2):    Printer.CurrentY = iPrint
            Printer.Print .Fields(3) & " - " & .Fields(4)
            '收文日期
            Printer.CurrentX = PLeft(3):    Printer.CurrentY = iPrint
            If .Fields(5) <> "//" Then
               St = .Fields(5)
            Else
               St = ""
            End If
            Printer.Print St
            '金額
            Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(6)
            'Modify By Sindy 2011/2/11 當事人1若為X65299時, 則當事人資料改抓當事人2
            If Left(Trim("" & .Fields(7)), 6) = "X65299" Then
               '客戶代號
               Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(15)
               '客戶名稱
               Printer.CurrentX = PLeft(6):    Printer.CurrentY = iPrint
               If IsNull(.Fields(16)) = False Then Printer.Print .Fields(16)
               '備註
               Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
               If IsNull(.Fields(18)) = False Then Printer.Print .Fields(18)
            '2011/2/11 End
            Else
               '客戶代號
               Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(7)
               '客戶名稱
               Printer.CurrentX = PLeft(6):    Printer.CurrentY = iPrint
               If IsNull(.Fields(8)) = False Then Printer.Print .Fields(8)
               '備註
               Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
               If IsNull(.Fields(10)) = False Then Printer.Print .Fields(10)
            End If
            If IsNull(RsTemp.Fields(0).Value) = False Then
               TmpArea = RsTemp.Fields(0).Value
            Else
               TmpArea = ""
            End If
            'Add By Cheng 2003/03/31
            '暫存列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "0", CheckStr(.Fields("HC01").Value), CheckStr(.Fields("HC02").Value), CheckStr(.Fields("HC03").Value), CheckStr(.Fields("HC04").Value)
            .MoveNext
            iPrint = iPrint + 300
        Loop
    End With
    Printer.EndDoc
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer
    i = 500
    If Page = "1" Then Printer.Orientation = 2
    Printer.Font.Size = 22
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = 6000:         Printer.CurrentY = i
    Printer.Print "顧問到期明細表"
    Printer.Font.Underline = False
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.CurrentX = 500:               Printer.CurrentY = i + 500
    Printer.Print "列印人 : " & strUserName
    Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
    Printer.Print "到期日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1))
    Printer.CurrentX = 13000:             Printer.CurrentY = i + 500
    Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
    Printer.CurrentX = 13000:             Printer.CurrentY = i + 800
    Printer.Print "頁次 : " & Page
    iPrint = i + 1100
End Sub

'Add By Cheng 2003/03/31
Private Sub CaseTitle1(ByVal Area As String, ByRef iPrint As Integer)
    Printer.CurrentX = 500:               Printer.CurrentY = iPrint
    Printer.Print "業務區：" & Area
    iPrint = iPrint + 300
    Printer.CurrentX = 500:               Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
    Printer.CurrentX = PLeft(0):          Printer.CurrentY = iPrint
    Printer.Print "智權人員"
    Printer.CurrentX = PLeft(1):          Printer.CurrentY = iPrint
    Printer.Print "顧問案號"
    Printer.CurrentX = PLeft(2):          Printer.CurrentY = iPrint
    Printer.Print "顧問期間"
    Printer.CurrentX = PLeft(3):          Printer.CurrentY = iPrint
    Printer.Print "收文日期"
    Printer.CurrentX = PLeft(4):          Printer.CurrentY = iPrint
    Printer.Print "金額"
    Printer.CurrentX = PLeft(5):          Printer.CurrentY = iPrint
    Printer.Print "客戶編號"
    Printer.CurrentX = PLeft(6):          Printer.CurrentY = iPrint
    Printer.Print "當事人名稱"
    Printer.CurrentX = PLeft(8):          Printer.CurrentY = iPrint
    Printer.Print "備　　註"
    iPrint = iPrint + 300
    Printer.CurrentX = 500:          Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1600
   PLeft(2) = 3200
   PLeft(3) = 5400
   PLeft(4) = 6900
   PLeft(5) = 7800
   PLeft(6) = 9300
   PLeft(7) = 10800
   PLeft(8) = 12500
   PLeft(9) = 14100
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'add by nickc 2005/04/01 刪除接洽結案單暫存資料，避免之前錯誤資料也印出來
   PUB_DeleteCaseCloseSheet strUserNum
   'Added by Lydia 2015/03/11 列印用
   PUB_SetPrinter Me.Name, Combo1, strPrint
   
   FormReset 'Added by Lydia 2015/03/13
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    'Modify By Cheng 2003/03/31
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 0, 1 '到期日期
        If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
        'Added by Lydia 2015/03/13 +是否寄發email
    Case 2, 5 '是否依業務區不同跳頁
        If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   
   Case 1 '到期日期
      'Add By Cheng 2002/09/09
      If blnClkSure = False Then
         If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
            If Val(Me.Text1(0).Text) > Val(Me.Text1(1).Text) Then
               MsgBox "到期日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(0).SetFocus
               Text1_GotFocus 0
               Exit Sub
            End If
         End If
      Else
         blnClkSure = False
      End If
    'Added by Lydia 2015/03/13 +是否寄發email
   Case 4
        If Me.Text1(3).Text <> "" And Me.Text1(4).Text <> "" Then
           If Me.Text1(3).Text > Me.Text1(4).Text Then
              MsgBox "業務範圍輸入錯誤!!!", vbExclamation + vbOKOnly
              Me.Text1(3).SetFocus
              Text1_GotFocus 3
              Exit Sub
           End If
        End If
   Case 5
         If Text1(Index) = "Y" Then
            GetEmailList
            cmdSend.SetFocus
         End If
   'end 2015/03/13
   End Select
   
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
            If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub

'Modified by Lydia 2021/09/01 +strLosCon
Private Function strGetcdnSQL(ByRef strLosCon As String) As String
 Dim i As Integer
 Dim strMid01 As String
 Dim strMid02 As String 'Added by Lydia 2021/09/01
 
   If Text1(0) = "" And Text1(1) <> "" Then
      strMid01 = " AND CP54<=" & Text1(1)
   ElseIf Text1(0) <> "" And Text1(1) <> "" Then
      strMid01 = " AND (CP54 BETWEEN " & ChangeTStringToWString(Text1(0)) & " AND " & ChangeTStringToWString(Text1(1)) & ")"
   End If
   strMid02 = strMid01 'Added by Lydia 2021/09/01
   
   'Added by Lydia 2015/03/ +業務區範圍 ,||''(排除運算,加速)
   If Text1(3).Text <> "" And Text1(4).Text <> "" Then
      strMid01 = strMid01 & " AND (CP12||''>='" & Text1(3).Text & "' AND CP12||''<='" & Text1(4).Text & "')"
      strMid02 = strMid02 & " AND (S1.ST15||''>='" & Text1(3).Text & "' AND S1.ST15||''<='" & Text1(4).Text & "')"   'Added by Lydia 2021/09/01
   ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
          strMid01 = strMid01 & " AND CP12||''<='" & Text1(4).Text & "'"
          strMid02 = strMid01 & " AND S1.ST15||''<='" & Text1(4).Text & "'"  'Added by Lydia 2021/09/01
   End If
   'Modified by Lydia 2021/09/01 +AND CP01='LA' AND SUBSTR(CP12,1,1)='S'
   strMid01 = strMid01 & " AND CP01='LA' AND SUBSTR(CP12,1,1)='S' AND CP10='0' AND CP57 IS NULL"
   strMid02 = strMid02 & " AND CP01='LA' AND SUBSTR(CP12,1,1)<>'S' AND CP10='0' AND CP57 IS NULL" 'Added by Lydia 2021/09/01
   strGetcdnSQL = strMid01
   strLosCon = strMid02  'Added by Lydia 2021/09/01
End Function
 'Added by Lydia 2015/03/11 列印用
Private Sub Form_Unload(Cancel As Integer)
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set frm073008 = Nothing
End Sub
'Added by Lydia 2015/03/13 + E-mail收件人Grid
Private Function GetEmailList() As String
  
  'Modified by Lydia 2021/09/01 區分案源和非案源
'  strExc(0) = strGetcdnSQL
'  strExc(0) = "SELECT ' ' a00,DECODE(CP12,A0901,A0902) A01,s2.st01 as A02,s2.st02 as A03,cp12 " & _
'              "FROM HIRECASE,CASEPROGRESS,STAFF s1,ACC090,staff s2 " & _
'              "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 " & _
'              "AND CP13=s1.ST01(+) and s2.st01(+)=a0908 AND CP12=A0901(+) " & strExc(0)
'  strExc(0) = strExc(0) & " and cp27 is null group by DECODE(CP12,A0901,A0902),s2.st01,s2.st02,cp12 order by cp12 "
  strExc(0) = strGetcdnSQL(strExc(5))
  '非案源
  strExc(0) = "SELECT ' ' a00,DECODE(CP12,A0901,A0902) A01,s2.st01 as A02,s2.st02 as A03,cp12 " & _
              "FROM HIRECASE,CASEPROGRESS,STAFF s1,ACC090,staff s2 " & _
              "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 " & _
              "AND CP13=s1.ST01(+) and s2.st01(+)=a0908 AND CP12=A0901(+) " & strExc(0)
  strExc(0) = strExc(0) & " and cp27 is null group by DECODE(CP12,A0901,A0902),s2.st01,s2.st02,cp12 "
  '案源
  strExc(5) = "SELECT ' ' a00,A0902 A01,s2.st01 as A02,s2.st02 as A03,S1.ST15 as cp12 " & _
              "FROM HIRECASE,CASEPROGRESS,STAFF s1,ACC090,staff s2,LawOfficeSource " & _
              "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 " & _
              "AND CP09=LOS06(+) AND substr(LOS04,1,5)=s1.ST01(+) and s2.st01(+)=a0908 AND S1.ST15=A0901(+) " & strExc(5)
  strExc(5) = strExc(5) & " and cp27 is null group by a0902,s2.st01,s2.st02,s1.st15 "
  strExc(0) = strExc(0) & " Union " & strExc(5) & " order by cp12 "
  'end 2021/09/01
  intI = 1
  If AdoEmail.State <> adStateClosed Then AdoEmail.Close
  Set AdoEmail = Nothing
  Set AdoEmail = ClsLawReadRstMsg(intI, strExc(0))
  
   DataGrid1.Enabled = True
   '+FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(AdoEmail, , , , Me.Name)
   If AdoEmail.RecordCount > 0 Then
   Else
      MsgBox "無符合資料！", vbExclamation
   End If
   
End Function

Private Function CheckCase() As Boolean
    CheckCase = False
    blnClkSure = False
    m_print = 0
    If ChkRange(Text1(0), Text1(1), "到期日期") = False Then
       blnClkSure = True
       Exit Function
    End If
    If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
       Me.Text1(0).SetFocus
       Text1_GotFocus 0
       Exit Function
    End If
    If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
       Me.Text1(1).SetFocus
       Text1_GotFocus 1
       Exit Function
    End If
    If Me.Text1(3).Text <> "" And Me.Text1(4).Text <> "" Then
       If Me.Text1(3).Text > Me.Text1(4).Text Then
          MsgBox "業務範圍輸入錯誤!!!", vbExclamation + vbOKOnly
          Me.Text1(3).SetFocus
          Text1_GotFocus 3
          Exit Function
       End If
    End If
    CheckCase = True
End Function
Private Sub cmdSend_Click()

    If CheckCase = False Then
       Exit Sub
    End If
    If Text1(5).Text <> "Y" Or AdoEmail.State = adStateClosed Then
       MsgBox "未確認E-mail收件人！", vbExclamation
       Text1(5).SetFocus
       Exit Sub
    End If
    Set AdoEmail = Adodc1.Recordset.Clone
    If AdoEmail.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        Set Adodc1.Recordset = PUB_CreateRecordset(AdoEmail, , , , Me.Name, mESeqNo)
        Call SendCase
        Screen.MousePointer = vbDefault
    Else
        MsgBox "無符合資料！", vbExclamation
    End If
    
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 1 Then
   pST01 = DataGrid1.Columns(ColIndex).Value
Else
   pST01 = ""
End If
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
   If ColIndex = 1 Then
    strSql = "select st02 from staff where st01='" & DataGrid1.Columns(ColIndex) & "' and st04=1 and st14 is null"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
       DataGrid1.Columns(ColIndex + 1) = "" & RsTemp.Fields(0)
    Else
       MsgBox "該員工編號不可設為收件人!", vbExclamation
       DataGrid1.Columns(ColIndex).Value = pST01
    End If
   End If
   Adodc1.Recordset.UPDATE
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If DataGrid1.col = 1 Then
      If KeyCode = vbKeyReturn Then
         If DataGrid1.row < Adodc1.Recordset.RecordCount - 1 Then
            SendKeys "{DOWN}"
         End If
      End If
   End If
End Sub

Private Sub DataGrid1_LostFocus()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Adodc1.Recordset.UPDATE
Checking:
   Exit Sub
End Sub
'清空表單
Private Sub FormReset()
   Dim oText As TextBox
   
   For Each oText In Text1
      oText.Text = ""
   Next
   
   If Not Adodc1.Recordset Is Nothing Then
      If Adodc1.Recordset.State = 1 Then
         Adodc1.Recordset.Close
         DataGrid1.Refresh
      End If
   End If
End Sub
'寄發e-mail資料
'注意:若有修改,請檢查frmAutoBatch.strMenu13 (每月自動通知個人)
Private Sub SendCase()
   Dim rsA As New ADODB.Recordset
   Dim stSQL As String, strPath As String
   Dim ff As Integer
   Dim TempFileName As String, strTemp(8) As String, strTFName As String
   Dim bolSend As Boolean
   Dim strTo As String, strApatch As String
   Dim stSQL2 As String 'Added by Lydia 2021/09/01
   
    strPath = PUB_Getdesktop
    strPath = strPath & "\顧問到期明細表\"
    
    If Dir(strPath, vbDirectory) = "" Then
       MkDir strPath
    End If
    
   strTFName = "顧問到期明細表" & Text1(0).Text & "-" & Text1(1).Text
   'Modified by Lydia 2021/09/01 區分案源和非案源
'   stSQL = strGetcdnSQL
'   stSQL = "SELECT s2.st01 as B00,DECODE(CP12,A0901,A0902) B01,DECODE(CP13,s1.ST01,s1.ST02) B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
'           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
'           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
'           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,CP13 " & _
'           "FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090,RDataFactory,STAFF s2 " & _
'           "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=s1.st01(+) AND CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) " & _
'           "AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
'           "and cp12=r005(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and s2.st01(+)=r003 and s2.st04=1 and s2.st14 is null " & stSQL
'   stSQL = stSQL & " and cp27 is null" & _
'           " ORDER BY 2,1,CP13,CP01||CP02||CP03||CP04 "
   stSQL = strGetcdnSQL(stSQL2)
   '非案源
   stSQL = "SELECT s2.st01 as B00,DECODE(CP12,A0901,A0902) B01,DECODE(CP13,s1.ST01,s1.ST02) B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,CP12,CP13 " & _
           "FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090,RDataFactory,STAFF s2 " & _
           "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=s1.st01(+) AND CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) " & _
           "AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
           "and cp12=r005(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and s2.st01(+)=r003 and s2.st04=1 and s2.st14 is null " & stSQL & _
           " and cp27 is null"
   '案源
   stSQL2 = "SELECT S2.ST01 AS B00,A0902 B01,S1.ST02 B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2, s1.ST15 as CP12, s1.st01 as CP13 " & _
           "FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090,RDataFactory,STAFF s2,LawOfficeSource " & _
           "WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP159=0 AND CP09=LOS06(+) AND SUBSTR(LOS04,1,5)=S1.ST01(+) AND S1.ST15=A0901(+) " & _
           "AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) " & _
           "and S1.ST15=r005(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and s2.st01(+)=r003 and s2.st04=1 and s2.st14 is null " & stSQL2 & _
           " and cp27 is null"
   stSQL = stSQL & " Union " & stSQL2 & " Order by CP12, CP13, B03 "
   'end 2021/09/01
   If rsA.State <> adStateClosed Then rsA.Close
  
    Set rsA = New ADODB.Recordset
    rsA.CursorLocation = adUseClient
    rsA.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    With rsA
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If TempFileName <> strTFName & "_" & Trim(.Fields("B01")) Then
                    strExc(10) = TempFileName: strApatch = ""
                    TempFileName = strTFName & "_" & Trim(.Fields("B01"))
                    If ff > 0 Then
                       Close #ff
                       If Len(Trim(strTemp(0))) > 0 Then
                          strTo = strTemp(0) '前一筆記錄的收信人
                          strApatch = strPath & strExc(10) & ".txt"  '前一個業務區附件
                          PUB_SendMail strUserNum, strTo, "", strExc(10), vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, , strApatch, True
                       End If
                    End If
                    ff = FreeFile
                    Open strPath & TempFileName & ".txt" For Output As ff
                    strExc(0) = convForm(" ", 30)
                    strExc(1) = convForm(" ", 90)
                    Print #ff, strExc(0) & TempFileName
                    Print #ff, "列印人：" & strUserName & strExc(1) & "列印日期：" & ChangeWStringToTString(strSrvDate(1))
                    Print #ff, "智權人員 顧問案號        顧問期間              收文日期        金額 客戶編號  當事人名稱                     備      註"
                    Print #ff, "======== =============== ===================== ========= ========== ========= ============================== =============================="
                End If
                strTemp(0) = "" & .Fields("B00")
                strTemp(1) = convForm("" & .Fields("B02"), 8)
                strTemp(2) = convForm("" & .Fields("B03"), 15)
                strTemp(3) = convForm(Trim("" & .Fields("B04")) & " - " & Trim("" & .Fields("B05")), 21)
                strTemp(4) = convForm(Trim("" & .Fields("B06")), 9)
                strTemp(5) = PUB_StrToStr(.Fields("cp16"), 10, True, True)
                strTemp(6) = convForm("" & .Fields("hc05"), 9)
                strTemp(7) = convForm(PUB_StrToStr(CheckStr("" & .Fields("B07")), 30), 30)
                strTemp(8) = convForm(PUB_StrToStr(CheckStr("" & .Fields("CU79_1")), 30), 30)
                
                If Left(Trim(.Fields("hc05")), 6) = "X65299" Then
                   strTemp(2) = convForm(PUB_StrToStr(Trim(strTemp(2)) & "（謝）", 15), 15) '顧問案號
                   '改用第２當事人
                   strTemp(6) = convForm("" & .Fields("hc24"), 9) '客戶代號
                   strTemp(7) = convForm(PUB_StrToStr(CheckStr("" & .Fields("B08")), 30), 30) '客戶名稱
                   strTemp(8) = convForm(PUB_StrToStr(CheckStr("" & .Fields("CU79_2")), 30), 30) '備註
                End If
            
                Print #ff, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & _
                      " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8)
                
                .MoveNext
            Loop
        'Added by Lydia 2015/04/08 Grid資料和Text條件不符
        Else
            If AdoEmail.RecordCount > 0 Then MsgBox "輸入條件與下方清單不一致!!!", vbExclamation
        End If
        If TempFileName <> "" Then
            Close ff
            strTo = strTemp(0) '最後一筆記錄的收信人
            strApatch = strPath & TempFileName & ".txt"
            strExc(10) = strPath & strExc(10) & ".txt"
             PUB_SendMail strUserNum, strTo, "", TempFileName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, , strApatch, True

        End If

    End With
    
    Set rsA = Nothing
End Sub

'end 'Added by Lydia 2015/03/13 + E-mail收件人Grid
