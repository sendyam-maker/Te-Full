VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100114_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人案件統計"
   ClientHeight    =   5736
   ClientLeft      =   1992
   ClientTop       =   1116
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9324
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel檔"
      Height          =   400
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   70
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7680
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin VB.TextBox txtCase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   7500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8448
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7524
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   70
      Width           =   900
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm100114_6.frx":0000
      Height          =   1305
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   2307
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "X1"
         Caption         =   "編號"
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
         DataField       =   "X2"
         Caption         =   "名稱"
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
         DataField       =   "FC04"
         Caption         =   "年度"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "X3"
         Caption         =   "期間"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "FC06"
         Caption         =   "系統別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FC07"
         Caption         =   "給案量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "FC08"
         Caption         =   "備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
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
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3360.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   468.283
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   455.811
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   612.284
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1560.189
         EndProperty
      EndProperty
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "統計方式："
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   4152
      TabIndex        =   23
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "統計方式："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   12
      Left            =   3216
      TabIndex        =   22
      Top             =   2400
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   4
      Left            =   5010
      TabIndex        =   21
      Top             =   400
      Width           =   2340
      VariousPropertyBits=   27
      Size            =   "4128;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   20
      Top             =   400
      Width           =   3180
      VariousPropertyBits=   27
      Size            =   "5609;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   19
      Top             =   699
      Width           =   8250
      VariousPropertyBits=   27
      Size            =   "14552;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   18
      Top             =   998
      Width           =   8250
      VariousPropertyBits=   27
      Size            =   "14552;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   17
      Top             =   1296
      Width           =   8256
      VariousPropertyBits=   27
      Size            =   "14563;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "互惠代理人資料："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   1440
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "3.案件如已更代或申請人已變動，則以系統類別*呈現此種狀況中其為原代理人或原申請人的案件數 "
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   3168
      Width           =   7824
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "2.近三年之定義為：以2018/11/27查詢為例，近三年之定義為2015/11/27~2018/11/27"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2904
      Width           =   6312
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "2.現代理人之定義：從申請時委辦本所或是後續程序將案件轉至本所的現代理人"
      Height          =   180
      Index           =   1
      Left            =   5808
      TabIndex        =   11
      Top             =   2352
      Visible         =   0   'False
      Width           =   6252
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "1.現代理人：系統類別+歷年案件數（括弧內為近三年案件數）"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   4905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件往來說明："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件往來："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   1704
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   180
      Index           =   4
      Left            =   4440
      TabIndex        =   7
      Top             =   400
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日文名稱："
      Height          =   180
      Index           =   3
      Left            =   30
      TabIndex        =   6
      Top             =   1296
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文名稱："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   5
      Top             =   998
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "中文名稱："
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   699
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   400
      Width           =   1080
   End
End
Attribute VB_Name = "frm100114_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/06 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lbl1(index)
'Create by sonia 2018/10/31
Option Explicit
'紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2018/12/25
Dim strMid01 As String, strMid02 As String
Dim bolExcel As Boolean
Dim intXlsCnt As Integer 'Added by Lydia 2019/05/06
Dim m_strKind As String 'Added by Lydia 2025/09/19

Public Sub PubShowNextData()
   Select Case cmdState
   Case 0
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 1
      fnCloseAllFrm100
   Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdState = -1
   
   'Added by Lydia 2018/12/25
   'Mark by Lydia 2024/04/22
   'If Pub_StrUserSt03 = "M51" Then
   '    cmdExcel.Visible = True
   'End If
   ''end 2018/12/25
   'end 2024/04/22
   
   'Added b y Lydia 2025/09/19
   Debug.Print "Load"
   m_strKind = frm100114_1.m_strTotKind
   If m_strKind = "1" Then
      'Modified by Lydia 2025/11/11 改選項說明:新申請案=>新案（委任申請案）
      lblKind = "1-新案（委任申請案）"
      lblMemo(1) = "2.現代理人之定義：從申請時委辦本所"
      lblMemo(3) = ""
   Else
      'Modified by Lydia 2025/11/11 改選項說明:案件數
      lblKind = "2-在案（目前代理案）"
      lblMemo(1) = "2.現代理人之定義：從申請時委辦本所或是後續程序將案件轉至本所的現代理人"
      'Modified by Lydia 2025/11/11 原本的lblMemo(1)隱藏; 4.=>3.
      lblMemo(3) = "3.案件如已更代或申請人已變動，則以系統類別*呈現此種狀況中其為原代理人或原申請人的案件數"
   End If
   'end 2025/09/19
End Sub

Sub StrMenu()
Dim strMidCon As String 'Added by Lydia 2019/05/06
Dim strSqlNow As String, strSQLpass As String 'Added by Lydia 2024/04/22
Dim strSqlAreaNow As String, strSqlAreaPass As String 'Added by Lydia 2024/05/24
   
   Debug.Print "Run"
   'Memo by Lydia 2025/08/15
   '*******************
      '整理規則如下:
      '
      '案件往來說明:
      '1.現代理人：系統類別+歷年案件數（括弧內為近三年案件數）
      '2.現代理人之定義：從申請時委辦本所或是後續程序將案件轉至本所的現代理人
      '3.近三年之定義為：以2018/11/27查詢為例，近三年之定義為2015/11/27~2018/11/27=>以系統當日起往前推3年的案件數
      '4.案件如已更代或申請人已變動，則以系統類別*呈現此種狀況中其為原代理人或原申請人的案件數（括弧內為近三年案件數）
      '
      '案件數的定義:
      '1.A類或B類收文;
      '2.排除假收文和假發文，或是發文日=閉卷日的收文
      '3.排除特定案件性質如下：回覆代理人、不續辦、取消收文、更換FC代理人、閉卷、催審、催公開、催提申、催收達、催款。
      'SELECT cpm01||cpm02 as t01,cpm03,sk01,sk03 FROM casepropertymap,systemkind WHERE cpm01 IN ('P','CFP','FCP','T','CFT','FCT','TF')
      'AND instr('回覆代理人、不續辦、取消收文、更換FC代理人、閉卷、催審、催公開、催提申、催收達、催款。',cpm03) > 0 and cpm01=sk01(+)
      'ORDER BY sk02,sk03,cpm01,cpm02;

      '系統別:
      'FCFP、FCFT：抓收文的FC代理人(CP139)，同時非專利／商標基本檔的FC代理人(PA75/TM44)
      'FMP: P案非台灣案並且有設定FC代理人 (pa75)
   '*******************
   
'Added by Lydia 2019/05/06 CF案件判斷案件日期為最小發文日;並且針對1.現在案件屬於A, 2.最初案件屬於A, 3.中間案件屬於A的情況都要抓到,所以先將資料丟暫存檔
'CREATE TABLE R100114_6 (ID VARCHAR2(6),FORMID VARCHAR2(20),PNO VARCHAR2(8),
'C01 VARCHAR2(3),C02 VARCHAR2(6),C03 VARCHAR2(1),C04 VARCHAR2(2),MINCP09 VARCHAR2(30),MAXCP09 VARCHAR2(30));
If bolExcel = False Or (bolExcel = True And intXlsCnt = 0) Then
    '清除暫存檔
    'Modified by Lydia 2024/04/22 改成模組
    'cnnConnection.Execute "DELETE FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID = '" & Me.Name & "' "
    'Modified by Lydia 2024/05/24
    'Modified by Lydia 2025/09/19 +選擇統計方式 m_strKind
    Call Pub_frm100114_6_StrMenu(strUserNum, Me.Name, Me.Tag, strMidCon, strSqlNow, strSQLpass, strSqlAreaNow, strSqlAreaPass, m_strKind)
End If
'strMidCon = "AND CP44 IS NOT NULL AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726')" 'Mark by Lydia 2024/04/22 改成模組
strSql = "INSERT INTO R100114_6 (ID,FORMID,PNO,C01,C02,C03,C04,MINCP09,MAXCP09) " & _
            "SELECT '" & strUserNum & "','" & Me.Name & "','" & Left(GetNewFagent(Me.Tag), 8) & "',CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
            "   WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 " & _
            "   AND CP158<>NVL(NVL(TM30,PA58),0) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
            "   AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP44||' '<>PA75||TM44||' ') " & strMidCon & _
            "   GROUP BY CP01,CP02,CP03,CP04"
cnnConnection.Execute strSql, intI

ClearQueryLog (Me.Name) 'Added by Lydia 2025/08/22 清除查詢印表記錄檔欄位
If bolExcel = False Then 'Added by Lydia 2018/12/25 判斷是否抓案件往來的語法
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   '顯示表單上頭資料
   lbl1(0).Caption = Me.Tag
   
   '檢查國內外權限
   If CheckSR12(Me.Tag) = False Then
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   
   strSql = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA77,NA03 FROM FAGENT,NATION WHERE FA01='" & Left(GetNewFagent(Me.Tag), 8) & "' AND FA02='" & Right(GetNewFagent(Me.Tag), 1) & "' AND FA10=NA01(+) "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       If Not IsNull(adoRecordset.Fields(0)) Then
           'Modified by Lydia 2018/11/28 "&"在畫面顯示為"_"
           lbl1(1).Caption = Replace("" & adoRecordset.Fields(0), "&", "＆")
       Else
           lbl1(1).Caption = ""
       End If
       If Not IsNull(adoRecordset.Fields(1)) Then
           'Modified by Lydia 2018/11/28 "&"在畫面顯示為"_"
           lbl1(2).Caption = Replace("" & adoRecordset.Fields(1), "&", "＆")
       Else
           lbl1(2) = ""
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           'Modified by Lydia 2018/11/28 "&"在畫面顯示為"_"
           lbl1(3) = Replace("" & adoRecordset.Fields(2), "&", "＆")
       Else
           lbl1(3) = ""
       End If
       If CheckStr(adoRecordset.Fields("FA77")) = "Y" Then
           lbl1(0).ForeColor = &HFF&
       Else
           lbl1(0).ForeColor = &H80000012
       End If
       lbl1(4) = "" & adoRecordset.Fields("NA03")
   Else
       lbl1(1).Caption = ""
       lbl1(2).Caption = ""
       lbl1(3).Caption = ""
       lbl1(4).Caption = ""
   End If

   '開始搜尋
   txtCase = ""
End If 'Added by Lydia 2018/12/25

'------------------現在案件屬於該編號
' 'Mark by Lydia 2024/04/22 改成模組
'   '代理人
'   If Left(GetNewFagent(Me.Tag), 1) = "Y" Then
'      '現代理人
'      strSql = "SELECT CP01,COUNT(*)||'('||SUM(NEW)||')' FROM ("
'      '專利FC代理人  FCP-055795應抓105/11/16案件轉至本所,不可抓11/11/11分割
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP05,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'
'      '商標FC代理人
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP05,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'
'      '專利及商標CF代理人(子案不算)
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP27,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN (SELECT SUBSTR(MAXCP09,9,9) FROM " & _
'                        "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        "(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                        "AND CP44||' '<>PA75||TM44||' ') " & _
'                        "AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "GROUP BY CP01,CP02,CP03,CP04)) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%') "
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN (SELECT SUBSTR(MAXCP09,9,9) FROM " & _
'                        "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        "(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                        "AND CP44||' '<>PA75||TM44||' ') " & _
'                        "AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "GROUP BY CP01,CP02,CP03,CP04)) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%') "
'      'Modified by Lydia 2019/05/06 現在案件屬於Y編號
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN (SELECT SUBSTR(MAXCP09,9,9) FROM " & _
'                        "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        "(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(NVL(TM30,PA58),0) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                        "AND CP44||' '<>PA75||TM44||' ') " & _
'                        "AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "GROUP BY CP01,CP02,CP03,CP04)) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%') "
'         strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
'                        "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left(GetNewFagent(Me.Tag), 8) & "') > 0 AND PNO='" & Left(GetNewFagent(Me.Tag), 8) & "' ) " & _
'                        "AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' " & strMidCon & _
'                        " GROUP BY CP01,CP02,CP03,CP04)) "
'   '申請人
'   Else
'      'Added by Lydia 2023/06/07 Owen製作下周日本關西地區拜訪客戶之排程表,其中有申請人,暫時參考代理人的寫法
'      strSql = "SELECT CP01,COUNT(*)||'('||SUM(NEW)||')' FROM ("
'      '專利
'      strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE INSTR(PA26||PA27||PA28||PA29||PA30,'" & Left(GetNewFagent(Me.Tag), 8) & "') > 0 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      '商標
'      strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        "(SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE INSTR(TM23||TM78||TM79||TM80||TM81,'" & Left(GetNewFagent(Me.Tag), 8) & "') > 0 AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C'" & _
'                        "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      'end 2023/06/07
'   End If
'   'Modified by Lydia 2019/05/06
'   'strSql = strSql & "GROUP BY CP01 ORDER BY 1 "
'   strSql = strSql & " ) GROUP BY CP01 ORDER BY 1 "
'end 2024/04/22
'Added by Lydia 2018/12/25 判斷是否抓案件往來的語法
If bolExcel = True Then
    'Modified by Lydia 2024/04/22 改成模組
    'strMid01 = strSql
    strMid01 = strSqlNow
Else
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2024/04/22 改成模組
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSqlNow, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         txtCase = txtCase & adoRecordset(0) & "-" & adoRecordset(1) & ";"
         adoRecordset.MoveNext
      Loop
   End If
   CheckOC
End If 'Added by Lydia 2018/12/25

'------------------曾經案件屬於該編號
'Mark by Lydia 2024/04/22 改成模組
'   '代理人
'   If Left(GetNewFagent(Me.Tag), 1) = "Y" Then
'      '原代理人
'      strSql = "SELECT CP01||'*',COUNT(*)||'('||SUM(NEW)||')' FROM ("
'      '專利FC代理人
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP05,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL AND CP09<'C' " & _
'                        "    AND SUBSTR(CP139,1,8)<>SUBSTR(PA75,1,8) AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL AND CP09<'C' " & _
'                        "    AND SUBSTR(CP139,1,8)<>SUBSTR(PA75,1,8) AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL AND CP09<'C' " & _
'                        "    AND SUBSTR(CP139,1,8)<>SUBSTR(PA75,1,8) AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'
'      '商標FC代理人
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP05,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL AND CP09<'C' " & _
'                        "     AND SUBSTR(CP139,1,8)<>SUBSTR(TM44,1,8) AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL AND CP09<'C' " & _
'                        "     AND SUBSTR(CP139,1,8)<>SUBSTR(TM44,1,8) AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE CP139 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL AND CP09<'C' " & _
'                        "     AND SUBSTR(CP139,1,8)<>SUBSTR(TM44,1,8) AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'
'      '專利及商標CF代理人(子案不算)
'      'Modified by Lydia 2018/11/27 改為以系統當日起往前推3年的案件數(ex.2018/11/27~2015/11/27)
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SUBSTR(CP27,1,4),TO_CHAR(SYSDATE,'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-12),'YYYY'),1,TO_CHAR(ADD_MONTHS(SYSDATE,-24),'YYYY'),1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MINCP09,9,9) FROM " & _
'                        " (SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        " 　(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "      AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP44||' '<>PA75||TM44||' ') " & _
'                        "  AND CP158>19221111 AND CP09<'C' " & _
'                        "  AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "  GROUP BY CP01,CP02,CP03,CP04) WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18) " & _
'                        ") AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%')"
'      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MINCP09,9,9) FROM " & _
'                        " (SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        " 　(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "      AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP44||' '<>PA75||TM44||' ') " & _
'                        "  AND CP158>19221111 AND CP09<'C' " & _
'                        "  AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "  GROUP BY CP01,CP02,CP03,CP04) WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18) " & _
'                        ") AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%')"
'      'Modified by Lydia 2019/05/06 現在案件不屬於Y編號
'      'strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MINCP09,9,9) FROM " & _
'                        " (SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        " 　(SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(NVL(TM30,PA58),0) AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' AND CP04='00' AND CP09<'C' " & _
'                        "      AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP44||' '<>PA75||TM44||' ') " & _
'                        "  AND CP158>19221111 AND CP09<'C' " & _
'                        "  AND CP01||CP10 NOT IN ('P902','P907','P913','P925','P937','CFP902','CFP907','CFP913','CFP925','CFP937','T703','T704','T718','T720','T726','TF703','TF704','TF718','TF720','TF726','CFT703','CFT704','CFT718','CFT720','CFT726') " & _
'                        "  GROUP BY CP01,CP02,CP03,CP04) WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18) " & _
'                        ") AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%')"
'      strSql = strSql & "Union SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP27)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS WHERE CP09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
'                        "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
'                        "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left(GetNewFagent(Me.Tag), 8) & "') = 0 AND PNO='" & Left(GetNewFagent(Me.Tag), 8) & "' ) " & _
'                        "AND CP44 LIKE '" & Left(GetNewFagent(Me.Tag), 8) & "%' " & strMidCon & _
'                        " GROUP BY CP01,CP02,CP03,CP04)) "
'   '申請人
'   Else
'      'Added by Lydia 2023/06/07 抓曾經是的案件
'      strSql = "SELECT CP01||'*',COUNT(*)||'('||SUM(NEW)||')' FROM ("
'      '專利
'      strSql = strSql & "SELECT DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,PATENT WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,PATENT WHERE INSTR(CP55||CP93||CP94||CP95||CP96,'" & Left(GetNewFagent(Me.Tag), 8) & "') > 0 AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL AND CP09<'C' " & _
'                        " AND INSTR(PA26||PA27||PA28||PA29||PA30,'" & Left(GetNewFagent(Me.Tag), 8) & "') = 0 AND SUBSTR(CP139,1,8)<>SUBSTR(PA75,1,8) AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 "
'      '商標
'      strSql = strSql & "Union SELECT DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01) CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP27,CP57,DECODE(SIGN((TO_CHAR(SYSDATE,'YYYYMMDD') - CP05)-30001),-1,1,0) NEW " & _
'                        "FROM CASEPROGRESS,TRADEMARK WHERE CP09 IN " & _
'                        " (SELECT SUBSTR(MIN(CP05||CP09),9) FROM CASEPROGRESS,TRADEMARK WHERE INSTR(CP55||CP93||CP94||CP95||CP96,'" & Left(GetNewFagent(Me.Tag), 8) & "') > 0 AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL AND CP09<'C' " & _
'                        " AND INSTR(TM23||TM78||TM79||TM80||TM81,'" & Left(GetNewFagent(Me.Tag), 8) & "') = 0 AND SUBSTR(CP139,1,8)<>SUBSTR(TM44,1,8) AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) GROUP BY CP01,CP02,CP03,CP04) AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 "
'      'end 2023/06/07
'   End If
'   'Modified by Lydia 2019/05/06
'   'strSql = strSql & "GROUP BY CP01 ORDER BY 1 "
'   strSql = strSql & " ) GROUP BY CP01 ORDER BY 1 "
'end 2024/04/22
         
'Added by Lydia 2018/12/25 判斷是否抓案件往來的語法
If m_strKind = "2" Then  'Added by Lydia 2025/09/19 選擇統計方式：2-案件數，才需要分析過去案件
   If bolExcel = True Then
       'Modified by Lydia 2024/04/22 改成模組
       'strMid02 = strSql
       strMid02 = strSQLpass
   Else
      CheckOC
      adoRecordset.CursorLocation = adUseClient
      'Modified by Lydia 2024/04/22 改成模組
      'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      adoRecordset.Open strSQLpass, cnnConnection, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
         '現在資料與原資料間以三個空格區隔
         txtCase = txtCase & "　　　"
         adoRecordset.MoveFirst
         Do While Not adoRecordset.EOF
            txtCase = txtCase & adoRecordset(0) & "-" & adoRecordset(1) & ";"
            adoRecordset.MoveNext
         Loop
      End If
      
      If txtCase = "" Then txtCase = "無案件！"
      'Added by Lydia 2025/08/22 記錄客戶/代理人編號(案件統計)及案件往來欄的內容即可，不必串接前一畫面的條件；結果筆數欄記錄互惠代理人資料Grid的筆數。
      pub_QL05 = pub_QL05 & ";" & lbl1(0) & "(案件統計)：" & IIf(txtCase = "", "無案件", txtCase)
      'end 2025/08/22
   End If 'Added by Lydia 2018/12/25
End If 'Added by Lydia 2025/09/19 選擇統計方式：2-案件數，才需要分析過去案件

If bolExcel = False Then 'Added by Lydia 2018/12/25 判斷是否抓案件往來的語法
   'Added by Lydia 2018/11/28 增加"互惠代理人資料"顯示
   strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
                    ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
                    " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01='" & Left(GetNewFagent(Me.Tag), 8) & "'" & _
                    " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
                    " and na01(+)=fa10"
   strExc(0) = "SELECT X.*,FC01||FC02||DECODE(FC03,NULL,'','-'||FC03) X1" & _
      ",DECODE(FC03,NULL,NVL(EN,NVL(JN,CN)),NVL(CEN,NVL(CJN,CCN))) X2" & _
      ",DECODE(FC05,'1','上半','下半') X3" & _
      " FROM (" & strExc(0) & ") X "
   strExc(0) = strExc(0) & "ORDER BY fc01,fc02,fc03,fc04 desc,fc05,fc06"

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Added by Lydia 2025/08/22
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount)
   Else
      InsertQueryLog (0)
   End If
   'end 2025/08/22
   Set Adodc1.Recordset = RsTemp
   'end 2018/11/28
   
   CheckOC
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End If 'Added by Lydia 2018/12/25

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100114_6 = Nothing
End Sub

'Added by Lydia 2018/12/25 列出所有國外部Y編號代理人下列資料，以利外專/外商評估2019年APAA邀請人數
Private Sub CmdExcel_Click()
'Mark by Lydia 2018/12/28 因為有些專案未載入excel,先Mark
MsgBox "因為有些專案未載入excel,先Mark"

Call Pub_ChkExcelPath 'Added by Lydia 2021/07/01 檢查xls資料夾的模組
   
''Memo by Lydia 2019/05/06 intXlsCnt控制清除暫存檔
'
'Dim strCon1 As String
'Dim intJ As Integer
'Dim rsRd As New ADODB.Recordset
'Dim xlsPoint1 As New Excel.Application
'Dim wksPoint1 As New Worksheet
'Dim strFileName As String '檔案名稱
'Dim iRow As Integer
'Dim xCols As Integer
'Dim tmpArr1 As Variant, tmpArr2 As Variant
'Dim strTmp As String
'
''抓所有Y 編號、國籍、英文名稱、
''性質 (a / B / c): a: 代理人律師事務所 B: 公司直接委辦 c: 其他
''最近新案[含分割/改請]收文日期:抓專利性質NewCasePtyList+3開頭, 商標案101
''最近其它收文日期: 所有A類收文
''互惠 (p / t): 當年有互惠記錄(107上半)
''FCP歷年件數(近3年件數: 系統日往前推3年)、FCT(同左)、FCL(同左)、CFP(同左)、CFT(同左)、CFL(同左)、其它案件(同左)
'
'    cmdExcel.Enabled = False
'    Screen.MousePointer = vbHourglass
'    bolExcel = True
'    strTmp = Me.Tag '記錄傳入Form的X/Y編號
'    Debug.Print "Start: " & Format(ServerTime, "000000")
'
'
'    strExc(1) = "select decode(substr(na02,1,1),'A','1亞洲','B','1亞洲',decode(substr(na02,1,2),'C0','1亞洲','C1','2美洲','C2','3歐洲','C3','4非洲','5大洋洲')) area, " & _
'                      "fa01||fa02 as fano, substr(na01,1,3) na01,na03,decode(fa05,null,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65) faname,decode(fa76,'A','代理人律師事務所','B','公司直接委辦','其他') ftype " & _
'                      "From fagent, nation where fa02='0' and fa01>='Y0000100' and fa10=na01(+) "
'    'strExc(1) = strExc(1) & "and fa10 ='013' " '測試抓特定國家
'    'Added by Lydia 2023/06/07 增加申請人
'    strExc(1) = strExc(1) & "and fa01 in ('Y5251900','Y5427100','Y5520100','Y5285800','Y3015000','Y2776600','Y2204600','Y5488800','Y5189000','Y5150800') "
'    strExc(1) = strExc(1) & " Union select decode(substr(na02,1,1),'A','1亞洲','B','1亞洲',decode(substr(na02,1,2),'C0','1亞洲','C1','2美洲','C2','3歐洲','C3','4非洲','5大洋洲')) area," & _
'                       "cu01||cu02 as cuno, substr(na01,1,3) na01,na03,decode(cu05,null,nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90) cuname, decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構') ftype " & _
'                       "From customer, nation where cu02='0' and cu10=na01(+) and cu01='X8368500' "
'    'end 2023/06/07
'    strExc(1) = strExc(1) & "order by area,na01,faname,fano "
'    intI = 1
'    Set rsRd = ClsLawReadRstMsg(intI, strExc(1))
'    If intI = 1 Then
'        strFileName = strExcelPath & strSrvDate(1) & "代理人案件統計分析表" & MsgText(43)
'        If Dir(strFileName) <> "" Then
'           Kill strFileName
'        End If
'        xlsPoint1.Workbooks.add
'        xlsPoint1.Visible = False '預設不顯示
'        rsRd.MoveFirst
'        iRow = 1
'        xCols = 1
'        xlsPoint1.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
'        Set wksPoint1 = xlsPoint1.Worksheets(1)
'        xlsPoint1.Sheets(1).Select '選擇工作表
'        strExc(1) = "洲別,代理人編號,國籍,國籍名稱,代理人名稱,性質,最近專利/商標之新案[含分割/改請]收文日期, 最近其它收文日期,互惠(P),互惠(T),案件往來"
'        strExc(2) = "9,10,8,10,25,12,20,10,7,7,50"
'        tmpArr1 = Split(strExc(1), ",")
'        tmpArr2 = Split(strExc(2), ",")
'        '欄位抬頭
'        For intJ = 0 To UBound(tmpArr1)
'            If Trim(tmpArr1(intJ)) <> "" Then
'                 strExc(3) = Pub_NumberToSystem26(intJ + 1)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = Trim(tmpArr1(intJ))
'                 xlsPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = Val("" & tmpArr2(intJ))
'                 xlsPoint1.Range(strExc(3) & ":" & strExc(3)).NumberFormatLocal = "@" '文字格式
'                 If InStr("G,H,I,J", strExc(3)) > 0 Then
'                    xlsPoint1.Range(strExc(3) & ":" & strExc(3)).HorizontalAlignment = xlCenter '水平置中
'                 End If
'            End If
'        Next intJ
'        xlsPoint1.Range(iRow & ":" & iRow).RowHeight = 38
'        xlsPoint1.Range(iRow & ":" & iRow).WrapText = True '自動換列
'        xlsPoint1.Range(iRow & ":" & iRow).VerticalAlignment = xlCenter  '垂直置中
'        xlsPoint1.Range(iRow & ":" & iRow).HorizontalAlignment = xlCenter '水平置中
'        iRow = iRow + 1
'        wksPoint1.Range(iRow & ":" & iRow).Select
'        xlsPoint1.ActiveWindow.FreezePanes = True '凍結窗格
'        wksPoint1.Range("A1").Select
'
'        Do While Not rsRd.EOF
'            xCols = 0
'            strCon1 = ""
'            '代理人基本資料
'            For intJ = 0 To rsRd.Fields.Count - 1
'                 xCols = xCols + 1
'                 strExc(3) = Pub_NumberToSystem26(xCols)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = "" & rsRd.Fields(intJ)
'            Next intJ
'             xCols = xCols + 1
'
'            '最近新案[含分割/改請]收文日期
'            strSql = "SELECT SQLDATET(MAX(NDATE)) FROM ("
'            '專利案:抓專利性質NewCasePtyList+3開頭
'            'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'            'strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C' " & _
'                        "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) AND (CP10 IN (" & GetAddStr(NewCasePtyList) & ") OR CP10 LIKE '3%') "
'            If Left("" & rsRd.Fields("fano"), 1) = "Y" Then  'Added by Lydia 2023/06/07 判斷
'               strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) AND (CP10 IN (" & GetAddStr(NewCasePtyList) & ") OR CP10 LIKE '3%') "
'
'               '商標案101
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C' " & _
'                           "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) AND CP10 ='101' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) AND CP10 ='101' "
'
'
'               '專利及商標CF代理人(子案不算)
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND (CP10 IN (" & GetAddStr(NewCasePtyList) & ") OR CP10 LIKE '3%') "
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 ='101' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP158<>NVL(PA58,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND (CP10 IN (" & GetAddStr(NewCasePtyList) & ") OR CP10 LIKE '3%') "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(TM30,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 ='101' "
'               strSql = strSql & ") "
'            'Added by Lydia 2023/06/07
'            Else '申請人
'               '專利
'               strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE INSTR(PA26||PA27||PA28||PA29||PA30,'" & Mid("" & rsRd.Fields("fano"), 1, 8) & "') > 0 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09<'C' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) AND (CP10 IN (" & GetAddStr(NewCasePtyList) & ") OR CP10 LIKE '3%') "
'               '商標案101
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE INSTR(TM23||TM78||TM79||TM80||TM81,'" & Mid("" & rsRd.Fields("fano"), 1, 8) & "') > 0 AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09<'C' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) AND CP10 ='101' "
'               strSql = strSql & ") "
'            'end 2023/06/07
'            End If
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'                 strExc(3) = Pub_NumberToSystem26(xCols)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = "" & RsTemp.Fields(0)
'            End If
'            xCols = xCols + 1
'
'            If Left("" & rsRd.Fields("fano"), 1) = "Y" Then  'Added by Lydia 2023/06/07 判斷
'               '最近其它收文日期: 所有A類收文(排除新案)
'               strSql = "SELECT SQLDATET(MAX(NDATE)) FROM ("
'               '專利案:抓專利性質NewCasePtyList+3開頭
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'               strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE PA75 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'
'               '商標案101
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND (CP158>19221111 OR (CP158=0 AND CP159=0)) AND CP10 <>'101' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE TM44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) AND CP10 <>'101' "
'
'               '專利及商標CF代理人(子案不算)
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 <>'101' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP158<>NVL(PA58,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(TM30,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 <>'101' "
'
'               '服務和商標
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MIN(CP05||CP09),1,8) FROM CASEPROGRESS,SERVICEPRACTICE " & _
'                           "WHERE SP26 like '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP09 LIKE 'A%' AND CP05 > 19221111 AND (CP158 > 19221111 OR (CP158=0 AND CP159=0)) "
'               'strSql = strSql & "Union SELECT SUBSTR(MIN(CP05||CP09),1,8) FROM CASEPROGRESS,LAWCASE " & _
'                           "WHERE LC27 like '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP09 LIKE 'A%' AND CP05 > 19221111 AND (CP158 > 19221111 OR (CP158=0 AND CP159=0)) "
'               strSql = strSql & "Union SELECT SUBSTR(MIN(CP05||CP09),1,8) FROM CASEPROGRESS,SERVICEPRACTICE " & _
'                           "WHERE SP26 like '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP09 LIKE 'A%' AND CP05 > 19221111 AND ((CP158>19221111 AND CP158<>NVL(SP16,0)) OR (CP158=0 AND CP159=0)) "
'               strSql = strSql & "Union SELECT SUBSTR(MIN(CP05||CP09),1,8) FROM CASEPROGRESS,LAWCASE " & _
'                           "WHERE LC27 like '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP09 LIKE 'A%' AND CP05 > 19221111 AND ((CP158>19221111 AND CP158<>NVL(LC09,0)) OR (CP158=0 AND CP159=0)) "
'
'               '服務和商標(CF代理人)
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
'                           "AND CP44||' '<>SP26||' ' "
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,LAWCASE WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & _
'                           "AND CP44||' '<>LC27||' ' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP158>19221111 AND CP158<>NVL(SP16,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
'                           "AND CP44||' '<>SP26||' ' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,LAWCASE WHERE CP158>19221111 AND CP158<>NVL(LC09,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & _
'                           "AND CP44||' '<>LC27||' ' "
'               strSql = strSql & ") "
'            'Added by Lydia 2023/06/07
'            Else  '申請人
'               '最近其它收文日期: 所有A類收文(排除新案)
'               strSql = "SELECT SQLDATET(MAX(NDATE)) FROM ("
'               '專利
'               strSql = strSql & "SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE INSTR(PA26||PA27||PA28||PA29||PA30,'" & Mid("" & rsRd.Fields("fano"), 1, 8) & "') > 0 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'
'               '商標案101
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP05||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE INSTR(TM23||TM78||TM79||TM80||TM81,'" & Mid("" & rsRd.Fields("fano"), 1, 8) & "') > 0 AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP09 LIKE 'A%' " & _
'                           "AND CP05>19221111 AND ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) AND CP10 <>'101' "
'
'               '專利及商標CF代理人(子案不算)
'               'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'               'strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 <>'101' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP158<>NVL(PA58,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
'                           "AND CP44||' '<>PA75||' ' AND CP10 NOT IN (" & GetAddStr(NewCasePtyList) & ") AND CP10 NOT LIKE '3%' "
'               strSql = strSql & "Union SELECT SUBSTR(MAX(CP27||CP09),1,8) NDATE FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(TM30,0) AND CP44 LIKE '" & Mid("" & rsRd.Fields("fano"), 1, 8) & "%' AND CP04='00' AND CP09 LIKE 'A%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                           "AND CP44||' '<>TM44||' ' AND CP10 <>'101' "
'               strSql = strSql & ") "
'
'            'end 2023/06/07
'            End If
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'                 strExc(3) = Pub_NumberToSystem26(xCols)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = "" & RsTemp.Fields(0)
'            End If
'            xCols = xCols + 1
'
'            '互惠 (P / T): 當年有互惠記錄(107上半)
'            strSql = "Select Count(*) cnt from FagentConfig  WHERE FC01||FC02='" & rsRd.Fields("fano") & "' and FC04='" & Left(strSrvDate(2), 3) & "' and FC05='1'  and FC06='CFP' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'                 strExc(3) = Pub_NumberToSystem26(xCols)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = IIf(Val("" & RsTemp.Fields(0)) > 0, "Y", "")
'            End If
'            xCols = xCols + 1
'            strSql = "Select Count(*) cnt from FagentConfig  WHERE FC01||FC02='" & rsRd.Fields("fano") & "' and FC04='" & Left(strSrvDate(2), 3) & "' and FC05='1'  and FC06='CFT' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'                 strExc(3) = Pub_NumberToSystem26(xCols)
'                 xlsPoint1.Range(strExc(3) & iRow).Value = IIf(Val("" & RsTemp.Fields(0)) > 0, "Y", "")
'            End If
'            xCols = xCols + 1
'
'            '案件往來
'            strMid01 = "": strMid02 = ""
'            Me.Tag = "" & rsRd.Fields("fano")
'            Call StrMenu
'            If strMid01 <> "" Then
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strMid01)
'                If intI = 1 Then
'                     RsTemp.MoveFirst
'                     Do While Not RsTemp.EOF
'                          strCon1 = strCon1 & RsTemp(0) & "-" & RsTemp(1) & ";"
'                          RsTemp.MoveNext
'                     Loop
'                End If
'            End If
'            If strMid02 <> "" Then
'                '現在資料與原資料間以三個空格區隔
'                If strCon1 <> "" Then strCon1 = strCon1 & "　　　"
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strMid02)
'                If intI = 1 Then
'                     RsTemp.MoveFirst
'                     Do While Not RsTemp.EOF
'                          strCon1 = strCon1 & RsTemp(0) & "-" & RsTemp(1) & ";"
'                          RsTemp.MoveNext
'                     Loop
'                End If
'            End If
'            strExc(3) = Pub_NumberToSystem26(xCols)
'            xlsPoint1.Range(strExc(3) & iRow).Value = IIf(strCon1 <> "", strCon1, "無案件！")
'            xCols = xCols + 1
'
'            iRow = iRow + 1
'            rsRd.MoveNext
'        Loop
'
'       '判斷版本
'       If Val(xlsPoint1.Version) < 12 Then
'            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'       Else
'            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'       End If
'
'       xlsPoint1.Workbooks.Close
'       xlsPoint1.Quit
'       Set wksPoint1 = Nothing
'       Set xlsPoint1 = Nothing
'        'Modify by Amy 2021/06/22 +strExcelPathN 改中文字顯示
'       MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & Replace(strFileName, strExcelPath, "")
'    End If
'
'    Debug.Print "End:   " & Format(ServerTime, "000000")
'    Me.Tag = strTmp   '還原Form的X/Y編號
'    bolExcel = False
'    cmdExcel.Enabled = True
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
End Sub

