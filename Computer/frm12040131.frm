VERSION 5.00
Begin VB.Form frm12040131 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "國內客戶名條"
   ClientHeight    =   4665
   ClientLeft      =   570
   ClientTop       =   4470
   ClientWidth     =   4755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Chk 
      Caption         =   "印所有客戶(原始)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2850
      TabIndex        =   28
      Top             =   2040
      Width           =   1800
   End
   Begin VB.CheckBox Chk 
      Caption         =   "只寄客戶2014後有收文"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   27
      Top             =   2040
      Width           =   2500
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "只寄研發處寄發雜誌名單"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   210
      TabIndex        =   26
      Top             =   2310
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2450
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "N"
      Top             =   465
      Width           =   345
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "Y"
      Top             =   3300
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   2850
      TabIndex        =   5
      Top             =   1740
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   1350
      TabIndex        =   4
      Top             =   1740
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1950
      TabIndex        =   8
      Top             =   3570
      Width           =   705
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1950
      TabIndex        =   9
      Top             =   3870
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   180
      TabIndex        =   15
      Top             =   2610
      Width           =   4275
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   240
         Width           =   3240
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2850
      MaxLength       =   3
      TabIndex        =   2
      Top             =   750
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1350
      MaxLength       =   3
      TabIndex        =   1
      Top             =   750
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3855
      TabIndex        =   11
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3030
      TabIndex        =   10
      Top             =   30
      Width           =   800
   End
   Begin VB.Label Label6 
      Caption         =   "是否含寄電子報對象:             (N: 不含)"
      Height          =   255
      Left            =   540
      TabIndex        =   25
      Top             =   465
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "雙排列印：　　　(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   585
      TabIndex        =   24
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label Label5 
      Caption         =   "雙排名條紙張, 上紙後不必調整上下位置!!!"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   4305
      Width           =   4110
   End
   Begin VB.Label Label4 
      Caption         =   "若不接續上次列印，則不需設定流水號!!!"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   510
      TabIndex        =   22
      Top             =   1470
      Width           =   4155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2580
      TabIndex        =   21
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "流水號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   20
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   540
      TabIndex        =   19
      Top             =   3630
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   540
      TabIndex        =   18
      Top             =   3930
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2580
      TabIndex        =   17
      Top             =   780
      Width           =   180
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2550
      TabIndex        =   14
      Top             =   1170
      Width           =   1395
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   1170
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "業務區： "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   540
      TabIndex        =   12
      Top             =   780
      Width           =   765
   End
End
Attribute VB_Name = "frm12040131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

' 流水號
Dim m_PageNo As Integer
'******** 90.11.14   nick
Dim m_PrinterName As String
Dim Prn As Printer
'**************************
Dim strSql As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add By Cheng 2002/02/30
Dim m_dbl_LeftMargin  As Double '橫軸偏移值
Dim m_dbl_TopMargin  As Double '縱軸偏移值
'Add by Morgan 2004/11/1
Const m_LabelWidth As Double = 5250 '單張寬度
Dim m_bolRigh As Boolean '右邊標籤
Dim m_dbl_LeftMargin_1st As Double '第一張橫軸位置

Public Sub cmdBack_Click()
    'Add By Cheng 2003/01/30
    '若印表機或偏移值有變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Or Me.Text1(7).Text <> Me.Text1(7).Tag Or Me.Text1(8).Text <> Me.Text1(8).Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, Me.Text1(7).Text, Me.Text1(8).Text, Me.Combo1.Text
    End If
    bolToEndByNick = True
    Unload Me
End Sub

Public Sub cmdPrint_Click()
   'Add by Amy 2020/06/20 判斷勾選
   If chk(0).Value = 0 And chk(1).Value = 0 And Chk2.Value = 0 Then
       MsgBox "請勾選列印資料！"
       Exit Sub
   End If
   If chk(0).Value = 1 And chk(1).Value = 1 Then
       MsgBox "「只寄客戶2014後收文」及「印所有客戶(原始)」只能擇一勾選！"
       Exit Sub
   End If
   If Len(Me.Text1(0).Text) > 0 And Len(Me.Text1(1).Text) > 0 Then
      If Me.Text1(0).Text > Me.Text1(1).Text Then
         MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.Text1(0).SetFocus
         Text1_GotFocus 0
         Exit Sub
      End If
   End If
   Me.lblName(0).Caption = StaffQuery(Me.Text1(2).Text): DoEvents
   If Len(Me.Text1(2).Text) > 0 And Len(Me.lblName(0).Caption) <= 0 Then
      MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.Text1(2).SetFocus
      Text1_GotFocus 2
      Exit Sub
   End If
   PUB_RestorePrinter Combo1 'Added by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
   PrintCase
End Sub

Sub PrintCase()
Dim i As Integer
Dim St As String
Dim Page As Integer, iPrint As Integer, IntF As Integer, PriType As Integer, j As Integer
Dim Prn As Printer
Dim nRow As Integer
Dim strCustNo As String '客戶代號
Dim strCustAdd As String '客戶地址
'Add By Cheng 2003/02/25
Dim StrSQLa As String
Dim RsA As New ADODB.Recordset
Dim ii As Double
Dim strRAndDList As String 'Add by Amy 2020/06/20 研發處寄發名單

On Error GoTo ErrHand
'92.2.28 SONIA 客戶國籍>='000' AND <='009' & 有智權人員 & 有聯絡地址或中文地址
'              但  變更名稱資料(即 CU02<>'0') 者不印
Screen.MousePointer = vbHourglass
strSql = " AND (CU10>='000' AND CU10<='009') AND CU32 IS NULL AND CU13 IS NOT NULL AND CU02='0' AND (CU23 IS NOT NULL OR CU31 IS NOT NULL)"
'業務區
If Len(Me.Text1(0).Text) > 0 Then
   strSql = strSql & " AND CU12>='" & Me.Text1(0).Text & "' "
End If
If Len(Me.Text1(1).Text) > 0 Then
   strSql = strSql & " AND CU12<='" & Me.Text1(1).Text & "' "
End If
'智權人員
If Len(Me.Text1(2).Text) > 0 Then
   strSql = strSql & " AND CU13='" & Me.Text1(2).Text & "' "
End If
'2009/4/16 add by sonia 是否含寄電子報對象
If Text1(5) = "N" Then
    strSql = strSql & " AND (CU20||CU116||CU117||CU118 IS NULL OR UPPER(CU20||CU116||CU117||CU118)='NO' OR CU132='N' OR INSTR(CU79,'90 fail')>0) "
End If
'2009/4/16 END

strExc(0) = ""
If chk(0).Value = 1 Or chk(1).Value = 1 Then
    'Memo 2020/06/22 若地址或客戶抬頭長度顯示有修改,需看GetRAndDList是否也需修改
    'Modify By Sindy 2015/2/6 mailzip檔案不要了,改統一抓取postzipdata檔案
    'strExc(0) = "SELECT CU30,DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)),DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15))," & _
    '            "SUBSTR(decode(cu104,null,CU04,cu104),1,20),SUBSTR(decode(cu104,null,CU04,cu104),21,20),CU08,CU01||CU02 FROM CUSTOMER,MAILZIP " & _
    '            " WHERE nvl(CU30,'　　　')=MZ01(+) AND RowNum = 1 " & strSql & " ORDER BY MZ02,MZ03,NVL(CU31,CU23),CU01 "
    'Add by Amy 2020/06/20 +研發處寄發名單 修改Order by 並搬至下方
    'strExc(0) = "SELECT CU30,DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)),DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15))," & _
            "SUBSTR(decode(cu104,null,CU04,cu104),1,20),SUBSTR(decode(cu104,null,CU04,cu104),21,20),CU08,CU01||CU02 FROM CUSTOMER " & _
            " WHERE RowNum = 1 " & strSql & " ORDER BY NVL(CU30,CU112),NVL(CU31,CU23),CU01 "
    strExc(0) = "SELECT NVL(CU30,CU112) as CU30,DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)) as Addr,DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15)) as Addr2," & _
                "SUBSTR(decode(cu104,null,CU04,cu104),1,20) as CmpN,SUBSTR(decode(cu104,null,CU04,cu104),21,20) as CmpN2,CU08 as Contact,CU01||CU02 as CusNo FROM CUSTOMER " & _
                " WHERE RowNum = 1 " & strSql
End If
'2015/2/6 END
'Modify by Amy 2020/06/20 +研發處寄發名單
If Chk2.Value = 1 Then
    strRAndDList = "Select '' as cu30,poc10 as Addr,'' as Addr2,poc03 as CmpN,'' as CmpN2,'' as Contact,poc01||poc02 as CusNo From PotCustomer1 Where InStr(Poc15,'研發處寄發雜誌名單')>0 "
End If
If strExc(0) <> MsgText(601) And strRAndDList <> MsgText(601) Then
    strExc(0) = strExc(0) & " Union " & strRAndDList
ElseIf strExc(0) = MsgText(601) And strRAndDList <> MsgText(601) Then
    strExc(0) = strRAndDList
End If
'end 2020/06/20
strExc(0) = strExc(0) & " Order by cu30,addr,CusNo "
intI = 0
'edit by nickc 2007/02/09 不用 dll 了
'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
If intI <> 1 Then
   Screen.MousePointer = vbDefault
   Exit Sub
End If
If RsTemp.RecordCount <= 0 Then
   MsgBox "資料庫無資料!!!", vbExclamation + vbOKOnly
   Screen.MousePointer = vbDefault
   Exit Sub
Else
    '接續上次列印
    If MsgBox("是否要接續上次列印???", vbExclamation + vbYesNo) = vbYes Then
        'Add by Amy 2020/01/17 避免抓到舊資料,按Yes 流水號必輸
        If Trim(Text1(3)) = MsgText(601) Or Trim(Text1(4)) = MsgText(601) Then
            MsgBox "請輸入流水號!!!", vbExclamation + vbOKOnly
            If Trim(Text1(3)) = MsgText(601) Then
                Text1(3).SetFocus
            Else
                Text1(4).SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'end 2020/01/17
        strExc(0) = ""
        'Add By Cheng 2003/02/26
        '流水號
        If Me.Text1(3).Text <> "" Then
            strExc(0) = strExc(0) & " And CAL02>=" & Val(Me.Text1(3).Text) & " "
        End If
        If Me.Text1(4).Text <> "" Then
            strExc(0) = strExc(0) & " And CAL02<=" & Val(Me.Text1(4).Text) & " "
        End If
        'Modify By Cheng 2003/02/26
'        strExc(0) = "Select CAL03,CAL04,CAL05,CAL06,CAL07,CAL08,CAL09,CAL02 From CustomerAddressList Where CAL01='" & strUserNum & "' Order By CAL02 "
        'Modify By Cheng 2003/04/29
        '不考慮使用者
'        strExc(0) = "Select CAL03,CAL04,CAL05,CAL06,CAL07,CAL08,CAL09,CAL02 From CustomerAddressList Where CAL01='" & strUserNum & "' " & strExc(0) & " Order By CAL02 "
        strExc(0) = "Select CAL03,CAL04,CAL05,CAL06,CAL07,CAL08,CAL09,CAL02 From CustomerAddressList Where CAL01=CAL01 " & strExc(0) & " Order By CAL02 "
        intI = 0
        'edit by nickc 2007/02/09 不用 dll 了
        'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI <> 1 Then
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    '不接續上次列印
    Else
        '不考慮使用者, 刪除所有暫存資料
        StrSQLa = "Delete From  CustomerAddressList "
        cnnConnection.Execute StrSQLa
        
        'Modify by Amy 2020/06/20 +if 可單獨列印
        ii = 1
        If chk(0).Value = 1 Or chk(1).Value = 1 Then
            'Modify by Morgan 2008/8/11 接洽人改先用聯絡人編號抓聯絡人檔,若有聯絡人編號但該聯絡人無地址時抓原客戶地址
            'Modify By Sindy 2015/2/6 mailzip檔案不要了,改統一抓取postzipdata檔案
    '        strExc(0) = "SELECT CU30,DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)),DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15))" & _
    '                    ",nvl(SUBSTR(nvl(cu104,CU04),1,20),CU05),decode(nvl(cu104,CU04),null,CU88,SUBSTR(nvl(cu104,CU04),21,20)),CU08,CU01||CU02" & _
    '                    " FROM (SELECT NVL(PCC21,CU30) CU30,NVL(PCC22,CU31) CU31,PCC05 CU08,CU01,CU02,CU104,CU04,CU88,CU05,CU23" & _
    '                    " FROM CUSTOMER,POTCUSTCONT WHERE PCC01(+)=CU01 AND PCC02(+)=CU127 " & strSql & ") X,MAILZIP " & _
    '                    " WHERE nvl(CU30,'　　　')=MZ01(+) ORDER BY MZ02,MZ03,NVL(CU31,CU23),CU01 "
            strExc(0) = "SELECT NVL(CU30,CU112),DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)),DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15))" & _
                        ",nvl(SUBSTR(nvl(cu104,CU04),1,20),CU05),decode(nvl(cu104,CU04),null,CU88,SUBSTR(nvl(cu104,CU04),21,20)),CU08,CU01||CU02" & _
                        " FROM (SELECT NVL(PCC21,CU30) CU30,NVL(PCC22,CU31) CU31,PCC05 CU08,CU01,CU02,CU104,CU04,CU88,CU05,CU23,CU112" & _
                        " FROM CUSTOMER,POTCUSTCONT WHERE PCC01(+)=CU01 AND PCC02(+)=CU127 " & strSql & ") X " & _
                        " ORDER BY NVL(CU30,CU112),NVL(CU31,CU23),CU01 "
            '2015/2/6 END
            intI = 0
            'edit by nickc 2007/02/09 不用 dll 了
            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI <> 1 Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            'ii = 1'Mark by Amy 2020/06/20 往上搬
    'Add by Morgan 2004/11/15 當客戶編號前6碼同且地址也相同時不印
             strCustNo = "No"
             strCustAdd = "Add"
    '2004/11/15
            RsA.CursorLocation = adUseClient
            RsA.Open "CustomerAddressList", cnnConnection, adOpenDynamic, adLockOptimistic
            While Not RsTemp.EOF
    '            strSQLA = "Insert Into CustomerAddressList Values ('" & strUserNum & "', " & ii & ",'" & rsTemp.Fields(0).Value & "','" & rsTemp.Fields(1).Value & "','" & rsTemp.Fields(2).Value & "','" & rsTemp.Fields(3).Value & "','" & rsTemp.Fields(4).Value & "','" & rsTemp.Fields(5).Value & "','" & rsTemp.Fields(6).Value & "' )"
    '            cnnConnection.Execute strSQLA
    
    'Modify by Morgan 2004/11/15 當客戶編號前6碼同且地址也相同時不印
                If Not (strCustNo = Left(RsTemp.Fields(6).Value, 6) And strCustAdd = "" & RsTemp.Fields(1).Value & "" & RsTemp.Fields(2).Value) Then
                   RsA.AddNew
                   RsA.Fields(0).Value = strUserNum '使用者名稱
                   RsA.Fields(1).Value = ii '流水號
                   RsA.Fields(2).Value = "" & RsTemp.Fields(0).Value '郵遞區號
                   RsA.Fields(3).Value = "" & RsTemp.Fields(1).Value '聯絡地址1
                   RsA.Fields(4).Value = "" & RsTemp.Fields(2).Value '聯絡地址2
                   RsA.Fields(5).Value = "" & RsTemp.Fields(3).Value '客戶抬頭1
                   RsA.Fields(6).Value = "" & RsTemp.Fields(4).Value '客戶抬頭2
                   RsA.Fields(7).Value = "" & RsTemp.Fields(5).Value '聯絡人
                   RsA.Fields(8).Value = "" & RsTemp.Fields(6).Value '客戶代號
                   RsA.UPDATE
                   strCustNo = Left(RsTemp.Fields(6).Value, 6)
                   strCustAdd = "" & RsTemp.Fields(1).Value & "" & RsTemp.Fields(2).Value
                   ii = ii + 1
                End If
    '2004/11/15 end
                RsTemp.MoveNext
            Wend
        End If
        'end 2020/06/20
        
        'Add by Amy2020/01/17 +只寄客戶2014後有收文
        If chk(0).Value = 1 Then
            If Get2014AfterData = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        'Add by Amy 2020/06/20 只寄研發處寄發雜誌名單
        ElseIf Chk2.Value = 1 Then
            If GetRAndDList("CustomerAddressList", ii) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        'end 2020/01/17
        'Add by Amy 2020/06/20 +if 單獨印「研發處寄發雜誌名單」序號於此排
        If (chk(0).Value = 0 And chk(1).Value = 0) Or (chk(1).Value = 1 And Chk2.Value = 1) Then
            Call SetTmpTB
        End If
        'end 2020/06/20
        
        'Modify By Cheng 2003/04/29
        '不考慮使用者
'        strExc(0) = "Select CAL03,CAL04,CAL05,CAL06,CAL07,CAL08,CAL09,CAL02 From CustomerAddressList Where CAL01='" & strUserNum & "' Order By CAL02 "
        strExc(0) = "Select CAL03,CAL04,CAL05,CAL06,CAL07,CAL08,CAL09,CAL02 From CustomerAddressList Order By CAL02 "
        intI = 0
        'edit by nickc 2007/02/09 不用 dll 了
        'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI <> 1 Then
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    'Add By Cheng 2002/02/30
    '設定偏移值
    'Modify by Morgan 2004/11/1 考慮雙排貼紙的左邊位置
    m_bolRigh = False
    m_dbl_LeftMargin = CDbl(Me.Text1(7).Text) * 567
    If Text2.Text = "Y" Then
      m_dbl_LeftMargin = m_dbl_LeftMargin + 200
    End If
    m_dbl_LeftMargin_1st = m_dbl_LeftMargin
    '2004/11/1 end
    
    m_dbl_TopMargin = CDbl(Me.Text1(8).Text) * 567
   IntF = 7
   Printer.Font.Size = 12
'Modify by Morgan 2011/10/24
'    'Modify By Cheng 2003/02/25
'    '設定列印有效區域
'   'Printer.Height = 2200
'   'Printer.Width = 10000
'    Printer.Height = 2160
'    'Modify by Morgan 2004/11/15 考慮雙排紙
'    'Printer.Width = 10000
'    Printer.Width = 20000
   Printer.PaperSize = PUB_GetPaperSize(13)
   If Printer.Width = Printer.ScaleWidth Then
      m_dbl_LeftMargin = m_dbl_LeftMargin + 250
   End If
   m_dbl_LeftMargin_1st = m_dbl_LeftMargin
'end 2011/10/24

   RsTemp.MoveFirst
   iPrint = 1
'Remove by Morgan 2004/11/15 改在產生資料時就過濾
'   strCustNo = "No"
'   strCustAdd = "Add"

   With RsTemp
      Do While Not .EOF
        'Add By Cheng 2003/02/25
        '設定流水號
        iPrint = RsTemp("CAL02").Value
'Remove by Morgan 2004/11/15 改在產生資料時就過濾
'         If strCustNo = Left(rsTemp.Fields(6).Value, 6) And strCustAdd = "" & rsTemp.Fields(1).Value & "" & rsTemp.Fields(2).Value Then
'            GoTo NextRecord
'         Else
'            strCustNo = Left(rsTemp.Fields(6).Value, 6)
'            strCustAdd = "" & rsTemp.Fields(1).Value & "" & rsTemp.Fields(2).Value
'         End If
         nRow = 0
         For i = 0 To IntF - 1
            'Modify By Cheng 2003/02/25
'            Printer.CurrentX = 1000
            Printer.CurrentX = 0 + m_dbl_LeftMargin
            If IsNull(.Fields(i)) = False Then
               If IsEmptyText(.Fields(i)) = False Then
                    'Modify By Cheng 2003/02/25
'                  Printer.CurrentY = i * 250
                  Printer.CurrentY = i * 250 + m_dbl_TopMargin
                  nRow = nRow + 1
               End If
            End If
'            If IsNull(.Fields(i)) = False Then
'               If IsEmptyText(.Fields(i)) = False Then
'                  'Modify By Cheng 2003/03/04
'                  If i = 3 Then
'                     If "" & .Fields(3).Value <> "" And "" & .Fields(5).Value = "" Then
'                        Printer.Print "" & .Fields(i) & "　　　　　　君　　鈞啟"
'                     Else
'                        Printer.Print "" & .Fields(i)
'                     End If
'                  ElseIf i = 5 Then
'                     If "" & .Fields(5).Value <> "" Then
'                        Printer.Print "" & .Fields(i) & "　　　　　　君　　鈞啟"
'                     Else
'                        Printer.Print "" & .Fields(i)
'                     End If
'                  Else
'                     Printer.Print "" & .Fields(i)
'                   End If
'               End If
'            End If
            If i = 5 Then
                '若有資料
                If "" & .Fields(i) <> "" Then
                    Printer.Print "" & .Fields(i) & "　　　　　　君　　鈞啟"
                '若無資料
                Else
                    Printer.Print "　　　　　　　　　　君　　鈞啟"
                End If
            Else
                Printer.Print "" & .Fields(i)
            End If
         Next
        'Modify By Cheng 2003/02/25
'         Printer.CurrentX = 5000
'         Printer.CurrentY = (i - 1) * 250
         Printer.CurrentX = 4000 + m_dbl_LeftMargin
         Printer.CurrentY = (i - 1) * 250 + m_dbl_TopMargin
         If m_PageNo > 0 Then
            Printer.Print Format(m_PageNo, "000000")
         Else
            Printer.Print Format(iPrint, "000000")
         End If
         iPrint = iPrint + 1
         'Modify by Morgan 2004/11/1 加雙排列印選項
         'Printer.NewPage
         If Text2.Text = "Y" Then
            If m_bolRigh = False Then
               m_dbl_LeftMargin = m_dbl_LeftMargin_1st + m_LabelWidth
            Else
               Printer.NewPage
               m_dbl_LeftMargin = m_dbl_LeftMargin_1st
            End If
            m_bolRigh = Not m_bolRigh
         Else
            Printer.NewPage
         End If
         
        'DoEvents
NextRecord:
        'Modify By Cheng 2003/02/26
        '不要刪除暫存資料
'        'Add By Cheng 2003/02/25
'        '刪除列印過的資料
'        strSQLA = "Delete From CustomerAddressList Where CAL01='" & strUserNum & "' And CAL02=" & .Fields("CAL02").Value
'        cnnConnection.Execute strSQLA
         .MoveNext
      Loop
   End With
   Printer.EndDoc
'   ShowPrintOk
    MsgBox "列印完成，總筆數為 " & Format(iPrint - 1, "#,##0") & " 筆資料!!!", vbExclamation + vbOKOnly
End If
Screen.MousePointer = vbDefault
Exit Sub
ErrHand:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
   Unload Me
End Select
End Sub

Private Sub Form_Load()

MoveFormToCenter Me

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint, Text1(7), Text1(8)   'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
chk(0).Value = 1 'Add by Amy 2020/01/17
Chk2.Value = 1 'Add by Amy 2020/06/20
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
    Case 2 '智權人員
        'Add By Cheng 2002/06/07
        Me.lblName(0).Caption = StaffQuery(Me.Text1(2).Text)
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   '2009/4/16 ADD BY SONIA
   Select Case Index
      Case 1
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
   End Select
   '2009/4/16 END
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 1
   'Add By Cheng 2002/06/07
   If Me.Text1(0).Text > Me.Text1(1).Text Then
      MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.Text1(0).SetFocus
      TextInverse Me.Text1(0)
   End If
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then
      Cancel = True
   End If
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    Set frm12040131 = Nothing
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
CheckKeyIn = True
Select Case intIndex
Case 0
   If Len(Me.Text1(1).Text) > 0 Then
      If Me.Text1(0).Text > Me.Text1(1).Text Then
         MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         CheckKeyIn = False
      End If
   End If
Case 1
   If Len(Me.Text1(0).Text) > 0 And Len(Me.Text1(1).Text) > 0 Then
      If Me.Text1(0).Text > Me.Text1(1).Text Then
         'Modify By Cheng 2002/06/07
         '改在LostFocus時判斷
'         MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
'         CheckKeyIn = False
         CheckKeyIn = True
      End If
   End If
Case 2
   Me.lblName(0).Caption = StaffQuery(Me.Text1(2).Text)
   'Add By Cheng 2002/06/07
   If Len(Me.Text1(2).Text) > 0 Then
      If Len(Me.lblName(0).Caption) <= 0 Then
         MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
         CheckKeyIn = False
      End If
   End If
End Select
End Function

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'Add by Morgan 2004/11/1
Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2.Text <> "Y" Then Text2.Text = "N"
End Sub

'Add by Amy 2020/01/17
Private Function Get2014AfterData() As Boolean
    Dim strExe As String
    
On Error GoTo ErrHand
    
    Call DelTmpTB(True)
    strExe = "Create Table PGMID.CUST_20190816 (CU01 VARCHAR2(8) NOT NULL,MAXCP05 NUMBER(8,0) NULL,GPMAXCP05 NUMBER(8,0) NULL," & _
                  "PRIMARY KEY (CU01))  PCTFREE 10 PCTUSED 40 INITRANS 1 MAXTRANS 255 " & _
                  "STORAGE ( INITIAL 1216K NEXT 208K MINEXTENTS 1 MAXEXTENTS 2147483645 PCTINCREASE 1) TABLESPACE USR"
    cnnConnection.Execute strExe
    
    strExe = "CREATE TABLE PGMID.CUST_20190816A (CU01 VARCHAR2(8) NOT NULL,MAXCP05 NUMBER(8,0) NULL," & _
                  "PRIMARY KEY (CU01))  PCTFREE 10 PCTUSED 40 INITRANS 1 MAXTRANS 255 " & _
                  "STORAGE ( INITIAL 1216K NEXT 208K MINEXTENTS 1 MAXEXTENTS 2147483645 PCTINCREASE 1) TABLESPACE USR"
    cnnConnection.Execute strExe
    
    strExe = "TrunCate Table CUST_20190816"
    cnnConnection.Execute strExe
    
    '將先前存於CustomerAddressList的顧問客戶名條 資料寫入 CUST_20190816
    strExe = "Insert Into CUST_20190816 (CU01) Select SUBSTR(CAL09,1,8) From CustomerAddressList"
    cnnConnection.Execute strExe
    
    '--所有收文客戶
    strExe = "TrunCate Table CUST_20190816A"
    cnnConnection.Execute strExe
    
    strExe = "INSERT INTO CUST_20190816A (Select PA26,MAX(CP05) CP05 From " & _
                "(Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(PA26,1,8) PA26 From CASEPROGRESS,PATENT,CUST_20190816 " & _
                 "Where SubStr(PA26,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                    "And CP01 IN ('P','CFP','FCP') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(PA27,1,8) From CASEPROGRESS,PATENT,CUST_20190816 " & _
                 "Where SubStr(PA27,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                    "And CP01 IN ('P','CFP','FCP') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(PA28,1,8) From CASEPROGRESS,PATENT,CUST_20190816 " & _
                "Where SubStr(PA28,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                    "And CP01 IN ('P','CFP','FCP') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(PA29,1,8) From CASEPROGRESS,PATENT,CUST_20190816 " & _
                "Where SubStr(PA29,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                    "And CP01 IN ('P','CFP','FCP') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(PA30,1,8) From CASEPROGRESS,PATENT,CUST_20190816 " & _
                "Where SubStr(PA30,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                    "And CP01 IN ('P','CFP','FCP') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) "
    '商標
    strExe = strExe & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(TM23,1,8) From CASEPROGRESS,TRADEMARK,CUST_20190816 " & _
                "Where SubStr(TM23,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=TM01 And CP02(+)=TM02 And CP03(+)=TM03 And CP04(+)=TM04 " & _
                    "And CP01 IN ('T','CFT','FCT','TF') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(TM78,1,8) From CASEPROGRESS,TRADEMARK,CUST_20190816 " & _
                "Where SubStr(TM78,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=TM01 And CP02(+)=TM02 And CP03(+)=TM03 And CP04(+)=TM04 " & _
                    "And CP01 IN ('T','CFT','FCT','TF') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(TM79,1,8) From CASEPROGRESS,TRADEMARK,CUST_20190816 " & _
                "Where SubStr(TM79,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=TM01 And CP02(+)=TM02 And CP03(+)=TM03 And CP04(+)=TM04 " & _
                    "And CP01 IN ('T','CFT','FCT','TF') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(TM80,1,8) From CASEPROGRESS,TRADEMARK,CUST_20190816 " & _
                "Where SubStr(TM80,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=TM01 And CP02(+)=TM02 And CP03(+)=TM03 And CP04(+)=TM04 " & _
                    "And CP01 IN ('T','CFT','FCT','TF') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(TM81,1,8) From CASEPROGRESS,TRADEMARK,CUST_20190816 " & _
                "Where SubStr(TM81,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=TM01 And CP02(+)=TM02 And CP03(+)=TM03 And CP04(+)=TM04 " & _
                    "And CP01 IN ('T','CFT','FCT','TF') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) "
    '法務
    strExe = strExe & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(LC11,1,8) From CASEPROGRESS,LAWCASE,CUST_20190816 " & _
                "Where SubStr(LC11,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=LC01 And CP02(+)=LC02 And CP03(+)=LC03 And CP04(+)=LC04 " & _
                    "And CP01 IN ('L','CFL','FCL','LIN') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(LC43,1,8) From CASEPROGRESS,LAWCASE,CUST_20190816 " & _
                "Where SubStr(LC43,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=LC01 And CP02(+)=LC02 And CP03(+)=LC03 And CP04(+)=LC04 " & _
                    "And CP01 IN ('L','CFL','FCL','LIN') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(LC44,1,8) From CASEPROGRESS,LAWCASE,CUST_20190816 " & _
                "Where SubStr(LC44,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=LC01 And CP02(+)=LC02 And CP03(+)=LC03 And CP04(+)=LC04 " & _
                    "And CP01 IN ('L','CFL','FCL','LIN') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(LC45,1,8) From CASEPROGRESS,LAWCASE,CUST_20190816 " & _
                "Where SubStr(LC45,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=LC01 And CP02(+)=LC02 And CP03(+)=LC03 And CP04(+)=LC04 " & _
                    "And CP01 IN ('L','CFL','FCL','LIN') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(LC46,1,8) From CASEPROGRESS,LAWCASE,CUST_20190816 " & _
                "Where SubStr(LC46,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=LC01 And CP02(+)=LC02 And CP03(+)=LC03 And CP04(+)=LC04 " & _
                    "And CP01 IN ('L','CFL','FCL','LIN') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) "
    '服務業務
    strExe = strExe & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(SP08,1,8) From CASEPROGRESS,SERVICEPRACTICE,CUST_20190816 " & _
                "Where SubStr(SP08,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=SP01 And CP02(+)=SP02 And CP03(+)=SP03 And CP04(+)=SP04 " & _
                    "And CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(SP58,1,8) From CASEPROGRESS,SERVICEPRACTICE,CUST_20190816 " & _
                "Where SubStr(SP58,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=SP01 And CP02(+)=SP02 And CP03(+)=SP03 And CP04(+)=SP04 " & _
                    "And CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(SP59,1,8) From CASEPROGRESS,SERVICEPRACTICE,CUST_20190816 " & _
                "Where SubStr(SP59,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=SP01 And CP02(+)=SP02 And CP03(+)=SP03 And CP04(+)=SP04 " & _
                    "And CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(SP65,1,8) From CASEPROGRESS,SERVICEPRACTICE,CUST_20190816 " & _
                "Where SubStr(SP65,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=SP01 And CP02(+)=SP02 And CP03(+)=SP03 And CP04(+)=SP04 " & _
                    "And CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(SP66,1,8) From CASEPROGRESS,SERVICEPRACTICE,CUST_20190816 " & _
                "Where SubStr(SP66,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=SP01 And CP02(+)=SP02 And CP03(+)=SP03 And CP04(+)=SP04 " & _
                    "And CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) "
    '顧問
    strExe = strExe & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(HC05,1,8) From CASEPROGRESS,HIRECASE,CUST_20190816 " & _
                "Where SubStr(HC05,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=HC01 And CP02(+)=HC02 And CP03(+)=HC03 And CP04(+)=HC04 " & _
                    "And CP01='LA' And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(HC24,1,8) From CASEPROGRESS,HIRECASE,CUST_20190816 " & _
                "Where SubStr(HC24,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=HC01 And CP02(+)=HC02 And CP03(+)=HC03 And CP04(+)=HC04 " & _
                    "And CP01='LA' And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(HC25,1,8) From CASEPROGRESS,HIRECASE,CUST_20190816 " & _
                "Where SubStr(HC25,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=HC01 And CP02(+)=HC02 And CP03(+)=HC03 And CP04(+)=HC04 " & _
                    "And CP01='LA' And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(HC26,1,8) From CASEPROGRESS,HIRECASE,CUST_20190816 " & _
                "Where SubStr(HC26,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=HC01 And CP02(+)=HC02 And CP03(+)=HC03 And CP04(+)=HC04 " & _
                    "And CP01='LA' And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
    "Union Select CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SubStr(HC27,1,8) From CASEPROGRESS,HIRECASE,CUST_20190816 " & _
                "Where SubStr(HC27,1,8)=CU01(+) And CU01 Is Not Null And CP01(+)=HC01 And CP02(+)=HC02 And CP03(+)=HC03 And CP04(+)=HC04 " & _
                    "And CP01='LA' And CP09<'B' And (CP158>0 OR (CP158=0 And CP159=0)) " & _
  ") Group BY PA26) "
  cnnConnection.Execute strExe
  
    '--更新資料
    strExe = "Update CUST_20190816 B Set B.MAXCP05=(Select A.MAXCP05 From CUST_20190816A A Where B.CU01=A.CU01)"
    cnnConnection.Execute strExe
    
    strExe = "Update CUST_20190816 B Set B.GPMAXCP05=(Select MAX(A.MAXCP05) From CUST_20190816A A Where SubStr(B.CU01,1,6)=SubStr(A.CU01,1,6))"
    cnnConnection.Execute strExe
    
    'COPY Customeraddresslist 改名Cust_List1014
    strExe = "Create Table Cust_List1014 as Select * From Customeraddresslist Where 1=2"
    cnnConnection.Execute strExe
    
    strExe = "Insert Into Cust_List1014 " & _
                  "Select * From Customeraddresslist " & _
                 "Where Substr(Cal09,1,8) In (Select Cu01 From Cust_20190816 Where Decode(SIGN(NVL(GPMAXCP05,0)-20140101),-1,0,1)=1)"
    cnnConnection.Execute strExe
    
    'Add by Amy 2020/06/20 +寄研發處寄發雜誌名單
    If Chk2.Value = 1 Then
       If GetRAndDList("Cust_List1014") = False Then
            Get2014AfterData = False
            Exit Function
       End If
    End If
    
    strExe = "Delete From Customeraddresslist"
    cnnConnection.Execute strExe
    
    'Modify by Amy 2020/06/20 +if 混合列印時,無Zip 排前,再排編號,地址
    If chk(0).Value = 1 And Chk2.Value = 1 Then
        strExe = "Insert Into Customeraddresslist (cal01,cal02,cal03,cal04,cal05,cal06,cal07,cal08,cal09) " & _
                    "Select Cal01,Row_Number() Over (Order By cal02),Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 " & _
                    "From ( Select CAL01,Decode(cal03,'　',1,Decode(SubStr(cal09,1,1),'X',2,3)) as Cal02,Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 " & _
                                "From Cust_List1014 Order by cal02,cal03,cal09,cal06 " & _
                            ")"

    Else
        strExe = "Insert Into Customeraddresslist (cal01,cal02,cal03,cal04,cal05,cal06,cal07,cal08,cal09) " & _
                    "Select Cal01,Row_Number() Over (Order By cal02),Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 From Cust_List1014 "
    End If
    
    cnnConnection.Execute strExe
    'end 2020/06/20
    
    Call DelTmpTB
    Get2014AfterData = True
    Exit Function
    
ErrHand:
    MsgBox Err.Description
End Function

Private Sub DelTmpTB(Optional ByVal IsFirst As Boolean = False)
    Dim RsQ As New ADODB.Recordset
    Dim strq As String
    Dim intR As Integer
    
    strq = "Select * From tab Where Upper(TNAME) =Upper('CUST_20190816' )"
    intR = 1
    If RsQ.State <> adStateClosed Then RsQ.Close
    Set RsQ = ClsLawReadRstMsg(intR, strq)
    If intR = 1 Then
        cnnConnection.Execute "Drop Table CUST_20190816"
    End If
    
    
    strq = "Select * From tab Where Upper(TNAME) =Upper('CUST_20190816A') "
    intR = 1
    If RsQ.State <> adStateClosed Then RsQ.Close
    Set RsQ = ClsLawReadRstMsg(intR, strq)
    If intR = 1 Then
        cnnConnection.Execute "Drop Table CUST_20190816A"
    End If
    
    '第一次刪除Cust_List1014
    If IsFirst = True Then
        strq = "Select * From tab Where Upper(TNAME) =Upper('Cust_List1014') "
        intR = 1
        If RsQ.State <> adStateClosed Then RsQ.Close
        Set RsQ = ClsLawReadRstMsg(intR, strq)
        If intR = 1 Then
            cnnConnection.Execute "Drop Table Cust_List1014"
        End If
    End If
    If RsQ.State <> adStateClosed Then RsQ.Close
    Set RsQ = Nothing
End Sub

'Add by Amy 2020/06/20 +研發處寄發雜誌名單
Private Function GetRAndDList(stTmpTB As String, Optional DouSeq As Double = 1) As Boolean
    Dim RsQ As ADODB.Recordset
    Dim intq As Integer
    Dim strq As String, strIns As String, strZip As String
    Dim strAddr As String, strCal04 As String, strCal05 As String, strCName As String, strCal06 As String, strCal07 As String
    Dim strCal08 As String, strCal09 As String
    
On Error GoTo ErrHand
    GetRAndDList = False
    
    strq = "Select * From PotCustomer1 Where InStr(Poc15,'研發處寄發雜誌名單')>0 "
    intq = 1
    Set RsQ = ClsLawReadRstMsg(intq, strq)
    If intq = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            strZip = "": strCal04 = "": strCal05 = "": strCal06 = "": strCal07 = "": strCal08 = ""
            '拆地址(Cal04/05)及Zip(Cal03)
            strAddr = "" & RsQ.Fields("poc10")
            Do While Left(strAddr, 1) >= "０" And Left(strAddr, 1) <= "９"
                strZip = strZip & Left(strAddr, 1)
                strAddr = Mid(strAddr, 2)
            Loop
            strAddr = Replace(RsQ.Fields("poc10"), strZip, "")
            strCal04 = Mid(strAddr, 1, 20)
            If Replace(strAddr, strCal04, "") <> MsgText(601) Then
                strCal05 = Mid(strAddr, 21, 15)
            End If
            If strZip = MsgText(601) Then strZip = "　" '空值設為全型空白
            
            '拆客戶名稱(Cal06/07)及聯絡人(Cal08)
            strCName = "" & RsQ.Fields("poc03")
            If InStr(strCName, "-") > 0 Then
                strCal08 = Mid(strCName, InStr(strCName, "-"))
                strCName = Replace(strCName, strCal08, "")
                strCal08 = Mid(strCal08, 2)
            End If
            strCal06 = Mid(strCName, 1, 20)
            If Replace(strCName, strCal06, "") <> MsgText(601) Then
                strCal07 = Mid(strAddr, 21, 20)
            End If
            strCal09 = RsQ.Fields("poc01") & RsQ.Fields("poc02") '潛客編號
            
            '新增至暫存檔(因需拆解Zip,無Zip需先印,故序號於最後加
            strIns = "Insert into " & stTmpTB & " (cal01,cal02,cal03,cal04,cal05,cal06,cal07,cal08,cal09) " & _
                        "Values ('" & strUserNum & "'," & DouSeq & ",'" & ChgSQL(strZip) & "'," & CNULL(ChgSQL(strCal04)) & "," & CNULL(ChgSQL(strCal05)) & _
                            "," & CNULL(ChgSQL(strCal06)) & "," & CNULL(ChgSQL(strCal07)) & _
                            "," & CNULL(ChgSQL(strCal08)) & "," & CNULL(ChgSQL(strCal09)) & " )"
            cnnConnection.Execute strIns
            DouSeq = DouSeq + 1 'CustomerAddressList cal02為key
            RsQ.MoveNext
        Loop
    End If
    
    Set RsQ = Nothing
    GetRAndDList = True
    Exit Function
    
ErrHand:
    MsgBox Err.Description
End Function

Private Sub SetTmpTB()
    Dim strExe As String
    
    '刪除暫存檔
    Call DelTmpTB(True)
    
    '為可分次列印,排序完需回寫 CustomerAddressList
    'COPY Customeraddresslist 改名Cust_List1014 ()
    strExe = "Create Table Cust_List1014 as Select * From CustomerAddressList "
    cnnConnection.Execute strExe
    
    strExe = "Delete From  CustomerAddressList "
    cnnConnection.Execute strExe
        
    '只印寄研發處寄發雜誌名單
    If chk(0).Value = 0 And chk(1).Value = 0 Then
        strExe = "Insert Into CustomerAddressList (cal01,cal02,cal03,cal04,cal05,cal06,cal07,cal08,cal09) " & _
                         "Select * From (Select Cal01,Row_Number() Over (Order By cal03,cal06,cal09) as Cal02,Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 " & _
                         "From Cust_List1014 ) Order By Cal02 "
         cnnConnection.Execute strExe
    '印所有客戶(原始)+研發處寄發雜誌名單
    ElseIf chk(1).Value = 1 And Chk2.Value = 1 Then
        strExe = "Insert Into CustomerAddressList (cal01,cal02,cal03,cal04,cal05,cal06,cal07,cal08,cal09) " & _
                        "Select Cal01,Row_Number() Over (Order By cal02),Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 " & _
                        "From ( Select CAL01,Decode(cal03,'　',1,Decode(SubStr(cal09,1,1),'X',2,3)) as Cal02,Cal03,Cal04,Cal05,Cal06,Cal07,Cal08,Cal09 " & _
                                    "From Cust_List1014 Order by cal02,cal03,cal09,cal06 " & _
                                    ")"
        cnnConnection.Execute strExe
    End If
End Sub
