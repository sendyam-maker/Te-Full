VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060304 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費通知函"
   ClientHeight    =   5556
   ClientLeft      =   4380
   ClientTop       =   2736
   ClientWidth     =   7272
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5556
   ScaleWidth      =   7272
   Begin VB.TextBox txtData 
      Height          =   280
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   3384
      Width           =   750
   End
   Begin VB.TextBox txtData 
      Height          =   280
      Index           =   1
      Left            =   3144
      TabIndex        =   2
      Top             =   432
      Width           =   900
   End
   Begin VB.TextBox txtData 
      Height          =   280
      Index           =   0
      Left            =   2070
      TabIndex        =   1
      Top             =   432
      Width           =   900
   End
   Begin VB.CommandButton cmdOK2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5568
      TabIndex        =   3
      Top             =   150
      Width           =   756
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整批期限："
      Height          =   204
      Index           =   2
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   1224
   End
   Begin VB.Frame Frame5 
      Caption         =   "改版前的選項"
      Height          =   1452
      Left            =   6888
      TabIndex        =   41
      Top             =   3216
      Visible         =   0   'False
      Width           =   4332
      Begin VB.CommandButton cmdOK 
         Caption         =   "確定(&O)"
         Height          =   400
         Index           =   0
         Left            =   1464
         TabIndex        =   53
         Top             =   96
         Width           =   756
      End
      Begin VB.OptionButton Option1 
         Caption         =   "整批"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   492
         Width           =   792
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   5
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   47
         Top             =   1116
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   6
         Left            =   2736
         MaxLength       =   7
         TabIndex        =   46
         Top             =   1116
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   1080
         TabIndex        =   42
         Top             =   360
         Width           =   3135
         Begin VB.OptionButton Option2 
            Caption         =   "1~10號"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "11~20號"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   44
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "21~月底"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   43
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上次通知日"
         Height          =   180
         Index           =   0
         Left            =   336
         TabIndex        =   52
         Top             =   876
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2160
         X2              =   2280
         Y1              =   1236
         Y2              =   1236
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   1
         Left            =   2796
         TabIndex        =   51
         Top             =   876
         Width           =   480
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2520
         X2              =   2640
         Y1              =   948
         Y2              =   948
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   50
         Top             =   876
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日"
         Height          =   180
         Index           =   2
         Left            =   336
         TabIndex        =   49
         Top             =   1152
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Enabled         =   0   'False
      Height          =   2440
      Left            =   8064
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   12
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   39
         Top             =   2180
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   11
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   38
         Top             =   2180
         Width           =   975
      End
      Begin VB.CheckBox Chk1 
         Caption         =   "代理人Y2006500(ARCO)和Y5285800(Shinjyu) + 4個月"
         Height          =   300
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   9
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   25
         Top             =   1360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   10
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   24
         Top             =   1360
         Width           =   975
      End
      Begin VB.CheckBox Chk1 
         Caption         =   "代理人Y5133301(北京銀龍)和Y53496 (Gotoh) + 4個月"
         Height          =   300
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   7
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   21
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   8
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   20
         Top             =   540
         Width           =   975
      End
      Begin VB.CheckBox Chk1 
         Caption         =   "代理人Y2062400(YAMASAKI && PARTNERS) + 7個月"
         Height          =   300
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   0
         Width           =   4695
      End
      Begin VB.Line Line4 
         X1              =   2910
         X2              =   3045
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Line Line3 
         X1              =   2910
         X2              =   3045
         Y1              =   2045
         Y2              =   2045
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   6
         Left            =   1950
         TabIndex        =   37
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   7
         Left            =   3180
         TabIndex        =   36
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日"
         Height          =   300
         Index           =   7
         Left            =   720
         TabIndex        =   35
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上次通知日"
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   34
         Top             =   1980
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   2910
         X2              =   3030
         Y1              =   1480
         Y2              =   1480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上次通知日"
         Height          =   180
         Index           =   5
         Left            =   720
         TabIndex        =   32
         Top             =   1160
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   5
         Left            =   3180
         TabIndex        =   31
         Top             =   1160
         Width           =   480
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2910
         X2              =   3030
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   4
         Left            =   1950
         TabIndex        =   30
         Top             =   1160
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上次通知日"
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   29
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   3
         Left            =   3180
         TabIndex        =   28
         Top             =   330
         Width           =   480
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2910
         X2              =   3030
         Y1              =   400
         Y2              =   400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   2
         Left            =   1950
         TabIndex        =   27
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日"
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   26
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日"
         Height          =   180
         Index           =   3
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2910
         X2              =   3030
         Y1              =   660
         Y2              =   660
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "設定請款單及定稿"
      Height          =   660
      Left            =   552
      TabIndex        =   14
      Top             =   4668
      Width           =   5505
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   12
         Top             =   240
         Width           =   4620
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   263
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   552
      TabIndex        =   13
      Top             =   3912
      Width           =   5505
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   240
         Width           =   4620
      End
      Begin VB.Label Label3 
         Caption         =   "印表機"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   263
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   4
      Left            =   3648
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2940
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   3
      Left            =   3402
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2940
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   2
      Left            =   2556
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2940
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   280
      Index           =   1
      Left            =   2070
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "FCP"
      Top             =   2940
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   228
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   2964
      Width           =   1272
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6360
      TabIndex        =   4
      Top             =   150
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1908
      Left            =   120
      TabIndex        =   40
      Top             =   864
      Width           =   7032
      _ExtentX        =   12404
      _ExtentY        =   3366
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      AllowUserResizing=   3
      FormatString    =   "V|代理人|名稱|前X月|上次通知日|(止)|下次繳費日|(止)"
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSForms.Label lblFM2 
      Height          =   252
      Left            =   2472
      TabIndex        =   55
      Top             =   3408
      Width           =   1020
      Size            =   "1799;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "管制人："
      Height          =   204
      Left            =   888
      TabIndex        =   54
      Top             =   3432
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2904
      X2              =   3360
      Y1              =   552
      Y2              =   552
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "※整批列印結束時會一併列印同區間內寰華案清單"
      ForeColor       =   &H00FF00FF&
      Height          =   220
      Left            =   108
      TabIndex        =   17
      Top             =   60
      Width           =   4584
   End
End
Attribute VB_Name = "frm060304"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/05/29 繳年費通知函仿照實審通知函(frm060325)，拆分給各區管制人；隱藏cmdOK(0)、Option1(0),Option2(0~2)、上次通知日、下次繳費日
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Const ET01 As String = "08"
Dim m_ii As Integer
Dim m_strTxt(1 To 20) As String '新增例外欄位使用
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
'Add by Morgan 2011/3/15
Dim strPrinter As String
'Add by Lydia 2014/10/23 增加代理人Y20624和整批條件選項
Dim mDate() As String ', mSrvDate As String 'Remove by Lydia 2016/04/19
Dim strPrinter2 As String 'Add By Sindy 2015/7/6
Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2015/10/30
Public m_SavePath As String '電子檔存放路徑 Add by Morgan 2009/4/13
'Added by Lydia 2018/07/16 將原本畫面中固定特定代理人的資料,改存在DB用Grid呈現,並且只要增加記錄不需改程式
Dim ExceptList As String '固定排除的代理人(整批)
Dim ChkInList As String '不同代理人用,區隔,存:Y編號|上次通知日起|上次通知日止|下次繳費日起|下次繳費日止
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim tmpList1 As String '產生的資料檢核表
Dim strBUser As String, strBDate1 As String, strBDate2 As String 'Added by Lydia 2024/05/29 記錄管制人和上次期限

Private Sub PrintPI(ByVal strTmp As String)
   'edit by nickc 2007/02/02
   'Dim pA(1 To T_PA) As String, A1K(1 To T_1K0) As String, lTmp As Long
   Dim pa() As String, A1K() As String, lTmp As Long
   'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
   ReDim A1K(1 To TF_1K0) As String
   
   Dim strStartDay As String, strEndDay As String, varYear As Variant, s As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim lTmp1 As Long
 
On Error GoTo ErrHnd

   cnnConnection.BeginTrans

   ChgCaseNo strTmp, pa
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      '若有FC代理人判斷FA39="Y" ,若無FC代理人則用申請人判斷 CU72="Y" , 才可新增請款單資料
        '若有代理人
      If pa(75) <> "" Then
        pa(75) = Left(pa(75) & "000000000", 9)
        StrSQLa = "Select FA39 From Fagent Where FA01='" & Left(pa(75), 8) & "' And FA02='" & Mid(pa(75), 9, 1) & "' " & _
                            " And FA39 IS NOT NULL AND FA39 ='Y' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            cnnConnection.RollbackTrans
            Exit Sub
        End If
        '若有申請人
      ElseIf pa(26) <> "" Then
        pa(26) = Left(pa(26) & "000000000", 9)
        StrSQLa = "Select CU72 From Customer Where CU01='" & Left(pa(26), 8) & "' And CU02='" & Mid(pa(26), 9, 1) & "' " & _
                            " And CU72 IS NOT NULL AND CU72 ='Y' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            cnnConnection.RollbackTrans
            Exit Sub
        End If
      '若無代理人與申請人
      Else
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            cnnConnection.RollbackTrans
            Exit Sub
      End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
      
      A1K(1) = AccAutoNo(MsgText(815), 5)
      '更新流水號
      AccSaveAutoNo MsgText(815), Right(A1K(1), 5)
      A1K(2) = strSrvDate(2)
        A1K(3) = PUB_GetA1K03(pa(1), pa(2), pa(3), pa(4))
        
      If GetMoneyDate(Val(pa(8)), pa(9), pa, strStartDay, strExc(1), strEndDay) Then
         varYear = Split(strExc(1), ",")
         On Error Resume Next
         s = Format(varYear(UBound(Split(pa(72), ",")) + 1))
         If Err.Number <> 0 Then s = 0
         On Error GoTo ErrHnd
      End If
      strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(pa(9)) & " AND YF02=" & CNULL(pa(8)) & " AND " & _
         "YF03='Y00000000' AND YF04='605' AND YF05=" & s
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      A1K(9) = 0
      lTmp = 0: lTmp1 = 0
      If intI = 1 Then lTmp = RsTemp.Fields(0): lTmp1 = RsTemp.Fields(1)
        A1K(9) = lTmp1
      strExc(0) = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & strSrvDate(2) & " ORDER BY USXR01 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then A1K(10) = RsTemp.Fields(0)
        Dim strDisc As String '折扣
        strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), "605", A1K(2)) / 100)
        'A1K11要先扣除折扣才存檔
      A1K(11) = Format(Val(A1K(9)) + lTmp - Val(lTmp * Val(strDisc)))
        If A1K(10) <> "" Then
            '美金取至整數位(無條件捨去)
            A1K(8) = Fix(Val(A1K(11)) / Val(A1K(10)))
        Else
            '美金取至整數位(無條件捨去)
            A1K(8) = Fix(Val(A1K(11)))
        End If
      A1K(13) = pa(1)
      A1K(14) = pa(2)
      A1K(15) = pa(3)
      A1K(16) = pa(4)
      A1K(18) = "USD"
      A1K(19) = strSrvDate(2)
      A1K(20) = ServerTime
      A1K(21) = strUserNum
         '列印對象及請款對象皆為代理人編號
         A1K(27) = PUB_GetA1K27(pa(1), pa(2), pa(3), pa(4), "605")
        If A1K(27) = "" Then A1K(27) = A1K(3)
         A1K(28) = PUB_GetA1K28(pa(1), pa(2), pa(3), pa(4), "605")
        If A1K(28) = "" Then A1K(28) = A1K(3)
        'Modify by Morgan 2004/12/16 改規則
        'A1K(4) = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4))
         A1K(4) = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4), A1K(28), "605")
        '2004/12/16 end
        
      If SaveNew1K0(A1K) Then
      
         strExc(1) = "INSERT INTO ACC1L0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L07,A1L08,A1L09,A1L10) VALUES " & _
            "('" & A1K(1) & "','001','" & pa(1) & "','605'," & lTmp & "," & lTmp * Val(strDisc) & "," & A1K(19) & "," & A1K(20) & ",'" & A1K(21) & "')"
        '案件性質代號後加99表此案件性質的規費(規費)
         strExc(2) = "INSERT INTO ACC1L0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L07,A1L08,A1L09,A1L10) VALUES " & _
            "('" & A1K(1) & "','002','" & pa(1) & "','60599'," & lTmp1 & ",0 ," & A1K(19) & "," & A1K(20) & ",'" & A1K(21) & "')"
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.ExecSQL(2, strExc) Then
         If ClsLawExecSQL(2, strExc) Then
            
            PUB_UpdateA1k08 A1K(1) 'Added by Morgan 2012/11/2 更新請款單外幣金額
   
            '更新進度檔的CP60
            StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And NP07='" & 年費 & "' And NP06 IS NULL "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strSql = "Update CaseProgress Set CP60 ='" & A1K(1) & "' Where CP09='" & rsA("NP01").Value & "' "
                cnnConnection.Execute strSql
            End If
            
            PUB_PointAutoassign A1K(1), True 'Add by Morgan 2010/4/21 自動分配點數
            
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '新增請款單列表資料
            'Remove by Morgan 2008/4/3 印地址條時已+
            'pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewDebitNoteList strUserNum, A1K(1), "" & pub_AddressListSN, IIf(m_bolEmail, "Y", ""), IIf(m_bolPlusPaper, "Y", "")
            
            'Added by Lydia 2016/11/21 整批列印:以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則新增列印清單
            If PUB_ChkAcc225MsgList(A1K(1), A1K(28), pa(1), pa(2), pa(3), pa(4), IIf(Option1(0).Value = True, Me.Caption, "")) Then
            End If
            'end 2016/11/21
            
A0:               '列印 P/I
                    
        Else
            GoTo ErrHnd
         End If
      End If
   End If
   
   cnnConnection.CommitTrans
   Exit Sub
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strTmp As String, rsAD As New ADODB.Recordset
   Dim strTxt(1 To 3) As String
   'edit by nickc 2007/02/02
   'Dim pA(1 To T_PA) As String, i As Integer, j As Integer, A1K(1 To T_1K0) As String
   'Modified by Lydia 2024/05/29
   'Dim pa() As String, i As Integer, j As Integer, A1K() As String
   'add by nickc 2007/02/02
   'ReDim pa(1 To TF_PA) As String
   'ReDim A1K(1 To TF_1K0) As String
   Dim tmpPA() As String
   ReDim tmpPA(1 To TF_PA) As String
   'end 2024/05/29
   Dim strPA25 As String, strDate As String
   Dim strSitu As String ' 定稿處理方式
   Dim strBillNo As String '待印請款單號 Add by Morgan 2011/6/24
   Dim strPA08 As String 'Added by Morgan 2013/7/19

   m_PrintRpt1 = False
   tmpList1 = "" 'Added by Lydia 2018/07/16
   Select Case Index
      Case 0 '確定
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
        
         '下次繳費日
         If Option1(0).Value = True Then
            If Me.Text1(5).Text = "" Then
               MsgBox "請輸入下次繳費起日!!!", vbExclamation + vbOKOnly
               Me.Text1(5).SetFocus
               Text1_GotFocus 5
               Exit Sub
            End If
            If Me.Text1(6).Text = "" Then
                MsgBox "請輸入下次繳費迄日!!!", vbExclamation + vbOKOnly
               Me.Text1(6).SetFocus
               Text1_GotFocus 6
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
               Me.Text1(5).SetFocus
               Text1_GotFocus 5
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
               Me.Text1(6).SetFocus
               Text1_GotFocus 6
               Exit Sub
            End If
            
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(5) & "-" & Text1(6) 'Add By Sindy 2010/12/7
            
            'Memo by Lydia 2024/05/29 刪除不用的Code

            'Added by Morgan 2016/1/13 曾經發生錯誤後又重新執行一次
            If MsgBox("本次為整批列印，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Sub
            End If
            'end 2016/1/13
            
            Screen.MousePointer = vbHourglass
            'Memo by Lydia 2024/05/29 刪除不用的Code
            
            'Added by Lydia 2016/04/19
            'Modified by Lydia 2018/07/16 抓特定代理人字串
            'Memo by Lydia 2024/05/29 刪除不用的Code
            tmpArr1 = Empty
            tmpArr2 = Empty
            If ChkInList <> "" Then
                tmpArr1 = Split(ChkInList, ",")
                For intI = 0 To UBound(tmpArr1)
                     If Trim(tmpArr1(intI)) <> "" Then
                          tmpArr2 = Split(tmpArr1(intI), "|")
                          pub_QL05 = pub_QL05 & ";" & tmpArr2(0) & " " & tmpArr2(3) & "-" & tmpArr2(4)
                     End If
                Next intI
            End If
            'end 2018/07/16
            
            'Add by Lydia 2014/10/23 -> 排除代理人Y2062400 0
            'Memo by Lydia 2024/05/29 刪除不用的Code
            'Modified by Lydia 2018/07/16 改成固定排除代理人ExceptList
            'strExc(0) = "select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,'1' ord1 from nextprogress A, patent " & _
               " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
               " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and (PA75 is null or (substr(PA75,1,8) <> 'Y2062400' and substr(PA75,1,8) <> 'Y5133301' and substr(PA75,1,8) <> 'Y2006500' and substr(PA75,1,8) <> 'Y5285800')) " & _
               " AND NP06 IS NULL and not exists" & _
               " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
               " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
            'Modify By Sindy 2021/4/27 + ,np23
            'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
            'strExc(0) = "select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'1' ord1 from nextprogress A, patent " & _
              " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
              " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and (PA75 is null or instr('" & ExceptList & "',substr(pa75,1,8))=0) " & _
              " AND NP06 IS NULL and not exists" & _
              " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
              " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
            strExc(0) = "select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'1' ord1 from nextprogress A, patent " & _
              " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
              " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and (PA75 is null or instr('" & ExceptList & "',substr(pa75,1,8))=0) " & _
              " AND NP06 IS NULL and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
            'end 2025/10/28
            'Modified by Lydia 2018/07/16 抓特定代理人字串
            'Memo by Lydia 2024/05/29 刪除不用的Code
            If ChkInList <> "" Then
                tmpArr1 = Empty
                tmpArr2 = Empty
                tmpArr1 = Split(ChkInList, ",")
                For intI = 0 To UBound(tmpArr1)
                     If Trim(tmpArr1(intI)) <> "" Then
                        tmpArr2 = Split(tmpArr1(intI), "|")
                        'Modify By Sindy 2021/4/27 + ,np23
                        'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
                        'strExc(0) = strExc(0) & " Union select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'2' ord1 from nextprogress A, patent " & _
                           " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(tmpArr2(3), 2) & " AND " & TransDate(tmpArr2(4), 2) & _
                           " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and substr(PA75,1,8)='" & tmpArr2(0) & "' " & _
                           " AND NP06 IS NULL and not exists" & _
                           " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
                           " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
                        strExc(0) = strExc(0) & " Union select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'2' ord1 from nextprogress A, patent " & _
                           " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(tmpArr2(3), 2) & " AND " & TransDate(tmpArr2(4), 2) & _
                           " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and substr(PA75,1,8)='" & tmpArr2(0) & "' " & _
                           " AND NP06 IS NULL and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
                        'end 2025/10/28
                     End If
                Next intI
            End If 'end 2018/07/16
            
            'Modified by Lydia 2016/12/08 + ord1 分作業階段排序
            strExc(0) = strExc(0) & " ORDER BY ord1,eMail,np02,np03,np04,np05"
            'end 2016/04/19
            
            intI = 1
            Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Add by Morgan 2011/7/8
               pub_OsPrinter = PUB_GetOsDefaultPrinter
               PUB_SetOsDefaultPrinter Combo2.Text
               PUB_SetWordActivePrinter
               'end 2011/7/8
               PUB_RestorePrinter Combo2.Text 'Add By Sindy 2015/7/6
               With rsAD
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
                  Do While Not .EOF
                     intI = 1
                     strReceiveNo = .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)
         'Add by Lydia 2014/10/23 抽出共用InsertData1,傳共用參數
                     tmpPA(1) = .Fields(0): tmpPA(2) = .Fields(1): tmpPA(3) = .Fields(2): tmpPA(4) = .Fields(3) 'NP02,NP03,NP04,NP05
                     tmpPA(5) = .Fields(4): tmpPA(6) = .Fields(5) '本所期限,法定期限
                     tmpPA(7) = "" & .Fields(7): tmpPA(8) = .Fields(8): tmpPA(9) = .Fields(9) 'NP15.備註,NP01.總收文號,NP22.序號 'Add By Sindy 2015/8/18
                     tmpPA(10) = "" & .Fields("np23") 'NP23.約定期限 'Add By Sindy 2021/4/27
                     'Call InsertData1(1, tmpPA(), strTxt(), strPA25, strDate, strSitu, strBillNo, strPA08)
                     If InsertData1(1, tmpPA(), strTxt(), strPA25, strDate, strSitu, strBillNo, strPA08) = False Then Exit Do
                     .MoveNext
                  Loop
               End With
               
               PUB_SetOsDefaultPrinter pub_OsPrinter 'Add by Morgan 2011/6/24
               PUB_RestorePrinter strPrinter2 'Add By Sindy 2015/7/6
         
               MsgBox "列印結束 !" & IIf(m_PrintRpt1 = True, vbCrLf & "定稿轉PDF存卷宗區有錯誤,詳情請看(" & PUB_Getdesktop & "\" & m_strFileName1 & ")", ""), vbInformation
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/12/7
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
            
            AddPLetterList 'Added by Morgan 2015/10/20
            'Modified by Morgan 2013/5/21
            'Memo by Lydia 2024/05/29 刪除不用的Code
            PUB_SaveLastDate Me.Name, "DATE1", Text1(5).Text
            PUB_SaveLastDate Me.Name, "DATE2", Text1(6).Text
            Label2(0).Caption = PUB_GetLastDate(Me.Name, "DATE1")
            Label2(1).Caption = PUB_GetLastDate(Me.Name, "DATE2")
            'end 2013/5/21
         
            'Modified by Lydia 2018/07/16 抓特定代理人字串
            'Memo by Lydia 2024/05/29 刪除不用的Code
         
            'Added by Lydia 2018/07/16 抓特定代理人字串
            If ChkInList <> "" Then
                tmpArr1 = Empty
                tmpArr2 = Empty
                tmpArr1 = Split(ChkInList, ",")
                For intI = 0 To UBound(tmpArr1)
                     If Trim(tmpArr1(intI)) <> "" Then
                          tmpArr2 = Split(tmpArr1(intI), "|")
                          PUB_SaveLastDate Me.Name, tmpArr2(0) & "-1", Trim(tmpArr2(3))
                          PUB_SaveLastDate Me.Name, tmpArr2(0) & "-2", Trim(tmpArr2(4))
                     End If
                Next intI
            End If
            
            Screen.MousePointer = vbDefault
            
         'Add by Lydia 2014/10/23 增加代理人Y20624和整批條件選項
         '本所案號
         'Else
         ElseIf Option1(1).Value = True Then
            strTmp = Text1(1) & Text1(2)
            If Me.Text1(2).Text = "" Then
                MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
                Me.Text1(2).SetFocus
               Text1_GotFocus 2
               Exit Sub
            End If
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/12/7
            If Text1(3).Text = "" Then
               Text1(3).Text = "0"  'Add by Lydia 2014/10/23-無資料,補零
               strTmp = strTmp & "0"
            Else
               strTmp = strTmp & Text1(3).Text
               pub_QL05 = pub_QL05 & "-" & Text1(3) 'Add By Sindy 2010/12/7
            End If
            If Text1(4).Text = "" Then
               Text1(4).Text = "00"  'Add by Lydia 2014/10/23-無資料,補零
               strTmp = strTmp & "00"
            Else
               strTmp = strTmp & Text1(4).Text
               pub_QL05 = pub_QL05 & "-" & Text1(4) 'Add By Sindy 2010/12/7
            End If
            Screen.MousePointer = vbHourglass
            strReceiveNo = strTmp
'Add by Lydia 2014/10/23 抽出共用
            '傳共用參數
            tmpPA(1) = Text1(1): tmpPA(2) = Text1(2): tmpPA(3) = Text1(3): tmpPA(4) = Text1(4)         '本所案號
            'Modify By Sindy 2015/7/6 +,GetEmailFlag(np02||np03||np04||np05) eMail,np15
            'Modify By Sindy 2021/4/27 + ,np23
            strExc(0) = "SELECT NVL(np08,'') as NP08,NVL(np09,'') as NP09,GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23 FROM NEXTPROGRESS WHERE " & ChgNextProgress(strReceiveNo) & " AND NP07=" & 年費 & _
                       " AND NP06 IS NULL ORDER BY NP08,NP09"
                       
            intI = 1
            Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Add by Morgan 2011/7/8
               pub_OsPrinter = PUB_GetOsDefaultPrinter
               PUB_SetOsDefaultPrinter Combo2.Text
               PUB_SetWordActivePrinter
               'end 2011/7/8
               PUB_RestorePrinter Combo2.Text 'Add By Sindy 2015/7/6
               
               tmpPA(5) = rsAD.Fields("NP08"): tmpPA(6) = rsAD.Fields("NP09")   '本所期限,法定期限
               tmpPA(7) = "" & rsAD.Fields("np15"): tmpPA(8) = rsAD.Fields("np01"): tmpPA(9) = rsAD.Fields("np22") 'NP15.備註,NP01.總收文號,NP22.序號 'Add By Sindy 2015/8/18
               tmpPA(10) = "" & rsAD.Fields("np23") 'NP23.約定期限 'Add By Sindy 2021/4/27
               Call InsertData1(2, tmpPA(), strTxt(), strPA25, strDate, strSitu, strBillNo, strPA08)
               
               PUB_SetOsDefaultPrinter pub_OsPrinter 'Add by Morgan 2011/6/24
               PUB_RestorePrinter strPrinter2 'Add By Sindy 2015/7/6
            Else
                InsertQueryLog (0)
                MsgBox "無符合條件之資料可列印 !", vbInformation
            End If

            Screen.MousePointer = vbDefault
            
            Set rsAD = Nothing 'Added by Lydia 2024/05/29

         End If
      Case 1 '結束
            Me.Enabled = False
            Unload Me
   End Select
   
End Sub

'Added by Lydia 2024/05/29 拆分給各區管制人
Private Sub cmdok2_Click()
Dim intErr As Integer
Dim rsAD As New ADODB.Recordset
Dim strTxt(1 To 3) As String
Dim tmpPA() As String
ReDim tmpPA(1 To TF_PA) As String
Dim strPA25 As String, strDate As String
Dim strSitu As String ' 定稿處理方式
Dim strBillNo As String '待印請款單號
Dim strPA08 As String

   intErr = -1
   If Option1(2).Value = True Then
      If Trim(txtData(0)) = "" Then
         MsgBox "請輸入整批期限起日!!!", vbExclamation + vbOKOnly
         intErr = 0
         GoTo ErrHandTag
      End If
      If Trim(txtData(1)) = "" Then
         MsgBox "請輸入整批期限止日!!!", vbExclamation + vbOKOnly
         intErr = 1
         GoTo ErrHandTag
      End If
      If Trim(txtData(0)) > Trim(txtData(1)) Then
         MsgBox "整批期限起日不可大於止日!!!", vbExclamation + vbOKOnly
         intErr = 0
         GoTo ErrHandTag
      End If
      If Trim(txtData(2)) = "" Or lblFM2.Caption = "" Or strBUser <> Trim(txtData(2)) Then
         MsgBox "請輸入管制人!!!", vbExclamation + vbOKOnly
         intErr = 2
         GoTo ErrHandTag
      End If
      'Added by Morgan 2016/1/13 曾經發生錯誤後又重新執行一次
      If MsgBox("本次為整批列印，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      'end 2016/1/13
   ElseIf Option1(1).Value = True Then
      If Len(Text1(2)) < 6 Then
         MsgBox "本所案號不可空白，請重新輸入 !", vbCritical + vbOKOnly
         Text1(2).SetFocus
         Text1_GotFocus 2
         Exit Sub
      End If
   Else
      MsgBox "請選擇整批／個案！", vbExclamation + vbOKOnly
      Exit Sub
   End If

   m_PrintRpt1 = False
   tmpList1 = ""
   ClearQueryLog (Me.Name)
   If Option1(2).Value = True Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txtData(0) & "-" & txtData(1) & ";"
      '抓特定代理人字串
      tmpArr1 = Empty
      tmpArr2 = Empty
      If ChkInList <> "" Then
         strExc(6) = ""
         tmpArr1 = Split(ChkInList, ",")
         For intI = 0 To UBound(tmpArr1)
            If Trim(tmpArr1(intI)) <> "" Then
               tmpArr2 = Split(tmpArr1(intI), "|")
               '依畫面輸入的期限，重新計算下次執行期限
               If Val(UBound(tmpArr2)) < 10 Then
                  '一般提早3個月催期限，特殊代理人提前Ｘ個月包含原本的3個月
                  strExc(3) = CompDate(1, Val(tmpArr2(UBound(tmpArr2))) - 3, TransDate(Left(txtData(0), 5) + "01", 2))
                  strExc(4) = CompDate(2, -1, CompDate(1, 1, strExc(3)))
                  tmpArr2(3) = TransDate(strExc(3), 1)
                  tmpArr2(4) = TransDate(strExc(4), 1)
                  strExc(6) = strExc(6) & "," & tmpArr2(0) & "|" & tmpArr2(1) & "|" & tmpArr2(2) & "|" & tmpArr2(3) & "|" & tmpArr2(4)
               End If
               pub_QL05 = pub_QL05 & ";" & tmpArr2(0) & " " & tmpArr2(3) & "-" & tmpArr2(4)
            End If
         Next intI
         ChkInList = Mid(strExc(6), 2)
      End If
      pub_QL05 = pub_QL05 & ";" & Label5.Caption & Trim(txtData(2))
      
      'Modified by Lydia 2024/06/04 debug: and (PA75 is null or instr('" & ExceptList & "',substr(pa75,1,8))=0) => 改為IIf(ExceptList <> "", " and instr('" & ExceptList & "',substr(pa75,1,8))=0", "")
      'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
      'strExc(0) = "select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'1' ord1 from nextprogress A, patent,fagent,nation " & _
            " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(txtData(0).Text, 2) & " AND " & TransDate(txtData(1).Text, 2) & _
            " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & IIf(ExceptList <> "", " and instr('" & ExceptList & "',substr(pa75,1,8))=0", "") & _
            " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na16='" & Trim(txtData(2)) & "' AND NP06 IS NULL and not exists" & _
            " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
            " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
      strExc(0) = "select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'1' ord1 from nextprogress A, patent,fagent,nation " & _
            " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(txtData(0).Text, 2) & " AND " & TransDate(txtData(1).Text, 2) & _
            " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & IIf(ExceptList <> "", " and instr('" & ExceptList & "',substr(pa75,1,8))=0", "") & _
            " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na16='" & Trim(txtData(2)) & "'" & _
            " AND NP06 IS NULL and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
      'end 2025/10/28
      '抓特定代理人字串
      If ChkInList <> "" Then
         tmpArr1 = Empty
         tmpArr2 = Empty
         tmpArr1 = Split(ChkInList, ",")
         For intI = 0 To UBound(tmpArr1)
            If Trim(tmpArr1(intI)) <> "" Then
               tmpArr2 = Split(tmpArr1(intI), "|")
               'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
               'strExc(0) = strExc(0) & " Union select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'2' ord1 from nextprogress A, patent,fagent,nation " & _
                  " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(tmpArr2(3), 2) & " AND " & TransDate(tmpArr2(4), 2) & _
                  " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and substr(PA75,1,8)='" & tmpArr2(0) & "' " & _
                  " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na16='" & Trim(txtData(2)) & "' AND NP06 IS NULL and not exists" & _
                  " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
                  " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
               strExc(0) = strExc(0) & " Union select np02,np03,np04,np05,NVL(np08,''),NVL(np09,''),GetEmailFlag(np02||np03||np04||np05) eMail,np15,np01,np22,np23,'2' ord1 from nextprogress A, patent,fagent,nation " & _
                  " WHERE NP02='FCP' and NP07=" & 年費 & " AND NP09 BETWEEN " & TransDate(tmpArr2(3), 2) & " AND " & TransDate(tmpArr2(4), 2) & _
                  " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and substr(PA75,1,8)='" & tmpArr2(0) & "' " & _
                  " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na16='" & Trim(txtData(2)) & "'" & _
                  " AND NP06 IS NULL and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
               'end 2025/10/28
            End If
         Next intI
      End If
      strExc(0) = strExc(0) & " ORDER BY ord1,eMail,np02,np03,np04,np05"
      
      intI = 1
      Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter Combo2.Text
         PUB_SetWordActivePrinter
         PUB_RestorePrinter Combo2.Text
         With rsAD
            InsertQueryLog (.RecordCount)
            Do While Not .EOF
               intI = 1
               strReceiveNo = .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)
               '抽出共用InsertData1,傳共用參數
               tmpPA(1) = .Fields(0): tmpPA(2) = .Fields(1): tmpPA(3) = .Fields(2): tmpPA(4) = .Fields(3) 'NP02,NP03,NP04,NP05
               tmpPA(5) = .Fields(4): tmpPA(6) = .Fields(5) '本所期限,法定期限
               tmpPA(7) = "" & .Fields(7): tmpPA(8) = .Fields(8): tmpPA(9) = .Fields(9) 'NP15.備註,NP01.總收文號,NP22.序號
               tmpPA(10) = "" & .Fields("np23")
               If InsertData1(1, tmpPA(), strTxt(), strPA25, strDate, strSitu, strBillNo, strPA08) = False Then Exit Do
               .MoveNext
            Loop
         End With
         
         PUB_SetOsDefaultPrinter pub_OsPrinter
         PUB_RestorePrinter strPrinter2
         MsgBox "列印結束 !" & IIf(m_PrintRpt1 = True, vbCrLf & "定稿轉PDF存卷宗區有錯誤,詳情請看(" & PUB_Getdesktop & "\" & m_strFileName1 & ")", ""), vbInformation
      Else
         InsertQueryLog (0)
         MsgBox "無符合條件之資料可列印 !", vbInformation
      End If

      Set rsAD = Nothing
      AddPLetterList
      Call SetDefDate("U")
      If ChkInList <> "" Then
         Call QueryData
      End If
      
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   ElseIf Option1(1).Value = True Then
      Call cmdok_Click(0)
   End If
   
   Exit Sub
   
ErrHandTag:
   If intErr >= 0 Then
      txtData(intErr).SetFocus
      Txtdata_GotFocus intErr
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer
   Dim tmpYY As String, tmpMM As String, tmpDay As Integer, tmpYM01 As String
   
    MoveFormToCenter Me
    intWhere = 國外_FC
    Option1_Click 0
    'Modified by Morgan 2013/5/21
    'Label2(0).Caption = GetSetting("TAIE", "FCP", "DATE1", "")
    'Label2(1).Caption = GetSetting("TAIE", "FCP", "DATE2", "")
    Label2(0).Caption = PUB_GetLastDate(Me.Name, "DATE1")
    Label2(1).Caption = PUB_GetLastDate(Me.Name, "DATE2")
    'end 2013/5/21
    
    'Add by Lydia 2014/10/23 增加代理人Y20624和整批條件選項
    'Remove by Lydia 2018/07/16 改成Grid '----Memo by Lydia 2024/05/29 刪除不用的Code
'Modified by Lydia 2024/05/29 拆分給各區管制人
'    Erase mDate
'    ReDim mDate(0 To 3)
'
'    '整批系統期限
'    'Modified by Lydia 2015/05/14 改接上次通知期限
'    'tmpYM01 = CompDate(1, 5, mSrvDate) '系統日期 + 5個月
'    tmpYM01 = CompDate(2, 1, DBDATE(Label2(1).Caption))
'
'    tmpYY = Mid(tmpYM01, 1, 4) '期限年
'    tmpMM = Mid(tmpYM01, 5, 2) '期限月
'    tmpDay = Mid(tmpYM01, 7, 2) '系統-日(數值)
'    mDate(0) = tmpYY & tmpMM   'YYYYMM
'    mDate(1) = GetLastDay(tmpYM01)
'    If tmpDay < 11 Then
'       Option2(0).Value = True
'       Option2_Click 0
'    ElseIf tmpDay < 21 Then
'       Option2(1).Value = True
'       Option2_Click 1
'    Else
'       Option2(2).Value = True
'       Option2_Click 2
'    End If
''---------------------------------------
   Dim tmpBol As Boolean
   Dim oObj As Object
   Option1(2).Value = True
   txtData(2) = strUserNum

   Call QueryData
   For Each oObj In txtData
      oObj.Tag = oObj.Text
   Next
   Call Txtdata_Validate(2, tmpBol)
'end 2024/05/29

    '設定印表機
'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
'end 2011/3/15

   MsgBox "本程式已改為直接列印定稿，請先選定印表機並放好定稿紙！", vbExclamation
   
   If Option1(0).Value = True Then 'Added by Lydia 2024/05/28
      strExc(6) = Right(Label2(0).Caption, 2) '上期起-日
      strExc(7) = Right(Label2(1).Caption, 2) '上期迄-日
      If (strExc(6) <> "01" And strExc(6) <> "11" And strExc(6) <> "21") Or (Val(strExc(7)) < 28 And Trim(Val(strExc(7)) + 1) <> Right(Text1(5).Text, 2)) Then
          MsgBox "請注意上次通知日與下次繳費日的銜接！"
      End If
   End If
    'end 2014/10/23
           
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印定稿整批列印清單
   'Modify By Sindy 2015/7/6 E化與非E化清單分開列印
   PUB_PrintLetterList strUserNum, "5", Combo2, strPrinter2, "and LL09 in('Ｅ','ｅ')" '"and GetEmailFlag(LL04||LL05||LL06||LL07) in('E','e')"
   PUB_PrintLetterList strUserNum, "5", Combo2, strPrinter2, "and LL09 is null", False '"and GetEmailFlag(LL04||LL05||LL06||LL07) is null", False
   '2015/7/6 END
   PUB_PrintLetterList strUserNum, "9", Combo2, strPrinter2, , False 'Added by Morgan 2015/10/20
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, "and LL02 like '%繳年費通知函' "

   '列印請款單
   PUB_PrintDebitNote strUserNum, Me.Combo2.Text
   '刪除請款單列表資料
   PUB_DeleteDebitNoteList strUserNum
   
   'Added by Lydia 2016/11/21
   '列印:國外固定寄催款單清單
   PUB_PrintAcc225List strUserNum, Me.Combo2.Text
   '刪除:國外固定寄催款單清單
   PUB_DeleteAcc225List strUserNum
   'end 2016/11/21
   
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2004/10/26 end

   Set frm060304 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   For Each txt In Text1
      txt.Enabled = False
   Next
   'Add by Lydia 2014/10/23 封鎖option2
   Frame3.Enabled = False
   
   Select Case Index
      Case 0
         Frame3.Enabled = True
         'Modified by Lydia 2016/04/19 下次繳費日期不可變更
         'Text1(5).Enabled = True
         'Text1(6).Enabled = True
      Case 1
         For i = 2 To 4 '鎖定FCP案
            Text1(i).Enabled = True
         Next
         Call QueryData  'Added by Lydia 2018/07/16
         
      'Remove by Lydia 2016/04/19 併入整批的1-10號
      ''Add by Lydia 2014/10/23
      'Case 2 '代理人Y20624
      '   Text1(7).Enabled = True
      '   Text1(8).Enabled = True
      '   Text1(7).Text = ChangeWStringToTString(mDate(2))
      '   Text1(8).Text = ChangeWStringToTString(mDate(3))
   End Select
End Sub

'Add by Lydia 2014/10/23 增加代理人Y20624和整批條件選項
Private Sub Option2_Click(Index As Integer)

On Error Resume Next

Chk1(0).Value = 0: Chk1(1).Value = 0 'Added by Lydia 2016/04/19
Chk1(2).Value = 0 'Added by Lydia 2016/11/04

Select Case Index
       Case 0   '1~10號
            'Added by Lydia 2016/04/19 特定代理人併入整批的1~10號
            Chk1(0).Value = 1: Chk1(1).Value = 1
            Chk1(2).Value = 1 'Added by Lydia 2016/11/04
            
            Text1(5).Text = ChangeWStringToTString(mDate(0) & "01")
            Text1(6).Text = ChangeWStringToTString(mDate(0) & "10")
       Case 1   '11~20號
            Text1(5).Text = ChangeWStringToTString(mDate(0) & "11")
            Text1(6).Text = ChangeWStringToTString(mDate(0) & "20")
       Case 2   '21~30號
            Text1(5).Text = ChangeWStringToTString(mDate(0) & "21")
            Text1(6).Text = ChangeWStringToTString(mDate(1))

End Select

Call QueryData  'Added by Lydia 2018/07/16 改Grid顯示

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
 Dim Cancel As Boolean
   Cancel = False
   If Option1(0).Value = True Then
      If Index = 6 Then
         If Not ChkRange(Text1(5), Text1(6), "下次繳費日") Then
         End If
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Option1(0).Value = True Then
      Select Case Index
         Case 5, 6
            If Text1(Index) <> "" Then
               If Not ChkDate(Text1(Index)) Then
                  Cancel = True
               End If
            End If
      End Select
   Else
      If Index = 1 Then
         If Text1(Index).Text <> "FCP" Then
            MsgBox "系統別必需為 FCP，請重新輸入 !", vbCritical
            Cancel = True
         End If
      ElseIf Index = 2 Then
         If Text1(Index).Text = "" Then
            MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   End If
   If Cancel Then TextInverse Text1(Index)
End Sub
'Add by Morgan 2006/5/4
'取得定稿處理方式--新規則
Private Function GetSitu(ByVal p_PA01 As String, ByVal p_PA02 As String, ByVal p_PA03 As String, ByVal p_PA04 As String) As String
   Dim stLanguage As String, stAutoPay As String, stReceiver As String
   
   p_PA03 = Right("0" & p_PA03, 1)
   p_PA04 = Right("00" & p_PA04, 2)
   
'Removed by Morgan 2017/4/26 發明已閉卷,改用一般定稿--David
'   'Added by Morgan 2013/8/22 FCP046754(定稿特別)--David
'   If p_PA01 & p_PA02 & p_PA03 & p_PA04 = "FCP046754000" Then
'      GetSitu = "07"
'      Exit Function
'   End If
'   'end 2013/8/22
'end 2017/4/26
                        
   '先抓基本檔設定
   strSql = "SELECT PA85,PA70 FROM PATENT WHERE PA01='" & p_PA01 & "' AND PA02='" & p_PA02 & "' AND PA03='" & p_PA03 & "' AND PA04='" & p_PA04 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      stLanguage = "" & RsTemp.Fields(0)
      'stAutoPay = "" & RsTemp.Fields(1) 'Removed by Morgan 2025/8/28 改在下面用函數設定
   End If
   '再抓代理人/客戶檔設定
   If stLanguage = "" Or stAutoPay = "" Then
      stReceiver = PUB_GetReceiver(p_PA01, p_PA02, p_PA03, p_PA04, "605", "1")
      If Left(stReceiver, 1) = "Y" Then
         strSql = "select FA31,FA41 from Fagent where fa01||fa02='" & stReceiver & "'"
      Else
         strSql = "select CU64,CU74 from Customer where cu01||cu02='" & stReceiver & "'"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
      If intI = 1 Then
         If stLanguage = "" Then stLanguage = "" & RsTemp.Fields(0)
         'If stAutoPay = "" Then stAutoPay = "" & RsTemp.Fields(1) 'Removed by Morgan 2025/8/28 改在下面用函數設定
      End If
   End If
   
   stAutoPay = PUB_GetAutoPay(p_PA01, p_PA02, p_PA03, p_PA04) 'Added by Morgan 2025/8/28
   
   Select Case stLanguage
      Case "1" '中文
         GetSitu = "01"
      Case "3" '日文
         '自動代繳
         If stAutoPay = "Y" Then
            GetSitu = "05"
         Else
            GetSitu = "04"
         End If
      Case Else '預設英文
         'Modified by Morgan 2025/8/28 英文定稿合併
'         '自動代繳
'         If stAutoPay = "Y" Then
'            GetSitu = "03"
'         Else
'            GetSitu = "02"
'            'Add by Morgan 2007/4/14 Nikon(Y45148)定稿特別
'            If Left(stReceiver, 6) = "Y45148" Then
'               GetSitu = "06"
'            End If
'            'end 2007/4/14
'         End If
         GetSitu = "02"
         'end 2025/8/28
   End Select
   
End Function

'取得繳年費相關費用資料
'Modify By Sindy 2015/7/6 +, ByRef strYear As String
Private Sub GetPatentYearFee(ByVal ET01 As String, ByVal strReceiveNo As String, ByVal strSitu As String, ByRef strYear As String)
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim arrPA72
   'Dim strYear As String
   Dim dblService As Double, dblFee As Double
   Dim bDisc As Boolean
   
   strYear = ""
   StrSQLa = "Select * From Patent Where " & ChgPatent(strReceiveNo)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Add by Morgan 2010/5/4
      'Modified by Morgan 2012/8/20
      'If rsA("pa75") = "Y52216000" Then
      If Left(rsA("pa26"), 6) = "X47047" Then
      
         m_ii = m_ii + 1
         m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & strSitu & "','" & strUserNum & "','報價不印','♀')"
      End If
      'end 2010/5/4
   
      '若有繳年費年度資料
      If "" & rsA("PA72").Value <> "" Then
         arrPA72 = Split(rsA("PA72").Value, ",")
         strYear = arrPA72(UBound(arrPA72))
         '抓下一次的繳年費資料
         strYear = Val("0" & strYear) + 1
         '若有最近繳費年度
         If strYear <> "" Then
            bDisc = PUB_GetFCPCaseDiscState(strReceiveNo) 'Add by Morgan 2008/4/17
            
            dblService = PUB_GetYF06("" & rsA("PA09").Value, "" & rsA("PA08").Value, "Y00000000", 年費, strYear, strYear)
            m_ii = m_ii + 1
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & strSitu & "','" & strUserNum & "','服務費','" & dblService & "')"
            
            dblFee = PUB_GetYF07("" & rsA("PA09").Value, "" & rsA("PA08").Value, "Y00000000", 年費, strYear, strYear)
            'Add by Morgan 2008/4/17
            '加判斷是否有減免
            If bDisc = True Then
               If Val(strYear) < 4 Then
                  dblFee = dblFee - 800
               ElseIf Val(strYear) < 7 Then
                  dblFee = dblFee - 1200
               End If
            End If
            
            m_ii = m_ii + 1
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & strSitu & "','" & strUserNum & "','規費','" & dblFee & "')"
            m_ii = m_ii + 1
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & strSitu & "','" & strUserNum & "','年費','" & dblService + dblFee & "')"
            m_ii = m_ii + 1
            '美金取整數位(無條件捨去)
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & strSitu & "','" & strUserNum & "','費用','" & Fix(Format((dblService + dblFee) / PUB_GetUSXRate)) & "')"
         End If
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'Add by Lydia 2014/10/23 增加代理人Y20624和整批條件選項 => 抽出共用
Private Function InsertData1(ByVal mKind As Integer, ByRef m_Pa() As String, ByRef mstrTxt() As String, ByRef mstrPA25 As String, _
                             ByRef mstrDate As String, ByRef mstrSitu As String, ByRef mstrBillNo As String, _
                             ByRef mstrPA08 As String) As Boolean
'mKind = 1.整批, 2.本所案號
'mstrTxt()= 1-申請人1, 2-FC代理人, 3-年費代理人
Dim strYear As String 'Add By Sindy 2015/7/6
'Add By Sindy 2015/10/30
Dim strLD03 As String
Dim strFileName As String, strFullFileName As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim strMsg As String
Dim strNewCP09 As String
'2015/10/30 END
Dim strPA10 As String 'Added by Morgan 2019/8/5

    InsertData1 = True
    'Modified by Morgan 2005/4/21 加年費代理人PA76
    'Modified by Morgan 2019/8/5 +pa10
    strExc(0) = "SELECT '',PA26,PA75,PA76,PA25,pa14+10000*(1+lastyear(pa72)) dt,pa08,pa10 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & _
       " AND (PA57<>'Y' OR PA57 IS NULL)"
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       Erase mstrTxt
       If Not IsNull(RsTemp.Fields(1)) Then mstrTxt(1) = RsTemp.Fields(1) '申請人1
       If Not IsNull(RsTemp.Fields(2)) Then mstrTxt(2) = RsTemp.Fields(2) 'FC代理人
       
       'Add by Morgan 2010/3/12
       mstrPA25 = "" & RsTemp.Fields("pa25")
       mstrDate = "" & RsTemp.Fields("dt")
       'end 2010/3/12
       
       mstrPA08 = "" & RsTemp.Fields("pa08") 'Added by Morgan 2013/7/19
       
       'Add by Morgan 2005/4/21
       mstrTxt(3) = "" & RsTemp.Fields("PA76") '年費代理人
       
       strPA10 = "" & RsTemp.Fields("pa10") 'Added by Morgan 2019/8/5
       
       '先取得處理狀況
       mstrSitu = GetSitu(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))
       
       EndLetter ET01, strReceiveNo & "&605", mstrSitu, strUserNum
       EndLetter "04", strReceiveNo & "&605", "98", strUserNum 'Add By Sindy 2015/7/21
       'Modify by Morgan 2005/4/21
       'If Not CU73FA40(mstrTxt(1), mstrTxt(2)) Then
       If Not CU73FA40(mstrTxt(1), mstrTxt(2), mstrTxt(3)) Then
          'Add By Sindy 2015/8/18
          'Modified by Lydia 2019/08/16 加入:FCP程序大項工作整批發文
          'If PUB_AddCP1913(m_PA(1), m_PA(2), m_PA(3), m_PA(4), m_PA(5), m_PA(6), m_PA(8), m_PA(9), , , strNewCP09) = False Then
          If PUB_AddCP1913(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), m_Pa(5), m_Pa(6), m_Pa(8), m_Pa(9), , , strNewCP09, , True, , , Me.Name) = False Then
             MsgBox m_Pa(1) & "-" & m_Pa(2) & "-" & m_Pa(3) & "-" & m_Pa(4) & "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
             InsertData1 = False
             Exit Function
          End If
          '2015/8/18 END
          
          m_LetterLanguage = PUB_GetLanguage(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) 'Add By Sindy 2015/9/21
          
          m_ii = 0
          'Add By Sindy 2015/7/21
          m_ii = m_ii + 1
          m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('04','" & strReceiveNo & "&605" & "','98','" & strUserNum & "','傳真頁數','2')"
          '2015/7/21 END
          m_ii = m_ii + 1
          'Modify By Sindy 2021/4/27
          If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','約定期限'," & CNULL(m_Pa(10)) & ")"
          Else
          '2021/4/27 END
            m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','本所期限'," & CNULL(m_Pa(5)) & ")"
          End If
          m_ii = m_ii + 1
          m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','法定期限'," & CNULL(m_Pa(6)) & ")"
          
          'Added by Morgan 2019/8/5
          '108.11.1新法設計案專用期由12年延長為15年
          'Removed by Morgan 2019/8/16 改寫共用例外欄位
          'If mstrPA08 = "3" Then
          '  '專用期更新前特殊控制,更新後可改都抓專用期止日
          '  If strSrvDate(1) < 20191101 Then
          '     strExc(1) = CompDate(2, -1, CompDate(0, 15, strPA10))
          '  Else
          '     strExc(1) = mstrPA25
          '  End If
          '  m_ii = m_ii + 1
          '  m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
          '     "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','設計案15年屆滿日','" & strExc(1) & "')"
          'End If
          'end 2019/8/16
          'end2019/8/5
          
          'Add by Morgan 2010/3/12
          '最後一年(下次年費起用日>專用期止日)
          If Val(mstrPA25) > 0 And Val(mstrDate) > Val(mstrPA25) Then
             m_ii = m_ii + 1
             m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','最後一年要印','♀')"
             
             m_ii = m_ii + 1
             m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','最後一年不印','♀')"
                
             'Added by Morgan 2019/8/5
            '108.11.1新法設計案專用期由12年延長為15年(專用期更新前特殊控制,更新後可連同定稿內例外欄位一併移除)
            If mstrPA08 = "3" Then
               If strSrvDate(1) < 20191101 Then
                  m_ii = m_ii + 1
                  m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','108/10/31前設計案最後一年要印','♀')"
                  m_ii = m_ii + 1
                  m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','108/10/31前設計案最後一年不印','♀')"
               End If
            End If
            'end 2019/8/5
          End If
          
          'Added by Morgan 2013/7/19
          '一案兩請提醒
          If mstrPA08 = "2" Then
             'Modified by Morgan 2017/1/24 +判斷發明無證書號才帶(+ and pa22 is null) --David
             strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
                " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_Pa(1) & "' and cm02='" & m_Pa(2) & "' and cm03='" & m_Pa(3) & "' and cm04='" & m_Pa(4) & "'" & _
                " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_Pa(1) & "' and cm06='" & m_Pa(2) & "' and cm07='" & m_Pa(3) & "' and cm08='" & m_Pa(4) & "') X" & _
                ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and pa22 is null"
             intI = 1
             Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                m_ii = m_ii + 1
                m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','一案兩請新型案要印','♀')"
                m_ii = m_ii + 1
                m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','發明案申請號','" & adoRecordset("pa11") & "')"
                m_ii = m_ii + 1
                m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','發明案彼所案號','" & IIf(IsNull(adoRecordset("pa77")), "", "" & adoRecordset("pa77")) & "')"
                m_ii = m_ii + 1
                m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','發明案本所案號','" & adoRecordset("CNo") & "')"
             End If
          End If
          'end 2013/7/19
          
          'Add by Morgan 2011/6/24
          If PUB_GetUnPaidBill(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), mstrBillNo) = True Then
             m_ii = m_ii + 1
             m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','有欠款才印','♀')"
             m_ii = m_ii + 1
             m_strTxt(m_ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                "('" & ET01 & "','" & strReceiveNo & "&605" & "','" & mstrSitu & "','" & strUserNum & "','有欠款不印','♀')"
          End If
          
          '取得年費相關資料
          GetPatentYearFee ET01, strReceiveNo, mstrSitu, strYear
          
          'edit by nickc 2007/02/05 不用 dll 了
          'If Not objLawDll.ExecSQL(m_ii, m_mstrTxt) Then
          If Not ClsLawExecSQL(m_ii, m_strTxt) Then
             MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
             InsertData1 = False
             Exit Function
          End If
          
          'Add By Sindy 2015/11/4 非整批才要列印承辦單
          'Modified by Lydia 2016/04/19 代理人也不印承辦單
          'If Option1(0).Value = False Then
          If Option1(1).Value = True Then
          '2015/11/4 END
            'Add By Sindy 2015/7/6 列印承辦單
            'Modified by Lydia 2019/03/04 更換類別代號;
            'Call PUB_PrintFCPEmpBill(m_PA(1), m_PA(2), m_PA(3), m_PA(4), ET01, , , m_PA(6), strYear & ";" & m_PA(7))
            Call PUB_PrintFCPEmpBill(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), "04", , , m_Pa(6), strYear & ";" & m_Pa(7))
          End If
          
          '新增地址條列表資料 '整批+Y20624
          If mKind = 1 Then pub_AddressListSN = pub_AddressListSN + 1 '請款單清單會用
          
          'Add by Morgan 2008/3/26 判斷是否產生電子檔
          m_bolEmail = PUB_GetEMailFlag(strReceiveNo, True, , m_bolPlusPaper)
          'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
          If m_bolPlusPaper Then
             m_iCopy = 0
          Else
             m_iCopy = 1
          End If
          'end 2009/10/20

          If m_bolEmail Then
             NowPrint strReceiveNo & "&605", ET01, mstrSitu, False, strUserNum, , , , , m_iCopy, , True, True
             'Modify By Sindy 2015/10/30 Mark
'             '本所案號
'             If mKind = 2 Then MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(Text1(1).Text) & " ]！"
          Else
          'end 2008/3/26
             'Add By Sindy 2015/7/6 下一程序備註有出現傳真或FAX,加印Cover Page
             If Trim(m_Pa(7)) <> "" And (InStr(m_Pa(7), "傳真") > 0 Or InStr(UCase(m_Pa(7)), UCase("fax")) > 0) Then
               '加英文傳真封面
               NowPrint strReceiveNo & "&605", "04", "98", False, strUserNum, , , , , 1
             End If
             '2015/7/6 END
             NowPrint strReceiveNo & "&605", ET01, mstrSitu, False, strUserNum
          End If

          'Add by Morgan 2011/6/24
          '列印請款單
          If mstrBillNo <> "" Then
            'Modify By Sindy 2017/4/19 呼叫請款單列印因產生PDF很容易當掉會出現無法預期的錯誤,而影響到整批作業
            '改只有紙本的才需要列印請款單
            If m_bolEmail = False Then
            '2017/4/19 END
               PUB_PrintBill mstrBillNo, Combo2.Text, m_bolEmail, m_bolPlusPaper, Me.Name, , 1
            End If
          End If

          '列印通知函
'          PUB_PrintLetter strReceiveNo & "&605"
          'end 2011/6/24
          'Modify By Sindy 2015/10/30 定稿轉PDF存卷宗區
          strFileName = m_Pa(1) & m_Pa(2) & IIf(m_Pa(4) <> "00", "-" & m_Pa(3) & "-" & m_Pa(4), IIf(m_Pa(3) <> "0", "-" & m_Pa(3), "")) & ".1913.CUS.PDF"
          PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
          strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
          cnnConnection.Execute strSql
          If PUB_PrintLetter(strReceiveNo & "&605", , , True, strFullFileName) = True Then
            Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續; ex. 9/16 整批列印,因為沒有檔案才發生上傳錯誤
            Set oFile = oFileSys.GetFile(strFullFileName)
            If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
                'Added by Lydia 2016/04/19 +檔名開頭
                'Modified by Lydia 2018/07/16 抓特定代理人字串
'                Select Case Left(mstrTxt(2), 8)
'                    Case "Y2062400": strExc(1) = Me.Caption & Me.Text1(7).Text & "-" & Me.Text1(8).Text & "Y2062400資料檢核表.txt"
'                    Case "Y5133301": strExc(1) = Me.Caption & Me.Text1(9).Text & "-" & Me.Text1(10).Text & "Y5133301資料檢核表.txt"
'                    Case "Y2006500": strExc(1) = Me.Caption & Me.Text1(11).Text & "-" & Me.Text1(12).Text & "Y2006500資料檢核表.txt" 'Added by Lydia 2016/11/04
'                    Case "Y5285800": strExc(1) = Me.Caption & Me.Text1(11).Text & "-" & Me.Text1(12).Text & "Y5285800資料檢核表.txt" 'Added by Lydia 2017/12/18
'                    Case Else:   strExc(1) = Me.Caption & Me.Text1(5).Text & "-" & Me.Text1(6).Text & "資料檢核表.txt"
'                End Select
                If ChkInList <> "" And InStr(ChkInList, Left(mstrTxt(2), 8)) > 0 Then
                    tmpArr1 = Empty
                    tmpArr2 = Empty
                    tmpArr1 = Split(ChkInList, ",")
                    If tmpList1 = "" Or (tmpList1 <> "" And InStr(tmpList1, Left(mstrTxt(2), 8)) = 0) Then
                        m_PrintRpt1 = False '建立檔案
                        tmpList1 = tmpList1 & "," & Left(mstrTxt(2), 8)
                    End If
                    For intI = 0 To UBound(tmpArr1)
                         If Trim(tmpArr1(intI)) <> "" And Mid(tmpArr1(intI), 1, 8) = Left(mstrTxt(2), 8) Then
                              tmpArr2 = Split(tmpArr1(intI), "|")
                              strExc(1) = Me.Caption & tmpArr2(3) & "-" & tmpArr2(4) & tmpArr2(0)
                              Exit For
                         End If
                    Next intI
                Else
                    'Modified by Lydia 2024/05/29
                    'strExc(1) = Me.Caption & Me.Text1(5).Text & "-" & Me.Text1(6).Text
                    strExc(1) = Me.Caption & IIf(Option1(2).Value = 1, Me.Text1(5).Text & "-" & Me.Text1(6).Text, Me.txtData(0) & "-" & Me.txtData(1))
                End If
                'end 2018/07/16
               'Modified by Lydia 2022/10/31 +& ";" & strMsg
               Call ReadTxt1(m_Pa(1) & "-" & m_Pa(2) & "-" & m_Pa(3) & "-" & m_Pa(4), strNewCP09, "定稿轉PDF失敗", strExc(1) & ";" & strMsg) 'Memo by Lydia 2018/07/16 儲存-資料檢核表
               'end 2016/04/19
            End If
            Kill strFullFileName
          End If
          '2015/10/30 END

          If Not m_bolEmail Or m_bolPlusPaper Then
'            'Add By Sindy 2015/9/21 日文定稿才要印地址條
'            If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'            '2015/9/21 END
               '新增地址條列表資料  '本所案號
               If mKind = 2 Then pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, "" & m_Pa(1), "" & m_Pa(2), "" & m_Pa(3), "" & m_Pa(4), "" & pub_AddressListSN, "0", "605"
'            End If
          End If
                    
          'PrintPI strReceiveNo 'Removed by Morgan 2017/6/26 程式有錯,不應該產生請款單,需要再討論--David
          
          If mKind = 1 Then  '整批+Y20624
            '新增整批定稿列印清單資料
            'Modified by Lydia 2016/04/19 代理人Y20624和代理人Y5133301併入整批的1~10號
            'If Option1(2).Value Then
            'Modified by Lydia 2018/07/16 抓特定代理人字串
'                If Left(mstrTxt(2), 8) = "Y2062400" Then
'                '  'Modify By Sindy 2015/7/6 +, IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "") +e化註記
'                   PUB_AddNewLetterList "繳年費通知函", Me.Text1(7).Text & "-" & Me.Text1(8).Text & "(Y2062400)", "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
'                ElseIf Left(mstrTxt(2), 8) = "Y5133301" Then
'                   PUB_AddNewLetterList "繳年費通知函", Me.Text1(9).Text & "-" & Me.Text1(10).Text & "(Y5133301)", "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
'                'end 2016/04/19
'                'Added by Lydia 2016/11/04
'                ElseIf Left(mstrTxt(2), 8) = "Y2006500" Then
'                   PUB_AddNewLetterList "繳年費通知函", Me.Text1(11).Text & "-" & Me.Text1(12).Text & "(Y2006500)", "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
'                'end 2016/11/04
'                'Added by Lydia 2017/12/18
'                ElseIf Left(mstrTxt(2), 8) = "Y5285800" Then
'                   PUB_AddNewLetterList "繳年費通知函", Me.Text1(11).Text & "-" & Me.Text1(12).Text & "(Y5285800)", "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
'                'end 2016/12/18
'                Else
'                  'Modify By Sindy 2015/7/6 +, IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "") +e化註記
'                  PUB_AddNewLetterList "繳年費通知函", Me.Text1(5).Text & "-" & Me.Text1(6).Text, "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
'                End If
                If ChkInList <> "" And InStr(ChkInList, Left(mstrTxt(2), 8)) > 0 Then
                    tmpArr1 = Empty
                    tmpArr2 = Empty
                    tmpArr1 = Split(ChkInList, ",")
                    For intI = 0 To UBound(tmpArr1)
                         If Trim(tmpArr1(intI)) <> "" And Mid(tmpArr1(intI), 1, 8) = Left(mstrTxt(2), 8) Then
                              tmpArr2 = Split(tmpArr1(intI), "|")
                              PUB_AddNewLetterList "繳年費通知函", tmpArr2(3) & "-" & tmpArr2(4) & "(" & tmpArr2(0) & ")", "" & m_Pa(1), "" & m_Pa(2), "" & m_Pa(3), "" & m_Pa(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
                              Exit For
                         End If
                    Next intI
                Else
                    'Modified by Lydia 2024/05/29
                    'PUB_AddNewLetterList "繳年費通知函", Me.Text1(5).Text & "-" & Me.Text1(6).Text, "" & m_PA(1), "" & m_PA(2), "" & m_PA(3), "" & m_PA(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
                    PUB_AddNewLetterList "繳年費通知函", IIf(Option1(2).Value = 1, Me.Text1(5).Text & "-" & Me.Text1(6).Text, Me.txtData(0) & "-" & Me.txtData(1)), "" & m_Pa(1), "" & m_Pa(2), "" & m_Pa(3), "" & m_Pa(4), IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
                End If
'end 2018/07/16
           ' PrintPI strReceiveNo
          End If
          If mKind = 2 Then '本所案號->列印結束
            MsgBox "列印結束 !" & IIf(m_PrintRpt1 = True, vbCrLf & "定稿轉PDF存卷宗區有錯誤,詳情請看(" & PUB_Getdesktop & "\" & m_strFileName1 & ")", ""), vbInformation
          End If
       End If
    Else 'If mKind = 2 Then '本所案號->無資料可供列印
       InsertQueryLog (0)
       MsgBox "無符合條件之資料可列印 !", vbInformation
    End If
End Function

'Add By Sindy 2015/10/30
'資料檢核表
'Modified by Lydia 2016/04/19 +傳檔名開頭strTit
Private Sub ReadTxt1(strCaseNo As String, strRecvNo As String, strNote As String, strTit As String)
Dim i As Integer
Dim strTemp(1 To 7) As String
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      'Modified by Lydia 2016/04/19
      'If Option1(2).Value Then
      '   m_strFileName1 = Me.Caption & Me.Text1(7).Text & "-" & Me.Text1(8).Text & "資料檢核表.txt"
      'Else
      '   m_strFileName1 = Me.Caption & Me.Text1(5).Text & "-" & Me.Text1(6).Text & "資料檢核表.txt"
      'End If
      m_strFileName1 = strTit & "資料檢核表.txt"
      'end 2016/04/19
      
      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
      'Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
      Print #ff1, "本所案號        總收文號   原因"
      Print #ff1, "=============== ========== ============================================="
   End If
   For i = 1 To 3
      strTemp(i) = ""
   Next i
   strTemp(1) = convForm(CheckStr(Trim(strCaseNo)), 15)
   strTemp(2) = convForm(CheckStr(Trim(strRecvNo)), 10)
   strTemp(3) = Trim(strNote)
   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3)
End Sub

'Added by Morgan 2015/10/20
Private Sub AddPLetterList()
   On Error GoTo ErrHnd
   'Modified by Lydia 2015/10/26 寰華案定義改為:新案發文人員為外專程序者
'   strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08,LL09)" & _
'      " select '" & strUserNum & "','寰華案繳年費通知函','" & Text1(5).Text & "-" & Text1(6).Text & "'" & _
'      ",np02,np03,np04,np05,nvl(pa75,pa26) LL08,Decode(GetEmailFlag(np02||np03||np04||np05),'E','Ｅ','e','ｅ') eMail  from nextprogress A, caseprogress, patent " & _
'      " WHERE NP09 BETWEEN " & DBDATE(Text1(5)) & " AND " & DBDATE(Text1(6)) & " AND NP02||NP07||NP06='P605'" & _
'      " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
'      " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
'      " and not exists" & _
'      " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
'      " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
   'Modified by Lydia 2019/05/09 +PK: 使用者帳號@電腦名稱(pub_HostName)
   'Added by Lydia 2024/05/29 拆分給各區管制人
   If Option1(2).Value = True Then
      'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
      'strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08,LL09)" & _
         " select '" & strUserNum & "@" & pub_HostName & "','寰華案繳年費通知函','" & txtData(0).Text & "-" & txtData(1).Text & "'" & _
         ",np02,np03,np04,np05,nvl(pa75,pa26) LL08,Decode(GetEmailFlag(np02||np03||np04||np05),'E','Ｅ','e','ｅ') eMail from nextprogress A, caseprogress, patent,staff,fagent,nation " & _
         " WHERE NP09 BETWEEN " & DBDATE(txtData(0)) & " AND " & DBDATE(txtData(1)) & " AND NP02||NP07||NP06='P605'" & _
         " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
         " and cp83=st01(+) and (st03='F22' or st01 is null) " & _
         " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
         " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na79='" & Trim(txtData(2)) & "' and not exists" & _
         " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
         " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
      strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08,LL09)" & _
         " select '" & strUserNum & "@" & pub_HostName & "','寰華案繳年費通知函','" & txtData(0).Text & "-" & txtData(1).Text & "'" & _
         ",np02,np03,np04,np05,nvl(pa75,pa26) LL08,Decode(GetEmailFlag(np02||np03||np04||np05),'E','Ｅ','e','ｅ') eMail from nextprogress A, caseprogress, patent,staff,fagent,nation " & _
         " WHERE NP09 BETWEEN " & DBDATE(txtData(0)) & " AND " & DBDATE(txtData(1)) & " AND NP02||NP07||NP06='P605'" & _
         " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
         " and cp83=st01(+) and (st03='F22' or st01 is null) " & _
         " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
         " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+) and na79='" & Trim(txtData(2)) & "'" & _
         " and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
      'end 2025/10/28
   Else
   'end 2024/05/29
      'Modified by Morgan 2025/10/28 +原規則遇延半年仍不辦但過期後又要辦時會誤判，故改判斷不續辦管制的期限才是延期後的期限 Ex:FCP-061556
      'strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08,LL09)" & _
         " select '" & strUserNum & "@" & pub_HostName & "','寰華案繳年費通知函','" & Text1(5).Text & "-" & Text1(6).Text & "'" & _
         ",np02,np03,np04,np05,nvl(pa75,pa26) LL08,Decode(GetEmailFlag(np02||np03||np04||np05),'E','Ｅ','e','ｅ') eMail from nextprogress A, caseprogress, patent,staff " & _
         " WHERE NP09 BETWEEN " & DBDATE(Text1(5)) & " AND " & DBDATE(Text1(6)) & " AND NP02||NP07||NP06='P605'" & _
         " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
         " and cp83=st01(+) and (st03='F22' or st01 is null) " & _
         " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
         " and not exists" & _
         " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
         " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
      strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08,LL09)" & _
         " select '" & strUserNum & "@" & pub_HostName & "','寰華案繳年費通知函','" & Text1(5).Text & "-" & Text1(6).Text & "'" & _
         ",np02,np03,np04,np05,nvl(pa75,pa26) LL08,Decode(GetEmailFlag(np02||np03||np04||np05),'E','Ｅ','e','ｅ') eMail from nextprogress A, caseprogress, patent,staff " & _
         " WHERE NP09 BETWEEN " & DBDATE(Text1(5)) & " AND " & DBDATE(Text1(6)) & " AND NP02||NP07||NP06='P605'" & _
         " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
         " and cp83=st01(+) and (st03='F22' or st01 is null) " & _
         " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
         " and not exists (select * from caseprogress B where cp09=A.np01 and cp10='907')"
      'end 2025/10/28
   End If
   cnnConnection.Execute strSql, intI
   Exit Sub
   
ErrHnd:
      MsgBox Err.Description, vbCritical, "寰華案清單產生失敗!!"
End Sub

'Added by Lydia 2016/04/22 保留代理人繳費期限
'Memo by Lydia 2024/05/29 刪除不用的Code

'Added by Lydia 2018/07/16 讀取DB記錄,判斷特定代理人
Private Sub QueryData()
Dim rsRead As New ADODB.Recordset

On Error GoTo ErrorHand2:
'將原本畫面中固定特定代理人的資料,改存在DB(FormDate)用Grid呈現,並且只要增加記錄不需改程式
'Step 1: 檢查現在到下期催繳日前是否有期限未通知
'Step 2: 增加相對應記錄
'insert into formdate (FD01,FD02,FD03,FD04,FD05,FD06,FD07,FD08) values('frm060304','Y5349600-1',1071101,'A3034',20180716,160000,5,4);
'insert into formdate (FD01,FD02,FD03,FD04,FD05,FD06,FD07,FD08) values('frm060304','Y5349600-2',1071130,'A3034',20180716,160000,5,4);

    Call SetGrd(True) '清空
    '非整批的1~10日,不可抓特定代理人的資料
    'Mark by Lydia 2024/05/29 拆分給各區管制人
    'If Option1(1).Value = True Or Option2(1).Value = True Or Option2(2).Value = True Then
    '    strSql = "select ' ' as V, fa01,fa05||' '||fa63||' '||fa64||' '||fa65 as fname " & _
                 ",a.fd08 months,null as fdate1, null as fdate2 " & _
                 ",null as edate1 " & _
                 ",null as edate2 "
    'Else
    'end 2024/05/29
        strSql = "select 'V' as V, fa01,fa05||' '||fa63||' '||fa64||' '||fa65 as fname " & _
                 ",a.fd08 as months,a.fd03 as fdate1, b.fd03 as fdate2 " & _
                 ",substr(to_char(add_months(to_date((a.fd03+19110000),'yyyymmdd'), 1 ),'yyyymmdd') -19110000,1,7) as edate1 " & _
                 ",substr(to_char(add_months(to_date((b.fd03+19110000),'yyyymmdd'), 1 ),'yyyymmdd') -19110000,1,7) as edate2 "
    'End If 'Mark by Lydia 2024/05/29 拆分給各區管制人
    'Modified by Lydia 2024/05/29
    'strSql = strSql & "from formdate a,formdate b,fagent " & _
             "where a.fd01='" & Me.Name & "' and nvl(a.fd07,'0')>='1' and instr(a.fd02,'-1') > 0 " & _
             "and a.fd01=b.fd01(+) and a.fd07=b.fd07(+) and instr(b.fd02,'-2') > 0 " & _
             "and substr(a.fd02,1,8)=fa01(+) and '0'=fa02(+) " & _
             "order by a.fd07 "
    strSql = strSql & "from formdate a,formdate b,fagent,nation " & _
             "where a.fd01='" & Me.Name & "' and nvl(a.fd07,'0')>='1' and instr(a.fd02,'-1') > 0 " & _
             "and a.fd01=b.fd01(+) and a.fd07=b.fd07(+) and instr(b.fd02,'-2') > 0 " & _
             "and substr(a.fd02,1,8)=fa01(+) and '0'=fa02(+) and fa10=na01(+) " & IIf(Trim(txtData(2)) <> "", "and na16='" & Trim(txtData(2)) & "' ", "") & _
             "order by a.fd07 "
    intI = 1
    Set rsRead = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
       MSHFlexGrid1.FixedCols = 0
       Set MSHFlexGrid1.Recordset = rsRead
       Call SetGrd
       MSHFlexGrid1.FixedCols = 5
    End If
    
    Call SetDefDate("R") 'Added by Lydia 2024/05/29
    Set rsRead = Nothing
    Exit Sub

ErrorHand2:
   If Err.Number > 0 Then
      MsgBox Err.Description
      Exit Sub
   End If
End Sub

'設定Grid
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iR As Integer
     
   arrGridHeadText = Array("V", "代理人編號", "代理人名稱", "前X月", "上次(起)", "上次(止)", "下次(起)", "下次(止)")
   arrGridHeadWidth = Array(300, 900, 1700, 600, 820, 820, 820, 820)
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
   
   ExceptList = ""
   ChkInList = ""
   For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       If iRow <> 1 Then
           MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       End If
   Next
   For intI = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.row = intI
        For iRow = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.col = iRow
           MSHFlexGrid1.CellBackColor = &H80000005
           '置中
           If iRow = 0 Or iRow > 2 Then
               MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
           End If
           
           '記錄固定排除和特定代理人
           If "" & MSHFlexGrid1.TextMatrix(intI, iRow) <> "" Then
                If iRow = 1 Then
                    ExceptList = ExceptList & "," & MSHFlexGrid1.TextMatrix(intI, iRow) '固定排除的代理人編號(整批)
                    If UCase("" & MSHFlexGrid1.TextMatrix(intI, 0)) = "V" Then '特定代理人資料
                        ChkInList = ChkInList & "," & MSHFlexGrid1.TextMatrix(intI, iRow)
                    End If
                ElseIf UCase("" & MSHFlexGrid1.TextMatrix(intI, 0)) = "V" And iRow >= 4 And iRow <= 7 Then
                    ChkInList = ChkInList & "|" & MSHFlexGrid1.TextMatrix(intI, iRow)
                    'Added by Lydia 2024/05/29 記錄前X月,放在最後面
                    If Option1(2).Value = True And iRow = 7 Then
                       ChkInList = ChkInList & "|" & MSHFlexGrid1.TextMatrix(intI, 3)
                    End If
                End If
           End If
        Next iRow
   Next intI
   
   If ExceptList <> "" Then ExceptList = Mid(ExceptList, 2)
   If ChkInList <> "" Then ChkInList = Mid(ChkInList, 2)
   MSHFlexGrid1.Visible = True
End Sub

'Added by Lydia 2024/05/29
Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txtData(Index) <> "" Then
            If ChkDate(txtData(Index)) = False Then
               GoTo EXITSUB
            End If
         End If
      Case 2
         If txtData(Index) <> "" Then
            lblFM2.Caption = GetStaffName(txtData(Index), True)
         Else
            lblFM2.Caption = ""
         End If
         If txtData(Index).Tag <> txtData(Index).Text Then
            Call QueryData
         End If
         txtData(Index).Tag = txtData(Index).Text
   End Select
   
   Exit Sub
   
EXITSUB:
   Cancel = True
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
End Sub

'Added by Lydia 2024/05/29
Private Sub SetDefDate(ByVal pStatus As String)

   If Trim(txtData(2)) <> "" Then
      strBUser = Trim(txtData(2))
      If pStatus = "U" And txtData(0) <> "" And txtData(1) <> "" Then
         If Val(strBDate2) < Val(txtData(0)) Then '記錄最後催的期限
            PUB_SaveLastDate Me.Name, strBUser & "-1", txtData(0)
            PUB_SaveLastDate Me.Name, strBUser & "-2", txtData(1)
            '抓特定代理人字串
            If ChkInList <> "" Then
               tmpArr1 = Empty
               tmpArr2 = Empty
               tmpArr1 = Split(ChkInList, ",")
               For intI = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intI)) <> "" Then
                     tmpArr2 = Split(tmpArr1(intI), "|")
                     PUB_SaveLastDate Me.Name, tmpArr2(0) & "-1", Trim(tmpArr2(3))
                     PUB_SaveLastDate Me.Name, tmpArr2(0) & "-2", Trim(tmpArr2(4))
                  End If
               Next intI
            End If
         End If
      End If
      strBDate1 = PUB_GetLastDate(Me.Name, strBUser & "-1")
      strBDate2 = PUB_GetLastDate(Me.Name, strBUser & "-2")
   Else
      strBUser = ""
      strBDate1 = ""
      strBDate2 = ""
   End If
End Sub


