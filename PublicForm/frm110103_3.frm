VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110103_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "閉卷"
   ClientHeight    =   5748
   ClientLeft      =   192
   ClientTop       =   972
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9324
   Begin VB.CheckBox ChkOutlook 
      Caption         =   "出OutLook草稿"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3750
      TabIndex        =   42
      Top             =   30
      Width           =   1600
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   3
      Left            =   4116
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3525
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆(&N)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   3
      Left            =   5352
      TabIndex        =   7
      Top             =   10
      Width           =   900
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1608
      TabIndex        =   23
      Top             =   168
      Width           =   2000
      Begin VB.Label lblCode 
         Height          =   180
         Index           =   2
         Left            =   1290
         TabIndex        =   24
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblCode 
         Height          =   180
         Index           =   1
         Left            =   948
         TabIndex        =   25
         Top             =   0
         Width           =   252
      End
      Begin VB.Label lblCode 
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1656
      TabIndex        =   18
      Top             =   168
      Visible         =   0   'False
      Width           =   2412
      Begin VB.Label lblTFCode 
         Height          =   180
         Index           =   3
         Left            =   1560
         TabIndex        =   19
         Top             =   0
         Width           =   372
      End
      Begin VB.Label lblTFCode 
         Height          =   180
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   372
      End
      Begin VB.Label lblTFCode 
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   0
         Width           =   372
      End
      Begin VB.Label lblTFCode 
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   852
      End
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   2
      Left            =   5790
      MaxLength       =   1
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   1
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   0
      Left            =   1056
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3525
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8412
      TabIndex        =   10
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6276
      TabIndex        =   8
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7104
      TabIndex        =   9
      Top             =   10
      Width           =   1284
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2040
      Left            =   75
      TabIndex        =   12
      Top             =   1440
      Width           =   9150
      _ExtentX        =   16150
      _ExtentY        =   3598
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "PS：請自行撰寫給代理人的指示信"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   4200
      Width           =   3495
   End
   Begin MSForms.ComboBox cboNote 
      Height          =   300
      Left            =   1050
      TabIndex        =   5
      Top             =   4755
      Width           =   8145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14367;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   390
      Left            =   1050
      TabIndex        =   6
      Top             =   5310
      Width           =   8145
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "14367;688"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboReason 
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   3840
      Width           =   8145
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14367;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1125
      TabIndex        =   11
      Top             =   420
      Width           =   8070
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14235;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   5355
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   115
      Left            =   120
      TabIndex        =   39
      Top             =   5100
      Width           =   8220
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "證書號數："
      Height          =   180
      Left            =   5130
      TabIndex        =   38
      Top             =   990
      Width           =   900
   End
   Begin VB.Label lblRegNo 
      Height          =   180
      Left            =   6060
      TabIndex        =   37
      Top             =   990
      Width           =   3165
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   120
      TabIndex        =   36
      Top             =   990
      Width           =   900
   End
   Begin VB.Label lblAppNo 
      Height          =   180
      Left            =   1065
      TabIndex        =   35
      Top             =   990
      Width           =   3075
   End
   Begin VB.Label lblAgent 
      Height          =   180
      Left            =   1125
      TabIndex        =   34
      Top             =   750
      Width           =   8085
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   750
      Width           =   930
   End
   Begin VB.Label Label3 
      Caption         =   "後續准駁簡單報告：             (Y：核准以及C類來函簡單報告)"
      Height          =   180
      Left            =   2490
      TabIndex        =   32
      Top             =   3555
      Width           =   5000
   End
   Begin VB.Label Label21 
      Caption         =   "案件備註："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   4785
      Width           =   975
   End
   Begin VB.Label lblChildCase 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "有子案或相關卷號"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7536
      TabIndex        =   30
      Top             =   3552
      Width           =   1680
   End
   Begin VB.Label lblSystem 
      Height          =   180
      Left            =   1104
      TabIndex        =   27
      Top             =   168
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   168
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "是否修改指示信內容：           （Y：Word）"
      Height          =   180
      Left            =   4020
      TabIndex        =   17
      Top             =   4500
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "是否列印指示信：          （N：不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   4500
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "閉卷原因："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3870
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "閉卷日期："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "本案期限："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1215
      Width           =   975
   End
End
Attribute VB_Name = "frm110103_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/10 改成Form2.0(lblAgent,cboReason,cboNote,txtCP64)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'2010/8/3 日期欄已修改 by sonia
Option Explicit

'intWhereComeFrom  1:frm110103_1     2:frm110103_2
Public intWhereComeFrom As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗
'intLeaveKind離開時，是0:結束1:回上一畫面
Public mPrev01 As Form 'Add by Sindy 2015/1/14 電子結案單須呼叫此作業

Dim bolLeave As Boolean, intLeaveKind As Integer
'strCaseCode上一畫面frm110103_2勾選的本所案號
'intTotalCaseCode上一畫面frm110103_2勾選的本所案號總數
'intNowCaseCode現在Query的收文號Index
Dim strCaseCode() As String, intTotalCaseCode As Integer, intNowCaseCode As Integer
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
'intWhere 國內,國外_CF,國外_FC
Dim intCaseKind As Integer, intWhere As Integer
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'儲存解除期限原因的編號
Dim strReasonNo() As String
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj011 As New prjTaieDll011.cls011
Dim strSql As String, SCp(1 To 79) As String, k As Integer, L As Integer, BolExit As Boolean
'Add By Cheng 2002/12/04
Dim m_blnFirstShow As Boolean
'add by nickc 2008/05/16 國家暫存
Dim strNation As String
Dim strMCaseCP09 As String 'Add by Morgan 2009/7/8 新多國主案收文號
Dim m_boleOrderLetter As Boolean 'Added by Morgan 2015/11/3 指示信電子化
Dim m_bolFMP As Boolean 'Add by Lydia 2016/10/19 判斷FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/09 是否為寰華案
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知
'Add by Amy 2025/06/02
Public intFCState As Integer '0-智權/1-FC商標/2-FC專利 發起之結案單
Dim strF0301 As String, bolInvoice As Boolean '結案單號/是否開請款單輸入
Dim strOutLookType As String 'Add by Amy 2025/07/10 "0":寄 工程師+承辦 / "1":寄 承辦

Private Sub ReadAllData()
Dim j As Integer, i As Integer
'Dim field() As String
Dim m_CCM04 As String, strRCodeN As String 'Add by Amy 2025/06/02 閉卷原因 代碼/名稱
strOutLookType = "" 'Add by Amy 2025/07/10
BolExit = False
If intWhereComeFrom = 2 Then
'   j = 1
'   For i = 1 To frm110103_2.grdDataList.Rows - 1
'          If frm110103_2.grdDataList.TextMatrix(i, 0) <> "" Then
'             ReDim Preserve strCaseCode(3, j)
'             strCaseCode(0, j) = frm110103_2.grdDataList.TextMatrix(i, 6)
'             strCaseCode(1, j) = frm110103_2.grdDataList.TextMatrix(i, 7)
'             strCaseCode(2, j) = frm110103_2.grdDataList.TextMatrix(i, 8)
'             strCaseCode(3, j) = frm110103_2.grdDataList.TextMatrix(i, 9)
'             j = j + 1
'          End If
'   Next
'   intTotalCaseCode = j - 1
'   intNowCaseCode = 1
   bolLeave = False
   intLeaveKind = 1
   SetDataListWidth
'   L = intNowCaseCode
   If intTotalCaseCode = 0 Then
      bolLeave = True
      Unload Me
      Exit Sub
   End If
'   For k = L To intTotalCaseCode
'      intNowCaseCode = k
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         strSql = "SELECT PA01 FROM PATENT WHERE PA01='" & strCaseCode(0, intNowCaseCode) & "' AND PA02='" & strCaseCode(1, intNowCaseCode) & "' AND PA03='" & strCaseCode(2, intNowCaseCode) & "' AND PA04='" & strCaseCode(3, intNowCaseCode) & "' AND (PA57 <>'Y' OR PA57 IS NULL) "
         strSql = strSql & " union all select TM01 FROM TRADEMARK WHERE TM01='" & strCaseCode(0, intNowCaseCode) & "' AND TM02='" & strCaseCode(1, intNowCaseCode) & "' AND TM03='" & strCaseCode(2, intNowCaseCode) & "' AND TM04='" & strCaseCode(3, intNowCaseCode) & "' AND (TM29 <>'Y' OR TM29 IS NULL) "
         strSql = strSql & " union all select LC01 FROM LAWCASE WHERE LC01='" & strCaseCode(0, intNowCaseCode) & "' AND LC02='" & strCaseCode(1, intNowCaseCode) & "' AND LC03='" & strCaseCode(2, intNowCaseCode) & "' AND LC04='" & strCaseCode(3, intNowCaseCode) & "' AND (LC08 <>'Y' OR LC08 IS NULL) "
         strSql = strSql & " union all select HC01 FROM HIRECASE WHERE HC01='" & strCaseCode(0, intNowCaseCode) & "' AND HC02='" & strCaseCode(1, intNowCaseCode) & "' AND HC03='" & strCaseCode(2, intNowCaseCode) & "' AND HC04='" & strCaseCode(3, intNowCaseCode) & "' AND (HC09 <>'Y' OR HC09 IS NULL) "
         strSql = strSql & " union all select SP01 FROM SERVICEPRACTICE WHERE SP01='" & strCaseCode(0, intNowCaseCode) & "' AND SP02='" & strCaseCode(1, intNowCaseCode) & "' AND SP03='" & strCaseCode(2, intNowCaseCode) & "' AND SP04='" & strCaseCode(3, intNowCaseCode) & "' AND (SP15 <>'Y' OR SP15 IS NULL) "
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount = 0 Then
            MsgBox "此 " & strCaseCode(0, intNowCaseCode) & "-" & strCaseCode(1, intNowCaseCode) & "-" & strCaseCode(2, intNowCaseCode) & "-" & strCaseCode(3, intNowCaseCode) & " 已閉卷！！"
            For j = 1 To frm110103_2.grdDataList.Rows - 1
               If frm110103_2.grdDataList.TextMatrix(j, 0) <> "" Then
                  If frm110103_2.grdDataList.TextMatrix(j, 6) = strCaseCode(0, intNowCaseCode) And frm110103_2.grdDataList.TextMatrix(j, 7) = strCaseCode(1, intNowCaseCode) And frm110103_2.grdDataList.TextMatrix(j, 8) = strCaseCode(2, intNowCaseCode) And frm110103_2.grdDataList.TextMatrix(j, 9) = strCaseCode(3, intNowCaseCode) Then
                     frm110103_2.grdDataList.TextMatrix(j, 0) = ""
                     frm110103_2.intChoose = frm110103_2.intChoose - 1
                  End If
               End If
            Next j
            bolLeave = True
            Me.Hide
            CheckOC
            If frm110103_2.intChoose = 0 Then
               BolExit = True
               CheckOC
               Unload Me
            Else
               intNowCaseCode = intNowCaseCode + 1
               'Added by Lydia 2021/07/05 因為現在已經是下一筆要處理的Index
               If intNowCaseCode > intTotalCaseCode Then
                    BolExit = True
                    intLeaveKind = 0
                    CheckOC
                    Unload Me
                    Exit Sub
               Else
               'end 2021/07/05
                    If intNowCaseCode = intTotalCaseCode Then
                       cmdOK(3).Visible = False
                    End If
                    ReadAllData
               End If 'Added by Lydia 2021/07/05
            End If
            'Unload Me
         Else
            Me.Show
         End If
         CheckOC
      End With
'      If BolExit = True Then
'         Exit For
'      End If
'   Next k
Else
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      strSql = "SELECT PA01 FROM PATENT WHERE PA01='" & strCaseCode(0, intNowCaseCode) & "' AND PA02='" & strCaseCode(1, intNowCaseCode) & "' AND PA03='" & strCaseCode(2, intNowCaseCode) & "' AND PA04='" & strCaseCode(3, intNowCaseCode) & "' AND (PA57 <>'Y' OR PA57 IS NULL) "
      strSql = strSql & " union all select TM01 FROM TRADEMARK WHERE TM01='" & strCaseCode(0, intNowCaseCode) & "' AND TM02='" & strCaseCode(1, intNowCaseCode) & "' AND TM03='" & strCaseCode(2, intNowCaseCode) & "' AND TM04='" & strCaseCode(3, intNowCaseCode) & "' AND (TM29 <>'Y' OR TM29 IS NULL) "
      strSql = strSql & " union all select LC01 FROM LAWCASE WHERE LC01='" & strCaseCode(0, intNowCaseCode) & "' AND LC02='" & strCaseCode(1, intNowCaseCode) & "' AND LC03='" & strCaseCode(2, intNowCaseCode) & "' AND LC04='" & strCaseCode(3, intNowCaseCode) & "' AND (LC08 <>'Y' OR LC08 IS NULL) "
      strSql = strSql & " union all select HC01 FROM HIRECASE WHERE HC01='" & strCaseCode(0, intNowCaseCode) & "' AND HC02='" & strCaseCode(1, intNowCaseCode) & "' AND HC03='" & strCaseCode(2, intNowCaseCode) & "' AND HC04='" & strCaseCode(3, intNowCaseCode) & "' AND (HC09 <>'Y' OR HC09 IS NULL) "
      strSql = strSql & " union all select SP01 FROM SERVICEPRACTICE WHERE SP01='" & strCaseCode(0, intNowCaseCode) & "' AND SP02='" & strCaseCode(1, intNowCaseCode) & "' AND SP03='" & strCaseCode(2, intNowCaseCode) & "' AND SP04='" & strCaseCode(3, intNowCaseCode) & "' AND (SP15 <>'Y' OR SP15 IS NULL) "
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount = 0 Then
         MsgBox "此 " & strCaseCode(0, intNowCaseCode) & "-" & strCaseCode(1, intNowCaseCode) & "-" & strCaseCode(2, intNowCaseCode) & "-" & strCaseCode(3, intNowCaseCode) & " 已閉卷！！"
         bolLeave = True
         Me.Hide
         CheckOC
         Unload Me
         Exit Sub
      Else
         Me.Show
      End If
      CheckOC
   End With
End If

Dim varSaveCursor, strReceiveCode As String, strReasonName() As String

On Error GoTo ErrHand
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
lblSystem = strCaseCode(0, intNowCaseCode)
If strCaseCode(0, intNowCaseCode) = 馬德里案 Then
   lblTFCode(0) = Left(strCaseCode(1, intNowCaseCode), 5)
   lblTFCode(1) = IIf(Right(strCaseCode(1, intNowCaseCode), 1) = "0", "", Right(strCaseCode(1, intNowCaseCode), 1))
   lblTFCode(2) = IIf(strCaseCode(2, intNowCaseCode) = "0", "", strCaseCode(2, intNowCaseCode))
   lblTFCode(3) = IIf(strCaseCode(3, intNowCaseCode) = "00", "", strCaseCode(3, intNowCaseCode))
Else
   lblCode(0) = strCaseCode(1, intNowCaseCode)
   lblCode(1) = IIf(strCaseCode(2, intNowCaseCode) = "0", "", strCaseCode(2, intNowCaseCode))
   lblCode(2) = IIf(strCaseCode(3, intNowCaseCode) = "00", "", strCaseCode(3, intNowCaseCode))
End If
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetSystemKind(strCaseCode(0, intNowCaseCode), intCaseKind, , intWhere) Then
   'If objPublicData.GetReceiveCode(strCaseCode(0, intNowCaseCode), strCaseCode(1, intNowCaseCode), strCaseCode(2, intNowCaseCode), strCaseCode(3, intNowCaseCode), strReceiveCode) = False Then GoTo Err1
   'If objPublicData.ReadAllData(strReceiveCode, cp(), field(), intCaseKind, intWhere) = False Then GoTo Err1
If ClsPDGetSystemKind(strCaseCode(0, intNowCaseCode), intCaseKind, , intWhere) Then
   If ClsPDGetReceiveCode(strCaseCode(0, intNowCaseCode), strCaseCode(1, intNowCaseCode), strCaseCode(2, intNowCaseCode), strCaseCode(3, intNowCaseCode), strReceiveCode) = False Then GoTo err1
   ReDim cp(TF_CP) As String
   cp(9) = strReceiveCode
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) = False Then GoTo err1
   'add by nickc 2008/05/16 暫存國家
   strNation = ""
   
   'Add by Lydia 2016/10/19 判斷FMP案
   'Modified by Morgan 2021/2/2
   'If Left(cp(12), 1) = "F" And cp(1) = "P" And field(9) <> "000" Then
   '   m_bolFMP = True
   'Else/
   '   m_bolFMP = False
   'End If
   m_bolFMP = PUB_ChkIsFMP(field(1), field(2), field(3), field(4), field(9))
   'end 2021/2/2
   'end 2016/10/19
   
   'Added by Lydia 2023/06/09 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, field(1), field(2), field(3), field(4))
   End If
   'end 2023/06/09
   'Added by Lydia 2023/07/28 FCP專利連結通知
   If field(1) = "FCP" Then
      m_PA177 = field(177)
   Else
      m_PA177 = ""
   End If
   'end 2023/07/28
   
   Select Case intCaseKind
   Case 專利
            strNation = field(9)
            'Add by Morgan 2010/7/15
            Label4 = "證書號："
            lblAppNo = field(11)
            lblRegNo = field(22)
            strExc(1) = ""
            If field(75) <> "" Then
               PUB_GetAgentName "1", field(75), strExc(1)
            End If
            lblAgent = field(75) & " " & strExc(1)
            'end 2010/7/15
   Case 商標
            strNation = field(10)
            'Add by Morgan 2010/7/15
            lblAppNo = field(12)
            Label4 = "審定號："
            lblRegNo = field(15)
            strExc(1) = ""
            If field(44) <> "" Then
               PUB_GetAgentName "1", field(44), strExc(1)
            End If
            lblAgent = field(44) & " " & strExc(1)
            'end 2010/7/15
   Case 法務
            strNation = field(15)
            'Add by Morgan 2010/7/15
            strExc(1) = ""
            If field(22) <> "" Then
               PUB_GetAgentName "1", field(22), strExc(1)
            End If
            lblAgent = field(22) & " " & strExc(1)
            Label18.Visible = False
            Label4.Visible = False
            'end 2010/7/15
   Case 顧問
            strNation = "000"
            'Add by Morgan 2010/7/15
            Label16(1).Visible = False
            Label18.Visible = False
            Label4.Visible = False
            'end 2010/7/15
   Case Else
            strNation = field(9)
            'Add by Morgan 2010/7/15
            lblAppNo = field(11)
            Label4 = "證書號："
            lblRegNo = field(14)
            strExc(1) = ""
            If field(26) <> "" Then
               PUB_GetAgentName "1", field(26), strExc(1)
            End If
            lblAgent = field(26) & " " & strExc(1)
            'end 2010/7/15
   End Select
   
   
   
   If intCaseKind = 顧問 Then
      SetNameToCombo cboCaseName, field(6), "", ""
   Else
      SetNameToCombo cboCaseName, field(5), field(6), field(7)
   End If
   Set grdDataList.Recordset = ReadCloseCaseDRst(intWhere, strCaseCode(0, intNowCaseCode), strCaseCode(1, intNowCaseCode), strCaseCode(2, intNowCaseCode), strCaseCode(3, intNowCaseCode))
   SetDataListVision grdDataList
   'modify by sonia 90.10.7
   For i = 1 To 4
      field(i) = strCaseCode(i - 1, intNowCaseCode)
   Next
   'edit by nickc 2006/06/22 從 dll 內 copy 出
   'Select Case obj011.CheckChildCaseOrCaseRelation(field())
   Select Case CheckChildCaseOrCaseRelation(field())
                Case 1, 2
                           lblChildCase.Visible = True
                Case 0
                           lblChildCase.Visible = False
                Case -1, -2
                           GoTo err1
   End Select
   
   'Modify by Amy 2025/06/02  +FC結案單電子化,「閉卷原因」改共用
'   Select Case ReadReasonOfRelief(strReasonNo(), strReasonName())
'                Case 1
'                           For i = 0 To UBound(strReasonNo)
'                                 cboReason.AddItem strReasonName(i)
'                           Next
'                           cboReason.ListIndex = 0
'                Case -1
'                           GoTo err1
'   End Select
   'Modify by Amy 2025/08/25 +if 原因代碼相同,但顯示名稱可能不同 ex:CFT 有智權部和外商的案子
   If UCase(TypeName(mPrev01)) <> UCase("frm210149_1") Then
      strExc(9) = "0"
      If field(1) = "FCT" Or ((field(1) = "T" Or field(1) = "CFT") And Left(PUB_GetST03(strUserNum), 1) = "F") Or (field(1) = "S" And strNation = "000") Then
         'Nvl(ROR03,ROR02) 商標專有名稱->原名稱
         strExc(9) = "1"
      ElseIf field(1) = "FCP" Or ((field(1) = "P" Or field(1) = "CFP") And Left(PUB_GetST03(strUserNum), 1) = "F") Or field(1) = "FG" Then
         'Nvl(ROR04,ROR02)  專利專有名稱->原名稱
         strExc(9) = "2"
      End If
      Call Pub_SetCloseReason(Val(strExc(9)), Me.Name, cboReason)
   End If
   'end 2025/08/25
   'end 2025/06/02
   
   ChkOutlook.Visible = False 'Add by Amy 2025/06/16
   If Not mPrev01 Is Nothing Then   'Added by Morgan 2021/5/4
      'Add By Sindy 2015/1/14 結案單電子化
      If UCase(mPrev01.Name) = UCase("frm210149_1") Then
         'Modify by Amy 2025/06/02 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            strF0301 = mPrev01.txtF0301
            intFCState = mPrev01.intFCState
            bolInvoice = mPrev01.bolInvoice
            'Add by Amy 2025/08/25 國內與國外結案單原因代碼相同,但顯示名稱可能不同
            '     ex:測式國外結案單操作 CFT-025335 原因15,國內無此代碼
            Call Pub_SetCloseReason(intFCState, Me.Name, cboReason)
            'end 2025/08/25
            '閉卷原因=結案記錄
            m_CCM04 = mPrev01.m_F0305
            strRCodeN = m_CCM04
            Call Pub_SetCloseReason(intFCState, Me.Name, , strRCodeN)
            If strRCodeN = MsgText(601) Then
               Me.cboReason = m_CCM04
            Else
               Me.cboReason = m_CCM04 & "--" & strRCodeN
            End If
            '若外專人員可勾選「出Outlook草稿」
            If intFCState = 2 And Left(PUB_GetST03(strUserNum), 2) = "F2" Then
               ChkOutlook.Visible = True
               '外專承辦於結案單勾選需請款項目,「出Outlook草稿」預設勾選
               'Modify by Amy 2025/07/10 +strOutLookType
               If Pub_ChkCloseInvoce(Me.Name, strF0301, field(1), field(2), field(3), field(4), , strOutLookType) = True Then
                  ChkOutlook.Value = vbChecked
               End If
            End If
         Else
            '結案記錄
            'Modify by Amy 2025/06/16 改抓共用
'            If Val(mPrev01.m_F0305) = 99 Then
'               Me.cboReason.ListIndex = Me.cboReason.ListCount - 1
'            Else
'               Me.cboReason.ListIndex = Val(mPrev01.m_F0305) - 1
'            End If
            m_CCM04 = Left(mPrev01.m_F0305, 2)
            strRCodeN = m_CCM04
            Call Pub_SetCloseReason(intFCState, Me.Name, , strRCodeN)
            If strRCodeN = MsgText(601) Then
               Me.cboReason = m_CCM04
            Else
               Me.cboReason = m_CCM04 & "--" & strRCodeN
            End If
            'end 2025/06/16
         End If
         '備註
         Me.txtCP64.Text = Trim(Me.txtCP64.Text) & Trim(mPrev01.m_F0306)
      End If
      '2015/1/14 END
   End If
Else
err1:
   frm110103_2.ReChoose intNowCaseCode, strCaseCode()
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
If Len(Trim(txtCaseField(0))) = 0 Then txtCaseField(0) = ChangeWStringToTString(GetTodayDate)
Screen.MousePointer = varSaveCursor
Exit Sub
ErrHand:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub
Private Function SaveDatabase() As Boolean
Select Case intCaseKind
             Case 專利
                        field(57) = "Y"
                        field(58) = IIf(CheckIsDate(txtCaseField(0), False), txtCaseField(0), ChangeTStringToWString(txtCaseField(0)))
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                        'field(59) = strReasonNo(cboReason.ListIndex)
                        field(59) = GetReason(cboReason)
             Case 商標
                        field(29) = "Y"
                        field(30) = IIf(CheckIsDate(txtCaseField(0), False), txtCaseField(0), ChangeTStringToWString(txtCaseField(0)))
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                        'field(31) = strReasonNo(cboReason.ListIndex)
                        field(31) = GetReason(cboReason)
             Case 法務
                        field(8) = "Y"
                        field(9) = IIf(CheckIsDate(txtCaseField(0), False), txtCaseField(0), ChangeTStringToWString(txtCaseField(0)))
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                        'field(10) = strReasonNo(cboReason.ListIndex)
                        field(10) = GetReason(cboReason)
             Case 顧問
                        field(9) = "Y"
                        field(10) = IIf(CheckIsDate(txtCaseField(0), False), txtCaseField(0), ChangeTStringToWString(txtCaseField(0)))
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                        'field(31) = strReasonNo(cboReason.ListIndex)
                        field(31) = GetReason(cboReason)
             Case Else
                        field(15) = "Y"
                        field(16) = IIf(CheckIsDate(txtCaseField(0), False), txtCaseField(0), ChangeTStringToWString(txtCaseField(0)))
                        'Modify by Amy 2025/06/02 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                        'field(17) = strReasonNo(cboReason.ListIndex)
                        field(17) = GetReason(cboReason)
End Select
'edit by nickc 2007/02/05 不用 dll 了
'If obj011.SaveCloseCaseData(intCaseKind, intWhere, cp(), field()) Then
If Cls011SaveCloseCaseData(intCaseKind, intWhere, cp(), field()) Then
   SaveDatabase = True
Else
   MsgBox "存檔失敗!!", vbCritical
End If
End Function

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer, varSaveCursor
Dim strTmp As String, bolChk As Boolean
Dim oErrMsg As String 'add by nickc 2005/04/22
Dim strUpdDate As String, strUpdTime As String, strF0308 As String, strF0309 As String 'Add By Sindy 2015/1/14
Dim m_CP12 As String, m_CP13 As String 'Add By Sindy 2015/1/15
Dim bolDelCM As Boolean 'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯
'Add by Amy 2018/06/27
Dim strTo As String, bolCommit As Boolean
Dim strOldF0308 As String, strCmd(1) As String 'Add by Amy 2021/06/23 記錄原程序人員(CFT/CFC/S可職代操作)/更新語法
Dim strReason As String 'Add by Amy 2025/06/02
Dim strSubject As String, strContent As String 'Add By Sindy 2025/6/4
Dim strNotPay As String, strCCD08 As String 'Add by Amy 2025/08/08
Dim bolMailF0202_3 As Boolean  'Add by Amy 2025/08/19 解除期限後是否寄信給補看人員
Dim bolOpen21H0Ok As Boolean 'Add by Amy 2025/10/20

oErrMsg = ""

Select Case Index
             Case 0 '確定
                       'Add by Amy 2025/10/28 有請款項目 且請款單輸入已開啟需關閉
                       If bolInvoice = True Then
                          If PUB_CheckFormExist("Frmacc21h0") = True Or PUB_CheckFormExist("Frmacc21h01") = True Then
                             MsgBox "請款單輸入已開啟請關閉", vbExclamation
                             Exit Sub
                          End If
                        End If
                        'end 2025/10/28
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 2
                               If txtCaseField(i).Enabled Then
                                  If CheckKeyIn(i) <> 1 And CheckKeyIn(i) <> 4 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                     'Modified by Morgan 2015/5/14
                                     'Exit For
                                     Screen.MousePointer = vbDefault
                                     Exit Sub
                                     'end 2015/5/14
                                  End If
                               End If
                        Next
                     
                     CheckFCPDualCase 'Added by Morgan 2015/5/18
         
                     '2006/3/29 ADD BY SONIA 已發文請求面詢407但無通知面詢1401且無面詢408之收文者提示訊息
                     CHECKFCP407 cp(1), cp(2), cp(3), cp(4)
                     '2006/3/29 END
                     
                     '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
                     'Pub_EndModCashMsg strNation  '2011/10/28 ADD BY SONIA
                     Pub_EndModCashMsg strNation, cp(1), cp(2), cp(3), cp(4)
                     
                     'Add by Morgan 2009/7/8
                     '多國主案閉卷彈新主案國家及案號之訊息
                     '新主案順序 美->日->德->英->韓->收文順序
                     strMCaseCP09 = ""
                     If cp(1) = "CFP" Then
                        strSql = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
                           ",na03,b.cp09 from caseprogress a,caserelation,patent,caseprogress b,nation where a.cp01='" & cp(1) & "'" & _
                           " and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp21 is null and a.cp31='Y' and a.cp27 is null and a.cp57 is null" & _
                           " and cr01(+)=a.cp01 and cr02(+)=a.cp02 and cr03(+)=a.cp03 and cr04(+)=a.cp04 and pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08 and pa57 is null" & _
                           " and b.cp01(+)=pa01 and b.cp02(+)=pa02 and b.cp03(+)=pa03 and b.cp04(+)=pa04 and b.cp31='Y' and b.cp21='Y'" & _
                           " and na01(+)=pa09" & _
                           " order by decode(pa09,'101','1','011','2','231','3','201','4','012','5',b.CP09) ASC"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strMCaseCP09 = RsTemp("cp09")
                           MsgBox "本案為多國主案,閉卷後將改設定【 " & RsTemp("CNo") & "(" & RsTemp("na03") & ") 】為主案!!"
                        End If
                     End If
                     
                     'Added by Morgan 2015/11/3 指示信電子化
                     'P非臺灣案指示信都要彈修改畫面來確認送判的內容
                     m_boleOrderLetter = False
                     'Modified by Morgan 2015/12/15 外專程序除外
                     If field(1) = "P" And field(9) <> "000" And txtCaseField(1) = "" And Left(Pub_StrUserSt03, 1) <> "F" Then
                        'Modified by Morgan 2015/12/14 無期限閉卷不會有指示信--韻丞 P-109170
                        'Modified by Morgan 2021/5/4 可能會有指示信 Ex:P-119291 --蕭茹曣
                        m_boleOrderLetter = True
                     End If
                     'end 2015/11/3
                     
                     'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯
                     If field(1) <> "FCP" And m_bolFMP = False And intCaseKind = 專利 And (field(8) = "1" Or field(8) = "2") Then
                        If PUB_DualCaseRelationExist(field) Then
                           If PUB_ChkCPExist(field, IIf(field(8) = "1", "101", "102"), 1) Then '判斷未發文的新案才取消關聯
                              bolDelCM = True
                           End If
                        End If
                     End If
                     'end 2016/10/19
                     'Modify by Amy 2021/06/23 取得案件表單主檔程序人員
                     If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                        strOldF0308 = GetFlow003Data(Trim(mPrev01.txtF0301), , "F0308")
                     End If
                     
                     'Added by Lydia 2022/05/18 一案二請新型案上年費不續辦/閉卷時，判斷發明案尚未核准 '2022/05/16 frm110101_2
                     If field(1) = "FCP" And field(8) = "2" And field(9) = "000" Then
                          If PUB_IsDualApply(field, strExc, , , , , , True) = True Then
                              strExc(0) = "select pa16 from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' and pa08='1' "
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                  '"是"按鈕'上不續辦/閉卷；"否"按鈕'新增以下
                                  If "" & RsTemp.Fields("pa16") = "" Then
                                     If MsgBox("為一案兩請且發明案尚未核准，是否閉卷？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                                         strExc(5) = PUB_GetFCPSalesNo(field(1), field(2), field(3), field(4), cp(10))
                                         strExc(6) = PUB_GetFCPProSup(strExc(5))
                                         strExc(7) = PUB_GetFCPHandler(field(1), field(2), field(3), field(4), cp(10))
                                         strExc(8) = PUB_GetFCPProSup(strExc(7))
                                         '行事曆
                                         If PUB_AddFCPStaffCalendar(CompWorkDay(4, strSrvDate(1)), "1", strExc(5) & "," & strExc(7), "一案二請因新型案接獲客戶閉卷通知且發明案尚未核准，待承辦通知客戶確認再上閉卷", strExc(5) & "," & strExc(7), "1", field(1), field(2), field(3), field(4)) = False Then
                                             Exit Sub
                                         End If
                                         strExc(1) = "【一案兩請】新型案接獲客戶年費閉卷通知且""發明案尚未核准""，請確認報告客戶。 Our Ref:" & field(1) & "-" & field(2) & IIf(field(3) <> "0", "-" & field(3), "") & IIf(field(4) <> "00", "-" & field(4), "") & "[INCOM.]"
                                         strExc(2) = "1. 新型案接獲客戶閉卷通知且""發明案尚未核准""，請確認報告客戶。" & vbCrLf & _
                                                          "2. 行事曆已自動新增一3天期限" & vbCrLf & _
                                                          "3. 承辦確認後再行通知程序上閉卷"
                                         strExc(3) = strExc(6) & ";" & strExc(7) & ";" & strExc(8) & ";backup"
                                         PUB_SendMail strUserNum, strExc(5), "", strExc(1), strExc(2), , , , , , strExc(3)
                                         Screen.MousePointer = vbDefault
                                         Exit Sub
                                     End If
                                  End If
                              End If
                          End If
                     End If
                     'end 2022/05/18
                     
                        'If i = 3 Then
                        '    If SaveDatabase Then
                        '       intNowCaseCode = intNowCaseCode + 1
                        '       If intNowCaseCode = intTotalCaseCode Then
                        '           frm110103_2.ReChoose intNowCaseCode, strCaseCode()
                        '           bolLeave = True
                        '           Unload Me
                        '        Else
                        '           If intNowCaseCode = intTotalCaseCode Then
                        '              cmdOK(3).Visible = False
                        '           End If
                        '           ReadAllData
                        '        End If
                        '    End If
                        'End If
                        '**************************************************
                        'add by nick 2005/04/22 transation
                        On Error GoTo CheckingErr
                        cnnConnection.BeginTrans
                        
                        'Add by Morgan 2009/7/8
                        If strMCaseCP09 <> "" Then
                            strSql = "UPDATE caseprogress SET cp21=null WHERE cp09='" & strMCaseCP09 & "' and cp21='Y'"
                            cnnConnection.Execute strSql
                        End If
                             
'Removed by Morgan 2012/10/1 改在列印結案單時提醒並列印於接洽單上
'                        'Add by Morgan 2009/10/15
'                        '大陸案一案兩請:新型年費欲結案時,若該發明案尚未核准公告,則發E-MAIL告知智權同仁及其所屬區主管
'                        '不可放後面執行否則年費期限會被解除或是取消收文
'                        If field(1) = "P" And field(9) = "020" And field(8) = "2" And Val(DBDATE(field(10))) >= 20091001 Then
'                           strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,nvl(np10,cp13) C2" & _
'                              " from (select cm05,cm06,cm07,cm08,cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
'                              " union select cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
'                              ",patent,nextprogress,caseprogress where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and (pa16 is null or pa16='2')" & _
'                              " and np02(+)=cm01 and np03(+)=cm02 and np04(+)=cm03 and np05(+)=cm04 and np06(+) is null and np07(+)='605'" & _
'                              " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and cp27(+) is null and cp10(+)='605' and cp57(+) is null and nvl(np10,cp13) is not null"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strExc(1) = field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4))
'                              strExc(1) = "提醒:" & strExc(1) & "大陸案為一案兩請,新型放棄續繳年費將同時放棄發明或實用新型間擇一選擇的權利。"
'                              strExc(2) = "" & RsTemp.Fields("C2")
'                              strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                                 " values ('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'                              cnnConnection.Execute strSql, intI
'
'                              strExc(0) = "select a0908 from staff,acc090 where st01='" & strExc(2) & "' and a0901(+)=st15 and a0908<>'" & strExc(2) & "'"
'                              intI = 1
'                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                              If intI = 1 Then
'                                 strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                                    " values ('" & strUserNum & "','" & RsTemp(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'                                 cnnConnection.Execute strSql, intI
'                              End If
'                           End If
'                        End If
'                        'end 2009/10/15
                        
                        ' nick 900802 改
                        'UPDATE 基本檔是否閉卷,閉卷日期,閉卷原因
                        'Modify By Cheng 2002/01/29
                        '更新各基本檔的備註進度
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305 拆至結案單主檔中
                        'strReason = strReasonNo(cboReason.ListIndex)
                        strReason = GetReason(cboReason)
                        Select Case Val(CheckSys(cp(1)))
                        Case 1
                              'Modify By Cheng 2002/01/29
'                             strSQL = "UPDATE PATENT SET PA57='Y',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & strReasonNo(cboReason.ListIndex) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                             'strSql = "UPDATE PATENT SET PA57='Y',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,PA91='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                             strSql = "UPDATE PATENT SET PA57='Y',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & strReason & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,PA91='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                             cnnConnection.Execute strSql
                             'Add By Cheng 2002/05/29
                             strSql = "UPDATE PATENT SET PA89='" & Me.txtCaseField(3).Text & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                             cnnConnection.Execute strSql

                        Case 2
                              'Modify By Cheng 2002/01/29
'                             strSQL = "UPDATE TRADEMARK SET TM29='Y',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & strReasonNo(cboReason.ListIndex) & "' WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                             'strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,TM58='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                             strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & strReason & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,TM58='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                             cnnConnection.Execute strSql
                        Case 3
                              'Modify By Cheng 2002/01/29
'                             strSQL = "UPDATE LAWCASE SET LC08='Y',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & strReasonNo(cboReason.ListIndex) & "' WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                             'strSql = "UPDATE LAWCASE SET LC08='Y',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,LC27='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                             strSql = "UPDATE LAWCASE SET LC08='Y',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & strReason & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,LC27='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                             cnnConnection.Execute strSql
                        Case 4
                              'Modify By Cheng 2002/01/29
'                             strSQL = "UPDATE HIRECASE SET HC09='Y',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & strReasonNo(cboReason.ListIndex) & "' WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                             'strSql = "UPDATE HIRECASE SET HC09='Y',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,HC12='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                             strSql = "UPDATE HIRECASE SET HC09='Y',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & strReason & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,HC12='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                             cnnConnection.Execute strSql
                        Case 5, 6, 7, 8
                              'Modify By Cheng 2002/01/29
'                             strSQL = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & strReasonNo(cboReason.ListIndex) & "' WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                             'strSql = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,SP18='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                             strSql = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & strReason & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,SP18='" & ChgSQL(Me.cboNote.Text) & "' ") & " WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                             cnnConnection.Execute strSql
                        Case Else
                        End Select
                        'UPDATE 進度檔,取消收文日期,取消收文原因
                        '93.10.5 MODIFY BY SONIA
                        'strSQL = "UPDATE CASEPROGRESS SET CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & strReasonNo(cboReason.ListIndex) & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
                        'strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & strReasonNo(cboReason.ListIndex) & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
                        strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & strReason & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
                        'end 2025/06/02
                        'Added by Lydia 2016/01/29 排除FCP案的代辦退費(實審,再審和再審延期)
                        If cp(1) = "FCP" Then
                           'Modified by Morgan 2022/11/23 +排除續行母案再審的代辦退費 Ex:FCP-067213 --Winfrey
                           strSql = strSql & "and cp09 not in (select a.cp09 from caseprogress a,caseprogress b where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435') " & _
                                    "union select a.cp09 from  caseprogress a,caseprogress b,nextprogress where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107' " & _
                                    "union select a.cp09 from  caseprogress a,caseprogress b,caseprogress c where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107') "
                        End If
                        cnnConnection.Execute strSql
                        '93.10.5 END
                        
                        'Add By Cheng 2002/01/22
                        '更新案件進度檔時, 當無發文日(CP27 Is Null)資料時, 才更新是否算案件數CP26為N
                        '93.10.5 CANCEL BY SONIA 改在前一句
                        'strSQL = "UPDATE CASEPROGRESS SET CP26='N' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP27 IS NULL "
                        'cnnConnection.Execute strSQL
                        '93.10.5 END
                                 
                                 'UPDATE CASEPROGRESS SET CP57=20010802,CP58='01' WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS WHERE CP01='FCP' AND CP02='022626')
                        'UPDATE 下一程序檔解除期限日期,解除期限原因
                        '93.10.5 MODIFY BY SONIA 只更新 是否續辦為 NULL 者
                        'strSQL = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & strReasonNo(cboReason.ListIndex) & "' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP11 IS NULL "
                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305 拆至結案單主檔中
                        'strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & strReasonNo(cboReason.ListIndex) & "' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP11 IS NULL AND NP06 IS NULL"
                        strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & strReason & "' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP11 IS NULL AND NP06 IS NULL"
                        'end 2025/06/02
                        '93.10.5 END
                        cnnConnection.Execute strSql
                        ' ADD 到案件進度檔
                           Dim strAutoNum As String
                           'Modify By Cheng 2002/10/01
'                           If objPublicData.GetAutoNumber("B", strAutoNum, True, False) Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAutoNumber("B", strAutoNum, True, True) Then
                           If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
                                CheckOC
                                strSql = "select au01||(au02-1911) from autonumber where au01='B'"
                                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                If Not adoRecordset.BOF Then adoRecordset.MoveFirst
                                If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Sub
                                'Modify By Sindy 2010/8/18 比對自動編號年度
                                'strAutoNum = CheckStr(adoRecordset.Fields(0).Value) & strAutoNum
                                strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNum
                                CheckOC
                                'Modify By Sindy 2015/5/19 +cp140
                                strSql = "insert into caseprogress ( " & _
                                          "cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10," & _
                                          "cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20," & _
                                          "cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30," & _
                                          "cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40," & _
                                          "cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50," & _
                                          "cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60," & _
                                          "cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76," & _
                                          "cp77,cp78,cp79,cp140) values "
                                'Set SCp() = cp()
                                For i = 1 To 79
                                   Select Case i
                                   '文字null
                                    'Modify By Cheng 2002/01/29
'                                   Case 8, 11, 12, 13, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63, 64
                                    'Modify By Cheng 2002/05/29
'                                   Case 8, 11, 12, 13, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                                   '92.1.25 MODIFY BY SONIA 取消收文日及原因要存
                                   'Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                                   Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63
                                   '92.1.25 END
                                        SCp(i) = "null "
                                   'Add By Cheng 2002/05/29
                                   Case 13 '智權人員
                                       'Add By Sindy 2015/1/15
                                       If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                                          m_CP13 = ShowCurrCP13(cp(1), cp(2), cp(3), cp(4), strNation)
                                          SCp(i) = "'" & m_CP13 & "'"
                                       Else
                                       '2015/1/15 END
                                          SCp(i) = "'" & frm110103_1.txtCaseField(4).Text & "'"
                                       End If
                                   Case 12 '業務區
                                       'Add By Sindy 2015/1/15
                                       If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                                          Call ShowCurrCP13(cp(1), cp(2), cp(3), cp(4), strNation, m_CP12)
                                          SCp(i) = "'" & m_CP12 & "'"
                                       Else
                                       '2015/1/15 END
                                          SCp(i) = "'" & GetST15(frm110103_1.txtCaseField(4).Text) & "'"
                                       End If
                                   '文字畫面上
                                   Case 14
                                        SCp(i) = "'" & strUserNum & "'"
                                   Case 1, 2, 3, 4
                                        SCp(i) = "'" & Trim(ChgSQL(cp(i))) & "'"
                                   Case 5, 27
                                        SCp(i) = GetTodayDate
                                   Case 9
                                        SCp(i) = "'" & strAutoNum & "'"
                                   '91.12.6 modify by sonia
                                   'Case 26, 20, 32
                                   '     SCp(i) = "'N'"
                                   Case 20
                                      If intWhere <> "2" Then
                                          SCp(i) = "'N'"
                                          '2013/8/13 add by sonia FMT要請款
                                          'Add By Sindy 2015/1/15
                                          If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                                             If cp(1) = "T" And Left(GetST15(m_CP13), 1) = "F" Then
                                                SCp(i) = "null "
                                             End If
                                          Else
                                          '2015/1/15 END
                                             If cp(1) = "T" And Left(GetST15(frm110103_1.txtCaseField(4).Text), 1) = "F" Then
                                                SCp(i) = "null "
                                             End If
                                             '2013/8/13 end
                                          End If
                                      'Add by Morgan 2007/7/23 改抓CPM設定
                                      'Modify by Amy 2025/06/03 +FG-陳亭妙 ex:FG-001323(亭妙已自行補上cp20=N)
                                      ElseIf cp(1) = "FCP" Or cp(1) = "FG" Then
                                          SCp(i) = CNULL(PUB_GetCP20(cp(1), Replace(SCp(10), "'", "")))
                                      Else
                                          SCp(i) = "null "
                                      End If
                                   Case 26, 32
                                        SCp(i) = "'N'"
                                   '91.12.6 end
                                   'Case 43
                                   '     SCp(i) = "'" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 0) & "'"
                                   Case 10
                                         Select Case Val(CheckSys(cp(1)))
                                         Case 1, 5         'patent
                                            SCp(i) = "'913'"
                                         Case 2, 6         'trademark
                                            SCp(i) = "'704'"
                                         Case 3, 4, 7, 8   'lawcase & hirecase
                                            SCp(i) = "'993'" 'Modify By Sindy 2011/10/26 999=>993.閉卷
                                         Case Else
                                         End Select
                                   'Add By Cheng 2002/01/29
                                   '進度備註
                                   Case 64
                                       'Modify By Sindy 2015/1/14 已在畫面上增加一個進度備註欄
'                                       If Len(Me.cboNote.Text) <= 0 Then
'                                         SCp(i) = "null "
'                                       Else
'                                           'Modify By Cheng 2002/11/26
''                                          SCp(i) = Me.cboNote.Text & " "
'                                         SCp(i) = "'" & Me.cboNote.Text & " " & "'"
'                                       End If
                                       If Len(Me.txtCP64.Text) <= 0 Then
                                         SCp(i) = "null"
                                       Else
                                           'Modify By Cheng 2002/11/26
'                                          SCp(i) = Me.cboNote.Text & " "
                                         SCp(i) = "'" & Me.txtCP64.Text & "'"
                                       End If
                                       '2015/1/14 END
                                   Case 65, 66, 67, 68, 69, 70
                                        SCp(i) = ""
                                   '92.1.25 ADD BY SONIA
                                   Case 57
                                        SCp(i) = ChangeTStringToWString(txtCaseField(0))
                                   Case 58
                                        'Modify By Cheng 2004/04/15
'                                        SCp(i) = strReasonNo(cboReason.ListIndex)
                                        'Modify by Amy 2025/06/16 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
                                        'SCp(i) = IIf(strReasonNo(cboReason.ListIndex) = "", "Null", CNULL(strReasonNo(cboReason.ListIndex)))
                                       'Modify by Amy 2025/07/10 +CNULL 避免數字前面的0被拿掉
                                       SCp(i) = CNULL(GetReason(cboReason))
                                       If SCp(i) = "" Then SCp(i) = "Null"
                                        'End
                                   '92.1.25 END
                                   '數字
                                   Case Else
                                        SCp(i) = "null "
                                   End Select
                                Next i
                                strSql = strSql & " ("
                                For i = 1 To 79
                                    Select Case i
                                    Case 65, 66, 67, 68, 69, 70
                                    Case Else
                                         strSql = strSql & SCp(i)
                                         If i <> 79 Then
                                            strSql = strSql & ","
                                         End If
                                    End Select
                                Next i
                                'Add By Sindy 2015/5/19 +結案單電子化 : CP140
                                If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                                   strSql = strSql & ",'" & mPrev01.txtF0301 & "'"
                                Else
                                   strSql = strSql & ",null"
                                End If
                                '2015/5/19 END
                                strSql = strSql & ") "
                                cnnConnection.Execute strSql
                                
                              'Added by Morgan 2015/11/3 指示信電子化
                              If m_boleOrderLetter Then
                                 'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                                 'strExc(1) = Pub_GetSpecMan("PS4") 'P案指示信判發人
                                strExc(1) = PUB_GetLetterJudgeNew("2", field(1), Replace(SCp(10), "'", ""), field(9))
                                 'end 2018/7/30
                                 PUB_AddAppForm strAutoNum, IIf(txtCaseField(2) = "Y", True, False), strExc(1)
                              End If
                              'end 2015/11/3
                        
                                'Add by Sindy 2013/04/12 更新c類的代理人及彼所案號，要在新增c類之後
                                Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
                                
                                bolLeave = True
                                intLeaveKind = 2
                                Me.Hide
                           Else
                              Screen.MousePointer = vbDefault
                              'edit by nickc 2005/04/22
                               'MsgBox ("自動給號錯誤")
                               'Exit Sub
                               oErrMsg = "自動給號錯誤"
                               GoTo CheckingErr
                           End If
                        bolLeave = True
                        
                        Pub_UpdateEndModCash cp(1), cp(2), cp(3), cp(4)  '2011/10/28 ADD BY SONIA
                        
                        'Add By Sindy 2023/12/13 檢查接洽單的Flow是否要結束
                        Call PUB_UpdateCRLFlowClose(cp(140), cp(9))
          
                        'Add By Sindy 2015/1/14 結案單電子化
                        If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                           bolMailF0202_3 = True 'Add by Amy 2025/08/19
                           intLeaveKind = 0
                           strUpdDate = strSrvDate(1)
                           strUpdTime = Right("000000" & ServerTime, 6)
'                           '記錄電子表單編號
'                           strSql = "update caseprogress set cp140='" & mPrev01.txtF0301 & "' where cp09=" & SCp(9)
'                           cnnConnection.Execute strSql
                           '卷宗區 : 新增一筆結案單Close至卷宗區
                           'Modify By Sindy 2020/2/19 電子檔名,本所案號使用函數 PUB_CaseNo2FileName
'                           strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
'                                    " values(" & SCp(9) & "," & _
'                                            "'" & cp(1) & Val(cp(2)) & IIf(cp(3) = "0" And cp(4) = "00", "", "-" & cp(3)) & IIf(cp(4) = "00", "", "-" & cp(4)) & "." & Replace(SCp(10), "'", "") & "." & EMP_結案單 & ".menu',0,'" & strUserNum & "'," & _
'                                            strUpdDate & "," & strUpdTime & "," & _
'                                            strUpdDate & "," & strUpdTime & ",'Y')"
                           strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                                    " values(" & SCp(9) & ",'" & PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & _
                                            "." & Replace(SCp(10), "'", "") & "." & EMP_結案單 & ".menu',0,'" & strUserNum & "'," & _
                                            strUpdDate & "," & strUpdTime & "," & _
                                            strUpdDate & "," & strUpdTime & ",'Y')"
                           cnnConnection.Execute strSql
                            'Modify by Amy 2021/06/23 職代需註明(代)
                            strSql = ""
                            'Modify by Amy 2025/06/02 目前只有內商人員不是掛在個人
                            'If (field(1) = "CFT" Or field(1) = "CFC" Or field(1) = "S") And strOldF0308 <> MsgText(601) Then
                            If Pub_StrUserSt03 <> "P21" And strOldF0308 <> MsgText(601) Then
                                If strOldF0308 <> strUserNum Then
                                    strSql = " ,F0208='(代)' "
                                End If
                                '葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現
                                'If strUserNum = "78011" Or strUserNum = "80030" Then
                                'Modify by Amy 2025/06/09 + strNation <> "000"
                                'Modify by Amy 2025/08/19 +外商案之外商程序人員=補看人員者(目前st03=F12的二級主管為 湘嫺),補看不出現,也不發信給補看
                                If ((field(1) = "CFT" Or field(1) = "CFC" Or (field(1) = "S" And strNation <> "000")) And (strUserNum = "78011" Or strUserNum = "80030")) _
                                  Or (intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日 And strUserNum = "79020") Then
                            'end 2025/06/02
                                    strCmd(0) = "update FLOW002 set " & _
                                        "F0205='" & strUpdDate & "'" & _
                                        ",F0206='" & strUpdTime & "'" & _
                                        ",F0207='3',F0204='" & strUserNum & "'" & _
                                        " where F0201='" & mPrev01.txtF0301 & "' and F0202='3' and F0207 is null "
                                    strCmd(1) = "Update FLOW003 Set F0309=" & CNULL(Flow_歸檔) & " Where F0301='" & mPrev01.txtF0301 & "'"
                                    bolMailF0202_3 = False 'Add by Amy 2025/08/19
                                End If
                            End If
                           '簽核檔-程序人員:3.已處理
                           strSql = "update FLOW002 set " & _
                                    "F0205='" & strUpdDate & "'" & _
                                    ",F0206='" & strUpdTime & "'" & _
                                    ",F0207='3',F0204='" & strUserNum & "'" & strSql & _
                                    " where F0201='" & mPrev01.txtF0301 & "' and F0202='2' and F0207 is null "
                           cnnConnection.Execute strSql
                           'end 2021/06/23
                           '讀取下一處理人員
                           'Modified by Morgan 2015/11/3 +傳m_boleOrderLetter
                           If GetNextProPerson_Flow(Trim(mPrev01.txtF0301), Trim(mPrev01.m_F0316), strF0308, strF0309, m_boleOrderLetter) = False Then GoTo CheckingErr
                           '流程備註檔
                           If Trim(mPrev01.txtNote.Text) <> "" Then
                              strSql = GetInsertFLOW004Sql(Trim(mPrev01.txtF0301), strUserNum, strUpdDate, strUpdTime, strF0309, ChgSQL(Trim(mPrev01.txtNote.Text)))
                              cnnConnection.Execute strSql
                           End If
                            'Add by Amy 2021/06/23 葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現
                            If strCmd(0) <> MsgText(601) Then
                                cnnConnection.Execute strCmd(0)
                                cnnConnection.Execute strCmd(1)
                            End If
                        End If
                        '2015/1/14 END
                        'Add by Amy 2025/08/08 外專結案單有勾 未付帳款 及有輸 管制催款日,加 未付帳款 行事曆提醒
                        If intFCState = "2" Then
                           If ChkCCD03(1, Me.Name, mPrev01.txtF0301, strNotPay, strCCD08) = True Then
                                strExc(1) = PUB_GetFCPHandler(field(1), field(2), field(3), field(4)) '程序管制人
                                strExc(2) = PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) '承辦案件管制人
                                strExc(0) = strExc(1) & "," & strExc(2)
                                If InStr(strNotPay, "追蹤欠款：") = 0 Then strNotPay = "追蹤欠款：" & strNotPay
                                strCCD08 = Val(strCCD08) + 19110000
                                If PUB_AddFCPStaffCalendar(strCCD08, 1, strExc(0), strNotPay, strExc(0), "1", field(1), field(2), field(3), field(4), strCCD08, , , mPrev01.txtF0301) Then
                                End If
                           End If
                        End If
                        'end 2025/08/08
                        'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯
                        If bolDelCM = True Then
                            strExc(0) = field(1): strExc(1) = field(2): strExc(2) = field(3): strExc(3) = field(4)
                            strExc(4) = "": strExc(5) = "": strExc(6) = "": strExc(7) = ""
                            If PUB_DeleteCaseRelation(strExc, 3) Then
                            End If
                        End If
                        'end 2016/10/19
                        'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：輸入閉卷913自動收文「通知資訊變更961」,發一封Email給承辦工程師
                        If field(1) = "FCP" And m_PA177 = "Y" Then
                           'Memo by Lydia 2025/04/02 模組內已去掉SCp(10)的單引號Replace
                           If PUB_GetFCPlinkMC("6", TransDate(txtCaseField(0), 2), field, strAutoNum, SCp(10)) = True Then
                           End If
                        End If
                        'end 2023/07/28
                                  
                        cnnConnection.CommitTrans
                        bolCommit = True
                        
                        '2015/8/10 add by sonia 專利案件閉卷時有新案翻譯尚未完稿要提醒(FCP-51551)
                        'modify by sonia 2015/9/4 再加cp05>20150101否則舊案無完稿也會有訊息
                        If intCaseKind = 專利 Then
                          strExc(0) = "select cp09,nvl(ep09,0) ep09,nvl(cp27,0) cp27 from caseprogress,engineerprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " and cp10='201' and cp09=ep02(+) and cp05>20150101"
                          intI = 1
                          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                          If intI = 1 Then
                             If Val("" & RsTemp("ep09")) = 0 Then
                                MsgBox "此案新案翻譯進度尚未完稿！"
                             ElseIf Val("" & RsTemp("cp27")) = 0 Then
                                MsgBox "此案新案翻譯進度尚未發文！"
                             End If
                          End If
                        End If
                        '2015/8/10 end
                        
                        'Added by Morgan 2025/3/28
                        If txtCaseField(1) = "" Then
                           bolChk = False
                           If txtCaseField(2) = "Y" Then
                              bolChk = True
                           End If
                           If m_boleOrderLetter Then
                              NowPrint strAutoNum, "15", "00", bolChk, strUserNum, 0, , , , , , , , , , , , strAutoNum
                              If bolChk = True Then
                                 frm1105_1.m_RecNo = strAutoNum
                                 'Modified by Lydia 2025/04/02 去掉案件性質SCp(10)的單引號Replace
                                 frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & Replace(SCp(10), "'", "") & ".DATA.PDF"
                                 frm1105_1.Show
                              End If
                           Else
                              NowPrint strAutoNum, "15", "00", bolChk, strUserNum
                           End If
                        End If
                        'end 2025/3/28
                        
                        '*** 通知補看人員 ***
                        'Add by Amy 2018/06/27 非P結案發mail通知補看人員
                        'Modfy by Amy 2018/08/31 CFP補看人員也不發mail
                        'Modify By Sindy 2025/6/4 內專不只有P,CFP案還有PS,CPS
                        'If UCase(mPrev01.Name) = UCase("frm210149_1") And cp(1) <> "P" And field(1) <> "CFP" Then
                        'Modify by Amy 2025/06/12 +UCase(mPrev01.Name) = UCase("frm210149_1") ,從閉卷進入不會有前畫面的結案單號會錯
                        'if Left(PUB_GetST03(strUserNum), 2) <> "P1" Then '非內專
                        If UCase(mPrev01.Name) = UCase("frm210149_1") And Left(PUB_GetST03(strUserNum), 2) <> "P1" Then '非內專
                        '2025/6/4 END
                           'Add by Amy 2021/06/28+ if 需發mail(葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現,故不需發mail)
                           'Modfiy by Amy 2025/06/02 +外專/外商 案件
                           'If (strCmd(0) = MsgText(601) And (cp(1) = "CFT" Or cp(1) = "CFC" Or cp(1) = "S")) Or (cp(1) <> "CFT" And cp(1) <> "CFC" And cp(1) <> "S") Then
                           'Modify By Sindy 2025/6/4 + And strNation <> "000")
                           'Modify by Amy 2025/06/09 FCT上線延後,故加intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日 判斷
                           'Modify by Amy 2025/08/19 +bolMailF0202_3 是否寄補看人員,改抓變數,目前只要有run frm210149_1 (T延展不會,也不會有補看人員),預設都寄,除非有設定
                           'If (strCmd(0) = MsgText(601) And (cp(1) = "CFT" Or cp(1) = "CFC" Or (cp(1) = "S" And strNation <> "000"))) _
                              Or (cp(1) <> "CFT" And cp(1) <> "CFC" And cp(1) <> "S" And intFCState <> 1) _
                              Or (intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日) Then
                                'Modify by Amy 2021/06/29 +本所案號
                                strTo = GetF0202_3(cp(1), cp(2), cp(3), cp(4))
                           If strSrvDate(1) >= FCT結案單電子化啟用日 And strTo <> "" Then
                                'Memo by Amy 葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現,故不需發-於上面更新Flow002時已設
                                'CF案且補看人員為 葉易雲(78011) 不發
                                If strTo = Pub_GetSpecMan("CFT62") Then
                                   bolMailF0202_3 = False
                                '外商案且補看人員為 湘嫺(79020) 不發,因目前FC補看人員都是湘嫺,故以部門判斷即可-Sindy
                                ElseIf PUB_GetST03(strTo) = "F12" Then
                                   bolMailF0202_3 = False
                                End If
                           End If
                           
                           If bolMailF0202_3 = True Then
                           'end 2025/08/19
                                If strTo <> MsgText(601) Then
                                    'Modify By Sindy 2025/6/4
                                    'PUB_SendMail strUserNum, strTo, "", cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已閉卷！", "如主旨"
                                    strContent = GetEMailContent_Flow(Trim(mPrev01.txtF0301), strSubject)
                                    PUB_SendMail strUserNum, strTo, "", strSubject, strContent
                                    '2025/6/4 END
                                End If
                            End If
                        End If
                        '*** End 通知補看人員 ***
                        
                        'Added by Lydia 2023/06/09 當寰華案在key閉卷按確認時，請判斷是否有相關香港案及澳門案未不續辦/閉卷，若有則發mail
                        If m_bolFMP2 = True And field(1) = "P" And field(9) = "020" Then
                           'Modified by Lydia 2023/06/28 傳入案件性質SCp(10)
                           'Modified by Lydia 2025/04/02 去掉案件性質SCp(10)的單引號Replace
                           Call PUB_CloseMailto013044("1", field(1), field(2), field(3), field(4), Replace(SCp(10), "'", ""))
                        End If
                        Call PUB_SendMailCache
                        'end 2023/06/09
                        
                        'Add by Amy 2025/06/02 外專請款通知,開啟Outlook
                        If intFCState = 2 And strSrvDate(1) >= FCP結案單電子化啟用日 And ChkOutlook.Value = vbChecked Then
                           strExc(9) = Replace(SCp(10), "'", "")
                           'Modify by Amy 2025/07/10 +strOutLookType(依Pub_ChkCloseInvoce函數回傳寄誰)
                           If Pub_CloseOutLook(Me.Name, strF0301, field(1), field(2), field(3), field(4), field(9), strExc(9), strOutLookType, oErrMsg) = False Then
                             If oErrMsg <> "" Then
                                MsgBox oErrMsg
                                If InStr(oErrMsg, "無C類來函掛工程師,不需出草稿") > 0 Then
                                   ChkOutlook.Value = vbUnchecked
                                End If
                             Else
                                MsgBox "開啟Outlook失敗,請洽電腦中心!"
                             End If
                           End If
                        End If
                        'end 2025/06/02
                        
                        'Add by Amy 2025/08/19+有請款資料彈訊息詢問是否輸請款單
                        If bolInvoice = True Then
                           'Modify by Amy 2025/10/20 有請款項目直接開請款單輸入不詢問-薛經理
'                          intI = MsgBox("開啟請款單輸入及Outlook草稿？" & vbCrLf & _
'                                      "是：開啟請款單輸入及Outlook草稿" & vbCrLf & _
'                                      "否：回待處理區列表", vbYesNo + vbDefaultButton2 + vbQuestion)
'                          If intI = vbYes Then
                             mPrev01.QueryData '更新前畫面資料
                             mPrev01.SetButtonEnable (False) '鎖住前畫面按鈕
                             mPrev01.SSTab1.Tab = 1
                             Screen.MousePointer = vbHourglass
                             'Modify by Amy 2025/11/11 原只帶第一個畫面,改外商有輸請款項目前3碼與未付款案件性質的金額相符,直接帶入第二個畫面
                              intLeaveKind = 4 'Add by Amy 2025/10/28
                              bolOpen21H0Ok = Pub_Open21H0(strF0301, Me.Name, mPrev01, field(1), field(2), field(3), field(4), oErrMsg)
                              If bolOpen21H0Ok = False Then
                                 MsgBox oErrMsg, vbExclamation
                                 If InStr(oErrMsg, "通知電腦中心") > 0 Then mPrev01.SetButtonEnable (True) '開放前畫面按鈕
                              End If
                              '開啟Outlook草稿
                              strExc(9) = Replace(SCp(10), "'", "")
                              If Pub_CloseOutLook_T(Me.Name, strF0301, field(1), field(2), field(3), field(4), field(9), strExc(9), "", oErrMsg) = False Then
                                 MsgBox "開啟Outlook失敗,請洽電腦中心!"
                              End If
                              Screen.MousePointer = vbDefault
'                          End If
                           'end 2025/11/11
                        End If
                        'end 2025/08/19
                        'end 2018/06/27
                           
                        If intWhereComeFrom = 2 Then
                           intNowCaseCode = intNowCaseCode + 1
                            'Modified by Lydia 2021/07/05 因為現在已經是下一筆要處理的Index; 桂英7/1閉卷X46746(天獅 ,ex.T-146528),cp13=MCTF03,  Grid按全選,執行時有彈訊息有1筆已閉卷,但是最後一筆沒有閉卷到
                            'If intNowCaseCode >= intTotalCaseCode Then
                            If intNowCaseCode > intTotalCaseCode Then
                               frm110103_2.ReChoose intNowCaseCode, strCaseCode()
                               bolLeave = True
                               Unload Me
                            Else
                               'Modified by Lydia 2021/07/05 因為現在已經是下一筆要處理的Index
                               'If intNowCaseCode = intTotalCaseCode - 1 Then
                               If intNowCaseCode = intTotalCaseCode Then
                                  cmdOK(3).Visible = False
                               End If
                               ReadAllData
                            End If
                        Else
                           Screen.MousePointer = vbDefault
                           Unload Me
                        End If
                        Screen.MousePointer = vbDefault
                        
                        '**************************************************************************************************

                        Screen.MousePointer = varSaveCursor
             Case 1, 2 '回前畫面, 結束
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           'Add By Sindy 2015/1/14 結案單電子化
                           'Modified by Morgan 2021/5/4
                           'If UCase(mPrev01.Name) <> UCase("Nothing") Then
                           If Not mPrev01 Is Nothing Then
                           'end 2021/5/4
                              intLeaveKind = 3
                           Else
                           '2015/1/14 END
                              intLeaveKind = 1
                              frm110103_2.ReChoose intNowCaseCode, strCaseCode()
                           End If
                        End If
                        bolLeave = False
                        Unload Me
             Case 3 '下一筆
                        If MsgBox("你並未存檔，確定到下一筆嗎?", vbYesNo + vbCritical) = vbYes Then
                           intNowCaseCode = intNowCaseCode + 1
                           If intNowCaseCode = intTotalCaseCode Then
                              cmdOK(3).Visible = False
                           End If
                           ReadAllData
                        End If
End Select
 'add by nickc2005/04/22 transation
     Exit Sub
CheckingErr:
    If oErrMsg = "" Then
         MsgBox Err.Description
    Else
         MsgBox oErrMsg
    End If
    'Modify by Amy 2018/06/27 +if
    If bolCommit = False Then cnnConnection.RollbackTrans
End Sub

Private Sub Form_Activate()
'Add By Cheng 2003/02/26
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'Add By Cheng 2002/12/04
If m_blnFirstShow = True Then
    m_blnFirstShow = False
    
    Me.Hide
    If intNowCaseCode = intTotalCaseCode Then
       cmdOK(3).Visible = False
    End If
    ReadAllData
    'Add By Cheng 2003/02/26
    If bolLeave = False Then
        '若系統類別為專利, 顯示不續辦但准通知欄位
        If Val(CheckSys(cp(1))) = 1 Then
            StrSQLa = "Select * From Patent Where " & ChgPatent(cp(1) & cp(2) & cp(3) & cp(4))
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                Me.txtCaseField(3).Text = "" & rsA("PA89").Value
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
        
        
      'Added by Morgan 2015/11/3 指示信電子化
      'P非臺灣案指示信都要彈修改畫面來確認送判的內容
      'Modified by Morgan 2015/12/15 外專程序除外
      'Modified by Morgan 2025/3/28 不必再排除外專--敏莉
      If field(1) = "P" And field(9) <> "000" Then
         txtCaseField(2) = "Y"
         'txtCaseField(2).Enabled = False 'Removed by Morgan 2025/3/28
         
         'Added by Morgan 2021/5/4 可能會有指示信 Ex:P-119291 --蕭茹曣
         Label9.Visible = True
         txtCaseField(1).Visible = True
         'txtCaseField(1) = "N" 'Removed by Morgan 2025/3/28 改預設要出--玲玲
         'end 2021/5/4
         
         'Added by Morgan 2025/3/28
         Label2(0).Visible = False
         Label10.Visible = True
         txtCaseField(2).Visible = True
         'end 2025/3/28
      End If
      'end 2015/11/3
        
    End If
End If
'Memo by Amy 2025/08/05 將[不續辦但准通知] 改為[後續准駁簡單報告]
'     基本上此欄位不應該在閉卷出現,為何出現於此支作業已不可考,避免有特殊狀況需存在-故與淑華確認後保留
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer

If intWhereComeFrom = 2 Then
   j = 1
   For i = 1 To frm110103_2.grdDataList.Rows - 1
          If frm110103_2.grdDataList.TextMatrix(i, 0) <> "" Then
             ReDim Preserve strCaseCode(3, j)
             strCaseCode(0, j) = frm110103_2.grdDataList.TextMatrix(i, 6)
             strCaseCode(1, j) = frm110103_2.grdDataList.TextMatrix(i, 7)
             strCaseCode(2, j) = frm110103_2.grdDataList.TextMatrix(i, 8)
             strCaseCode(3, j) = frm110103_2.grdDataList.TextMatrix(i, 9)
             j = j + 1
          End If
   Next
Else
   ReDim Preserve strCaseCode(3, 1)
   'Modify By Sindy 2015/1/14
   If UCase(mPrev01.Name) = UCase("frm210149_1") Then
      strCaseCode(0, 1) = mPrev01.m_CP01
      strCaseCode(1, 1) = mPrev01.m_CP02
      strCaseCode(2, 1) = mPrev01.m_CP03
      strCaseCode(3, 1) = mPrev01.m_CP04
   Else
   '2015/1/14 END
      strCaseCode(0, 1) = frm110103_1.txtSystem
      If frm110103_1.txtSystem = 馬德里案 Then
         strCaseCode(1, 1) = frm110103_1.txtTFCode(0) + IIf(frm110103_1.txtTFCode(1) = "", "0", frm110103_1.txtTFCode(1))
         strCaseCode(2, 1) = IIf(frm110103_1.txtTFCode(2) = "", "0", frm110103_1.txtTFCode(2))
         strCaseCode(3, 1) = IIf(frm110103_1.txtTFCode(3) = "", "00", frm110103_1.txtTFCode(3))
      Else
         strCaseCode(1, 1) = frm110103_1.txtCode(0)
         strCaseCode(2, 1) = IIf(frm110103_1.txtCode(1) = "", "0", frm110103_1.txtCode(1))
         strCaseCode(3, 1) = IIf(frm110103_1.txtCode(2) = "", "00", frm110103_1.txtCode(2))
      End If
   End If
   j = 1
End If

MoveFormToCenter Me
intTotalCaseCode = j - 1

If intTotalCaseCode = 0 Then cmdOK(3).Enabled = False

intNowCaseCode = 1
bolLeave = False
intLeaveKind = 1
SetDataListWidth
m_blnFirstShow = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'   PUB_SendMailCache 'Add by Morgan 2009/10/15
   If intWhereComeFrom = 1 Then
      Select Case intLeaveKind
         Case 0 '結束
            'Add By Sindy 2015/1/14 結案單電子化
            If UCase(mPrev01.Name) = UCase("frm210149_1") Then
               intFCState = Empty 'Add by Amy 2025/06/02
               'Add by Amy 2023/02/14 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
               Call Pub_ChkLock(3, mPrev01.Name, "D", , Replace(lblCode(0), "-", ""))
               frm210149.Hide
               frm210149.QueryData
               frm210149.Show
               Unload mPrev01
            Else
            '2015/1/14 END
               Unload frm110103_1
            End If
         Case 1 '回前畫面
            frm110103_1.Show
         Case 2
            frm110103_1.Show
            frm110103_1.Cleartxt
         'Add By Sindy 2015/1/14
         Case 3 '電子化:回前畫面
            'Add by Amy 2025/07/31 國外結案單 需開啟frm210149_1的按鈕
            If UCase(mPrev01.Name) = UCase("frm210149_1") And intFCState > 0 Then
               mPrev01.SetButtonEnable (True)
            End If
            'end 2025/0731
            mPrev01.Show
         '2015/1/14 END
      Case 4 '外商電子化開請款單輸入
         '結案單有Run 請款單輸入,關閉此表單,前畫面(frm210149_1)需保留
      End Select
   Else
      If intLeaveKind = 1 Then
         frm110103_2.Show
      Else
         Unload frm110103_2
      End If
   End If
   'Modify by Amy 2025/10/28 +if 結案單要Run 請款單輸入
   'Memo 結案單沒Run 請款單輸入,畫面才需結束(請款單輸入需看frm210149_1資料輸)
   If intLeaveKind <> 4 Then
      Set mPrev01 = Nothing 'Add By Sindy 2015/2/13
   End If
   'Add By Cheng 2002/07/18
   Set frm110103_3 = Nothing
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtCaseField_GotFocus (Index)
End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
CheckKeyIn = -1
Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           '2010/8/3 加val
                           If Val(txtCaseField(intIndex)) <= Val(GetTaiwanTodayDate) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(8003)
                           End If
                         End If
             Case 1
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 2
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   'TextInverse txtCaseField(Index)
   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   If Index = 4 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
             Case 1, 2
                        KeyAscii = UpperCase(KeyAscii)
                        'Added by Morgan 2025/3/28
                        If Index = 1 Then
                           If KeyAscii <> 78 And KeyAscii <> 8 Then
                              KeyAscii = 0
                           End If
                        Else
                           If KeyAscii <> 89 And KeyAscii <> 8 Then
                              KeyAscii = 0
                           End If
                        End If
                        'end 2025/3/28
             Case 3 '不續辦但准通知
                        KeyAscii = UpperCase(KeyAscii)
                        If KeyAscii <> 89 And KeyAscii <> 8 Then
                           KeyAscii = 0
                        End If
End Select
End Sub
Private Sub lblSystem_Change()
'Add By Cheng 2002/05/29
Dim Rs As New ADODB.Recordset

If lblSystem = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If

'Add By Cheng 2002/05/29
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open " Select SK02 From SystemKind Where SK01='" & Me.lblSystem.Caption & "'", cnnConnection, adOpenStatic, adLockReadOnly
If Rs.RecordCount > 0 Then
   If Rs.Fields(0).Value = "1" Then
      Me.txtCaseField(3).Enabled = True
   Else
      Me.txtCaseField(3).Enabled = False
   End If
Else
   Me.txtCaseField(3).Enabled = False
End If
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing

End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(1900, 1000, 1000, 2000, 1200, 1200)
SetGridDataListWidth grdDataList, varGridWidth()
End Sub

Private Function ReadCloseCaseDRst(ByRef intWhere As Integer, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
Dim strSql As String, rsRecordset As New ADODB.Recordset
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPublicData As Object, strSQL1 As String, strSQL2 As String, i As Integer
Dim strSQL1 As String, strSQL2 As String, i As Integer

On Error GoTo ErrHand
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'edit by nickc 2008/05/16 修正一開始就有的錯誤
'strSQL1 = "select decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) s02,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) s03,np13 s04,np14 s05,'' s06 from nextprogress,casepropertymap where np02=cpm01 and np07=cpm02 and np02=" + CNULL(strCode1) + " and np03=" + CNULL(strCode2) + " and np04=" + CNULL(strCode3) + " and np05=" + CNULL(strCode4) & " AND NP06 IS NULL"
'strSQL2 = "select decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01,decode(cp06,null,'',substr(cp06,1,4)-1911||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) s02,decode(cp07,null,'',substr(cp07,1,4)-1911||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) s03,cp08 s04,nvl(cp40,nvl(cp50,cp55)) s05,cp09 s06 from caseprogress,casepropertymap where cp01=cpm01 and cp10=cpm02 and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) & " AND CP27 IS NULL AND CP57 IS NULL AND CP06 IS NOT NULL AND CP09<'C'"
strSQL1 = "select decode(" & CNULL(strNation) & "," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) s02,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) s03,np13 s04,np14 s05,'' s06 from nextprogress,casepropertymap where np02=cpm01 and np07=cpm02 and np02=" + CNULL(strCode1) + " and np03=" + CNULL(strCode2) + " and np04=" + CNULL(strCode3) + " and np05=" + CNULL(strCode4) & " AND NP06 IS NULL"
strSQL2 = "select decode(" & CNULL(strNation) & "," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01,decode(cp06,null,'',substr(cp06,1,4)-1911||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) s02,decode(cp07,null,'',substr(cp07,1,4)-1911||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) s03,cp08 s04,nvl(cp40,nvl(cp50,cp55)) s05,cp09 s06 from caseprogress,casepropertymap where cp01=cpm01 and cp10=cpm02 and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) & " AND CP27 IS NULL AND CP57 IS NULL AND CP06 IS NOT NULL AND CP09<'C'"

strSql = "select s01 案件性質,s02 本所期限,s03 法定期限,s04 機關文號,s05 相關人,s06 收文號 from (" + strSQL1 + " union " + strSQL2 + ") order by s02,s06"
'edit by nickc 2007/02/06 不用 dll 了
'Set ReadCloseCaseDRst = objPublicData.ReadRst(strSQL)
Set ReadCloseCaseDRst = ClsPDReadRst(strSql)
err1:
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
Exit Function
ErrHand:
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
End Function

Private Function ReadReasonOfRelief(ByRef strReasonNo() As String, ByRef strReasonName() As String) As Integer
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo ErrHand

strSql = "select ror01,ror02 from reasonofrelief"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   Do While Not rsRecordset.EOF
         ReDim Preserve strReasonNo(i) As String
         ReDim Preserve strReasonName(i) As String
         strReasonNo(i) = rsRecordset.Fields(0)
         strReasonName(i) = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         i = i + 1
         rsRecordset.MoveNext
   Loop
   ReadReasonOfRelief = 1
Else
   ReadReasonOfRelief = 0
End If
Exit Function
ErrHand:
ShowMsg MsgText(8001)
ReadReasonOfRelief = -1
End Function

'Added by Morgan 2015/5/18
'FCP台灣新型年費解除期限一案兩請提醒
Private Sub CheckFCPDualCase()
   
   If field(1) = "FCP" And field(9) = "000" And field(8) = "2" Then
      '若發明案尚未審定或核駁且未閉卷時，提醒使用者
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) " & _
         " from (select cm05,cm06,cm07,cm08,cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
         " union select cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
         ",patent,nextprogress,caseprogress where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1' AND pa57 is null and (pa16 is null or pa16='2')" & _
         " and np02(+)=cm01 and np03(+)=cm02 and np04(+)=cm03 and np05(+)=cm04 and np06(+) is null and np07(+)='605'" & _
         " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and cp27(+) is null and cp10(+)='605' and cp57(+) is null and np01||cp09 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(2) = "此案為一案兩請且發明案 " & RsTemp(0) & " 尚未審定，請將卷宗交業務承辦告知客戶新型專利權若因未繳年費而當然消滅者，則將不予專利！"
         MsgBox strExc(2), vbExclamation, "一案兩請新型結案提醒"
      End If
      
   End If

End Sub

'Add by Amy 2025/06/02 原因
Private Function GetReason(ByVal stCbo As String) As String
   If stCbo <> "" Then
      GetReason = Mid(stCbo, 1, Val(InStr(stCbo, "--")) - 1)
   End If
End Function
