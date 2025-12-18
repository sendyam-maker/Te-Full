VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210124_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "定稿報價查詢(修改)"
   ClientHeight    =   6684
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8748
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6684
   ScaleWidth      =   8748
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   30
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   7875
      Begin MSForms.TextBox txtOldAddr 
         Height          =   300
         Left            =   1200
         TabIndex        =   37
         Top             =   120
         Width           =   6492
         VariousPropertyBits=   671105055
         MaxLength       =   180
         Size            =   "11451;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNewAddr 
         Height          =   300
         Left            =   1200
         TabIndex        =   36
         Top             =   450
         Width           =   6495
         VariousPropertyBits=   671105055
         MaxLength       =   180
         Size            =   "11451;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "原註冊地址："
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "新地址："
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   450
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   0
      TabIndex        =   19
      Top             =   5370
      Width           =   8655
      Begin VB.CheckBox Check4 
         Caption         =   "境外公司"
         Height          =   285
         Left            =   2640
         TabIndex        =   34
         Top             =   0
         Width           =   1965
      End
      Begin VB.TextBox txtCRL116 
         Height          =   300
         Left            =   4800
         MaxLength       =   60
         TabIndex        =   26
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCRL115 
         Height          =   300
         Left            =   2310
         MaxLength       =   20
         TabIndex        =   25
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txtCRL114 
         Height          =   300
         Left            =   450
         MaxLength       =   20
         TabIndex        =   24
         Top             =   960
         Width           =   1245
      End
      Begin VB.CheckBox Check3 
         Caption         =   "同上"
         Height          =   285
         Left            =   810
         TabIndex        =   23
         Top             =   645
         Width           =   675
      End
      Begin VB.TextBox txtCRL99 
         Height          =   300
         Left            =   810
         MaxLength       =   10
         TabIndex        =   20
         Top             =   0
         Width           =   1275
      End
      Begin MSForms.TextBox txtCRL100 
         Height          =   300
         Left            =   1500
         TabIndex        =   22
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12568;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCRL101 
         Height          =   300
         Left            =   810
         TabIndex        =   21
         Top             =   315
         Width           =   7815
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13785;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "特殊ID統 一編號輸8個0"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   5670
         TabIndex        =   33
         Top             =   60
         Width           =   2205
      End
      Begin VB.Label Label4 
         Caption         =   "E-Mail(財務):"
         Height          =   225
         Index           =   6
         Left            =   3720
         TabIndex        =   32
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "傳真:"
         Height          =   225
         Index           =   5
         Left            =   1830
         TabIndex        =   31
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "電話:"
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   30
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "郵寄地址:"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "營業地址:"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   345
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "統一編號:"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   60
         Width           =   1005
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "以 DEBIT NOTE 請款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   5010
      Width           =   2370
   End
   Begin VB.CheckBox Check2 
      Caption         =   "請貼印花"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2715
      TabIndex        =   17
      Top             =   5010
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "參考報價(&R)"
      Height          =   400
      Index           =   3
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Index           =   2
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5535
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4725
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2205
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2310
      Width           =   6270
      _ExtentX        =   11070
      _ExtentY        =   3895
      _Version        =   393216
      BackColor       =   -2147483633
      Cols            =   3
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   405
      Left            =   1110
      TabIndex        =   15
      Top             =   4560
      Width           =   5325
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9393;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDesc 
      Height          =   1185
      Left            =   3600
      TabIndex        =   43
      Top             =   570
      Width           =   2760
      Caption         =   "lblFM2"
      Size            =   "4868;2090"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblMemo 
      Height          =   405
      Left            =   870
      TabIndex        =   42
      Top             =   1860
      Width           =   5595
      Caption         =   "lblFM2"
      Size            =   "9869;714"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblProperty 
      Height          =   255
      Left            =   1170
      TabIndex        =   41
      Top             =   1560
      Width           =   1905
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3360;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   1320
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4075;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   16
      Top             =   4650
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "備註："
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   14
      Top             =   1830
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   1578
      Width           =   900
   End
   Begin VB.Label lblCaseNo 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1620
      TabIndex        =   12
      Top             =   1074
      Width           =   1500
   End
   Begin VB.Label lblCountry 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1215
      TabIndex        =   11
      Top             =   822
      Width           =   1905
   End
   Begin VB.Label lblAppName 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1035
      TabIndex        =   10
      Top             =   570
      Width           =   2490
   End
   Begin VB.Label Label4 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   9
      Top             =   1326
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "本所號/分所號："
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   1074
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   7
      Top             =   822
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   255
      Left            =   270
      TabIndex        =   6
      Top             =   570
      Width           =   720
   End
End
Attribute VB_Name = "frm210124_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2022/01/04 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblCaseName、lblProperty、lblMemo、lblDesc、txtCRL100、txtCRL101、Combo1
'Memo by Lydia 2019/07/01 表單名稱:專業部定稿報價查詢(修改)=>定稿報價查詢(修改)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/4/23
Option Explicit

Dim ii As Integer, jj As Integer '迴圈共用
Dim m_iRow As Integer '本次點選列數
Dim m_iCol As Integer '本次點選行數

Public m_LC01 As String
Public m_LC02 As String
Public m_iRowID  As Integer
'Added by Lydia 2016/05/27
Dim strCP(0 To 3) As String '本所案號
Dim m_Nation As String '國別代號
Dim m_Kind As String '專利/商標 種類
Dim m_Kind2 As String  'Added by Lydia 2016/11/11 目前案件准駁
Dim m_PA21TM20 As String 'Added by Lydia 2017/02/13 發證日


Private Sub cmdOK_Click(Index As Integer)
   Dim stTO As String 'Add by Amy 2024/05/15
   
   Select Case Index
      Case 0, 1, 2
         If Index = 0 Then
            If TxtValidate = False Then Exit Sub
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            If FormSave = False Then
               Screen.MousePointer = vbDefault
               MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
               Exit Sub
               
            'Add By Sindy 2019/9/17
            Else
               If Combo1.Enabled = True And Me.Frame1.Visible = True Then
                  'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
                  If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
                      stTO = Pub_GetSpecMan("財務處應收處理人員")
                  Else
                     stTO = Pub_GetSpecMan("財務處總帳人員")
                  End If
                  PUB_SendMail strUserNum, stTO, "", _
                     lblCaseNo & "之" & lblProperty & "(" & frm210124_1.Tag & ")指定收據抬頭之相關資料！（非客戶新抬頭，請依此資料建立收據抬頭基本資料）", _
                     "收據抬頭:" & Combo1.Text & vbCrLf & _
                     "營業地址:" & txtCRL101.Text & vbCrLf & _
                     "郵寄地址:" & IIf(Check3.Value = 1, Check3.Caption, txtCRL100.Text) & vbCrLf & _
                     "統一編號:" & txtCRL99.Text & vbCrLf & _
                     "是否境外公司:" & IIf(Check4.Value = 1, "是", "否") & vbCrLf & _
                     "電 話:" & txtCRL114.Text & vbCrLf & _
                     "傳 真:" & txtCRL115.Text & vbCrLf & _
                     "財務E -MAIL:" & txtCRL116.Text & vbCrLf
                     'end 2024/05/15
               End If
            End If
            '2019/9/17 END
            Screen.MousePointer = vbDefault
         End If
         frm210124.Tag = Index
         Unload Me
      '參考報價
      Case 3
         GetPrePrice
   End Select
End Sub

'Modified by Lydia 2016/05/27 改成可以不顯示表單
'Private Sub GetPrePrice()
Private Sub GetPrePrice(Optional ByVal bolFormShow As Boolean = True, Optional ByRef PreList As String)
   Dim stAppNo As String, stAppCountry As String, strAppKind As String, strProperty As String, strDate As String
   Dim strKey As String, strSys As String
   Dim dblYear As Double 'Add by Morgan 2010/1/11 繳費年度
   'Added by Lydia 2016/05/27
   Dim rsAD As New ADODB.Recordset
   Dim inR As Integer
   
   PreList = ""
   'end 2016/05/27
   
   If m_iRowID > 0 Then
      With frm210124.grdDataList
         'Modified by Morgan 2015/7/1 加收據抬頭欄位,索引+1
         strKey = .TextMatrix(m_iRowID, 10) & .TextMatrix(m_iRowID, 11)
         stAppNo = .TextMatrix(m_iRowID, 12)
         stAppCountry = .TextMatrix(m_iRowID, 13)
         strAppKind = .TextMatrix(m_iRowID, 14)
         strProperty = .TextMatrix(m_iRowID, 15)
         strDate = .TextMatrix(m_iRowID, 16)
         strSys = .TextMatrix(m_iRowID, 17)
         dblYear = Val("" & .TextMatrix(m_iRowID, 21)) 'Add by Morgan 2010/1/11 繳費年度
          'Add by Lydia 2014/10/29 　LC02="0"＝＞為自動發證國家之證書號輸入(frm05010403_2)時產生，因為無下一程序所以預設LC02=NP22=0
         'If .TextMatrix(m_iRowID, 11) = "0" Then strSys = "3" 'Removed by Morgan 2015/7/2 考慮商標及TC也有自動發證,改前畫面直接改Grid的值
         'Modified by Lydia 2016/05/27
         'If PUB_GetOldPrice(stAppNo, stAppCountry, strAppKind, strProperty, RsTemp, strDate, strKey, strSys, dblYear) = True Then
         '   Set frm880014.grdDataList.Recordset = RsTemp
         '   Set frm880014.fmParent = Me
         '   frm880014.Show vbModal
         'End If
         If PUB_GetOldPrice(stAppNo, stAppCountry, strAppKind, strProperty, rsAD, strDate, strKey, strSys, dblYear) = True Then
            If bolFormShow Then '開啟表單
               Set frm880014.grdDataList.Recordset = rsAD
               Set frm880014.fmParent = Me
               frm880014.Show vbModal
            Else               '讀取資料
               If rsAD.RecordCount > 0 Then
                  rsAD.MoveFirst
                  Do While Not rsAD.EOF
                     PreList = PreList & rsAD("報價日") & " " & convForm(rsAD("本所案號"), 15) & " " & convForm(rsAD("案件名稱"), 30) & " " & convForm(rsAD("項目"), 10) & " " & convForm(rsAD("年度"), 4) & " " & convForm(rsAD("報價"), 15) & vbCrLf
                     rsAD.MoveNext
                  Loop
               End If
            End If
         End If
      End With
   End If
End Sub

Private Sub Combo1_LostFocus()
   'Add By Sindy 2019/9/17 若有輸入收據抬頭且字數>=4，
   '請檢查若不存在於客戶檔及抬頭檔則加畫面讓使用者輸入收據抬頭的相關資料
   If Combo1.Tag <> Combo1 Then
      If Len(Trim(Combo1.Text)) >= 4 Then
         If PUB_ChkTitleNmExist(Combo1.Text, False) = "" Then
            Frame1.Visible = True
            Me.Width = 8835
            'Modified by Lydia 2022/01/04 Me.Height = 6930=>7065
            Me.Height = 7065
            txtCRL99.SetFocus
         Else
            Frame1.Visible = False
            Me.Width = 6615
            'Modified by Lydia 2022/01/04 Me.Height = 5700=>5775
            Me.Height = 5700
         End If
      End If
   End If
   '2019/9/17 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtInput.Visible = False
   'Added by Lydia 2016/05/27 抓案件資料
   If m_LC01 <> "" Then GetData
End Sub

Private Sub SetBox()
   Dim lngLeft As Long, lngTop As Long
   With grdDataList
      If .row > 0 And (.col = 2) Then
         If .TextMatrix(.row, 0) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            m_iRow = .row: m_iCol = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
            If .TextMatrix(.row, 0) = "點數" Then
               txtInput.Locked = True
            Else
               txtInput.Locked = False
            End If
         End If
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210124_1 = Nothing
End Sub

Private Sub GrdDataList_Click()
   With grdDataList
      .row = .MouseRow
      .col = .MouseCol
      'Modify by Morgan 2008/11/17 +判斷點數有值且可修改的才可輸入
      If .TextMatrix(.row, 5) = "" Then
         SetBox
      End If
   End With
End Sub

Private Sub grdDataList_Scroll()
   If txtInput.Visible = True Then
      txtInput_LostFocus
   End If
End Sub

Private Sub txtCRL100_GotFocus()
   OpenIme
   TextInverse txtCRL100
End Sub

'Modified by Lydia 2022/01/04 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtCRL100_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub txtCRL100_Validate(Cancel As Boolean)
   If txtCRL100.Enabled = False Then Exit Sub

   If txtCRL100.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL100, txtCRL100.MaxLength) Then
      Call txtCRL100_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL101_GotFocus()
   OpenIme
   TextInverse txtCRL101
End Sub

'Modified by Lydia 2022/01/04 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtCRL101_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub txtCRL101_Validate(Cancel As Boolean)
   If txtCRL101.Enabled = False Then Exit Sub

   If txtCRL101.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL101, txtCRL101.MaxLength) Then
      Call txtCRL101_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL114_GotFocus()
   CloseIme
   TextInverse txtCRL114
End Sub
Private Sub txtCRL115_GotFocus()
   CloseIme
   TextInverse txtCRL115
End Sub
Private Sub txtCRL116_GotFocus()
   CloseIme
   TextInverse txtCRL116
End Sub
Private Sub txtCRL116_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub txtCRL116_Validate(Cancel As Boolean)
   If txtCRL116.Enabled = False Then Exit Sub
   
   If txtCRL116.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(txtCRL116.Text)
End Sub

Private Sub txtCRL99_GotFocus()
   TextInverse txtCRL99
   CloseIme
End Sub

Private Sub txtCRL99_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRL99_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If txtCRL99.Enabled = False Then Exit Sub

   If txtCRL99.Text = "" Then Exit Sub
   If Trim(txtCRL99) = "境外" Then Exit Sub 'Add by Amy 2016/05/23
   
   txtCRL99.Text = Trim(PUB_StringFilter(txtCRL99.Text)) 'Add By Sindy 2014/4/11 瑞婷反應智權同仁在複製貼上時多貼到空白格
   If GetTextLength(txtCRL99.Text) <> 8 Then
      Call txtCRL99_GotFocus
      strTmp = "統編必須是8碼 ! 請確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(1, txtCRL99.Text) = False Then
      Call txtCRL99_GotFocus
      strTmp = "統一編號錯誤，是否確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         If UpdateVar = True Then
            GoNext
         Else
            TextInverse txtInput
         End If
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Sub GoNext()
   With grdDataList
      .col = 2
      If .row < .Rows - 1 Then
         .row = .row + 1
         If .TextMatrix(.row, 5) = "" Then
            SetBox
         Else
            txtInput.Visible = False
            cmdOK(0).SetFocus
         End If
      Else
         txtInput.Visible = False
         cmdOK(0).SetFocus
      End If
      
   End With
End Sub

Public Sub SetDataListWidth()
   With grdDataList
      .FormatString = "項目|專業部金額|智權部金額|點數"
      .FixedCols = 1
      .RowHeightMin = txtInput.Height
      ii = 0
      .ColWidth(ii) = 2400
      .ColAlignment(ii) = flexAlignLeftCenter
      
      ii = ii + 1
      .ColWidth(ii) = 1400
      .ColAlignment(ii) = flexAlignRightCenter
      
      ii = ii + 1
      .ColWidth(ii) = 1400
      .ColAlignment(ii) = flexAlignRightCenter
      
      ii = ii + 1
      .ColWidth(ii) = 600
      .ColAlignment(ii) = flexAlignRightCenter
            
      For ii = ii + 1 To .Cols - 1
         .ColWidth(ii) = 0
      Next
   End With
   SetGridColor
End Sub

Private Sub SetGridColor()
   Dim lngColor As Long
   With grdDataList
      lngColor = vbWhite
      For ii = 1 To .Rows - 1
         .row = ii
         'Modify by Morgan 2008/11/17 +判斷可修改的才變色
         If .TextMatrix(ii, 5) = "" Then
            .col = 2: .CellBackColor = lngColor
         End If
      Next
   End With
End Sub

Private Function TxtValidate() As Boolean
'Added by Lydia 2016/05/27
Dim dblFeeChange As Double, dblNewPt As Double
Dim tmpCp10 As String
'end 2016/05/27
'Added by Lydia 2016/10/05
Dim stCont1 As String 'CFP控制年費項目合成一封信
Dim bolMsgChk As Boolean 'CFP控制年費詢問
Dim stSub1 As String '催XXX
Dim bCancel As Boolean
Dim tmpBol As Boolean 'Added by Lydia 2022/05/30

   If strCP(0) = "" Then GetData 'Added by Lydia 2016/05/27
   
   'Move by Lydia 2016/06/04 從下面移上來
   'Added by Morgan 2015/7/2
   If Combo1.Enabled = True Then
      If Trim(Combo1) = "" Then
         MsgBox "收據抬頭不可空白!!", vbCritical
         Exit Function
      'Add By Sindy 2019/9/17
      Else
         If Me.Frame1.Visible = True Then
            '收據抬頭>=4碼時則統一編號欄必須有值
            '營業地址及郵寄地址不可空白
            If txtCRL99.Enabled = True And Len(Trim(txtCRL99.Text)) = 0 Then
               MsgBox "請輸入統一編號！", vbExclamation
               txtCRL99.SetFocus
               Exit Function
            End If
            '統一編號欄位不是"境外"才檢查
            If txtCRL101.Enabled = True And Trim(txtCRL101.Text) = "" And Trim(txtCRL99.Text) <> "境外" Then
               MsgBox "請輸入營業地址！", vbExclamation
               txtCRL101.SetFocus
               Exit Function
            End If
            If txtCRL100.Enabled = True And Trim(txtCRL100.Text) = "" And Check3.Value = 0 Then
               MsgBox "請輸入郵寄地址！", vbExclamation
               txtCRL100.SetFocus
               Exit Function
            End If
            If txtCRL114.Enabled = True And Len(Trim(txtCRL114.Text)) = 0 Then
               MsgBox "請輸入電話！", vbExclamation
               txtCRL114.SetFocus
               Exit Function
            End If
            txtCRL99_Validate bCancel
            If bCancel = True Then
               txtCRL99.SetFocus
               Exit Function
            End If
            txtCRL101_Validate bCancel
            If bCancel = True Then
               txtCRL101.SetFocus
               Exit Function
            End If
            txtCRL100_Validate bCancel
            If bCancel = True Then
               txtCRL100.SetFocus
               Exit Function
            End If
            txtCRL116_Validate bCancel
            If bCancel = True Then
               txtCRL116.SetFocus
               Exit Function
            End If
         End If
      End If
   End If
   'end 2015/7/2
   
   Pub_Send_CFPdg = False 'Added by Lydia 2016/09/02
   
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 2) = "" Then
            MsgBox "[" & .TextMatrix(ii, 0) & "]尚未輸入，請確認！"
            Exit Function
         'Added by Lydia 2016/05/27 CFP控制年費.延展費及維持費智權同仁可加的點數
         Else
            dblFeeChange = Val(.TextMatrix(ii, 2)) - Val(.TextMatrix(ii, 1))
            dblNewPt = Val(.TextMatrix(ii, 3))
            tmpCp10 = ""
            'Added by Lydia 2016/09/29 EPC進入各國年費分別判斷
            'Modified by Lydia 2016/11/11 母案核准後才進入國家階段,尚未核准時以一般個案判斷
            'If m_Nation = "221" Then
            'Modified by Lydia 2017/02/13 進入國家階段以發證日為準 (BY 甄妮)
            'If m_Nation = "221" And m_Kind2 = "1" Then
            If m_Nation = "221" And m_PA21TM20 <> "" Then
                If Right(Trim(.TextMatrix(ii, 0)), 2) = "年費" Then
                   tmpCp10 = "605"
                End If
                strExc(7) = Trim(lblCountry) & "進入" & Mid(Trim(.TextMatrix(ii, 0)), 1, Len(Trim(.TextMatrix(ii, 0))) - 2) 'EPC各國
                If tmpCp10 <> "" Then stSub1 = "年費" 'Added by Lydia 2016/10/05
            Else '原: 非EPC
                Select Case Trim(.TextMatrix(ii, 0))
                    Case "年費": tmpCp10 = "605"
                    Case "維持費": tmpCp10 = "606"
                    Case "延展費": tmpCp10 = "607"
                    Case Else: tmpCp10 = ""
                End Select
                strExc(7) = Trim(lblCountry) 'Added by Lydia 2016/09/29
                If tmpCp10 <> "" Then stSub1 = stSub1 & IIf(stSub1 <> "", "、", "") & Trim(.TextMatrix(ii, 0))  'Added by Lydia 2016/10/05
            End If
            'end 2016/09/29
            
            
            'Added by Lydia 2016/10/05 同一筆進度的所有報價資料都要顯示
            If strCP(0) = "CFP" And tmpCp10 <> "" Then
                'Move by Lydia 2016/10/05 移到上面抓報價資料
                '抓下一程序的本所期限;因為馬來西亞新型605+607和俄羅斯設計605+607可能有合併催函的情況,所以用案件性質抓.
                strExc(8) = ""
                strSql = "select np08 from nextprogress where np01='" & m_LC01 & "' and np06 is null and np07='" & tmpCp10 & "' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then strExc(8) = "" & RsTemp(0)
                
                strExc(5) = PUB_GetCPMnexttimes(strCP(0), strCP(1), strCP(2), strCP(3), m_Nation, m_Kind, tmpCp10, tmpCp10)
                Call GetPrePrice(False, strExc(6)) '讀取參考報價
                stCont1 = stCont1 & "申請國家：" & strExc(7) & vbCrLf & _
                                    "案件性質：" & Trim(strExc(5)) & vbCrLf & _
                                    "本所期限：" & IIf(strExc(8) <> "", ChangeWStringToTDateString(strExc(8)), "") & vbCrLf & _
                                    "程序報價：" & Format(Val(.TextMatrix(ii, 1)), DDollar2) & Mid(.TextMatrix(ii, 1), InStr(.TextMatrix(ii, 1), "(")) & vbCrLf & _
                                    "智權報價：" & Format(Val(.TextMatrix(ii, 2)), DDollar2) & "(" & dblNewPt & ")" & vbCrLf & _
                                    IIf(Len(strExc(6)) > 0, "參考報價：" & vbCrLf & strExc(6), "") & vbCrLf & _
                                    IIf(Val(.TextMatrix(ii, 1)) <> Val(.TextMatrix(ii, 2)), "此案欲調整點數為" & dblNewPt & "點，請您批示！若您同意請轉寄郵件給預設程序人員, 謝謝 !" & vbCrLf, "") & vbCrLf
                                    
                'Added by Lydia 2016/10/11 加區隔線
                If ii <> .row Then stCont1 = stCont1 & String(100, "-") & vbCrLf & vbCrLf

            End If
            'end 2016/10/05
                
            'Modified by Lydia 2016/10/11 只問一次 +bolMsgChk=false
            If strCP(0) = "CFP" And Abs(dblFeeChange) <> 0 And bolMsgChk = False And ((tmpCp10 = "605" And dblNewPt > CFP_dg605) Or (tmpCp10 = "606" And dblNewPt > CFP_dg606) Or (tmpCp10 = "607" And dblNewPt > CFP_dg607)) Then
                'Move by Lydia 2016/10/05 移到上面抓報價資料

                '若有調整點數時並超過點數上限發E-mail給主管批示，若選否則不可存檔；若選擇是則開E-MAIL畫面
                If MsgBox(Trim(.TextMatrix(ii, 0)) & "已超過點數上限" & IIf(tmpCp10 = "605", CFP_dg605, IIf(tmpCp10 = "606", CFP_dg606, CFP_dg607)) & "點，是否發E-mail給主管批示？", vbCritical + vbYesNo, "CFP控制點數") = vbYes Then
                   bolMsgChk = True 'Added by Lydia 2016/10/05
                   'Move by Lydia 2016/10/05 移到最後寄信
                Else
                   Exit Function
                End If
            'Added by Lydia 2016/06/04 因為自動發證國的領證費要開收據,所以業務修改點數要發mail通知給程序(輸入人員)
            ElseIf m_LC02 = "0" And Abs(dblFeeChange) <> 0 Then
                   'modify by sonia 2016/10/28 +cp10
                   strSql = "select cp65,cp10 from caseprogress where cp09=" & CNULL(m_LC01)
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                   If intI = 1 Then
                      strExc(0) = strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & " " & Trim(lblProperty) & "修改報價點數通知！"
                      'modify by sonia 2016/10/28 T,CFT,TC之註冊證發信內容不同
                      'strExc(1) = "本所案號：" & strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & vbCrLf & _
                                  "案件名稱：" & Trim(lblCaseName) & vbCrLf & _
                                  "申請人：" & Trim(lblAppName) & vbCrLf & _
                                  "申請國家：" & Trim(lblCountry) & vbCrLf & _
                                  "案件性質：" & Trim(lblProperty) & vbCrLf & _
                                  "原始報價點數：" & Trim(Val(Mid(.TextMatrix(ii, 1), InStr(.TextMatrix(ii, 1), "(") + 1))) & vbCrLf & _
                                  "業務修改點數：" & Trim(Val(.TextMatrix(ii, 3))) & vbCrLf & vbCrLf & _
                                  "請確認後修改進度檔之費用及點數!"
                      strExc(1) = "本所案號：" & strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & vbCrLf & _
                                  "案件名稱：" & Trim(lblCaseName) & vbCrLf & _
                                  "申請人：" & Trim(lblAppName) & vbCrLf & _
                                  "申請國家：" & Trim(lblCountry) & vbCrLf & _
                                  "案件性質：" & Trim(lblProperty) & vbCrLf & _
                                  "原始報價點數：" & Trim(Val(Mid(.TextMatrix(ii, 1), InStr(.TextMatrix(ii, 1), "(") + 1))) & vbCrLf & _
                                  "業務修改點數：" & Trim(Val(.TextMatrix(ii, 3))) & vbCrLf & vbCrLf
                      If (strCP(0) = "T" Or strCP(0) = "TC" Or strCP(0) = "CFT") And RsTemp(1) = "1701" Then
                          strExc(1) = strExc(1) & "商標註冊證案，系統已依業務修改點數更新進度檔之費用及點數!"
                      Else
                          strExc(1) = strExc(1) & "請確認後修改進度檔之費用及點數!"
                      End If
                      'end 2016/10/28
                      PUB_SendMail strUserNum, "" & RsTemp(0), "", strExc(0), strExc(1)
                   End If
            'end 2106/06/04
            End If
         'end 2016/05/27
         End If
      Next
   End With
   
   'Move by Lydia 2016/10/05 移到最後寄信
   If bolMsgChk = True Then
        strExc(1) = ""
        strExc(1) = GetFLOW001Person(Trim(frm210124.txtSales.Text), Flow_接洽單)
        If strExc(1) <> "" Then
           strSql = "select lc10,lc11 from lettercache where lc01='" & m_LC01 & "' and lc02='" & m_LC02 & "'"
           intI = 1: strExc(2) = "": strExc(3) = ""
           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
              strExc(2) = "" & RsTemp.Fields("lc10")
              strExc(3) = "" & RsTemp.Fields("lc11")
           End If
           strExc(3) = CompWorkDay(5, strExc(3)) '定稿暫存檔建檔日+4個工作天
           strExc(5) = PUB_GetCPMnexttimes(strCP(0), strCP(1), strCP(2), strCP(3), m_Nation, m_Kind, tmpCp10, tmpCp10)
           Call GetPrePrice(False, strExc(6)) '讀取參考報價
           'Modified by Lydia 2022/05/30 改用frm880019
           'frm880005.bolCCList = True
           'frm880005.txtEmail(0) = strExc(1)
           'frm880005.txtEmail(0).Tag = "CFPdg" 'Added by Lydia 2016/09/02
           ''Modified by Lydia 2016/10/05
           ''frm880005.txtEmail(1) = strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & " 催" & Trim(.TextMatrix(ii, 0)) & "期限報價提高點數請示！請於" & ChangeWStringToTDateString(strExc(3)) & "前批示，謝謝！"
           'frm880005.txtEmail(1) = strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & " 催" & stSub1 & "期限報價提高點數請示！請於" & ChangeWStringToTDateString(strExc(3)) & "前批示，謝謝！"
           ''Modified by Lydia 2016/08/31 報價金額後方要加點數
           ''Modified by Lydia 2016/09/29 Trim(lblCountry) =>strExc(7)
           ''Modified by Lydia 2016/10/05 CFP控制年費項目合成一封信
           ''frm880005.txtEmail(2) = "本所案號：" & strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & vbCrLf & _
                                   "案件名稱：" & Trim(lblCaseName) & vbCrLf & _
                                   "申請人：" & Trim(lblAppName) & vbCrLf & _
                                   "申請國家：" & strExc(7) & vbCrLf & _
                                   "案件性質：" & Trim(strExc(5)) & vbCrLf & _
                                   "本所期限：" & IIf(strExc(8) <> "", ChangeWStringToTDateString(strExc(8)), "") & vbCrLf & _
                                   "程序報價：" & Format(Val(.TextMatrix(ii, 1)), DDollar2) & Mid(.TextMatrix(ii, 1), InStr(.TextMatrix(ii, 1), "(")) & vbCrLf & _
                                   "智權報價：" & Format(Val(.TextMatrix(ii, 2)), DDollar2) & "(" & dblNewPt & ")" & vbCrLf & _
                                   IIf(Len(strExc(6)) > 0, "參考報價：" & vbCrLf & strExc(6), "" & vbCrLf) & vbCrLf & _
                                   "此案欲調整點數為" & dblNewPt & "點，請您批示！若您同意請轉寄郵件給" & GetStaffName(strExc(2)) & ", 謝謝 !"
           'frm880005.txtEmail(2) = "本所案號：" & strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & vbCrLf & _
                                   "案件名稱：" & Trim(lblCaseName) & vbCrLf & _
                                   "申請人：" & Trim(lblAppName) & vbCrLf & Replace(stCont1, "預設程序人員", GetStaffName(strExc(2)))
           'frm880005.Show vbModal
           ''注意在測試時,若沒有按傳送,需要手動跳到已發送狀態,不然會多檢查一次
           ''Modified by Lydia 2016/08/31 frm880005.bolLeave在Unload時，會變false,改設在共用變數
           'If Pub_Send_CFPdg = False Then
           '   Exit Function
           'End If
           frm880019.txtReceiver = strExc(1)
           frm880019.txtSubject = strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & " 催" & stSub1 & "期限報價提高點數請示！請於" & ChangeWStringToTDateString(strExc(3)) & "前批示，謝謝！"
           frm880019.txtContent = "本所案號：" & strCP(0) & "-" & Val(strCP(1)) & IIf(strCP(2) & strCP(3) = "000", "", "-" & strCP(2) & "-" & strCP(3)) & vbCrLf & _
                                   "案件名稱：" & Trim(lblCaseName) & vbCrLf & _
                                   "申請人：" & Trim(lblAppName) & vbCrLf & Replace(stCont1, "預設程序人員", GetStaffName(strExc(2)))
           frm880019.cmdAttach.Visible = False
           frm880019.SetParent Me
           frm880019.Show vbModal
           tmpBol = frm880019.m_bolDone '是否傳送成功
           Unload frm880019
           If tmpBol = False Then
                MsgBox "送信失敗，請重新Email !", vbCritical
                Exit Function
           End If
           'end 2022/05/30
        End If
   End If
   'end 2016/10/05
   
    'Added by Lydia 2022/01/04 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If txtCRL100 <> "" Or txtCRL101 <> "" Then
        If PUB_ChkUniText(Me, , True, "TextBox") = False Then
            Exit Function
        End If
    End If
    If Combo1.Text <> "" Then
        If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
            Exit Function
        End If
    End If
    'end 2022/01/04
    
   TxtValidate = True
   
End Function

Private Function FormSave() As Boolean
   
   Dim dblTotChange As Double
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   With grdDataList
      For ii = 1 To .Rows - 1
         strSql = "update LetterCacheVar set lcv06='" & .TextMatrix(ii, 2) & "'" & _
            " where lcv01='" & m_LC01 & "' and lcv02='" & m_LC02 & "' and lcv03='" & .TextMatrix(ii, 4) & "'"
         cnnConnection.Execute strSql, intI
         
         strSql = "update LetterCacheVar set lcv06='" & .TextMatrix(ii, 3) & "'" & _
            " where lcv01='" & m_LC01 & "' and lcv02='" & m_LC02 & "' and lcv03='" & .TextMatrix(ii, 4) & "點數'"
         cnnConnection.Execute strSql, intI
         
         dblTotChange = dblTotChange + (Val(.TextMatrix(ii, 2)) - Val(.TextMatrix(ii, 1)))
         '2009/11/9 add by sonia 大陸商標領證費更新進度檔費用及點數,開立收據
         If Mid(m_LC01, 1, 1) = "C" And .TextMatrix(ii, 4) = "領證費" And .TextMatrix(ii, 2) <> "" Then
            '2011/8/25 modify by sonia 大陸商標才能直接更新,CFT不行,故加入cp01='T'
            '2011/11/3 MODIFY BY SONIA 葉芳如說CFT也可以改 CFT-014016
            'modify by sonia 2016/10/28 TC-10827桂英說要加TC
            strSql = "update caseprogress set cp16='" & .TextMatrix(ii, 2) & "',cp18='" & .TextMatrix(ii, 3) & "'" & _
                     " where cp09='" & m_LC01 & "' and cp10='1701' and cp01 IN ('T','CFT','TC') "
            cnnConnection.Execute strSql, intI
         End If
         '2009/11/9 end
      Next
   End With
   
   'If dblTotChange <> 0 Then 'Remove by Morgan 2011/2/18 都要重新算,因為有可能又改回來
       'Modified by Lydia 2018/06/04 O12語法會出錯,index有影響 ; lcv01 => ''||lcv01
      strSql = "update LetterCacheVar set lcv06=lcv04+(" & dblTotChange & ")" & _
         " where ''||lcv01='" & m_LC01 & "' and lcv02='" & m_LC02 & "' and lcv03='費用合計'"
      cnnConnection.Execute strSql, intI
      
      strSql = "update LetterCacheVar set lcv06=lcv04+(" & dblTotChange & ")" & _
         " where ''||lcv01='" & m_LC01 & "' and lcv02='" & m_LC02 & "' and lcv03='費用總計'"
      cnnConnection.Execute strSql, intI
      
      strSql = "update LetterCacheVar set lcv06=lcv04+(" & dblTotChange / 1000 & ")" & _
         " where ''||lcv01='" & m_LC01 & "' and lcv02='" & m_LC02 & "' and lcv03='點數合計'"
      cnnConnection.Execute strSql, intI
   'End If
   
   'Remove by Morgan 2008/6/4 此處只修改報價不確認，統一在前畫面做
   'strSQL = "UPDATE LETTERCACHE SET LC07='" & strUserNum & "',LC08='" & strSrvDate(1) & "',LC09=TO_CHAR(SYSDATE,'HH24MISS')" & _
   '   " where lc01='" & m_LC01 & "' and lc02='" & m_LC02 & "'"
   'cnnConnection.Execute strSQL, intI
   
   'Added by Morgan 2015/7/2
   If Combo1.Enabled = True Then
      If Combo1.Tag <> Combo1 Then
         strSql = "update lettercache set lc16='" & ChgSQL(Combo1) & "' where lc01='" & m_LC01 & "' and lc02='" & m_LC02 & "'"
         cnnConnection.Execute strSql, intI
      End If
      'Added by Morgan 2015/12/2
      strExc(1) = ""
      If Check1.Value = vbChecked Then strExc(1) = Check1.Caption
      If Check2.Value = vbChecked Then strExc(1) = strExc(1) & IIf(strExc(1) = "", "", ";") & Check2.Caption
      strExc(1) = Trim(strExc(1))
      PUB_UpdateCP64Tag m_LC01, "開收據提醒", strExc(1)
      'end 2015/12/2
   End If
   'end 2015/7/2
   
   'Added by Lydia 2016/05/27 是否需主管簽核,更新為'Y';整批列印時,延後一天產生(LC06 或LC11 + 1 天)
   If Pub_Send_CFPdg Then
      'Modified by Lydia 2016/08/31 因為已確認請主管簽核,所以直接歸到已確認報價
      strSql = "update lettercache set lc17='Y',lc06=decode(lc06,null,null,to_number(to_char(to_date(lc06,'YYYYMMDD')+1,'YYYYMMDD')))," & _
               " lc11=decode(lc11,null,null,to_number(to_char(to_date(lc11,'YYYYMMDD')+1,'YYYYMMDD'))) " & _
               ",LC07='" & Trim(frm210124.txtSales.Text) & "',LC08='" & strSrvDate(1) & "',LC09=" & CNULL(Format(ServerTime, "000000")) & _
               " where lc01='" & m_LC01 & "' and lc02='" & m_LC02 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2016/05/27
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

'Added by Morgan 2015/12/2
Public Sub SetCheck()
   strExc(1) = PUB_ReadCP64Tag(m_LC01, "開收據提醒")
   If InStr(strExc(1), Check1.Caption) > 0 Then
      Check1.Value = vbChecked
   End If
   If InStr(strExc(1), Check2.Caption) > 0 Then
      Check2.Value = vbChecked
   End If
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Visible = True Then
      If UpdateVar = True Then
         txtInput.Visible = False
      Else
         txtInput.SetFocus
      End If
   End If
End Sub

'更新報價
Private Function UpdateVar() As Boolean
   Dim dblFeeChange As Double, dblNewPt As Double, bolCancel As Boolean
   'Dim strCP As Variant 'Remove by Lydia 2016/05/27
   
   If strCP(0) = "" Then GetData 'Added by Lydia 2016/05/27
   
   With grdDataList
      '費用變更時需更新點數
      If txtInput.Text <> txtInput.Tag Then
         dblFeeChange = Format(Val(txtInput)) - Val(.TextMatrix(m_iRow, 1))
         'strCP = Split(lblCaseNo, "-") 'Remove by Lydia 2016/05/27
         'Modify By Sindy 2013/4/23
         If Trim(lblCountry) = "歐洲聯盟" And strCP(0) = "CFT" And Trim(lblProperty) = "註冊證" Then
            If Val(txtInput) > 25000 Then
               MsgBox "CFT歐洲聯盟領證費以 $25,000 為上限！"
               bolCancel = True
               GoTo gotoExit
            End If
         ElseIf Abs(dblFeeChange) > 5000 Then
         '2013/4/23 End
           
            'Added by Lydia 2016/09/06 CFP年費,維持費,延展費的控制改成主管同意核准
            'Modified by Lydia 2016/10/05 EPC年費抓法不同
            'If Not (strCP(0) = "CFP" And InStr("年費,維持費,延展費", .TextMatrix(m_iRow, 0)) > 0) Then
            'Modified by Lydia 2016/11/11
            'If Not (strCP(0) = "CFP" And InStr("年費,維持費,延展費", IIf(m_Nation = "221", Right(.TextMatrix(m_iRow, 0), 2), .TextMatrix(m_iRow, 0))) > 0) Then
            'Modified by Lydia 2017/02/13 進入國家階段以發證日為準 (BY 甄妮)
            'If Not (strCP(0) = "CFP" And InStr("年費,維持費,延展費", IIf(m_Nation = "221" And m_Kind2 = "1", Right(.TextMatrix(m_iRow, 0), 2), .TextMatrix(m_iRow, 0))) > 0) Then
            If Not (strCP(0) = "CFP" And InStr("年費,維持費,延展費", IIf(m_Nation = "221" And m_PA21TM20 <> "", Right(.TextMatrix(m_iRow, 0), 2), .TextMatrix(m_iRow, 0))) > 0) Then
               MsgBox "費用調整不可超過 $5000！"
               bolCancel = True
               GoTo gotoExit
            End If
         End If
         If .TextMatrix(m_iRow, 3) <> "" Then
            dblFeeChange = Format(txtInput.Text) - Format(txtInput.Tag)
            dblNewPt = Val(.TextMatrix(m_iRow, 3)) + dblFeeChange / 1000
            If dblNewPt < 0 Then
               MsgBox "費用調整不可使點數小於 0！"
               bolCancel = True
               GoTo gotoExit
            Else
               .TextMatrix(m_iRow, 3) = dblNewPt
            End If
         End If
      End If
gotoExit:
      If bolCancel = False Then
         .TextMatrix(m_iRow, m_iCol) = Format(txtInput.Text)
         UpdateVar = True
      End If
   End With
End Function

''Added by Lydia 2016/05/27 取得年費 第X年
'Private Function GetCPMnexttimes(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String, ByVal strKEY05 As String, ByVal strKey06 As String, ByVal bstNP07 As String, ByVal p_stNP07 As String)
''strKey01~04    案號
''strKey05       國別
''strKey06       專利/商標種類
''bstNP07        傳入的案件性質(可多筆,以","區隔)
''p_stNP07       指定案件性質
'Dim arrNP07 As Variant
'Dim strCPM03s As String
'Dim m_iFixNo As Integer '修法次數
'Dim strYear As String '抓下次繳費年度
'Dim m_Nexttimes As String '抓下次繳費次數
'Dim strYF15 As String '年費年度說明
'Dim strKey(0 To 4) As String
'Dim strFeeType As String
'Dim strFeeYear As String
'Dim m_CaseFee(1 To 2) As String
'Dim aryCaseFee As Variant
'Dim rsRd As New ADODB.Recordset
'Dim inA As Integer
'Dim iX As Integer
'Dim Str01 As String
'Dim i As Integer
'
'   If (strKEY01 = "P" Or strKEY01 = "CFP") And (InStr(bstNP07, "605") > 0 Or InStr(bstNP07, "606") > 0 Or InStr(bstNP07, "607") > 0) Then
'       '設定本所案號
'       strKey(0) = ""
'       strKey(1) = strKEY01
'       strKey(2) = strKEY02
'       strKey(3) = strKEY03
'       strKey(4) = strKEY04
'       arrNP07 = Empty
'       arrNP07 = Split(bstNP07, ",")
'       strCPM03s = ""
'       For iX = LBound(arrNP07) To UBound(arrNP07)
'          '抓案件性質名稱
'          Str01 = "": strFeeYear = ""
'          If strKEY05 = "000" And (InStr(1, strKEY01, "P") > 0 Or InStr(1, strKEY01, "T") > 0) Then
'             strSql = "select DECODE(CPM03,'（無）',CPM04,CPM03) CPM03 from CasePropertyMap where cpm01='" & strKEY01 & "' and cpm02='" & arrNP07(i) & "'"
'          Else
'             strSql = "select DECODE(CPM04,'（無）',CPM03,CPM04) CPM03 from CasePropertyMap where cpm01='" & strKEY01 & "' and cpm02='" & arrNP07(i) & "'"
'          End If
'          inA = 1
'          Set rsRd = ClsLawReadRstMsg(inA, strSql)
'          If inA = 1 Then
'             Str01 = Str01 & IIf(Str01 <> "", ",", "") & rsRd.Fields("CPM03")
'          End If
'          'CFP-大馬新型延展費,俄羅斯設計延展費
'          If (strKEY01 = "CFP" And strKEY05 = "018" And strKey06 = "2" And arrNP07(i) = "607") Or _
'             (strKEY01 = "CFP" And strKEY05 = "233" And strKey06 = "3" And arrNP07(i) = "607") Then
'               strFeeYear = " " & PUB_GetExpYF607(strKEY05, strKEY01, strKEY02, strKEY03, strKEY04, strKEY05)
'
'          ElseIf (arrNP07(i) = "605" Or arrNP07(i) = "606" Or arrNP07(i) = "607") Then
'              '取得繳年費的資料(抓修法次數)
'              If GetMoneyDate(strKey06, strKEY05, strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
'                 '取得下次繳費次數/年度
'                 m_Nexttimes = PUB_Getnexttimes(strKEY01, strKEY02, strKEY03, strKEY04, strYear)
'                 If m_Nexttimes <> "" Then
'                    If p_stNP07 = "605" Then '605.年費
'                       aryCaseFee = Split(m_CaseFee(2), ",")
'                       strFeeType = PUB_GetNa20Na22Na24(strKEY05, strKey06)
'                       If strKEY05 = "017" And strFeeType = "605" And m_Nexttimes = "1" Then
'                          strYF15 = PUB_GetXDesc(strKEY01, strKEY02, strKEY03, strKEY04)
'                       Else
'                          strYF15 = PUB_GetYF15(strKEY05, strKey06, "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'                       End If
'                       strFeeYear = " " & strYF15
'
'                    ElseIf p_stNP07 = "606" Or p_stNP07 = "607" Then '606.維持費 607.延展費
'                       '年度說明
'                       strFeeType = PUB_GetNa20Na22Na24(strKEY05, strKey06)
'                       strYF15 = PUB_GetYF15(strKEY05, strKey06, "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'                       strFeeYear = " " & strYF15
'                    End If
'                 End If
'              End If
'          End If
'          strCPM03s = strCPM03s & Str01 & strFeeYear & ","
'       Next
'       If Right(strCPM03s, 1) = "," Then strCPM03s = Mid(strCPM03s, 1, Len(strCPM03s) - 1)
'   End If
'   GetCPMnexttimes = strCPM03s
'End Function
'Added by Lydia 2016/05/27
Private Sub GetData()
    'Modified by Lydia 2016/11/11 + decode(pa01,null,tm16,pa16) kind2
    'Modified by Lydia 2017/02/13 +發證日+ decode(pa01,null,tm20,pa21) pa21tm20
    strSql = "select cp01,cp02,cp03,cp04,nvl(pa09,tm10) nation,nvl(pa08,tm08) kind,decode(pa01,null,tm16,pa16) kind2,decode(pa01,null,tm20,pa21) pa21tm20 from caseprogress,patent,trademark where cp09='" & m_LC01 & "' " & _
             "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
       strCP(0) = RsTemp(0)
       strCP(1) = RsTemp(1)
       strCP(2) = RsTemp(2)
       strCP(3) = RsTemp(3)
       m_Nation = "" & RsTemp(4)
       m_Kind = "" & RsTemp(5)
       m_Kind2 = "" & RsTemp(6) 'Added by Lydia 2016/11/11
       m_PA21TM20 = "" & RsTemp(7) 'Added by Lydia 2017/02/13
    End If
End Sub

Private Sub Check3_Click()
Dim strText As String
   
   If Check3.Value = 1 Then
      strText = txtCRL101.Text
      If strText = "" Then
         MsgBox "請輸入營業地址！", vbExclamation
         Check3.Value = 0
         txtCRL101.SetFocus
         Exit Sub
      Else
         txtCRL100.Text = ""
      End If
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = 1 Then
      txtCRL99 = "境外"
   Else
      txtCRL99 = ""
   End If
End Sub
