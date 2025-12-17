VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc14p0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國內收據產生特殊請款單"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdWord 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word(&W)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7350
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   150
      Width           =   1455
   End
   Begin VB.TextBox txtNo 
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
      Left            =   5340
      MaxLength       =   9
      TabIndex        =   3
      Top             =   180
      Width           =   1890
   End
   Begin VB.TextBox txtCaseNo 
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
      Left            =   1770
      MaxLength       =   12
      TabIndex        =   0
      Top             =   180
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Frmacc14p0.frx":0000
      Left            =   1770
      List            =   "Frmacc14p0.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1230
      Width           =   1920
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收據編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3930
      TabIndex        =   2
      Top             =   180
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   3180
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   2970
      Width           =   1635
   End
   Begin VB.TextBox txtRate 
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
      Left            =   5340
      TabIndex        =   5
      Top             =   1230
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3435
      Left            =   60
      TabIndex        =   7
      Top             =   1650
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|公司別|收據日期|智權人員|收據編號|案件性質|結清|合併|NTD規費|NTD服務費|外幣規費"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "注意：在產生特殊請款單時，不要使用Word！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   5130
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSForms.Label lblCustCaseNo 
      Height          =   285
      Left            =   5340
      TabIndex        =   16
      Top             =   630
      Width           =   3030
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "5345;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單幣別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   53
      Left            =   465
      TabIndex        =   15
      Top             =   1290
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶案件案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3825
      TabIndex        =   14
      Top             =   630
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   675
      TabIndex        =   12
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "匯　　率："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4245
      TabIndex        =   11
      Top             =   1290
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   675
      TabIndex        =   10
      Top             =   630
      Width           =   1050
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Left            =   1770
      TabIndex        =   9
      Top             =   960
      Width           =   7170
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "12647;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label lblNation 
      Height          =   285
      Left            =   1770
      TabIndex        =   8
      Top             =   630
      Width           =   1860
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "3281;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "Frmacc14p0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已修改
'Create By Sindy 2014/4/22
Option Explicit

Dim m_dftColor As Long '預設顏色
Dim m_dftColor2 As Long '預設顏色2
Dim m_dftColor3 As Long '點選顏色
Dim iRow As Integer, iCol As Integer
Dim m_PA01 As String, m_PA02 As String, m_PA03 As String, m_PA04 As String
Dim m_PADate As String, m_PANo As String, m_PAKind As String
Dim m_FileName As String


Private Function TxtValidate() As Boolean
Dim i As Integer, strPriNo As String
Dim bolIsSel As Boolean
   
   TxtValidate = False
   
   If Trim(Combo1.Text) = "" Then
      MsgBox "請點選請款單幣別！"
      Combo1.SetFocus
      Exit Function
   End If
   
   If Val(txtRate) = 0 Then
      MsgBox "請輸入匯率！"
      txtRate.SetFocus
      Exit Function
   End If
   
   '檢查第一筆及最後一筆是否為相同收據號碼,若是,則自動打勾
   If MSHFlexGrid2.TextMatrix(1, 4) <> "" Then
      If MSHFlexGrid2.TextMatrix(1, 4) = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 4) Then
         MSHFlexGrid2.TextMatrix(1, 0) = "V"
         SetColor 1, m_dftColor3
      End If
   End If
   '檢查是否有收據號碼已勾但其他同收據號碼未勾選的
   strPriNo = ""
   For i = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "V" Then
         strPriNo = MSHFlexGrid2.TextMatrix(i, 4)
      Else
         If strPriNo <> "" And strPriNo = MSHFlexGrid2.TextMatrix(i, 4) Then
            MSHFlexGrid2.TextMatrix(i, 0) = "V"
            SetColor i, m_dftColor3
         End If
      End If
   Next i
   strPriNo = ""
   For i = MSHFlexGrid2.Rows - 1 To 1 Step -1
      If MSHFlexGrid2.TextMatrix(i, 0) = "V" Then
         strPriNo = MSHFlexGrid2.TextMatrix(i, 4)
      Else
         If strPriNo <> "" And strPriNo = MSHFlexGrid2.TextMatrix(i, 4) Then
            MSHFlexGrid2.TextMatrix(i, 0) = "V"
            SetColor i, m_dftColor3
         End If
      End If
   Next i
   '至少要勾選一張收據號碼
   bolIsSel = False
   For i = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "V" Then
         bolIsSel = True
         Exit For
      End If
   Next i
   If bolIsSel = False Then
      MsgBox "至少要勾選一張收據號碼！"
      Exit Function
   End If
   
   '檢查NTD規費欄為0時外幣規費欄才可為0
   For i = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "V" Then
         If Val(MSHFlexGrid2.TextMatrix(i, 10)) = 0 And Val(MSHFlexGrid2.TextMatrix(i, 8)) > 0 Then
            MsgBox "外幣規費不可為0！"
            MSHFlexGrid2.col = 10
            MSHFlexGrid2.row = i
            iRow = i: iCol = 10
            SetBox
            Exit Function
         End If
      End If
   Next i
   
   TxtValidate = True
End Function

Private Sub JCallWordPrint()
Dim strNo As String, strFileName As String, strA0k02 As String
Dim i As Integer, jj As Integer
Dim strName As String
Dim strText As String
Dim dblSerFee As Double, dblFee As Double
Dim dblTotSerFee As Double, dblTotFee As Double
Dim iRows As Integer
   
On Error GoTo ErrHand
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   
   strNo = ""
   For jj = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(jj, 0) = "V" Then
         If strNo <> MSHFlexGrid2.TextMatrix(jj, 4) Then
            '先存檔
            If strNo <> "" Then
               '頁尾
               With g_WordAp
                  .Selection.Tables(1).Rows(8).Select
                  .Selection.Cells(2).Select
                  .Selection.TypeText Text:="NTD" & IIf(dblTotSerFee = 0, 0, Format(dblTotSerFee, DDollar))
                  .Selection.Tables(1).Rows(8).Select
                  .Selection.Cells(3).Select
                  .Selection.TypeText Text:="NTD" & IIf(dblTotFee = 0, 0, Format(dblTotFee, DDollar))
                  .Selection.Tables(1).Rows(9).Select
                  .Selection.Cells(2).Select
                  .Selection.TypeText Text:="NTD" & Format((dblTotSerFee + dblTotFee), DDollar)
                  .Selection.Find.ClearFormatting
                  .Selection.Find.Text = "|#新台幣總額#|"
                  .Selection.Find.Replacement.Text = ""
                  .Selection.Find.Forward = True
                  .Selection.Find.Wrap = wdFindContinue
                  .Selection.Find.Format = False
                  .Selection.Find.MatchCase = False
                  .Selection.Find.MatchWholeWord = False
                  .Selection.Find.MatchWildcards = False
                  .Selection.Find.MatchSoundsLike = False
                  .Selection.Find.MatchAllWordForms = False
                  .Selection.Find.MatchByte = True
                  .Selection.Find.Execute
                  .Selection.Delete
                  .Selection.Font.ColorIndex = wdBlack
                  .Selection.TypeText ChangeNumber(CStr((dblTotSerFee + dblTotFee)))
                  .ActiveDocument.Save
                  .ActiveDocument.Close
               End With
            End If
            dblTotSerFee = 0: dblTotFee = 0: iRows = 1
            '開新檔
            '檔名:客戶編號-收據編號-收據抬頭.doc
            'strFileName = MSHFlexGrid2.TextMatrix(jj, 11) & "-" & MSHFlexGrid2.TextMatrix(jj, 4) & "-" & MSHFlexGrid2.TextMatrix(jj, 12) & ".doc"
            '檔名:收據抬頭2個字-本所案號.doc
            strFileName = Left(MSHFlexGrid2.TextMatrix(jj, 12), 2) & "-" & m_PA01 & m_PA02 & m_PA03 & m_PA04 & ".doc"
            If Dir(PUB_Getdesktop & "\" & strFileName) <> "" Then
               Kill PUB_Getdesktop & "\" & strFileName
            End If
            g_WordAp.Documents.Open App.path & "\" & m_FileName
            g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & strFileName
            g_WordAp.ActiveDocument.Close
            g_WordAp.Documents.Open PUB_Getdesktop & "\" & strFileName
            '頁首
            With g_WordAp
               .Selection.WholeStory
               .Selection.Copy
               For i = 1 To 10
                  strName = ""
                  strText = ""
                  If i = 1 Then
                     strName = "請款日期"
                     strA0k02 = DBDATE(MSHFlexGrid2.TextMatrix(jj, 2))
                     strText = Left(strA0k02, 4) & "年" & Mid(strA0k02, 5, 2) & "月" & Right(strA0k02, 2) & "日"
                  ElseIf i = 2 Then
                     strName = "收據單號"
                     strText = MSHFlexGrid2.TextMatrix(jj, 4)
                  ElseIf i = 3 Then
                     strName = "客戶案號"
                     strText = lblCustCaseNo
                  ElseIf i = 4 Then
                     strName = "申請日期"
                     strText = m_PADate
                  ElseIf i = 5 Then
                     strName = "本所案號"
                     strText = m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 = "000", "", "-" & m_PA03 & "-" & m_PA04)
                  ElseIf i = 6 Then
                     strName = "申請國家"
                     strText = lblNation
                  ElseIf i = 7 Then
                     strName = "申請案號"
                     strText = m_PANo
                  ElseIf i = 8 Then
                     strName = "申請種類"
                     strText = m_PAKind
                  ElseIf i = 9 Then
                     strName = "案件名稱"
                     strText = LblCaseName
                  Else
                     strName = "案件性質"
                     strText = ""
                  End If
                  If Trim(strName) <> "" Then
                     .Selection.Find.ClearFormatting
                     .Selection.Find.Text = "|#" & strName & "#|"
                     .Selection.Find.Replacement.Text = ""
                     .Selection.Find.Forward = True
                     .Selection.Find.Wrap = wdFindContinue
                     .Selection.Find.Format = False
                     .Selection.Find.MatchCase = False
                     .Selection.Find.MatchWholeWord = False
                     .Selection.Find.MatchWildcards = False
                     .Selection.Find.MatchSoundsLike = False
                     .Selection.Find.MatchAllWordForms = False
                     .Selection.Find.MatchByte = True
                     .Selection.Find.Execute
                     .Selection.Delete
                     .Selection.Font.ColorIndex = wdBlack
                     .Selection.TypeText strText
                  End If
               Next i
            End With
         End If
         strNo = MSHFlexGrid2.TextMatrix(jj, 4)
         iRows = iRows + 1
         If MSHFlexGrid2.TextMatrix(jj, 7) = "Y" Then
            dblSerFee = Val(MSHFlexGrid2.TextMatrix(jj, 8)) + Val(MSHFlexGrid2.TextMatrix(jj, 9))
         Else
            dblSerFee = Val(MSHFlexGrid2.TextMatrix(jj, 9))
         End If
         dblFee = Val(MSHFlexGrid2.TextMatrix(jj, 10))
         dblTotSerFee = dblTotSerFee + dblSerFee
         If dblFee > 0 Then
            dblTotFee = dblTotFee + (dblFee * txtRate)
         End If
         '明細
         With g_WordAp
            .Selection.Tables(1).Rows(iRows).Select
            .Selection.Cells(1).Select
            .Selection.TypeText Text:=MSHFlexGrid2.TextMatrix(jj, 5)
            .Selection.Tables(1).Rows(iRows).Select
            .Selection.Cells(2).Select
            .Selection.TypeText Text:=IIf(dblSerFee = 0, 0, Format(dblSerFee, DDollar))
            .Selection.Tables(1).Rows(iRows).Select
            .Selection.Cells(3).Select
            .Selection.TypeText Text:=IIf(dblFee = 0, 0, Format(dblFee, DDollar))
            If iRows = 2 Then
               .Selection.Tables(1).Rows(iRows).Select
               .Selection.Cells(4).Select
               .Selection.TypeText Text:="1" & Left(Trim(Combo1.Text), 3) & ":" & txtRate & "NTD"
            End If
         End With
      End If
   Next jj
   '最後一筆存檔
   If strNo <> "" Then
      '頁尾
      With g_WordAp
         .Selection.Tables(1).Rows(8).Select
         .Selection.Cells(2).Select
         .Selection.TypeText Text:="NTD" & IIf(dblTotSerFee = 0, 0, Format(dblTotSerFee, DDollar))
         .Selection.Tables(1).Rows(8).Select
         .Selection.Cells(3).Select
         .Selection.TypeText Text:="NTD" & IIf(dblTotFee = 0, 0, Format(dblTotFee, DDollar))
         .Selection.Tables(1).Rows(9).Select
         .Selection.Cells(2).Select
         .Selection.TypeText Text:="NTD" & Format((dblTotSerFee + dblTotFee), DDollar)
         .Selection.Find.ClearFormatting
         .Selection.Find.Text = "|#新台幣總額#|"
         .Selection.Find.Replacement.Text = ""
         .Selection.Find.Forward = True
         .Selection.Find.Wrap = wdFindContinue
         .Selection.Find.Format = False
         .Selection.Find.MatchCase = False
         .Selection.Find.MatchWholeWord = False
         .Selection.Find.MatchWildcards = False
         .Selection.Find.MatchSoundsLike = False
         .Selection.Find.MatchAllWordForms = False
         .Selection.Find.MatchByte = True
         .Selection.Find.Execute
         .Selection.Delete
         .Selection.Font.ColorIndex = wdBlack
         .Selection.TypeText ChangeNumber(CStr((dblTotSerFee + dblTotFee)))
         .ActiveDocument.Save
      End With
   End If
   g_WordAp.ActiveDocument.Close
   g_WordAp.Quit
   Set g_WordAp = Nothing
   MsgBox "檔案已產生至桌面！"
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
      If Not g_WordAp Is Nothing Then
         g_WordAp.Quit
         Set g_WordAp = Nothing
      End If
   End If
End Sub

Private Sub cmdWord_Click()
   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      Call JCallWordPrint
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
On Error GoTo Checking
   
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            doQuery
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, , MsgBox(5)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   
   If Option1(0).Value = True Then
      If Trim(txtCaseNo) = "" Then
         MsgBox "本所案號不可空白！", , MsgText(5)
         txtCaseNo.SetFocus
         Exit Function
      End If
      If Len(Trim(txtCaseNo)) < 10 Then
         MsgBox "請輸入完整的本所案號！", , MsgText(5)
         txtCaseNo.SetFocus
         Exit Function
      End If
   Else
      If Trim(txtNo) = "" Then
         MsgBox "收據編號不可空白！", , MsgText(5)
         txtNo.SetFocus
         Exit Function
      End If
   End If
   
   FormCheck = True
End Function

Private Sub doQuery()
   Dim strCon As String
   Dim bolChk As Boolean, strTmp As String
   
   lblNation.Caption = ""
   lblCustCaseNo.Caption = ""
   LblCaseName.Caption = ""
   MSHFlexGrid2.Clear
   GridHead
   cmdWord.Enabled = False
   
   Screen.MousePointer = vbHourglass
   
   '收據編號
   If Option1(1).Value = True Then
      txtCaseNo.Text = ""
      strExc(0) = "select a0j02 from acc0k0,acc0j0 where a0k01='" & txtNo & "' and a0k01=a0j13(+) order by a0j02 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCaseNo.Text = "" & RsTemp.Fields(0)
      Else
         MsgBox "無此收據編號！", , MsgText(5)
         Exit Sub
      End If
      strCon = " and a0k01='" & txtNo & "' and a0k01=a0j13(+)"
   '本所案號
   Else
      txtNo.Text = ""
      strCon = " and a0j02='" & txtCaseNo & "' and a0j13=a0k01(+)"
   End If
   
   strKey1 = Left(Trim(txtCaseNo.Text), Len(Trim(txtCaseNo.Text)) - 9)
   StrKey2 = Mid(Trim(txtCaseNo.Text), Len(Trim(txtCaseNo.Text)) - 9 + 1, 6)
   strKey3 = Mid(Trim(txtCaseNo.Text), Len(Trim(txtCaseNo.Text)) - 3 + 1, 1)
   strKey4 = Right(Trim(txtCaseNo.Text), 2)
   strExc(0) = "select pa09,na03,pa05||pa06||pa07,pa48,pa01,pa02,pa03,pa04,pa10,pa11,pa08,'P' as SysID from patent,nation where pa01='" & strKey1 & "' and pa02='" & StrKey2 & "' and pa03='" & strKey3 & "' and pa04='" & strKey4 & "' and pa09=na01(+)" & _
               " Union select tm10,na03,tm05||tm06||tm07,tm35,tm01,tm02,tm03,tm04,tm11,tm12,tm08,'T' as SysID from trademark,nation where tm01='" & strKey1 & "' and tm02='" & StrKey2 & "' and tm03='" & strKey3 & "' and tm04='" & strKey4 & "' and tm10=na01(+)" & _
               " Union select sp09,na03,sp05||sp06||sp07,sp29,sp01,sp02,sp03,sp04,sp10,sp11,'','S' as SysID from servicepractice,nation where sp01='" & strKey1 & "' and sp02='" & StrKey2 & "' and sp03='" & strKey3 & "' and sp04='" & strKey4 & "' and sp09=na01(+)" & _
               " Union select lc15,na03,lc05||lc06||lc07,lc17,lc01,lc02,lc03,lc04,0,'','','L' as SysID from lawcase,nation where lc01='" & strKey1 & "' and lc02='" & StrKey2 & "' and lc03='" & strKey3 & "' and lc04='" & strKey4 & "' and lc15=na01(+)" & _
               " Union select '','',hc06,'',hc01,hc02,hc03,hc04,0,'','','H' as SysID from hirecase where hc01='" & strKey1 & "' and hc02='" & StrKey2 & "' and hc03='" & strKey3 & "' and hc04='" & strKey4 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblNation.Caption = "" & RsTemp.Fields(1)
      LblCaseName.Caption = "" & RsTemp.Fields(2)
      lblCustCaseNo.Caption = "" & RsTemp.Fields(3)
      m_PA01 = RsTemp.Fields("pa01")
      m_PA02 = RsTemp.Fields("pa02")
      m_PA03 = RsTemp.Fields("pa03")
      m_PA04 = RsTemp.Fields("pa04")
      m_PADate = "" & RsTemp.Fields("pa10")
      If m_PADate <> "" Then m_PADate = ChangeWStringToWDateString(m_PADate)
      m_PANo = "" & RsTemp.Fields("pa11")
      m_PAKind = "" & RsTemp.Fields("pa08")
      If RsTemp.Fields("SysID") = "P" Then
         If RsTemp.Fields("pa09") = 台灣國家代號 Or m_PA01 = "CFP" Then
            bolChk = False
         Else
            bolChk = True
         End If
         Call ClsPDGetPatentTrademarkKind(專利, m_PAKind, strTmp, bolChk, RsTemp.Fields("pa09"))
         m_PAKind = strTmp
      ElseIf RsTemp.Fields("SysID") = "T" Then
         m_PAKind = GetTradeMarkName(m_PAKind, 0)
      End If
   Else
      MsgBox "無此本所案號！", , MsgText(5)
      Exit Sub
   End If
   
   strExc(0) = "select ' ' V,a0k05,sqldatet(a0k02) a0k02,st02,a0k01,cpmNm,a0k37,a0j07,a0j10-nvl(a1u09,0) a0j10_2,a0j09-nvl(a1u07,0) a0j09_2,' ' inpFee,a0k03,a0k04" & _
               " from acc1u0,(" & _
               " select a0k05,a0k02,st02,a0k01,a0j01,a0j07,a0j25,a0k10,decode(a0j04,'000',cpm03,cpm04) cpmNm,a0k37,a0j10,a0j09,a0k03,a0k04" & _
               " From acc0j0, acc0k0, staff, caseprogress, casepropertymap" & _
               " Where (a0k09 Is Null Or a0k09 = 0)" & _
               " and (a0k37<>'N' or a0k37 is null)" & _
               " and a0k20=st01(+)" & _
               " and a0j01=cp09(+)" & _
               " and cp01=cpm01(+) and cp10=cpm02(+)" & strCon & ") d" & _
               " where d.a0k01=a1u02(+)" & _
               " and d.a0j01=a1u03(+)" & _
               " and d.a0k10=a1u01(+)" & _
               " order by a0k01 asc,a0j25 asc"
   intI = 1
   Set MSHFlexGrid2.Recordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "無任何請款單資料！", , MsgText(5)
   Else
      cmdWord.Enabled = True
   End If
   GridHead
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub GridHead()
   With MSHFlexGrid2
      .Visible = False
      .Cols = 13
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "V"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 500: .Text = "公司別"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignCenterCenter
      
      .col = 2: .ColWidth(2) = 800: .Text = "收據日期"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 3: .ColWidth(3) = 800: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      
      .col = 4: .ColWidth(4) = 900: .Text = "收據編號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      
      .col = 5: .ColWidth(5) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignLeftCenter
      
      .col = 6: .ColWidth(6) = 400: .Text = "結清"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      
      .col = 7: .ColWidth(7) = 400: .Text = "合併"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(7) = flexAlignCenterCenter
      
      .col = 8: .ColWidth(8) = 1000: .Text = "NTD規費"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 9: .ColWidth(9) = 1000: .Text = "NTD服務費"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 10: .ColWidth(10) = 1000: .Text = "外幣規費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(10) = flexAlignRightCenter
      
      .col = 11: .ColWidth(11) = 1000: .Text = "客戶編號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(11) = flexAlignRightCenter
      
      .col = 12: .ColWidth(12) = 1000: .Text = "收據抬頭"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(12) = flexAlignRightCenter
            
      For intI = 11 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub UpdateCol()
   Dim ii As Integer
   
   If txtInput <> txtInput.Tag Then
      With MSHFlexGrid2
      If iCol = 10 Then
         .TextMatrix(iRow, iCol) = Val(txtInput.Text)
      Else
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = .TextMatrix(iRow, 0) Then
               .TextMatrix(ii, iCol) = txtInput.Text
            End If
         Next
      End If
      End With
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   '底色
   m_dftColor = &HFFFFFF
   '底色2
   m_dftColor2 = RGB(&HFF, &HFA, &HCD)
   '底色3
   m_dftColor3 = &HFFC0C0
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原W:9120 H:5700
   Me.Width = 9225
   Me.Height = 5970
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   lblNation.Caption = ""
   lblCustCaseNo.Caption = ""
   LblCaseName.Caption = ""
   txtInput.Visible = False
   MSHFlexGrid2.Clear
   GridHead
   cmdWord.Enabled = False
   
   Combo1.AddItem ""
   strExc(0) = "SELECT A1Y01||'-'||A1Y02 FROM ACC1Y0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         Combo1.AddItem RsTemp.Fields(0)
         RsTemp.MoveNext
      Loop
   End If
   
   Call Option1_Click(0)
   
   m_FileName = "$$國內收據特殊請款單(昆盈).doc"
   If Dir(App.path & "\" & m_FileName) <> "" Then
      Kill App.path & "\" & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000003-0-00")
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand

   If Not g_WordAp Is Nothing Then
      g_WordAp.Visible = True
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
   
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc14p0 = Nothing
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo CloseWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

Private Sub SetColor(pRow As Integer, pColor As Long)
   With MSHFlexGrid2
   .row = pRow
   For intI = 0 To .Cols - 1
      .col = intI
      .CellBackColor = pColor
   Next
   End With
End Sub

Private Sub MSHFlexGrid2_Click()
   Dim iCurCol As Integer, iCurRow As Integer
   
   With MSHFlexGrid2
   If .MouseRow > 0 And .MouseRow < .Rows And .MouseCol < 18 Then
      iCurRow = .MouseRow
      iCurCol = .MouseCol
      .Visible = False
      
      .row = iCurRow
      .col = 0
      If Trim(.TextMatrix(.row, .col)) = "" Then
         .TextMatrix(.row, .col) = "V"
         SetColor iCurRow, m_dftColor3
      Else
         .TextMatrix(.row, .col) = ""
         SetColor iCurRow, m_dftColor
      End If
      
      .col = iCurCol
      iRow = .row: iCol = .col
      If .col = 10 Then SetBox
           
      .Visible = True
   End If
   End With
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtInput.Visible = False
End Sub

Private Sub MSHFlexGrid2_Scroll()
   If txtInput.Visible = True Then
      SetBox False
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         txtCaseNo.Enabled = True
         txtNo.Enabled = False
      Case 1
         txtCaseNo.Enabled = False
         txtNo.Enabled = True
   End Select
End Sub

Private Sub txtCaseNo_GotFocus()
   InverseTextBox txtCaseNo
End Sub

Private Sub txtCaseNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtInput_GotFocus()
   InverseTextBox txtInput
   CloseIme
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Locked = False Then UpdateCol
   txtInput.Visible = False
End Sub

Private Sub SetBox(Optional pbolSetValue As Boolean = True)
   Dim lngLeft As Long, lngTop As Long
   Dim ii As Integer
   
   With MSHFlexGrid2
   If .LeftCol > .col Or .TopRow > .row Then
      txtInput.Visible = False
   Else
      txtInput.FontName = .CellFontName
      txtInput.FontSize = .CellFontSize
      If .CellAlignment < 3 Then
         txtInput.Alignment = 0 '靠左
      ElseIf .CellAlignment < 6 Then
         txtInput.Alignment = 2 '置中
      ElseIf .CellAlignment < 9 Then
         txtInput.Alignment = 1 '靠右
      Else
         txtInput.Alignment = 0 '靠左
      End If
      If pbolSetValue = True Then
         txtInput.Text = .TextMatrix(.row, .col)
      End If
      txtInput.Tag = txtInput.Text
      txtInput.Width = .ColWidth(.col) + 10
      txtInput.Height = .RowHeight(.row) - 5
      lngLeft = .Left + 20
      lngTop = .Top + .RowHeight(0) + 20
      For ii = .LeftCol To .col - 1
         lngLeft = lngLeft + .ColWidth(ii)
      Next
      For ii = .TopRow To .row - 1
         lngTop = lngTop + .RowHeight(ii)
      Next
      txtInput.Left = lngLeft: txtInput.Top = lngTop
      If txtInput.Left + txtInput.Width < .Left + .Width Then
         txtInput.Visible = True
         txtInput.SetFocus
         TextInverse txtInput
         iRow = .row: iCol = .col
      Else
         txtInput.Visible = False
      End If
   End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   
   If KeyAscii = vbKeyReturn Then
      UpdateCol
      GoNext
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
   '外幣規費欄位
   ElseIf iCol = 10 Then
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 46 Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
   End If
End Sub

Private Sub GoNext()
   With MSHFlexGrid2
      .col = 10
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      SetBox
   End With
End Sub

Private Sub txtNo_GotFocus()
   InverseTextBox txtNo
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtRate_GotFocus()
   InverseTextBox txtRate
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
End Sub
