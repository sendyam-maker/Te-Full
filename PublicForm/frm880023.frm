VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880023 
   BorderStyle     =   1  '單線固定
   Caption         =   "輸入"
   ClientHeight    =   8184
   ClientLeft      =   -12
   ClientTop       =   276
   ClientWidth     =   8352
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8184
   ScaleWidth      =   8352
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1764
      Left            =   48
      TabIndex        =   5
      Top             =   4944
      Width           =   5676
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1368
         Left            =   168
         TabIndex        =   6
         Top             =   288
         Width           =   5328
         _ExtentX        =   9398
         _ExtentY        =   2413
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   48
      TabIndex        =   2
      Top             =   3456
      Width           =   7860
      Begin VB.OptionButton Option1 
         Caption         =   "列暫收　(請於備註說明暫收原因.僅能是沖客戶之後案件之費用)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Index           =   1
         Left            =   168
         TabIndex        =   4
         Top             =   768
         Width           =   7596
      End
      Begin VB.OptionButton Option1 
         Caption         =   "退客戶　(請二個月內提供客戶帳號資料交財務退客戶)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Index           =   0
         Left            =   168
         TabIndex        =   3
         Top             =   312
         Width           =   7596
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   144
      TabIndex        =   1
      Top             =   2736
      Width           =   4788
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   2496
      Left            =   108
      TabIndex        =   0
      Top             =   108
      Width           =   4836
      VariousPropertyBits=   -1395638245
      ScrollBars      =   2
      Size            =   "8530;4403"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm880023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2022/1/22
Option Explicit

Dim oTextBox As Object
Public p_iChoice As Integer, p_sReturn As String 'Added by Morgan 2023/12/6

Dim intLastRow As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub Command1_Click()
   Dim ii As Integer
   
   If TextBox1.Visible Then
      oTextBox.Text = TextBox1.Text
      oTextBox.SelStart = TextBox1.SelStart
      
   'Added by Morgan 2023/12/6
   ElseIf p_iChoice = 1 Then
      If Option1(0).Value = True Then
         p_sReturn = "2"
      ElseIf Option1(1).Value = True Then
         p_sReturn = "1"
      Else
         MsgBox "請選擇溢收款處理方式！", vbExclamation
         Exit Sub
      End If
   'end 2023/12/6
   
   'Added by Morgan 2025/11/7
   '表格點選一筆
   ElseIf p_iChoice = 2 Then
      With MSHFlexGrid1
        .Visible = False
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = "v" Then Exit For
        Next ii
        .Visible = True
        If ii = .Rows Then
            If MsgBox("尚未點選資料，是否確定要離開！", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
        Else
            p_sReturn = .TextMatrix(ii, 2)
        End If
      End With
   End If
   Unload Me
End Sub

Public Sub SetTextBox(pTextBox As Object)
   Set oTextBox = pTextBox
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   Dim iVSpace As Integer, iHSpace As Integer
   
   If Not bolActivated Then
      'Modified by Morgan 2025/11/7
      'iVSpace = Me.Width - TextBox1.Width
      'iHSpace = Me.Height - TextBox1.Height
      iVSpace = 300
      iHSpace = Command1.Height + 700
      'end 2025/11/7
      
      'Added by Morgan 2023/12/6
      '智權繳款作業-溢收款處理方式
      'Added by Morgan 2023/12/6
      TextBox1.Visible = False
      Frame1.Visible = False
      Frame2.Visible = False 'Added by Morgan 2025/11/7
      If p_iChoice = 1 Then
         Me.Caption = "溢收款處理方式"
         Frame1.Visible = True
         Frame1.Top = TextBox1.Top
         Frame1.Left = TextBox1.Left
         Command1.Top = Frame1.Top + Frame1.Height + 50
         'Modified by Morgan 2025/11/7
         'Command1.Left = Frame1.Left + Frame1.Width / 2 - Command1.Width / 2
         Command1.Left = Frame1.Left
         Command1.Width = Frame1.Width
         'end 2025/11/7
         Me.Width = Frame1.Width + iVSpace
         Me.Height = Frame1.Height + iHSpace
         
      'Added by Morgan 2025/11/7
      '表格資料點選
      ElseIf p_iChoice = 2 Then
         Frame2.Visible = True
         Me.Caption = "請點選一筆"
         Frame2.Top = TextBox1.Top
         Frame2.Left = TextBox1.Left
         Command1.Top = Frame2.Top + Frame2.Height + 50
         Command1.Left = Frame2.Left
         Command1.Width = Frame2.Width
         Me.Width = Frame2.Width + iVSpace
         Me.Height = Frame2.Height + iHSpace
      'end 2025/11/7
      
      Else
         TextBox1.Visible = True
      'end 2023/12/6
      
         TextBox1.Text = oTextBox.Text
         TextBox1.Width = oTextBox.Width
         If oTextBox.Height > 2500 Then
            TextBox1.Height = oTextBox.Height
         End If
         TextBox1.SelStart = oTextBox.SelStart
         TextBox1.MaxLength = oTextBox.MaxLength
         
         Command1.Top = TextBox1.Top + TextBox1.Height + 50
         'Modified by Morgan 2025/11/7
         'Command1.Left = TextBox1.Left + TextBox1.Width / 2 - Command1.Width / 2
         Command1.Left = TextBox1.Left
         Command1.Width = TextBox1.Width
         'end 2025/11/7
         Me.Width = TextBox1.Width + iVSpace
         Me.Height = TextBox1.Height + iHSpace
      End If
      bolActivated = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If p_iChoice = 0 Then
      Set frm880023 = Nothing
   End If
End Sub

'Add By Sindy 2022/10/13 + PUB_RefreshText
Private Sub TextBox1_Change()
   PUB_RefreshText TextBox1
End Sub

'Added by Morgan 2025/11/7
Public Sub SetGridHead(pRst As ADODB.Recordset, Optional pColNameList As String, Optional pColWidthList As String, Optional pColAlignList As String)
   Dim ii As Integer, jj As Integer, kk As Integer
   Dim strList As String, arrList() As String
   Set MSHFlexGrid1.Recordset = pRst
   FixGrid MSHFlexGrid1
   If pColNameList & pColWidthList & pColAlignList <> "" Then
      
      With MSHFlexGrid1
         .Visible = False
         For ii = 1 To 3
            If ii = 1 Then
               strList = pColNameList
            ElseIf ii = 2 Then
               strList = pColWidthList
            Else
               strList = pColAlignList
            End If
            If strList <> "" Then
               arrList = Split(strList, ",")
               .row = 0
               For jj = LBound(arrList) To UBound(arrList)
                  .col = jj
                  .CellAlignment = flexAlignCenterCenter
                  If ii = 1 Then
                     .Text = arrList(jj)
                  ElseIf ii = 2 Then
                     .ColWidth(.col) = Val(arrList(jj))
                  ElseIf ii = 3 Then
                     .ColAlignment(.col) = Val(arrList(jj))
                  End If
               Next
               If ii = 1 Then
                  For kk = jj To .Cols - 1
                     .col = kk: .ColWidth(kk) = 0
                  Next
               End If
            End If
         Next
        .Visible = True
      End With
   End If
End Sub

'Added by Morgan 2025/11/7
Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "v" Then
      If MSHFlexGrid1.Text = "點數" Then
         If m_blnColOrderAsc = True Then
            MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Added by Morgan 2025/11/7
Private Sub MSHFlexGrid1_SelChange()
   GridClick MSHFlexGrid1, intLastRow, 0
   Command1.SetFocus
End Sub

