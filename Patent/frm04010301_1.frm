VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010301_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "P案各式申請書"
   ClientHeight    =   5760
   ClientLeft      =   -4290
   ClientTop       =   1170
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9345
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   810
      MaxLength       =   1
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   600
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "P"
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   0
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   6120
      MaxLength       =   7
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   7530
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8364
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   5
      Top             =   1200
      Width           =   4935
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "8705;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "<--text6不可刪除"
      Height          =   180
      Left            =   1260
      TabIndex        =   18
      Top             =   5460
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   5160
      TabIndex        =   15
      Top             =   660
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   1020
      TabIndex        =   13
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   5160
      TabIndex        =   12
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   6120
      TabIndex        =   11
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1200
      Width           =   765
   End
End
Attribute VB_Name = "frm04010301_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,MSHFlexGrid1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Add by Morgan 2011/9/22 Copy frm040103_1 來改
'工程師用的各式申請書
Option Explicit
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 0 '確定
         Text6 = "" 'Added by Morgan 2023/5/19
         Me.Tag = ""
         pa(10) = ""
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               Me.Tag = MSHFlexGrid1.TextMatrix(i, 2)
               pa(10) = MSHFlexGrid1.TextMatrix(i, 7)
               Exit For
            End If
         Next
         
         If pa(10) = 補文件 Then
            Set frm04010304_1.oParentForm = Me
            frm04010304_1.iFrom = 1
            frm04010304_1.Caption = "P案" & frm04010304_1.Caption
            frm04010304_1.Text6 = "Y" '預設修改
            frm04010304_1.Show
         'Added by Morgan 2023/5/18 +244補中文說明書,232補優先權證明
         ElseIf pa(10) = "244" Or pa(10) = "232" Then
            Text6 = "3"
            Set frm04010304_1.oParentForm = Me
            frm04010304_1.Show
         'Add By Sindy 2022/8/25
         Else
            MsgBox "無申請書!", vbCritical
            Exit Sub
            '2022/8/25 END
         'Modify By Sindy 2020/3/27 Mark
'         Else
'            Set frm04010310_1.oParentForm = Me
'            frm04010310_1.iFrom = 1
'            frm04010310_1.Caption = "P案" & frm04010310_1.Caption
'            frm04010310_1.m_strCP10 = pa(10)
'            frm04010310_1.Show
         End If
         
         cmdOK(1).SetFocus
         Me.Hide
      Case 1 '尋找
         Label4 = ""
         Label6 = ""
         Combo1.Clear
         MSHFlexGrid1.Clear
         GridHead
         
         If Text3 = "" Then Text3 = "0"
         If Text4 = "" Then Text4 = "00"
         pa(1) = Text1
         pa(2) = Text2
         pa(3) = Text3
         pa(4) = Text4
         
         If pa(1) = "P" Then
            If ClsPDReadPatentDatabase(pa(), intWhere) Then
               If pa(9) = 台灣國家代號 Then
                  AddCboName Combo1, pa(5), pa(6), pa(7)
                  Text5.Text = pa(10)
                  Label4.Caption = pa(11)
                  Label6.Caption = pa(22)
                  
               Else
                  MsgBox "本案件非台灣案！"
                  Text2.SetFocus
                  Exit Sub
               End If
            Else
               Text2.SetFocus
               Exit Sub
            End If
         ElseIf pa(1) = "PS" Then
            If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
               AddCboName Combo1, pa(5), pa(6), pa(7)
               Text5.Text = pa(10)
               Label4.Caption = pa(11)
            Else
               Text2.SetFocus
               Exit Sub
            End If
         End If
                  
         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,staff.st02 as st1,staff1.st02 as st2," & _
            "cp64,cp10 from caseprogress, casepropertymap,staff,staff staff1 where " & _
            ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and " & _
            "( cp09<'C' ) and cp01=cpm01(+) and " & _
            "cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
         GridHead
         If Me.MSHFlexGrid1.Rows = 2 Then
            MSHFlexGrid1_Click
         End If
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
 Dim i As Integer, j As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
         For j = 0 To .Cols - 1
            .col = j
            .CellBackColor = .BackColor
         Next
      Next
   End With
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()


   MoveFormToCenter Me
   intWhere = 國內
'   Combo1.ListIndex = 0
   Label4 = ""
   Label6 = ""
   InitGrid 8, MSHFlexGrid1
   GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010301_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
      GridClick MSHFlexGrid1, intLastRow, 0
      cmdOK(0).SetFocus
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "P" And Text1 <> "PS" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "案件備註"
      .col = 7: .ColWidth(7) = 0
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

Public Sub ClearForm()
   Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Label4 = Empty
   Label6 = Empty
   Combo1.Clear
   InitGrid 8, MSHFlexGrid1
   GridHead
   Text1.SetFocus
End Sub

