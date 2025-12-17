VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc14o1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "INVOICE編號作廢作業"
   ClientHeight    =   4080
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6228
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6228
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc14o1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   114
      Width           =   350
   End
   Begin VB.CommandButton cmdProc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "編號作廢(&E)"
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
      Left            =   1752
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   3600
      Width           =   2196
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Top             =   107
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   2796
      Left            =   72
      TabIndex        =   2
      Top             =   552
      Width           =   5988
      _ExtentX        =   10562
      _ExtentY        =   4932
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "V|收據編號|收據日期|收據金額|客戶編號|智權人員|收據抬頭"
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "INVOICE編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   168
      TabIndex        =   4
      Top             =   168
      Width           =   1236
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc14o1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2023/11/13
Option Explicit

Private Sub SetGrd(Optional ByVal bolReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1           2           3           4           5           6           7           8
   arrGridHeadText = Array("V", "收據編號", "收據日期", "收據金額", "智權人員", "客戶編號", "收據抬頭", "外幣金額")
   arrGridHeadWidth = Array(200, 1000, 800, 1000, 1000, 1000, 3000, 0)
   If bolReset = True Then
      GRD1.Clear
      GRD1.Rows = 2
   End If
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub CmdProc_Click()
Dim strTmp As String

   If Trim(txtNo.Text) <> "" And GRD1.Rows >= 2 Then
      strExc(1) = Trim(txtNo)
      strSql = "UPDATE ACC0K0 SET A0K40=NULL WHERE A0K40='" & strExc(1) & "' "
      cnnConnection.Execute strSql
      Call ClearAll(True)
      MsgBox "請自行刪除作廢INVOICE之WORD檔案。" & "INVOICE編號：" & strExc(1), vbInformation
   End If
End Sub

Private Sub Command3_Click()
   Call doQuery
End Sub

Private Sub doQuery()
Dim Rs As ADODB.Recordset
   
   Call SetGrd(True)
   
   Screen.MousePointer = vbHourglass
   
   strExc(0) = "select '' as V, a0k01,substr(sqldatet(a0k02),1,9) a0k02,sum(a0j09-nvl(a1u07,0)+a0j10-nvl(a1u09,0)) amt,st02,a0k03,a0k04,cu196,a0k40" & _
               " from acc1u0,(" & _
               "select a0k01,a0k02,st02,a0j01,a0j07,a0j25,a0k10,a0j09,a0j10,a0k03,a0k04,cu196,a0k40" & _
               " From acc0j0,acc0k0,staff,customer" & _
               " Where (a0k09 Is Null Or a0k09 = 0)" & _
               " and (a0k37<>'N' or a0k37 is null)" & _
               " and a0k20=st01(+)" & _
               " and a0k40='" & txtNo & "' and a0k01=a0j13 and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)) d" & _
               " where d.a0k01=a1u02(+)" & _
               " and d.a0j01=a1u03(+)" & _
               " and d.a0k10=a1u01(+)" & _
               " group by a0k01,a0k02,st02,a0k03,a0k04,cu196,a0k40"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "無任何INVOICE資料！", , MsgText(5)
      CmdProc.Enabled = False
   Else
      GRD1.FixedCols = 0
      Set GRD1.Recordset = Rs
      Call SetGrd
      CmdProc.Enabled = True
   End If
   
   Set Rs = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub Grd1_Click()
Dim intRow As Integer

   With GRD1
      If .MouseRow > 0 Then
         intRow = .MouseRow
         .row = intRow
         GridClick GRD1, intRow, 0, 0
      End If
   End With
End Sub

Private Sub Form_Load()
   
   '表單初始化
   PUB_InitForm Me, 6300, 4500, strBackPicPath4
      
   Call ClearAll
   Call SetGrd(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)

   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear

   Set Frmacc14o1 = Nothing
End Sub

Private Sub ClearAll(Optional ByVal bolReset As Boolean = True)
   If bolReset = True Then
      txtNo.Text = ""
   End If
   Call SetGrd(True)
   CmdProc.Enabled = False
End Sub

Private Sub txtNo_GotFocus()
   TextInverse txtNo
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
