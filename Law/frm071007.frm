VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071007 
   BorderStyle     =   1  '單線固定
   Caption         =   "收件人資料輸入"
   ClientHeight    =   5370
   ClientLeft      =   1320
   ClientTop       =   555
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   4752
      TabIndex        =   4
      Top             =   48
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "收件人資料"
      Height          =   4716
      Left            =   264
      TabIndex        =   5
      Top             =   552
      Width           =   5535
      Begin VB.CommandButton cmdInput 
         Caption         =   "新增(&A)"
         Default         =   -1  'True
         Height          =   400
         Left            =   3024
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   360
         Width           =   800
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "刪除(&D)"
         Height          =   400
         Left            =   3852
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   360
         Width           =   800
      End
      Begin MSForms.TextBox txtReceiver 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   418
         Width           =   1695
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   30
         Size            =   "2990;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox List1 
         Height          =   3300
         Left            =   420
         TabIndex        =   3
         Top             =   960
         Width           =   4695
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "8281;5821"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         Caption         =   "收件人："
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   433
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm071007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ;List1、txtReceiver
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim intRec As Integer, intNowRecd As Integer

Private Sub cmdBack_Click()
 Dim strName As String, i As Integer
   For i = 0 To List1.ListCount - 1
     strName = strName + List1.List(i) + ","
   Next
   strPublicTemp = strName
   frm071006.Show
   Unload Me
   frm071006.Command2.SetFocus
End Sub

Private Sub cmdCancel_Click()
   txtReceiver = ""
   List1.RemoveItem intNowRecd
   cmdInput.Enabled = False
   cmdCancel.Enabled = False
End Sub

Private Sub cmdInput_Click()
   cmdInput.Enabled = False
   cmdCancel.Enabled = False
   List1.AddItem txtReceiver
   txtReceiver = ""
   txtReceiver.SetFocus
End Sub

Private Sub Form_Activate()
   txtReceiver.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdInput.Enabled = False
   cmdCancel.Enabled = False
   'GetData
   ReadTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071007 = Nothing
End Sub

Private Sub List1_Click()
   cmdCancel.Enabled = True
   'Modified by Lydia 2021/12/02 改成Form 2.0; 沒有此屬性
   'intNowRecd = List1.Columns
   intNowRecd = List1.ListIndex
   txtReceiver = List1.Text
End Sub

Private Sub txtReceiver_Change()
   If txtReceiver <> "" Or IsNull(txtReceiver) Then
      cmdInput.Enabled = True
   End If
End Sub

Private Sub txtReceiver_GotFocus()
   TextInverse txtReceiver
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtReceiver.IMEMode = 1
   OpenIme
End Sub

'Modified by Lydia 2021/09/14 改成Form2.0 ;
'Private Sub txtReceiver_KeyPress(KeyAscii As Integer)
Private Sub txtReceiver_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtReceiver_Validate(Cancel As Boolean)
   If txtReceiver <> "" Or IsNull(txtReceiver) Then
      txtReceiver = UCase(txtReceiver)
      cmdInput.SetFocus
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub GetData()
   strExc(0) = "select cp50 from caseprogress where cp43='" + frm071006.lbeNum + "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Do While Not RsTemp.EOF
         If IsNull(RsTemp.Fields(0).Value) = False Then
            List1.AddItem RsTemp.Fields(0).Value
         End If
         RsTemp.MoveNext
      Loop
   End If
End Sub
Private Sub ReadTemp()
  Dim strCP50 As Variant
  Dim i As Integer
  strCP50 = Split(strPublicTemp, ",")
  For i = 0 To UBound(strCP50) - 1
      List1.AddItem strCP50(i)
  Next i
End Sub
