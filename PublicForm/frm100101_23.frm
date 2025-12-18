VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_23 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "個人行事曆資料查詢"
   ClientHeight    =   5750
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6780
      TabIndex        =   11
      Top             =   75
      Width           =   1230
   End
   Begin VB.TextBox textSS 
      Height          =   264
      Index           =   1
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   4
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox textSS 
      Height          =   270
      Index           =   2
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   3
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox textSS 
      Height          =   264
      Index           =   3
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1230
      Width           =   975
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8070
      TabIndex        =   0
      Top             =   75
      Width           =   800
   End
   Begin MSForms.TextBox txtSS04 
      Height          =   3900
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   7116
      VariousPropertyBits=   -1472184293
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12552;6879"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   3525
      Size            =   "6218;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "員工姓名："
      Height          =   180
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "期限日期："
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "序　　號："
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   1590
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(lbl1(0),textSS(4)改為txtSS04)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

Public cmdState As Integer
Dim strTmp As String

Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As TextBox
Dim idx As Integer

Private Sub DataGrid1_Click()
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_23 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
   End Select
End Sub

Sub StrMenu()
Dim strKey  As String, strKey1 As String, StrKey2 As String
Dim varRef As Variant
Dim ii As Integer
   
   varRef = Split(Me.Tag, "-")
   For ii = LBound(varRef) To UBound(varRef)
      If ii = 0 Then strKey = varRef(ii)
      If ii = 1 Then strKey1 = varRef(ii)
      If ii = 2 Then StrKey2 = varRef(ii)
   Next ii
   
   strExc(0) = "select * from Staff_Schedule where ss01='" & strKey & "' and ss02=" & strKey1 & " and ss03=" & StrKey2 & " "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ShowRecord RsTemp, strKey, strKey1, StrKey2
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset, strKey As String, strKey1 As String, StrKey2 As String)
   Dim rsSS As ADODB.Recordset
   
   ClearField
   SetCtrlReadOnly True
   Set rsSS = p_Rst.Clone
   With rsSS
      If .RecordCount > 0 Then
         For Each oText In textSS
            idx = oText.Index
            oText.Text = "" & .Fields("SS" & Format(idx, "0#"))
            If idx = 2 Then
               If Trim(oText.Text) <> "" Then
                  oText.Text = ChangeTStringToTDateString(ChangeWStringToTString(Trim(oText.Text)))
               End If
            End If
         Next
         txtSS04.Text = "" & RsTemp.Fields("ss04")    'add by sonia 2022/1/22 textSS(4)改為txtSS04
         '員工姓名
         lbl1(0).Caption = GetPrjSalesNM(RsTemp.Fields("ss01"))
      End If
   End With
End Sub

Private Sub ClearField()
   Dim oLabel As LABEL
   For Each oText In textSS
      oText.Text = Empty
   Next
   txtSS04.Text = ""  'add by sonia 2022/1/22 textSS(4)改為txtSS04
   'modify by sonia 2022/1/22
   'For Each oLabel In lbl1
   '   oLabel.Caption = Empty
   'Next
   lbl1(0).Caption = ""
   'end 2022/1/22
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In textSS
      oText.Locked = bLocked
   Next
   txtSS04.Locked = bLocked   'add by sonia 2022/1/22 textSS(4)改為txtSS04
End Sub

'Add By Sindy 2010/11/26
Private Sub textSS_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1
         KeyAscii = UpperCase(KeyAscii)
      Case 2
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub
