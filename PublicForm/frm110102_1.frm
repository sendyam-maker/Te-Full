VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110102_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "取消收文"
   ClientHeight    =   5670
   ClientLeft      =   1875
   ClientTop       =   1530
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8328
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7500
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Height          =   400
      Index           =   2
      Left            =   4272
      TabIndex        =   10
      Top             =   384
      Width           =   800
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1008
      MaxLength       =   3
      TabIndex        =   0
      Top             =   384
      Width           =   732
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1776
      TabIndex        =   13
      Top             =   384
      Width           =   2415
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3816
      Left            =   48
      TabIndex        =   9
      Top             =   1824
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6720
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   285
      Left            =   1764
      TabIndex        =   14
      Top             =   432
      Visible         =   0   'False
      Width           =   2412
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   5
         TabIndex        =   4
         Top             =   0
         Width           =   852
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   1
         Left            =   960
         MaxLength       =   1
         TabIndex        =   5
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   2
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   6
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   204
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Top             =   0
         Width           =   492
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   276
      Left            =   996
      TabIndex        =   8
      Top             =   1272
      Width           =   8175
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14420;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   180
      Left            =   1035
      TabIndex        =   22
      Top             =   1605
      Width           =   3135
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNumber1 
      Height          =   220
      Left            =   1800
      TabIndex        =   21
      Top             =   756
      Width           =   2412
   End
   Begin VB.Label lblNumber2 
      Height          =   220
      Left            =   996
      TabIndex        =   20
      Top             =   1026
      Width           =   3972
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   36
      TabIndex        =   19
      Top             =   1323
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "申請人："
      Height          =   180
      Left            =   48
      TabIndex        =   18
      Top             =   1620
      Width           =   972
   End
   Begin VB.Label lblNumber 
      Caption         =   "審定號數/證書號數："
      Height          =   180
      Left            =   48
      TabIndex        =   17
      Top             =   729
      Width           =   1692
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號："
      Height          =   180
      Left            =   36
      TabIndex        =   16
      Top             =   1026
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   48
      TabIndex        =   15
      Top             =   432
      Width           =   972
   End
End
Attribute VB_Name = "frm110102_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/7 改成Form2.0(cboCaseName,lblAgent,grdDataList)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'edit by nickc 2007/02/06 不用 dll 了
'Dim obj011 As Object
'Add By Sindy 2022/6/15


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean

Select Case Index
             Case 0
                        frm110102_2.Show
                        Me.Hide
             Case 1
                        Unload Me
             Case 2
                        If txtSystem = 馬德里案 Then
                           bolRt = CheckKeyIn1(3)
                        Else
                           bolRt = CheckKeyIn2(2)
                        End If
                        If bolRt Then GetCancelReceivedDayCaseData
End Select
End Sub

Private Sub GetCancelReceivedDayCaseData(Optional bolBlank As Boolean = False)
'TF為馬德里案，另外判斷
If bolBlank = True Then
   GetCancelReceivedDayData "", "", "", ""
Else
   If txtSystem = 馬德里案 Then
      GetCancelReceivedDayData txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))
   Else
      GetCancelReceivedDayData txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
   End If
End If
End Sub

Private Sub GetCancelReceivedDayData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor
Dim i As Integer

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/06 不用 dll 了
'Set grdDataList.Recordset = obj011.ReadCancelReceivedDayRst(strCode1, strCode2, strCode3, strCode4)
Set grdDataList.Recordset = Cls011ReadCancelReceivedDayRst(strCode1, strCode2, strCode3, strCode4)
         'ADD BY SONIA 2016/8/31
         For i = 1 To grdDataList.Rows - 1
            grdDataList.TextMatrix(i, 2) = grdDataList.TextMatrix(i, 2) & PUB_GetRelateCasePropertyName(grdDataList.TextMatrix(i, 0), "1")
         Next i
         'END 2016/8/31
SetDataListWidth
SetDataListVision grdDataList
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 7
   cmdOK(0).Enabled = True
   cmdOK(0).Default = True
   If grdDataList.Rows = 2 Then
      grdDataList.row = 1
      grdDataList.col = 0
      If Len(grdDataList.Text) <> 0 Then
         cmdOK_Click (0)
      End If
   End If
Else
   cmdOK(0).Enabled = False
   cmdOK(2).Default = True
   '911024 nick 若沒有資料，秀訊息
   If strCode2 <> "" Then
        ShowNoData
        txtSystem.SetFocus
   End If
End If
Screen.MousePointer = varSaveCursor
End Sub

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(1200, 900, 1500, 900, 900, 900, 900, 900, 1100)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

Private Sub Form_Activate()
   'GetCancelReceivedDayCaseData
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetDataListWidth
'edit by nickc 2007/02/06 不用 dll 了
'If obj011 Is Nothing Then
'   Set obj011 = CreateObject("prjTaieDll011.cls011")
'   Set obj011.Connection = cnnConnection
'End If
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm110102_1 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click 0
End Sub

Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetCancelReceivedDayCaseData True

End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
If Not (FMP2open = True And (txtSystem.Text = "P" Or txtSystem.Text = "PS")) Then
    If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
       ShowMsg MsgText(1056)
       Cancel = True
       txtSystem_GotFocus
    End If
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetCancelReceivedDayCaseData True
End Sub

Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetCancelReceivedDayCaseData True
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub

Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
  'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2)
      End If
   End If
    If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
    'If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
       SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
       lblNumber1 = strNumber1
       lblNumber2 = strNumber2
       lblAgent = strCustomer
       CheckKeyIn1 = True
    End If
    
Else
   CheckKeyIn1 = True
End If
End Function

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
CloseIme
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
  'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2)
      End If
   End If
    If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
   ' If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
       SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
       lblNumber1 = strNumber1
       lblNumber2 = strNumber2
       lblAgent = strCustomer
       CheckKeyIn2 = True
    End If
Else
   CheckKeyIn2 = True
End If
End Function

Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then grdDataList_DblClick
End Sub

Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      'edit by nickc 2005/07/21 修 bug
      'ShowBar grdDataList, intLastRow, 7
      ShowBar grdDataList, intLastRow, 8
      blnOKtoShow = True
   End If
End If
End Sub

Public Sub Cleartxt()
txtCode(0) = ""
txtCode(1) = ""
txtSystem = ""
txtCode(2) = ""
cboCaseName.Clear
lblNumber1.Caption = ""
lblNumber2.Caption = ""
lblAgent.Caption = ""
txtSystem.SetFocus
End Sub
