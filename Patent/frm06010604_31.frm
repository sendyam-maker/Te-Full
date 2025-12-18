VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010604_31 
   BorderStyle     =   1  '單線固定
   Caption         =   "已收文領證通知"
   ClientHeight    =   3420
   ClientLeft      =   645
   ClientTop       =   1110
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印聯絡單(&P)"
      Height          =   400
      Index           =   4
      Left            =   4512
      TabIndex        =   25
      Top             =   70
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   6756
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "E-mail(&S)"
      Height          =   400
      Index           =   3
      Left            =   5832
      TabIndex        =   1
      Top             =   70
      Width           =   900
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   6
      Left            =   1140
      TabIndex        =   28
      Top             =   2640
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCasePropertyName 
      Height          =   285
      Left            =   1740
      TabIndex        =   27
      Top             =   2280
      Width           =   2265
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCountryName 
      Height          =   285
      Left            =   5700
      TabIndex        =   26
      Top             =   1200
      Width           =   2265
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1140
      TabIndex        =   0
      Top             =   1560
      Width           =   6852
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "13652;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   285
      Left            =   5940
      TabIndex        =   24
      Top             =   1920
      Width           =   2025
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   9
      Left            =   5100
      TabIndex        =   23
      Top             =   3000
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   8
      Left            =   1140
      TabIndex        =   22
      Top             =   3000
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   7
      Left            =   5100
      TabIndex        =   21
      Top             =   2640
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   5
      Left            =   5100
      TabIndex        =   20
      Top             =   2280
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   4
      Left            =   1140
      TabIndex        =   19
      Top             =   2280
      Width           =   555
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "979;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   3
      Left            =   5100
      TabIndex        =   18
      Top             =   1920
      Width           =   795
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1402;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   17
      Top             =   1920
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   1
      Left            =   5100
      TabIndex        =   16
      Top             =   1200
      Width           =   525
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "926;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseField 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   15
      Top             =   1200
      Width           =   2772
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   4140
      TabIndex        =   13
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "規費："
      Height          =   180
      Left            =   4140
      TabIndex        =   11
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "此案件已收文領證，是否向智權人員發出E-mail !!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   720
      Width           =   4275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4140
      TabIndex        =   6
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "費用："
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Left            =   4140
      TabIndex        =   4
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4140
      TabIndex        =   3
      Top             =   1200
      Width           =   900
   End
End
Attribute VB_Name = "frm06010604_31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo By Sindy 2021/11/22 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim cp(1 To T_CP) As String
Dim cp() As String
Dim intLeaveKind As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組

   Select Case Index
      Case 3
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.txtEmail(1) = "此案通知被異議，但日前巳收文領證，請通知客戶。"
         'frm880005.txtEmail(2) = "智權人員姓名：" + lblSalesName + vbCrLf + vbCrLf + _
               "案件名稱：" + cboCaseName.List(0) + vbCrLf + vbCrLf + _
               "收文日：" + lblCaseField(2).Caption + vbTab + vbTab + vbTab + "案件性質：" + lblCasePropertyName + vbCrLf + vbCrLf + _
               "費用：" + lblCaseField(6) + vbTab + vbTab + vbTab + "規費：" + lblCaseField(7) + vbTab + vbTab + vbTab + "點數：" + lblCaseField(5) + vbCrLf + vbCrLf + _
               "本所期限：" + lblCaseField(8) + vbTab + vbTab + vbTab + "法定期限：" + lblCaseField(9) + vbCrLf + vbCrLf + _
               "此案件已可發文但尚未收款，請儘快收款以便發文。"
         'frm880005.Show vbModal
         m_StrTo = lblCaseField(3)
         m_StrSub = "此案通知被異議，但日前已收文領證，請通知客戶。"
         m_StrCont = "智權人員姓名：" + lblSalesName + vbCrLf + vbCrLf + _
               "案件名稱：" + cboCaseName.List(0) + vbCrLf + vbCrLf + _
               "收文日：" + lblCaseField(2).Caption + vbTab + vbTab + vbTab + "案件性質：" + lblCasePropertyName + vbCrLf + vbCrLf + _
               "費用：" + lblCaseField(6) + vbTab + vbTab + vbTab + "規費：" + lblCaseField(7) + vbTab + vbTab + vbTab + "點數：" + lblCaseField(5) + vbCrLf + vbCrLf + _
               "本所期限：" + lblCaseField(8) + vbTab + vbTab + vbTab + "法定期限：" + lblCaseField(9) + vbCrLf + vbCrLf + _
               "此案件已可發文但尚未收款，請儘快收款以便發文。"
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
      Case 4
         PrintEmail
   End Select
   Unload Me
End Sub

Private Sub PrintEmail()
'edit by nickc 2007/02/06 不用 dll 了
 'Dim objPrintDllPublic As Object, intCaseKind As Integer, varSaveCursor
 Dim intCaseKind As Integer, varSaveCursor
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetSystemKind(cp(1), intCaseKind) Then
   If ClsPDGetSystemKind(cp(1), intCaseKind) Then
      'edit by nickc 2007/02/06 不用 dll 了
      'Set objPrintDllPublic = CreateObject("prjPrintDllPublic.clsPrintPublic")
      'objPrintDllPublic.PrintEmail intCaseKind, intPWhere, cp(9), strUserName
      'Set objPrintDllPublic = Nothing
      ClsPPPrintEmail intCaseKind, intPWhere, cp(9), strUserName
   End If
End Sub

Private Sub ReadAllData()
 On Error GoTo ErrHnd
   intLeaveKind = Left(frm06010604_31.Tag, 1)
    'Modify By Cheng 2002/12/03
'   cp(9) = Right(frm06010604_31.Tag, (frm06010604_31.Tag) - 1)
   cp(9) = Right(frm06010604_31.Tag, Len(frm06010604_31.Tag) - 1)
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intPWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intPWhere) Then
      If cp(1) = 馬德里案 Then
         lblCaseField(0) = cp(1) + " - " + Left(cp(2), 5) + _
            IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
            IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
            IIf(cp(4) = "00", "", " - " + cp(4))
      Else
         lblCaseField(0) = cp(1) + " - " + cp(2) + _
           IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
           IIf(cp(4) = "00", "", " - " + cp(4))
      End If
      lblCaseField(3) = cp(13)
      lblCaseField(4) = cp(10)
      lblCaseField(5) = cp(18)
      lblCaseField(6) = cp(16)
      lblCaseField(7) = cp(17)
      If intPWhere <> 國外_CF Then
         lblCaseField(2) = ChangeTStringToTDateString(cp(5))
         lblCaseField(8) = ChangeTStringToTDateString(cp(6))
         lblCaseField(9) = ChangeTStringToTDateString(cp(7))
      Else
         lblCaseField(2) = ChangeWStringToWDateString(cp(5))
         lblCaseField(8) = ChangeWStringToWDateString(cp(6))
         lblCaseField(9) = ChangeWStringToWDateString(cp(7))
      End If
      Select Case intLeaveKind
         'Remove by Morgan 2011/10/5 不再使用
         'Case 4
         '   lblCaseField(1) = frm04010504_3.MPa9
         '   SetComboToCombo cboCaseName, frm04010504_3.Combo1
            
         Case 6
            lblCaseField(1) = frm06010604_3.MPa9
            SetComboToCombo cboCaseName, frm06010604_3.Combo1
      End Select
   Else
      MsgBox "讀取CaseProgress檔時發生錯誤!!", vbCritical
   End If
   Exit Sub
ErrHnd:
   MsgBox Err.Description
   Resume
End Sub

Private Sub Form_Activate()
   ReadAllData
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010604_31 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, bolIsChina As Boolean

Select Case Index
   Case 1
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
      If ClsPDGetNation(lblCaseField(Index), strTemp) Then
         lblCountryName.Caption = strTemp
      End If
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
         lblSalesName.Caption = strTemp
      End If
   Case 4
      If lblCaseField(1) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
      If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
         lblCasePropertyName = strTemp
      End If
End Select
End Sub
