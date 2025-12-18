VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040106_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外指示信"
   ClientHeight    =   3330
   ClientLeft      =   285
   ClientTop       =   2775
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印聯絡單(&P)"
      Height          =   405
      Index           =   4
      Left            =   3240
      TabIndex        =   29
      Top             =   70
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   1
      Left            =   6708
      TabIndex        =   3
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "E-mail(&S)"
      Height          =   405
      Index           =   3
      Left            =   4560
      TabIndex        =   1
      Top             =   70
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繼續作業(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5484
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   1500
      Width           =   6855
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12091;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   1860
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCountryName 
      Height          =   255
      Left            =   5400
      TabIndex        =   27
      Top             =   1140
      Width           =   2535
      VariousPropertyBits=   27
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCasePropertyName 
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   2220
      Width           =   2535
      VariousPropertyBits=   27
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   25
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   24
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   23
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   22
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   21
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   20
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   19
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   18
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   17
      Top             =   1140
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   1140
      Width           =   2775
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   4080
      TabIndex        =   14
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "規費："
      Height          =   180
      Left            =   4080
      TabIndex        =   12
      Top             =   2580
      Width           =   540
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "此程序未收款，是否向智權人員發出E-mail !!"
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
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4080
      TabIndex        =   7
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "費用："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   2580
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Left            =   4080
      TabIndex        =   5
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4080
      TabIndex        =   4
      Top             =   1140
      Width           =   900
   End
End
Attribute VB_Name = "frm040106_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Morgan 2021/12/13 改成Form2.0 (cboCaseName,lblSalesName..)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim cp(1 To T_CP) As String
Dim cp() As String

Dim intLeaveKind As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組

   Select Case Index
      Case 0
         intLeaveKind = 1
      Case 1
         intLeaveKind = 0
      Case 3
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.txtEmail(1) = "請儘快收款以便發文"
         'frm880005.txtEmail(2) = "智權人員姓名：" + lblSalesName + vbCrLf + vbCrLf + _
               "本所案號：" + lblCaseField(0).Caption + vbCrLf + vbCrLf + _
               "案件名稱　" + cboCaseName.List(0) + vbCrLf + vbCrLf + _
               "收文日：" + lblCaseField(2).Caption + vbTab + vbTab + vbTab + "案件性質：" + lblCasePropertyName + vbCrLf + vbCrLf + _
               "費用：" + lblCaseField(6) + vbTab + vbTab + vbTab + "規費：" + lblCaseField(7) + vbTab + vbTab + vbTab + "點數：" + lblCaseField(5) + vbCrLf + vbCrLf + _
               "本所期限：" + lblCaseField(8) + vbTab + vbTab + vbTab + "法定期限：" + lblCaseField(9) + vbCrLf + vbCrLf + _
               "以上案件已可發文但尚未收款，請儘快收款以便發文。"
         'frm880005.Show vbModal
         'If frm880005.bolLeave Then
         '   intLeaveKind = 1
         'Else
         '   Exit Sub
         'End If
         m_StrTo = lblCaseField(3)
         m_StrSub = "請儘快收款以便發文"
         m_StrCont = "智權人員姓名：" + lblSalesName + vbCrLf + vbCrLf + _
               "本所案號：" + lblCaseField(0).Caption + vbCrLf + vbCrLf + _
               "案件名稱　" + cboCaseName.List(0) + vbCrLf + vbCrLf + _
               "收文日：" + lblCaseField(2).Caption + vbTab + vbTab + vbTab + "案件性質：" + lblCasePropertyName + vbCrLf + vbCrLf + _
               "費用：" + lblCaseField(6) + vbTab + vbTab + vbTab + "規費：" + lblCaseField(7) + vbTab + vbTab + vbTab + "點數：" + lblCaseField(5) + vbCrLf + vbCrLf + _
               "本所期限：" + lblCaseField(8) + vbTab + vbTab + vbTab + "法定期限：" + lblCaseField(9) + vbCrLf + vbCrLf + _
               "以上案件已可發文但尚未收款，請儘快收款以便發文。"
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
      Case 4
         PrintEmail
         intLeaveKind = 1
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
   cp(9) = frm040106_1.Tag
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
      lblCaseField(1) = frm040106_1.MainPa9
      SetComboToCombo cboCaseName, frm040106_1.Combo1
   Else
      MsgBox "讀取CaseProgress檔時發生錯誤!!", vbCritical
   End If
   Exit Sub
ErrHnd:
   MsgBox Err.Description
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
   Select Case intLeaveKind
      Case 0 '停止
         frm040106_1.bolLeave = True
      Case 1 '繼續
         frm040106_1.bolLeave = False
   End Select
   Set frm040106_2 = Nothing
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

