VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050208_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF代理人報價附件資料"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8175
   Begin VB.TextBox textCQ11 
      Height          =   270
      Left            =   1440
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   6315
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   180
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7050
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox textCQ04 
      Height          =   270
      Left            =   90
      TabIndex        =   10
      Top             =   2580
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   7410
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2340
      Width           =   735
   End
   Begin VB.ListBox lstAtt 
      Height          =   1860
      ItemData        =   "frm050208_1.frx":0000
      Left            =   1185
      List            =   "frm050208_1.frx":0007
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   6210
   End
   Begin VB.TextBox textCQ01 
      Height          =   300
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   0
      Top             =   570
      Width           =   1485
   End
   Begin VB.TextBox textCQ02 
      Height          =   300
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   1
      Top             =   900
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   150
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox textCQ03 
      Height          =   990
      Left            =   1185
      TabIndex        =   2
      Top             =   1230
      Width           =   6210
      VariousPropertyBits=   -1466941413
      MaxLength       =   4000
      ScrollBars      =   2
      Size            =   "10954;1746"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabelCQ01 
      Height          =   240
      Left            =   2700
      TabIndex        =   14
      Top             =   600
      Width           =   5325
      Caption         =   " lblFM2"
      Size            =   "9393;423"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   8070
      Caption         =   "CREATE : lblFM2"
      Size            =   "14235;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "附　件："
      Height          =   180
      Index           =   3
      Left            =   405
      TabIndex        =   9
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(民國年月日)"
      Height          =   180
      Index           =   0
      Left            =   2700
      TabIndex        =   8
      Top             =   930
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "內　容："
      Height          =   180
      Index           =   2
      Left            =   405
      TabIndex        =   7
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   615
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日　期："
      Height          =   180
      Index           =   17
      Left            =   405
      TabIndex        =   5
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frm050208_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; Label23、LabelCQ01、textCQ03
'Create By Sindy 2012/12/19
Option Explicit

Public m_CurrKEY1 As String
Public m_CurrKEY2 As String

Private Const cTableName As String = "CFQUOTATION" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱

Private Sub cmdExit_Click()
   Unload Me
   frm050208.Show
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetCtrlReadOnly True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   PUB_KillTempFile "$$*.*" 'Added by Lydia 2017/08/09 清除暫存檔
   
   Set frm050208_1 = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Public Function UpdateCtrlData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   UpdateCtrlData = False
   strSql = "SELECT * FROM CFQuotation " & _
            "WHERE CQ01='" & m_CurrKEY1 & "' and CQ02=" & Val(DBDATE(m_CurrKEY2))
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      UpdateCtrlData = True
      If IsNull(rsTmp.Fields("CQ01")) = False Then: textCQ01 = rsTmp.Fields("CQ01")
      LabelCQ01 = GetPrjName1(textCQ01)
      If IsNull(rsTmp.Fields("CQ02")) = False Then: textCQ02 = TAIWANDATE(rsTmp.Fields("CQ02"))
      If IsNull(rsTmp.Fields("CQ03")) = False Then: textCQ03 = rsTmp.Fields("CQ03")
      If IsNull(rsTmp.Fields("CQ04")) = False Then: textCQ04 = rsTmp.Fields("CQ04")
      SetList lstAtt, textCQ04
      'Added by Lydia 2017/08/09
      If IsNull(rsTmp.Fields("CQ11")) = False Then: textCQ11 = rsTmp.Fields("CQ11")
      
      cmdOpenAtt.Enabled = True
      Call UpdateCUID(rsTmp)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("CQ05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CQ05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CQ06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ07")) = False Then
         strTemp = rsSrcTmp.Fields("CQ07")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CQ08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CQ09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ10")) = False Then
         strTemp = rsSrcTmp.Fields("CQ10")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textCQ01.Locked = bEnable
   textCQ02.Locked = bEnable
   If bEnable Then textCQ01.BackColor = &H8000000F Else textCQ01.BackColor = &H80000005
   If bEnable Then textCQ02.BackColor = &H8000000F Else textCQ02.BackColor = &H80000005
   textCQ03.Locked = bEnable
   textCQ04.Locked = bEnable
End Sub

Private Sub ClearField()
   textCQ01 = Empty
   LabelCQ01 = Empty
   textCQ02 = Empty
   textCQ03 = Empty
   textCQ04 = Empty
   Label23.Caption = Empty
   
   lstAtt.Clear
   cmdOpenAtt.Enabled = False
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And textCQ11.Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(textCQ11.Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, "(") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      Else
      'end 2017/08/09
          PUB_OpenFtpFile textCQ01, lstAtt.Text, Winsock1, "4"
      End If 'end 2017/08/09
   End If
End Sub
