VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180301_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人/繪圖人員外出記錄"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Tag             =   "加班資料"
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   2
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   24
      Top             =   600
      Width           =   945
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   3
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   23
      Top             =   960
      Width           =   945
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   4
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   22
      Top             =   1320
      Width           =   945
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   9
      Left            =   6885
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Top             =   960
      Width           =   435
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   8
      Left            =   6495
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   960
      Width           =   315
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   7
      Left            =   5475
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txtOG 
      Height          =   270
      Index           =   6
      Left            =   4845
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Top             =   960
      Width           =   525
   End
   Begin VB.TextBox txtOG 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   4845
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   705
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   360
      Left            =   6630
      TabIndex        =   0
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8040
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSMask.MaskEdBox mebOG 
      Height          =   270
      Index           =   19
      Left            =   4845
      TabIndex        =   7
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   476
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebOG 
      Height          =   270
      Index           =   20
      Left            =   6120
      TabIndex        =   8
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   476
      _Version        =   393216
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSForms.Label Label27 
      Height          =   195
      Left            =   5010
      TabIndex        =   29
      Top             =   5130
      Width           =   3735
      VariousPropertyBits=   27
      Caption         =   "Update ID:           Date         Time             "
      Size            =   "6588;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   5010
      TabIndex        =   28
      Top             =   4830
      Width           =   3735
      VariousPropertyBits=   27
      Caption         =   "Create ID:           Date         Time             "
      Size            =   "6588;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtOG2 
      Height          =   810
      Index           =   12
      Left            =   1500
      TabIndex        =   27
      Top             =   3390
      Width           =   7005
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "12356;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtOG2 
      Height          =   810
      Index           =   11
      Left            =   1500
      TabIndex        =   26
      Top             =   2520
      Width           =   7005
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "12356;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtOG2 
      Height          =   810
      Index           =   10
      Left            =   1500
      TabIndex        =   25
      Top             =   1650
      Width           =   7005
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "12356;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   255
      Index           =   1
      Left            =   2490
      TabIndex        =   21
      Top             =   1350
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   255
      Index           =   0
      Left            =   2490
      TabIndex        =   20
      Top             =   990
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblStarW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2490
      TabIndex        =   19
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(格式 : HH:mm) "
      Height          =   180
      Index           =   9
      Left            =   7260
      TabIndex        =   18
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   8
      Left            =   600
      TabIndex        =   17
      Top             =   3435
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外出時數："
      Height          =   180
      Index           =   5
      Left            =   3900
      TabIndex        =   16
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   1380
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4995
      X2              =   7125
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外出人員："
      Height          =   180
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   3
      Left            =   3900
      TabIndex        =   13
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外出日期："
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "序號："
      Height          =   180
      Index           =   1
      Left            =   3900
      TabIndex        =   11
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "地點："
      Height          =   180
      Index           =   6
      Left            =   600
      TabIndex        =   10
      Top             =   1725
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Index           =   7
      Left            =   600
      TabIndex        =   9
      Top             =   2565
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5865
      X2              =   6015
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frm180301_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Create by Sindy 2013/6/26
Option Explicit

Public m_OG01 As String
Dim m_PrevForm As Form '前一畫面 'Add By Sindy 2013/6/21


Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   
   m_OG01 = Right("000000" & m_OG01, 6)
   strSql = "SELECT OG01,OG02-19110000 AS OG02,OG03,OG04,OG06,OG07,OG08,OG09,OG10,OG11,OG12,OG19,OG20" & _
            ",A.ST02 AS D01, B.ST02 AS D02,OG13,OG14,OG15,OG16,OG17,OG18" & _
            " From OutGoing, STAFF A,STAFF B Where A.ST01(+)=OG03 AND B.ST01(+)=OG04 AND OG01='" & m_OG01 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      txtOG(1) = m_OG01
      For intI = 1 To 4
         txtOG(intI) = "" & rsTmp.Fields("OG" & Format(intI, "00"))
      Next intI
      
      'Add By Sindy 2020/9/7 顯示星期幾
      If Val(txtOG(2)) > 0 Then
         LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtOG(2)), "####/##/##"))) & ")"
      End If
      
      For intI = 6 To 9 '12
         txtOG(intI) = "" & rsTmp.Fields("OG" & Format(intI, "00"))
      Next intI
      For intI = 10 To 12
         txtOG2(intI) = "" & rsTmp.Fields("OG" & Format(intI, "00"))
      Next intI
      
      For intI = 19 To 20
         mebOG(intI) = "" & rsTmp.Fields("OG" & Format(intI, "00"))
      Next intI
      
      For intI = 1 To 2
         lblDisp(intI - 1) = "" & rsTmp.Fields("D" & Format(intI, "00"))
      Next intI
      Call UpdateCUID(rsTmp)
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      frm180301_01.Show
      Exit Sub
   End If
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
    
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ClearField()
   Dim oText As TextBox, oLabel As Label, oMaskEdBox As MaskEdBox
   Dim oText2 As Object
   
   For Each oText In txtOG
      oText.Text = ""
   Next
   
   'Add By Sindy 2021/5/28
   For Each oText2 In txtOG2
      oText2.Text = ""
   Next
   '2021/5/28 END
   
   For Each oMaskEdBox In mebOG
      oMaskEdBox.Text = "00:00"
   Next
   For Each oText2 In lblDisp
      oText2.Caption = ""
   Next
   
   Label26.Caption = ""
   Label27.Caption = ""
   LblStarW.Caption = "" 'Add By Sindy 2020/9/7
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '清空欄位值
   ClearField
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set m_PrevForm = Nothing 'Add By Sindy 2013/6/21
   Set frm180301_04 = Nothing
End Sub

Private Sub cmdExit_Click()
   'Add By Sindy 2013/6/21
   If UCase(m_PrevForm.Name) = UCase("frm160012") Then
      m_PrevForm.Show
      Set m_PrevForm = Nothing
      Unload Me
   Else
   '2013/6/21 END
      Set m_PrevForm = Nothing
      Unload Me
      Unload frm180301_01
      Unload frm180301
   End If
End Sub

Private Sub cmdQueryNext_Click()
   Unload Me
   'Modify By Sindy 2013/6/21
'   frm180301_01.Show
'   frm180301_01.PubShowNextData
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   Set m_PrevForm = Nothing
   '2013/6/21 END
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
   
   If IsNull(rsSrcTmp.Fields("og13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og13")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("og13"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("og14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og14")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("og14"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("og15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og15")) = False Then
         strTemp = rsSrcTmp.Fields("og15")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("og16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og16")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("og16"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("og17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og17")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("og17"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("og18")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("og18")) = False Then
         strTemp = rsSrcTmp.Fields("og18")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
   Label27.Caption = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'Add By Sindy 2013/6/21
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub
