VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010029_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "登錄取消發文"
   ClientHeight    =   3510
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   Begin VB.TextBox txtCP132 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1890
      MaxLength       =   100
      TabIndex        =   0
      Text            =   "980504"
      Top             =   1860
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2385
      TabIndex        =   3
      Top             =   45
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   45
      Width           =   1200
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "frm010029_1.frx":0000
      Left            =   180
      List            =   "frm010029_1.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin MSForms.TextBox txtCP131 
      Height          =   900
      Left            =   180
      TabIndex        =   1
      Top             =   2490
      Width           =   4335
      VariousPropertyBits=   -1466941413
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "7646;1587"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "發文室取消發文日："
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1860
      Width           =   1665
   End
   Begin VB.Label Label4 
      Caption         =   "發文室取消發文備註："
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   2190
      Width           =   4350
   End
   Begin VB.Label Label2 
      Caption         =   "欲取消發文之案件："
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   4350
   End
End
Attribute VB_Name = "frm010029_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 txtCP131
'Added by Morgan 2014/4/29
Option Explicit

Public strNoList As String     '收文號清單
Public strCP131 As String      '發文室取消發文備註(回傳)
Public strCP132 As String      '發文室取消發文日期(回傳)
Public BolOk As Boolean        'True: 確定  False: 取消


Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   '確定
   If Index = 0 Then
      'Add by Amy 2021/12/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

      PUB_FilterFormText Me 'Add by Sindy 2009/05/13 修正畫面所有含跳行符號的文字框
      
      '檢查資料正確性
      Cancel = False
      txtCP131_Validate Cancel
      If Cancel = True Then
         Exit Sub
      End If
      
      strCP132 = Trim(txtCP132.Text)
      strCP131 = Trim(txtCP131.Text)
      BolOk = True
      
   '回前畫面(取消)
   Else
      strCP132 = ""
      strCP131 = ""
      BolOk = False
   End If
   Me.Hide
End Sub

Public Function CheckShowList() As Boolean
   Dim stSQL As String, stDesc As String
   Dim intIdx As Integer, ii As Integer
   Dim ArrCP09
   Dim stCon As String
   
   CheckShowList = False
   lstData.Clear
   '預設值
   txtCP132 = strSrvDate(2)
   
   ArrCP09 = Split(strNoList, ",")
   stSQL = ""
   For ii = LBound(ArrCP09) To UBound(ArrCP09)
      If ArrCP09(ii) <> "" Then
         stSQL = IIf(stSQL = "", "", stSQL & " UNION ALL ") & "SELECT CP01,CP02,CP03,CP04,CP09,CP43,CPM03 FROM CaseProgress,casepropertymap WHERE CP09='" & ArrCP09(ii) & "' AND CP01=CPM01(+) and CP10=CPM02(+)"
      End If
   Next
   
   '取得案號、案件性質
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If Not IsNull(RsTemp("cp43")) Then
            stDesc = PUB_GetRelateCasePropertyName(RsTemp("cp09"), "1")
         Else
            stDesc = ""
         End If
         intIdx = lstData.ListCount
         lstData.AddItem RsTemp("cp01") & "-" & RsTemp("cp02") & IIf(RsTemp("cp03") & RsTemp("cp04") = "000", "", "-" & RsTemp("cp03") & "-" & RsTemp("cp04")) & "　　" & RsTemp("cp09") & "　　" & RsTemp("cpm03") & stDesc, intIdx
         RsTemp.MoveNext
      Loop
      CheckShowList = True
   End If
   BolOk = True
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub txtCP131_GotFocus()
   TextInverse txtCP131
   OpenIme
End Sub

Private Sub txtCP131_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If IsEmptyText(txtCP131) = False Then
      If CheckLengthIsOK(txtCP131, 100) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文室取消發文備註內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCP131_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub
