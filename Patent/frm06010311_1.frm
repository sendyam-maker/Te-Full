VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010311_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-異議/舉發延期"
   ClientHeight    =   5700
   ClientLeft      =   75
   ClientTop       =   990
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7590
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010311_1.frx":0000
      Left            =   1320
      List            =   "frm06010311_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   57
      Top             =   1170
      Width           =   615
   End
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   6420
      TabIndex        =   54
      Top             =   1140
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   55
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   56
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   11
      Top             =   5340
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6420
      TabIndex        =   14
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4395
      TabIndex        =   12
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5250
      TabIndex        =   13
      Top             =   45
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   21
      Top             =   510
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1632
      MaxLength       =   6
      TabIndex        =   20
      Top             =   510
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2472
      MaxLength       =   1
      TabIndex        =   19
      Top             =   510
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2712
      MaxLength       =   2
      TabIndex        =   18
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4980
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2580
      Width           =   300
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2940
      Width           =   735
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   3
      Top             =   3270
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "1720;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   3600
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   3885
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   250
      Size            =   "7170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   4155
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "7170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   4470
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "7170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   4755
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "7170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   10
      Top             =   5025
      Width           =   4065
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "7170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   2625
      Width           =   975
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   7
      Size            =   "1720;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   5970
      TabIndex        =   4
      Top             =   3270
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5040
      TabIndex        =   53
      Top             =   3330
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "公告號或專利號數只列印申請書, 不存檔"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3120
      TabIndex        =   52
      Top             =   5370
      Width           =   3150
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "公告號或專利號數"
      Height          =   180
      Left            =   240
      TabIndex        =   51
      Top             =   5340
      Width           =   1440
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "對造號數:"
      Height          =   180
      Left            =   240
      TabIndex        =   50
      Top             =   3330
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(中):"
      Height          =   180
      Left            =   240
      TabIndex        =   49
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(英):"
      Height          =   180
      Left            =   240
      TabIndex        =   48
      Top             =   3885
      Width           =   1425
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(日):"
      Height          =   180
      Left            =   240
      TabIndex        =   47
      Top             =   4155
      Width           =   1425
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(中):"
      Height          =   180
      Left            =   240
      TabIndex        =   46
      Top             =   4470
      Width           =   1065
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(英):"
      Height          =   180
      Left            =   240
      TabIndex        =   45
      Top             =   4755
      Width           =   1065
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(日):"
      Height          =   180
      Left            =   240
      TabIndex        =   44
      Top             =   5025
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   7200
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   7200
      Y1              =   2535
      Y2              =   2535
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   43
      Top             =   2160
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   42
      Top             =   2160
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   8
      Left            =   2580
      TabIndex        =   41
      Top             =   2955
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期　:"
      Height          =   180
      Left            =   240
      TabIndex        =   40
      Top             =   2595
      Width           =   1125
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4080
      TabIndex        =   39
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   4080
      TabIndex        =   38
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   1830
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   36
      Top             =   510
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4080
      TabIndex        =   35
      Top             =   1500
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   1500
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   4080
      TabIndex        =   31
      Top             =   840
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   1170
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   29
      Top             =   840
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   28
      Top             =   840
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1950
      TabIndex        =   27
      Top             =   1170
      Width           =   4365
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   26
      Top             =   1500
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   25
      Top             =   1500
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   24
      Top             =   1830
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   23
      Top             =   1830
      Width           =   2640
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4657;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   3300
      TabIndex        =   22
      Top             =   2640
      Width           =   2880
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "延期案件性質:"
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   2970
      Width           =   1125
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   16
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   765
   End
End
Attribute VB_Name = "frm06010311_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/8 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer, intLastRow As Integer
Dim m_CP17 As String


Private Sub cmdOK_Click(Index As Integer)
 Dim bolChk As Boolean, strTmp As String
   Select Case Index
      Case 0 '確定
        'Add By Cheng 2003/06/24
        If Me.Text9.Text = "" Then
            MsgBox "請輸入延期案件性質或點選未收文期限資料!!!", vbExclamation + vbOKOnly
            Me.Text9.SetFocus
            Text9_GotFocus
            Exit Sub
        End If
         
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                 Exit Sub
             End If
         End If
         'end 2020/02/17
         
         'Add by Sindy 2021/11/8 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         Select Case Text9.Text
            Case 異議_專
               strTmp = "04"
            Case 舉發
               strTmp = "06"
            Case Else
         End Select
         StartLetter "01", strReceiveNo, strTmp
         NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
      Case 1 '回前畫面
         frm060103_1.Show
      Case 2 '結束
         Unload frm060103_1
   End Select
   Unload Me
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 5) As String, strTmp As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
         
    EndLetter ET01, ET02, ET03, strUserNum
    ii = 0
    Select Case Me.Text9.Text
    Case "801" '異議
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','申請書法條','" & IIf(pa(8) = "1", "第41條", IIf(pa(8) = "2", "第102條", "第115條")) & "')"
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','規費','" & m_CP17 & "')"
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','列印備註','" & Text8.Text & "')"
    Case "803" '舉發
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','申請書法條','" & IIf(pa(8) = "1", "第67條", IIf(pa(8) = "2", "第107條", "第128條")) & "')"
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','規費','" & m_CP17 & "')"
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','列印備註','" & Text8.Text & "')"
   End Select
    
   If ii <> 0 Then
       'edit by nickc 2007/02/05 不用 dll 了
       'If Not objLawDll.ExecSQL(ii, strTxt) Then
       If Not ClsLawExecSQL(ii, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
       End If
   End If

End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = pa(5)
      Case "英"
         Label12(3) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label12(3) = pa(7)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5(0).Text = strSrvDate(2)
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010311_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object, i As Integer
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Text5(0) = pa(10)
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      Label12(3) = pa(5)
   End If
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP17,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP110 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0)
         If Label12(0).Caption <> "延期" Then Text9.Enabled = False
         Text9 = .Fields(4)
         Label12(8).Caption = .Fields(0)
         If Me.Text9.Text = "404" Then
            Me.Text9.Text = ""
            Label12(8).Caption = ""
         End If
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
      If Not IsNull(.Fields(5)) Then Label12(9) = .Fields(5)
      If Not IsNull(.Fields(6)) Then Label12(10) = .Fields(6)
      If Not IsNull(.Fields(7)) Then m_CP17 = .Fields(7)
      For i = 4 To 10
         If Not IsNull(.Fields(i + 4)) Then Text5(i) = .Fields(i + 4)
      Next
   End If
   End With
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
End Sub

Private Sub Text5_LostFocus(Index As Integer)
   Select Case Index
      Case 7
         If Text5(5) = "" And Text5(6) = "" And Text5(7) = "" Then
            MsgBox "對造案件名稱不可同時空白 !", vbCritical
            Text5(5).SetFocus
         End If
      Case 10
         If Text5(8) = "" And Text5(9) = "" And Text5(10) = "" Then
            MsgBox "對造名稱不可同時空白 !"
            Text5(8).SetFocus
         End If
   End Select

End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         Cancel = Not ChkLetterDate(Text5(Index).Text)
      Case 4
         If Text5(4) = "" Then
            MsgBox "對造號數不可空白 !", vbCritical
            Cancel = True
         End If
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   Cancel = False
   If Text8 = "" Then
      MsgBox "公告號或專利號數不可空白 !", vbCritical
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text8
End Sub

Private Sub Text9_Change()
   ' Me.Label18(2).Visible = False
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
Dim strTempName As String
   
    If Me.Text9.Text = "" Then Exit Sub
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.GetCaseProperty("FCP", Text9, strTempName, False) Then
    If ClsPDGetCaseProperty("FCP", Text9, strTempName, False) Then
        Label12(8) = strTempName
    Else
        Label12(8) = ""
        Cancel = True
    End If
    If Cancel = True Then TextInverse Text9
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   TextInverse Text5(Index)
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
'檢查對造號數
If Text5(4) = "" Then
   MsgBox "對造號數不可空白 !", vbCritical
   Me.Text5(4).SetFocus
   Text5_GotFocus 4
   Exit Function
End If
If GetTextLength(Text5(4)) > 20 And InStr(Text5(4), ",") = 0 And InStr(Text5(4), ";") = 0 Then
    MsgBox "對造號數內容太長,無法寫入基本檔欄位中,請洽電腦中心", vbOKOnly, "檢核資料"
    Me.Text5(4).SetFocus
    Text5_GotFocus 4
    Exit Function
End If
'檢查對造案件名稱
If Text5(5) = "" And Text5(6) = "" And Text5(7) = "" Then
   MsgBox "對造案件名稱不可同時空白 !", vbCritical
   Text5(5).SetFocus
   Text5_GotFocus 5
   Exit Function
End If
'檢查對告名稱
If Text5(8) = "" And Text5(9) = "" And Text5(10) = "" Then
   MsgBox "對造名稱不可同時空白 !"
   Text5(8).SetFocus
   Text5_GotFocus 8
   Exit Function
End If
'檢查公告號或專利號數
If Text8 = "" Then
   MsgBox "公告號或專利號數不可空白 !", vbCritical
   Me.Text8.SetFocus
   Text8_GotFocus
   Exit Function
End If

   'Add by Morgan 2005/8/8
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If

TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim intStep As Integer
 
FormSave = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   intStep = 1
   
   'Modify by morgan 2005/8/8 加 cp110
   strExc(intStep) = "UPDATE CASEPROGRESS SET cp36=" & CNULL(ChgSQL(Text5(4))) & "," & _
      "cp37=" & CNULL(ChgSQL(Text5(5))) & ",cp38=" & CNULL(ChgSQL(Text5(6))) & "," & _
      "cp39=" & CNULL(ChgSQL(Text5(7))) & ",cp40=" & CNULL(ChgSQL(Text5(8))) & "," & _
      "cp41=" & CNULL(ChgSQL(Text5(9))) & ",cp42=" & CNULL(ChgSQL(Text5(10))) & "" & _
      ",cp110=" & CNULL(m_CP110) & " where cp09='" & strReceiveNo & "'"
    cnnConnection.Execute strExc(intStep)
    
   intStep = intStep + 1
   strExc(intStep) = "Update Patent Set PA05='" & Me.Text5(5).Text & "' ,PA06=" & CNULL(ChgSQL(Text5(6))) & ",PA07='" & Me.Text5(7).Text & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strExc(intStep)
   
   cnnConnection.CommitTrans
   Exit Function

CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function
'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
