VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060503 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費資料維護"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7464
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7464
   Begin VB.OptionButton Option1 
      Caption         =   "之後"
      Height          =   195
      Index           =   2
      Left            =   3780
      TabIndex        =   11
      Top             =   3660
      Width           =   705
   End
   Begin VB.OptionButton Option1 
      Caption         =   "當天"
      Height          =   195
      Index           =   0
      Left            =   2370
      TabIndex        =   9
      Top             =   3660
      Width           =   705
   End
   Begin VB.OptionButton Option1 
      Caption         =   "之前"
      Height          =   195
      Index           =   1
      Left            =   3060
      TabIndex        =   10
      Top             =   3660
      Width           =   705
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060503.frx":0000
      Left            =   1125
      List            =   "frm060503.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   22
      Top             =   2220
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5535
      TabIndex        =   15
      Top             =   60
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6420
      TabIndex        =   16
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "FCP"
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   2
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   3
      Top             =   540
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   4
      Top             =   510
      Width           =   800
   End
   Begin MSForms.TextBox Text5 
      Height          =   288
      Index           =   6
      Left            =   3840
      TabIndex        =   47
      Top             =   3276
      Width           =   1116
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1968;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "待收款後辦案:                         (輸入管制日期)"
      Height          =   180
      Index           =   18
      Left            =   2664
      TabIndex        =   48
      Top             =   3330
      Width           =   3528
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   1230
      TabIndex        =   12
      Top             =   3945
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   5
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   1365
      Index           =   5
      Left            =   1230
      TabIndex        =   13
      Top             =   4260
      Width           =   3210
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "5662;2408"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   1230
      TabIndex        =   8
      Top             =   3630
      Width           =   1110
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1958;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   1230
      TabIndex        =   7
      Top             =   3300
      Width           =   390
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "688;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   2970
      Width           =   390
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "688;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   2070
      TabIndex        =   6
      Top             =   2970
      Width           =   390
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "688;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   5580
      TabIndex        =   14
      Top             =   4230
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   15
      Left            =   270
      TabIndex        =   46
      Top             =   4005
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   11
      Left            =   2160
      TabIndex        =   45
      Top             =   3975
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2275;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   16
      Left            =   270
      TabIndex        =   44
      Top             =   4290
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人："
      Height          =   180
      Index           =   17
      Left            =   4485
      TabIndex        =   43
      Top             =   4290
      Width           =   1080
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   10
      Left            =   4470
      TabIndex        =   42
      Top             =   2550
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2275;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   5055
      TabIndex        =   41
      Top             =   570
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2275;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1125
      TabIndex        =   40
      Top             =   2550
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2275;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   9
      Left            =   270
      TabIndex        =   39
      Top             =   2550
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   2
      Left            =   4275
      TabIndex        =   38
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   10
      Left            =   3645
      TabIndex        =   37
      Top             =   2550
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   8
      Left            =   1770
      TabIndex        =   36
      Top             =   2220
      Width           =   5550
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9790;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   1125
      TabIndex        =   35
      Top             =   1890
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4470
      TabIndex        =   34
      Top             =   1560
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1125
      TabIndex        =   33
      Top             =   1560
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4470
      TabIndex        =   32
      Top             =   1230
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1125
      TabIndex        =   31
      Top             =   1230
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   30
      Top             =   1230
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Index           =   4
      Left            =   3645
      TabIndex        =   29
      Top             =   1230
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Index           =   5
      Left            =   270
      TabIndex        =   28
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Index           =   6
      Left            =   3645
      TabIndex        =   27
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Index           =   7
      Left            =   270
      TabIndex        =   26
      Top             =   1890
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   8
      Left            =   270
      TabIndex        =   25
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "巳繳年度:"
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   24
      Top             =   900
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   9
      Left            =   1125
      TabIndex        =   23
      Top             =   900
      Width           =   5730
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10107;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   225
      X2              =   7500
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   225
      X2              =   7300
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "暫不繳納：          (Y:暫不繳)"
      Height          =   180
      Index           =   14
      Left            =   270
      TabIndex        =   21
      Top             =   3330
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "指定日期："
      Height          =   180
      Index           =   13
      Left            =   270
      TabIndex        =   20
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第           至          年"
      Height          =   180
      Index           =   12
      Left            =   1215
      TabIndex        =   19
      Top             =   3030
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "繳納年度："
      Height          =   180
      Index           =   11
      Left            =   270
      TabIndex        =   18
      Top             =   3030
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   17
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm060503"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/8/27
Option Explicit

Dim m_PA05 As String, m_PA06 As String, m_PA07 As String, m_PA08 As String, m_PA09 As String, m_PA25 As String
Dim m_CP07 As String, m_CP110 As String, m_AgentName As String
Dim m_bolAddCanlendar As Boolean, m_CP13 As String 'Added by Morgan 2025/7/23

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      If TxtValidate Then
         'Add by Sindy 2021/11/25 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If FormSave Then
            FormClear True
            Text2.SetFocus
         Else
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         End If
      End If
   Case 2
      Unload Me
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(8) = m_PA05
      Case "英"
         Label2(8) = m_PA06
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(8) = m_PA07
   End Select
End Sub

Private Sub Command1_Click()
   Dim cp(4) As String
   Dim PA143 As String
   
   FormClear
   
   If Text1.Text = "" Or Text2.Text = "" Then
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      Text1.SetFocus
      Exit Sub
   End If
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   
   strExc(0) = "select * from patent where pa01='" & Text1 & "' and pa02='" & Text2 & "' and pa03='" & Text3 & "' and pa04='" & Text4 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Label2(9) = "" & .Fields("pa72")
      For intI = 0 To 4
         If Not IsNull(.Fields("pa" & (26 + intI))) Then
            If ClsLawLawGetName(.Fields("pa" & (26 + intI)), strExc(1)) Then
               Label2(3 + intI) = strExc(1)
            End If
         End If
      Next
      m_PA05 = "" & .Fields("PA05")
      m_PA06 = "" & .Fields("PA06")
      m_PA07 = "" & .Fields("PA07")
      m_PA08 = "" & .Fields("PA08")
      m_PA09 = "" & .Fields("PA09")
      m_PA25 = "" & .Fields("pa25")
      Combo1.ListIndex = 0
      Combo1_Click
      PA143 = "" & .Fields("PA143")
      End With
      
      'Added by Morgan 2014/7/16
      strExc(0) = Right("" & RsTemp("pa72"), 2)
      If Left(strExc(0), 1) = "," Then strExc(0) = Mid(strExc(0), 2)
      If Val(strExc(0)) < 6 Then
         strExc(1) = ""
         If PUB_GetCaseDiscStat(RsTemp("pa01") & RsTemp("pa02") & RsTemp("pa03") & RsTemp("pa04"), strExc(1)) = "Y" Then
            If InStr(strExc(1), "3") > 0 Then
               MsgBox "此案申請人已設定為中小企業,年費請以紙本親送!!", vbExclamation
               Text1.SetFocus
               Exit Sub
            End If
         End If
      End If
      'end 2014/7/16
      
      strExc(0) = "select * from caseprogress where cp01='" & Text1 & "' and cp02='" & Text2 & "' and cp03='" & Text3 & "' and cp04='" & Text4 & "' and cp10='605' and cp27||cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            Label2(1) = .Fields("cp09")
            Label2(10) = ChangeWStringToTDateString("" & .Fields("cp07"))
            Text5(0) = "" & .Fields("cp53")
            Text5(1) = "" & .Fields("cp54")
            If .Fields("cp141") = "4" Then
               Text5(2) = "Y"
            End If
            If Not IsNull(.Fields("cp142")) Then
               Text5(3) = TransDate(.Fields("cp142"), 1)
'               'Add By Sindy 2022/6/21 亭妙需求:
'               If Text5(2) <> "Y" Then
'                  '1.系統日< 指定日期 時，暫不繳:  Y
'                  '2.系統日=及> 指定日期 時，暫不繳納: __(空)
'                  If Val(strSrvDate(2)) < Val(Text5(3)) Then
'                     Text5(2) = "Y"
'                  End If
'               End If
'               '2022/6/21 END
            End If
            
            'Add By Sindy 2021/4/20
            If "" & .Fields("CP164") = "1" Then
               Option1(0).Value = True
            ElseIf "" & .Fields("CP164") = "2" Then
               Option1(1).Value = True
            'Add By Sindy 2022/6/21
            ElseIf "" & .Fields("CP164") = "3" Then
               Option1(2).Value = True
            End If
            '2021/4/20 END
            
            If ClsPDGetStaff("" & .Fields("CP13"), strExc(1)) Then
               Label2(2) = strExc(1)
            End If
      
            Text5(4) = "" & .Fields("cp14")
            Text5_Validate 4, False
            Text5(5) = "" & .Fields("cp64")
            m_CP110 = "" & .Fields("cp110")
            If PA143 = "N" Then
               lstNameAgent.Enabled = False
               m_CP110 = ""
            Else
               cp(1) = .Fields("cp01")
               cp(2) = .Fields("cp02")
               cp(3) = .Fields("cp03")
               cp(4) = .Fields("cp04")
               'Modified by Morgan 2021/2/3 +"605"
               PUB_SetOurAgent lstNameAgent, cp(), m_CP110, "605", True
               'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
               lstNameAgent.Height = 1300
               lstNameAgent.Width = 1300
            End If
            m_CP07 = "" & .Fields("cp07")
            m_CP13 = "" & .Fields("cp13") 'Added by Morgan 2025/7/23
         End With
         If Text5(0) = "" Then Text5(0).SetFocus
      Else
         MsgBox "本案無未發文年費，請重新輸入 !", vbCritical
         Text1.SetFocus
         Exit Sub
      End If
   Else
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      Text1.SetFocus
      Exit Sub
   End If
   
   cmdok(0).Enabled = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear
   
   'Add By Sindy 2021/4/29
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      Option1(0).Visible = True
      Option1(1).Visible = True
      Option1(2).Visible = True 'Add By Sindy 2022/6/21
   Else
      Option1(0).Visible = False
      Option1(1).Visible = False
      Option1(2).Visible = False 'Add By Sindy 2022/6/21
   End If
   '2021/4/29 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060503 = Nothing
End Sub

' 清除資料表
Private Sub FormClear(Optional pbAll As Boolean)
   Dim oLabel As Object
   Dim oText As Object
   
   If pbAll Then
      Text1 = "FCP"
      Text2 = ""
      Text3 = ""
      Text4 = ""
   End If
   
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   
   For Each oText In Text5
      oText.Text = ""
   Next
   
   lstNameAgent.Clear
   cmdok(0).Enabled = False
End Sub

Private Sub lstNameAgent_Validate(Cancel As Boolean)
   If lstNameAgent.Enabled = False Then Exit Sub
   
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/11 員工編號已可非數字需做轉換
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

Private Sub Text1_Change()
   If Me.cmdok(0).Enabled = True Then
      FormClear
   End If
End Sub

Private Sub Text2_Change()
   If Me.cmdok(0).Enabled = True Then
      FormClear
   End If
End Sub

Private Sub Text3_Change()
   If Me.cmdok(0).Enabled = True Then
      FormClear
   End If
End Sub

Private Sub Text4_Change()
   If Me.cmdok(0).Enabled = True Then
      FormClear
   End If
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   TextInverse Text5(Index)
   If Index = 5 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
   Case 0, 1, 3
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
      End If
   Case 2
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      'Added by Morgan 2025/7/23
      ElseIf KeyAscii = Asc("Y") Then
         Text5(6).Enabled = True
      Else
         Text5(6) = ""
         Text5(6).Enabled = False
      'end 2025/7/23
      End If
   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Dim strStartDate As String
   
   Select Case Index
   Case 0
      If Text5(0) <> "" Then
         strExc(1) = Right(Label2(9), 2)
         If Left(strExc(1), 1) = "," Then strExc(1) = Mid(strExc(1), 2)
         If Text5(0) <> Val(strExc(1)) + 1 Then
            MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
            Cancel = True
         End If
      End If
   Case 1
      If Text5(1) <> "" Then
         If Val(Text5(1)) < Val(Text5(0)) Then
            MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
            Cancel = True
         Else
            If m_PA25 = "" Then
               If GetMoneyDate(Val(m_PA08) + 10, m_PA09, strExc, strExc(1), strExc(2), m_PA25) = False Then '抓專用期起止日
                  Cancel = True
                  Exit Sub
               End If
            End If
            strExc(0) = Label2(1)
            strExc(1) = Text1
            strExc(2) = Text2
            strExc(3) = Text3
            strExc(4) = Text4
            
            If GetMoneyDate(Val(m_PA08), m_PA09, strExc, strStartDate, strExc(5)) = False Then    '抓年費起算起日
               Cancel = True
               Exit Sub
            End If
            
            strExc(0) = CompDate(0, Text5(1) - 1, strStartDate)
            If Val(strExc(0)) > Val(m_PA25) Then
               MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
               Cancel = True
            End If
         End If
      End If
   Case 3
      If Text5(3) <> "" Then
         If Not ChkDate(Text5(3)) Then
            Cancel = True
         ElseIf m_CP07 <> "" And Val(DBDATE(Text5(3))) > Val(m_CP07) Then
            MsgBox "指定日期不可大於法定期限 !", vbCritical
            Cancel = True
         End If
      End If
   Case 4
      If ClsPDGetStaff(Text5(4), strExc(1)) Then
         Label2(11) = strExc(1)
      Else
         Cancel = True
      End If
   
   'Added by Morgan 2025/7/23
   Case 6
      If Text5(Index) <> "" Then
         If Not ChkDate(Text5(Index)) Then
            Cancel = True
         ElseIf m_CP07 <> "" And Val(DBDATE(Text5(Index))) > Val(m_CP07) Then
            MsgBox "管制日期不可大於法定期限 !", vbCritical
            Cancel = True
         End If
      End If
      
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If Text5(0) = "" Then
      MsgBox "起始繳費年度不可空白！", vbExclamation
      Text5(0).SetFocus
      Exit Function
   Else
      Text5_Validate 0, bCancel
      If bCancel Then Exit Function
   End If
   If Text5(1) = "" Then
      MsgBox "繳費年度不可空白！", vbExclamation
      Text5(1).SetFocus
      Exit Function
   Else
      Text5_Validate 1, bCancel
      If bCancel Then Exit Function
   End If
   Text5_Validate 3, bCancel
   If bCancel Then Exit Function
   If Text5(4) = "" Then
      MsgBox "承辦人不可空白！", vbExclamation
      Text5(4).SetFocus
      Exit Function
   Else
      Text5_Validate 4, bCancel
      If bCancel Then Exit Function
   End If
   
   lstNameAgent_Validate bCancel
   If bCancel Then Exit Function
   
   'Add By Sindy 2021/4/20 檢查指定送件日相關欄位
   'Modify By Sindy 2022/6/21 + And Option1(2).Value = False
   If Val(Text5(3).Text) > 0 And Option1(0).Visible = True Then
      If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
         MsgBox "有輸入指定送件日，當天 或 之前 或 之後 請擇一。", vbExclamation
         Exit Function
      End If
   Else
      Option1(0).Value = False
      Option1(1).Value = False
      Option1(2).Value = False 'Add By Sindy 2022/6/21
   End If
   '2021/4/20 END
   
   'Added by Morgan 2025/7/23
   m_bolAddCanlendar = False
   If Text5(6) <> "" Then
      strExc(0) = "select * from Staff_Calendar where sc04='待收款後辦案' and sc05='" & Text1 & "' and sc06='" & Text2 & "' and sc07='" & Text3 & "' and sc08='" & Text4 & "' and sc18 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "該筆資料已存在行事曆，請至行事曆資料維護更新！", vbExclamation
      ElseIf intI = 0 Then
         m_bolAddCanlendar = True
      Else
         Exit Function
      End If
   End If
   'end 2025/7/23
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim stCP141 As String
   Dim strCon As String
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   'Added by Morgan 2025/7/23
   If m_bolAddCanlendar Then
      strExc(1) = DBDATE(Text5(6)) '管制日期
      strExc(2) = Text5(4) '年費承辦人
      strExc(5) = strExc(2)
      strExc(0) = PUB_GetFCPHandler(Text1, Text2, Text3, Text4, "605") '程序管制人
      If InStr(strExc(5), strExc(0)) = 0 Then
         strExc(5) = strExc(5) & "," & strExc(0)
      End If
      strExc(5) = strExc(5) & "," & m_CP13 '+收文智權人員
      strExc(0) = PUB_GetFCPSalesNo(Text1, Text2, Text3, Text4, "605") '承辦智權人員
      If InStr(strExc(5), strExc(0)) = 0 Then
         strExc(5) = strExc(5) & "," & strExc(0)
      End If
      '提醒人員:該道年費承辦人、該區程序(抓取順序: 1. 年費代理人(如有) 2. 國外代理人) + 該道年費智權人員、該區承辦(抓取順序: 1. 年費代理人(如有) 2. 國外代理人)
      '可解除人員=提醒人員
      PUB_AddFCPStaffCalendar strExc(1), "1", strExc(5), "待收款後辦案", strExc(5), "1", Text1, Text2, Text3, Text4
      Text5(5) = "待收款後辦案管制" & ChangeTStringToTDateString(Text5(6)) & "；" & Text5(5)
   End If
   'end 2025/7/23
   
   
   
   If Text5(2) = "Y" Then
      stCP141 = "4"
   ElseIf Text5(3) <> "" Then
      stCP141 = "3"
   Else
      stCP141 = "1"
   End If
   
   'Modify By Sindy 2021/4/20
   strCon = ""
   'Added by Morgan 2025/7/23 若指定日期清除,指定日期方式也要清
   If Text5(3) = "" Then
      strCon = ",cp164=''"
   'end 2025/7/23
   ElseIf Option1(0).Value = True Then
      strCon = ",cp164='1'"
   ElseIf Option1(1).Value = True Then
      strCon = ",cp164='2'"
   'Add By Sindy 2022/6/21
   ElseIf Option1(2).Value = True Then
      strCon = ",cp164='3'"
   End If
   '2021/4/20 END
   strSql = "update caseprogress set cp53='" & Text5(0) & "',cp54='" & Text5(1) & "',cp141='" & stCP141 & "'" & _
      ",cp142=" & CNULL(DBDATE(Text5(3)), True) & strCon & ",cp14='" & Text5(4) & "',cp64='" & ChgSQL(Text5(5)) & "',cp110='" & m_CP110 & "'" & _
      " where cp09='" & Label2(1) & "'"
   Pub_SeekTbLog strSql 'Added by Morgan 2020/3/4
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function
