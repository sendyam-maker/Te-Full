VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010517 
   BorderStyle     =   1  '單線固定
   Caption         =   "去電記錄"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9315
   Begin VB.OptionButton Option2 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   3735
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   795
      Width           =   1230
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號："
      Height          =   180
      Left            =   405
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   795
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7065
      TabIndex        =   5
      Top             =   690
      Width           =   800
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2280
      Width           =   885
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1980
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   0
      Top             =   750
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   7470
      TabIndex        =   14
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8310
      TabIndex        =   13
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   6675
      MaxLength       =   2
      TabIndex        =   4
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   6435
      MaxLength       =   1
      TabIndex        =   3
      Top             =   750
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5595
      MaxLength       =   6
      TabIndex        =   2
      Top             =   750
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   5115
      MaxLength       =   3
      TabIndex        =   1
      Top             =   750
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frm04010517.frx":0000
      Height          =   1650
      Left            =   135
      TabIndex        =   29
      Top             =   3840
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   2910
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Left            =   1395
      TabIndex        =   10
      Top             =   2610
      Width           =   885
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   5115
      TabIndex        =   8
      Top             =   1980
      Width           =   1695
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2990;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   840
      Left            =   135
      TabIndex        =   11
      Top             =   2940
      Width           =   8940
      VariousPropertyBits=   -1467989989
      Size            =   "15769;1482"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7095
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12515;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "去電內容："
      Height          =   180
      Index           =   0
      Left            =   8040
      TabIndex        =   33
      Top             =   2640
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   2340
      TabIndex        =   32
      Top             =   2655
      Width           =   3015
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5318;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   135
      Top             =   570
      Width           =   8880
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   28
      Top             =   1770
      Width           =   7080
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "12488;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   27
      Top             =   2310
      Width           =   2550
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4498;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   26
      Top             =   2310
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3149;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "存檔後會自動發郵件給王副總及副本收受者。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   11
      Left            =   180
      TabIndex        =   25
      Top             =   60
      Width           =   6195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受者："
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   24
      Top             =   2655
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "記錄發信收受者："
      Height          =   180
      Index           =   9
      Left            =   4005
      TabIndex        =   23
      Top             =   2310
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "處理人員："
      Height          =   180
      Index           =   8
      Left            =   225
      TabIndex        =   22
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分機號碼："
      Height          =   180
      Index           =   7
      Left            =   225
      TabIndex        =   21
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "去電對象："
      Height          =   180
      Index           =   6
      Left            =   4005
      TabIndex        =   20
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   19
      Top             =   1770
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   5115
      TabIndex        =   18
      Top             =   1230
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4604;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Index           =   3
      Left            =   4005
      TabIndex        =   17
      Top             =   1230
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   1170
      TabIndex        =   16
      Top             =   1230
      Width           =   1770
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3122;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   225
      TabIndex        =   15
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   4
      Left            =   225
      TabIndex        =   12
      Top             =   1500
      Width           =   900
   End
End
Attribute VB_Name = "frm04010517"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (MSHFlexGrid1,Text6,Text2,Text9,Combo1,Label2)
'Created by Lydia 2015/11/05 去電記錄
Option Explicit

Dim m_CP13 As String
Dim m_CaseName As String
Dim m_CP36 As String
Dim m_strReceiver As String 'Added by Lydia 2023/04/24

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         cmdOK(0).Enabled = False
         If TxtValidate = True Then
            If FormSave = True Then
               PUB_SendMailCache
               FormClear
               If Option1.Value Then
                  Text5.SetFocus
               Else
                  Text1(1).SetFocus
               End If
            Else
               cmdOK(0).Enabled = True
            End If
         Else
            cmdOK(0).Enabled = True
         End If
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   If GetCaseData = False Then
      If Option1.Value Then
         Text5.SetFocus
         Text5_GotFocus
      Else
         Text1(1).SetFocus
         Text1_GotFocus 1
      End If
   Else
      cmdOK(0).Enabled = True
   End If
End Sub

Private Sub Form_Load()
   'Added by Lydia 2023/04/24 修改王副總退休之相關控制
   If strSrvDate(1) >= "20230511" Then
       Label1(11).Caption = "存檔後會自動發郵件給李經理及副本收受者。"
       m_strReceiver = "99050"
   ElseIf strSrvDate(1) >= "20230501" Then
       Label1(11).Caption = "存檔後會自動發郵件給王副總、李經理及副本收受者。"
       m_strReceiver = "71011;99050"
   Else
       Label1(11).Caption = "存檔後會自動發郵件給王副總及副本收受者。"
       m_strReceiver = "71011"
   End If
   'end 2023/04/24
   
   MoveFormToCenter Me
   FormClear True

End Sub

'游經理跟秀玲反應不知道按了什麼畫面被關掉了
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MsgBox("是否確定要結束？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010517 = Nothing
End Sub

Private Sub Option1_Click()
   Text1(1).Enabled = False
   Text1(2).Enabled = False
   Text1(3).Enabled = False
   Text1(4).Enabled = False
   Text5.Enabled = True
   Text5.SetFocus
End Sub

Private Sub Option2_Click()
   Text1(1).Enabled = True
   Text1(2).Enabled = True
   Text1(3).Enabled = True
   Text1(4).Enabled = True
   Text5.Enabled = False
   Text1(1).SetFocus
End Sub

Private Sub Text1_Change(Index As Integer)
   If Text1(Index).Tag <> "" Then
      FormClear False
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Change()
   Label2(5) = ""
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'修正輸入中文轉員工編號後中文名稱不會顯示問題
Private Sub Text2_Validate(Cancel As Boolean)
   Text2 = Trim(Text2)
   If Text2 <> "" And Text2.Tag <> Text2 Then
      'Modified by Morgan 2017/10/25 不必抓第1碼否則6字頭的員工號會被剔除
      'If Left(Text2, 1) > "6" And Left(Text2, 1) < "F" Then
      If Text2 > "6" And Text2 < "F" Then
         If ClsPDGetStaff(Text2, strExc(1)) Then
            Label2(5) = strExc(1)
         Else
            Cancel = True
            Text2_GotFocus
         End If
      Else
         If GetIdFromName(Text2, strExc(1)) Then
            strExc(0) = Text2
            Text2 = strExc(1)
            Label2(5) = strExc(0)
         Else
            Cancel = True
            Text2_GotFocus
         End If
      End If
      Text2.Tag = Text2
   End If
End Sub

Private Sub Text5_Change()
   If Text5.Tag <> "" Then
      FormClear False
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   CloseIme
End Sub

Private Function FormSave() As Boolean
   Dim strCP09 As String, strCP12 As String
   Dim strCP64 As String, strDate As String
   Dim strReceiver As String, strCC As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   
   strCP09 = AutoNo("B", 6)
   strCP12 = GetSalesArea(m_CP13)
   strCP64 = "去電對象：" & Text6 & ", 分機號碼：" & Text7 & ", 去電內容：" & Text9
   
   If Text2 <> "" Then
      strCC = Text2 & ";"
   End If
   
   '同時上發文日為系統日，進度備註存"去電對象，分機號碼，去電內容'；不存本所期限及承辦期限。
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
      "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP64) VALUES " & _
      "('" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "'," & strSrvDate(1) & _
      ",'" & strCP09 & "','955','90'," & CNULL(strCP12) & "," & CNULL(m_CP13) & _
      ",'" & Text8 & "','N','N','" & strSrvDate(1) & "','N','" & ChgSQL(m_CP36) & "','" & ChgSQL(strCP64) & "') "
   cnnConnection.Execute strSql, intI
   
   '記錄發信收受者,預設71011王副總
   'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
   'strReceiver = "71011"
   strReceiver = m_strReceiver
   
   strExc(1) = Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & " (" & strCP09 & ") 去電記錄通知!!"
   'Modified by Lydia 2017/07/27 改變文字大小
   'strExc(2) = GetMailText
   strExc(2) = GetMailTextNew
   strExc(3) = PUB_LeftB(strExc(2), 4000)
   strExc(4) = Mid(strExc(2), Len(strExc(3)) + 1)
   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc10,mc11)" & _
      " values ('" & strUserNum & "','" & strReceiver & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
      ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "','" & strCC & "','Y','" & ChgSQL(strExc(4)) & "')"
   cnnConnection.Execute strSql, intI
   

   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Function TxtValidate() As Boolean
  
   If Text7 = "" Then
      MsgBox "請輸入分機號碼!", vbInformation
      Text7.SetFocus
      Exit Function
   ElseIf Text6 = "" Then
      MsgBox "請輸入去電對象!", vbInformation
      Text6.SetFocus
      Exit Function
   End If
   If Text9 = "" Then
      MsgBox "請輸入去電內容!", vbInformation
      Text9.SetFocus
      Exit Function
   End If
   If Text8 = "" Then
      MsgBox "請輸入處理人員!", vbInformation
      Text8.SetFocus
      Exit Function
   ElseIf Label2(3) = "" Then
      MsgBox "處理人員輸入錯誤!", vbExclamation
      Text8.SetFocus
      Exit Function
   End If
   
   If Text2 <> "" Then
      If Label2(5) = "" Then
         MsgBox "副本收受者輸入錯誤!", vbExclamation
         Text2.SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function
Private Function GetCaseData() As Boolean
   Dim bolFound As Boolean
   Dim ii As Integer

   m_CaseName = ""
   m_CP36 = ""
   
   '可用本所案號尋找
   If Option1.Value Then
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName" & _
         " from patent,customer where pa11='" & ChgSQL(Text5) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   Else
      If Text1(3) = "" Then Text1(3) = "0"
      If Text1(4) = "" Then Text1(4) = "00"
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName" & _
         " from patent,customer where pa01='" & Text1(1) & "' and pa02='" & Text1(2) & "' and pa03='" & Text1(3) & "' and pa04='" & Text1(4) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      bolFound = True
   ElseIf intI = 0 And Option1.Value Then
      '抓對造號數
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa10,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName" & _
         " from caseprogress,patent,customer where cp36='" & ChgSQL(Text5) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_CP36 = Text5
         bolFound = True
      End If
   End If
   
   If Not bolFound Then
      MsgBox "申請案號輸入錯誤！", vbExclamation
   
   ElseIf Pub_StrUserSt03 = "F22" And RsTemp.Fields("pa01") = "P" Then
      MsgBox "不可輸入 P 案！", vbExclamation
      
   Else
      With RsTemp
      
      If Option2.Value Then Text5 = "" & .Fields("pa11")
      Text1(1) = "" & .Fields("pa01")
      Text1(1).Tag = Text1(1)
      Text1(2) = "" & .Fields("pa02")
      Text1(2).Tag = Text1(2)
      Text1(3) = "" & .Fields("pa03")
      Text1(3).Tag = Text1(3)
      Text1(4) = "" & .Fields("pa04")
      Text1(4).Tag = Text1(4)
      Text5.Tag = Text5
      
      If Not IsNull(.Fields("pa05")) Then
         m_CaseName = .Fields("pa05")
      ElseIf Not IsNull(.Fields("pa06")) Then
         m_CaseName = .Fields("pa06")
      ElseIf Not IsNull(.Fields("pa07")) Then
         m_CaseName = .Fields("pa07")
      End If
      AddCboName Combo1, "" & .Fields("pa05"), "" & .Fields("pa06"), "" & .Fields("pa07")
      m_CP13 = PUB_GetAKindSalesNo(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text)
      Label2(0) = GetPrjSalesNM(m_CP13)
      Label2(1) = "" & .Fields("pa47")
      Label2(2) = "" & .Fields("CuName")
      End With
      
      MSHFlexGrid1.Visible = False
      '所有排序條件加DESC,改同共同查詢(王副總)
      strExc(0) = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質" & _
         ",CP43 as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員" & _
         ",SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日" & _
         " from caseprogress,PATENT,casepropertymap,staff s1,staff s2" & _
         " WHERE cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cp01=cpm01(+) and cp10=cpm02(+) and s1.st01(+)=cp14 and s2.st01(+)=cp13" & _
         " ORDER BY SQLDatet2(CP05) DESC,CP66 DESC,CP67 DESC,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      If intI = 1 Then
         With MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 4) <> "" Then
                .TextMatrix(ii, 3) = .TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(ii, 2), "1")
            End If
         Next
         
         End With
      End If
      MSHFlexGrid1.Visible = True
      
      GetCaseData = True
   End If
End Function

Private Sub FormClear(Optional pbolAll As Boolean)
   Dim oLabel As Object
   If pbolAll Then
      Text1(1) = ""
      Text1(2) = ""
      Text1(3) = ""
      Text1(4) = ""
      Text5 = ""
   Else
      If Option1.Value Then
         Text1(1) = ""
         Text1(2) = ""
         Text1(3) = ""
         Text1(4) = ""
      Else
         Text5 = ""
      End If
   End If
   Text1(1).Tag = ""
   Text1(2).Tag = ""
   Text1(3).Tag = ""
   Text1(4).Tag = ""
   Text2 = ""
   Text2.Tag = ""
   Text5.Tag = ""
   Text6 = ""
   Text7 = ""
   Text9 = ""
   Combo1.Clear
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   GridHead
   Me.cmdOK(0).Enabled = False
   
   '預設處理人員為系統操作者,不可變更
   Text8.Text = strUserNum
   Text8.Tag = Text8.Text
   Text8.Enabled = False
   Label2(3).Caption = strUserName
   '記錄發信收受者,預設71011王副總
   'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
   'Label2(4).Caption = GetStaffName("71011")
   Label2(4).Caption = PUB_ReadUserData(m_strReceiver)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   OpenIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub

Private Sub Text8_Change()
   Label2(3) = ""
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'修正輸入中文轉員工編號後中文名稱不會顯示問題
Private Sub Text8_Validate(Cancel As Boolean)
   Text8 = Trim(Text8)
   If Text8 <> "" And Text8.Tag <> Text8 Then
      If Text8 > "6" And Text8 < "F" Then
         If ClsPDGetStaff(Text8, strExc(1)) Then
            Label2(3) = strExc(1)
         Else
            Cancel = True
            Text8_GotFocus
         End If
      Else
         If GetIdFromName(Text8, strExc(1)) Then
            strExc(0) = Text8
            Text8 = strExc(1)
            Label2(3) = strExc(0)
         Else
            Cancel = True
            Text8_GotFocus
         End If
      End If
      Text8.Tag = Text8
   End If
End Sub

Private Function GetIdFromName(ByVal pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02='" & ChgSQL(pName) & "' and st04='1' and st01>'6' and st01<'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function

Private Sub Text9_GotFocus()
   TextInverse Text9
   OpenIme
End Sub

Private Function GetMailText() As String
   Dim strText As String
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=600 STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=5><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(去電記錄通知)"
   strText = strText & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">日　期</FONT></TD>"
   strText = strText & "<TD COLSPAN=""3"" VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text5 & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(2) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">本所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>分所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & IIf(Label2(1) = "", "　", Label2(1)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">案件名稱</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & m_CaseName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">去電對象</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text6 & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">分機號碼</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text7 & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">處理人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(3) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">智權人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(0) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>去電內容：<BR>" & Text9 & "</FONT></TD></TR>"
   strText = strText & "</TABLE>"
   
   'Removed by Morgan 2020/3/30 取消公司名稱
   'strText = strText & "<TABLE WIDTH=600>"
   'strText = strText & "<TR border=0><TD VALIGN=""TOP"" HEIGHT=21>"
   'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
   'strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""RIGHT"">台一國際專利法律事務所</FONT></TD></TR>"
   'strText = strText & "</TABLE>"
   'end 2020/3/30
   
   GetMailText = strText
End Function

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .Cols = 11
      .row = 0
      .col = 0: .Text = "V"
      .ColWidth(0) = 0 '180
      .col = 1: .Text = "收文日"
      .ColWidth(1) = 800 '788
      .CellAlignment = flexAlignRightCenter
      .col = 2: .Text = "總收文號"
      .ColWidth(2) = 1000 '938
      .CellAlignment = flexAlignLeftCenter
      .col = 3: .Text = "案件性質"
      .ColWidth(3) = 2000 ' 950
      .CellAlignment = flexAlignLeftCenter
      .col = 4: .Text = "相關收文號"
      .ColWidth(4) = 0
      .CellAlignment = flexAlignLeftCenter
      .col = 5: .Text = "承辦人"
      .ColWidth(5) = 650 ' 593
      .CellAlignment = flexAlignLeftCenter
      .col = 6: .Text = "智權人員"
      .ColWidth(6) = 650 ' 593
      .CellAlignment = flexAlignLeftCenter
      .col = 7: .Text = "本所期限"
      .ColWidth(7) = 820 '788
      .CellAlignment = flexAlignRightCenter
      .col = 8: .Text = "法定期限"
      .ColWidth(8) = 820 '788
      .CellAlignment = flexAlignRightCenter
      .col = 9: .Text = "發文日"
      .ColWidth(9) = 800 '788
      .CellAlignment = flexAlignRightCenter
      .col = 10: .Text = "取消收文日"
      .ColWidth(10) = 1000 '788
      .CellAlignment = flexAlignLeftCenter
      .Visible = True
   End With
End Sub

'Added by Lydia 2017/07/26 因為W7的解析度不同,改變字體
Private Function GetMailTextNew() As String
   Dim strText As String
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=620px STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37><span style=""font-size:22px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(去電記錄通知)"
   strText = strText & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""15%"" VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">日　期</FONT></span></TD>"
   strText = strText & "<TD WIDTH=""30%"" VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""15%"" VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">申請案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text5 & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">申請人</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(2) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">本所案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">分所案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & IIf(Label2(1) = "", "　", Label2(1)) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">案件名稱</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & m_CaseName & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">去電對象</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text6 & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">分機號碼</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text7 & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">處理人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(3) & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">智權人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(0) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P>去電內容：<BR>" & Text9 & "</FONT></span></TD></TR>"
   strText = strText & "</TABLE>"
   
   'Removed by Morgan 2020/3/30 取消公司名稱
   'strText = strText & "<TABLE WIDTH=620px>"
   'strText = strText & "<TR border=0><TD VALIGN=""TOP"" HEIGHT=21><span style=""font-size:18px"">"
   'strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""RIGHT"">台一國際專利法律事務所</FONT></span></TD></TR>"
   'end 2020/3/30
   'strText = strText & "</TABLE>"
   
   GetMailTextNew = strText
End Function
