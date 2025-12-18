VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090217_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例資料彙整作業"
   ClientHeight    =   4695
   ClientLeft      =   1830
   ClientTop       =   1950
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8055
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1125
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1140
      TabIndex        =   1
      Top             =   645
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   1140
      TabIndex        =   2
      Top             =   930
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   1140
      TabIndex        =   3
      Top             =   1230
      Width           =   1215
   End
   Begin VB.ComboBox cboPC19 
      Height          =   300
      ItemData        =   "frm090217_2.frx":0000
      Left            =   1140
      List            =   "frm090217_2.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1500
      Width           =   1215
   End
   Begin VB.TextBox txtPC20 
      Height          =   270
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   6675
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton Command 
      Caption         =   "刪除(&D)"
      Height          =   405
      Index           =   1
      Left            =   5895
      TabIndex        =   15
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton Command 
      Caption         =   "修改(M)"
      Height          =   405
      Index           =   0
      Left            =   5115
      TabIndex        =   14
      Top             =   70
      Width           =   756
   End
   Begin MSForms.TextBox Text1 
      Height          =   615
      Index           =   10
      Left            =   1140
      TabIndex        =   12
      Top             =   3150
      Width           =   6735
      VariousPropertyBits=   -1467989989
      MaxLength       =   400
      ScrollBars      =   2
      Size            =   "11880;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1140
      TabIndex        =   6
      Top             =   2085
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   3
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1740
      TabIndex        =   7
      Top             =   2085
      Width           =   855
      VariousPropertyBits=   671107099
      MaxLength       =   6
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   2700
      TabIndex        =   8
      Top             =   2085
      Width           =   255
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   7
      Left            =   3060
      TabIndex        =   9
      Top             =   2085
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   1140
      TabIndex        =   10
      Top             =   2370
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "11880;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   1140
      TabIndex        =   11
      Top             =   2760
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "11880;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   11
      Left            =   1155
      TabIndex        =   13
      Top             =   3780
      Visible         =   0   'False
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "主類："
      Height          =   180
      Left            =   450
      TabIndex        =   32
      Top             =   405
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "次類："
      Height          =   180
      Left            =   450
      TabIndex        =   31
      Top             =   690
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "次次類："
      Height          =   180
      Left            =   210
      TabIndex        =   30
      Top             =   990
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "備用類："
      Height          =   180
      Left            =   330
      TabIndex        =   29
      Top             =   1275
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   180
      Left            =   90
      TabIndex        =   28
      Top             =   2130
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      Caption         =   "主旨："
      Height          =   180
      Left            =   450
      TabIndex        =   27
      Top             =   2415
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      Caption         =   "案例字號："
      Height          =   180
      Left            =   90
      TabIndex        =   26
      Top             =   2775
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      Caption         =   "案情摘要："
      Height          =   180
      Left            =   90
      TabIndex        =   25
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Create    ID  DATE ："
      Height          =   180
      Left            =   165
      TabIndex        =   24
      Top             =   4080
      Width           =   1620
   End
   Begin VB.Label Label14 
      Caption         =   "Update   ID  DATE ：  "
      Height          =   180
      Left            =   165
      TabIndex        =   23
      Top             =   4365
      Width           =   1620
   End
   Begin MSForms.Label Label15 
      Height          =   255
      Left            =   1845
      TabIndex        =   22
      Top             =   4080
      Width           =   4650
      VariousPropertyBits=   27
      Size            =   "8202;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   1845
      TabIndex        =   21
      Top             =   4380
      Width           =   4665
      VariousPropertyBits=   27
      Size            =   "8229;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   1305
      X2              =   3285
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "彙整旗標："
      Height          =   180
      Left            =   105
      TabIndex        =   20
      Top             =   3810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "（ 0:未彙整, 1:已彙整 ）"
      Height          =   180
      Left            =   1815
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書類型："
      Height          =   180
      Left            =   165
      TabIndex        =   18
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書日期："
      Height          =   180
      Left            =   165
      TabIndex        =   17
      Top             =   1815
      Width           =   900
   End
End
Attribute VB_Name = "frm090217_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text1,Label15,Label16)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim SYSERR As Integer

Public Function Process(ByVal strKey As String) As Boolean

On Error GoTo ErrFlg

   Dim pemain As New ADODB.Recordset, i As Integer
   
   If pemain.State = adStateOpen Then pemain.Close
   strExc(0) = "select pc01,pc02,pc03,pc04,pc05,pc06,pc07,pc08,pc09,pc10,pc11,s1.st02,pc13,pc14,s2.st02,pc16,pc17,PC01||PC02||PC03||PC04 AS A, PC18, PC19, PC20 from patentcase,staff s1,staff s2 where pc12=s1.st01(+) and pc15=s2.st01(+) and PC01||PC02||PC03||PC04='" & strKey & "'"
   pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If pemain.EOF And pemain.BOF Then
      MsgBox "資料庫內無資料", vbInformation
   Else
      If IsNull(pemain.Fields(0).Value) Then
          Combo1.Text = ""
      Else
          Combo1.Text = pemain.Fields(0).Value
      End If
      If IsNull(pemain.Fields(1).Value) Then
          Combo2.Text = ""
      Else
          Combo2.Text = pemain.Fields(1).Value
      End If
      If IsNull(pemain.Fields(2).Value) Then
          Combo3.Text = ""
      Else
          Combo3.Text = pemain.Fields(2).Value
      End If
      If IsNull(pemain.Fields(3).Value) Then
          Combo4.Text = ""
      Else
          Combo4.Text = pemain.Fields(3).Value
      End If
      For i = 4 To 10
         Text1(i) = "" & pemain.Fields(i).Value
      Next i
      Label15.Caption = CheckStr(pemain.Fields(11).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(12).Value))) & "      " & Format(CheckStr(pemain.Fields(13).Value), "@@:@@")
      Label16.Caption = CheckStr(pemain.Fields(14).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(15).Value))) & "      " & Format(CheckStr(pemain.Fields(16).Value), "@@:@@")
      Text1(11) = "" & pemain.Fields("PC18").Value
      cboPC19.ListIndex = -1
      If Not IsNull(pemain.Fields("PC19")) Then
         If Val(pemain.Fields("PC19")) >= 0 And Val(pemain.Fields("PC19")) <= 2 Then
            cboPC19.ListIndex = Val(pemain.Fields("PC19"))
         End If
      End If
      txtPC20.Text = ChangeWStringToTString(CheckStr("" & pemain.Fields("PC20").Value))
      Call locktext
      Process = True
   End If
   
ErrFlg:
   Set pemain = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
            
End Function

Private Sub Command_Click(Index As Integer)

   Dim iIdx As Integer, iRecAff As Long
   
On Error GoTo ErrFlg

   Select Case Index
   
      Case 0 '修改
         
         For iIdx = 8 To 10
            Text1_LostFocus iIdx
            If SYSERR = 1 Then Exit Sub
         Next
         
         Dim stPC19 As String, stPC20 As String, stPC07 As String, stPC08 As String
            
         If cboPC19.ListIndex >= 0 Then
            stPC19 = "'" & cboPC19.ListIndex & "'"
         Else
            stPC19 = "Null"
         End If
         
         If txtPC20 <> "" Then
            stPC20 = Val(txtPC20) + 19110000
         Else
            stPC20 = "NULL"
         End If
         
         stPC07 = IIf(Len(Trim(Text1(4))) <> 0, IIf(Len(Trim(Text1(5))) <> 0, IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text), ""), "")
         stPC08 = IIf(Len(Trim(Text1(4))) <> 0, IIf(Len(Trim(Text1(5))) <> 0, IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text), ""), "")
      
         cnnConnection.Execute "begin user_data.user_enabled:=1; UPDATE PATENTCASE  SET PC05='" & ChgSQL(Text1(4)) & "',PC06='" & ChgSQL(Text1(5)) & "',PC07='" & ChgSQL(stPC07) & "',PC08='" & ChgSQL(stPC08) & "',PC09='" & ChgSQL(Text1(8)) & "',PC10='" & ChgSQL(Text1(9)) & "',PC11='" & ChgSQL(Text1(10)) & "',PC18='" & ChgSQL(Text1(11)) & "',PC19=" & stPC19 & ",PC20=" & stPC20 & " WHERE  PC01='" & Combo1.Text & "' AND PC02='" & Combo2.Text & "' AND PC03='" & Combo3.Text & "' AND PC04='" & Combo4.Text & "'; end;", iRecAff
         frm090217_1.Show
         frm090217_1.Process
         Unload Me
         
      Case 1 '刪除
         If MsgBox("是否要刪除此筆資料", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            cnnConnection.Execute "delete PATENTCASE where  PC01='" & Combo1.Text & "' AND PC02='" & Combo2.Text & "' AND PC03='" & Combo3.Text & "' AND PC04='" & Combo4.Text & "'", iRecAff
            If iRecAff = 1 Then
               frm090217_1.Show
               frm090217_1.Process
               Unload Me
            Else
               MsgBox "無資料可更新，可能已被刪除，請重新查詢！", vbExclamation
            End If
         End If
          
      Case 2 '回前畫面
         frm090217_1.Show
         frm090217_1.CMDOK(3).SetFocus
         Unload Me
   End Select
   
ErrFlg:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Command(0).Default = True
   cboPC19.AddItem "判決", 0
   cboPC19.AddItem "決定書", 1
   cboPC19.AddItem "其他", 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090209_2 = Nothing
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index).Locked = True Then Exit Sub
   Select Case Index
      Case 11
         If Text1(Index) = "" Then Text1(Index) = "0"
   End Select
End Sub
'Add by Morgan 2004/5/3
Private Sub txtPC20_GotFocus()
   TextInverse txtPC20
End Sub
'Add by Morgan 2004/5/3
Private Sub txtPC20_Validate(Cancel As Boolean)
   If txtPC20.Locked = False Then
      If Len(Trim(txtPC20)) <> 0 Then
        If CheckIsTaiwanDate(txtPC20) = False Then
            txtPC20.SetFocus
            txtPC20_GotFocus
            Cancel = True
        ElseIf Val(strSrvDate(1)) < Val(txtPC20) + 19110000 Then
            MsgBox "文書日期不可大於系統日", vbExclamation, "USER 輸入錯誤！！"
            txtPC20.SetFocus
            txtPC20_GotFocus
            Cancel = True
        End If
      End If
   End If
End Sub

Private Sub locktext()   '鎖住輸入項
   Dim oText

   Combo1.Locked = True
   Combo2.Locked = True
   Combo3.Locked = True
   Combo4.Locked = True
   
   For Each oText In Text1
      oText.Locked = False
   Next
   cboPC19.Locked = False

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Select Case Index
      Case 8, 9, 10
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(Index).IMEMode = 1
         OpenIme
      Case Else
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(Index).IMEMode = 2
         CloseIme
   End Select
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index >= 4 And Index <= 7 Then
      KeyAscii = UpperCase(KeyAscii)
   ElseIf Index = 11 Then
      If KeyAscii <> 49 And KeyAscii <> 48 And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 8 '主旨
         SYSERR = 0
         If Me.Text1(Index).Text <> "" Then
            If CheckLengthIsOK(Me.Text1(Index).Text, 60) = False Then
               Me.Text1(Index).SetFocus
               SYSERR = 1
               Exit Sub
            End If
         End If
      Case 9 '案例字號
         SYSERR = 0
         If Me.Text1(Index).Text <> "" Then
            If CheckLengthIsOK(Me.Text1(Index).Text, 60) = False Then
               Me.Text1(Index).SetFocus
               SYSERR = 1
               Exit Sub
           End If
         End If
      Case 10 '案情摘要
         SYSERR = 0
         If Me.Text1(Index).Text <> "" Then
            If CheckLengthIsOK(Me.Text1(Index).Text, 400) = False Then
               Me.Text1(Index).SetFocus
               SYSERR = 1
               Exit Sub
            End If
         End If
   End Select
End Sub
