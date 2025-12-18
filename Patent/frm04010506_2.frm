VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010506_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "異議/舉發受理函輸入"
   ClientHeight    =   3540
   ClientLeft      =   -2208
   ClientTop       =   3552
   ClientWidth     =   8724
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8724
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   1650
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3150
      Width           =   684
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6540
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5712
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7764
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4920
      MaxLength       =   9
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1110
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2400
      Width           =   4305
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   6015
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10610;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Caption         =   "對造案件數代號:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3150
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   8520
      Y1              =   2256
      Y2              =   2256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   8520
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   5520
      TabIndex        =   26
      Top             =   2760
      Width           =   732
   End
   Begin VB.Label Label14 
      Caption         =   "列印客戶通知函:"
      Height          =   252
      Left            =   3480
      TabIndex        =   25
      Top             =   2760
      Width           =   1332
   End
   Begin VB.Label Label13 
      Caption         =   "(1.受理  2.對方延期)"
      Height          =   252
      Left            =   1680
      TabIndex        =   24
      Top             =   2760
      Width           =   1572
   End
   Begin VB.Label Label12 
      Caption         =   "程序:"
      Height          =   252
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   612
   End
   Begin VB.Label Label11 
      Caption         =   "機關文號:"
      Height          =   252
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "來函收文日:"
      Height          =   252
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label9 
      Height          =   252
      Left            =   1320
      TabIndex        =   20
      Top             =   1800
      Width           =   2052
   End
   Begin VB.Label Label6 
      Caption         =   "收文號:"
      Height          =   252
      Left            =   3480
      TabIndex        =   19
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   " "
      Height          =   252
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   2652
   End
   Begin VB.Label Label4 
      Caption         =   "案件性質:"
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label3 
      Height          =   252
      Left            =   1200
      TabIndex        =   16
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號:"
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label19 
      Caption         =   "對造號數:"
      Height          =   252
      Left            =   3480
      TabIndex        =   14
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   4440
      TabIndex        =   13
      Top             =   720
      Width           =   2652
   End
   Begin VB.Label Label7 
      Caption         =   "專利名稱:"
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   852
   End
End
Attribute VB_Name = "frm04010506_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim pmain As New ADODB.Recordset, PSUB As New ADODB.Recordset
Dim UserStaff
Dim m_CP09 As String
'Add By Cheng 2002/12/28
Dim m_PA09 As String '申請國家
'Added by Morgan 2014/1/14
Public m_DocWord As String 'Added by Morgan 2014/4/17
Public m_DocNo As String
Public m_AppNo As String
Dim m_NewCP09 As String
Dim m_PA26 As String
'end 2014/1/14
Dim m_PA75 As String 'Added by Morgan 2014/7/22
Dim stCP10 As String 'Add by Lydia 2014/11/18 改全域變數
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/17
Dim m_bolFMP As Boolean 'Added by Lydia 2022/10/11 是否為FMP案

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
       Case 0 '確定
         If Text5.Text = "" Then MsgBox "程序不可為空值", vbInformation: Text5.SetFocus: Exit Sub
         '檢查對造案件數代號
         If m_PA09 = "000" Then
            If Me.Text23.Text = "" Then
               MsgBox "請輸入對造案件數代號!!!", vbInformation
               Me.Text23.SetFocus
               Text23_GotFocus
               Exit Sub
            'Add by Morgan 2007/1/25 台灣對造案件數代號第一碼必須為 N -- 敏惠
            ElseIf Left(Me.Text23.Text, 1) <> "N" Then
               MsgBox "對造案件數代號第一碼必須為 N !!!", vbInformation
               Me.Text23.SetFocus
               Text23_GotFocus
               Exit Sub
            End If
         End If
        'Add By Cheng 2003/03/26
        '檢查機關文號
        If pmain.Fields(11).Value = 台灣國家代號 Then
            If Me.Text7.Tag = Me.Text7.Text Then
                MsgBox "請輸入機關文號!!!", vbExclamation + vbOKOnly
                Me.Text7.SetFocus
                Text7_GotFocus
                Exit Sub
            End If
        End If
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add By Sindy 2022/7/1
         If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
            If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
               Exit Sub
            End If
         End If
         '2022/7/1 END
         
         'Add By Cheng 2003/01/03
         Screen.MousePointer = vbHourglass
         'Modify By Cheng 2002/11/06
'         insertdata
         If insertdata = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         'Add By Cheng 2003/01/03
         Screen.MousePointer = vbDefault
         strKey1 = "1"
         If Text6.Text <> "N" Then
            'Modify By Cheng 2002/11/05
            '修正定稿別
'            NowPrint Label5.Caption, "9", "00", False, strUserNum, 0
            'Modify By Cheng 2002/12/28
'            NowPrint Label5.Caption, "09", "00", False, strUserNum, 0
            '若為受理
            If Me.Text5.Text = "1" Then
               If m_PA09 = "000" Then  '國內
                  NowPrint Label5.Caption, "09", "00", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
               Else        '大陸
                  NowPrint Label5.Caption, "09", "02", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
               End If
            '若為對方延期
            Else
                NowPrint Label5.Caption, "09", "01", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
            End If
         End If
       'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If m_PA09 = "000" And Text1 = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (m_PA09 = "000" Or m_PA09 = "020") And Text1 = "P" Then
         If (m_PA09 = "000" Or m_PA09 = "020") And Text1 = "P" And m_bolFMP = False Then
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg Text1, Text2, Text3, Text4, stCP10, m_NewCP09
            PUB_TaiwanCInputMsg Text1, Text2, Text3, Text4, stCP10, m_PA09, m_NewCP09
         End If
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010506_1
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         ElseIf Me.m_DocNo <> "" Then
         'Added by Morgan 2014/1/14
         'If Me.m_DocNo <> "" Then
         '2016/10/5 END
            Unload frm04010506_1
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
         
           'Modify By Cheng 2003/01/03
   '         frm04010506_1.Text7 = ""
            frm04010506_1.Show
           'Add By Cheng 2003/01/03
            frm04010506_1.Clear
            Unload Me
         End If 'Added by Morgan 2014/1/14
         
       Case 1
         frm04010506_1.Show
         Unload Me
       Case 2
         Unload frm04010506_1
         Unload Me
End Select
End Sub

Public Sub SetData(ByVal strData As String)
   m_CP09 = strData
End Sub

Private Sub Form_Activate()
   If pmain.State = adStateOpen Then pmain.Close
   strExc(0) = "SELECT ST01 FROM STAFF WHERE ST02='" & strUserName & "'"
   pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   UserStaff = pmain.Fields(0).Value
   If pmain.State = adStateOpen Then pmain.Close
   strExc(0) = "SELECT CP01,CP02,CP03,CP04,PA05,PA06,PA07," & _
      "DECODE(PA09,'000',CPM03,CPM04),CP36,CP09,CP08,PA09,CP13,CP12,PA26,PA75 FROM " & _
      "CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP09='" & m_CP09 & "' AND " & _
      "CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND " & _
      "CP01=CPM01 AND CP10=CPM02"
   pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   Text1.Text = pmain.Fields(0).Value
   Text2.Text = pmain.Fields(1).Value
   Text3.Text = pmain.Fields(2).Value
   Text4.Text = pmain.Fields(3).Value
   If IsNull(pmain.Fields(4).Value) Then
      Combo1.AddItem "中: ", 0
      Combo1.Text = "中: "
   Else
      Combo1.AddItem "中: " + pmain.Fields(4).Value, 0
      Combo1.Text = "中: " + pmain.Fields(4).Value
   End If
   
   If IsNull(pmain.Fields(5).Value) Then
      Combo1.AddItem "英: ", 1
   Else
      Combo1.AddItem "英: " + pmain.Fields(5).Value, 1
   End If
   If IsNull(pmain.Fields(6).Value) Then
      Combo1.AddItem "日: ", 2
   Else
      Combo1.AddItem "日: " + pmain.Fields(6).Value, 2
   End If
   If IsNull(pmain.Fields(8).Value) Then
      Label2.Caption = ""
   Else
        'Modify by Morgan 2004/2/3
        '若CP36後三碼為NXX或PXX時不顯示後三碼
        'Label2.Caption = pmain.Fields(8).Value
        If InStr(1, "NP", Left(Right(pmain.Fields(8).Value, 3), 1)) > 0 Then
            Label2.Caption = Left(pmain.Fields(8).Value, Len(pmain.Fields(8).Value) - 3)
            Text23 = Right(pmain.Fields(8).Value, 3)
        Else
            Label2.Caption = pmain.Fields(8).Value
        End If
        'Modify End 2004/2/3
   End If
   Label5.Caption = pmain.Fields(9).Value
   Label3.Caption = pmain.Fields(7)
   Label9.Caption = frm04010506_1.Text5

   ' 90.10.5 modify by sonia (機關文號設預設內容)
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
'   Text7.Text = "（" & strTmp & "）智專三(三）字第號"
   If pmain.Fields(11).Value = "000" Then
      Text7.Text = "（" & strTmp & "）智專一（二）字第號"
        'Add By Cheng 2003/03/26
        '記錄機關文號的預設值
        Me.Text7.Tag = Me.Text7.Text
   End If
   
   'Added by Morgan 2014/1/14
   'Modified by Morgan 2014/4/17 +發文字
   If m_DocWord <> "" Then
      Text7 = m_DocWord & "字第" & m_DocNo & "號"
   ElseIf m_DocNo <> "" Then
      Text7 = Replace(Text7, "第號", "第" & m_DocNo & "號")
   End If
   m_PA26 = "" & pmain.Fields("pa26")
   'end 2014/1/14
   m_PA75 = "" & pmain.Fields("pa75") 'Added by Morgan 2014/7/22
   
   If m_AppNo <> "" Then Text23 = Mid(m_AppNo, 10) 'Added by Morgan 2014/6/4
   
    'Add By Cheng 2002/12/28
    '取得申請國家
    m_PA09 = "" & pmain.Fields(11).Value
    
   'Added by Lydia 2022/10/11
   If Left("" & pmain.Fields("CP12"), 1) = "F" And m_PA09 <> "000" Then
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010506_1.m_strIR01
   m_strIR02 = frm04010506_1.m_strIR02
   m_strIR03 = frm04010506_1.m_strIR03
   m_strIR04 = frm04010506_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010506_2 = Nothing
End Sub

Private Sub Text23_GotFocus()
    'Add By Cheng 2003/01/27
    '欄位值反白
    TextInverse Me.Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/01/27
    '轉換為大寫
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
If Text5.Text <> "1" And Text5.Text <> "2" Then
   MsgBox "程序只可輸入 1 或 2", vbInformation
   Text5.SetFocus
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
   Cancel = True
Else
   Cancel = False
End If
Exit Sub
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Text6.Text <> "" And Text6.Text <> "N" Then
   MsgBox "只可輸入空白或 N", vbInformation
   Text6.SetFocus
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6)
   Cancel = True
Else
   Cancel = False
End If
Exit Sub
End Sub

Private Sub Text7_GotFocus()
'Text7.SelStart = 0
'Text7.SelLength = Len(Text7)
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'將游標設定在機關文號欄的"專"的後面
With Me.Text7
   If Len("" & .Text) > 0 Then
        'Modify By Cheng 2002/10/28
'      intPos = InStr("" & .Text, "專")
      intPos = InStr("" & .Text, "字")
'      If intPos > 0 Then
      If intPos - 1 > 0 Then
         .SelStart = intPos - 1
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If pmain.Fields(11).Value = "000" And Text7.Text = "" Then
   MsgBox "申請國家為台灣時機關文號不可為空值", vbInformation
   Text7.SetFocus
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
   Cancel = True
   Exit Sub
Else
   Cancel = False
End If
'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
'If CheckLengthIsOK(Text7, 40) = False Then
If CheckLengthIsOK(Text7, Text7.MaxLength) = False Then
   Text7.SetFocus
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
   Cancel = True
Else
   Cancel = False
End If
End Sub

Private Function GetName(county As String, SYS As String, NUM As String)
Dim gtname As New ADODB.Recordset
'edit by nickc 2007/02/08
'strExc(0) = "select decode( '" & conuty & "','000',cpm03,cpm04) from casepropertymap where cpm01='" & SYS & "' and cpm02='" & NUM & "'"
strExc(0) = "select decode( '" & county & "','000',cpm03,cpm04) from casepropertymap where cpm01='" & SYS & "' and cpm02='" & NUM & "'"

gtname.Open strExc(0), cnnConnection
If gtname.BOF And gtname.EOF Then GetName = ""
GetName = gtname.Fields(0).Value
End Function

'Modify By Cheng 2002/11/06
'Private Sub insertdata()
Private Function insertdata() As Boolean
 Dim autonum As String
 Dim pro As String
  'Add by Morgan 2004/2/9
 Dim stCP12 As String, stCP13 As String
 'Dim stCP10 As String 'Add by Lydia 2014/11/18 改全域變數
 
 'Add By Cheng 2002/11/06
 On Error GoTo ErrorHandler
 insertdata = True
 cnnConnection.BeginTrans
    'Add By Cheng 2003/01/27
    '更新原進度檔資料的對造號數
    'Modify by Morgan 2004/3/31
    'strSQL = "Update CaseProgress Set CP36=CP36||'" & Me.Text23.Text & "' Where CP09='" & m_CP09 & "' "
    strSql = "Update CaseProgress Set CP36='" & Me.Label2.Caption & Me.Text23.Text & "' Where CP09='" & m_CP09 & "' "
    'Add By Cheng 2003/02/06
    '執行更新動作
    cnnConnection.Execute strSql
    autonum = AutoNo("C", 6)
    
'Removed by Morgan 2014/4/16 沒用
'    If Text5.Text = "1" Then
'       pro = "1607"
'    ElseIf Text5.Text = "2" Then
'       pro = "1611"
'    End If
'end 2014/4/16

    'Modify By Cheng 2003/01/27
    '更新對造號數
'   strExc(0) = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
'      "CP12,CP13,CP14,CP20,cp26,CP27,CP32,CP43) VALUES ('" & Text1.Text & "','" & _
'      Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & _
'      ChangeTStringToWString(frm04010506_1.Text5.Text) & "','" & Text7.Text & "','" & _
'      autonum & "',DECODE('" & Text5 & "','1','1803','2','1804'),'" & _
'      pmain.Fields(13).Value & "','" & pmain.Fields(12).Value & "','" & UserStaff & _
'      "','N','N','" & GetTodayDate & "' ,'N','" & Label5.Caption & "') "
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    
    'Modify by Morgan 2004/2/9
   ' strExc(0) = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP12,CP13,CP14,CP20,cp26,CP27,CP32,CP36,CP43) VALUES ('" & Text1.Text & "','" & _
      Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & _
      ChangeTStringToWString(frm04010506_1.Text5.Text) & "','" & Text7.Text & "','" & _
      autonum & "',DECODE('" & Text5 & "','1','1803','2','1804'),'" & _
      pmain.Fields(13).Value & "','" & PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text) & "','" & UserStaff & _
      "','N','N','" & GetTodayDate & "' ,'N','" & Me.Label2.Caption & Me.Text23.Text & "','" & Label5.Caption & "') "
   
    stCP10 = IIf(Text5 = "1", "1803", "1804")
    stCP13 = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
    stCP12 = GetSalesArea(stCP13)
    'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
    'Modified by Morgan 2020/1/17 +m_bolNoCP27
    strExc(0) = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP12,CP13,CP14,CP20,cp26,CP27,CP32,CP36,CP43,CP119) VALUES ('" & Text1.Text & "','" & _
      Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & _
      ChangeTStringToWString(frm04010506_1.Text5.Text) & "','" & Text7.Text & "','" & _
      autonum & "','" & stCP10 & "','" & _
      stCP12 & "','" & stCP13 & "','" & UserStaff & _
      "','N','N'," & IIf(m_bolNoCP27, "NULL", strSrvDate(1)) & " ,'N','" & Me.Label2.Caption & Me.Text23.Text & "'" & _
      ",'" & Label5.Caption & "'," & DBDATE(Label9) & ") "
   
    'Modify end 2004/2/9
   cnnConnection.Execute strExc(0)
   
   'Added by Morgan 2014/1/14
   m_NewCP09 = autonum
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, m_NewCP09, Text1, Text2, Text3, Text4, stCP10
   End If
   'end 2014/1/14
   
   'Added by Morgan 2014/4/14 電子化-新增信函進度檔
   If m_PA09 = "000" Then
      strExc(1) = ""
      If Text6 <> "N" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(Text1, stCP10, , , Text1, Text2, Text3, Text4)
         strExc(1) = PUB_GetLetterJudgeNew("1", Text1, stCP10)
      End If
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      PUB_AddLetterProgress m_NewCP09, 1, IIf(Text6 <> "N", True, False), strExc(1), False, m_PA26, stCP10, m_PA75
      
   'Added by Morgan 2016/6/16 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      strExc(1) = ""
      If Text6 <> "N" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(Text1, stCP10, , m_PA09, Text1, Text2, Text3, Text4)
         strExc(1) = PUB_GetLetterJudgeNew("1", Text1, stCP10, m_PA09, , , IIf(Left(stCP12, 1) = "F", True, False))
      End If
      PUB_AddLetterProgress m_NewCP09, 2, IIf(Text6 <> "N", True, False), strExc(1), False, m_PA26, stCP10, m_PA75
   'end 2016/6/16
   
   End If
   'end 2014/4/14
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010506_1", IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
   End If
   '2016/10/5 END
   
    'Add By Cheng 2002/11/06
    cnnConnection.CommitTrans
    Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    insertdata = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
End If

   'Added by Morgan 2014/5/15 電子化-檢查pdf檔
   If m_PA09 = "000" Then
      If PUB_CheckPDF(Text1, Text2, Text3, Text4, 1, m_DocNo) = False Then
         Exit Function
      End If
   End If
   'end 2014/5/15
   
   'Added by Morgan 2020/1/17
   '大陸案,有通知函,程序承辦,非掛號(無期限)
   m_bolNoCP27 = False
   'Removed by Morgan 2024/1/30 取消--郭
   'If m_PA09 = "020" And Text6 <> "N" Then
   '   If PUB_GetCustomerValue(m_PA26, "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/17
   
TxtValidate = True
End Function

