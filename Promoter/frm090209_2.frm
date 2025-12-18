VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090209_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報簡訊資料彙整作業"
   ClientHeight    =   4065
   ClientLeft      =   1830
   ClientTop       =   1950
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7260
   Begin VB.CommandButton Command 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton Command 
      Caption         =   "刪除(&D)"
      Height          =   405
      Index           =   1
      Left            =   5220
      TabIndex        =   8
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton Command 
      Caption         =   "修改(M)"
      Height          =   405
      Index           =   0
      Left            =   4350
      TabIndex        =   7
      Top             =   70
      Width           =   840
   End
   Begin MSForms.TextBox Text 
      Height          =   1545
      Index           =   6
      Left            =   1050
      TabIndex        =   6
      Top             =   2070
      Width           =   6135
      VariousPropertyBits=   -1467989989
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "10821;2725"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   5
      Left            =   1050
      TabIndex        =   5
      Top             =   1695
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   4
      Left            =   1050
      TabIndex        =   4
      Top             =   1365
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   1050
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   1050
      TabIndex        =   1
      Top             =   1050
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   465
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   5
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "索引只可輸入P或U開頭"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   5310
      TabIndex        =   20
      Top             =   1410
      Width           =   1920
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   180
      Left            =   1470
      TabIndex        =   19
      Top             =   1410
      Width           =   3720
   End
   Begin VB.Label Label9 
      Height          =   180
      Left            =   1050
      TabIndex        =   18
      Top             =   780
      Width           =   1215
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1050
      TabIndex        =   17
      Top             =   3750
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label8"
      Size            =   "5794;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "Create  ID ："
      Height          =   180
      Left            =   -240
      TabIndex        =   16
      Top             =   3750
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容摘要："
      Height          =   180
      Left            =   -240
      TabIndex        =   15
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告日期："
      Height          =   180
      Left            =   -240
      TabIndex        =   14
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "索引："
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   1365
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "國際分類："
      Height          =   180
      Left            =   0
      TabIndex        =   12
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告號數："
      Height          =   180
      Left            =   -240
      TabIndex        =   11
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告頁數："
      Height          =   180
      Left            =   -240
      TabIndex        =   10
      Top             =   510
      Width           =   1215
   End
End
Attribute VB_Name = "frm090209_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text,Label8)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, PESUB As New ADODB.Recordset
Dim i As Integer, UserStaff As String, s As Integer

Private Sub Command_Click(Index As Integer)
'Add By Cheng 2003/03/03
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnBegin As Boolean '判斷Transaction
Dim Cancel As Boolean

On Error GoTo ErrorHandler
blnBegin = False
Select Case Index
Case 0 '修改
      '2010/5/21 ADD BY SONIA
      Cancel = False
      Text_Validate 6, Cancel
      If Cancel = True Then
         Text(6).SetFocus
         Text_GotFocus 6
         Exit Sub
      End If
      '2010/5/21 END
      
    'Modify By Cheng 2003/03/03
    '若未修改公告頁數
    If Me.Text(0).Text = Me.Text(0).Tag Then
        cnnConnection.Execute "UPDATE BULLETINBRIEF SET BB03='" & Text(1).Text & "',BB04='" & Text(2).Text & "',BB05='" & Text(3).Text & "',BB06='" & Text(4).Text & "',BB07=" & ChangeTStringToWString(Text(5).Text) & ",BB08='" & Text(6).Text & "' WHERE BB02='" & Label9.Caption & "' and BB01='" & Text(0).Text & "' "
    '若有修改公告頁數
    Else
        StrSQLa = "Select * From BulletinBrief Where BB01='" & Me.Text(0).Tag & "' And BB02='" & Me.Label9.Caption & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        '若有搜尋到欲更新的資料
        If rsA.RecordCount > 0 Then
            cnnConnection.BeginTrans
            blnBegin = True
            'Modify by Morgan 2004/5/25
            '其他欄位也要一起更新
            'strSQLA = "Insert Into BulletinBrief Values ('" & Me.Text(0).Text & "','" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "'," & CNULL(rsA.Fields(6).Value) & ",'" & ChgSQL(rsA.Fields(7).Value) & "','" & rsA.Fields(8).Value & "','" & rsA.Fields(9).Value & "'," & CNULL(rsA.Fields(10).Value) & "," & CNULL(rsA.Fields(11).Value) & " ) "
            StrSQLa = "Insert Into BulletinBrief(BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08,BB09,BB10,BB11,BB12) Values ('" & Me.Text(0).Text & "','" & rsA.Fields(1).Value & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "','" & Text(4).Text & "'," & CNULL(ChangeTStringToWString(Text(5).Text)) & ",'" & ChgSQL(Text(6).Text) & "','" & rsA.Fields(8).Value & "','" & rsA.Fields(9).Value & "'," & CNULL(rsA.Fields(10).Value) & "," & CNULL(rsA.Fields(11).Value) & " ) "
            '刪除原資料
            cnnConnection.Execute "DELETE BULLETINBRIEF WHERE BB02='" & Label9.Caption & "' and BB01='" & Me.Text(0).Tag & "' "
            '新增新資料
            cnnConnection.Execute StrSQLa
            cnnConnection.CommitTrans
            blnBegin = False
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    frm090209_1.Show
    frm090209_1.Process
    Unload Me
Case 1 '刪除
    cnnConnection.Execute "DELETE BULLETINBRIEF WHERE BB02='" & Label9.Caption & "' and BB01='" & Text(0).Text & "' "
    frm090209_1.Show
    frm090209_1.Process
    Unload Me
Case 2
    frm090209_1.Show
    Unload Me
End Select
'Add By Cheng 2003/03/03
Exit Sub
ErrorHandler:
    If blnBegin = True Then cnnConnection.RollbackTrans
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Activate()
'If pemain.State = adStateOpen Then pemain.Close
'strExc(0) = "SELECT ST01 FROM STAFF WHERE ST02='" & strUserName & "'"
'pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'If pemain.BOF And pemain.EOF Then MsgBox "無此LOGIN人員之資料", vbInformation: Unload Me
'UserStaff = pemain.Fields(0).Value
'pemain.Close
'strExc(0) = "SELECT BB01,BB03,BB04,BB05,BB06,BB07,BB08,BB02 FROM BULLETINBRIEF WHERE BB02='" & frm090209_1.NUM & "'"
'pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'For i = 0 To 6
'  If i = 5 Then
'        Text(i).Text = ChangeWStringToTString(pemain.Fields(i).Value)
'  Else
'        Text(i).Text = pemain.Fields(i).Value
'  End If
'Next i
'  Label9.Caption = pemain.Fields(7).Value
'  Label8.Caption = UserStaff
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Command(0).Default = True
'pemain.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/03/04
frm090209_1.cmdOK(3).SetFocus
Set frm090209_2 = Nothing
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Text(Index).SelStart = 0
    Text(Index).SelLength = Len(Text(Index))
End Sub

Sub Process(Strindex As String, StrIndex2 As String)
'Add By Cheng 2003/03/03
'記錄公告頁數
Me.Text(0).Tag = "" & Strindex
strExc(0) = "SELECT BB01,BB03,BB04,BB05,BB06,BB07,BB08,bb02,st02,bb11,bb12 FROM BULLETINBRIEF,staff WHERE BB01='" & Strindex & "' AND BB02='" & StrIndex2 & "' and bb10=st01(+) "
If pemain.State = adStateOpen Then pemain.Close
pemain.CursorLocation = adUseClient
pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If pemain.RecordCount <> 0 Then
For i = 0 To 6
  If i = 5 Then
        Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i).Value))
  Else
        Text(i).Text = CheckStr(pemain.Fields(i).Value)
  End If
Next i
If PESUB.State = adStateOpen Then PESUB.Close
strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(4).Text & "'"
PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If Not PESUB.BOF And Not PESUB.EOF Then
   Label10.Caption = PESUB.Fields(0).Value
Else
   Label10.Caption = ""
End If
PESUB.Close


Label9.Caption = CheckStr(pemain.Fields(7).Value)
Label8.Caption = CheckStr(pemain.Fields(8).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(9).Value))) & "      " & Format(CheckStr(pemain.Fields(10).Value), "@@:@@")
End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
Case 1, 2, 3, 4
 KeyAscii = UpperCase(KeyAscii)
Case Else
End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
   Case 4
      If PESUB.State = adStateOpen Then PESUB.Close
        strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(4).Text & "'"
        PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
        If PESUB.BOF And PESUB.EOF Then
            MsgBox "無此索引代號", vbInformation
            Label10.Caption = ""
            Text(Index).SetFocus
            Text(Index).SelStart = 0
            Text(Index).SelLength = Len(Text(Index))
            Exit Sub
        End If
        If Not PESUB.BOF Then PESUB.MoveFirst
        Label10.Caption = PESUB.Fields(0).Value
        PESUB.Close
   Case 5
         If Len(Trim(Text(Index))) <> 0 Then
           If CheckIsTaiwanDate(Text(Index)) = False Then
               Text(Index).SetFocus
               Text(Index).SelStart = 0
               Text(Index).SelLength = Len(Text(Index))
               Exit Sub
           Else
               If Val(strSrvDate(1)) < Val(Text(Index)) + 19110000 Then
                  s = MsgBox("公告日期不可大於系統日", , "USER 輸入錯誤！！")
                  Text(Index).SetFocus
                  Text_GotFocus (Index)
                  Exit Sub
               End If
           End If
        End If

   Case Else
   End Select
End Sub

'2010/5/21 ADD BY SONIA 自Text_LostFocus移過來,回前畫面時不檢查
Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 6
         If Len(Trim(Text(Index))) <> 0 Then
           If CheckLengthIsOK(Text(Index), 300) = False Then
               Text(Index).SetFocus
               Text(Index).SelStart = 0
               Text(Index).SelLength = Len(Text(Index))
               Exit Sub
           End If
         End If
      Case Else
   End Select
End Sub
'2010/5/21 END
