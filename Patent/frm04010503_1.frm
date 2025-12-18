VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010503_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁函輸入"
   ClientHeight    =   5745
   ClientLeft      =   -2790
   ClientTop       =   1710
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1560
      Width           =   1632
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   11
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7524
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   180
      TabIndex        =   12
      Top             =   600
      Width           =   9012
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   364
         Width           =   1632
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   372
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   3270
         TabIndex        =   2
         Top             =   372
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4350
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "P"
         Top             =   364
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4845
         MaxLength       =   6
         TabIndex        =   4
         Top             =   364
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5700
         MaxLength       =   1
         TabIndex        =   5
         Top             =   364
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   5970
         MaxLength       =   2
         TabIndex        =   6
         Top             =   364
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6540
         TabIndex        =   7
         Top             =   312
         Width           =   800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3672
      Left            =   180
      TabIndex        =   9
      Top             =   1920
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      FocusRect       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   948
   End
End
Attribute VB_Name = "frm04010503_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intLastRow As Integer
Dim intWhere As Integer
Dim m_PA09 As String 'Added by Morgan 2012/4/20
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
'end 2014/1/14
'Added by Morgan 2014/4/17
Public m_DocWord As String
Public m_DeadLine As String
'end 2014/4/17
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
'2016/10/5 END


Public Sub Clear()
   Text7 = Empty
   InitGrid 9, MSHFlexGrid1
   GridHead
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   Dim strTmp(1 To 2) As String
   'Add By Cheng 2002/06/21
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
 
   m_PA09 = "" 'Added by Morgan 2012/4/20
   
   intI = 0
   'Modify By Cheng 2002/06/21
   '選擇申請案號
   If Me.Option1(0).Value Then
      If Text7 = "" Then
         MsgBox "申請案號不得空白，請重新輸入 !", vbCritical
         Me.Text7.SetFocus
         Text7_GotFocus
         Exit Sub
      End If
      'strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
      '   "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent where PA01='P' AND " & _
      '   "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "' union " & _
      '   "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
      '   "nvl(cp37,nvl(cp38,cp38)) as Name," & _
      '   "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' from caseprogress where " & _
      '   "CP01='P' AND cp36='" & Text7 & "' and (cp01,cp02,cp03,cp04) not in " & _
      '   "(select pa01,pa02,pa03,pa04 from patent where PA01='P' AND pa11='" & Text7 & "' and " & _
      '   "pa09='" & 台灣國家代號 & "')"
      'Modified by Morgan 2012/4/19 ''-->pa09
    '  strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,pa09 cty from patent where PA01='P' AND " & _
         "pa11='" & Text7 & "' union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
         "nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 from caseprogress,patent where " & _
         "CP01='P' AND cp36='" & Text7 & "' and (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='P' AND pa11='" & Text7 & "' and " & _
         "pa09='" & 台灣國家代號 & "') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
   
      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
      '設別名f0,+FMP2openSQL
       strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,pa09 cty from patent f0 where PA01='P' AND " & _
         "pa11='" & Text7 & "' " & FMP2openSQL
       strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
        'Add by Lydia 2014/10/31 先判斷外專程序人員權限。
        If FMP2open = True And FMP2openSQL <> "" Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
            Me.Text7.SetFocus
            Text7_GotFocus
            Exit Sub
           End If
        End If
        '抓對造資料 '測試大陸案(P-105534)以申請案號查詢,會比本所案號查詢多一筆(Union的),其意義目前未知,將來可考慮省略Union部份
       strExc(0) = strExc(0) & " union select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
         "nvl(cp37,nvl(cp38,cp38)) as Name,nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 " & _
         "from caseprogress f2,patent where CP01='P' AND cp36='" & Text7 & "' " & FMP2openSQL & _
         " and (cp01,cp02,cp03,cp04) not in (select pa01,pa02,pa03,pa04 from patent where PA01='P' AND pa11='" & Text7 & "' and " & _
         "pa09='" & 台灣國家代號 & "') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
       strExc(0) = Replace(strExc(0), "f0.CP", "f2.CP")
   '選擇本所案號
   Else
      If Me.Text1.Text = "" Then
         MsgBox "系統類別不得空白，請重新輸入 !", vbCritical
         Me.Text1.SetFocus
         Text1_GotFocus
         Exit Sub
      Else
         If Me.Text1.Text <> "P" And Me.Text1.Text <> "PS" Then
            MsgBox "系統類別輸入錯誤，請重新輸入 !", vbCritical
            Me.Text1.SetFocus
            Text1_GotFocus
            Exit Sub
         End If
      End If
      If Me.Text2.Text = "" Then
         MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
         Me.Text2.SetFocus
         Text2_GotFocus
         Exit Sub
      End If
      If Me.Text1.Text = "P" Then
         '先檢查申請國家是否為台灣
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         StrSQLa = "SELECT * FROM PATENT WHERE PA01='" & Me.Text1.Text & "' AND PA02='" & Me.Text2.Text & "' AND PA03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' AND PA04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA("PA09") = 台灣國家代號 Then
               MsgBox "本案申請國家為台灣, 請改以申請案號查詢!!!", vbExclamation + vbOKOnly
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Sub
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         
         '專利基本檔
         'Modified by Morgan 2012/4/19 ''-->pa09
'         strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
            "'' as RName,'',pa01,pa02,pa03,pa04,pa09 cty from patent where PA01='" & Me.Text1.Text & "' AND pa02='" & Me.Text2.Text & "' and pa03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and pa04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' " & _
            " union " & _
            "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
            "nvl(cp37,nvl(cp38,cp38)) as Name," & _
            "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 from caseprogress,patent where " & _
            "CP01='" & Me.Text1.Text & "' AND cp02='" & Me.Text2.Text & "' and cp03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and cp04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' and (cp01,cp02,cp03,cp04) not in " & _
            "(select pa01,pa02,pa03,pa04 from patent where PA01='" & Me.Text1.Text & "' AND pa02='" & Me.Text2.Text & "' and pa03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and pa04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "'" & _
            ")  and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
            
      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
      '設別名f0,+FMP2openSQL
          strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
            "'' as RName,'',pa01,pa02,pa03,pa04,pa09 cty from patent f0 where PA01='" & Me.Text1.Text & "' AND pa02='" & Me.Text2.Text & "' " & _
            "and pa03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and pa04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' " & FMP2openSQL
        strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
        'Add by Lydia 2014/10/31 先判斷外專程序人員權限。
        If FMP2open = True And FMP2openSQL <> "" Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
            Me.Text2.SetFocus
            Text2_GotFocus
            Exit Sub
           End If
        End If
        '抓對造資料
          strExc(0) = strExc(0) & " union select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
            "nvl(cp37,nvl(cp38,cp38)) as Name,nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 " & _
            "from caseprogress f2,patent where CP01='" & Me.Text1.Text & "' AND cp02='" & Me.Text2.Text & "' " & _
            "and cp03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and cp04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' " & FMP2openSQL & _
            "and (cp01,cp02,cp03,cp04) not in (select pa01,pa02,pa03,pa04 from patent where PA01='" & Me.Text1.Text & "' AND pa02='" & Me.Text2.Text & "' and pa03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and pa04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "'" & _
            ")  and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
        strExc(0) = Replace(strExc(0), "f0.CP", "f2.CP")
        
      Else
         '先檢查申請國家是否為台灣
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         StrSQLa = "SELECT * FROM SERVICEPRACTICE WHERE SP01='" & Me.Text1.Text & "' AND SP02='" & Me.Text2.Text & "' AND SP03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' AND SP04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA("SP09") = 台灣國家代號 Then
               MsgBox "本案申請國家為台灣, 請改以申請案號查詢!!!", vbExclamation + vbOKOnly
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Sub
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         
         '服務業務基本檔
         'Modified by Morgan 2012/4/19 ''--sp09
         'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
         '設別名f0,+FMP2openSQL
         strExc(0) = "select " & ChgService("", 1) & " as No,nvl(sp05,nvl(sp06,sp07)) as Name," & _
            "'' as RName,'',sp01,sp02,sp03,sp04,sp09 cty from ServicePractice f0 where SP01='" & Me.Text1.Text & "' AND Sp02='" & Me.Text2.Text & "' and " & _
            "Sp03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and Sp04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' " & FMP2openSQL
        strExc(0) = Replace(strExc(0), "f0.CP", "f0.SP")
        'Add by Lydia 2014/10/31 先判斷外專程序人員權限。
        If FMP2open = True And FMP2openSQL <> "" Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
            Me.Text2.SetFocus
            Text2_GotFocus
            Exit Sub
           End If
        End If
         strExc(0) = strExc(0) & " union select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
            "nvl(cp37,nvl(cp38,cp38)) as Name,nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' " & _
            "from caseprogress f2 where CP01='" & Me.Text1.Text & "' AND cp02='" & Me.Text2.Text & "' " & _
            "and cp03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and cp04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "' " & _
            FMP2openSQL & "and (cp01,cp02,cp03,cp04) not in (select Sp01,Sp02,Sp03,Sp04 from ServicePractice where SP01='" & Me.Text1.Text & "' AND Sp02='" & Me.Text2.Text & "' " & _
            "and Sp03='" & IIf(Me.Text3.Text = "", "0", Me.Text3.Text) & "' and Sp04='" & IIf(Me.Text4.Text = "", "00", Me.Text4.Text) & "'" & _
            ")"
           strExc(0) = Replace(strExc(0), "f0.CP", "f2.CP")
      End If
   End If
 
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))

   If intI = 1 Then m_PA09 = "" & RsTemp("cty") 'Added by Morgan 2012/4/20
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   ' 90.06.29 modify by louis 只有一筆時直接進入到下一個畫面
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      MSHFlexGrid1_Click
      'If m_PA09 = 台灣國家代號 Then 'Added by Morgan 2012/4/20
      'Add By Sindy 2017/10/16 + Or (Text5.Tag = m_RDate And m_RDate <> "")
      If m_PA09 = 台灣國家代號 Or (Text5.Tag = m_RDate And m_RDate <> "") Then 'Added by Morgan 2012/4/20
         FormConfirm
      End If 'Added by Morgan 2012/4/20
   End If
   
   'Added by Morgan 2012/4/20
   '非台灣要清除來函收文日重新輸入
   If m_PA09 <> "" And m_PA09 <> 台灣國家代號 Then
      'Add By Sindy 2017/10/16
      If Not (Text5.Tag = m_RDate And m_RDate <> "") Then
      '2017/10/16 END
         Text5 = ""
         Text5.SetFocus
      End If
   End If
   'end 2012/4/20
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2017/12/27
   If m_strIR01 <> "" And m_Done = False Then
      Option1(0).Value = True
      'Text7.Text = m_AppNo
      Text5.Text = m_RDate
      Text5.Tag = m_RDate 'Add By Sindy 2017/10/16
      'Command1.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   'Added by Morgan 2014/1/14
   ElseIf m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text7.Text = m_AppNo
      Text5.Text = m_RDate
      Text5.Tag = m_RDate 'Add By Sindy 2017/10/16
      Command1.Value = True
      m_Done = True
   End If
   'end 2014/1/14
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   InitGrid 9, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   'Add By Cheng 2002/06/21
   SendKeys "{Tab}"
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010503_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 8
   cmdOK(0).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 '申請案號
   Me.Text7.Enabled = True
   Me.Text2.Enabled = False
   Me.Text3.Enabled = False
   Me.Text4.Enabled = False
   Me.Text7.SetFocus
Case 1 '本所案號
   Me.Text7.Enabled = False
   Me.Text2.Enabled = True
   Me.Text3.Enabled = True
   Me.Text4.Enabled = True
   Me.Text2.SetFocus
End Select
End Sub

Private Sub Text1_GotFocus()
TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 80 And KeyAscii <> 83 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Me.Text1.Text <> "P" And Me.Text1.Text <> "PS" Then
   MsgBox "系統類別只能輸入 P 或 PS !!!", vbExclamation + vbOKOnly
   Cancel = True
   Me.Text1.SetFocus
   Text1_GotFocus
End If
End Sub

Private Sub Text2_GotFocus()
TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   cmdOK(0).Default = True 'Added by Morgan 2020/9/15
End Sub

Private Sub Text5_LostFocus()
   Command1.Default = True 'Added by Morgan 2020/9/15
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         Text5 = TransDate(Text5, 1) 'Add by Morgan 2009/7/31 改可輸西元年但自動轉民國年
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If Text5 = "" Then
      MsgBox "來函收文日不可空白 !", vbCritical
      Text5.SetFocus
      Exit Function
      
   'Add by Morgan 2009/7/31
   Else
      Text5_Validate Cancel
      If Cancel = True Then
         Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
      
   End If
   TxtValidate = True
   
End Function

' 確認鈕
Private Sub FormConfirm()

   Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
 
   If TxtValidate = False Then Exit Sub
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 8) = "v" Then
            bolChk = True
            For j = 1 To 4
               strExc(j) = .TextMatrix(i, j + 3)
            Next
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Add By Sindy 2017/12/27
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2017/12/27 END
   'Add By Sindy 2016/10/5
   frm04010503_2.m_strIR01 = m_strIR01
   frm04010503_2.m_strIR02 = m_strIR02
   frm04010503_2.m_strIR03 = m_strIR03
   frm04010503_2.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010503_2.Show
   'Modify By Cheng 2002/06/21
'   Text7.SetFocus
   If Me.Option1(0).Value Then
      Option1_Click 0
   Else
      Option1_Click 1
   End If
   Me.Hide
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 4000: .Text = "專利名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "相關人"
      For i = 3 To 7
         .col = i: .ColWidth(i) = 0
      Next
      .col = 8: .ColWidth(8) = 0
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
