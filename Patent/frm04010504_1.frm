VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010504_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函輸入"
   ClientHeight    =   5750
   ClientLeft      =   -2690
   ClientTop       =   1410
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9340
   Begin VB.Frame Frame1 
      Height          =   1152
      Left            =   168
      TabIndex        =   14
      Top             =   624
      Width           =   9012
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "P"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   8
         Top             =   780
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   192
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1070
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1070
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "專利號數"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1070
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7530
      TabIndex        =   12
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8376
      TabIndex        =   13
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1860
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3252
      Left            =   180
      TabIndex        =   11
      Top             =   2340
      Width           =   9012
      _ExtentX        =   15893
      _ExtentY        =   5733
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   9180
      Y1              =   2256
      Y2              =   2256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   9180
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   300
      TabIndex        =   15
      Top             =   1860
      Width           =   948
   End
End
Attribute VB_Name = "frm04010504_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intCols As Integer
Dim intWhere As Integer
Dim m_PA09 As String 'Added by Morgan 2012/4/19
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
'end 2014/1/14
'Added by Morgan 2014/4/17
Public m_DocWord As String
Public m_DeadLine As String
Public m_NewCP10 As String
'end 2014/4/17
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
'2016/10/5 END


Public Sub Clear()
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text6 = Empty
   Text7 = Empty
   InitGrid 9, MSHFlexGrid1
   GridHead
   Option1_Click 0
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Private Sub SetGridData(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   Dim nCol As Integer
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
      nRow = MSHFlexGrid1.Rows - 1
      For nCol = 0 To 7
         If Not IsNull(rsTmp.Fields(nCol)) Then
            MSHFlexGrid1.TextMatrix(nRow, nCol) = rsTmp.Fields(nCol)
         End If
      Next nCol
      rsTmp.MoveNext
   Loop
End Sub

Private Sub QueryByPA11()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim sq01 As Integer, sq02 As Integer, sq03 As Integer 'Add by Lydia 2014/10/31
   InitGrid 9, MSHFlexGrid1
   GridHead
   'Modified by Morgan 2012/4/19 ''-->sp09
   'strSql = "select " & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
            "'' as RName,'',SP01,SP02,SP03,SP04,sp09 from SERVICEPRACTICE where SP01='PS' AND " & _
            "SP11='" & Text7 & "'"
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   '設別名f0,+FMP2openSQL
   strSql = "select " & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
            "'' as RName,'',SP01,SP02,SP03,SP04,sp09 from SERVICEPRACTICE f0 where SP01='PS' AND " & _
            "SP11='" & Text7 & "' " & FMP2openSQL
   strSql = Replace(strSql, "f0.CP", "f0.SP")
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   sq01 = rsTmp.RecordCount 'Add by Lydia 2014/10/31
   If rsTmp.RecordCount > 0 Then
      m_PA09 = "" & rsTmp("sp09") 'Added by Morgan 2012/4/19
      SetGridData rsTmp
   End If
   rsTmp.Close
   
   'Modified by Morgan 2012/4/19 ''-->pa09
   'strSql = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
            "'' as RName,'',pa01,pa02,pa03,pa04,pa09 from patent where PA01='P' AND " & _
            "pa11='" & Text7 & "'"
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   '設別名f0,+FMP2openSQL
   strSql = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
            "'' as RName,'',pa01,pa02,pa03,pa04,pa09 from patent f0 where PA01='P' AND " & _
            "pa11='" & Text7 & "' " & FMP2openSQL

   strSql = Replace(strSql, "f0.CP", "f0.PA")
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   sq02 = rsTmp.RecordCount
   If rsTmp.RecordCount > 0 Then
      m_PA09 = "" & rsTmp("pa09") 'Added by Morgan 2012/4/19
      SetGridData rsTmp
   End If
   rsTmp.Close
   
   'Modified by Morgan 2012/4/19 ''-->pa09
   strSql = "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No,nvl(cp37,nvl(cp38,cp38)) as Name," & _
            "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 from caseprogress,patent " & _
            "where (CP01='P' OR CP01='PS') AND cp36='" & Text7 & "' " & _
            "and (cp01,cp02,cp03,cp04) not in " & _
            "(select pa01,pa02,pa03,pa04 from patent where PA01='P' AND " & _
            "pa11='" & Text7 & "' UNION " & _
            "select SP01,SP02,SP03,SP04 from SERVICEPRACTICE where SP01='PS' AND " & _
            "SP11='" & Text7 & "')" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   sq03 = rsTmp.RecordCount
   If rsTmp.RecordCount > 0 Then
      m_PA09 = "" & rsTmp("pa09") 'Added by Morgan 2012/4/19
      SetGridData rsTmp
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Add by Lydia 2014/10/31 提示訊息
   If sq01 + sq02 + sq03 = 0 Then
      If FMP2open = True Then
         MsgBox "權限不足 !", vbInformation
      Else
         MsgBox "資料庫查無資料 !", vbInformation
      End If
   End If
      
   GridHead
End Sub

Private Sub Command1_Click()
   'Added by Morgan 2012/4/19
   m_PA09 = ""
   intI = 1
   strExc(0) = "select " & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
         "'' as RName,SP11,SP01,SP02,SP03,SP04,SP09 from SERVICEPRACTICE where rownum<1"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set MSHFlexGrid1.Recordset = RsTemp
   End If
   GridHead
   'end 2012/4/19
   
   intI = 0
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
      QueryByPA11
   ElseIf Option1(1).Value = True Then
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
     ' strExc(0) = "select " & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
         "'' as RName,SP11,SP01,SP02,SP03,SP04,SP09 from SERVICEPRACTICE where SP01='" & Text1 & _
         "' and SP02='" & Text2 & "' and SP03='" & Text3 & "' and SP04='" & Text4 & _
         "' union " & _
         "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,pa11,pa01,pa02,pa03,pa04,PA09 from patent where pa01='" & Text1 & _
         "' and pa02='" & Text2 & "' and " & _
         "pa03='" & Text3 & "' and pa04='" & Text4 & "'"
 'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        '設別名f0,+FMP2openSQL
          strExc(0) = "select " & ChgService("", 1) & " as No,nvl(sp05,nvl(sp06,sp07)) as Name," & _
            " '' as RName,SP11,SP01,SP02,SP03,SP04,SP09 from ServicePractice f0 where SP01='" & Me.Text1.Text & "' AND Sp02='" & Me.Text2.Text & "' " & _
            " and SP03='" & Text3 & "' and SP04='" & Text4 & "'" & FMP2openSQL
          strExc(0) = Replace(strExc(0), "f0.CP", "f0.SP")
          strExc(0) = strExc(0) & " union " & _
            "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
            "'' as RName,pa11,pa01,pa02,pa03,pa04,PA09 from patent f2 where pa01='" & Text1 & _
            "' and pa02='" & Text2 & "' and pa03='" & Text3 & "' and pa04='" & Text4 & "' " & FMP2openSQL
          strExc(0) = Replace(strExc(0), "f0.CP", "f2.PA")
          
         'Add by Lydia 2014/10/31 提示訊息
         If FMP2open = True And FMP2openSQL <> "" Then intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 0 And FMP2open = True And FMP2openSQL <> "" Then
            MsgBox "權限不足 !", vbInformation
            Exit Sub
         End If
        '.end 'Add by Lydia 2014/10/31 提示訊息
   
      If intI = 1 Then
         If RsTemp.Fields("sp09") = 台灣國家代號 And RsTemp.Fields("sp11") <> "" Then
            MsgBox "本案申請國家為台灣且已有申請案號, 請改以申請案號查詢!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
         If intI = 1 Then
            
            'Added by Morgan 2012/4/19
            '非台灣要清除來函收文日重新輸入
            m_PA09 = "" & RsTemp.Fields("sp09")
            If m_PA09 <> 台灣國家代號 Then
            'end 2012/4/19
               Set MSHFlexGrid1.Recordset = RsTemp
               GridHead
            Else
               If TxtValidate = False Then Exit Sub
               '進入畫面二
               strExc(1) = Text1
               strExc(2) = Text2
               strExc(3) = Text3
               strExc(4) = Text4
               'Add By Sindy 2017/12/27
               If m_strIR01 <> "" Then
                  If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
                     MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                     Exit Sub
                  End If
               End If
               '2017/12/27 END
               'Add By Sindy 2016/10/5
               frm04010504_2.m_strIR01 = m_strIR01
               frm04010504_2.m_strIR02 = m_strIR02
               frm04010504_2.m_strIR03 = m_strIR03
               frm04010504_2.m_strIR04 = m_strIR04
               '2016/10/5 END
               frm04010504_2.Show
               frm04010504_2.JumpIfOneRecord
               Me.Hide
            End If
         End If
      End If
   ElseIf Option1(2).Value = True Then
      If Text6 = "" Then MsgBox "專利號數不得空白，請重新輸入 !", vbCritical: Exit Sub
      'Modified by Morgan 2012/4/19 ''-->pa09
      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
      '設別名f0,+FMP2openSQL
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,pa09 from patent f0 where PA01='P' AND " & _
         "pa22='" & Text6 & "' " & FMP2openSQL & " union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No,nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,pa09 from caseprogress,patent where " & _
         "CP01='P' AND cp36='" & Text6 & "' and (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='P' AND pa22='" & Text6 & "')" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
         strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
         'Add by Lydia 2014/10/31 提示訊息
         If FMP2open = True And FMP2openSQL <> "" Then intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 0 And FMP2open = True And FMP2openSQL <> "" Then
            MsgBox "權限不足 !", vbInformation
            Exit Sub
         End If
        '.end 'Add by Lydia 2014/10/31 提示訊息
     
      If intI = 1 Then m_PA09 = "" & RsTemp.Fields("pa09") 'Added by Morgan 2012/4/19
      
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
   End If
   
   'Added by Morgan 2012/4/19
   '非台灣要清除來函收文日重新輸入
   If m_PA09 <> "" And m_PA09 <> 台灣國家代號 Then
      'Add By Sindy 2017/10/16
      If Not (Text5.Tag = m_RDate And m_RDate <> "") Then
      '2017/10/16 END
         Text5 = ""
         Text5.SetFocus
      End If
   End If
   'end 2012/4/19
   
   ' 90.07.05 modify by louis (只有一筆時直接進入到下一個畫面)
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      GridClick MSHFlexGrid1, 1, 8
      'If m_PA09 = 台灣國家代號 Then 'Added by Morgan 2012/4/19
      'Add By Sindy 2017/10/16 + Or (Text5.Tag = m_RDate And m_RDate <> "")
      If m_PA09 = 台灣國家代號 Or (Text5.Tag = m_RDate And m_RDate <> "") Then 'Added by Morgan 2012/4/19
         FormConfirm
      End If 'Added by Morgan 2012/4/19
   End If
   
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
   Option1_Click (0)
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010504_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 8
   cmdOK(0).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
 On Error Resume Next
   Select Case Index
      Case 0
         Text7.Enabled = True
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text6.Enabled = False
         Text7.SetFocus
      Case 1
         Text7.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text6.Enabled = False
         Text1.SetFocus
      Case 2
         Text7.Enabled = False
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text6.Enabled = True
         Text6.SetFocus
   End Select
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "P" And Text1 <> "PS" Then
      MsgBox "系統別錯誤，請重新輸入 !", vbCritical
      Cancel = True
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
   If Text7.Enabled Then Text7.SetFocus
   'Add By Sindy 2017/12/27
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2017/12/27 END
   'Add By Sindy 2016/10/5
   frm04010504_2.m_strIR01 = m_strIR01
   frm04010504_2.m_strIR02 = m_strIR02
   frm04010504_2.m_strIR03 = m_strIR03
   frm04010504_2.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010504_2.Show
   frm04010504_2.JumpIfOneRecord
   'Command1.SetFocus
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
      .CellAlignment = flexAlignCenterCenter
      For i = 3 To 8
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_LostFocus()
   Command1.Default = True 'Added by Morgan 2020/9/15
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
