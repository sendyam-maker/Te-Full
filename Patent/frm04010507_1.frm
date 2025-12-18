VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010507_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   930
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9330
   Begin VB.TextBox txtBillno 
      Height          =   285
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1590
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   3
      Top             =   660
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(F)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6690
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   8364
      TabIndex        =   9
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   7536
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "P"
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      MaxLength       =   1
      TabIndex        =   5
      Top             =   5400
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3315
      Left            =   120
      TabIndex        =   7
      Top             =   1980
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   5847
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
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   1260
      Width           =   8115
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14314;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   3480
      TabIndex        =   21
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "帳單編號:"
      Height          =   180
      Left            =   5250
      TabIndex        =   20
      Top             =   1650
      Width           =   765
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Index           =   2
      Left            =   1080
      TabIndex        =   19
      Top             =   1650
      Width           =   4020
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "7091;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Index           =   1
      Left            =   4440
      TabIndex        =   18
      Top             =   960
      Width           =   3150
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "5556;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Index           =   0
      Left            =   1080
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "3651;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "結果：             1.已收達 2.已提申"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "專利號數："
      Height          =   180
      Left            =   3480
      TabIndex        =   15
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1650
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm04010507_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1,Label5)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Public m_TEXT2 As String     '2008/11/28 ADD BY SONIA 記錄是否以申請案號條件輸資料

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer
'Removed by Morgan 2023/5/17 沒用了
'Dim bolHaveAppNo As Boolean 'Add by Morgan 2010/4/15 是否已有申請號
'end 2023/5/17
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2016/10/7 END


Public Sub SetInputFocus()
   'Add By Cheng 2002/12/18
   '2008/11/28 ADD BY SONIA 決定游標停放位置
   'Text1(0).SetFocus
   'TextInverse Me.Text1(0)
   If m_TEXT2 <> "" Then
      Text2.SetFocus
      TextInverse Text2
   Else
      Text1(1).SetFocus
      TextInverse Me.Text1(1)
   End If
   '2008/11/28 END
End Sub

Public Sub Clear()
    'Modify By Cheng 2002/12/18
    '保留原輸入條件
   'Text1(0) = Empty
    '2009/3/11 modify by sonia 因加入申請案號條件改為要清掉輸入條件
   Text1(1) = Empty
   Text1(2) = Empty
   Text1(3) = Empty
   Label5(0) = Empty
   Label5(1) = Empty
   Label5(2) = Empty
   Combo1.Clear
   Text3 = Empty
   Text2 = Empty        '2009/3/11 add

   'Modified by Morgan 2012/3/27 改 9 欄
   InitGrid 9, MSHFlexGrid1
   GridHead
   'cmdOK(0).Enabled = False
   'Add By Cheng 2003/02/14
   Me.Command1.Default = True
   
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 0
         If Text3.Text = "" Then
            MsgBox "請選擇結果 !", vbCritical
            Text3.SetFocus
         Else
            '2008/11/28 ADD BY SONIA 結果為 已提申 時, 必須以申請案號條件輸入
            'Modify by Morgan 2010/4/16 有申請案號的才要(FMP案會缺文件而無申請號)
            'If Text3.Text = "2" And Text2.Text = "" Then
            'Modified by Morgan 2023/5/17
            'If Text3.Text = "2" And Text2.Text = "" And bolHaveAppNo Then
            '   MsgBox "結果為 已提申 時, 必須以申請案號條件輸入 !", vbInformation
            '   Exit Sub
            If Text3.Text = "2" And pa(11) <> "" Then
               If Text2.Text = "" Then
                  MsgBox "結果為 已提申 時, 必須輸入申請案號 !", vbInformation
                  Text2.SetFocus
                  Exit Sub
               ElseIf Text2.Text <> pa(11) Then
                  MsgBox "申請案號輸入錯誤 !", vbExclamation
                  Text2.SetFocus
                  Text2_GotFocus
                  Exit Sub
               End If
            'end 2023/5/17
            End If
            '2008/11/28 END
            For i = 1 To MSHFlexGrid1.Rows - 1
               If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
                  '92.5.6 add by sonia
                  If Text3.Text = "2" And MSHFlexGrid1.TextMatrix(i, 4) = "" Then
                     MsgBox "結果為 已提申 時, 請選擇已發文資料 !", vbInformation
                     Exit Sub
                  End If
                  '92.5.6 end
                  '2006/5/9 ADD BY SONIA
                  If Text3.Text = "2" And InStr((CaseMapIn & "301,302,303"), MSHFlexGrid1.TextMatrix(i, 7)) > 0 Then
                     MsgBox "新申請案不可在此輸入已提申資料 !", vbInformation
                     Exit Sub
                  End If
                  '2006/5/9 END
                  'Add by Morgan 2004/3/23
                  '結果輸入'1'已收達時不可選則已收達的資料
                  'Modified by Morgan 2019/2/27 改可重複輸入已收達--玲玲
                  'If Text3.Text = "1" And MSHFlexGrid1.TextMatrix(i, 5) <> "" Then
                  '   MsgBox "結果為 已收達 時, 請選擇未收達資料 !", vbInformation
                  '   Exit Sub
                  'End If
                  'end 2019/2/27
                  'Add end
                  bolChk = True
                  Me.Tag = MSHFlexGrid1.TextMatrix(i, 1)
                  strExc(1) = Text1(0)
                  strExc(2) = Text1(1)
                  strExc(3) = Text1(2)
                  strExc(4) = Text1(3)
                  Exit For
               End If
            Next
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
            Me.Hide
            'Add By Sindy 2016/10/7
            frm04010507_2.m_strIR01 = m_strIR01
            frm04010507_2.m_strIR02 = m_strIR02
            frm04010507_2.m_strIR03 = m_strIR03
            frm04010507_2.m_strIR04 = m_strIR04
            '2016/10/7 END
            frm04010507_2.Show
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()

   '搜尋資料
   '2008'11/28 add by sonia 加入已提申時要以申請案號輸入
   m_TEXT2 = ""
   If Text2 <> "" Then
      m_TEXT2 = Text2
      'Add by Lydia 2014/10/31 設別名f0,+FMP2openSQL
      strExc(0) = "SELECT PA01,PA02,PA03,PA04 FROM PATENT f0 WHERE PA11='" & Text2 & "' AND PA09<>'000' " & FMP2openSQL
      strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
            If FMP2open = True And FMP2openSQL <> "" Then
               MsgBox "權限不足 !", vbInformation
            Else
               MsgBox "資料庫內無此申請案號之資料", vbInformation
            End If
         Text2.SetFocus
         Me.Command1.Default = True
         Exit Sub
      Else
         Text1(0) = RsTemp.Fields("PA01")
         Text1(1) = RsTemp.Fields("PA02")
         Text1(2) = RsTemp.Fields("PA03")
         Text1(3) = RsTemp.Fields("PA04")
      End If
   End If
   '2008/11/28 end
   Text1_LostFocus 3
End Sub

Private Sub Form_Activate()
   Dim i As Integer, j As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
         For j = 0 To .Cols - 1
            .col = j
            .CellBackColor = .BackColor
         Next
      Next
   End With
   Text3.Text = ""
   m_TEXT2 = ""    '2008/12/1 ADD BY SONIA
   
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      Text1(0).Text = m_strCP01
      Text1(1).Text = m_strCP02
      Text1(2).Text = m_strCP03
      Text1(3).Text = m_strCP04
      Text2.Text = m_AppNo
      Command1.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()

 Dim Lbl As Object
   MoveFormToCenter Me
   intWhere = 國內
   'Modified by Morgan 2012/3/27 改 9 欄
   InitGrid 9, MSHFlexGrid1
   GridHead
   For Each Lbl In Label5
      Lbl.Caption = ""
   Next
   'cmdOK(0).Enabled = False
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010507_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim ii As Integer
    
    GridClick MSHFlexGrid1, intLastRow, 0
    'Add By Cheng 2003/04/02
    If Me.MSHFlexGrid1.Rows > 1 Then
        Me.Text3.SetFocus
        For ii = 1 To Me.MSHFlexGrid1.Rows - 1
            '若為勾選此筆
            If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
                'Modify by Morgan 改控制輸入本所號時預設已收達，申請號預設已提申，不必再管案件性質--玲玲
                ''若案件性質為年費(605)
                'If Me.MSHFlexGrid1.TextMatrix(ii, 7) = "605" Then
                '     Me.Text3.Text = "2"
                '     Text3_GotFocus
                If Text2 <> "" Then
                    Me.Text3.Text = "2"
                    Text3_GotFocus
                Else
                    Me.Text3.Text = "1"
                    Text3_GotFocus
                End If
            End If
        Next ii
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Dim Lbl As Object, strTmp As String, bolChk As Boolean
   'Add By Cheng 2002/07/08
   Dim StrSQLa As String
   
    'Add By Cheng 2003/02/14
    '2008/11/28 MODIFY BY SONIA
    'If Index = 3 Then
    If Index = 3 And Text1(1) <> "" Then
        'Modify By Cheng 2003/04/02
'        If MSHFlexGrid1.Rows = 2 Then
'           GridClick MSHFlexGrid1, 1, 0
'           Text3.SetFocus
'        End If
        'Add By Cheng 2002/07/08
        StrSQLa = ""
        Select Case Index
           Case 3
               
              For Each Lbl In Label5
                 Lbl.Caption = ""
              Next
              If Text1(2) = "" Then Text1(2) = "0"
              If Text1(3) = "" Then Text1(3) = "00"
              
       'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(0), Text1(1), Text1(2), Text1(3)) = False Then
             If Text2 <> "" Then
                Text2.SetFocus
             Else
                Text1(1).SetFocus
             End If
             Exit Sub
           End If
         End If
        
              pa(1) = Text1(0)
              pa(2) = Text1(1)
              pa(3) = Text1(2)
              pa(4) = Text1(3)
              
              'Removed by Morgan 2023/5/17 沒用了
              'bolHaveAppNo = False 'Add by Morgan 2010/4/16
              'end 2023/5/17
              If pa(1) = "P" Then
                 If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
                    AddCboName Combo1, pa(5), pa(6), pa(7)
                    Label5(0) = pa(11)
                    Label5(1) = pa(22)
                    If pa(26) <> "" Then
                       'edit by nickc 2007/02/05 不用 dll 了
                       'If objLawDll.LawGetName(pa(26), strTmp) Then
                       If ClsLawLawGetName(pa(26), strTmp) Then
                          Label5(2) = strTmp
                       End If
                    End If
                    'Removed by Morgan 2023/5/17 沒用了
                    'If pa(11) <> "" Then bolHaveAppNo = True 'Add by Morgan 2010/4/16
                    'end 2023/5/17
                 Else
                    'Text1(Index).SetFocus
                    Text1(1).SetFocus
                    'Add By Cheng 2003/02/14
                    Me.Command1.Default = True
                    Exit Sub
                 End If
              ElseIf pa(1) = "PS" Then
                 If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
                    AddCboName Combo1, pa(5), pa(6), pa(7)
                    Label5(0) = pa(11)
                    If pa(26) <> "" Then
                       'edit by nickc 2007/02/05 不用 dll 了
                       'If objLawDll.LawGetName(pa(26), strTmp) Then
                       If ClsLawLawGetName(pa(26), strTmp) Then
                          Label5(2) = strTmp
                       End If
                    End If
                    'Removed by Morgan 2023/5/17 沒用了
                    'If pa(11) <> "" Then bolHaveAppNo = True 'Add by Morgan 2010/4/16
                    'end 2023/5/17
                 End If
              End If
              If pa(9) = 台灣國家代號 Then
                 strTmp = "CPM03"
                 '92.6.28 add by sonia
                 MsgBox "此案件之申請國家為 台灣 !!", vbOKOnly + vbCritical, "檢核資料"
                 Text1(1).SetFocus
                 Me.Command1.Default = True
                 Exit Sub
                 '92.6.28 end
              Else
                 strTmp = "CPM04"
              End If
              'Modify By Cheng 2002/04/15
        '         strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & ",NVL(FA05,NVL(FA04,FA06))," & _
        '            SQLDate("CP46") & " FROM caseprogress, casepropertymap,FAGENT where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
        '            " and cp27 is NOT null and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') and CP47 IS NULL AND CP24 IS NULL AND " & _
        '            "CP61 IS NULL AND cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+)"
              'Modify By Cheng 2002/07/08
              '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
        '         strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & ",NVL(FA05,NVL(FA04,FA06))," & _
        '            SQLDate("CP46") & " FROM caseprogress, casepropertymap,FAGENT where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
        '            " and cp27 is NOT null and ( cp09<'C' ) and CP47 IS NULL AND CP24 IS NULL AND " & _
        '            "CP61 IS NULL AND cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+)"
        ' 91.09.13 modify by louis
        '         strSQLA = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
        '         strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & "," & strSQLA & " " & _
        '            SQLDate("CP46") & " FROM caseprogress, casepropertymap,FAGENT,SystemKind where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
        '            " and cp27 is NOT null and ( cp09<'C' ) and CP47 IS NULL AND CP24 IS NULL AND " & _
        '            "CP61 IS NULL AND cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP01=SK01(+) "
        '         strSQLA = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
                'Modify By Cheng 2003/04/02
'              strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & "," & strSQLA & " " & _
'                 SQLDate("CP46") & ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & " FROM caseprogress, casepropertymap,FAGENT,SystemKind where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                 " and cp27 is NOT null and ( cp09<'C' ) and CP47 IS NULL AND CP24 IS NULL AND " & _
'                 "cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP01=SK01(+) " & _
'                 "ORDER BY SORTFIELD DESC "
               '2005/4/11 MODIFY BY SONIA 取消CP47 IS NULL的控制
               'strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & "," & strSQLA & " " & _
               '  SQLDate("CP46") & ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD, CP10  FROM caseprogress, casepropertymap,FAGENT,SystemKind where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               '  " and ( cp09<'C' ) and CP47 IS NULL AND CP24 IS NULL AND " & _
               '  "cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP01=SK01(+) " & _
               '  "ORDER BY SORTFIELD DESC "
               
               '20140327START MODIFY By eric  CP24 IS NULL 改為 (CP24 IS NULL or CP10='601')
               'Modify by Morgan 2010/1/15 +CP27>0 控制
               'Modified by Morgan 2012/3/27 +CP64
               strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & "," & StrSQLa & " " & _
                 SQLDate("CP46") & ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD, CP10,CP64  FROM caseprogress, casepropertymap,FAGENT,SystemKind where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                 " and ( cp09<'C' ) and (CP24 IS NULL OR CP10='601') AND CP27>0 AND " & _
                 "cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP01=SK01(+) " & _
                 "ORDER BY SORTFIELD DESC "
               'strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & "," & StrSQLa & " " & _
               '  SQLDate ("CP46") & ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD, CP10,CP64  FROM caseprogress, casepropertymap,FAGENT,SystemKind where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               '  " and ( cp09<'C' ) and CP24 IS NULL AND CP27>0 AND " & _
               '  "cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP01=SK01(+) " & _
               '  "ORDER BY SORTFIELD DESC "
               '2005/4/8 END
               '20140327END
              
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
              If intI <> 2 Then
                 Set MSHFlexGrid1.Recordset = RsTemp
              End If
              GridHead
              '若有搜尋到資料
              If MSHFlexGrid1.Rows > 1 Then
              '   cmdOK(0).Enabled = True
                 Me.cmdOK(0).Default = True
              '若無資料
              Else
                 MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "檢核資料"
                 Text1(1).SetFocus
                 Me.Command1.Default = True
              End If
        End Select
        If MSHFlexGrid1.Rows = 2 Then
           GridClick MSHFlexGrid1, 1, 0
           Text3.SetFocus
        End If
    End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim Lbl As Object, strTmp As String, bolChk As Boolean
   Select Case Index
      Case 0 '系統類別
         If Text1(Index) <> "P" And Text1(Index) <> "PS" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            TextInverse Text1(Index)
            Cancel = True
         End If
      Case 3 '多國多類碼
         'For Each Lbl In Label5
         '   Lbl.Caption = ""
         'Next
         'If Text1(2) = "" Then Text1(2) = "0"
         'If Text1(3) = "" Then Text1(3) = "00"
         'pa(1) = Text1(0)
         'pa(2) = Text1(1)
         'pa(3) = Text1(2)
         'pa(4) = Text1(3)
         '
         'If pa(1) = "P" Then
         '   If objPublicData.ReadPatentDatabase(pa(), intWhere) Then
         '      AddCboName Combo1, pa(5), pa(6), pa(7)
         '      Label5(0) = pa(11)
         '      Label5(1) = pa(22)
         '      If pa(26) <> "" Then
         '         If objLawDll.LawGetName(pa(26), strTmp) Then
         '            Label5(2) = strTmp
         '         End If
         '      End If
         '   Else
         '      Text1(Index).SetFocus
         '      Exit Sub
         '   End If
         'ElseIf pa(1) = "PS" Then
         '   If objPublicData.ReadServicePracticeDatabase(pa(), intWhere) Then
         '      AddCboName Combo1, pa(5), pa(6), pa(7)
         '      Label5(0) = pa(11)
         '      If pa(26) <> "" Then
         '         If objLawDll.LawGetName(pa(26), strTmp) Then
         '            Label5(2) = strTmp
         '         End If
         '      End If
         '   End If
         'End If
         'If pa(10) = 台灣國家代號 Then
         '   strTmp = "CPM03"
         'Else
         '   strTmp = "CPM04"
         'End If
         '
         'strExc(0) = "select '',cp09," & SQLDate("CP05") & "," & strTmp & "," & SQLDate("CP27") & ",NVL(FA05,NVL(FA04,FA06))," & _
         '   SQLDate("CP46") & " FROM caseprogress, casepropertymap,FAGENT where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         '   " and cp27 is NOT null and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') and CP47 IS NULL AND CP24 IS NULL AND " & _
         '   "CP61 IS NULL AND cp01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+)"
         'intI = 1
         'Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         'If intI <> 2 Then
         '   Set MSHFlexGrid1.Recordset = rsTemp
         'End If
         'GridHead
         'If MSHFlexGrid1.Rows > 1 Then
         '   cmdOK(0).Enabled = True
         'Else
         '   MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "檢核資料"
         'End If
   End Select
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "代理人收達日"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 0: .Text = "代理人"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 0: .Text = "案件性質代號"
      .col = 8: .ColWidth(8) = 2350: .Text = "進度備註" 'Added by Morgan 2012/3/27
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
