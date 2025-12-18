VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010501_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   3210
   ClientLeft      =   135
   ClientTop       =   1800
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8445
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   3
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2784
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   2
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2784
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   1
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2184
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2484
      Width           =   6960
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm04010501_2.frx":0000
      Left            =   960
      List            =   "frm04010501_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   1020
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6264
      TabIndex        =   5
      Top             =   72
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5472
      TabIndex        =   4
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7488
      TabIndex        =   6
      Top             =   72
      Width           =   800
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   8340
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8340
      Y1              =   2076
      Y2              =   2076
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   5
      Left            =   1200
      TabIndex        =   28
      Top             =   1800
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3281;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   960
      TabIndex        =   27
      Top             =   1560
      Width           =   3660
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6456;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   5430
      TabIndex        =   26
      Top             =   1320
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   25
      Top             =   1320
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2540;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   1680
      TabIndex        =   24
      Top             =   1050
      Width           =   5640
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "9948;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   5430
      TabIndex        =   23
      Top             =   720
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   2784
      Width           =   588
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4620
      TabIndex        =   21
      Top             =   2784
      Width           =   768
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "實審通知日期:"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   2184
      Width           =   1188
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   2484
      Width           =   768
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   4620
      TabIndex        =   16
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4620
      TabIndex        =   13
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   765
   End
End
Attribute VB_Name = "frm04010501_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer
Dim m_DefaultPrinter As String
Dim cp(9 To 13) As String
Dim m_NewCP09 As String 'Modified by Morgan 2014/4/14 改全域變數
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
'end 2014/1/14
Public m_DocWord As String 'Added by Morgan 2014/4/17
Dim strCP10 As String 'Add by Lydia 2014/11/18 改全域變數
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_bolFMP As Boolean 'Added by Lydia 2022/10/11 是否為FMP案
Dim m_bolTW121Case As Boolean, m_1911Msg As String 'Added by Morgan 2023/4/20 台灣有主張國內優先權案

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer
Dim strTmp As String  'ADD BY SONIA 2014/6/16

   Select Case Index
      Case 1 '回前畫面
         'frm04010501.Command1_Click
         frm04010501_1.Show
         Unload Me
      Case 2 '結束
         Unload frm04010501
         Unload frm04010501_1
         Unload Me
      Case 0 '確定
        'Add By Cheng 2003/06/03
        Screen.MousePointer = vbHourglass
         ' 90.08.09 modify by louis
         If IsEmptyText(Text5(1)) = True Then
            MsgBox "請輸入實審通知日期或文件齊備日期", vbOKOnly + vbCritical, "檢核資料"
            Text5(1).SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         ' 90.07.17 modify by louis
         'If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         '2011/12/5 MODIFY BY SONIA 加再審P-099864
         If cp(10) <> 異議_專 And cp(10) <> 舉發 And cp(10) <> 答辯 Then
            If Text5(2) = "" Then MsgBox "申請案號不可空白 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            If Text5(3) = "" Then MsgBox "申請日不可空白 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         End If
        'Add By Cheng 2003/03/26
        '檢查機關文號
        If pa(9) = 台灣國家代號 Then
            If Me.Text5(0).Tag = Me.Text5(0).Text Then
                MsgBox "請輸入機關文號!!!", vbExclamation + vbOKOnly
                Me.Text5(0).SetFocus
                Text5_GotFocus 0
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
         
         'Add By Sindy 2022/7/1
         If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
            If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         '2022/7/1 END
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理人 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         
         If m_1911Msg <> "" Then MsgBox m_1911Msg, vbInformation 'Added by Morgan 2023/4/20
         
'Remove by Morgan 2008/8/18 已改開窗定稿
'        '若智權人員為北所人員才要列印地址條
'        If GetST06(PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)) = "1" Then
'            '新增列印地址條資料
'            pub_AddressListSN = pub_AddressListSN + 1
'            PUB_AddNewAddressList strUserNum, pa(1), pa(2), pa(3), pa(4), "" & pub_AddressListSN, "0"
'        End If
        
         '通知函
         '92.7.29 modify by sonia
         'NowPrint cp(9), "04", "00", False, strUserNum, 0
'CANCEL BY SONIA 2014/6/16 P-107758
'         If pa(10) <> "" Then
            '94.2.21 MODIFY BY SONIA 取消申請日限制及不判斷專利種類改判斷案件性質102或302者
            'If pa(9) = "000" And pa(8) = "2" And pa(23) = "1" And pa(10) >= 920701 Then
            'Modify by Morgan 2006/10/5 加203主動修正,204修正,2007/3/29 加判斷專利種類
            'If pA(9) = "000" And pA(23) = "1" And (cp(10) = "102" Or cp(10) = "302") Then
            '94.2.21 END
            If pa(9) = "000" And pa(23) = "1" And pa(8) = "2" And (cp(10) = "102" Or cp(10) = "302" Or cp(10) = "203" Or cp(10) = "204") Then
            'end 2006/10/5
               
               If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                     NowPrint cp(9), "04", "02", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09 'add by toni  20080905 for 大對台定稿
               Else
                     StartLetter "04", "01" 'Added by Morgan 2023/4/20
                     NowPrint cp(9), "04", "01", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
               End If
            Else
               If pa(23) = "1" Then 'ADD BY SONIA 2014/6/16 爭議案之申請案號欄應抓相關總收文號之對造號數
                  If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        NowPrint cp(9), "04", "02", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09 'add by toni  20080905 for 大對台定稿
                  Else
                        StartLetter "04", "00" 'Added by Morgan 2023/4/20
                        NowPrint cp(9), "04", "00", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
                  End If
               'ADD BY SONIA 2014/6/16 爭議案之申請案號欄應抓相關總收文號之對造號數 P-107923,107758
               Else
                  strTmp = "03"       '台->台
                  If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                     strTmp = "04"    '大->台
                  End If
                  StartLetter "04", strTmp
                  NowPrint cp(9), "04", strTmp, False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
               End If
               'END 2014/6/16
            End If
'CANCEL BY SONIA 2014/6/16 P-107758
'         Else
'            NowPrint cp(9), "04", "00", False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
'         End If
         '92.7.29 end
         'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), strCP10, m_NewCP09
            PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), strCP10, pa(9), m_NewCP09
         End If
         Screen.MousePointer = vbDefault
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010501
            Unload frm04010501_1
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
            Unload frm04010501
            Unload frm04010501_1
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
            Unload frm04010501_1
            frm04010501.Show
            frm04010501.Clear
            Unload Me
         End If 'Added by Morgan 2014/1/14
         
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(1) = pa(5)
      Case "英"
         Label2(1) = pa(6)
      Case "日"
         Label2(1) = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010501_1.m_strIR01
   m_strIR02 = frm04010501_1.m_strIR02
   m_strIR03 = frm04010501_1.m_strIR03
   m_strIR04 = frm04010501_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   'Add By Cheng 2002/10/31
   '游標預設在機關文號
   SendKeys "{Tab}"
   
   'Added by Morgan 2019/9/6
   '電子公文時預設在申請日--玲玲
   If m_DocWord <> "" Then
      SendKeys "{Tab}"
   End If
   'end 2019/9/6
End Sub

Public Function QueryData() As Boolean
 Dim strTmp As String
   Combo1.ListIndex = 0
   cp(9) = strExc(2)
   cp(10) = strExc(3)
   cp(12) = strExc(4)
   cp(13) = strExc(5)
   With frm04010501_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
   End With
   ReadPatent
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
      
   Text5(0).Text = "（" & strTmp & "）智專一（二）字第號"
   
    'Add By Cheng 2003/03/26
    '記錄機關文號預設值
    Me.Text5(0).Tag = Me.Text5(0).Text
   
   'Added by Morgan 2014/1/14
   'Modified by Morgan 2014/4/17 +發文字
   If m_DocWord <> "" Then
     Text5(0) = m_DocWord & "字第" & m_DocNo & "號"
   ElseIf m_DocNo <> "" Then
      Text5(0) = Replace(Text5(0), "第號", "第" & m_DocNo & "號")
   End If
   'end 2014/1/14
   
   If frm04010501.Option1(0).Value Then
      Text5(2).Enabled = False
      Text5(3).Enabled = False
   End If
   '92.7.29 add by sonia
   If pa(10) <> "" Then
      '94/2/22 取消申請日條件, 專利種類新型改判斷案件性質102或302者
      'If pa(9) = "000" And pa(8) = "2" And pa(23) = "1" And pa(10) >= 920701 Then
      If pa(9) = "000" And pa(23) = "1" And (cp(10) = "102" Or cp(10) = "302") Then
         Label17 = "文件齊備日期:"
      End If
   End If
   '92.7.29 end
   
   'Add by Morgan 2004/7/15
   '需輸入申請日並檢查與原值是否一致，確認後存檔
   'Modify by Morgan 2004/7/21
   '新申請案才要控制
   'Modify by Morgan 2004/8/19 加 304,305,306,307
   'If InStr("101,102,103,104,105", cp(10)) > 0 Then
   If InStr("101,102,103,104,105,304,305,306,307", cp(10)) > 0 Then
      Text5(3).Text = ""
      Text5(3).Enabled = True
   End If
   
   'Added by Lydia 2022/10/11
   If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010501_2 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTempName As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
      
   Label2(5) = frm04010501_1.Label2(5)
   Text5(1) = frm04010501_1.Label2(5)
   
   Text5(2) = ""
   Text5(3) = ""
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label2(1) = pa(5)
      If pa(8) <> "" Then ChgType (2) ' Label2(2)
      If pa(9) <> "" Then ChgType (3) ' Label2(3)
      If pa(75) <> "" Then ChgType (4) ' Label2(4)
      Label2(0) = pa(11)
      Text5(2) = pa(11)
      Text5(3) = pa(10)
   End If
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, False, 台灣國家代號) = 1 Then
            Label2(2) = strTempName
         End If
      Case 4
         'Modify By Cheng 2002/07/08
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'         If objPublicData.GetAgent(pa(75), strTempName) Then
         If PUB_GetAgentName(pa(1), pa(75), strTempName) Then
            Label2(4) = strTempName
            ChgType = True
         End If
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strTempName) = True Then
         If ClsPDGetNation(pa(9), strTempName) = True Then
            Label2(3) = strTempName
            ChgType = True
         End If
   End Select
End Function

' 儲存資料表
Private Function FormSave() As Boolean
   Dim i As Integer, bolChk As Boolean, strTxt(1 To 5) As String
   Dim strCP09 As String ', strCP10 As String ---Add by Lydia 2014/11/18 改全域變數
   Dim strCP13 As String
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim ii As Integer
   Dim stCP12 As String, stCP13 As String
 
On Error GoTo ErrHnd
   FormSave = True
   cnnConnection.BeginTrans
   
   If IsEmptyText(Text5(1)) = True Then
      MsgBox "請輸入實審通知日期或文件齊備日期", vbOKOnly + vbCritical, "檢核資料"
      Exit Function
   End If
    
   stCP13 = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
   stCP12 = GetSalesArea(stCP13)
   m_NewCP09 = AutoNo("C", 6)
   strCP10 = 通知實審日
   If pa(10) <> "" Then
      '94.2.2 MODIFY BY SONIA 不考慮申請日
      'If pa(9) = "000" And pa(8) = "2" And pa(23) = "1" And pa(10) >= 920701 Then
      '94.2.21 MODIFY BY SONIA 不判斷專利種類改判斷案件性質為102及302者
      'If pa(9) = "000" And pa(8) = "2" And pa(23) = "1" Then
      'Modify by Morgan 2006/5/17 加204修正
      'If pA(9) = "000" And pA(23) = "1" And (cp(10) = "102" Or cp(10) = "302") Then
      If pa(9) = "000" And pa(23) = "1" And (cp(10) = "102" Or cp(10) = "302" Or cp(10) = "204") Then
         strCP10 = "1217"
         'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
         strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
            "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,cp119) VALUES ('" & Text1 & "','" & Text2 & "','" & _
            Text3 & "','" & Text4 & "'," & DBDATE(Text5(1)) & ",'" & _
            Text5(0) & "','" & m_NewCP09 & "','" & strCP10 & "','" & stCP12 & "','" & _
            stCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
            cp(9) & "'," & DBDATE(Label2(5)) & ")"
      Else
         
         'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
         strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
            "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,cp119) VALUES ('" & Text1 & "','" & Text2 & "','" & _
            Text3 & "','" & Text4 & "'," & DBDATE(Text5(1)) & ",'" & _
            Text5(0) & "','" & m_NewCP09 & "','" & strCP10 & "','" & stCP12 & "','" & _
            stCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
            cp(9) & "'," & DBDATE(Label2(5)) & ")"
      End If
   Else
      'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
      strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,cp119) VALUES ('" & Text1 & "','" & Text2 & "','" & _
         Text3 & "','" & Text4 & "'," & DBDATE(Text5(1)) & ",'" & _
         Text5(0) & "','" & m_NewCP09 & "','" & strCP10 & "','" & stCP12 & "','" & _
         stCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
         cp(9) & "'," & DBDATE(Label2(5)) & ")"
   End If
    'Modify end 2004/2/9
    
   'Added by Morgan 2014/1/14
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, m_NewCP09, pa(1), pa(2), pa(3), pa(4), strCP10
   End If
   'end 2014/1/14
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010501", IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
   End If
   '2016/10/5 END
   
   '92.7.29 end
   '92.2.19 END
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(1)
   strTxt(2) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & cp(9) & "' AND NP07='" & 通知實審日 & "'"
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(2)
    
    
   'Added by Morgan 2014/4/14 電子化-新增信函進度檔
   'Modified by Morgan 2015/5/12 改放在新增來函後否則函數內用收文號會抓不到本所號
   If pa(9) = "000" Then
      'Modified by Morgan 2018/8/1
      'strExc(1) = PUB_GetLetterJudge(pa(1), strCP10, , , pa(1), pa(2), pa(3), pa(4))
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), strCP10)
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      PUB_AddLetterProgress m_NewCP09, 1, True, strExc(1), False, pa(26), strCP10, pa(75)
   End If
   'end 2014/4/14
   
   i = 3
   
'2015/1/30 cancel by sonia 前申請案號輸入已更新,此處再更新則會覆蓋掉原申請案號
'   If Left(cp(10), 1) = "3" And frm04010501.Tag = "1" Then
'      strTxt(3) = "UPDATE CASEPROGRESS SET CP30='" & pa(11) & "' WHERE CP09='" & cp(9) & "'"
'        'Add By Cheng 2002/11/06
'        cnnConnection.Execute strTxt(3)
'      i = i + 1
'   End If
'2015/1/30
   
   '若案件性質不為異議或舉發, 則更新申請日及申請案號
   '2011/12/5 MODIFY BY SONIA 加再審P-099864
   If cp(10) <> 異議_專 And cp(10) <> 舉發 And cp(10) <> 答辯 Then
      strTxt(i) = "UPDATE PATENT SET PA10=" & TransDate(Text5(3), 2) & ",PA11='" & Text5(2) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(i)
      i = i + 1
   End If
    'Modify By Cheng 2002/11/06
'   FormSave = objLawDll.ExecSQL(I, strTxt)
'   FormSave = objLawDll.ExecSQL(i - 1, strTxt)
    'Modify By Cheng 2003/05/28
'   frm04010501.AddCust pa(26), pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
  
   'Added by Morgan 2022/11/21 111/8/19請作單--陳玲玲
   '台灣案若有主張國內優先權時基礎案要自動新增1911通知暫不續行審查
   'Added by Morgan 2023/4/20
   m_bolTW121Case = False
   m_1911Msg = ""
   'end 2023/4/20
   If pa(9) = "000" Then
      'Modified by Morgan 2023/4/18 分割案除外
      strExc(0) = "select * from caseprogress a,pridate,patent" & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
         " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='121' and cp27>0" & _
         " and not exists(select * From caseprogress where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='307')" & _
         " and pd01(+)=cp01 and pd02(+)=cp02 and pd03(+)=cp03 and pd04(+)=cp04 and pd07='000'" & _
         " and pa09(+)=pd07 and pa10(+)=pd05 and pa11(+)=pd06"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_bolTW121Case = True 'Added by Morgan 2023/4/20
         With RsTemp
         'Modified by Morgan 2023/4/20 基礎案為本所案件且尚未審定
         If Not IsNull(.Fields("pa01")) And IsNull(.Fields("pa16")) Then
            strExc(9) = AutoNo("C", 6)
            strExc(2) = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            strExc(3) = GetSalesArea(strExc(2))
            strExc(4) = .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") & .Fields("pa04") = "000", "", "-" & .Fields("pa03") & "-" & .Fields("pa04"))
            strExc(5) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
            strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp43,cp64 )" & _
               " select cp01,cp02,cp03,cp04," & strSrvDate(1) & " cp05,'" & strExc(9) & "' cp09" & _
               ",'1911' cp10,'" & strExc(3) & "' cp12,'" & strExc(2) & "' cp13,'" & strUserNum & "' cp14" & _
               ",'N' cp20,'N' cp26," & strSrvDate(1) & " cp27,'N' cp32,cp09 cp43,'" & strExc(5) & "通知實審;'" & _
               " from caseprogress where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "'" & _
               " and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and cp10 in ('101','102')"
            cnnConnection.Execute strSql, intI
            m_1911Msg = "基礎案 (" & strExc(4) & ") 已上暫不續行審查！"
            
            'Added by Morgan 2023/5/3
            '複製公文到母案
            If m_DocNo <> "" Then
               strExc(1) = PUB_GetEDocFileName(pa(1), pa(2), pa(3), pa(4), strCP10)
               strExc(2) = PUB_GetEDocFileName(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), "1911")
               CopyCPP m_NewCP09, strExc(1), strExc(9), strExc(2)
            End If
            'end 2023/5/3
         End If
         End With
      End If
   End If
   'end 2022/11/21
   
   cnnConnection.CommitTrans
   Exit Function
ErrHnd:
   cnnConnection.RollbackTrans
   FormSave = False
End Function

Private Sub Text5_GotFocus(Index As Integer)
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'將游標設定在機關文號欄的"字"的前面
If Index <> 0 Then
   InverseTextBox Text5(Index)
Else
   With Me.Text5(Index)
      If Len("" & .Text) > 0 Then
         'Modify by Morgan 2004/7/28
         '預設游標改在二
         'intPos = InStr("" & .Text, "字")
         intPos = InStr("" & .Text, "二")
         If intPos - 1 >= 0 Then
            .SelStart = intPos - 1
            .SelLength = 0
         End If
      End If
   End With
End If
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0 '機關文號
         If Text5(Index) = "" Then
            MsgBox "機關文號不可空白 !", vbCritical
            Cancel = True
         Else
            'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
            If CheckLengthIsOK(Text5(Index), Text5(Index).MaxLength) = False Then
               Cancel = True
            End If
         End If
      Case 1 '實審通知日
         If IsEmptyText(Text5(Index)) = False Then
            If ChkDate(Text5(Index)) Then
               If Val(Text5(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "日期不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         End If
      Case 3 '申請日
         If Text5(Index) <> "" Then
            If ChkDate(Text5(Index)) Then
               If Val(Text5(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "日期不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         End If
      Case 2 '申請案號
         If Text5(Index) <> "" Then
            '2005/6/14 MODIFY BY SONIA
            'If Not ChkAppNo(Text5(Index).Text, pa(8), 0) Then
            If Not ChkAppNo(Text5(Index).Text, pa(8), 0, Val(pa(23))) Then
            '2005/6/14 END
               Cancel = True
            End If
         End If
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text5
   If objTxt.Enabled = True Then
      Cancel = False
      Text5_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text5(objTxt.Index).SetFocus
         Text5_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/7/15
'需輸入申請日並檢查與原值是否一致，確認後存檔
If Text5(3).Enabled = True Then
   If Val(pa(10)) <> 0 Then
      If Val(Text5(3)) <> Val(pa(10)) Then
         'Modified by Morgan 2014/5/20 改控制不同不可存檔--玲玲
         'If MsgBox("新輸入的申請日與原申請日期 " & pa(10) & " 不同，確定要更改？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
          MsgBox "輸入的申請日與原申請日期 " & pa(10) & " 不同，請更正！", vbCritical
            Text5_GotFocus 3
            Text5(3).SetFocus
            Exit Function
         'End If
      End If
   End If
End If

'Added by Morgan 2014/5/15 電子化-檢查pdf檔
If pa(9) = "000" Then
   If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1, m_DocNo) = False Then
      Exit Function
   End If
End If
'end 2014/5/15

TxtValidate = True
End Function

'Add By Cheng 2003/05/30
'取得員工所別
Private Function GetST06(strST01 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetST06 = ""
   StrSQLa = "Select ST06 From Staff Where ST01='" & strST01 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetST06 = "" & rsA("ST06").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Function

'ADD BY SONIA 2014/6/16
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, i As Integer
   
   EndLetter ET01, cp(9), ET03, strUserNum
   i = 1
   
   strExc(0) = "SELECT CP36 FROM CASEPROGRESS WHERE CP01='" & Text1 & "' AND CP02='" & Text2 & _
      "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND CP09='" & cp(9) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','相關總收文號對造號數','" & "" & RsTemp.Fields(0) & "')"
      i = i + 1
   End If
   
   'Added by Morgan 2023/4/20
   If m_bolTW121Case Then
      strExc(0) = "select * from pridate,patent,patenttrademarkmap" & _
         " where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "'" & _
         " and pd04='" & pa(4) & "' and pd07='000'" & _
         " and pa09(+)=pd07 and pa10(+)=pd05 and pa11(+)=pd06" & _
         " and ptm01(+)='1' and ptm02(+)=pd08"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','基礎案申種類','" & .Fields("ptm03") & "')"
         i = i + 1
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','基礎案申請號','" & .Fields("pd06") & "')"
         i = i + 1
         If Not IsNull(.Fields("pa05")) Then
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','基礎案名稱','「" & .Fields("pa05") & "」')"
            i = i + 1
         End If
         If .Fields("pd08") = "1" Then
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','基礎發明案不公開','並不予公開')"
            i = i + 1
         End If
         '基礎案為本所發明案且未審定且已發文實審且未有審查意見來函
         If .Fields("pa08") = "1" And IsNull(.Fields("pa16")) Then
            strExc(0) = "select * from caseprogress a" & _
               " where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "'" & _
               " and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "'" & _
               " and cp10='416' and cp27>0 and not exists(select * from caseprogress b" & _
               " where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='1202')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
                  "','可退審查費才印','♀')"
               i = i + 1
            End If
         End If
         End With
      End If
   End If
   'end 2023/4/20
   
   If Not ClsLawExecSQL(i - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub
'END 2014/6/16

'Added by Morgan 2023/5/3
'複製卷宗區檔案
Private Sub CopyCPP(pSrcCPP01 As String, pSrcCPP02 As String, pToCPP01 As String, pToCPP02 As String)
   Dim stTempPath As String
   Dim fs, f
   
   stTempPath = App.path & "\" & strUserNum
   If Dir(stTempPath, vbDirectory) = "" Then
      MkDir stTempPath
   End If
   
   If PUB_GetAttachFile_CPP(pSrcCPP01, pSrcCPP02, stTempPath) Then
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(pSrcCPP02)
      If Not SaveAttFile_PDF(pToCPP01, pSrcCPP02, pToCPP02, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True) Then
         Err.Raise 999, , "複製卷宗區檔案失敗！作業中斷！"
      End If
   End If
   Set fs = Nothing
   Set f = Nothing
End Sub
