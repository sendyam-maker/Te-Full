VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060201_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸專利公報資料維護"
   ClientHeight    =   3948
   ClientLeft      =   48
   ClientTop       =   2328
   ClientWidth     =   4824
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3948
   ScaleWidth      =   4824
   Begin VB.TextBox text11 
      Height          =   270
      Left            =   1728
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2592
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton buttonCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   3468
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   2640
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox text01 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   2892
   End
   Begin VB.TextBox text02 
      Height          =   270
      Left            =   1740
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1644
      Width           =   1092
   End
   Begin VB.TextBox text03 
      Height          =   270
      Left            =   1740
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1008
      Width           =   1092
   End
   Begin VB.TextBox text04 
      Height          =   270
      Left            =   1740
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1332
      Width           =   852
   End
   Begin VB.TextBox text05 
      Height          =   270
      Left            =   3300
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1332
      Width           =   852
   End
   Begin VB.TextBox text06_1 
      Height          =   270
      Left            =   1740
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1956
      Width           =   852
   End
   Begin VB.TextBox text07 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2892
   End
   Begin MSForms.ComboBox text10 
      Height          =   300
      Left            =   1740
      TabIndex        =   8
      Top             =   3525
      Width           =   2895
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "5106;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text08 
      Height          =   300
      Left            =   1740
      TabIndex        =   7
      Top             =   3195
      Width           =   2895
      VariousPropertyBits=   671107099
      Size            =   "5106;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text09 
      Height          =   300
      Left            =   1740
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
      VariousPropertyBits=   671107099
      Size            =   "5106;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text06_2 
      Height          =   300
      Left            =   2640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1935
      Width           =   1995
      VariousPropertyBits=   671105055
      Size            =   "3519;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "是否列印通知函  :              (N:不印)"
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   2595
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label11 
      Caption         =   "申請人地址  :"
      Height          =   255
      Left            =   180
      TabIndex        =   24
      Top             =   3555
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   180
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "公告號 :"
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "公告日 :"
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "公報 :"
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "代理事務所 :"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "本所案號 :"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   180
      TabIndex        =   17
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "申請人  :"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "卷"
      Height          =   255
      Left            =   2700
      TabIndex        =   15
      Top             =   1350
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "號"
      Height          =   255
      Left            =   4380
      TabIndex        =   14
      Top             =   1350
      Width           =   255
   End
End
Attribute VB_Name = "frm04060201_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/24 改成Form2.0 (text06_2,text09,text08,text10)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'Memo by Morgan 2008/11/18
'國內案公告不必控管是否有國外案未發文，若要管制則將於相同案之分案、發文、提申也加入國內案檢查

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim m_EditMode As Integer
Dim m_DataKey As String
Dim m_LastDate As String
Dim m_CurrNumber As String
Dim m_CurrCPB02 As String
Dim m_CurrCPB03 As String
Dim m_CurrCPB04 As String
Dim m_CurrCPB05 As String
'Add by Morgan 2004/2/7
'新增性質=通知公告(1208)的進度檔用
Dim stCP(1 To 14) As String
'add by nickc 2005/06/17
Dim m_HaveHK As Boolean
Dim m_HaveHKInCP As String
Dim m_HaveHKInNP As String
Dim m_SendHKMail As Boolean
Dim m_HK_CP01 As String
Dim m_HK_CP02 As String
Dim m_HK_CP03 As String
Dim m_HK_CP04 As String
Dim cm(7) As String
Dim m_HKMailID As String
Dim i As Integer



' 清除欄位內的資料
Public Sub Clear()
   text01 = Empty
   text02 = Empty
   text03 = Empty
   text04 = Empty
   text05 = Empty
   text06_1 = Empty
   text06_2 = Empty
   text07 = Empty
   text08 = Empty
   text09 = Empty
   Text10 = Empty
End Sub

'使用者按下確定的按鍵
Private Sub buttonOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nNo As Long
   Dim strTmp As String
    'Add By Cheng 2003/01/09
    Dim rsA As New ADODB.Recordset
    Dim StrSQLa As String
   
   Select Case m_EditMode
   ' 新增, 修改
   Case 0, 1:
      If CheckDataValid() = True Then
'add by nickc 2005/06/16 取有關大陸的香港關聯判斷
      If m_EditMode = "0" Then
         m_HaveHK = False
         m_HaveHKInCP = ""
         m_HaveHKInNP = ""
         m_SendHKMail = False
         m_HKMailID = ""
         If stCP(9) = "020" Then
            'edit by nickc 2006/05/05
            'm_HaveHK = ChkCMIsExist013(stCP(1), stCP(2), stCP(3), stCP(4))
            m_HaveHK = ChkCMIsExist013(stCP(1), stCP(2), stCP(3), stCP(4), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04)
            If m_HaveHK = True Then
'edit by nickc 2006/05/05
'               For i = 1 To 4
'                  cm(i - 1) = stCP(i)
'               Next
'               If obj003.GetCaseMap(cm, 4) = True Then
'                  m_HK_CP01 = cm(4)
'                  m_HK_CP02 = cm(5)
'                  m_HK_CP03 = cm(6)
'                  m_HK_CP04 = cm(7)
'               End If
               m_HaveHKInCP = Chk013Have111(stCP(1), stCP(2), stCP(3), stCP(4), m_HKMailID)
               If m_HaveHKInCP = "" Then 'Added by Morgan 2016/1/12 '沒有CP才抓NP否則m_HKMailID會被清除
                  m_HaveHKInNP = Chk013Have111(stCP(1), stCP(2), stCP(3), stCP(4), m_HKMailID, "NP")
               End If
            End If
         End If
      End If
        'Modify By Cheng 2002/11/06
'         OnWork
        If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
        
'Removed by Morgan 2016/6/17 公告通知已併入證書函(電子化後通知函要配來函)
'         'Add By Cheng 2002/06/27
'         If Text11 = "Y" Then
'            strTmp = Replace(Me.text07.Text, "-", "")
'            'Modify By Cheng 2003/01/09
'            '以收文號傳入列印定稿
''            NowPrint strTmp & "&000", "10", "02", False, strUserNum, 0
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            StrSQLa = "Select CP09 From CaseProgress WHERE " & ChgCaseprogress(strTmp) & " AND CP09 <'B' And CP05 IS NOT NULL AND CP09 IS NOT NULL ORDER BY CP05 DESC, CP09 DESC "
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'                'edit by nickc 2005/06/17
'                'NowPrint "" & rsA.Fields(0).Value, "10", "02", False, strUserNum, 0
'                  If m_HaveHK = False Then
'                     NowPrint "" & rsA.Fields(0).Value, "10", "02", False, strUserNum, 0
'                  Else
'                     If m_HaveHKInCP <> "" Then
'                        NowPrint "" & rsA.Fields(0).Value, "10", "08", False, strUserNum, 0
'                     Else
'                        NowPrint "" & rsA.Fields(0).Value, "10", "02", False, strUserNum, 0
'                        NowPrint "" & rsA.Fields(0).Value, "08", "12", False, strUserNum, 0
'                     End If
'                  End If
'            End If
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'end 2016/6/17

            frm04060201_1.AddCPB01 strTmp
            'add by nickc 2005/06/17 發 mail
            If m_SendHKMail = True And m_HKMailID <> "" And m_HaveHKInCP <> "" Then
               Call PUB_SendMail(strUserNum, m_HKMailID, m_HaveHKInCP, "大陸案(" & stCP(1) & "-" & stCP(2) & "-" & stCP(3) & "-" & stCP(4) & ")已公告，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "", "")
            End If
'         End If 'Removed by Morgan 2016/6/17
         
         Select Case m_EditMode
            Case 0:
               'nNo = Val(text02.Text)
               'nNo = nNo + 1
               'm_CurrCPB02 = CStr(nNo)
               m_CurrCPB03 = text03
               m_CurrCPB04 = text04
               m_CurrCPB05 = text05
         End Select
         Me.Hide
         frm04060201_1.Show
         frm04060201_1.SetInputCPB01
      End If
   ' 刪除
   Case 3:
      strTit = "詢問"
      strMsg = "是否要刪除此筆資料?"
      If MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit) = vbYes Then
        'Modify By Cheng 2002/11/06
'         OnWork
        If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         Me.Hide
         frm04060201_1.Show
         frm04060201_1.SetInputCPB01
      End If
   Case Else:
      Me.Hide
      frm04060201_1.Show
      frm04060201_1.SetInputCPB01
   End Select
   frm04060201_1.UpdateRecord m_DataKey
End Sub

' 使用者按下取消的按鍵
Private Sub buttonCancel_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
       
    'Modify By Cheng 2003/02/26
    '不要詢問直接退出畫面
'   strTit = "詢問"
'   strMsg = "你並未存檔, 確定離開嗎?"
'   nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'   If nResponse = vbYes Then
      Me.Hide
      frm04060201_1.Show
      frm04060201_1.SetInputCPB01
'   End If
End Sub

Public Sub UpdateState()
   Select Case m_EditMode
      Case 2, 3:
         text02.Locked = True
         text03.Locked = True
         text04.Locked = True
         text05.Locked = True
         text06_1.Locked = True
         text08.Locked = True
         text09.Locked = True
         Text10.Locked = True
      Case Else:
         text02.Locked = False
         text03.Locked = False
         text04.Locked = False
         text05.Locked = False
         text06_1.Locked = False
         text08.Locked = False
         text09.Locked = False
         Text10.Locked = False
   End Select
   text01.BackColor = &H8000000F
   text06_2.BackColor = &H8000000F
   text07.BackColor = &H8000000F
End Sub

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nRet As Boolean
   Dim nResponse
   CheckDataValid = False
   
   'Added by Morgan 2021/12/24 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
       Exit Function
   End If
   'end 2021/12/24
   
   If IsEmptyText(text02) = True Then
      strMsg = "請輸入公告號"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(text03) = True Then
      strMsg = "請輸入公告日"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   'Modify By Cheng 2002/10/23
   '卷期可空白
'   'Modify By Cheng 2002/06/27
'   If Me.text04.Enabled Then
'      If IsEmptyText(text04) = True Then
'         strMsg = "請輸入公告卷號"
'         strTit = "資料檢核"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         GoTo ExitSub
'      End If
'   End If
'   If Me.text05.Enabled Then
'      If IsEmptyText(text05) = True Then
'         strMsg = "請輸入公告期數"
'         strTit = "資料檢核"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         GoTo ExitSub
'      End If
'   End If
   If IsEmptyText(text06_1) = True Then
      strMsg = "請輸入代理事務所"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 案件名稱有輸入, 其餘未輸入
   If IsEmptyText(text08) = False Then
      If IsEmptyText(text09) = True Then
         strMsg = "請輸入申請人"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If IsEmptyText(Text10) = True Then
         strMsg = "請輸入申請人地址"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 申請人有輸入, 其餘未輸入
   If IsEmptyText(text09) = False Then
      If IsEmptyText(text08) = True Then
         strMsg = "請輸入案件名稱"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If IsEmptyText(Text10) = True Then
         strMsg = "請輸入申請人地址"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 申請人地址有輸入, 其餘未輸入
   If IsEmptyText(Text10) = False Then
      If IsEmptyText(text08) = True Then
         strMsg = "請輸入案件名稱"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If IsEmptyText(text09) = True Then
         strMsg = "請輸入申請人"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
      
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub text02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 公告號
Private Sub text02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCode As String
   
   Cancel = False
   If IsEmptyText(text02) = False Then
      If Mid(text02, 1, 2) <> "CN" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公告號不正確"
         nResponse = MsgBox(strMsg, vbCritical + vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      ' 依申請案號的第三碼來判斷
      'Modify by Morgan 2005/4/28 申請號改西元年開頭的改抓第五碼判斷
      'Select Case Mid(text01, 3, 1)
      If Len(text01) > 10 Then
         strCode = Mid(text01, 5, 1)
      Else
         strCode = Mid(text01, 3, 1)
      End If
      Select Case strCode
         Case "1":
            'Modify By Cheng 2003/01/02
'            If Right(text02, 1) <> "A" Then
            If Right(text02, 1) <> "C" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "公告號不正確"
               nResponse = MsgBox(strMsg, vbCritical + vbOKOnly, strTit)
               GoTo EXITSUB
            End If
         Case "2":
            If Right(text02, 1) <> "Y" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "公告號不正確"
               nResponse = MsgBox(strMsg, vbCritical + vbOKOnly, strTit)
               GoTo EXITSUB
            End If
         Case "3":
            If Right(text02, 1) <> "D" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "公告號不正確"
               nResponse = MsgBox(strMsg, vbCritical + vbOKOnly, strTit)
               GoTo EXITSUB
            End If
         Case Else:
      End Select
   End If
EXITSUB:
End Sub

Private Sub text03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text03) = False Then
      If CheckIsTaiwanDate(text03, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
      End If
   End If
End Sub

Private Sub text06_1_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   text06_2 = Empty
   If text06_1 <> Empty Then
      If UpdateCtrlData(1) = False Then
         Cancel = True
         strMsg = "無此事務所資料"
         strTit = "錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text06_1_GotFocus
      End If
   End If
End Sub
Public Sub UpdateData()
   ' 更新 Caption
   Dim strCap As String
   Dim strTmp As String
   
   Clear
   
   strCap = "專利公報資料維護"
   Select Case m_EditMode
      Case 0:
         strTmp = " -- 新增"
         text02 = m_CurrNumber
         text03 = m_LastDate
      Case 1:
         strTmp = " -- 修改"
      Case 2:
         strTmp = " -- 查詢"
      Case 3:
         strTmp = " -- 刪除"
   End Select
   Caption = strCap & strTmp
   ' 更新第一個欄位
   text01 = m_DataKey
   MoveFormToCenter Me
   ' 更新內容
   UpdateCtrlData (0)
   UpdateState
   
   If m_EditMode = 0 Then
      SetPrevState
   End If
End Sub
Public Function UpdateCtrlData(ByVal nAction As Integer) As Boolean
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
         
   UpdateCtrlData = True
   Select Case nAction
      ' 依照申請案號帶出所有相關資料
      Case 0:
         If m_EditMode = 1 Or m_EditMode = 2 Or m_EditMode = 3 Then
            Set rsTmp = New ADODB.Recordset
            strSql = "Select * from CPBulletin where CPB01 = '" & m_DataKey & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("CPB02")) = False Then
                  text02 = rsTmp.Fields("CPB02")
               End If
               If IsNull(rsTmp.Fields("CPB03")) = False Then
                  text03 = ChangeWStringToTString(rsTmp.Fields("CPB03"))
               End If
               If IsNull(rsTmp.Fields("CPB04")) = False Then
                  text04 = rsTmp.Fields("CPB04")
               End If
               If IsNull(rsTmp.Fields("CPB05")) = False Then
                  text05 = rsTmp.Fields("CPB05")
               End If
               If IsNull(rsTmp.Fields("CPB06")) = False Then
                  text06_1 = rsTmp.Fields("CPB06")
               End If
               If IsNull(rsTmp.Fields("CPB07")) = False Then
                  text08 = rsTmp.Fields("CPB07")
               End If
               If IsNull(rsTmp.Fields("CPB08")) = False Then
                  text09 = rsTmp.Fields("CPB08")
                  text09_Validate False
               End If
               If IsNull(rsTmp.Fields("CPB09")) = False Then
                  Text10 = rsTmp.Fields("CPB09")
               End If
               If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(1)
               End If
               If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(2)
               End If
            Else
               UpdateCtrlData = False
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         End If
         If m_EditMode = 0 Then
            UpdateCtrlData = UpdateCtrlData(2)
         End If
      ' 依照大陸事務所代號帶出事務所名稱
      Case 1:
         Set rsTmp = New ADODB.Recordset
         strSql = "SELECT * FROM CAgent WHERE FNM01 = '" & text06_1 & "'"
         text06_2 = Empty
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            text06_2 = rsTmp.Fields("FNM02")
         Else
            UpdateCtrlData = False
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      ' 依照申請案號帶出本所案號
      Case 2:
         Set rsTmp = New ADODB.Recordset
         strSql = "SELECT * FROM Patent " & _
                  "WHERE PA11 = '" & m_DataKey & "' AND " & _
                  "PA09 = '020'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            text07 = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
            'Add by Morgan 2004/2/7
            stCP(1) = rsTmp.Fields("PA01")
            stCP(2) = rsTmp.Fields("PA02")
            stCP(3) = rsTmp.Fields("PA03")
            stCP(4) = rsTmp.Fields("PA04")
            'add by nickc 2005/06/17
            stCP(9) = rsTmp.Fields("pa09")
            'stCP(14) = rsTmp.Fields("pa14")
         Else
            'Add by Morgan 2004/2/7
            stCP(1) = "": stCP(2) = "": stCP(3) = "": stCP(4) = ""
            'add by nickc 2005/06/17
            stCP(9) = "": stCP(14) = ""
            UpdateCtrlData = False
         End If
         rsTmp.Close
         Set rsTmp = Nothing
   End Select
   Set rsTmp = Nothing
End Function

' 此模組在處理資料到資料庫的工作
'Modify By Cheng 2002/11/06
'Public Sub OnWork()
Public Function OnWork() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCPB03 As String
'Add by Morgan 2004/2/7
Dim stCP09 As String, stCP12 As String, stCP13 As String
Dim strPromoteDate As String '2010/1/19 add by sonia
Dim m_bolFMP As Boolean, stNP23 As String  'Added by Lydia 2025/10/29

'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnWork = True
cnnConnection.BeginTrans
   
   m_bolFMP = PUB_ChkIsFMP(stCP(1), stCP(2), stCP(3), stCP(4), stCP(9)) 'Added by Lydia 2025/10/29
   
   strCPB03 = Empty
   If IsEmpty(text03) = False Then
      strCPB03 = ChangeTStringToWString(text03)
   End If
   Select Case m_EditMode
      ' 新增資料到國內專利公報檔
      Case 0:
         If strCPB03 <> Empty Then
            strSql = "INSERT INTO CPBulletin " & _
                     "(CPB01, CPB02, CPB03, CPB04, CPB05, CPB06, CPB07, CPB08, CPB09) " & _
                     "VALUES ('" & text01 & "','" & text02 & "'," & strCPB03 & ",'" & text04 & "','" & _
                     text05 & "','" & text06_1 & "','" & text08 & "','" & text09 & "','" & Text10 & "')"
         Else
            strSql = "INSERT INTO CPBulletin " & _
                     "(CPB01, CPB02, CPB04, CPB05, CPB06, CPB07, CPB08, CPB09) " & _
                     "VALUES ('" & text01 & "','" & text02 & "','" & text04 & "','" & _
                     text05 & "','" & text06_1 & "','" & text08 & "','" & text09 & "','" & Text10 & "')"
         End If
         
         cnnConnection.Execute strSql
         
         ' 更新專利基本檔的公告日及公告號
         Select Case Mid(m_DataKey, 3, 1)
            Case "1":
'               strSQL = "UPDATE Patent SET PA12 = " & ChangeTStringToWString(text03) & ", " & _
'                                          "PA13 = '" & text02 & "' " & _
'                        "WHERE PA11 = '" & m_DataKey & "'"
               strSql = "UPDATE Patent SET PA12 = " & ChangeTStringToWString(text03) & ", " & _
                                          "PA15 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & m_DataKey & "'" & " And PA09 = '020'"
               cnnConnection.Execute strSql
            Case Else:
'               strSQL = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                                          "PA15 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & m_DataKey & "'"
               strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                                          "PA15 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & m_DataKey & "'" & " And PA09 = '020'"
               cnnConnection.Execute strSql
         End Select
         
         ' 檢查大陸公報開拓客戶檔(不存在此申請人及地址則新增一筆)
         If IsEmpty(text09) = False And IsEmpty(Text10) = False Then
            strSql = "SELECT * FROM CBECustom " & _
                     "WHERE CBEC01 = '" & text09 & "' AND " & _
                           "CBEC02 = '" & Text10 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               strSql = "INSERT INTO CBECustom (CBEC01, CBEC02) " & _
                        "VALUES ('" & text09 & "','" & Text10 & "')"
               cnnConnection.Execute strSql
            End If
            rsTmp.Close
         End If
         
         'Add by Morgan 2004/2/3
         '若為本所案件時，即已申請案號抓Patent，有抓到時新增該筆案號之案件進度檔
         'CP05=CP27=系統日,CP09='CXXX',CP10='1208',CP12,CP13,CP14=USERNUM,CP20=CP26=CP32='N', 其他欄位=NULL
         If text07 <> "" Then
            stCP13 = PUB_GetAKindSalesNo(stCP(1), stCP(2), stCP(3), stCP(4))
            stCP12 = GetSalesArea(stCP13)
            stCP09 = AutoNo("C", 6)
            strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32)" & _
                    " VALUES ('" & stCP(1) & "','" & stCP(2) & "','" & stCP(3) & "','" & stCP(4) & "','" & strSrvDate(1) & "','" & stCP09 & "'" & _
                    ",'1208','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N','" & strSrvDate(1) & "','N')"
           cnnConnection.Execute strSql
         End If
         'Add End 2004/2/3
   
   
   'add by nickc 2005/06/17 大陸香港相關案
   If stCP(9) = "020" Then
   Dim tmpCp06 As String
   Dim tmpCp07 As String
   Dim strProgressNo As String
   '檢查有無香港
      If m_HaveHK = True Then
            tmpCp07 = PUB_GetWorkDay1(CompDate(1, 6, ChangeTStringToWString(text03)), True)
            'Added by Lydia 2025/10/29
            stNP23 = ""
            If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
               tmpCp06 = PUB_GetPOurDeadline(tmpCp07, stCP(9), stNP23, stCP(1), "111")
            Else
            'end 2025/10/29
               tmpCp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, tmpCp07)), True)
            End If 'Added by Lydia 2025/10/29
          '檢查有無收香港的 111
          If m_HaveHKInCP <> "" Then
               '更新期限，上發 mail tag
               strSql = "Update CaseProgress Set CP06=" & tmpCp06 & ",CP07=" & tmpCp07 & " Where CP09='" & m_HaveHKInCP & "' "
               cnnConnection.Execute strSql
               '更新齊備日
               strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & m_HaveHKInCP & "' "
               cnnConnection.Execute strSql
               m_SendHKMail = True
               
               If PUB_IfSetCP48(m_HaveHKInCP) Then 'Add by Morgan 2010/10/6
               
                  '2010/1/19 add by sonia 更新承辦期限
                  strPromoteDate = Pub_GetHandleDay(stCP(1), "013", "111", , tmpCp06)
                  If strPromoteDate <> "" Then
                     strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & m_HaveHKInCP & "' "
                     cnnConnection.Execute strSql
                  End If
                  '2010/1/19 end
                  
               End If 'Add by Morgan 2010/10/6
          Else
               '檢查有無 np 的香港 111
               If m_HaveHKInNP <> "" Then
                  '更新期限
                  'Modified by Lydia 2025/10/29 +NP23
                  'strSql = "Update nextProgress Set nP08=" & tmpCp06 & ",nP09=" & tmpCp07 & " Where np01='" & m_HaveHKInNP & "' "
                  strSql = "Update nextProgress Set nP08=" & tmpCp06 & ",nP09=" & tmpCp07 & ",NP23=" & IIf(Trim(stNP23) = "", "NP23", stNP23) & " Where np01='" & m_HaveHKInNP & "' AND NP07='111' "
                  cnnConnection.Execute strSql
               Else
                  '新增 np 期限
                  strProgressNo = GetNextProgressNo
                  'Modified by Lydia 2025/10/29 +NP23
                  strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09," & _
                      "NP10,NP22,NP23) select max(cp09),cp01,cp02,cp03,cp04,111," & tmpCp06 & "," & tmpCp07 & ",'" & _
                      PUB_GetAKindSalesNo(m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04) & "','" & strProgressNo & "'," & CNULL(stNP23, True) & _
                      " from caseprogress where cp01='" & m_HK_CP01 & "' and cp02='" & m_HK_CP02 & "' and cp03='" & m_HK_CP03 & _
                      "' and cp04='" & m_HK_CP04 & "' and cp09<'C' group by cp01,cp02,cp03,cp04 "
                  cnnConnection.Execute strSql
               End If
          End If
      End If
   End If
         
      ' 更新專利公報檔的資料
      Case 1:
         If strCPB03 <> Empty Then
            strSql = "UPDATE CPBulletin " & _
                     "SET CPB01='" & text01 & "'," & "CPB02='" & text02 & "'," & "CPB03=" & strCPB03 & "," & "CPB04='" & text04 & "'," & _
                     "CPB05='" & text05 & "'," & "CPB06='" & text06_1 & "'," & "CPB07='" & text08 & "'," & "CPB08='" & text09 & "'," & _
                     "CPB09='" & Text10 & "' " & _
                     "WHERE CPB01='" & text01 & "'"
         Else
            strSql = "UPDATE CPBulletin " & _
                  "SET CPB01='" & text01 & "'," & "CPB02='" & text02 & "'," & "CPB04='" & text04 & "'," & _
                  "CPB05='" & text05 & "'," & "CPB06='" & text06_1 & "'," & "CPB07='" & text08 & "'," & "CPB08='" & text09 & "'," & _
                  "CPB09='" & Text10 & "' " & _
                  "WHERE CPB01='" & text01 & "'"
         End If
         cnnConnection.Execute strSql
         
         ' 更新專利基本檔的公告日及公告號
         Select Case Mid(m_DataKey, 3, 1)
            Case "1":
'               strSQL = "UPDATE Patent SET PA12 = " & ChangeTStringToWString(text03) & ", " & _
'                                          "PA13 = '" & text02 & "' " & _
'                        "WHERE PA11 = '" & m_DataKey & "'"
               strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                                          "PA15 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & m_DataKey & "'" & " And PA09 = '" & "000" & "'"
               cnnConnection.Execute strSql
            Case Else:
'               strSQL = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
'                                          "PA15 = '" & text02 & "' " & _
'                        "WHERE PA11 = '" & m_DataKey & "'"
               strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                                          "PA15 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & m_DataKey & "'" & " And PA09 = '" & "000" & "'"
               cnnConnection.Execute strSql
         End Select
         
         ' 檢查大陸公報開拓客戶檔(不存在此申請人及地址則新增一筆)
         If IsEmpty(text09) = False And IsEmpty(Text10) = False Then
            strSql = "SELECT * FROM CBECustom " & _
                     "WHERE CBEC01 = '" & text09 & "' AND " & _
                           "CBEC02 = '" & Text10 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               strSql = "INSERT INTO CBECustom (CBEC01, CBEC02) " & _
                        "VALUES ('" & text09 & "','" & Text10 & "')"
               cnnConnection.Execute strSql
            End If
            rsTmp.Close
         End If
      Case 3:
         strSql = "DELETE FROM CPBulletin WHERE CPB01 = '" & m_DataKey & "'"
         cnnConnection.Execute strSql
   End Select
'Add By Cheng 2002/11/06
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnWork = False
End Function

' 設定編輯資料的模式 (新增或修改)
Public Sub SetMode(ByVal nMode As Integer)
   Select Case nMode
      Case 0, 1, 2, 3:
         m_EditMode = nMode
      Case Else:
   End Select
   'Add By Cheng 2002/06/27
   '若為新增狀態
   If m_EditMode = 0 Or m_EditMode = 1 Then
      If Mid(text01.Text, 3, 1) = "1" Or Mid(text01.Text, 3, 1) = "2" Then
         Me.text04.Enabled = False
         Me.text05.Enabled = False
      ElseIf Mid(text01.Text, 3, 1) = "3" Then
         Me.Text11.Text = "N"
      End If
   End If
   '若為修改狀況
   If m_EditMode = 1 Then
      Me.Text11.Text = "N"
   End If
   
End Sub

' 設定控制項中初始的值 (申請案號的值)
Public Sub SetData(ByVal textKey As String)
   m_DataKey = textKey
   If m_EditMode = 0 Then
      If ChkPatent(m_DataKey) Then Text11 = "Y"
   Else
      Text11 = "N"
   End If
End Sub

Private Function ChkPatent(PA11 As String) As Boolean
   ChkPatent = False
   strExc(0) = "SELECT COUNT(*) FROM PATENT WHERE PA11='" & PA11 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If RsTemp.Fields(0) > 0 Then ChkPatent = True
End Function

Private Sub SetPrevState()
   text02 = m_CurrCPB02
   'Modify By Cheng 2002/06/27
'   text03 = m_CurrCPB03
'   text04 = m_CurrCPB04
'   text05 = m_CurrCPB05
'   If text02 = Empty Then
'      text02.SetFocus
'   ElseIf text03 = Empty Then
'      text03.SetFocus
'   ElseIf text04 = Empty Then
'      text04.SetFocus
'   ElseIf text05 = Empty Then
'      text05.SetFocus
'   Else
'      If text06_1.Locked = False Then
'         text06_1.SetFocus
'      End If
'   End If
   'Add By Cheng 2002/06/26
   If Me.text03.Enabled Then Me.text03.SetFocus

End Sub

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As Object)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text02_GotFocus()
    'Modify By Cheng 2003/02/14
'    InverseAll text02
    Me.text02.SelStart = 2
    Me.text02.SelLength = Len(Me.text02.Text) - 2
End Sub

Private Sub text03_GotFocus()
   InverseAll text03
End Sub

Private Sub text04_GotFocus()
   InverseAll text04
End Sub

Private Sub text05_GotFocus()
   InverseAll text05
End Sub

Private Sub text06_1_GotFocus()
   InverseAll text06_1
End Sub

Private Sub text08_GotFocus()
   InverseAll text08
   'edit by nickc 2007/07/11 切換輸入法改用API
   'text08.IMEMode = 1
   OpenIme
End Sub

Private Sub text08_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   If StrLength(text08) > 80 Then
      Cancel = True
      strMsg = "案件名稱太長"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: text08.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub text09_GotFocus()
   InverseAll text09
   'edit by nickc 2007/07/11 切換輸入法改用API
   'text09.IMEMode = 1
   OpenIme
End Sub

Private Sub text09_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   
   Text10.Clear
   strSql = "SELECT * FROM CBECustom " & _
            "WHERE CBEC01 = '" & text09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("CBEC02")) = False Then
            If IsEmpty(rsTmp.Fields("CBEC02")) = False Then
               Text10.AddItem rsTmp.Fields("CBEC02")
            End If
         End If
         rsTmp.MoveNext
      Loop
      
      If Text10.ListCount > 0 Then: Text10.ListIndex = 0
   End If
   rsTmp.Close
      
   If StrLength(text09) > 40 Then
      Cancel = True
      strMsg = "申請人太長"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
      
   Set rsTmp = Nothing
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: text09.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub


Private Sub Text10_GotFocus()
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.Text)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'text10.IMEMode = 1
   OpenIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = vbKeyDown Then
      'Modified by Morgan 2021/12/24 Form2.0無hWnd屬性,可改直接用增加的 DropDown 方法
      'SendMessage text10.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
      Text10.DropDown
      'end 2021/12/24
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If StrLength(Text10) > 70 Then
      Cancel = True
      strMsg = "申請人地址太長"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: text10.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 8 Then KeyAscii = 0
   If KeyAscii = 89 Then
      If Not ChkPatent(m_DataKey) Then
         MsgBox "本所案件才可列印通知函 !", vbCritical
         KeyAscii = 89
      End If
   End If
End Sub



