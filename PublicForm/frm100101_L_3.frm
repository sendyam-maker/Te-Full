VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_L_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "搬移檔案"
   ClientHeight    =   5970
   ClientLeft      =   50
   ClientTop       =   300
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8970
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   6230
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   7220
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2280
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   7610
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2280
      Width           =   492
   End
   Begin VB.TextBox txtSystem 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5490
      MaxLength       =   3
      TabIndex        =   2
      Top             =   2280
      Width           =   732
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   8130
      Picture         =   "frm100101_L_3.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2280
      Width           =   350
   End
   Begin VB.ListBox lstAtt 
      Height          =   940
      ItemData        =   "frm100101_L_3.frx":0102
      Left            =   1140
      List            =   "frm100101_L_3.frx":0109
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   7730
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   6930
      TabIndex        =   10
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   11
      Top             =   90
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3110
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   2810
      Width           =   4400
      _ExtentX        =   7761
      _ExtentY        =   5486
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V  |  總收文號  |  收文日  |  案件性質  | 發文日"
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
      _Band(0).Cols   =   5
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   1140
      TabIndex        =   18
      Top             =   2010
      Width           =   7455
      Begin VB.OptionButton Option1 
         Caption         =   "其他案件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   3390
         TabIndex        =   1
         Top             =   30
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本案件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   30
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3110
      Index           =   1
      Left            =   4500
      TabIndex        =   9
      Top             =   2810
      Width           =   4400
      _ExtentX        =   7761
      _ExtentY        =   5486
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V  |  總收文號  |  收文日  |  案件性質  | 發文日"
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
      _Band(0).Cols   =   5
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   8920
      Y1              =   1720
      Y2              =   1720
   End
   Begin VB.Label Label1 
      Caption         =   "欲移動的檔案："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Index           =   5
      Left            =   180
      TabIndex        =   22
      Top             =   780
      Width           =   890
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      ForeColor       =   &H00000000&
      Height          =   230
      Index           =   4
      Left            =   4530
      TabIndex        =   21
      Top             =   2310
      Width           =   920
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   3
      Left            =   4530
      TabIndex        =   20
      Top             =   2580
      Width           =   960
   End
   Begin MSForms.Label lblCaseName2 
      Height          =   290
      Left            =   5510
      TabIndex        =   19
      Top             =   2580
      Width           =   3350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5909;512"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "請點選上列電子檔欲移至那一個案號 並且 勾選該案號的某一道程序"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   2
      Left            =   1290
      TabIndex        =   17
      Top             =   1770
      Width           =   6660
   End
   Begin VB.Label Label1 
      Caption         =   "移動方式："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   16
      Top             =   2070
      Width           =   950
   End
   Begin MSForms.Label lblCaseName 
      Height          =   290
      Left            =   1130
      TabIndex        =   15
      Top             =   390
      Width           =   5720
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10089;512"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   240
      Index           =   18
      Left            =   150
      TabIndex        =   14
      Top             =   420
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblCaseNo 
      Caption         =   "lblCaseNo"
      Height          =   225
      Left            =   1140
      TabIndex        =   12
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frm100101_L_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/23 Form2.0已修改
'Create By Sindy 2015/5/25
Option Explicit

Public m_strSaveFiles As String
Public strRecvNo As String
Public m_Nation As String
Dim ii As Integer, jj As Integer
'Private Declare Function SendMessageByNum Lib "user32" ()
''  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
'  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim m_PrevForm As Form '前一畫面
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_mouseRow As Long


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim intChkCnt As Integer
Dim strNewCP09 As String
Dim strNewCP10 As String
Dim strCPP01 As String
Dim strCPP02 As String
Dim strOldCP10 As String
Dim strNewCPP02 As String, strOldCPP02 As String
Dim strTemp As String
Dim arrData
Dim adoRst As ADODB.Recordset
Dim intGrd_idx As Integer 'Add By Sindy 2017/5/10
Dim strOldCPP06 As String 'Add By Sindy 2019/2/19
Dim strOldCPP07 As String 'Add By Sindy 2019/2/19
Dim bolConn As Boolean 'Add By Sindy 2019/2/19
'Added by Lydia 2019/03/06
Dim strOldCP01 As String, strOldCP02 As String, strOldCP03 As String, strOldCP04 As String '移檔前的本所案號
Dim strNewCP01 As String, strNewCP02 As String, strNewCP03 As String, strNewCP04 As String '移檔後的本所案號

On Error GoTo ErrHnd
      
   If Option1(0).Value = True Then '本案件
      intGrd_idx = 0
      'Added by Lydia 2019/03/06
      strNewCP01 = SystemNumber(lblCaseNo, 1)
      strNewCP02 = SystemNumber(lblCaseNo, 2)
      strNewCP03 = SystemNumber(lblCaseNo, 3)
      strNewCP04 = SystemNumber(lblCaseNo, 4)
   Else '其他案件
      intGrd_idx = 1
      'Added by Lydia 2019/03/06
      strNewCP01 = txtSystem
      strNewCP02 = txtCode(0)
      strNewCP03 = txtCode(1)
      strNewCP04 = txtCode(2)
   End If
   
   '確定
   If Index = 0 Then
      intChkCnt = 0
      'Added by Lydia 2019/03/06
      strOldCP01 = SystemNumber(lblCaseNo, 1)
      strOldCP02 = SystemNumber(lblCaseNo, 2)
      strOldCP03 = SystemNumber(lblCaseNo, 3)
      strOldCP04 = SystemNumber(lblCaseNo, 4)
      'end 2019/03/06
      
      For ii = 1 To GRD1(intGrd_idx).Rows - 1
         GRD1(intGrd_idx).row = ii
         GRD1(intGrd_idx).col = 1
         If GRD1(intGrd_idx).CellBackColor = &HFFC0C0 Then
            intChkCnt = intChkCnt + 1
            Exit For
         End If
      Next ii
      If intChkCnt = 0 Then
         MsgBox "請勾選要移至那一道程序！"
         Exit Sub
      Else
         If MsgBox("確定要移動畫面上的電子檔嗎？", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
      
      For ii = 1 To GRD1(intGrd_idx).Rows - 1
         GRD1(intGrd_idx).row = ii
         GRD1(intGrd_idx).col = 1
         If GRD1(intGrd_idx).CellBackColor = &HFFC0C0 Then
            '新文號
            strNewCP09 = Trim(GRD1(intGrd_idx).TextMatrix(ii, 1))
            strNewCP10 = Trim(GRD1(intGrd_idx).TextMatrix(ii, 5))
            For jj = 0 To lstAtt.ListCount - 1
               strTemp = lstAtt.List(jj)
               arrData = Split(strTemp, "  ")
               If UBound(arrData) <> 1 Then
                  MsgBox "欲搬移的電子檔資料有誤！"
                  Exit Sub
               End If
               
               Screen.MousePointer = vbHourglass
               
               '舊文號
               strCPP01 = Trim(arrData(0))
               strCPP02 = Trim(arrData(1)): strOldCPP02 = Trim(arrData(1))
               '取得舊案件性質
               'Modified by Lydia 2018/02/02 FCP客戶提供文件處理後，D類收文會刪除
               'strSql = "SELECT cp10 FROM caseprogress WHERE cp09='" & strCPP01 & "'"
               strSql = "SELECT 1 as ord1, cp10 FROM caseprogress WHERE cp09='" & strCPP01 & "'"
               If Left(strCPP01, 1) = "D" Then
                   strSql = strSql & " Union select 2 ord1 , '1920' as cp10 from (select * from custsupportdoc where csd05='" & strCPP01 & "' and nvl(csd11,0)>0 ) "
               End If
               strSql = strSql & " order by ord1"
               'end 2018/02/02
               intI = 1
               Set adoRst = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strOldCP10 = adoRst.Fields("cp10")
               End If
               'Modify By Sindy 2016/4/18 + And strOldCP10 <> ""
'               If InStr(strCPP02, "." & strOldCP10) > 0 And strOldCP10 <> "" Then
'                  strNewCPP02 = Replace(strCPP02, "." & strOldCP10, "." & strCP10)
'               Else
'                  If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, strCP10, strCPP02, strNewCPP02, True, 1) = False Then Exit Sub
'               End If
               'Modify By Sindy 2017/5/11
               If InStr(strCPP02, "." & strOldCP10) > 0 And strOldCP10 <> "" Then
                  'Modify By Sindy 2018/2/1
                  'strCPP02 = Replace(strCPP02, "." & strOldCP10, "." & strNewCP10)
                  strCPP02 = Replace(strCPP02, "." & strOldCP10 & ".", "." & strNewCP10 & ".")
               End If
               If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, strNewCP10, strCPP02, strNewCPP02, True, 1) = False Then Exit Sub
               '2017/5/11 END
               'Add By Sindy 2019/2/19
               If IsRecordExist(strCPP01, strOldCPP02, strOldCPP06, strOldCPP07, False) = False Then
                  Screen.MousePointer = vbDefault
                  MsgBox strOldCPP02 & " 此檔案不存在,無法移檔！"
                  Exit Sub
               End If
               '2019/2/19 END
               '更新資料
               If IsRecordExist(strNewCP09, strNewCPP02) = False Then
                  'Modify By Sindy 2016/4/18 + ,cpp10='X'
                  'Modify By Sindy 2018/3/8 取消,cpp10='X' ex.P-113775,AA7008005,*.msg
                  'Modify By Sindy 2018/5/9 + ,cpp10=decode(cpp10,'U','X',cpp10) ex.因為若為本所案號的回覆單歸入文號要將CPP10='U'改為CPP10='X'
                  'Add By Sindy 2019/2/19 增加判斷寄件備份的儲存
                  If UCase(Right(strOldCPP02, Len(".Email.menu"))) = UCase(".Email.menu") Then
                     cnnConnection.BeginTrans: bolConn = True
                     strSql = "update casepaperpdf set cpp01='" & strNewCP09 & "',cpp02='" & strNewCPP02 & "' where cpp01='" & strCPP01 & "' and upper(cpp02)='" & UCase(strOldCPP02) & "'"
                     Pub_SeekTbLog strSql 'Add By Sindy 2022/1/14
                     cnnConnection.Execute strSql
                     '日期要更新跟寄件備份同日期,時間才能查看
                     strSql = "update casepaperpdf set cpp06='" & strOldCPP06 & "',cpp07='" & strOldCPP07 & "' where cpp01='" & strNewCP09 & "' and upper(cpp02)='" & UCase(strNewCPP02) & "'"
                     Pub_SeekTbLog strSql 'Add By Sindy 2022/1/14
                     cnnConnection.Execute strSql
                     strSql = "update smailbackup set smb01='" & strNewCP09 & "' where smb01='" & strCPP01 & "' and smb02=" & strOldCPP06 & " and smb03=" & strOldCPP07
                     Pub_SeekTbLog strSql 'Add By Sindy 2022/1/14
                     cnnConnection.Execute strSql
                     cnnConnection.CommitTrans: bolConn = False
                  Else
                  '2019/2/19 END
                     strSql = "update casepaperpdf set cpp01='" & strNewCP09 & "',cpp02='" & strNewCPP02 & "',cpp10=decode(cpp10,'U','X','C','X',cpp10) where cpp01='" & strCPP01 & "' and upper(cpp02)='" & UCase(strOldCPP02) & "'"
                     Pub_SeekTbLog strSql 'Add By Sindy 2022/1/14
                     cnnConnection.Execute strSql
                  End If
                  'Add By Sindy 2017/6 暫時先不考慮回寫國外部收件夾資料裡的總收文號
'                  strSql = "select ii01,ii02,ii03,ii19 from ipdeptinput" & _
'                           " where ii19='" & strCPP01 & "'" & _
'                           " and instr('" & strOldCPP02 & "',ii01||substr('000000'||ii02,-6)||'.'||ii03)>0"
'                  intI = 1
'                  Set adoRst = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     If adoRst.RecordCount = 1 Then
'                        strSql = "update ipdeptinput set ii19='" & strNewCP09 & "'" & _
'                                 " where ii01=" & adoRst.Fields("ii01") & _
'                                 " and ii02=" & adoRst.Fields("ii02") & _
'                                 " and ii03='" & adoRst.Fields("ii03") & "'"
'                        cnnConnection.Execute strSql
'                     End If
'                  End If
                  
                  'Added by Lydia 2019/03/06 FCP之公告公報1228增加判斷是否有公告本
                  If (strOldCP01 = "FCP" And strOldCP10 = "1228" And InStr(UCase(strOldCPP02), ".GAZ.PDF") > 0) Or _
                         (strNewCP01 = "FCP" And strNewCP10 = "1228" And InStr(UCase(strNewCPP02), ".GAZ.PDF") > 0) Then
                      Call UpdateCP121(strCPP01, "1228", "GAZ")
                      Call UpdateCP121(strNewCP09, "1228", "GAZ")
                  End If
                  'end 2019/03/06
               End If
            Next jj
            Call m_PrevForm.ReadAttachFile
            Exit For
         End If
      Next ii
'      If intChkCnt = 0 Then
'         MsgBox "請勾選要移至那一道程序！"
'         Exit Sub
'      End If
   End If

   Set adoRst = Nothing
   Unload Me
   Screen.MousePointer = vbDefault

   Exit Sub

ErrHnd:
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   Screen.MousePointer = vbDefault
   Set adoRst = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCP09 As String, ByVal strFileName As String, _
   Optional ByRef strCPP06 As String, Optional ByRef strCPP07 As String, _
   Optional ByVal bolShowMsg As Boolean = True) As Boolean
Dim adoRst As ADODB.Recordset

   IsRecordExist = False
   
   'Modify By Sindy 2019/2/19 + ,cpp06,cpp07
   strSql = "SELECT cpp01,cpp06,cpp07 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and upper(cpp02)=upper('" & strFileName & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCPP06 = "" & adoRst.Fields("cpp06") 'Add By Sindy 2019/2/19
      strCPP07 = "" & adoRst.Fields("cpp07") 'Add By Sindy 2019/2/19
      IsRecordExist = True
      If bolShowMsg = True Then
         MsgBox "文號：" & strCP09 & " 附件：" & strFileName & " 已存在！"
      End If
   End If

   Set adoRst = Nothing
End Function

Private Sub Command5_Click()
   If txtSystem <> "" And txtCode(0) <> "" Then
      If txtCode(1) = "" Then txtCode(1) = "0"
      If txtCode(2) = "" Then txtCode(2) = "00"
   End If
   
   If txtSystem = "" Then
      MsgBox "系統別不可空白!", vbExclamation
      If txtSystem.Enabled = True Then txtSystem.SetFocus
      Exit Sub
   End If
   If txtCode(0) = "" Then
      MsgBox "案號不可空白!", vbExclamation
      If txtCode(0).Enabled = True Then txtCode(0).SetFocus
      Exit Sub
   End If
   
   If lblCaseNo.Caption = txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2) Then
      MsgBox "輸入的本所案號不可相同!", vbExclamation
      If txtCode(0).Enabled = True Then txtCode(0).SetFocus
      Exit Sub
   End If
   
   Call QueryData(1)
End Sub

Private Sub Form_Load()
Dim sFile

   MoveFormToCenter Me
   
   lstAtt.Clear
   If m_strSaveFiles <> "" Then
      sFile = Split(m_strSaveFiles, "&")
      For ii = 0 To UBound(sFile)
         lstAtt.AddItem sFile(ii), 0
         SetListScroll lstAtt
      Next ii
   End If

   Me.lblCaseNo = m_PrevForm.lblCaseNo
   Me.lblCaseName = m_PrevForm.lblCaseName
   
   Call SetGrd(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm100101_L_3 = Nothing
End Sub

Public Function QueryData(Index As Integer) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   m_mouseRow = 0
   QueryData = False
   '清空及預設欄位值
   GRD1(Index).Clear
   Call SetGrd(Index)
   
   If Index = 0 Then
      m_CP01 = SystemNumber(lblCaseNo, 1)
      m_CP02 = SystemNumber(lblCaseNo, 2)
      m_CP03 = SystemNumber(lblCaseNo, 3)
      m_CP04 = SystemNumber(lblCaseNo, 4)
   Else
      m_CP01 = txtSystem.Text
      m_CP02 = Left(Trim(txtCode(0).Text) & "00000", 6)
      m_CP03 = txtCode(1).Text
      m_CP04 = txtCode(2).Text
      lblCaseName2.Caption = ""
   End If
   
   '檢查使用者權限
   If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04) = False Then
      Exit Function
   End If
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Modify By Sindy 2017/7/26 del : and cp09 not in(" & strRecvNo & ")
   strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp05) as 收文日,NVL(DECODE('" & m_Nation & "','000',CPM03,CPM04),CP10) as 案件性質,sqldatet(cp27) as 發文日,CP10,CP43" & _
            " From caseprogress,casepropertymap" & _
            " Where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)"
   strSql = strSql & " order by CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1(Index).Recordset = rsTmp
      QueryData = True
      GRD1(Index).col = 0
      GRD1(Index).row = 1
      If Index = 1 Then
         lblCaseName2.Caption = GetPrjName(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04)
      End If
      cmdOK(Index).Default = True
      'Add By Sindy 2017/7/24 判斷有相關總收文號才做較快
      For ii = 1 To GRD1(Index).Rows - 1
         If GRD1(Index).TextMatrix(ii, 6) <> "" Then
            GRD1(Index).TextMatrix(ii, 3) = GRD1(Index).TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(GRD1(Index).TextMatrix(ii, 1), "1")
         End If
      Next ii
      '2017/7/24 END
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      If Index = 0 Then
         Unload Me
         Exit Function
      End If
      Screen.MousePointer = vbDefault
   End If
   rsTmp.Close

   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Set rsTmp = Nothing
End Function

Private Sub SetGrd(Index As Integer, Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   '                        0    1           2         3           4         5       6
   arrGridHeadText = Array("V", "總收文號", "收文日", "案件性質", "發文日", "CP10", "CP43")
   arrGridHeadWidth = Array(200, 950, 800, 1250, 800, 0, 0)
   GRD1(Index).Visible = False
   GRD1(Index).Cols = UBound(arrGridHeadText) + 1
   If bolSetRow = True Then
      GRD1(Index).Rows = 2
   End If
   For iRow = 0 To GRD1(Index).Cols - 1
      GRD1(Index).row = 0
      GRD1(Index).col = iRow
      If bolSetRow = True Then
         GRD1(Index).Text = arrGridHeadText(iRow)
      End If
      GRD1(Index).ColWidth(iRow) = arrGridHeadWidth(iRow)
      If bolSetRow = True Then
         GRD1(Index).CellAlignment = flexAlignCenterCenter
      End If
   Next
   GRD1(Index).Visible = True
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long

   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next

   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim oldRow As Long

getGrdColRow GRD1(Index), x, y, nCol, nRow
GRD1(Index).col = nCol
GRD1(Index).row = nRow
oldRow = m_mouseRow

GRD1(Index).Visible = False
If GRD1(Index).MouseRow <> 0 And Trim(GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 1)) <> "" Then
   m_mouseRow = GRD1(Index).MouseRow
   'If oldRow <> m_mouseRow Then
   If oldRow <> m_mouseRow And oldRow <= GRD1(Index).Rows - 1 Then
      GRD1(Index).row = oldRow
      GRD1(Index).col = 1
      If GRD1(Index).CellBackColor = &HFFC0C0 Then
         '清除反白
         GRD1(Index).TextMatrix(oldRow, 0) = ""
         For jj = 1 To GRD1(Index).Cols - 1
            GRD1(Index).col = jj
            GRD1(Index).CellBackColor = QBColor(15)
         Next jj
      End If
   End If

   GRD1(Index).row = GRD1(Index).MouseRow
   GRD1(Index).col = 1
   If GRD1(Index).CellBackColor = &HFFC0C0 Then
      '清除反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = ""
      For jj = 1 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = QBColor(15)
      Next jj
   Else
      '資料列反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = "V"
      For jj = 1 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = &HFFC0C0
      Next jj
   End If
End If
GRD1(Index).Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   '點選其他案件
   If Index = 1 Then
      txtSystem.Enabled = True
      txtCode(0).Enabled = True
      txtCode(1).Enabled = True
      txtCode(2).Enabled = True
      Command5.Enabled = True
      'Call QueryData(0)
      GRD1(0).Enabled = False
      txtSystem.SetFocus
      Command5.Default = True
   '點選本案件
   Else
      Call QueryData(0)
      GRD1(0).Enabled = True
      '其他案件相關物件清空或鎖住
      txtSystem.Enabled = False: txtSystem = ""
      txtCode(0).Enabled = False: txtCode(0) = ""
      txtCode(1).Enabled = False: txtCode(1) = ""
      txtCode(2).Enabled = False: txtCode(2) = ""
      Command5.Enabled = False
      '清空及預設欄位值
      GRD1(1).Clear
      Call SetGrd(1)
      lblCaseName2.Caption = ""
   End If
End Sub

Private Sub txtSystem_GotFocus()
   CloseIme
   
   If txtSystem.Enabled = True Then
      txtSystem.SetFocus
      txtSystem.SelStart = 0
      txtSystem.SelLength = Len(txtSystem)
   End If
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   CloseIme
   
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
