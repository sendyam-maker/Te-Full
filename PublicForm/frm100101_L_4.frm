VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_L_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "複製檔案"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9255
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "新增 →"
      Height          =   400
      Index           =   0
      Left            =   4830
      TabIndex        =   7
      Top             =   3450
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "← 刪除"
      Height          =   400
      Index           =   1
      Left            =   4830
      TabIndex        =   8
      Top             =   4110
      Width           =   870
   End
   Begin VB.ListBox lstAtt 
      Height          =   780
      ItemData        =   "frm100101_L_4.frx":0000
      Left            =   1140
      List            =   "frm100101_L_4.frx":0007
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   8055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   7170
      TabIndex        =   9
      Top             =   120
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   8190
      TabIndex        =   10
      Top             =   120
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "目的地："
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
      Height          =   3855
      Left            =   5490
      TabIndex        =   14
      Top             =   1770
      Width           =   3705
      Begin VB.ListBox List1 
         Height          =   3480
         ItemData        =   "frm100101_L_4.frx":0013
         Left            =   270
         List            =   "frm100101_L_4.frx":0015
         MultiSelect     =   2  '進階多重選取
         TabIndex        =   15
         Top             =   270
         Width           =   3345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "複製到："
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
      Height          =   3855
      Left            =   30
      TabIndex        =   16
      Top             =   1770
      Width           =   5385
      Begin VB.TextBox textCP10 
         Height          =   300
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   5
         Top             =   540
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Default         =   -1  'True
         Height          =   300
         Left            =   3750
         Picture         =   "frm100101_L_4.frx":0017
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   240
         Width           =   350
      End
      Begin VB.TextBox txtSystem 
         Height          =   264
         Left            =   1110
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   732
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   3225
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   2835
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   2295
         Left            =   90
         TabIndex        =   17
         Top             =   1470
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   4048
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
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   180
         Index           =   6
         Left            =   150
         TabIndex        =   25
         Top             =   570
         Width           =   945
      End
      Begin VB.Label LblCP10 
         BackColor       =   &H80000005&
         Height          =   180
         Left            =   1755
         TabIndex        =   24
         Top             =   570
         Width           =   1695
      End
      Begin MSForms.Label lblCaseName2 
         Height          =   525
         Left            =   1140
         TabIndex        =   20
         Top             =   870
         Width           =   4125
         BackColor       =   -2147483643
         Size            =   "7276;926"
         BorderColor     =   -2147483633
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   18
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Label lblRecvNo 
      BackColor       =   &H80000005&
      Caption         =   "lblRecvNo"
      Height          =   225
      Left            =   4620
      TabIndex        =   27
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   1
      Left            =   3630
      TabIndex        =   26
      Top             =   30
      Width           =   945
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   150
      TabIndex        =   23
      Top             =   1290
      Visible         =   0   'False
      Width           =   735
      VariousPropertyBits=   746604571
      Size            =   "1296;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   660
      Left            =   1140
      TabIndex        =   22
      Top             =   300
      Width           =   5895
      BackColor       =   -2147483643
      Size            =   "10398;1164"
      BorderColor     =   -2147483633
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   18
      Left            =   150
      TabIndex        =   21
      Top             =   300
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "檔案來源："
      Height          =   225
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   1020
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Top             =   30
      Width           =   945
   End
   Begin VB.Label lblCaseNo 
      BackColor       =   &H80000005&
      Caption         =   "lblCaseNo"
      Height          =   225
      Left            =   1140
      TabIndex        =   11
      Top             =   30
      Width           =   2085
   End
End
Attribute VB_Name = "frm100101_L_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2021/10/21 Form2.0已修改
Option Explicit

Public m_strSaveFiles As String
Public strRecvNo As String
Public m_CP10 As String
Dim m_Nation As String
Dim ii As Integer, jj As Integer
'Private Declare Function SendMessageByNum Lib "user32" ()
''  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
'  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim m_PrevForm As Form '前一畫面
Dim m_mouseRow As Long
Dim m_AttachPath As String
Dim strOldCP01 As String, strOldCP02 As String, strOldCP03 As String, strOldCP04 As String
Dim m_strUserRight As String '使用者系統類別使用權限
Dim m_arrUserRight '使用者系統類別使用權限陣列
Dim blnUserRight As Boolean '是否有此系統類別權限


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strNewCP09 As String
Dim strNewCP10 As String
Dim strOldCPP01 As String, strOldCPP02 As String
Dim strNewCPP02 As String
Dim strTemp As String
Dim arrData
Dim Cancel As Boolean
Dim strNewCP01 As String, strNewCP02 As String, strNewCP03 As String, strNewCP04 As String
Dim fs, f

On Error GoTo ErrHnd
   
'   Cancel = False
'   textCP10_Validate Cancel
'   If Cancel = True Then
'      Exit Sub
'   End If
   
   '確定
   If Index = 0 Then
      '檢查...
      If List1.ListCount = 0 Then
         MsgBox "請新增案號！", vbExclamation
         Exit Sub
      End If
      
      If MsgBox("確定複製檔案到下列案號嗎？", vbYesNo, "詢問") = vbNo Then
         Exit Sub
      End If
      
      '檢查要複製的電子檔
      For jj = 0 To lstAtt.ListCount - 1
         strTemp = lstAtt.List(jj)
         arrData = Split(strTemp, "  ")
         strOldCPP01 = Trim(arrData(0))
         strOldCPP02 = Trim(arrData(1))
         If UBound(arrData) <> 1 Then
            MsgBox "欲複製的電子檔資料有誤！", vbExclamation
            Exit Sub
         Else
            If Dir(m_AttachPath & "\" & strOldCPP02) <> "" Then
               Kill m_AttachPath & "\" & strOldCPP02
            End If
            '下載檔案到暫存區:
            '卷宗區
            If UCase(TypeName(m_PrevForm)) = UCase("frm100101_L") Then
               If PUB_GetAttachFile_CPP(strOldCPP01, strOldCPP02, m_AttachPath) = False Then
                  MsgBox "檔案下載失敗[ " & strOldCPP02 & " ]！", vbCritical
                  Exit Sub
               End If
            '原始檔區
            Else
               If PUB_GetAttachFile_Org(strOldCPP01, strOldCPP02, m_AttachPath) = False Then
                  MsgBox "檔案下載失敗[ " & strOldCPP02 & " ]！", vbCritical
                  Exit Sub
               End If
            End If
         End If
      Next jj
      
      Screen.MousePointer = vbHourglass
      
      '切換至來源目錄
      If m_AttachPath <> "." Then ChDir m_AttachPath
      For ii = 0 To List1.ListCount - 1
         strTemp = List1.List(ii)
         arrData = Split(strTemp, " ")
         strNewCP01 = SystemNumber(CStr(arrData(0)), 1)
         strNewCP02 = SystemNumber(CStr(arrData(0)), 2)
         strNewCP03 = SystemNumber(CStr(arrData(0)), 3)
         strNewCP04 = SystemNumber(CStr(arrData(0)), 4)
         strNewCP09 = Trim(arrData(1))
         strNewCP10 = Trim(arrData(2))
         
         For jj = 0 To lstAtt.ListCount - 1
            '原檔案名稱轉換為新檔案名稱
            Call TransCopyData(lstAtt.List(jj), strOldCPP01, strOldCPP02, _
                  strNewCP01, strNewCP02, strNewCP03, strNewCP04, strNewCP10, strNewCPP02)
            
            If Right(Trim(UCase(strOldCPP02)), 4) = UCase(".MSG") Or _
               Right(Trim(UCase(strOldCPP02)), 4) = UCase(".EML") Then
               Text2 = GetMsgFileText(m_AttachPath & "\" & strOldCPP02) '取得主旨
'               If Right(Trim(UCase(strOldCPP02)), 4) = UCase(".EML") Then
'                  Text2 = CStr(sFile(ii))
'               End If
            End If
            
            '檢查此檔案是否已存在
            If IsRecordExist(strNewCP09, strNewCPP02) = False Then
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(m_AttachPath & "\" & strOldCPP02)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg strOldCPP02 & MsgText(9221)
                  GoTo ErrHnd
               End If
               
               'Add by Sindy 2021/11/2 檢查畫面的物件是否含有Unicode文字
               Call PUB_ChkUniText(Me, , , "TextBox", , True)
               '2021/11/2 END
               
               '存檔:
               '卷宗區
               If UCase(TypeName(m_PrevForm)) = UCase("frm100101_L") Then
                  If SaveAttFile_PDF(strNewCP09, m_AttachPath & "\" & strOldCPP02, strNewCPP02, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(UCase(Right(strOldCPP02, 4)) = ".PDF", False, True), "A", , , , , Text2) = True Then
                     'bolAdd = True
                  Else
                     GoTo ErrHnd
                  End If
               '原始檔區
               Else
                  If SaveAttFile_Org(strNewCP09, m_AttachPath & "\" & strOldCPP02, strNewCPP02, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = True Then
                     'bolAdd = True
                  Else
                     GoTo ErrHnd
                  End If
               End If
'               Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
            End If
         Next jj
      Next ii
      
      ChDir App.path '目錄切回,釋放資料夾權限
      
      MsgBox "已複製完成！", vbExclamation
   End If

   Unload Me
   Screen.MousePointer = vbDefault

   Exit Sub

ErrHnd:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'原檔案名稱轉換為新檔案名稱
Private Function TransCopyData(ByVal strOldData As String, ByRef strOldCPP01 As String, ByRef strOldCPP02 As String, _
   ByVal strNewCP01 As String, ByVal strNewCP02 As String, ByVal strNewCP03 As String, ByVal strNewCP04 As String, _
   ByVal strNewCP10 As String, ByRef strNewCPP02 As String)
Dim arrData As Variant
Dim adoRst As ADODB.Recordset
Dim strOldCP10 As String
Dim strFileCaseOld As String, strFileCaseNew As String

   arrData = Split(strOldData, "  ")
   '舊文號
   strOldCPP01 = Trim(arrData(0))
   strOldCPP02 = Trim(arrData(1))
   '取得舊案件性質
   'FCP客戶提供文件處理後，D類收文會刪除
   strSql = "SELECT 1 as ord1, cp10 FROM caseprogress WHERE cp09='" & strOldCPP01 & "'"
   If Left(strOldCPP01, 1) = "D" Then
       strSql = strSql & " Union select 2 ord1 , '1920' as cp10 from (select * from custsupportdoc where csd05='" & strOldCPP01 & "' and nvl(csd11,0)>0 ) "
   End If
   strSql = strSql & " order by ord1"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strOldCP10 = adoRst.Fields("cp10")
   End If
   
   strFileCaseOld = Trim(strOldCP01) & Format(Val(strOldCP02), "000000") & _
        IIf(Trim(strOldCP03) & Trim(strOldCP04) <> "" And Trim(strOldCP03) & Trim(strOldCP04) <> "000", "-" & strOldCP03, "") & _
        IIf(Trim(strOldCP04) <> "" And Trim(strOldCP04) <> "00", "-" & Format(strOldCP04, "00"), "")
   If InStr(strOldCPP02, strFileCaseOld) = 0 Then
      strFileCaseOld = Trim(strOldCP01) & Val(strOldCP02) & _
         IIf(Trim(strOldCP03) & Trim(strOldCP04) <> "" And Trim(strOldCP03) & Trim(strOldCP04) <> "000", "-" & strOldCP03, "") & _
         IIf(Trim(strOldCP04) <> "" And Trim(strOldCP04) <> "00", "-" & Format(strOldCP04, "00"), "")
   End If
   
   strFileCaseNew = Trim(strNewCP01) & Format(Val(strNewCP02), "000000") & _
        IIf(Trim(strNewCP03) & Trim(strNewCP04) <> "" And Trim(strNewCP03) & Trim(strNewCP04) <> "000", "-" & strNewCP03, "") & _
        IIf(Trim(strNewCP04) <> "" And Trim(strNewCP04) <> "00", "-" & Format(strNewCP04, "00"), "")
   
   strNewCPP02 = strOldCPP02
   '置換案號
   If strFileCaseOld <> strFileCaseNew Then
      strNewCPP02 = Replace(strNewCPP02, strFileCaseOld & ".", strFileCaseNew & ".")
   End If
   If InStr(strNewCPP02, strFileCaseNew) = 0 Then
      strNewCPP02 = strFileCaseNew & "." & strNewCPP02
   End If
   '置換案件性質
   If strOldCP10 <> strNewCP10 Then
      strNewCPP02 = Replace(strNewCPP02, "." & strOldCP10 & ".", "." & strNewCP10 & ".")
   End If
   
   Set adoRst = Nothing
End Function

' 若為郵件要取主旨內容
Private Function GetMsgFileText(ByVal m_SourceFile) As String
Dim objOutLook As Object
Dim objMail As Object
Dim strTempFileName As String
Dim varTemp As Variant
   
   GetMsgFileText = ""
   varTemp = Split(m_SourceFile, ".")
   If UBound(varTemp) > 0 Then
      If UCase(varTemp(UBound(varTemp))) <> UCase("msg") Then Exit Function '非郵件,離開
   Else
      Exit Function
   End If
   
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItemFromTemplate(m_SourceFile)
   Me.Text2 = objMail.Subject 'Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
   GetMsgFileText = ChgSQL(Me.Text2) '要用文字框存放，因才能把unicode去掉
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCP09 As String, ByVal strFileName As String, _
   Optional ByRef strCPP06 As String, Optional ByRef strCPP07 As String, _
   Optional ByVal bolShowMsg As Boolean = True) As Boolean
Dim adoRst As ADODB.Recordset

   IsRecordExist = False
   
   '卷宗區
   If UCase(TypeName(m_PrevForm)) = UCase("frm100101_L") Then
      strSql = "SELECT cpp01,cpp06,cpp07 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and upper(cpp02)=upper('" & strFileName & "')"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCPP06 = "" & adoRst.Fields("cpp06") '異動日期
         strCPP07 = "" & adoRst.Fields("cpp07") '異動時間
         IsRecordExist = True
         If bolShowMsg = True Then
            MsgBox "文號：" & strCP09 & " 附件：" & strFileName & " 已存在！", vbExclamation
         End If
      End If
   '原始檔區
   ElseIf UCase(TypeName(m_PrevForm)) = UCase("frm100101_M") Then
      strSql = "SELECT cpf01,cpf06,cpf07 FROM casepaperfile WHERE cpf01='" & strCP09 & "' and upper(cpf02)=upper('" & strFileName & "')"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCPP06 = "" & adoRst.Fields("cpf06") '異動日期
         strCPP07 = "" & adoRst.Fields("cpf07") '異動時間
         IsRecordExist = True
         If bolShowMsg = True Then
            MsgBox "文號：" & strCP09 & " 附件：" & strFileName & " 已存在！", vbExclamation
         End If
      End If
   Else
      IsRecordExist = True
      MsgBox "無讀取到檔案資料！", vbExclamation
   End If
   
   Set adoRst = Nothing
End Function

Private Sub Command1_Click(Index As Integer)
Dim strTmp As String
Dim intChkCnt As Integer
Dim adoRst As ADODB.Recordset
Dim Cancel As Boolean
Dim strData As String
Dim arrData As Variant
Dim strNewCP01 As String, strNewCP02 As String, strNewCP03 As String, strNewCP04 As String
Dim strNewCP09 As String, strNewCP10 As String, strNewCPP02 As String
Dim strOldCPP01 As String, strOldCPP02 As String, strOldCP10 As String
   
   If Index = 1 Then
      If List1.ListCount = 0 Then
         MsgBox "無資料列可刪除！", vbExclamation
         Exit Sub
      End If
   Else
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
   End If
   
   '新增
   If Index = 0 And txtCode(0).Text <> "" Then
      strTmp = txtSystem & "-" & txtCode(0)
      If txtCode(1).Text = "" Then
         strTmp = strTmp & "-" & "0"
      Else
         strTmp = strTmp & "-" & txtCode(1).Text
      End If
      If txtCode(2).Text = "" Then
         strTmp = strTmp & "-" & "00"
      Else
         strTmp = strTmp & "-" & txtCode(2).Text
      End If
      
'      If Trim(textCP10) = "" And Index = 2 Then
'         MsgBox "案件性質不可空白！", vbExclamation
'         textCP10.SetFocus
'         Exit Sub
'      '新增案號
'      ElseIf Trim(textCP10) <> "" And Index = 2 Then
'         '檢查案號,是否有此案件性質
'         '若有，抓此案號該案件性質收文日最大的收文號。
'         strSql = "SELECT cp09,cp10 FROM caseprogress WHERE cp01='" & txtSystem & "'" & _
'                  " and cp02='" & txtCode(0) & "' and cp03='" & txtCode(1) & "' and cp04='" & txtCode(2) & "'" & _
'                  " and cp10='" & textCP10 & "'" & _
'                  " order by CP05 desc,CP09 desc"
'         intI = 1
'         Set adoRst = ClsLawReadRstMsg(intI, strSql)
'         If intI = 0 Then
'            MsgBox "該案號無此案件性質！", vbExclamation
'            txtCode(0).SetFocus
'            Set adoRst = Nothing
'            Exit Sub
'         Else
'            '案號 + ' ' + 總收文號 + ' ' + 案件性質
'            strData = strTmp & " " & Trim(adoRst.Fields("cp09")) & " " & Trim(adoRst.Fields("cp10"))
'         End If
'         Set adoRst = Nothing
'
''         If GRD1.Rows > 1 Then
''            '清空及預設欄位值
''            GRD1.Clear
''            Call SetGrd
''         End If
'
'      '新增指定總收文號
'      Else
         For ii = 1 To GRD1.Rows - 1
            GRD1.row = ii
            GRD1.col = 1
            If GRD1.CellBackColor = &HFFC0C0 Then
               intChkCnt = intChkCnt + 1
               '案號 + ' ' + 總收文號 + ' ' + 案件性質
               strData = strTmp & " " & Trim(GRD1.TextMatrix(ii, 1)) & " " & Trim(GRD1.TextMatrix(ii, 5))
               Exit For
            End If
         Next ii
         If intChkCnt = 0 Then
            MsgBox "請勾選要複製到那一道程序！", vbExclamation
            Exit Sub
         End If
'      End If
      
      If strData <> "" Then
         arrData = Split(strData, " ")
         strNewCP01 = SystemNumber(CStr(arrData(0)), 1)
         strNewCP02 = SystemNumber(CStr(arrData(0)), 2)
         strNewCP03 = SystemNumber(CStr(arrData(0)), 3)
         strNewCP04 = SystemNumber(CStr(arrData(0)), 4)
         strNewCP09 = Trim(arrData(1))
         strNewCP10 = Trim(arrData(2))
         
         'Add By Sindy 2022/5/30 開放複製卷宗區檔案至相同案號的不同程序的功能
         If InStr(strRecvNo, Trim(arrData(1))) > 0 Then
            MsgBox "不可複製到同一道程序！", vbExclamation
            Exit Sub
         End If
         '2022/5/30 END
         
         'Modify By Sindy 2022/5/5
         For ii = 0 To List1.ListCount - 1
            'Add By Sindy 2022/5/30 Mark: 開放複製卷宗區檔案至相同案號的不同程序的功能
'            If InStr(List1.List(ii), strNewCP01 & "-" & strNewCP02 & "-" & strNewCP03 & "-" & strNewCP04) > 0 Then
'               MsgBox "此案號( " & strData & " )已重覆，新增失敗！", vbExclamation
'               Exit Sub
'            End If
            If InStr(List1.List(ii), strNewCP09) > 0 Then
               MsgBox "此總收文號( " & strData & " )已重覆，新增失敗！", vbExclamation
               Exit Sub
            End If
         Next ii
         '2022/5/5 END
         
         For jj = 0 To lstAtt.ListCount - 1
            '原檔案名稱轉換為新檔案名稱
            Call TransCopyData(lstAtt.List(jj), strOldCPP01, strOldCPP02, _
                  strNewCP01, strNewCP02, strNewCP03, strNewCP04, strNewCP10, strNewCPP02)
            '檢查此檔案是否已存在
            If IsRecordExist(strNewCP09, strNewCPP02) = True Then
               Exit Sub
            End If
         Next jj
         
         '新增至案號區
         List1.AddItem strData
      End If
      
   '刪除
   Else
      If List1.ListCount > 0 Then
         ii = 0
         Do While ii < List1.ListCount
            If List1.Selected(ii) = True Then
               List1.RemoveItem ii
               ii = ii - 1
            End If
            ii = ii + 1
         Loop
      End If
   End If
'   Command5.Default = True
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim adoRst As ADODB.Recordset
Dim strData1 As String, strData2 As String, strData3 As String, strData4 As String, strData5 As String
   
   TxtValidate = False
   
   If txtSystem <> "" And txtCode(0) <> "" Then
      If txtCode(1) = "" Then txtCode(1) = "0"
      If txtCode(2) = "" Then txtCode(2) = "00"
   End If
   
   If txtSystem = "" Then
      MsgBox "系統別不可空白！", vbExclamation
      If txtSystem.Enabled = True Then txtSystem.SetFocus
      Exit Function
   End If
   If txtCode(0) = "" Then
      MsgBox "案號不可空白！", vbExclamation
      If txtCode(0).Enabled = True Then txtCode(0).SetFocus
      Exit Function
   End If
   
   'Modify By Sindy 2022/5/30 Mark: 開放複製卷宗區檔案至相同案號的不同程序的功能
   If lblCaseNo.Caption = txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2) Then
'      MsgBox "輸入的本所案號不可相同！", vbExclamation
'      If txtCode(0).Enabled = True Then txtCode(0).SetFocus
'      Exit Function
      'Modify By Sindy 2022/7/7
      textCP10.Text = ""
      LblCP10.Caption = ""
      '2022/7/7 END
   End If
   
   '檢查使用者權限
   If CheckSR09(strUserNum, txtSystem, "Y", , txtSystem, txtCode(0), txtCode(1), txtCode(2)) = False Then
      Exit Function
   End If
   
   m_Nation = GetPrjNation1(txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2))
   If Trim(textCP10) <> "" Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '檢查必須為同一FC代理人或同一申請人或同一CF代理人:
   'FC代理人
   strSql = "select tm44 from trademark where tm01='" & strOldCP01 & "' and tm02='" & strOldCP02 & "' and tm03='" & strOldCP03 & "' and tm04='" & strOldCP04 & "'" & _
            " union select pa75 from patent where pa01='" & strOldCP01 & "' and pa02='" & strOldCP02 & "' and pa03='" & strOldCP03 & "' and pa04='" & strOldCP04 & "'" & _
            " union select sp26 from servicepractice where sp01='" & strOldCP01 & "' and sp02='" & strOldCP02 & "' and sp03='" & strOldCP03 & "' and sp04='" & strOldCP04 & "'" & _
            " union select lc22 from lawcase where lc01='" & strOldCP01 & "' and lc02='" & strOldCP02 & "' and lc03='" & strOldCP03 & "' and lc04='" & strOldCP04 & "'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData1 = "" & adoRst.Fields(0)
      If strData1 <> "" Then
         strSql = "select tm44 from trademark where tm01='" & txtSystem & "' and tm02='" & txtCode(0) & "' and tm03='" & txtCode(1) & "' and tm04='" & txtCode(2) & "'" & _
                  " union select pa75 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "'" & _
                  " union select sp26 from servicepractice where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & txtCode(1) & "' and sp04='" & txtCode(2) & "'" & _
                  " union select lc22 from lawcase where lc01='" & txtSystem & "' and lc02='" & txtCode(0) & "' and lc03='" & txtCode(1) & "' and lc04='" & txtCode(2) & "'"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Left(strData1, 8) = Left("" & adoRst.Fields(0), 8) Then
               Set adoRst = Nothing
               TxtValidate = True
               Exit Function
            End If
         End If
      End If
   End If
   'CF代理人
   strSql = "select * from (" & _
            "Select CP44,CP45,CP05,CP09,CP27,2 as iSort From Caseprogress Where CP01='" & strOldCP01 & "' and CP02='" & strOldCP02 & "' and CP03='" & strOldCP03 & "' and CP04='" & strOldCP04 & "' AND CP09 <'C' AND CP27 IS NOT NULL AND CP57 IS NULL" & _
            " union Select CP44,CP45,CP05,CP09,CP27,2 as iSort From Caseprogress Where CP01='" & strOldCP01 & "' and CP02='" & strOldCP02 & "' and CP03='" & strOldCP03 & "' and CP04='" & strOldCP04 & "' AND CP09 <'C' AND CP57 IS NULL AND CP44 IS NOT NULL" & _
            ") order by iSort asc,nvl(CP27,CP05) DESC, CP09 DESC"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData1 = "" & adoRst.Fields(0)
      If strData1 <> "" Then
         strSql = "select * from (" & _
                  "Select CP44,CP45,CP05,CP09,CP27,2 as iSort From Caseprogress Where CP01='" & txtSystem & "' and CP02='" & txtCode(0) & "' and CP03='" & txtCode(1) & "' and CP04='" & txtCode(2) & "' AND CP09 <'C' AND CP27 IS NOT NULL AND CP57 IS NULL" & _
                  " union Select CP44,CP45,CP05,CP09,CP27,2 as iSort From Caseprogress Where CP01='" & txtSystem & "' and CP02='" & txtCode(0) & "' and CP03='" & txtCode(1) & "' and CP04='" & txtCode(2) & "' AND CP09 <'C' AND CP57 IS NULL AND CP44 IS NOT NULL" & _
                  ") order by iSort asc,nvl(CP27,CP05) DESC, CP09 DESC"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Left(strData1, 8) = Left("" & adoRst.Fields(0), 8) Then
               Set adoRst = Nothing
               TxtValidate = True
               Exit Function
            End If
         End If
      End If
   End If
   '同一申請人
   strSql = "select tm23,tm78,tm79,tm80,tm81 from trademark where tm01='" & strOldCP01 & "' and tm02='" & strOldCP02 & "' and tm03='" & strOldCP03 & "' and tm04='" & strOldCP04 & "'" & _
            " union select pa26,pa27,pa28,pa29,pa30 from patent where pa01='" & strOldCP01 & "' and pa02='" & strOldCP02 & "' and pa03='" & strOldCP03 & "' and pa04='" & strOldCP04 & "'" & _
            " union select sp08,sp58,sp59,sp65,sp66 from servicepractice where sp01='" & strOldCP01 & "' and sp02='" & strOldCP02 & "' and sp03='" & strOldCP03 & "' and sp04='" & strOldCP04 & "'" & _
            " union select lc11,lc43,lc44,lc45,lc46 from lawcase where lc01='" & strOldCP01 & "' and lc02='" & strOldCP02 & "' and lc03='" & strOldCP03 & "' and lc04='" & strOldCP04 & "'" & _
            " union select hc05,hc24,hc25,hc26,hc27 from hirecase where hc01='" & strOldCP01 & "' and hc02='" & strOldCP02 & "' and hc03='" & strOldCP03 & "' and hc04='" & strOldCP04 & "'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData1 = "" & adoRst.Fields(0)
      strData2 = "" & adoRst.Fields(1)
      strData3 = "" & adoRst.Fields(2)
      strData4 = "" & adoRst.Fields(3)
      strData5 = "" & adoRst.Fields(4)
      If strData1 <> "" Then
         strSql = "select tm23||','||tm78||','||tm79||','||tm80||','||tm81 from trademark where tm01='" & txtSystem & "' and tm02='" & txtCode(0) & "' and tm03='" & txtCode(1) & "' and tm04='" & txtCode(2) & "'" & _
            " union select pa26||','||pa27||','||pa28||','||pa29||','||pa30 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "'" & _
            " union select sp08||','||sp58||','||sp59||','||sp65||','||sp66 from servicepractice where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & txtCode(1) & "' and sp04='" & txtCode(2) & "'" & _
            " union select lc11||','||lc43||','||lc44||','||lc45||','||lc46 from lawcase where lc01='" & txtSystem & "' and lc02='" & txtCode(0) & "' and lc03='" & txtCode(1) & "' and lc04='" & txtCode(2) & "'" & _
            " union select hc05||','||hc24||','||hc25||','||hc26||','||hc27 from hirecase where hc01='" & txtSystem & "' and hc02='" & txtCode(0) & "' and hc03='" & txtCode(1) & "' and hc04='" & txtCode(2) & "'"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If (Left(strData1, 8) <> "" And InStr(adoRst.Fields(0), Left(strData1, 8)) > 0) Or _
               (Left(strData2, 8) <> "" And InStr(adoRst.Fields(0), Left(strData2, 8)) > 0) Or _
               (Left(strData3, 8) <> "" And InStr(adoRst.Fields(0), Left(strData3, 8)) > 0) Or _
               (Left(strData4, 8) <> "" And InStr(adoRst.Fields(0), Left(strData4, 8)) > 0) Or _
               (Left(strData5, 8) <> "" And InStr(adoRst.Fields(0), Left(strData5, 8)) > 0) Then
               Set adoRst = Nothing
               TxtValidate = True
               Exit Function
            Else
               Set adoRst = Nothing
               MsgBox "必須為同一FC代理人或同一申請人或同一CF代理人！", vbExclamation
               Exit Function
            End If
         Else
            Set adoRst = Nothing
            MsgBox txtSystem & "-" & txtCode(0) & _
               IIf(txtCode(1) & txtCode(2) <> "000", "-" & txtCode(1) & "-" & txtCode(2), "") & _
               "無資料！", vbExclamation
            Exit Function
         End If
      Else
         Set adoRst = Nothing
         MsgBox "必須為同一FC代理人或同一申請人或同一CF代理人！", vbExclamation
         Exit Function
      End If
   Else
      Set adoRst = Nothing
      MsgBox strOldCP01 & "-" & strOldCP02 & _
         IIf(strOldCP03 & strOldCP04 <> "000", "-" & strOldCP03 & "-" & strOldCP04, "") & _
         "無資料！", vbExclamation
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub Command5_Click()
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   Call QueryData
End Sub

Private Sub Form_Load()
Dim sFile

   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   lstAtt.Clear
   If m_strSaveFiles <> "" Then
      sFile = Split(m_strSaveFiles, "&")
      For ii = 0 To UBound(sFile)
         lstAtt.AddItem sFile(ii), 0
         SetListScroll lstAtt
      Next ii
   End If
   
   m_strUserRight = GetSystemKindByNick
   If m_strUserRight <> "" Then
      m_arrUserRight = Split(m_strUserRight, ",")
   End If

   Me.lblCaseNo = m_PrevForm.lblCaseNo
   Me.lblRecvNo = Replace(strRecvNo, "'", "") 'Add By Sindy 2022/5/30
   Me.lblCaseName = m_PrevForm.lblCaseName
   strOldCP01 = SystemNumber(lblCaseNo, 1)
   strOldCP02 = SystemNumber(lblCaseNo, 2)
   strOldCP03 = SystemNumber(lblCaseNo, 3)
   strOldCP04 = SystemNumber(lblCaseNo, 4)
   
   Me.textCP10 = Me.m_CP10
   Call textCP10_Validate(False)
   
   Call SetGrd
   
   '卷宗區
   If UCase(TypeName(m_PrevForm)) = UCase("frm100101_L") Then
      Me.Caption = "複製檔案 - 卷宗區"
   '原始檔區
   Else
      Me.Caption = "複製檔案 - 原始檔區"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm100101_L_4 = Nothing
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String

   QueryData = False
   '清空及預設欄位值
   GRD1.Clear
   Call SetGrd
   
   m_CP01 = txtSystem.Text
   m_CP02 = Left(Trim(txtCode(0).Text) & "00000", 6)
   m_CP03 = txtCode(1).Text
   m_CP04 = txtCode(2).Text
   lblCaseName2.Caption = ""
      
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   m_Nation = GetPrjNation1(txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2))
   strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp05) as 收文日,NVL(DECODE('" & m_Nation & "','000',CPM03,CPM04),CP10) as 案件性質,sqldatet(cp27) as 發文日,CP10,CP43" & _
            " From caseprogress,casepropertymap" & _
            " Where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)"
   If textCP10 <> "" Then
      strSql = strSql & " and cp10='" & textCP10 & "'"
   End If
   strSql = strSql & " order by CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      QueryData = True
      GRD1.col = 0
      GRD1.row = 1
      lblCaseName2.Caption = GetPrjName(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04)
'      cmdOK(0).Default = True
      '判斷有相關總收文號才做較快
      For ii = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(ii, 6) <> "" Then
            GRD1.TextMatrix(ii, 3) = GRD1.TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(GRD1.TextMatrix(ii, 1), "1")
         End If
      Next ii
      
      If rsTmp.RecordCount = 1 Then
         GRD1.col = 0
         GRD1.row = 1
         '資料列反白
         GRD1.TextMatrix(GRD1.row, 0) = "V"
         For jj = 1 To GRD1.Cols - 1
            GRD1.col = jj
            GRD1.CellBackColor = &HFFC0C0
         Next jj
'         Command1(0).Default = True
      End If
   Else
      ShowNoData
   End If
   rsTmp.Close
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Set rsTmp = Nothing
End Function

Private Sub SetGrd(Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   '                        0    1           2         3           4         5       6
   arrGridHeadText = Array("V", "總收文號", "收文日", "案件性質", "發文日", "CP10", "CP43")
   arrGridHeadWidth = Array(200, 950, 800, 1250, 800, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If bolSetRow = True Then
      GRD1.Rows = 2
   End If
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      If bolSetRow = True Then
         GRD1.Text = arrGridHeadText(iRow)
      End If
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If bolSetRow = True Then
         GRD1.CellAlignment = flexAlignCenterCenter
      End If
   Next
   GRD1.Visible = True
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

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
Dim oldRow As Long

getGrdColRow GRD1, X, Y, nCol, nRow
If nCol < 0 Then Exit Sub
GRD1.col = nCol
GRD1.row = nRow
oldRow = m_mouseRow

GRD1.Visible = False
If GRD1.MouseRow <> 0 And Trim(GRD1.TextMatrix(GRD1.MouseRow, 1)) <> "" Then
   m_mouseRow = GRD1.MouseRow
   If oldRow <> m_mouseRow And oldRow <= GRD1.Rows - 1 Then
      GRD1.row = oldRow
      GRD1.col = 1
      If GRD1.CellBackColor = &HFFC0C0 Then
         '清除反白
         GRD1.TextMatrix(oldRow, 0) = ""
         For jj = 1 To GRD1.Cols - 1
            GRD1.col = jj
            GRD1.CellBackColor = QBColor(15)
         Next jj
      End If
   End If

   GRD1.row = GRD1.MouseRow
   GRD1.col = 1
   If GRD1.CellBackColor = &HFFC0C0 Then
      '清除反白
      GRD1.TextMatrix(GRD1.MouseRow, 0) = ""
      For jj = 1 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = QBColor(15)
      Next jj
   Else
      '資料列反白
      GRD1.TextMatrix(GRD1.MouseRow, 0) = "V"
      For jj = 1 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = &HFFC0C0
      Next jj
   End If
End If
GRD1.Visible = True
End Sub

Private Sub textCP10_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
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

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   LblCP10 = Empty
   If IsEmptyText(textCP10) = False Then
      If m_Nation < "010" Then
          LblCP10.Caption = GetCaseTypeName(IIf(txtSystem <> "", txtSystem, strOldCP01), textCP10, 0)
      Else
          LblCP10.Caption = GetCaseTypeName(IIf(txtSystem <> "", txtSystem, strOldCP01), textCP10, 1)
      End If
      If IsEmptyText(LblCP10) = True Then
          Cancel = True
          strTit = "檢核資料"
          strMsg = "案件性質代號不存在"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP10_GotFocus
          Exit Sub
      End If
   End If
End Sub

'Private Sub txtSystem_LostFocus()
'   blnUserRight = False
'   If Me.txtSystem.Text <> "" Then
'      If m_strUserRight <> "" Then
'         For ii = LBound(m_arrUserRight) To UBound(m_arrUserRight)
'            If m_arrUserRight(ii) = Me.txtSystem.Text Then
'               blnUserRight = True
'            End If
'         Next ii
'         If blnUserRight = False Then
'            MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
'            Me.txtSystem.SetFocus
'            txtSystem_GotFocus
'            Exit Sub
'         End If
'      Else
'         MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
'         Me.txtSystem.SetFocus
'         txtSystem_GotFocus
'         Exit Sub
'      End If
'   End If
'End Sub
