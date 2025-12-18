VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_L_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "新增檔案"
   ClientHeight    =   4430
   ClientLeft      =   50
   ClientTop       =   300
   ClientWidth     =   8530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4430
   ScaleWidth      =   8530
   StartUpPosition =   3  '系統預設值
   Begin VB.OptionButton Option1 
      Caption         =   "契約書(.AGR.PDF)"
      Height          =   255
      Index           =   11
      Left            =   1050
      TabIndex        =   6
      Top             =   2100
      Width           =   2235
   End
   Begin VB.OptionButton Option1 
      Caption         =   "客戶確收(.CACK.PDF)"
      Height          =   255
      Index           =   10
      Left            =   3330
      TabIndex        =   12
      Top             =   1800
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "存卷資料(.INFO.PDF)"
      Height          =   255
      Index           =   9
      Left            =   3330
      TabIndex        =   14
      Top             =   2400
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其他各式附件(.ATT.PDF)"
      Height          =   255
      Index           =   8
      Left            =   3330
      TabIndex        =   13
      Top             =   2100
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "接洽單(.ORDER.PDF)"
      Height          =   255
      Index           =   7
      Left            =   1035
      TabIndex        =   10
      Top             =   3300
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "寄出郵件(.Tx.msg)"
      Height          =   255
      Index           =   6
      Left            =   5700
      TabIndex        =   16
      Top             =   2100
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "外來郵件(.Rx.msg)"
      Height          =   255
      Index           =   5
      Left            =   5700
      TabIndex        =   15
      Top             =   1800
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "銷案銷帳單(.Off.PDF)"
      Height          =   255
      Index           =   4
      Left            =   1035
      TabIndex        =   9
      Top             =   3000
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "通知函(.Cus.PDF)"
      Height          =   255
      Index           =   3
      Left            =   1035
      TabIndex        =   8
      Top             =   2700
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其他 (保留原檔名)(.PDF)"
      Height          =   255
      Index           =   2
      Left            =   1035
      TabIndex        =   11
      Top             =   3600
      Width           =   3435
   End
   Begin VB.OptionButton Option1 
      Caption         =   "客戶資料(.Case.)"
      Height          =   255
      Index           =   1
      Left            =   1050
      TabIndex        =   7
      Top             =   2400
      Width           =   2235
   End
   Begin VB.OptionButton Option1 
      Caption         =   "回覆單(.Reply.PDF)"
      Height          =   255
      Index           =   0
      Left            =   1035
      TabIndex        =   5
      Top             =   1800
      Width           =   2265
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "加入檔案(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6870
      TabIndex        =   0
      Top             =   90
      Width           =   930
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3810
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      VariousPropertyBits=   746604571
      Size            =   "1296;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   $"frm100101_L_2.frx":0000
      ForeColor       =   &H000000C0&
      Height          =   1155
      Left            =   5700
      TabIndex        =   28
      Top             =   2550
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   27
      Top             =   1530
      Width           =   960
   End
   Begin VB.Label lblAppl 
      Height          =   225
      Left            =   1035
      TabIndex        =   26
      Top             =   1530
      Width           =   5625
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   225
      Index           =   3
      Left            =   60
      TabIndex        =   25
      Top             =   1260
      Width           =   960
   End
   Begin VB.Label lblNa03 
      Height          =   225
      Left            =   1035
      TabIndex        =   24
      Top             =   1260
      Width           =   2145
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Left            =   1035
      TabIndex        =   23
      Top             =   660
      Width           =   5625
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   22
      Top             =   660
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "備註：匯入時，檔案將搬移至系統中。"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   90
      TabIndex        =   21
      Top             =   4110
      Width           =   3525
   End
   Begin VB.Label lblCP10Nm 
      Height          =   225
      Left            =   4185
      TabIndex        =   20
      Top             =   990
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   225
      Index           =   1
      Left            =   3210
      TabIndex        =   19
      Top             =   990
      Width           =   960
   End
   Begin VB.Label lblCP09 
      Height          =   225
      Left            =   1035
      TabIndex        =   18
      Top             =   990
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   18
      Left            =   60
      TabIndex        =   17
      Top             =   990
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "檔案類型："
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   945
   End
   Begin VB.Label lblCaseNo 
      Caption         =   "lblCaseNo"
      Height          =   225
      Left            =   1050
      TabIndex        =   2
      Top             =   390
      Width           =   2085
   End
End
Attribute VB_Name = "frm100101_L_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/2 Form2.0已修改
'Create By Sindy 2015/5/25
Option Explicit

Public m_identity As String
Public m_CP09 As String
Public m_CP10 As String
Public m_CPP11 As String 'Add By Sindy 2023/2/18 電子表單單號
Public m_CP10Nm As String
Public m_Nation As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_MousePointer As Integer
Dim ii As Integer, jj As Integer
Dim m_PrevForm As Form '前一畫面
Dim m_CP16 As String 'Add By Sindy 2022/9/13
Dim m_CP13 As String 'Add By Sindy 2025/10/23


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'新增
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim stReName As String
   Dim bolAdd As Boolean
   Dim intChkCnt As Integer
   Dim strFile As String
   Dim strSecName As String
   Dim bolSel As Boolean
   'Dim strCPP04 As String 'Add By Sindy 2017/6/20
   Dim strYourCP01 As String, strYourCP02 As String, strYourCP03 As String, strYourCP04 As String
   Dim strOurCP01 As String, strOurCP02 As String, strOurCP03 As String, strOurCP04 As String
   Dim strMailDate As String, strMailTime As String
   
On Error GoTo ErrHnd
   
   bolSel = False
   'Modify By Sindy 2018/10/22
   For ii = 0 To 11 '10
      If Option1(ii).Visible = True And Option1(ii).Value = True Then
         bolSel = True
      End If
   Next ii
   If bolSel = False Then
      MsgBox "請勾選欲新增的檔案類型！"
      Exit Sub
   End If
   
   If Option1(0).Value = True Then
      strSecName = EMP_回覆單
      'Add By Sindy 2021/8/25
      'Modify By Sindy 2021/9/22 + 開放身份為智權人員時,可以放C類回覆單
      If Left(m_CP09, 1) <> "A" And Left(m_CP09, 1) <> "B" And _
         Not (m_identity = "S" And Left(m_CP09, 1) = "C") Then
         MsgBox "總收文號必須是 A 或 B 類收文！", vbExclamation
         Exit Sub
      End If
      '2021/8/25 END
   'Add By Sindy 2021/9/22
   ElseIf Option1(10).Value = True Then
      strSecName = "CACK"
      'P117291(DB0024249)通知期限無法新增[客戶確收(.CACK)]:mark
'      If Left(m_CP09, 1) <> "A" And Left(m_CP09, 1) <> "B" And Left(m_CP09, 1) <> "C" Then
'         MsgBox "總收文號必須是 A 或 B 或 C 類收文！", vbExclamation
'         Exit Sub
'      End If
   ElseIf Option1(1).Value = True Then
      strSecName = EMP_客戶資料
   'Add By Sindy 2015/7/27
   ElseIf Option1(3).Value = True Then
      strSecName = EMP_通知函
   ElseIf Option1(4).Value = True Then
      strSecName = EMP_銷案銷帳單
      'Add By Sindy 2021/8/25
      If Left(m_CP09, 1) <> "A" And Left(m_CP09, 1) <> "B" Then
         'Modify By Sindy 2022/9/13 秀玲\阿蓮:CFT-22804銷帳，因直接領證，帳款輸在"註冊證"進度; 增加費用判斷
         If Val(m_CP16) = 0 Then
            MsgBox "總收文號必須是 A 或 B 類收文！", vbExclamation
            Exit Sub
         End If
      End If
      '2021/8/25 END
   '2015/7/27 END
   'Add By Sindy 2017/6/14
   ElseIf Option1(5).Value = True Then
      strSecName = "Rx" '郵件msg檔
   ElseIf Option1(6).Value = True Then
      strSecName = "Tx" '寄件備份msg檔
   '2017/6/14 END
   'Add By Sindy 2018/10/5
   ElseIf Option1(7).Value = True Then
      strSecName = EMP_接洽單
   'Add By Sindy 2018/10/22
   ElseIf Option1(8).Value = True Then
      strSecName = "ATT" '其他各式附件
   ElseIf Option1(9).Value = True Then
      strSecName = EMP_存卷資料
   'Add By Sindy 2023/2/7
   ElseIf Option1(11).Value = True Then
      strSecName = "AGR" '契約書
   End If
   
   bolAdd = False
   'Modify By Sindy 2023/2/7
   'stFileName = "*.*"
   If Option1(5).Value = True Or _
      Option1(6).Value = True Then
      stFileName = "*.msg"
   ElseIf Option1(1).Value = True Then
      stFileName = "*.*"
   Else
      stFileName = "*.pdf"
   End If
   '2023/2/7 END
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      '.Filter = "All Files (*.*)|*.*"
      .Filter = "All Files (" & stFileName & ")|" & stFileName & ""
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               'Add By Sindy 2013/10/9
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  GoTo EXITSUB
               End If
               '2013/10/9 END
               'Added by Lydia 2018/10/09 因為卷宗區不直接顯示查名附件(TS.PDF), 所以限制不可用TS.PDF (ex.T217581.101.TS.PDF=>T217581.101.TS.Case.PDF)
               If (m_CP01 = "T" Or m_CP01 = "TS") And Option1(1).Value = False _
                       And Right(UCase(CStr(sFile(ii))), Len("." & UCase(TMQ_查名作業 & ".pdf"))) = UCase("." & TMQ_查名作業 & ".pdf") Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & UCase(TMQ_查名作業 & ".pdf") & "為查名單附件，不可使用於檔案命名"
                  GoTo EXITSUB
               End If
               'end 2018/10/09
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               
               '除了客戶資料,郵件msg,寄件備份msg外,均要PDF檔
               If Option1(1).Value = False And _
                  Option1(5).Value = False And _
                  Option1(6).Value = False Then
                  If Right(Trim(UCase(stFileName)), 4) <> UCase(".PDF") Then
                     MsgBox "格式不符,只可存放.PDF檔!!"
                     GoTo EXITSUB
                  End If
               'Add By Sindy 2017/6/14
               'Modify By Sindy 2017/8/16 + eml And Right(Trim(UCase(stFileName)), 4) <> UCase(".EML")
               ElseIf Option1(5).Value = True Or _
                      Option1(6).Value = True Then
                  If Right(Trim(UCase(stFileName)), 4) <> UCase(".MSG") Then
                     MsgBox "格式不符,只可存放.MSG!!"
                     GoTo EXITSUB
                  End If
               '2017/6/14 END
               End If
               
               'Add By Sindy 2025/10/23 銷案銷帳單(.Off.PDF)及接洽單(.Order.PDF)二種類型，
               '若操作人員為ST15為SXX或該案件之智權人員時，不出現此二種選項；
               '同時若選擇「其他」類型，檔案名稱也不可有OFF.PDF及ORDER.PDF的字樣。
               If Left(Pub_StrUserSt15, 1) = "S" Or m_CP13 = strUserNum Then
                  If Option1(2).Value = True Then
                     If Right(Trim(UCase(stFileName)), 8) = UCase(".Off.PDF") Then
                        MsgBox "不可新增銷案銷帳單(.Off.PDF)!!"
                        GoTo EXITSUB
                     End If
                     If Right(Trim(UCase(stFileName)), 10) = UCase(".Order.PDF") Then
                        MsgBox "不可新增接洽單(.Order.PDF)!!"
                        GoTo EXITSUB
                     End If
                  End If
               End If
               '2025/10/23 END
   
               '檢查檔名規則
               'Add By Sindy 2017/6/23
               strMailDate = "0"
               Text2 = GetMsgFileText(stFileName, strMailDate, strMailTime) '取得主旨
               If Right(Trim(UCase(stFileName)), 4) = UCase(".MSG") Or _
                  Right(Trim(UCase(stFileName)), 4) = UCase(".EML") Then
'                  If PUB_IPDeptGetCaseNo(Text2, "YOURREF", m_CP01, m_CP02, m_CP03, m_CP04) = False Then
'                     If PUB_IPDeptGetCaseNo(Text2, "OURREF", m_CP01, m_CP02, m_CP03, m_CP04) = False Then
'                        MsgBox "解析不到本所案號！"
'                        GoTo EXITSUB
'                     End If
'                  End If
                  If Right(Trim(UCase(stFileName)), 4) = UCase(".EML") Then
                     Text2 = CStr(sFile(ii))
                  End If
                  'Modify By Sindy 2019/1/11
                  'Modify By Sindy 2019/3/6 + , , , "L"
                  Call PUB_IPDeptGetCaseNo(Text2, "YOURREF", strYourCP01, strYourCP02, strYourCP03, strYourCP04, , , "L", , False)
                  Call PUB_IPDeptGetCaseNo(Text2, "OURREF", strOurCP01, strOurCP02, strOurCP03, strOurCP04, , , "L", , False)
                  If strYourCP02 = "" And strOurCP02 = "" Then
                     MsgBox "解析不到本所案號！"
                     GoTo EXITSUB
                  Else
                     If (strYourCP01 & "-" & strYourCP02 & "-" & strYourCP03 & "-" & strYourCP04) <> lblCaseNo And _
                        (strOurCP01 & "-" & strOurCP02 & "-" & strOurCP03 & "-" & strOurCP04) <> lblCaseNo Then
                        MsgBox "主旨的本所案號(" & _
                               IIf(strYourCP02 <> "", strYourCP01 & "-" & strYourCP02 & "-" & strYourCP03 & "-" & strYourCP04, "") & _
                               IIf(strOurCP02 <> "", IIf(strYourCP02 <> "", "、", "") & strOurCP01 & "-" & strOurCP02 & "-" & strOurCP03 & "-" & strOurCP04, "") & _
                               ")並非為此案件，不可新增！"
                        GoTo EXITSUB
                     End If
                  End If
               Else
               '2017/6/23 END
                  If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "Y", m_CP10, , , False, , , strSecName) = False Then
                     GoTo EXITSUB
                  End If
               End If
               If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, ChgSQL(CStr(sFile(ii))), stReName, True, 1, , , m_CP09, strSecName) = False Then GoTo EXITSUB
               'Add By Sindy 2017/6/23 郵件ReName
               stReName = PUB_ReMailFileName(m_CP09, stReName, strMailDate, strMailTime, strSecName)
               
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  GoTo EXITSUB
               'Add By Sindy 2014/3/11
               ElseIf f.Size > 5242880 Then
                  'If Pub_StrUserSt15 = "P13" Then
                     If MsgBox("檔案過大（容量超過5MB），確認是否要上傳？", vbYesNo, "警告") = vbNo Then
                        GoTo EXITSUB
                     End If
                  'End If
               '2014/3/11 END
               End If
               '2013/9/6 END
               
               'Add by Sindy 2021/11/2 檢查畫面的物件是否含有Unicode文字
               Call PUB_ChkUniText(Me, , , "TextBox", , True)
               '2021/11/2 END

               If IsRecordExist(stReName) = False Then
                  '存檔
                  'Modify By Sindy 2023/2/18 + m_CPP11
                  If SaveAttFile_PDF(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(UCase(Right(stFileName, 4)) = ".PDF", False, True), "A", , , m_CPP11, , Text2) = True Then
                     bolAdd = True
                  Else
                     GoTo EXITSUB
                  End If
                  Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
'                  Pub_SaveLog strUserNum, "新增卷宗區附件：" & sFile(ii), m_CP01, m_CP02, m_CP03, m_CP04, m_CP09
               End If
            Next ii
            Call ChkCP121 'Add By Sindy 2013/10/30
         '單檔
         Else
            'stFileName = GetFileName(.FileName)
            'Modify By Sindy 2013/10/9
            'strFile = GetFileName(.FileName)
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               GoTo EXITSUB
            End If
            '2013/10/9 END
        
            'Added by Lydia 2018/10/09 因為卷宗區不直接顯示查名附件(TS.PDF), 所以限制不可用TS.PDF (ex.T217581.101.TS.PDF=>T217581.101.TS.Case.PDF)
            If (m_CP01 = "T" Or m_CP01 = "TS") And Option1(1).Value = False _
                    And Right(UCase(strFile), Len("." & UCase(TMQ_查名作業 & ".pdf"))) = UCase("." & TMQ_查名作業 & ".pdf") Then
                        MsgBox strFile & vbCrLf & vbCrLf & UCase(TMQ_查名作業 & ".pdf") & "為查名單附件，不可使用於檔案命名"
                        GoTo EXITSUB
            End If
            'end 2018/10/09
                           
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            stFileName = .FileName
            
            '除了客戶資料,郵件msg,寄件備份msg外,均要PDF檔
            If Option1(1).Value = False And _
               Option1(5).Value = False And _
               Option1(6).Value = False Then
               If Right(Trim(UCase(stFileName)), 4) <> UCase(".PDF") Then
                  MsgBox "格式不符,只可存放.PDF檔!!"
                  GoTo EXITSUB
               End If
            'Add By Sindy 2017/6/14 And Right(Trim(UCase(stFileName)), 4) <> UCase(".EML")
            ElseIf Option1(5).Value = True Or _
                   Option1(6).Value = True Then
               If Right(Trim(UCase(stFileName)), 4) <> UCase(".MSG") Then
                  MsgBox "格式不符,只可存放.MSG檔!!"
                  GoTo EXITSUB
               End If
            '2017/6/14 END
            End If
            
            'Add By Sindy 2025/10/23 銷案銷帳單(.Off.PDF)及接洽單(.Order.PDF)二種類型，
            '若操作人員為ST15為SXX或該案件之智權人員時，不出現此二種選項；
            '同時若選擇「其他」類型，檔案名稱也不可有OFF.PDF及ORDER.PDF的字樣。
            If Left(Pub_StrUserSt15, 1) = "S" Or m_CP13 = strUserNum Then
               If Option1(2).Value = True Then
                  If Right(Trim(UCase(stFileName)), 8) = UCase(".Off.PDF") Then
                     MsgBox "不可新增銷案銷帳單(.Off.PDF)!!"
                     GoTo EXITSUB
                  End If
                  If Right(Trim(UCase(stFileName)), 10) = UCase(".Order.PDF") Then
                     MsgBox "不可新增接洽單(.Order.PDF)!!"
                     GoTo EXITSUB
                  End If
               End If
            End If
            '2025/10/23 END
            
            '檢查檔名規則
            'Add By Sindy 2017/6/23
            strMailDate = "0"
            Text2 = GetMsgFileText(stFileName, strMailDate, strMailTime) '取得主旨
            If Right(Trim(UCase(stFileName)), 4) = UCase(".MSG") Or _
               Right(Trim(UCase(stFileName)), 4) = UCase(".EML") Then
               If Right(Trim(UCase(stFileName)), 4) = UCase(".EML") Then
                  Text2 = strFile
               End If
               'Modify By Sindy 2019/1/11
               'Modify By Sindy 2019/3/6 + , , , "L"
               Call PUB_IPDeptGetCaseNo(Text2, "YOURREF", strYourCP01, strYourCP02, strYourCP03, strYourCP04, , , "L", , False)
               Call PUB_IPDeptGetCaseNo(Text2, "OURREF", strOurCP01, strOurCP02, strOurCP03, strOurCP04, , , "L", , False)
               If strYourCP02 = "" And strOurCP02 = "" Then
                  MsgBox "解析不到本所案號！"
                  GoTo EXITSUB
               Else
                  If (strYourCP01 & "-" & strYourCP02 & "-" & strYourCP03 & "-" & strYourCP04) <> lblCaseNo And _
                     (strOurCP01 & "-" & strOurCP02 & "-" & strOurCP03 & "-" & strOurCP04) <> lblCaseNo Then
                     MsgBox "主旨的本所案號(" & _
                            IIf(strYourCP02 <> "", strYourCP01 & "-" & strYourCP02 & "-" & strYourCP03 & "-" & strYourCP04, "") & _
                            IIf(strOurCP02 <> "", IIf(strYourCP02 <> "", "、", "") & strOurCP01 & "-" & strOurCP02 & "-" & strOurCP03 & "-" & strOurCP04, "") & _
                            ")並非為此案件，不可新增！"
                     GoTo EXITSUB
                  End If
               End If
'               Call PUB_IPDeptGetCaseNo(Text2, "YOURREF", m_CP01, m_CP02, m_CP03, m_CP04)
'               If (m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) <> lblCaseNo Then
'                  Call PUB_IPDeptGetCaseNo(Text2, "OURREF", m_CP01, m_CP02, m_CP03, m_CP04)
'                  If (m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) <> lblCaseNo Then
'                     m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = ""
'                  End If
'               End If
'               If m_CP01 = "" And m_CP02 = "" Then
'                  MsgBox "解析不到本所案號！"
'                  GoTo EXITSUB
'               End If
            Else
            '2017/6/23 END
               If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, "Y", m_CP10, , , False, , , strSecName) = False Then
                  GoTo EXITSUB
               End If
            End If
            If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, ChgSQL(strFile), stReName, True, 1, , , m_CP09, strSecName) = False Then GoTo EXITSUB
            stReName = PUB_ReMailFileName(m_CP09, stReName, strMailDate, strMailTime, strSecName) 'Add By Sindy 2017/6/23 郵件ReName
            
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               GoTo EXITSUB
            'Add By Sindy 2014/3/11
            ElseIf f.Size > 5242880 Then
               'If Pub_StrUserSt15 = "P13" Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     GoTo EXITSUB
                  End If
               'End If
            '2014/3/11 END
            End If
            '2013/9/6 END
            
            'Add by Sindy 2021/11/2 檢查畫面的物件是否含有Unicode文字
            Call PUB_ChkUniText(Me, , , "TextBox", , True)
            '2021/11/2 END
            
            If IsRecordExist(stReName) = False Then
               '存檔
               'Modify By Sindy 2023/2/18 + m_CPP11
               If SaveAttFile_PDF(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(UCase(Right(stFileName, 4)) = ".PDF", False, True), "A", , , m_CPP11, , Text2) = True Then
                  bolAdd = True
               Else
                  GoTo EXITSUB
               End If
               Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
'               Pub_SaveLog strUserNum, "新增卷宗區附件：" & strFile, m_CP01, m_CP02, m_CP03, m_CP04, m_CP09
               Call ChkCP121 'Add By Sindy 2013/10/30
            End If
         End If
EXITSUB:
         If bolAdd = True Then
            Call m_PrevForm.ReadAttachFile
         End If
      End If
      ChDir App.path 'Add By Sindy 2020/1/13 釋放資料夾權限
   End With
   
   Unload Me
   Screen.MousePointer = m_MousePointer
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

' 若為郵件要取主旨內容
Private Function GetMsgFileText(ByVal m_SourceFile, ByRef strMailDate As String, ByRef strMailTime As String) As String
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
   
   Call PUB_ReadMailText(objMail, , , , , strMailDate, strMailTime) 'Add By Sindy 2025/2/17
End Function

'' 若為郵件修改檔名為日期時間加序號
'Private Function ReMsgFileName(ByVal m_ReFileName As String) As String
'Dim strTempFileName As String
'Dim strTempFileName1 As String, strTempFileName2 As String
'Dim adoRst As ADODB.Recordset
'Dim intRow As Integer
'Dim varTemp As Variant
'
'   ReMsgFileName = m_ReFileName
'   intRow = 0
'   varTemp = Split(m_ReFileName, ".")
'   For ii = 0 To 1 'UBound(sFile)
'      strTempFileName1 = strTempFileName1 & varTemp(ii) & "."
'   Next ii
'   ii = UBound(varTemp)
'   If UCase(varTemp(ii)) <> UCase("msg") Then Exit Function '非郵件,離開
'
'   strTempFileName1 = strTempFileName1 & strSrvDate(1) & Right("000000" & ServerTime, 6) & "."
'   strTempFileName2 = varTemp(ii - 1) & "." & varTemp(ii)
'GotoChk:
'   If intRow > 0 Then
'      strTempFileName = strTempFileName1 & intRow & "." & strTempFileName2
'   Else
'      strTempFileName = strTempFileName1 & strTempFileName2
'   End If
'   strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01='" & m_CP09 & "' and upper(cpp02)=upper('" & ChgSQL(strTempFileName) & "')"
'   intI = 1
'   Set adoRst = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      intRow = intRow + 1
'      GoTo GotoChk
'   End If
'   ReMsgFileName = strTempFileName
'
'   Set adoRst = Nothing
'End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal stFileName As String) As Boolean
Dim adoRst As ADODB.Recordset
   
   IsRecordExist = False
   
   strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01='" & m_CP09 & "' and upper(cpp02)=upper('" & ChgSQL(stFileName) & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      IsRecordExist = True
      MsgBox "附件 " & stFileName & " 已存在！"
   End If
   
   Set adoRst = Nothing
End Function

'Add By Sindy 2013/10/30 檢查是否為電子送件,若是,電子檔是否全數歸檔
Private Sub ChkCP121()
   If m_CP01 = "P" And m_Nation = "000" Then
      'Modify By Sindy 2014/5/21 Mark 因UpdateCP121會檢查新案是否已歸足,若未歸足會重新檢核
'      '檢查新申請案
'      strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
'                  " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
'                  " and cp10 in(" & NewCasePtyList & ")" & _
'                  " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
'                  " and cp01=cpm01(+) and cp10=cpm02(+)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Call UpdateCP121(RsTemp.Fields("cp09"), RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"), m_CP09)
'      End If
      '檢查此筆文號
      strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
                  " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                  " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                  " and cp09='" & m_CP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Call UpdateCP121(m_CP09, RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"))
      End If
   'Added by Lydia 2019/03/06 FCP之公告公報1228增加判斷是否有公告本
   ElseIf m_CP01 = "FCP" And m_CP10 = "1228" Then
        Call UpdateCP121(m_CP09, "1228", "GAZ")
   End If
   '2013/10/30 END
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Screen.MousePointer = m_MousePointer
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_MousePointer = Screen.MousePointer
   
   Me.lblCaseNo = m_PrevForm.lblCaseNo
   Me.lblCaseName = m_PrevForm.lblCaseName
   Me.lblCP09 = m_PrevForm.m_CP09
   Me.lblCP10Nm = m_PrevForm.m_CP10Nm
   Me.lblNa03 = GetPrjNationName(m_PrevForm.m_Nation)
   Me.lblAppl = GetPrjPeople1(m_PrevForm.m_Appl)
   'Modify By Sindy 2015/7/30 開放檔案室可以用其他
   'Modify By Sindy 2023/2/7 取消點選其他的限制
'   If m_identity = "C" Or m_identity = "F" Or m_identity = "W" Then '電腦中心和程序人員才能看到其他
      Option1(2).Visible = True
'   Else
'      Option1(2).Visible = False
'   End If
   m_CP01 = SystemNumber(lblCaseNo, 1)
   m_CP02 = SystemNumber(lblCaseNo, 2)
   m_CP03 = SystemNumber(lblCaseNo, 3)
   m_CP04 = SystemNumber(lblCaseNo, 4)
   'Add By Sindy 2022/9/13
   strSql = "SELECT cp09,cp16 FROM caseprogress WHERE cp09='" & m_CP09 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_CP16 = "" & RsTemp.Fields("cp16")
   End If
   
   'Add By Sindy 2025/10/23 銷案銷帳單及接洽單二種類型，
   '若操作人員為ST15為SXX或該案件之智權人員時，不出現此二種選項；
   m_CP13 = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04) '目前智權人員
   If Left(Pub_StrUserSt15, 1) = "S" Or m_CP13 = strUserNum Then
      Option1(4).Enabled = False
      Option1(7).Enabled = False
   End If
   '2025/10/23 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm100101_L_2 = Nothing
End Sub
