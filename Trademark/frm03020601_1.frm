VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020601_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-延期"
   ClientHeight    =   5700
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   9204
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9204
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   1425
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   11
      Top             =   3465
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "附送書件"
      Height          =   855
      Left            =   2940
      TabIndex        =   50
      Top             =   3180
      Width           =   4575
      Begin VB.CheckBox chkAtt1 
         Caption         =   "基本資料表"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Tag             =   ".contact.pdf"
         Top             =   240
         Value           =   1  '核取
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "委任書"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Tag             =   ".poa.pdf"
         Top             =   495
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "優先權證明文件"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   14
         Tag             =   ".PRI.pdf"
         Top             =   255
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "移轉契約"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   15
         Tag             =   ".asasignment.pdf"
         Top             =   510
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "更名證明文件"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   16
         Tag             =   ".change.pdf"
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.TextBox Text6 
      Height          =   264
      Left            =   1425
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3180
      Width           =   372
   End
   Begin VB.TextBox textCP27 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4920
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8340
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6390
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7230
      TabIndex        =   19
      Top             =   70
      Width           =   1080
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1425
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2565
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020601_1.frx":0000
      Left            =   1260
      List            =   "frm03020601_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1162
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   0
      Top             =   510
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   1
      Top             =   510
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   2
      Top             =   510
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2655
      MaxLength       =   2
      TabIndex        =   3
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2565
      Width           =   300
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1425
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1485
      Left            =   180
      TabIndex        =   17
      Top             =   4095
      Width           =   8955
      _ExtentX        =   15790
      _ExtentY        =   2625
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7530
      TabIndex        =   7
      Top             =   2580
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;980"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblFee 
      Caption         =   "規費 :"
      Height          =   180
      Left            =   240
      TabIndex        =   51
      Top             =   3510
      Width           =   810
   End
   Begin VB.Label Label16 
      Caption         =   "延期月數 :"
      Height          =   180
      Left            =   210
      TabIndex        =   49
      Top             =   3225
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發文日期:"
      Height          =   180
      Left            =   4110
      TabIndex        =   48
      Top             =   2910
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6540
      TabIndex        =   47
      Top             =   2610
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   9120
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   9120
      Y1              =   2490
      Y2              =   2490
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   46
      Top             =   2130
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1260
      TabIndex        =   45
      Top             =   2130
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   300
      Index           =   8
      Left            =   2220
      TabIndex        =   44
      Top             =   2910
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "1931;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期　:"
      Height          =   180
      Left            =   210
      TabIndex        =   43
      Top             =   2610
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4020
      TabIndex        =   42
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   4020
      TabIndex        =   41
      Top             =   1806
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   40
      Top             =   1806
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   39
      Top             =   510
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4020
      TabIndex        =   38
      Top             =   1482
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   37
      Top             =   1482
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   36
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   35
      Top             =   834
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   4020
      TabIndex        =   34
      Top             =   834
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   33
      Top             =   1222
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   32
      Top             =   840
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   31
      Top             =   834
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1980
      TabIndex        =   30
      Top             =   1162
      Width           =   7110
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "12541;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   29
      Top             =   1484
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   28
      Top             =   1482
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   27
      Top             =   1806
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   26
      Top             =   1806
      Width           =   4200
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "7408;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   3240
      TabIndex        =   25
      Top             =   2610
      Width           =   2880
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "延期案件性質:"
      Height          =   180
      Left            =   210
      TabIndex        =   24
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "未收文期限　:"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   3795
      Width           =   1125
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   4020
      TabIndex        =   22
      Top             =   2130
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   210
      TabIndex        =   21
      Top             =   2130
      Width           =   765
   End
End
Attribute VB_Name = "frm03020601_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/04 Form2.0已修改; Label2(index)、lstNameAgent、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String
Dim tm() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_strNPReceiveNo As String '點選未收的期限的收文號
Dim m_CP43 As String 'Modify By Sindy 2014/3/5
Dim m_CP43cp10 As String 'Added by Lydia 2022/09/28 相關總收文號的案件性質
Dim m_CP10 As String 'Added by Lydia 2020/09/29 案件性質
'Added by Lydia 2019/03/26
Dim m_CP118  As String '是否電子送件
Dim m_CaseNo As String '電子送件-本所案號
Dim m_F21st07 As String 'FCT程序分機
Dim m_DocNo As String '機關文號
Dim m_NewCP07 As String '延期後的法限

Private Sub cmdok_Click(Index As Integer)
 Dim bolChk As Boolean, strTmp As String
'Added by Lydia 2019/02/21
Dim strFolder As String, strFileName As String
Dim strContent As String 'Added by Lydia 2019/08/2

   Select Case Index
      Case 0 '確定
         If Me.Text9.Text = "" Then
            MsgBox "請輸入延期案件性質或點選未收文期限資料!!!", vbExclamation + vbOKOnly
            Me.Text9.SetFocus
            Text9_GotFocus
            Exit Sub
         End If
         
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         Select Case Text9.Text
            Case "201" '補正
               strTmp = "01"
            Case "202" '申請意見書
               strTmp = "02"
            Case Else
         End Select
         'Added by Lydia 2019/03/26 +電子送件申請書=補正申請書
         If m_CP118 = "Y" Then
             strTmp = "10"
         End If
         'end 2019/03/26
         strLetterDate = Text5.Text
         If strTmp = "" Then
            MsgBox "該性質並無申請書！"
         'Added by Lydia 2019/03/26 電子送件申請書
         ElseIf m_CP118 = "Y" Then
            'Added by Lydia 2019/03/28
            m_DocNo = ""
            m_NewCP07 = ""
            StartLetter "90", strReceiveNo, strTmp '取得機關文號和法限
            EndLetter "90", strReceiveNo, strTmp, strUserNum
            'end 2019/03/28
            m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
            '桌面上建立案號資料夾
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
            End If

            '申請書
            If StartLetter2("90", strTmp, strReceiveNo) = False Then Exit Sub
            'Added by Lydia 2019/08/21 判斷要基本資料表,先不存檔
            If chkAtt1(0).Value = 1 Then
                 NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                 strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
            Else
            'end 2019/08/21
                NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
                Call PUB_MakeDoc(strContent, strFileName)
            End If
            
            '基本資料表
            'Move by Lydia 2019/08/21 從申請書上面移下來
            If chkAtt1(0).Value = 1 Then 'Added by Lydia 2019/04/11 若不勾選基本資料表不用產生.contact檔案
                'Modified by Lydia 2020/12/31 電子送件-基本資料表03=>11
                If StartLetter2("90", "11", strReceiveNo) = False Then Exit Sub
                'Modified by Lydia 2019/08/21 統一將基本資料表要和申請書放在同一份文件
                'NowPrint strReceiveNo, "90", "03", False, strUserNum, , , True, strContent
                'strFileName = strFolder & "\" & m_CaseNo & ".contact"
                'Call PUB_MakeDoc(strContent, strFileName)
                'Modified by Lydia 2020/12/31 電子送件-基本資料表03=>11
                NowPrint strReceiveNo, "90", "11", False, strUserNum, , strContent, True, strContent
                If strFileName = "" Then strFileName = strFolder & "\" & m_CaseNo & ".contact"
                'Modified by Lydia 2020/09/25 增加分節處理頁碼
                'Call PUB_MakeDoc(strContent, strFileName)
                strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
                Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
                'end 2019/08/21
                'end 2020/09/25
            End If
            
         'end 2019/03/26
         Else  '紙本申請書
            'StartLetter "90", Text1 & Text2 & Text3 & Text4 & "&303", strTmp
            'NowPrint Text1 & Text2 & Text3 & Text4 & "&303", "90", strTmp, bolChk, strUserNum
            StartLetter "90", strReceiveNo, strTmp
            NowPrint strReceiveNo, "90", strTmp, bolChk, strUserNum, 0
         End If
         frm030206_1.Show
         '回到原畫面要清除畫面
         frm030206_1.ClearForm
      Case 1 '回前畫面
         frm030206_1.Show
      Case 2 '結束
         Unload frm030206_1
   End Select
   Unload Me
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP07 As String
   
   EndLetter ET01, ET02, ET03, strUserNum
   ii = 0
   Select Case Me.Text9.Text
      Case "201" '補正
         '點選非延期案進入
         If Me.Text9.Enabled = False Then
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
         '點選延期案進入
         Else
            '若有點選未收文期限
            If m_strNPReceiveNo <> "" Then
                'StrSQLa = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
                StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & m_strNPReceiveNo & "') "
            '若未點選未收文期限
            Else
                'StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
                StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
            End If
         End If
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            Select Case "" & rsA("CP10").Value
            Case "101" '申請
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                    "','案件種類','註冊')"
            Case "102" '延展
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                    "','案件種類','延展註冊')"
            Case "301" '變更
               '判斷是否有審定號
               If Trim(Label12(2)) = "" Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                     "','案件種類','註冊前變更')"
               Else
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                     "','案件種類','註冊變更')"
               End If
            Case "501" '移轉
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                    "','案件種類','移轉登記')"
            Case "502" '授權
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                    "','案件種類','授權登記')"
            End Select
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
   End Select
   
   '點選非延期案進入
   If Me.Text9.Enabled = False Then
      ii = ii + 1
      'Modify By Sindy 2012/5/4
      'strCP07 = DateAdd("m", 1, ChangeWStringToWDateString(DBDATE(Label12(10))))
      strCP07 = DBDATE(DateAdd("m", Val(Text6), ChangeWStringToWDateString(DBDATE(Label12(10)))))
      '2012/5/4 End
      strCP07 = PUB_FCTGetDelaySpecDay(strCP07, m_CP43) 'Modify By Sindy 2014/3/5
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','法定期限','" & DBDATE(strCP07) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','機關文號','" & Label12(7) & "')"
      'Added by Lydia 2019/03/28
      m_DocNo = Label12(7)
      m_NewCP07 = strCP07
      
   '點選延期案進入
   Else
      StrSQLa = "Select * From NextProgress Where NP01=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') " & _
                        "AND NP02='" & tm(1) & "' AND NP03='" & tm(2) & "' AND NP04='" & tm(3) & "' AND NP05='" & tm(4) & "' " & _
                        "AND NP07=" & Me.Text9.Text & " " & _
                        "AND NP06 IS NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         ii = ii + 1
         'Modify By Sindy 2012/5/4
         'strCP07 = DateAdd("m", 1, ChangeWStringToWDateString("" & rsA("NP09").Value))
         strCP07 = DBDATE(DateAdd("m", Val(Text6), ChangeWStringToWDateString("" & rsA("NP09").Value)))
         strCP07 = PUB_FCTGetDelaySpecDay(strCP07, m_CP43) 'Modify By Sindy 2014/3/5
         '2012/5/4 End
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','法定期限','" & DBDATE(strCP07) & "')"
         m_NewCP07 = strCP07 'Added by Lydia 2019/03/28
      'Add By Sindy 2015/6/23 抓不到下一程序,則抓進度檔
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         StrSQLa = "Select CP07 From CaseProgress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','法定期限','" & DBDATE("" & rsA.Fields("cp07")) & "')"
               m_NewCP07 = "" & rsA.Fields("cp07") 'Added by Lydia 2019/03/28
         End If
      '2015/6/23 END
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      
      StrSQLa = "Select * From CaseProgress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','機關文號','" & "" & rsA("CP08").Value & "')"
         m_DocNo = "" & rsA.Fields("cp08") 'Added by Lydia 2019/03/28
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   
   'Add By Sindy 2016/5/31
   If tm(8) = "7" Then '7.證明標章
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','證明標章','證明標章')"
   End If
   '2016/5/31 END
   
   If ii <> 0 Then
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = tm(5)
      Case "英"
         Label12(3) = tm(6)
      Case "日"
         Label12(3) = tm(7)
   End Select
End Sub

Private Sub Form_Activate()
Me.Text7.SetFocus
End Sub

Private Sub Form_Load()
Dim tKind As String 'Added by Lydia 2019/03/26 特殊申請書

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm030206_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      tKind = .Text6.Text   'Added by Lydia 2019/03/26
      If tKind = "2" Then m_CP118 = "Y" 'Added by Lydia 2019/07/10
      strReceiveNo = .Tag
   End With
   ReDim tm(TF_TM)
   ReadTradeMark
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modified by Lydia 2021/08/04 傳入案件性質、Form 2.0
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   'Added by Lydia 2021/08/04 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1500
   lstNameAgent.Width = 1300
      
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
'   If Text9 = "201" Or Text9 = "202" Then
'      Text7 = "Y"
'   End If
   'Added by Lydia 2019/03/26 電子送件(外商是先收文延期才做申請書，所以是從延期進度進入by阿蓮)
   If tKind = "2" Then
       m_CP118 = "Y"
       Frame1.Visible = True
       
   Else
       Frame1.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020601_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsTemp1 As New ADODB.Recordset
'Modified by Lydia 2021/08/04
'Dim Lbl As LABEL
Dim Lbl As Object

   For Each Lbl In Label12
      Lbl = ""
   Next
   tm(1) = Text1
   tm(2) = Text2
   tm(3) = Text3
   tm(4) = Text4
   If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
      Text5 = tm(11)
      Label12(1) = tm(12)
      Label12(2) = tm(15)
      Label12(3) = tm(5)
   End If
   
   'Modified by Lydia 2019/03/26 +cp118,cp17,FCT程序分機
   'strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP84,CP110,CP118 " & _
      "from caseprogress,casepropertymap,staff,staff staff1 " & _
      "where cp09='" & strReceiveNo & "' " & _
      "AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) " & _
      "and cp13=staff1.st01(+) "
   strExc(0) = "select cpm03,s1.st02 as st1,s2.st02 as st2,cp43,cp10,cp06,cp07,cp84,cp110,cp118,cp17,s3.st07 " & _
                    "from caseprogress,casepropertymap,staff s1 ,staff s2,staff s3 " & _
                    "where cp09='" & strReceiveNo & "' " & _
                    "and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) " & _
                    "and cp13=s2.st01(+) and s2.st57=s3.st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      'Added by Lydia 2019/03/26
      txtFee.Text = Format(Val("" & .Fields("CP17")), "#,##0")
      m_CP118 = "" & .Fields("CP118")
      'Modified by Lydia 2019/07/10
      'If m_CP118 <> "" Then m_CP118 = "Y"
      If m_CP118 <> "" Then
         txtFee.Text = Val("" & .Fields("CP17"))
         m_CP118 = "Y"
      End If
      'end 2019/07/10
      m_F21st07 = "" & .Fields("st07") 'FCT程序分機
      'end 2019/03/26
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0) '案件性質
         If Label12(0).Caption <> "延期" Then Text9.Enabled = False
         '延期案件性質
         Text9 = .Fields(4)
         Label12(8).Caption = .Fields(0)
         If Me.Text9.Text = "303" Then
            Me.Text9.Text = ""
            Label12(8).Caption = ""
         End If
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1) '承辦人
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2) '智權人員
      If Not IsNull(.Fields(3)) Then
         '相關總收文號
         m_CP43 = .Fields(3) 'Add By Sindy 2012/3/7
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields("CP05")) Then Label12(6) = TransDate(rsTemp1.Fields("CP05"), 1) '來函收文日
            If Not IsNull(rsTemp1.Fields("CP08")) Then Label12(7) = rsTemp1.Fields("CP08") '機關文號
            m_CP43cp10 = "" & rsTemp1.Fields("cp10") 'Added by Lydia 2022/09/28
         End If
      End If
      If Not IsNull(.Fields(5)) Then Label12(9) = TransDate(.Fields(5), 1) '本所期限
      If Not IsNull(.Fields(6)) Then Label12(10) = TransDate(.Fields(6), 1) '法定期限
   End If
   End With
   
   If Label12(0).Caption = "延期" Then 'Modify By Sindy 2012/5/4 +if
      '抓本所案號相同且是否續辦為NULL的下一程序資料
      '2012/5/4 Modify By Sindy 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
      strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13,NP14," & SQLDate("NP11") & ",NP01,NP07 " & _
         "FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
         "WHERE " & ChgNextProgress(tm(1) & tm(2) & tm(3) & tm(4)) & _
         " AND NP06 IS NULL " & _
         " AND NP02=CPM01(+) " & _
         " AND NP07=CPM02(+) " & strNpSqlOfNoSalesDuty
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   End If
   GridHead
   
   'Add By Sindy 2012/3/7
   '帶延期案件性質
   strExc(0) = ""
   If Left(m_CP43, 1) = "C" Then
      strExc(0) = "select np07 from NEXTPROGRESS " & _
                  "where np01='" & m_CP43 & "' " & _
                  "and np06 is null "
'Modify By Sindy 2012/5/4 Mark
'   Else
'      strExc(0) = "select cp10 from caseprogress " & _
'                  "where cp43='" & strCP43 & "' " & _
'                  "and cp27 is null and cp57 is null "
   End If
   If strExc(0) <> "" Then
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields(0)) = False Then
            Text9 = RsTemp.Fields(0)
            Label12(8) = GetPrjState6(tm(1), Text9)
         End If
      End If
   End If
   
   'Add By Sindy 2012/5/4 預設延期月數
   If m_CP43 <> "" Then
      Text6 = GetDelayMonth(m_CP43)
      If Text6 = "" Then Text6.Text = "1"
   End If
   '2012/5/4 End
   
End Sub

Private Sub MSHFlexGrid1_Click()
Dim ii As Integer
    GridClick MSHFlexGrid1, intLastRow, 0
    If Me.Text9.Enabled = True Then
        Me.Text9.Text = ""
        Me.Label12(8).Caption = ""
        m_strNPReceiveNo = ""
        For ii = 1 To Me.MSHFlexGrid1.Rows - 1
            If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
               Me.Text9.Text = Me.MSHFlexGrid1.TextMatrix(ii, 8)
               Me.Label12(8).Caption = Me.MSHFlexGrid1.TextMatrix(ii, 1)
               m_strNPReceiveNo = Me.MSHFlexGrid1.TextMatrix(ii, 7)
               'Add By Sindy 2012/5/4 預設延期月數
               Text6 = GetDelayMonth(m_strNPReceiveNo)
               If Text6 = "" Then Text6.Text = "1"
               '2012/5/4 End
               Exit For
            End If
        Next ii
    End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

'Add by Sindy 2012/5/4
Private Sub Text6_GotFocus()
   TextInverse Text6
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
'End 2012/5/4

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1200: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1200: .Text = "解除期限日期"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 0: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 0: .Text = "下一程序"
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
Dim strTempName As String
   
    If Me.Text9.Text = "" Then Exit Sub
    If ClsPDGetCaseProperty("FCT", Text9, strTempName, False) Then
        Label12(8) = strTempName
    Else
        Label12(8) = ""
        Cancel = True
    End If
    If Cancel = True Then TextInverse Text9
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim strSqlText As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Or _
      m_strNPReceiveNo <> "" Or _
      textCP27 <> "" Then
      strSql = " UPDATE CASEPROGRESS SET "
      If lstNameAgent.Visible = True Then
         If strSqlText = "" Then
            strSqlText = " cp110=" & CNULL(m_CP110)
         Else
            strSqlText = strSqlText & " ,cp110=" & CNULL(m_CP110)
         End If
      End If
      If m_strNPReceiveNo <> "" Then
         If strSqlText = "" Then
            strSqlText = " cp43=" & CNULL(m_strNPReceiveNo)
         Else
            strSqlText = strSqlText & " ,cp43=" & CNULL(m_strNPReceiveNo)
         End If
      End If
      If textCP27 <> "" Then
         If strSqlText = "" Then
            strSqlText = " cp27=" & ChangeTStringToWString(textCP27)
         Else
            strSqlText = strSqlText & " ,cp27=" & ChangeTStringToWString(textCP27)
         End If
      End If
      strSql = strSql & strSqlText & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   'Added by Lydia 2019/03/26 預設為電子送件
   If m_CP118 = "Y" Then
        'Modified by Morgan 2019/7/17 目前FCT尚未自動扣款
        'strSql = " UPDATE CASEPROGRESS SET CP118='A' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
        strSql = " UPDATE CASEPROGRESS SET CP118='Y' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
        cnnConnection.Execute strSql
   End If
   'end 2019/03/26
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/08/04 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If

End Sub

'Add By Sindy 2010/4/16
Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

'Add By Sindy 2010/4/16
' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nick 2004/08/31
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

'Added by Lydia 2019/03/26 各式申請書-電子送件申請書
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String) As Boolean
   Dim strTxt(1 To 30) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2019/03/27
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '申請人資料
   'Modified by Lydia 2020/09/29 案件性質統一為延期303,而非收文號之案件性質
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, tm(), False)
   'Modified by Lydia 2023/11/08 原本預設抓申請人基本檔之地址;現在改成預設抓案件申請人資料之地址
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, "303", tm(), False)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, "303", tm(), True)
   
   '出名代理人
   'Modified by Lydia 2019/03/27 改成共用模組取得資料
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "FCT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If
   'end 2019/03/27
   
   If iET03 = "03" Then '基本資料表
        ii = ii + 1
        'FCT程序分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCT程序分機','" & m_F21st07 & "')"
   End If
   
   If iET03 = "10" Then '申請書
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & txtFee.Text & "')"
        strTmp = ""
        'Added by Lydia 2022/09/28 其對應之相關總收文號為「電話通知」時，申請書之申請內容第一點請帶：一、敬覆  鈞局XX年XX月XX日之電話通知。(日期為「電話通知」之收文日)
        If m_CP43 <> "" And m_CP43cp10 = "1727" Then
             'Modified by Lydia 2022/10/07
             'strTmp = "　　敬覆　鈞局" & Val(Left(Label12(6), 3)) & "年" & Val(Mid(Label12(6), 4, 2)) & "月" & Val(Right(Label12(6), 2)) & "日之電話通知。"
             strTmp = "  1. 敬覆　鈞局" & Val(Left(Label12(6), 3)) & "年" & Val(Mid(Label12(6), 4, 2)) & "月" & Val(Right(Label12(6), 2)) & "日之電話通知。"
        'end 2022/09/28
             'Added by Lydia 2023/07/05 相關總收文號為「電話通知」帶出申請內容第2點,其中內文日期為畫面上的法定期限+延期月數
             strTmp = strTmp & vbCrLf & "  2. 茲因申請人刻正蒐集相關資料中，不克於期限內補正，謹請　鈞局賜准延緩至" & Val(PUB_DBYEAR(m_NewCP07)) - 1911 & "年" & PUB_DBMONTH(m_NewCP07) & "月" & PUB_DBDAY(m_NewCP07) & "日補正，實感德便。"
        'Added by Lydia 2019/03/28 依紙本定稿的方式決定內文
        'Modified by Lydia 2022/09/28
        'If Me.Text9.Text = "202" Then '申請意見書(相關總收文的下一程序性質)
        ElseIf Me.Text9.Text = "202" Then '申請意見書(相關總收文的下一程序性質)
            'Modified by Lydia 2022/10/07
            'strTmp = "　　敬覆　鈞局" & m_DocNo & "核駁理由先行通知書，請准予延緩提出意見書期間至" & Val(PUB_DBYEAR(m_NewCP07)) - 1911 & "年" & PUB_DBMONTH(m_NewCP07) & "月" & PUB_DBDAY(m_NewCP07) & "日。"
                             strTmp = "  1. 敬覆　鈞局" & m_DocNo & "核駁理由先行通知書。" & vbCrLf
            strTmp = strTmp & "  2. 茲因申請人刻正蒐集相關資料中，不克於期限內提出意見書，謹請　鈞局賜准延緩至" & Val(PUB_DBYEAR(m_NewCP07)) - 1911 & "年" & PUB_DBMONTH(m_NewCP07) & "月" & PUB_DBDAY(m_NewCP07) & "日補呈意見書，實感德便。"
        Else
            'Modified by Lydia 2022/10/07
            'strTmp = "　　敬覆　鈞局" & m_DocNo & "函，請准予延緩補正期間至" & Val(PUB_DBYEAR(m_NewCP07)) - 1911 & "年" & PUB_DBMONTH(m_NewCP07) & "月" & PUB_DBDAY(m_NewCP07) & "日。"
                           strTmp = "  1. 敬覆　鈞局" & m_DocNo & "書函。" & vbCrLf
            strTmp = strTmp & "  2. 茲因申請人刻正蒐集相關資料中，不克於期限內補正，謹請　鈞局賜准延緩至" & Val(PUB_DBYEAR(m_NewCP07)) - 1911 & "年" & PUB_DBMONTH(m_NewCP07) & "月" & PUB_DBDAY(m_NewCP07) & "日補正，實感德便。"
        End If
        If strTmp <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1', '" & strTmp & "')"
        End If
        'end 2019/03/28
        
        '附送書件
        For intI = 0 To 4
             If chkAtt1(intI).Value = 1 Then
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(intI).Caption & "', '" & m_CaseNo & chkAtt1(intI).Tag & "')"
             End If
        Next intI
        'Added by Lydia 2019/04/11 若不勾選基本資料表，則附件名稱「未變更本案基本資料」並且不用產生.contact檔案
        If chkAtt1(0).Value = 0 Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(0).Caption & "', '未變更本案基本資料')"
        End If
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

