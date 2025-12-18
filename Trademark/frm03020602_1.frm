VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020602_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-催審, 暫緩審理, 減縮商品, 其他"
   ClientHeight    =   4272
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4272
   ScaleWidth      =   9300
   Begin VB.TextBox txtTM136 
      Height          =   270
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   44
      Top             =   2580
      Width           =   300
   End
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2895
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "附送書件"
      Height          =   1575
      Left            =   3690
      TabIndex        =   14
      Top             =   2580
      Width           =   2655
      Begin VB.CheckBox chkAtt1 
         Caption         =   "電子收據"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   46
         Tag             =   ".receipt.pdf"
         Top             =   735
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "基本資料表"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
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
         TabIndex        =   9
         Tag             =   ".poa.pdf"
         Top             =   495
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "優先權證明文件"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Tag             =   ".PRI.pdf"
         Top             =   735
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "移轉契約"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Tag             =   ".asasignment.pdf"
         Top             =   990
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "更名證明文件"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Tag             =   ".change.pdf"
         Top             =   1245
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8340
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6390
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7230
      TabIndex        =   16
      Top             =   70
      Width           =   1080
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2865
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2475
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020602_1.frx":0000
      Left            =   1260
      List            =   "frm03020602_1.frx":000D
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
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3255
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1:電子 2:紙本"
      Height          =   180
      Index           =   1
      Left            =   1500
      TabIndex        =   45
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "證書形式："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   43
      Top             =   2640
      Width           =   900
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7680
      TabIndex        =   13
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
      Left            =   180
      TabIndex        =   42
      Top             =   2940
      Width           =   810
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6690
      TabIndex        =   41
      Top             =   2640
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   150
      X2              =   9090
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   9090
      Y1              =   2460
      Y2              =   2460
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   40
      Top             =   2130
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1260
      TabIndex        =   39
      Top             =   2130
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期　:"
      Height          =   180
      Left            =   1650
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4020
      TabIndex        =   37
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   4020
      TabIndex        =   36
      Top             =   1806
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   35
      Top             =   1806
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   34
      Top             =   510
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
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
      TabIndex        =   33
      Top             =   1482
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   32
      Top             =   1482
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   31
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   30
      Top             =   834
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   4020
      TabIndex        =   29
      Top             =   834
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   28
      Top             =   1162
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   27
      Top             =   840
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   26
      Top             =   834
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1980
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   1484
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   23
      Top             =   1482
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   22
      Top             =   1806
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   21
      Top             =   1800
      Width           =   4140
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "7302;503"
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
      Left            =   180
      TabIndex        =   20
      Top             =   3300
      Width           =   2880
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   4020
      TabIndex        =   19
      Top             =   2130
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   210
      TabIndex        =   18
      Top             =   2130
      Width           =   765
   End
End
Attribute VB_Name = "frm03020602_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/04 Form2.0已修改; Label2(index)、lstNameAgent
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
'Memo by Lydia 2019/03/26 表單抬頭從"申請意見書, 其他"改為"催審, 暫緩審理, 減縮商品, 其他"
'Memo by Lydia 2021/02/05 增加案件性質：308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、717商標註冊費繳費單、725商標規費退費申請書(代辦退費)
Option Explicit

Dim strReceiveNo As String
Dim tm() As String, m_CP110 As String, m_AgentName As String
Dim m_CP43 As String 'Add By Sindy 2016/5/10
Dim m_CP10 As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_strNPReceiveNo As String '點選未收的期限的收文號
Dim m2_CP27 As String '相關總收文號-發文日期
'Added by Lydia 2019/03/26
Dim m_CP118  As String '是否電子送件
Dim m_CaseNo As String '電子送件-本所案號
Dim m_F21st07 As String 'FCT程序分機
Dim strAppDetail As String '申請內容
Dim m2_CP10 As String '相關總收文號-案件性質
Dim m2_CP10ex As String 'Added by Lydia 2023/11/30 相關總收文號-案件性質=>智慧局的指定名稱
Dim oObj As Control 'Added by Lydia 2023/11/30

Private Sub cmdok_Click(Index As Integer)
 Dim bolChk As Boolean, strTmp As String
'Added by Lydia 2019/02/21
Dim strFolder As String, strFileName As String
Dim strContent As String 'Added by Lydia 2019/08/21
Dim ET03 As String, ET03_1 As String 'Added by Lydia 2021/02/05

   Select Case Index
      Case 0 '確定
         
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         strTmp = "01"
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
            m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
            '桌面上建立案號資料夾
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
            End If
            'Addded by Lydia 2021/02/05 電子送件-定稿別
            '開放第三階段：308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、717商標註冊費繳費單、725商標規費退費申請書(代辦退費)
            If m_CP10 = "308" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "註冊申請案分割申請書"
            ElseIf m_CP10 = "313" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "註冊指定使用商品服務減縮申請書"
            ElseIf m_CP10 = "304" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "英文證明書申請書"
            'Added by Lydia 2024/08/06
            ElseIf m_CP10 = "309" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "中文證明書申請書"
            'end 2024/08/06
            ElseIf m_CP10 = "717" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "商標註冊費繳費單申請書"
            ElseIf m_CP10 = "725" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "商標規費退費申請書"
            'Added by Lydia 2023/01/05
            ElseIf m_CP10 = "729" Then
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "商標復權申請書"
            'Added by Lydia 2023/11/30 增加各商標種類的電子送件申請書
            ElseIf m_CP10 = "306" Then  '自請撤回
               If m2_CP10 = "101" Then  '註冊申請案
                  ET03 = "21"
               Else
                  ET03 = "22"
               End If
               ET03_1 = "11"  '一般-基本資料表
               strFileName = m2_CP10ex & "撤回申請書"
            ElseIf m_CP10 = "307" Then  '自請拋棄商標權
               ET03 = "21"
               ET03_1 = "11"  '一般-基本資料表
               strFileName = "商標權拋棄申請書"
            'end 2023/11/30
            Else
                ET03 = "10"
                ET03_1 = "11" '一般-基本資料表
                strFileName = "補正申請書-商簡A"
            End If
            'end 2021/02/05
            
            '申請書
            'Modified by Lydia 2021/02/05 改成變數
            'If StartLetter2("90", strTmp, strReceiveNo) = False Then Exit Sub
            If StartLetter2("90", ET03, strReceiveNo, "2") = False Then Exit Sub
            
            'Added by Lydia 2019/08/21 判斷要基本資料表,先不存檔
            If chkAtt1(0).Value = 1 Then
                 'Modified by Lydia 2021/02/05 改成變數
                 'NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                 'strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
                 NowPrint strReceiveNo, "90", ET03, False, strUserNum, , , True, strContent
                 strFileName = strFolder & "\" & m_CaseNo & "." & strFileName
                 'end 2021/02/05
            Else
            'end 2019/08/21
                'Modified by Lydia 2021/02/05 改成變數
                'NowPrint strReceiveNo, "90", strTmp, False, strUserNum, , , True, strContent
                'strFileName = strFolder & "\" & m_CaseNo & ".補正申請書-商簡A"
                NowPrint strReceiveNo, "90", ET03, False, strUserNum, , , True, strContent
                strFileName = strFolder & "\" & m_CaseNo & "." & strFileName
                'end 2021/02/05
                Call PUB_MakeDoc(strContent, strFileName)
            End If
            
            '基本資料表
            'Move by Lydia 2019/08/21 從申請書上面移下來
            If chkAtt1(0).Value = 1 Then 'Added by Lydia 2019/04/11 若不勾選基本資料表不用產生.contact檔案
                'Modified by Lydia 2020/12/31 電子送件-基本資料表03=>11
                'Modified by Lydia 2021/02/05 改成變數
                'If StartLetter2("90", "11", strReceiveNo) = False Then Exit Sub
                If StartLetter2("90", ET03_1, strReceiveNo, "1") = False Then Exit Sub
                'Modified by Lydia 2019/08/21 統一將基本資料表要和申請書放在同一份文件
                'NowPrint strReceiveNo, "90", "03", False, strUserNum, , , True, strContent
                'strFileName = strFolder & "\" & m_CaseNo & ".contact"
                'Call PUB_MakeDoc(strContent, strFileName)
                'Modified by Lydia 2020/12/31 電子送件-基本資料表03=>11
                'Modified by Lydia 2021/02/05 改成變數
                'NowPrint strReceiveNo, "90", "11", False, strUserNum, , strContent, True, strContent
                'If strFileName = "" Then strFileName = strFolder & "\" & m_CaseNo & ".contact"  '已不用
                NowPrint strReceiveNo, "90", ET03_1, False, strUserNum, , strContent, True, strContent
                'end 2021/02/05
                'Modified by Lydia 2020/09/25 增加分節處理頁碼
                'Call PUB_MakeDoc(strContent, strFileName)
                strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
                Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
                'end 2019/08/21
                'end 2020/09/25
            End If
            
         'end 2019/03/26
         Else '紙本
            'StartLetter "90", Text1 & Text2 & Text3 & Text4 & "&202", strTmp
            'NowPrint Text1 & Text2 & Text3 & Text4 & "&202", "90", strTmp, bolChk, strUserNum
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
Dim strCaseDate As String
Dim strCP43 As String, strCP10 As String, strCaseType As String
Dim strCP10_C2 As String 'Add By Sindy 2013/1/10
Dim strCP08 As String 'Add By Sindy 2016/5/10
   
   'Add By Sindy 2012/7/3
   StrSQLa = "Select C1.CP43,C2.CP10,C2.CP27,C3.CP10,C3.CP27,C2.CP43 From Caseprogress C1,Caseprogress C2,Caseprogress C3 Where C1.CP09='" & strReceiveNo & "' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) "
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strCP43 = "" & rsA.Fields(0)
      strCP10_C2 = "" & rsA.Fields(1) 'Add By Sindy 2013/1/10
      If strCP43 <> "" And Left(strCP43, 1) = "C" Then
         strCP10 = "" & rsA.Fields(3)
      Else
         strCP10 = "" & rsA.Fields(1)
      End If
   End If
   Select Case strCP10
      Case "101" '申請
         strCaseType = "註冊"
      Case Else
         Select Case strCP10
            Case "102" '延展
               strCaseType = "延展註冊"
            Case "301" '變更
               '判斷是否有審定號
               If Trim(Label12(2)) = "" Then
                  strCaseType = "註冊前變更"
               Else
                  strCaseType = "註冊變更"
               End If
            Case "501" '移轉
               strCaseType = "移轉登記"
            Case "502" '授權
               strCaseType = "授權登記"
         End Select
   End Select
   '2012/7/3 End
   
   EndLetter ET01, ET02, ET03, strUserNum
   ii = 0
      
   'Add By Sindy 2012/7/3
   If strCaseType <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','案件種類','" & ChgSQL(strCaseType) & "')"
   End If
   '2012/7/3 End
   
   Select Case m_CP10
      Case "202" '申請意見書
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','機關文號','" & Label12(7) & "')"
      'Add By Sindy 2012/6/1
      Case "305" '催審
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','主旨補充內文','謹請　鈞局儘速審理事。')"
      Case "310" '暫緩審理
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','主旨補充內文','謹請　鈞局暫緩審理事。')"
      Case "313" '減縮商品
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','主旨補充內文','減縮商品事。')"
      '2012/6/1 End
   End Select
   
   'Add By Sindy 2016/5/10
   strCP08 = Label12(7)
   If m_CP10 = "306" Then '自請撤回
      '抓出相關總收文號最大C類來函
      StrSQLa = "Select cp09,cp10,cp08" & _
                " From Caseprogress" & _
                " Where CP43='" & m_CP43 & "' AND substr(CP09,1,1)='C'" & _
                " order by cp05 desc"
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      strCP08 = ""
      If rsA.RecordCount > 0 Then
         strCP10_C2 = "" & rsA.Fields("cp10")
         strCP08 = "" & rsA.Fields("cp08")
      End If
   End If
   '2016/5/10 END
   
   strAppDetail = "" 'Added by Lydia 2019/03/26
   '申請意見書, 催審, 706.其他
   If m_CP10 <> "202" And m_CP10 <> "305" And m_CP10 <> "706" Then
      'If Trim(Label12(7)) = "" Then '無機關文號
      If Trim(strCP08) = "" Then '無機關文號
         strCaseDate = m2_CP27 '相關總收文號-發文日期
         If strCaseDate <> "" Then
            If Len(strCaseDate) = 6 Then strCaseDate = "0" & Trim(strCaseDate)
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','說明一','本案業於" & Val(Left(strCaseDate, 3)) & "年" & Mid(strCaseDate, 4, 2) & "月" & Right(strCaseDate, 2) & "日提出申請在案。')"
            strAppDetail = strAppDetail & "　　一、本案業於" & Val(Left(strCaseDate, 3)) & "年" & Mid(strCaseDate, 4, 2) & "月" & Right(strCaseDate, 2) & "日提出申請在案。" & vbCrLf 'Added by Lydia 2019/03/26
         Else
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','說明一','本案業於　年　月　日提出申請在案。')"
            strAppDetail = strAppDetail & "　　一、本案業於　年　月　日提出申請在案。" & vbCrLf 'Added by Lydia 2019/03/26
         End If
      Else
         ii = ii + 1
         'Modify By Sindy 2013/1/10
         If strCP10_C2 = "1202" Then '核駁前先行通知
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','說明一','敬覆　鈞局" & strCP08 & "核駁理由先行通知書。')"
            strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & strCP08 & "核駁理由先行通知書。" & vbCrLf 'Added by Lydia 2019/03/26
         Else
         '2013/1/10 End
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','說明一','敬覆　鈞局" & strCP08 & "函。')"
                strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & strCP08 & "函。" & vbCrLf 'Added by Lydia 2019/03/26
         End If
      End If
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
   
   Set rsA = Nothing
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
      tKind = .Text6  'Added by Lydia 2019/03/26
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
   
   'Added by Lydia 2019/03/26 電子送件
   If tKind = "2" Then
       m_CP118 = "Y"
       Frame1.Visible = True
       'Added by Lydia 2021/02/05
       txtFee = Format(txtFee, "#####0")
       'Added by Lydia 2021/05/31 717 商標註冊費:預設不附基本資料
       If m_CP10 = "717" Then
           chkAtt1(0).Value = False
       End If
       'end 2021/05/31
   Else
       Frame1.Visible = False
   End If
   
    'Added by Lydia 2022/12/28 判斷717註冊費繳費
    'Modified by Lydia 2023/01/05 增加729復權
    If InStr("717,729", m_CP10) > 0 Then
       Label4(0).Visible = True: Label4(1).Visible = True
       txtTM136.Visible = True
    Else
       Label4(0).Visible = False: Label4(1).Visible = False
       txtTM136.Visible = False
    End If
    'end 2022/12/28
    'Added by Lydia 2023/11/30 代辦退費(725):增加電子收據
    If m_CP10 = "725" Then
       chkAtt1(5).Left = chkAtt1(2).Left
       chkAtt1(5).Visible = True
    End If
    'end 2023/11/30
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020602_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsTemp1 As New ADODB.Recordset
Dim tmpArr As Variant 'Added by Lydia 2023/11/30
   
   'Modified by Lydia 2023/11/30 改用oObj
   For Each oObj In Label12
      oObj = ""
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
      txtTM136 = tm(136) 'Added by Lydia 2022/12/26
   End If
   
   'Modified by Lydia 2019/03/26 +cp118,cp17,FCT程序分機
   'strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP84,CP110 " & _
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
      txtFee = Format(Val("" & .Fields("CP17")), "#,##0")
      m_CP118 = "" & .Fields("CP118")
      'Modified by Lydia 2019/07/05 規費有千分位,會造成轉檔錯誤
      'If m_CP118 <> "" Then m_CP118 = "Y"
      If m_CP118 <> "" Then
         m_CP118 = "Y"
         txtFee = Val("" & .Fields("CP17"))
      End If
      'end 2019/07/05
      m_F21st07 = "" & .Fields("st07") 'FCT程序分機
      'end 2019/03/26
      m_CP43 = "" & .Fields("cp43") 'Add By Sindy 2016/5/10
      m_CP10 = "" & .Fields("CP10")
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0) '案件性質
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1) '承辦人
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2) '智權人員
      If Not IsNull(.Fields(3)) Then
         '相關總收文號
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields("CP05")) Then Label12(6) = TransDate(rsTemp1.Fields("CP05"), 1) '來函收文日
            If Not IsNull(rsTemp1.Fields("CP08")) Then Label12(7) = rsTemp1.Fields("CP08") '機關文號
            If Not IsNull(rsTemp1.Fields("CP27")) Then m2_CP27 = ChangeWStringToTString(rsTemp1.Fields("CP27")) '發文日期
            If Not IsNull(rsTemp1.Fields("CP10")) Then m2_CP10 = rsTemp1.Fields("CP10") 'Added by Lydia 2020/12/31相關總收文號-案件性質
            'Added by Lydia 2023/11/30 自請撤回(306):分申請案和非申請案
            If m_CP10 = "306" Then
               If m2_CP10 = "101" Then
                  m2_CP10ex = "註冊申請案自請"
               Else
                  '因為智慧局名稱有部份與CPM03不同，改成指定名稱;延展(102)、補證(103) 、註冊前變更(301)、註冊變更(301)、英證(304)、註冊前分割(308)、註冊後分割(308)、商品減縮(313)、
                                                                 '移轉(501)、授權(502)、再授權(504)、異議(601)、評定(602)、廢止(605)、代辦退費(725)、設定質權(506)
                  'Modified by Lydia 2024/08/09 +中證(309)
                  strExc(1) = "延展(102)、補證(103)、註冊前變更(3010)、註冊變更(3011)、英證(304)、中證(309)、" & _
                              "註冊前分割(3080)、註冊後分割(3081)、商品減縮(313)、移轉(501)、授權(502)、" & _
                              "再授權(504)、異議案(601)、評定案(602)、廢止案(605)、退費(725)、" & _
                              "質權(506)"
                  tmpArr = Empty
                  tmpArr = Split(strExc(1), "、")
                  For intI = 0 To UBound(tmpArr)
                     If Trim(tmpArr(intI)) <> "" And m2_CP10ex = "" Then
                        If InStr(Trim(tmpArr(intI)), m2_CP10 & IIf(InStr("301,308", m2_CP10) > 0, IIf(tm(15) <> "", "1", "0"), "")) > 0 Then
                           m2_CP10ex = Mid(Trim(tmpArr(intI)), 1, InStr(Trim(tmpArr(intI)), "(") - 1)
                           Exit For
                        End If
                     End If
                  Next intI
               End If
               If m2_CP10ex = "" Then
                  MsgBox "目前無【" & Label12(0) & PUB_GetRelateCasePropertyName(strReceiveNo, "1") & "】的電子送件申請書！", vbCritical + vbOKOnly, "自請撤回申請書"
                  cmdOK(0).Enabled = False
               End If
            End If
            'end 2023/11/30
         End If
      End If
      If Not IsNull(.Fields(5)) Then Label12(9) = TransDate(.Fields(5), 1) '本所期限
      If Not IsNull(.Fields(6)) Then Label12(10) = TransDate(.Fields(6), 1) '法定期限
   End If
   End With
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

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

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & " WHERE CP09='" & strReceiveNo & "'"
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

'Added by Lydia 2019/03/26 各式申請書-電子送件申請書
'Modified by Lydia 2021/02/05 +iKind = 1.基本資料表, 2.申請書
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
   Dim strTxt(1 To 30) As String, strTmp As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer, jj As Integer
   Dim strCP07 As String
   Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2019/03/27
   Dim intA As Integer 'Added by Lydia 2021/02/05
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '申請人資料
   'Modified by Lydia 2020/09/29 +案件性質
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, tm(), False)
   'Modified by Lydia 2023/11/08 原本預設抓申請人基本檔之地址;現在改成預設抓案件申請人資料之地址
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), True)
   
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
   
   'Modified by Lydia 2021/02/05
   'If iET03 = "03" Then '基本資料表
   If iKind = "1" Then
        ii = ii + 1
        'FCT程序分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCT程序分機','" & m_F21st07 & "')"
   End If
   
   'Modified by Lydia 2021/02/05
   'If iET03 = "10" Then '申請書
   If iKind = "2" Then
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & txtFee.Text & "')"
        'Added by Lydia 2022/12/19 註冊證形式
        If strSrvDate(1) >= "20230101" Then
           If tm(136) = "1" Then
              strExc(1) = "電子"
           ElseIf tm(136) = "2" Then
              strExc(1) = "紙本"
           Else
              strExc(1) = "電子/紙本"
           End If
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','註冊證形式','" & strExc(1) & "')"
        End If
        'end 2022/12/19
        
        'Added by Lydia 2020/12/31 相關總收文號-發文日期
        If m_CP10 <> "202" And m_CP10 <> "305" And m_CP10 <> "706" Then
            If Trim(Label12(7)) = "" Then '無機關文號
                If m2_CP27 <> "" Then
                    m2_CP27 = Val(m2_CP27) - 19110000
                    If Len(m2_CP27) = 6 Then m2_CP27 = "0" & Trim(m2_CP27)
                    strAppDetail = strAppDetail & "　　一、本案業於" & Val(Left(m2_CP27, 3)) & "年" & Mid(m2_CP27, 4, 2) & "月" & Right(m2_CP27, 2) & "日提出申請在案。" & vbCrLf
                Else
                    strAppDetail = strAppDetail & "　　一、本案業於　年　月　日提出申請在案。" & vbCrLf
                End If
            Else
                If m2_CP10 = "1202" Then '核駁前先行通知
                   strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Trim(Label12(7).Caption) & "核駁理由先行通知書。" & vbCrLf
                'Added by Lydia 2022/09/28 其對應之相關總收文號為「電話通知」時，申請書之申請內容第一點請帶：一、敬覆  鈞局XX年XX月XX日之電話通知。(日期為「電話通知」之收文日)
                ElseIf m2_CP10 = "1727" Then
                   strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Val(Left(Label12(6), 3)) & "年" & Val(Mid(Label12(6), 4, 2)) & "月" & Val(Right(Label12(6), 2)) & "日之電話通知。"
                'end 2022/09/28
                Else
                   strAppDetail = strAppDetail & "　　一、敬覆　鈞局" & Trim(Label12(7).Caption) & "函。" & vbCrLf
                End If
            End If
        End If
        'end 2020/12/31
        
        '申請內容
        jj = 0
        strTmp = ""
        'Modified by Lydia 2023/11/30 改用For Each
        For Each oObj In chkAtt1
          If oObj.Index > 0 Then
             If oObj.Value = 1 Then
                 jj = jj + 1
                 strTmp = strTmp & vbCrLf & "　　　　" & jj & ". " & oObj.Caption
             End If
          End If
        Next
        'Modified by Lydia 2020/12/26
        If strTmp <> "" Then
              ii = ii + 1
              If strAppDetail <> "" Then
                   strTmp = strAppDetail & "　　二、補正如下：" & strTmp
              Else
                   strTmp = "　　一、補正如下：" & strTmp
              End If
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1', " & CNULL(ChgSQL(strTmp)) & ")"
        End If
        
        'Added by Lydia 2022/04/22  沒有勾選附件也要新增記錄
        If strAppDetail <> "" And strTmp = "" Then
              ii = ii + 1
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1', " & CNULL(ChgSQL(strAppDetail)) & ")"
        End If
        'end 2022/04/22
        
        'Added by Lydia 2021/02/05 商品服務類別及名稱: 308.分割,313.減縮商品
        If m_CP10 = "308" Or m_CP10 = "313" Then
            strExc(1) = "": strExc(2) = "": strExc(3) = ""
            strExc(0) = BeforePrintGetDBData("TMGoods:" & tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4) & "-||區隔", True)
            strTmp = ""
            If Trim(strExc(0)) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(strExc(0), "||")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          '減縮商品
                          If m_CP10 = "313" Then
                                strExc(2) = strExc(2) & _
                                                 "【擬減縮商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(Mid(strExc(1), 1, InStr(strExc(1), "：") - 1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                                strExc(3) = strExc(3) & _
                                                 "【減縮後指定商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(Mid(strExc(1), 1, InStr(strExc(1), "：") - 1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          '分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                          ElseIf m_CP10 = "308" Then
                                strExc(2) = strExc(2) & _
                                                 "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                                 "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                          End If
                          jj = jj + 1
                          strTmp = strTmp & "," & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1)
                     End If
                 Next intA
            ElseIf tm(9) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(tm(9), ",")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          '減縮商品
                          If m_CP10 = "313" Then
                                strExc(2) = strExc(2) & _
                                                 "【擬減縮商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(strExc(1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                                strExc(3) = strExc(3) & _
                                                 "【減縮後指定商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(strExc(1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          '分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                          ElseIf m_CP10 = "308" Then
                                strExc(2) = strExc(2) & _
                                                 "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                                 "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          End If
                          jj = jj + 1
                          strTmp = strTmp & "," & strExc(1)
                     End If
                 Next intA
            Else
                '減縮商品
                If m_CP10 = "313" Then
                     strExc(2) = strExc(2) & _
                                      "【擬減縮商品或服務名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                     strExc(3) = strExc(3) & _
                                      "【減縮後指定商品或服務名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                '分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                ElseIf m_CP10 = "308" Then
                      strExc(2) = strExc(2) & _
                                       "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                       "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                       "　　【類別】　　　　　　　　　" & vbCrLf & _
                                       "　　【商品服務名稱】　　　　　" & vbCrLf
                End If
                strTmp = strTmp & ",01"
            End If
            ii = ii + 1
            If m_CP10 = "308" Then '分割
                 tmpArr1 = Empty
                 tmpArr1 = Split(Mid(strTmp, 2), ",")
                 For intA = 0 To UBound(tmpArr1)
                    If Trim(tmpArr1(intA)) <> "" Then
                        strExc(2) = strExc(2) & _
                                         "【分割後商品服務類別名稱或證明標的內容2】  " & vbCrLf & _
                                         "　　【分割序號】　　　　　　　02" & vbCrLf & _
                                         "　　【類別】　　　　　　　　　" & Trim(tmpArr1(intA)) & vbCrLf & _
                                         "　　【商品服務名稱】　　　　　" & vbCrLf
                    End If
                 Next intA
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','分割後商品服務類別名稱','" & ChgSQL(strExc(2)) & "')"
                 ii = ii + 1
                 '因為無法抓子案的明確件數，所以預設1件
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','分割件數','1')"
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','他案辦理日期','" & ChangeTStringToTDateString(strSrvDate(2)) & "')"
            ElseIf m_CP10 = "313" Then '減縮商品
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','擬減縮商品服務類別名稱','" & ChgSQL(strExc(2)) & "')"
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','減縮後指定商品服務類別名稱','" & ChgSQL(strExc(3)) & "')"
            End If
        'Modified by Lydia 2024/08/06 +309中文證明書 => Or m_CP10 = "309"
        ElseIf m_CP10 = "304" Or m_CP10 = "309" Then    '英文證明書
             strExc(0) = BeforePrintGetDBData("TMGoods:" & tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4) & "-中文", True)
             If strExc(0) <> "" Then
                 '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
                 If InStr(tm(9), ",") = 0 Then
                      strExc(0) = Mid(strExc(0), InStr(strExc(0), "：") + 1)
                 End If
                 If Trim(strExc(0)) <> "" Then
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱中文','" & ChgSQL(strExc(0)) & "')"
                 End If
             End If
             If m_CP10 = "304" Then 'Added by Lydia 2024/08/06
               strExc(1) = BeforePrintGetDBData("TMGoods:" & tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4) & "-英文", True)
               If strExc(1) <> "" Then
                  '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
                  If InStr(tm(9), ",") = 0 Then
                       strExc(1) = Mid(strExc(1), InStr(strExc(1), "：") + 1)
                  End If
                  If Trim(strExc(1)) <> "" Then
                      ii = ii + 1
                      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱英文','" & ChgSQL(strExc(1)) & "')"
                  End If
               End If
             End If 'Added by Lydia 2024/0/06
        ElseIf m_CP10 = "725" Then '代辦退費
            strTmp = "（　　）智商/慧商　　　　字第　　　　　　　　　　號函"
            If Label12(7).Caption <> "" Then strTmp = Label12(7).Caption
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','機關文號','" & strTmp & "')"
            '抓電子收據資枓(參考內商使用電子收據資料)
            If m_CP43 <> "" Then
               strSql = "select cp09,cp64 from caseprogress where cp09='" & m_CP43 & "' and instr(cp64,'收據號碼:')>0"
               intA = 1
               Set RsTemp = ClsLawReadRstMsg(intA, strSql)
               If intA = 1 Then
                  strTmp = ""
                  If InStr(RsTemp.Fields("cp64"), "收據號碼:") > 0 Then
                     strTmp = Mid(RsTemp.Fields("cp64"), InStr(RsTemp.Fields("cp64"), "收據號碼:") + 5, 11)
                  End If
                  If strTmp <> "" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','電子收據號碼','" & strTmp & "')"
                     '電子收據紀錄檔
                     strSql = "select er01,er03 from ereceipt where er01='" & strTmp & "'"
                     intA = 1
                     Set RsTemp = ClsLawReadRstMsg(intA, strSql)
                     If intA = 1 Then
                        strTmp = Val("" & RsTemp.Fields("er03"))
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','電子收據規費','" & strTmp & "')"
                     End If
                  End If
               End If
            End If   'If m_CP43 <> "" Then
        End If
        'end 2021/02/05
        
        '附送書件
        'Modified by Lydia 2023/11/30 改用For Each
        For Each oObj In chkAtt1
           If oObj.Value = 1 Then
             ii = ii + 1
             strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & oObj.Caption & "', '" & m_CaseNo & oObj.Tag & "')"
           End If
        Next
        
        'Added by Lydia 2019/04/11 若不勾選基本資料表，則附件名稱「未變更本案基本資料」並且不用產生.contact檔案
        If chkAtt1(0).Value = 0 Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(0).Caption & "', '未變更本案基本資料')"
        End If
        'Added by Lydia 2023/11/30
        If m_CP10 = "306" Then  '自請撤回->非申請案
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定名稱','" & m2_CP10ex & "')"
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定名稱2','" & Replace(m2_CP10ex, "案", "") & "')"
        End If
        'end 2023/11/30
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

