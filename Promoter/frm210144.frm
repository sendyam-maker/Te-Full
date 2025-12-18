VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210144 
   BorderStyle     =   1  '單線固定
   Caption         =   "寄發文件"
   ClientHeight    =   5772
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8928
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      BorderStyle     =   0  '沒有框線
      Height          =   375
      Left            =   1140
      TabIndex        =   31
      Top             =   1500
      Visible         =   0   'False
      Width           =   2295
      Begin MSForms.TextBox txtInput 
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1635
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "2884;661"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   4410
      TabIndex        =   27
      Top             =   0
      Width           =   4515
      Begin VB.CommandButton cmdBrowser 
         BackColor       =   &H0000FF00&
         Caption         =   "點我展開"
         Height          =   345
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   29
         Top             =   0
         Width           =   4515
      End
      Begin VB.ComboBox cboAtt 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm210144.frx":0000
         Left            =   900
         List            =   "frm210144.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   28
         Top             =   330
         Width           =   3645
      End
      Begin VB.Label lblAttCnt 
         Alignment       =   2  '置中對齊
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  '單線固定
         Caption         =   " PDF:(0)"
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
         Height          =   300
         Left            =   -15
         TabIndex        =   30
         Top             =   330
         Width           =   930
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5115
      Left            =   4410
      TabIndex        =   0
      Top             =   630
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   9022
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   3420
      TabIndex        =   5
      Top             =   0
      Width           =   960
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1785
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   4365
      _ExtentX        =   7684
      _ExtentY        =   3154
      _Version        =   393216
      Cols            =   9
      FixedCols       =   4
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "親送|寄送|不寄|實體|本所案號|案件性質|案件名稱|收文日|備註"
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
      _Band(0).Cols   =   9
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2175
      Left            =   30
      TabIndex        =   7
      Top             =   3180
      Width           =   4365
      _ExtentX        =   7684
      _ExtentY        =   3831
      _Version        =   393216
      Cols            =   9
      FixedCols       =   2
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|案件性質|案件名稱|發文室發文日|本所期限|備註"
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
      _Band(0).Cols   =   9
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   765
      Left            =   -45
      TabIndex        =   8
      Top             =   2400
      Width           =   4515
      Begin VB.CommandButton CmdOk1 
         Caption         =   "更換"
         Height          =   315
         Index           =   3
         Left            =   2448
         TabIndex        =   36
         Top             =   150
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton CmdOk1 
         Caption         =   "下載"
         Height          =   315
         Index           =   2
         Left            =   1896
         TabIndex        =   35
         Top             =   150
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton CmdOk1 
         Caption         =   "進度"
         Height          =   315
         Index           =   0
         Left            =   756
         TabIndex        =   20
         Top             =   150
         Width           =   540
      End
      Begin VB.CommandButton CmdOk1 
         Caption         =   "卷宗區"
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   18
         Top             =   150
         Width           =   660
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "EMail"
         Height          =   315
         Index           =   1
         Left            =   1296
         TabIndex        =   10
         Top             =   150
         Width           =   600
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "確認"
         Height          =   315
         Index           =   1
         Left            =   3624
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "直寄："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   510
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(不寄--請取回紙本。)"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Index           =   1
         Left            =   2745
         TabIndex        =   16
         Top             =   510
         Width           =   1680
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   3180
         TabIndex        =   11
         Top             =   216
         Width           =   312
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   -45
      TabIndex        =   12
      Top             =   5250
      Width           =   4470
      Begin VB.CommandButton CmdOk1 
         Caption         =   "進度"
         Height          =   315
         Index           =   1
         Left            =   900
         TabIndex        =   21
         Top             =   150
         Width           =   660
      End
      Begin VB.CommandButton CmdOk1 
         Caption         =   "卷宗區"
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   19
         Top             =   150
         Width           =   800
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "EMail"
         Height          =   315
         Index           =   2
         Left            =   1575
         TabIndex        =   14
         Top             =   150
         Width           =   600
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "確認"
         Height          =   315
         Index           =   2
         Left            =   2880
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   2
         Left            =   2385
         TabIndex        =   15
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "重整"
      Height          =   345
      Left            =   2760
      TabIndex        =   22
      Top             =   0
      Width           =   600
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "電子證書"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   2
      Left            =   2295
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2070
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   930
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "3175;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "全E化"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   3900
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H000080FF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3675
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3090
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "E化"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   3315
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "確認人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(雙擊預覽)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   270
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "非直寄(親送/寄送/不寄)："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   2010
   End
End
Attribute VB_Name = "frm210144"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/18 改成Form2.0 (MSHFlexGrid1,MSHFlexGrid2,txtInput,Combo1)
'Created by Morgan 2014/3/27
'Memo by Lydia 2019/07/01 表單名稱:文件寄送確認=>寄發文件
Option Explicit

Dim iPrevRow1 As Integer, iPrevRow2 As Integer '前次點選列
Dim lTotRows1 As Long, lSelRows1 As Long, lTotRows2 As Long, lSelRows2 As Long
Dim m_blnColOrderAsc1 As Boolean, m_blnColOrderAsc2 As Boolean '欄位資料由小到大排序

Dim m_AttachPath As String
Dim m_AttachPath2 As String 'Added by Morgan 2020/2/18
Dim oFileSys As New FileSystemObject
Dim oFile As File

Dim m_bolBranch As Boolean
Dim m_InputGrid As MSHFlexGrid, m_InputCol As Integer, m_InputRow As Integer
Public cmdState As Integer '紀錄作用按鍵
Dim m_MailDoneList1 As String, m_MailDoneList2 As String 'Added by Morgan 2015/6/16 已EMail記錄
Dim m_idxECase1 As Integer, m_idxECase2 As Integer 'Added by Morgan 2015/6/16
Dim m_idxPaper As Integer 'Added by Morgan 2018/10/24 是否有紙本
Dim m_AttachPath3 As String 'Added by Morgan 2025/5/6 Confidential資料夾


'Added by Morgan 2019/3/4
Private Sub cboAtt_Click()
   Dim hLocalFile As Long
   Dim arrFileName() As String
   
   arrFileName = Split(cboAtt.List(cboAtt.ListIndex), Chr(9))
   'Modified by Morgan 2020/2/18
   'WebBrowser1.Navigate m_AttachPath & "\" & arrFileName(0): DoEvents
   'Modified by Morgan 2021/5/7
   'WebBrowser1.Navigate m_AttachPath2 & "\" & arrFileName(0): DoEvents
   strExc(1) = m_AttachPath2 & "\" & arrFileName(0)
   strExc(2) = strExc(1) & "_Copy"
   If Dir(strExc(2)) = "" Then
      FileCopy strExc(1), strExc(2)
   End If
   WebBrowser1.Navigate strExc(2): DoEvents
   'end 2021/5/7
End Sub

Private Sub cmdEmail_Click(Index As Integer)
   If Index = 1 Then
      fnOpen1 2
   Else
      fnOpen2 2
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim iRow As Integer, bContinue As Boolean
   Dim idxCP09 As Integer  'Added by Morgan 2015/6/16
   Dim stCheck As String 'Added by Morgan 2022/7/4
   
   SetMouseBusy
   bContinue = False
   If Index = 1 Then
      idxCP09 = GetFieldId("CP09", MSHFlexGrid1) 'Added by Morgan 2015/6/16
      With MSHFlexGrid1
      For iRow = 1 To .Rows - 1
         stCheck = .TextMatrix(iRow, 0) & .TextMatrix(iRow, 1) & .TextMatrix(iRow, 2)
         'Modified by Morgan 2022/7/4 +只E不寄紙本
         'If stCheck = "V" Then
         If stCheck = "V" Or stCheck = "E" Then
         'end 2022/7/4
            bContinue = True
            'Mofieied by Morgan 2015/6/16
            'E化案件檢查
            'Exit For
            If .TextMatrix(iRow, m_idxECase1) = "Y" Then
               'Modified by Morgan 2021/9/30 改檢查LP39
               'If InStr(m_MailDoneList1, .TextMatrix(iRow, idxCP09)) = 0 Then
               If ChkEMailRec(.TextMatrix(iRow, idxCP09)) = False Then
               'end 2021/9/30
                  If MsgBox(.TextMatrix(iRow, 4) & "為E化案件但尚未EMail，是否仍要繼續確認？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                     SetMouseReady
                     Exit Sub
                  End If
               End If
            End If
            'end 2015/6/16
         End If
      Next
      End With
   Else
      idxCP09 = GetFieldId("CP09", MSHFlexGrid2) 'Added by Morgan 2015/6/16
      With MSHFlexGrid2
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "V" Then
            bContinue = True
            'Mofieied by Morgan 2015/6/16
            'E化案件檢查
            'Exit For
            If .TextMatrix(iRow, m_idxECase2) = "Y" Then
               'Modified by Morgan 2021/9/30 改檢查LP39
               'If InStr(m_MailDoneList2, .TextMatrix(iRow, idxCP09)) = 0 Then
               If ChkEMailRec(.TextMatrix(iRow, idxCP09)) = False Then
               'end 2021/9/30
                  If MsgBox(.TextMatrix(iRow, 1) & "為E化案件但尚未EMail，是否仍要繼續確認？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                     SetMouseReady
                     Exit Sub
                  End If
               End If
            End If
            'end 2015/6/16
         End If
      Next
      End With
   End If
   
   If bContinue = False Then
      MsgBox "請先勾選(V)要確認的資料列！", vbInformation
   ElseIf Index = 1 Then
      FormSave1
   Else
      FormSave2
   End If
   SetMouseReady
End Sub

Private Sub fnOpen1(Index As Integer)
   Dim stFiles As String
   Dim stFileName As String
   Dim arrFileName() As String
   Dim idx As Integer, stCP09 As String
   Dim bJoin As Boolean
   Dim stFileNameList As String
   Dim bDone As Boolean 'Added by Morgan 2015/6/16.
   Dim stFileDescs As String 'Added by Morgan 2019/3/4
   Dim strNo As String 'Add by Amy 2019/11/05
   
   If iPrevRow1 = 0 Then
      'Modified by Morgan 2014/12/18
      'MsgBox "請先點選欲" & IIf(Index = 1, "預覽", "開啟") & "的資料列！", vbInformation
      MsgBox "請先點選欲" & IIf(Index = 1, "預覽", " Email ") & "的資料列！", vbInformation
      'end 2014/12/18
   'Added by Morgan 2018/10/24
   ElseIf Index = 2 And MSHFlexGrid1.TextMatrix(iPrevRow1, m_idxECase1) = "E" Then
      MsgBox "全E化客戶案件，不可在本作業 EMail！" & vbCrLf & vbCrLf & "※ 確認後統一由程序 EMail", vbExclamation
   'end 2018/10/24
   Else
      SetMouseBusy
      With MSHFlexGrid1
      If .TextMatrix(iPrevRow1, 4) <> "" Then
         stFiles = ""
         If Index = 1 Then
            bJoin = True
         Else
            'ClosePDF
            bJoin = False
         End If
        
         '檢查客戶函是否已轉檔
         intI = GetFieldId("cpp02", MSHFlexGrid1)
         If .TextMatrix(iPrevRow1, intI) = "" Then
            MsgBox "轉檔未完成，請稍後再試！", vbExclamation
            SetMouseReady
            Exit Sub
         End If
         stCP09 = GetValue(iPrevRow1, "cp09", MSHFlexGrid1)
         'Modified by Morgan 2016/3/2 改呼叫共用
         'If GetAttachFile(stCP09, stFiles, m_AttachPath, bJoin, GetValue(iPrevRow1, "cp10", MSHFlexGrid1)) = True Then
         'Modified by Morgan 2019/3/4 +stFileDescs
         'Modified by Morgan 2020/2/18
         'If PUB_GetAttachFile4Cust(stCP09, stFiles, m_AttachPath, bJoin, GetValue(iPrevRow1, "cp10", MSHFlexGrid1), stFileDescs) = True Then
         m_AttachPath2 = m_AttachPath & "\" & stCP09
         If PUB_GetAttachFile4Cust(stCP09, stFiles, m_AttachPath2, bJoin, GetValue(iPrevRow1, "cp10", MSHFlexGrid1), stFileDescs) = True Then
         'end 2020/2/17
         
            If stFileDescs <> "" Then
               SetValue iPrevRow1, "FDesc", stFileDescs, MSHFlexGrid1
            Else
               stFileDescs = GetValue(iPrevRow1, "FDesc", MSHFlexGrid1)
            End If
            
            arrFileName = Split(stFiles, ";")
            For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
               If arrFileName(idx) <> "" Then
                  'Modified by Morgan 2020/2/18
                  'stFileName = m_AttachPath & "\" & arrFileName(idx)
                  stFileName = m_AttachPath2 & "\" & arrFileName(idx)
                  'end 2020/2/17
                  If Index = 1 Then
                     WebBrowser1.Navigate stFileName
                     SetAttList stFileDescs
                     Exit For
                  Else
                     'Modified by Morgan 2014/12/18
                     'OpenPdf stFileName
                     stFileNameList = stFileName & ";" & stFileNameList
                     'end 2014/12/18
                  End If
               End If
            Next
            
            'Added by Morgan 2014/12/18
            If Index = 1 Then
            'end 2014/12/18
               '上已讀取註記
               SetValue iPrevRow1, "Read", "Y", MSHFlexGrid1
            'Added by Morgan 2014/12/18
            'Modified by Morgan 2021/5/7
            'ElseIf stFileNameList <> "" Then
            ElseIf PUB_ChkAttFile(stFileNameList) = True Then
            'end 2021/5/7
               SetMouseReady
               'Modified by Morgan 2015/6/16
               'Modify by Amy 2019/11/05 T字頭使用PUB_SettingTeMail
               strNo = GetValue(iPrevRow1, "CaseNo", MSHFlexGrid1)
               'Modify By Sindy 2021/1/14 + (PUB_GetST03(strUserNum) = "P21" Or Trim(Left(Combo1.Text, 6)) = "P2002")
               'Modify By Sindy 2023/2/6 改檢查屬MCT的案件 有TM44代理人為台灣案
'               If Left(strNo, 1) = "T" And _
'                  (PUB_GetST03(strUserNum) = "P21" Or Trim(Left(Combo1.Text, 6)) = "P2002") Then
               If Left(strNo, 1) = "T" And _
                  (GetPrjPeopleNum6(.TextMatrix(iPrevRow1, 4)) <> "" And GetPrjNation1(.TextMatrix(iPrevRow1, 4)) = "000") Then
               '2023/2/6 END
                  If strSrvDate(1) >= T商標電子化啟用日 Then
                     
                     'Added by Morgan 2025/8/20
                     'MCT案件要自動抓存在公用電腦的請款單(請款單號.invoice.pdf)
                     strExc(1) = GetValue(iPrevRow1, "cp10", MSHFlexGrid1)
                     If strExc(1) = "1101" Then
                        strExc(0) = "select a1k01 from caseprogress a,caseprogress b,acc1k0,staff where a.cp09='" & stCP09 & "' and b.cp09(+)=a.cp43 and a1k01(+)=b.cp60 and st01(+)=a1k21 and st03 like 'P2%'"
                     Else
                        strExc(0) = "select a1k01 from caseprogress,acc1k0,staff where cp09='" & stCP09 & "' and a1k01(+)=cp60 and st01(+)=a1k21 and st03 like 'P2%'"
                     End If
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(1) = RsTemp(0) & ".invoice.pdf"
                        strExc(2) = Pub_GetSpecMan("MCTInvoicePath") & "\" & strExc(1)
                        If Dir(strExc(2)) <> "" Then
                           strExc(1) = m_AttachPath2 & "\" & strExc(1)
                           FileCopy strExc(2), strExc(1)
                           stFileNameList = stFileNameList & strExc(1) & ";"
                        End If
                     End If
                     'end 2025/8/20
                     
                     'Add By Sindy 2024/10/16 + bolReadLP42=True
                     bDone = PUB_SettingTeMail(Me, PUB_DownloadOftPath("F23", ""), Mid(strNo, 1, Len(strNo) - 9), Mid(strNo, Len(strNo) - 8, 6), Mid(strNo, Len(strNo) - 2, 1), Mid(strNo, Len(strNo) - 1, Len(strNo)), _
                                         stFileNameList, GetValue(iPrevRow1, "cp10", MSHFlexGrid1), stCP09, , , , , stCP09, , , , , , True)
                  End If
               Else
                  'Modified by Morgan 2021/9/1 Email內文要帶出法定期限及本所期限，年費則多列待繳費年度(先做D類)
                  'PUB_ShowMailForm stCP09, stFileNameList, GetValue(iPrevRow1, "案件性質", MSHFlexGrid1), bDone, , , , True, , , , , , , , , , , , stCP09
                  PUB_SetDateAndFeeYear stCP09, strExc(6), strExc(7), strExc(8)
                  strExc(9) = GetValue(iPrevRow1, "案件性質", MSHFlexGrid1) & IIf(strExc(8) = "", "", " [ " & strExc(8) & " ] ")
                  'Add By Sindy 2024/10/16 + bolReadLP42=True
                  PUB_ShowMailForm stCP09, stFileNameList, strExc(9), bDone, , strExc(6), strExc(7), True, , , , , , , , , , , , stCP09, , , True
                  'end 2021/9/1
               End If
               'end 2019/11/05
               If bDone = True Then
                  m_MailDoneList1 = m_MailDoneList1 & ";" & stCP09
               End If
               'end 2015/6/16
               SetMouseBusy
            End If
            'end 2014/12/18
         End If
      End If
      End With
   End If
   SetMouseReady
End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
   MSHFlexGrid2.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
   MSHFlexGrid2.MousePointer = vbDefault
End Sub

Private Sub fnOpen2(Index As Integer)
   Dim stFiles As String
   Dim stFileName As String
   Dim arrFileName() As String
   Dim idx As Integer, stCP09 As String
   Dim bJoin As Boolean
   Dim stFileNameList As String
   Dim bDone As Boolean 'Added by Morgan 2015/6/16
   Dim stFileDescs As String 'Added by Morgan 2019/3/4
   Dim strNo As String 'Add by Amy 2019/11/05
   
   If iPrevRow2 = 0 Then
      MsgBox "請先點選欲" & IIf(Index = 1, "預覽", "開啟") & "的資料列！", vbInformation
   'Added by Morgan 2018/10/24
   ElseIf Index = 2 And MSHFlexGrid2.TextMatrix(iPrevRow2, m_idxECase2) = "E" Then
      MsgBox "全E化客戶案件，不可在本作業 EMail！" & vbCrLf & vbCrLf & "※ 統一由程序 EMail", vbExclamation
   'end 2018/10/24
   Else
      SetMouseBusy
      With MSHFlexGrid2
      If .TextMatrix(iPrevRow2, 1) <> "" Then
         stFiles = ""
         If Index = 1 Then
            bJoin = True
         Else
            'ClosePDF
            bJoin = False
         End If
         
         '檢查客戶函是否已轉檔
         If GetValue(iPrevRow2, "cpp02", MSHFlexGrid2) = "" Then
            MsgBox "轉檔未完成，請稍後再試！", vbExclamation
            SetMouseReady
            Exit Sub
         End If
         stCP09 = GetValue(iPrevRow2, "cp09", MSHFlexGrid2)
         'Modified by Morgan 2016/3/2 改呼叫共用
         'If GetAttachFile(stCP09, stFiles, m_AttachPath, bJoin, GetValue(iPrevRow2, "cp10", MSHFlexGrid2)) = True Then
         'Modified by Morgan 2020/2/18 暫存檔已改用本所號,遇同案號不同收文號,因檔名相同會不再下載故路徑再加收文號
         'If PUB_GetAttachFile4Cust(stCP09, stFiles, m_AttachPath, bJoin, GetValue(iPrevRow2, "cp10", MSHFlexGrid2), stFileDescs) = True Then
         m_AttachPath2 = m_AttachPath & "\" & stCP09
         If PUB_GetAttachFile4Cust(stCP09, stFiles, m_AttachPath2, bJoin, GetValue(iPrevRow2, "cp10", MSHFlexGrid2), stFileDescs) = True Then
         'end 2020/2/18
            If stFileDescs <> "" Then
               SetValue iPrevRow2, "FDesc", stFileDescs, MSHFlexGrid2
            Else
               stFileDescs = GetValue(iPrevRow2, "FDesc", MSHFlexGrid2)
            End If
            
            arrFileName = Split(stFiles, ";")
            For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
               If arrFileName(idx) <> "" Then
                  'Modified by Morgan 2020/2/18
                  'stFileName = m_AttachPath & "\" & arrFileName(idx)
                  stFileName = m_AttachPath2 & "\" & arrFileName(idx)
                  'end 2020/2/18
                  If Index = 1 Then
                     WebBrowser1.Navigate stFileName
                     SetAttList stFileDescs
                     Exit For
                  Else
                     'Modified by Morgan 2014/12/18
                     'OpenPdf stFileName
                     stFileNameList = stFileName & ";" & stFileNameList
                     'end 2014/12/18
                  End If
               End If
            Next
            'Added by Morgan 2014/12/18
            If Index = 1 Then
            'end 2014/12/18
               '上已讀取註記
               SetValue iPrevRow2, "Read", "Y", MSHFlexGrid2
            'Added by Morgan 2014/12/18
            'Modified by Morgan 2021/5/7
            'Else
            ElseIf PUB_ChkAttFile(stFileNameList) = True Then
            'end 2021/5/7
               SetMouseReady
               'Modified by Morgan 2015/6/16
               'Modify by Amy 2019/11/05 T字頭使用PUB_SettingTeMail
               'Modified by Morgan 2021/1/13 更正變數 iPrevRow1-->iPrevRow2,MSHFlexGrid1-->MSHFlexGrid2
               strNo = GetValue(iPrevRow2, "CaseNo", MSHFlexGrid2)
               'Modify By Sindy 2021/1/14 (PUB_GetST03(strUserNum) = "P21" Or Trim(Left(Combo1.Text, 6)) = "P2002")
               'Modify By Sindy 2023/2/6 改檢查屬MCT的案件 有TM44代理人為台灣案
'               If Left(strNo, 1) = "T" And _
'                  (PUB_GetST03(strUserNum) = "P21" Or Trim(Left(Combo1.Text, 6)) = "P2002") Then
               If Left(strNo, 1) = "T" And _
                  (GetPrjPeopleNum6(.TextMatrix(iPrevRow2, 1)) <> "" And GetPrjNation1(.TextMatrix(iPrevRow2, 1)) = "000") Then
               '2023/2/6 END
                  If strSrvDate(1) >= T商標電子化啟用日 Then
                      'Add By Sindy 2024/10/16 + bolReadLP42=True
                      bDone = PUB_SettingTeMail(Me, PUB_DownloadOftPath("F23", ""), Mid(strNo, 1, Len(strNo) - 9), Mid(strNo, Len(strNo) - 8, 6), Mid(strNo, Len(strNo) - 2, 1), Mid(strNo, Len(strNo) - 1, Len(strNo)), _
                                          stFileNameList, GetValue(iPrevRow2, "cp10", MSHFlexGrid2), stCP09, , , , , stCP09, , , , , , True)
                  End If
               Else
                  'Modified by Morgan 2021/6/16 直寄也要更新EMail記錄
                  'Modified by Morgan 2021/9/1 Email內文要帶出法定期限及本所期限，年費則於案件性質後加待繳費年度(先做D類)
                  'PUB_ShowMailForm stCP09, stFileNameList, GetValue(iPrevRow2, "案件性質", MSHFlexGrid2), bDone, , , , True, , , , , , , , , , , , stCP09
                  PUB_SetDateAndFeeYear stCP09, strExc(6), strExc(7), strExc(8)
                  strExc(9) = GetValue(iPrevRow2, "案件性質", MSHFlexGrid2) & IIf(strExc(8) = "", "", " [ " & strExc(8) & " ] ")
                  'Add By Sindy 2024/10/16 + bolReadLP42=True
                  PUB_ShowMailForm stCP09, stFileNameList, strExc(9), bDone, , strExc(6), strExc(7), True, , , , , , , , , , , , stCP09, , , True
                  'end 2021/9/1
               End If
               'end 2019/11/05
               If bDone = True Then
                  m_MailDoneList2 = m_MailDoneList2 & ";" & stCP09
               End If
               'end 2015/6/16
               SetMouseBusy
            End If
            'end 2014/12/18
         End If
      End If
      End With
      SetMouseReady
   End If
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index '紀錄作用按鍵
   PubShowNextData
   cmdState = -1
   Exit Sub
End Sub

Private Sub cmdRefresh_Click()
   'Added by Morgan 2019/3/4
   WebBrowser1.Navigate "about:blank": DoEvents
   SetAttList
   'end 2019/3/4
   
   KillTemp 'Added by Morgan 2018/7/9 前次開啟的客戶函要清除
   SetMouseBusy
   OpenTable1
   OpenTable2
   SetMouseReady
End Sub

Private Sub cmdBrowser_Click()
   If cmdBrowser.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      'Modified by Morgan 2014/12/17
      'cmdRefresh(1).Value = True
      'cmdRefresh(2).Value = True
      cmdRefresh.Value = True
      Combo1.Tag = Combo1
   End If
End Sub

Private Sub Form_Activate()
   Static bDone As Boolean
   
   If Screen.ActiveForm.Name <> Me.Name Then Exit Sub 'Added by Morgan 2021/7/23
   
   If Me.WindowState = 0 Then Me.WindowState = 2 'Added by Morgan 2014/5/29
   If bDone = False Then
      Combo1_Click
      bDone = True
   End If
   
End Sub

Private Sub Form_Load()
   'Removed by Morgan 2014/5/14
   'strExc(1) = GetLocalIP
   'If InStr(strExc(1), "192.168.") = 1 Then
   '   If Val(Mid(strExc(1), 9)) > 1 Then
   '      m_bolBranch = True
   '   End If
   'End If
   'Removed by Morgan 2022/2/18 改以選單的確認人員判斷所別
   'If pub_strUserOffice > "1" And pub_strUserOffice < "5" Then
   '   m_bolBranch = True
   'End If
   'end 2014/5/14
   
   MoveFormToCenter Me
   
   'Modified by Morgan 2014/5/9 考慮多人系統共用問題改放員工編號資料夾
   'm_AttachPath = App.path & "\" & Pub_GetSpecMan("EDocPath")
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      KillTemp
   End If
   
   'WebBrowser1.Navigate "about:blank"
   'SetFileAssociation
   SetCombo1
   
   Me.WindowState = 2 'Added by Morgan 2014/5/15
   Frame3.Width = txtInput.Width 'Added by Morgan 2022/1/18
   
   'Added by Morgan 2025/5/7
   '開放"附件要加印機密的客戶"的智權人員可以下載及更換
   strExc(2) = ""
   If Pub_StrUserSt03 = "M51" Then
      strExc(2) = "Y"
   Else
      strExc(1) = Pub_GetSpecMan("附件要加印機密的客戶")
      strExc(0) = "select * from customer where instr('" & strExc(1) & "',cu01)>0 and cu13='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(2) = "Y"
      End If
   End If
   If strExc(2) = "Y" Then
      CmdOk1(2).Visible = True
      CmdOk1(3).Visible = True
      m_AttachPath3 = Pub_GetSpecMan("寄發文件下載區") & "\" & strUserNum
   End If
   'end 2025/5/7
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If cmdBrowser.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub

Private Sub SetGrid1(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   'Modified by Morgan 2015/8/31 +報價
   'Modified by Morgan 2021/10/20 +實體
   If m_bolBranch = True Then
      arrGridHeadWidth = Array(235, 235, 235, 235, 1035, 840, 1160, 1500, 1500, 860, 1900, 600)
   Else
      arrGridHeadWidth = Array(235, 235, 235, 235, 1035, 840, 1160, 0, 1500, 860, 1900, 600)
   End If
   
   iUbound = UBound(arrGridHeadWidth)

   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2

      iPrevRow1 = 0
      lTotRows1 = 0
      lSelRows1 = 0
      lblCount(1) = lSelRows1 & " / " & lTotRows1
   End If
   .FixedCols = 5
   'Modified by Morgan 2021/10/20 +實體
   .FormatString = "親送|寄送|不寄|實體|本所案號|案件性質|案件名稱|發文室發文時間|專業部發文時間|報價|備註|通知函"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub OpenTable1()
Dim iRow As Integer, idx As Integer, idxCP09 As Integer
Dim stCon As String
Dim idxCaseNo As Integer 'Added by Morgan 2015/6/16
Dim idxColor As Integer  'Added by Morgan 2015/6/24
Dim arrID                'add by sonia 2016/7/1
Dim stUserNo As String, stST06 As String 'Added by Morgan 2022/2/18
Dim idxECert As Integer 'Added by Morgan 2023/2/2
   
   
   If Trim(Left(Combo1.Text, 6)) <> "" Then
      'modify by sonia 2016/7/1 因S29故改寫法
      'stCon = " and cp13='" & Trim(Left("" & Combo1.Text, 6)) & "'"
      arrID = Split(Combo1.Text, " ")
      'Modified by Morgan 2019/10/24 原CP13條件改為LP06(當客戶換業務時方便批次改確認人員)
      'stCon = " and cp13='" & Trim(Left("" & arrID(0), 6)) & "'"
      stUserNo = Trim(Left("" & arrID(0), 6))
      stCon = " and lp06='" & stUserNo & "'"
      'end 2016/7/1
   End If
   
   'Added by Morgan 2022/2/18
   '是否分所改以選單內的人員判斷
   m_bolBranch = False
   If stUserNo <> "" Then
      stST06 = PUB_GetST06(stUserNo)
   Else
      stST06 = pub_strUserOffice
   End If
   If stST06 > "1" And stST06 < "5" Then
      m_bolBranch = True
   End If
   'end 2022/2/18
      
   
   SetGrid1 True
   
   'Modified by Morgan 2014/5/21 分所案件未發文不要顯示
   'Modified by Morgan 2014/8/4 +lp05>0
   'Modified by Morgan 2015/6/16 +CaseNo,ECase
   'Modified by Morgan 2018/10/24 +cp27>19221111(CFP案期限通知報價轉定稿後才會上發文日)
   'Modified by Morgan 2019/3/4 +FDesc
   'Modified by Morgan 2019/10/24 原CP13條件改為LP06(當客戶換業務時方便批次改確認人員)
   'Modify by Amy 2019/11/05 增加Trademark,ServicePractice
'   strExc(0) = "select '' 親送,'' 寄送,'' 不寄" & _
'      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
'      ",decode(pa09,'000',cpm03,cpm04) 案件性質,nvl(pa05,pa06) 案件名稱,rtrim(sqldatet(cp127)||' '||sqltime6(cp128)) 發文室發文時間,rtrim(sqldatet(cp27)||' '||sqltime6(cp82)) 專業部發文時間" & _
'      ",decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價" & _
'      ",'' 備註,decode(cpp02,null,'轉檔中','OK') 客戶函" & _
'      ",'' Read,cpp02,cp09,cp127,cp10,CP01||CP02||CP03||CP04 as CaseNo,LP26 ECase,LP29,LP30,LP35,decode(CP154,'QPGMR','N','Y') Paper,'' FDesc" & _
'      " From letterprogress, caseprogress, SetSpecMan, staff, patent, casepropertymap, (select cpp01,min(cpp02) cpp02" & _
'      " from letterprogress,casepaperpdf where lp05>0 and lp07=0 AND LP10='Y' and lp11 is null and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')>0 And cpp10<>'D' group by cpp01)" & _
'      " where lp05>0 and lp07=0 AND LP10='Y' and lp11 is null and cp09(+)=lp01 and cp27>19221111 and ocode(+)='A7'" & stCon & _
'      " and st01(+)=lp06 and ( instr('234',st06)=0 or lp15='Y' or instr(';'||replace(oMan,',',';')||';',';'||lp06||';')>0) and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10 and cpp01(+)=cp09" & _
'      " order by cp27 desc,cp82 desc,cp09 desc"
    'Modify by Amy 2020/02/03 +and instr(upper(cpp02),'.CUS.PDF')>0
    'Modify by Sindy 2020/2/6 +,Decode(pa01,null,Decode(tm01,null,sp08,tm23),pa26) tm23
    'Moidfy by Amy 2020/03/03 案件名稱原只抓中、英
    'Modified by Morgan 2021/10/20 +LP51
    'Modified by Morgan 2023/2/2 +ECert
    'Modified by Morgan 2025/4/25 電子證書改判斷基本檔(因MCT案都設無實體)
    strExc(0) = "select '' 親送,'' 寄送,'' 不寄" & _
      ",LP51 實體,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",Decode(Decode(pa01,null,Decode(tm01,null,sp09,tm10),pa09), '000',cpm03,cpm04) 案件性質,Decode(pa01,null,Decode(tm01,null,nvl(sp05,nvl(sp06,sp07)),nvl(tm05,nvl(tm06,tm07))),nvl(pa05,nvl(pa06,pa07))) 案件名稱" & _
      ",rtrim(sqldatet(cp127)||' '||sqltime6(cp128)) 發文室發文時間,rtrim(sqldatet(cp27)||' '||sqltime6(cp82)) 專業部發文時間" & _
      ",decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價" & _
      ",'' 備註,decode(cpp02,null,'轉檔中','OK') 客戶函" & _
      ",'' Read,cpp02,cp09,cp127,cp10,CP01||CP02||CP03||CP04 as CaseNo,LP26 ECase,LP29,LP30,LP35,decode(LP51,'','N','Y') Paper,'' FDesc,Decode(pa01,null,Decode(tm01,null,sp08,tm23),pa26) tm23" & _
      ",decode(pa01||pa09||cp10||pa178,'P00016031','Y')||decode(tm01||tm10||cp10||tm136,'T00017011','Y') ECert From letterprogress, caseprogress, SetSpecMan, staff, patent, casepropertymap,TradeMark,ServicePractice," & _
            "(select cpp01,min(cpp02) cpp02" & _
            " from letterprogress,casepaperpdf where lp05>0 and lp07=0 AND LP10='Y' and lp11 is null and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')>0 And cpp10<>'D' group by cpp01)" & _
      " where lp05>0 and lp07=0 AND LP10='Y' and lp11 is null and cp09(+)=lp01 and cp27>19221111 and ocode(+)='A7'" & stCon & _
      " and st01(+)=lp06 and ( instr('234',st06)=0 or lp15='Y' or instr(';'||replace(oMan,',',';')||';',';'||lp06||';g')>0) and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 And tm02(+)=cp02 And tm03(+)=cp03 And tm04(+)=cp04 And sp01(+)=cp01 And sp02(+)=cp02 And sp03(+)=cp03 And sp04(+)=cp04 " & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and cpp01(+)=cp09" & _
      " order by cp27 desc,cp82 desc,cp09 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      SetGrid1
      lTotRows1 = RsTemp.RecordCount
      lblCount(1) = lSelRows1 & " / " & lTotRows1
      .col = 1: .row = 1
      '.Visible = True
      
      'Added by Morgan 2014/7/22 --經理
      idx = GetFieldId("案件性質", MSHFlexGrid1)
      idxCP09 = GetFieldId("CP09", MSHFlexGrid1)
      idxCaseNo = GetFieldId("CaseNo", MSHFlexGrid1) 'Added by Morgan 2015/6/16
      m_idxECase1 = GetFieldId("ECase", MSHFlexGrid1) 'Added by Morgan 2015/6/16
      m_idxPaper = GetFieldId("Paper", MSHFlexGrid1) 'Added by Morgan 2018/10/24
      idxECert = GetFieldId("ECert", MSHFlexGrid1) 'Added by Morgan 2023/2/2
   
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, idxCP09), "1")
         'Added by Morgan 2015/6/16
         'E化客戶案件要變色
         'If PUB_GetEMailFlag(.TextMatrix(iRow, idxCaseNo)) = True Then
         '   .TextMatrix(iRow, m_idxECase1) = "Y"
         If .TextMatrix(iRow, m_idxECase1) <> "" Then
            'E化
            If .TextMatrix(iRow, m_idxECase1) = "Y" Then
               idxColor = 0
            '全E化
            Else
               idxColor = 1
            End If
            .row = iRow
            For intI = 0 To 2
               .col = intI
               .CellBackColor = lblColor(idxColor).BackColor
            Next
            lblColor(idxColor).Visible = True
            lblColorDesc(idxColor).Visible = True
         End If
         'end 2015/6/16
         
         'Added by Morgan 2023/2/2 電子證書變綠色
         If .TextMatrix(iRow, idxECert) = "Y" Then
            idxColor = 2
            .row = iRow
            .col = 3
            .CellBackColor = lblColor(idxColor).BackColor
            lblColor(idxColor).Visible = True
            lblColorDesc(idxColor).Visible = True
         End If
         'end 2023/2/2
      Next
      'end 2014/7/22
      .Visible = True
      End With
      SelectRow 1, MSHFlexGrid1, iPrevRow1
   Else
      'Removed by Morgan 2014/6/4 取消--經理
      'MsgBox "無未發文待確認資料！", vbExclamation
   End If
End Sub

Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With FlexGrid
   nCol = .col
   If pPrevRow > 0 Then
      If pPrevRow <> pRow Then
         .row = pPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
         Next
      End If
   End If

   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   pPrevRow = pRow
   End With
End Sub

Private Sub SetGrid2(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer
   'Modified by Morgan 2015/8/31 +報價
   'Modified by Lydia 2016/08/22 +CP64
   arrGridHeadWidth = Array(225, 1035, 1200, 1200, 1500, 840, 860, 1900, 600, 0)
   iUbound = UBound(arrGridHeadWidth)

   With MSHFlexGrid2
   If pReset = True Then
      .Clear
      .Rows = 2

      iPrevRow2 = 0
      lTotRows2 = 0
      lSelRows2 = 0
      lblCount(2) = lSelRows2 & " / " & lTotRows2
   End If
   .FixedCols = 2
   'Modified by Lydia 2016/08/22 + CP64
   .FormatString = "V|本所案號|案件性質|案件名稱|發文室發文時間|本所期限|報價|備註|通知函|CP64"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub OpenTable2()
Dim iRow As Integer, idx As Integer, idxCP09 As Integer
Dim stCon As String
Dim idxCaseNo As Integer 'Added by Morgan 2015/6/16
Dim idxColor As Integer 'Added by Morgan 2015/6/24
Dim arrID                'add by sonia 2016/7/12
   
   If Trim(Left(Combo1.Text, 6)) <> "" Then
      'modify by sonia 2016/7/12 因S29故改寫法
      'stCon = " and cp13='" & Trim(Left("" & Combo1.Text, 6)) & "'"
      arrID = Split(Combo1.Text, " ")
      'Modified by Morgan 2019/10/24 原CP13條件改為LP06(當客戶換業務時方便批次改確認人員)
      'stCon = " and cp13='" & Trim(Left("" & arrID(0), 6)) & "'"
      stCon = " and lp06='" & Trim(Left("" & arrID(0), 6)) & "'"
      'end 2016/7/12
   End If
   
   SetGrid2 True
   'Modified by Morgan 2014/7/28 +lp03>0
   'Modified by Morgan 2015/6/16 +CaseNo,ECase
   'Modified by Lydia 2016/08/22 + CP64
   'Modify by Amy 2019/11/05 +Trademark,ServicePractice
   'Modify by Sindy 2020/2/6 +,Decode(pa01,null,Decode(tm01,null,sp08,tm23),pa26) tm23
   'Modify by Amy 2020/03/03 案件名稱原只抓中、英
   'Modified by Morgan 2021/11/18 全E化不判斷發文室發文(承辦人跑歷程的來函輸入時還沒有定稿不會上系統發文)
   'Modified by Morgan 2022/2/14 +cp27>19221111(CFP案期限通知報價轉定稿後才會上發文日)
   'Modified by Morgan 2022/2/16 +lp05>0 全E化可能會先有發文室發文日但尚未判發
   strExc(0) = "select '' V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",Decode(Decode(pa01,null,Decode(tm01,null,sp09,tm10),pa09), '000',cpm03,cpm04) 案件性質,Decode(pa01,null,Decode(tm01,null,nvl(sp05,nvl(sp06,sp07)),nvl(tm05,nvl(tm06,tm07))),nvl(pa05,nvl(pa06,pa07))) 案件名稱" & _
      ",rtrim(sqldatet(cp127)||' '||sqltime6(cp128)) 發文室發文時間" & _
      ",sqldatet(cp06) 本所期限,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價" & _
      ",'' 備註,decode(cpp02,null,'轉檔中','OK') 客戶函,CP64" & _
      ",'' Read,cpp02,cp09,cp10,CP01||CP02||CP03||CP04 as CaseNo,LP26 ECase,LP29,LP30,LP35,'' FDesc,Decode(pa01,null,Decode(tm01,null,sp08,tm23),pa26) tm23" & _
      " From letterprogress,caseprogress, patent, casepropertymap,TradeMark,ServicePractice," & _
            "(select cpp01,min(cpp02) cpp02" & _
            " from letterprogress,casepaperpdf where  lp03>0 and lp07=0 AND LP10='Y' and lp11='Y' AND (LP15='Y' or LP26='E') " & _
            " and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')>0 group by cpp01)" & _
      " Where lp03>0 and lp05>0 and lp07=0 AND LP10='Y' and lp11='Y' AND LP15='Y' and cp09(+)=lp01 and cp27>19221111" & stCon & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 And tm02(+)=cp02 And tm03(+)=cp03 And tm04(+)=cp04 And sp01(+)=cp01 And sp02(+)=cp02 And sp03(+)=cp03 And sp04(+)=cp04 " & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and cpp01(+)=cp09" & _
      "" & _
      " order by cp127 desc,cp128 desc,cp09 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid2
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      lTotRows2 = RsTemp.RecordCount
      SetGrid2
      .col = 1: .row = 1
      SelectRow 1, MSHFlexGrid2, iPrevRow2
      lblCount(2) = lSelRows2 & " / " & lTotRows2
      '.Visible = True
      
      'Added by Morgan 2014/7/22 --經理
      idx = GetFieldId("案件性質", MSHFlexGrid2)
      idxCP09 = GetFieldId("CP09", MSHFlexGrid2)
      idxCaseNo = GetFieldId("CaseNo", MSHFlexGrid2) 'Added by Morgan 2015/6/16
      m_idxECase2 = GetFieldId("ECase", MSHFlexGrid2) 'Added by Morgan 2015/6/16
      
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, idxCP09), "1")
         'Added by Morgan 2015/6/16
         'E化客戶案件要變色
         'If PUB_GetEMailFlag(.TextMatrix(iRow, idxCaseNo)) = True Then
         '   .TextMatrix(iRow, m_idxECase2) = "Y"
         If .TextMatrix(iRow, m_idxECase2) <> "" Then
            If .TextMatrix(iRow, m_idxECase2) = "Y" Then
               idxColor = 0
            Else
               idxColor = 1
            End If
            .row = iRow
            .col = 0
            .CellBackColor = lblColor(idxColor).BackColor
            lblColor(idxColor).Visible = True
            lblColorDesc(idxColor).Visible = True
         End If
         'end 2015/6/16
      Next
      'end 2014/7/22
      .Visible = True
      End With
      
   Else
      'Removed by Morgan 2014/6/4 取消--經理
      'MsgBox "無直寄待確認資料！", vbExclamation
   End If
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

Private Function GetValue(pRow As Integer, pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As String
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String, ByRef FlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Function FormSave1() As Boolean
   Dim iRow As Integer, idxCP09 As Integer, iChoice As String, idxMemo As Integer
   'Add by Sindy 2020/2/6
   Dim idxCaseNo As Integer, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   Dim salesNo As String, salesArea As String
   '2020/2/6 END
   Dim strCaseNo As String, bolMCTF As Boolean 'Add By Sindy 2021/4/28
   Dim stCheck As String 'Added by Morgan 2022/7/4
   
On Error GoTo ErrHnd
   
   idxCP09 = GetFieldId("cp09", MSHFlexGrid1)
   idxMemo = GetFieldId("備註", MSHFlexGrid1)
   'Add by Sindy 2020/2/6
   idxCaseNo = GetFieldId("本所案號", MSHFlexGrid1)
   'idxtm23 = GetFieldId("tm23", MSHFlexGrid1)
   '2020/2/6 END
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      stCheck = .TextMatrix(iRow, 0) & .TextMatrix(iRow, 1) & .TextMatrix(iRow, 2)
      'Modified by Morgan 2022/7/4 +只E不寄紙本
      'If stCheck = "V" Then
      If stCheck = "V" Or stCheck = "E" Then
      'end 2022/7/4
         If .TextMatrix(iRow, 0) = "V" Then
            iChoice = 0
         ElseIf .TextMatrix(iRow, 1) = "V" Then
            iChoice = 1
         'Added by Morgan 2022/7/4
         ElseIf .TextMatrix(iRow, 1) = "E" Then
            iChoice = 1
         'end 2022/7/4
         ElseIf .TextMatrix(iRow, 2) = "V" Then
            iChoice = 2
         End If
         
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT

         'Add By Sindy 2021/4/28
         strCaseNo = Trim(.TextMatrix(iRow, idxCaseNo)) & "-0-00"
         strCP01 = SystemNumber(strCaseNo, 1)
         strCP02 = SystemNumber(strCaseNo, 2)
         strCP03 = SystemNumber(strCaseNo, 3)
         strCP04 = SystemNumber(strCaseNo, 4)
         bolMCTF = False
         If Mid(PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04), 1, 4) = "MCTF" Then
            bolMCTF = True
         End If
         '2021/4/28 END
         
         'Modify By Sindy 2021/4/6 + lp32=null
         '大至台-整批延展通知，調整程序可以做寄送，發文室可以做確認
         If iChoice = "1" And Trim(Left(Combo1.Text, 6)) = "P2002" Then
            strSql = "update letterprogress set lp32=null,lp06='" & strUserNum & "'" & _
                     ",lp07=" & strSrvDate(1) & ",lp11='" & iChoice & "'" & _
                     ",lp12='" & ChgSQL(.TextMatrix(iRow, idxMemo)) & "'" & _
                     " where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
            cnnConnection.Execute strSql, intI
         Else
         '2021/4/6 END
            'Modify By Sindy 2021/4/28 + MCT案件親送,不經發文室
            strSql = "update letterprogress set lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & _
                     ",lp11='" & iChoice & "',lp12='" & IIf(bolMCTF = True And iChoice = "0", "不經發文室;", "") & ChgSQL(.TextMatrix(iRow, idxMemo)) & "'" & _
                     " where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
            cnnConnection.Execute strSql, intI
         End If

         'Added by Morgan 2018/9/11
         '若E化要寄送時，進度檔發文人員日期時間要清除發文室才能作業 Ex.P-117994 補充說明 (107/9/10發文)
         If .TextMatrix(iRow, m_idxECase1) = "Y" Then
            '2不寄
            'Modify By Sindy 2021/4/28 + MCT案件親送,不經發文室
            If iChoice = "2" Or _
               (bolMCTF = True And iChoice = "0") Then
            '2021/4/28 END
               'Added by Morgan 2020/3/19 E化不寄也要上發文室發文以便寄件查詢能看到並可重EMail
               strSql = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp127 is null"
               cnnConnection.Execute strSql, intI
               If intI > 0 Then
                  strSql = "update caseprogress set cp154='QPGMR' where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
                  cnnConnection.Execute strSql, intI
               End If
               'end 2020/3/19
               
            '0親送/1寄送
            'Modified by Morgan 2018/9/18 先取消,分所案件為北所發文給分所的日期不可清除
            'Modified by Morgan 2019/10/30 +判斷北所人員操作時清除
            ElseIf pub_strUserOffice = "1" Then
               strSql = "update caseprogress set cp154=NULL,cp127=NULL,cp128=NULL where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp154='QPGMR' and cp127>0"
               cnnConnection.Execute strSql, intI
            End If
            
         'Added by Morgan 2022/3/21
         '全E化有實體
         ElseIf .TextMatrix(iRow, m_idxECase1) = "E" And .TextMatrix(iRow, m_idxPaper) = "Y" Then
            '不寄也要上發文日(寄件查詢要能看到)
            'Modified by Morgan 2022/7/4 +只E不寄紙本
            'If iChoice = "2" Then
            'Modified by Morgan 2025/2/26 親送也不須經發文室
            If iChoice = "2" Or iChoice = "0" Or .TextMatrix(iRow, 1) = "E" Then
            'end 2022/7/4
               strSql = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp127 is null"
               cnnConnection.Execute strSql, intI
               If intI > 0 Then
                  strSql = "update caseprogress set cp154='QPGMR' where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
                  cnnConnection.Execute strSql, intI
               End If
               
               'Added by Morgan 2022/7/4
               If .TextMatrix(iRow, 1) = "E" Then
                  strSql = "update letterprogress set lp12='只E不寄紙本;'||lp12 where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
                  cnnConnection.Execute strSql, intI
               End If
               'end 2022/7/4
               
            '清除QPGMR的發文紀錄
            Else
               strSql = "update caseprogress set cp154=NULL,cp127=NULL,cp128=NULL where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp154='QPGMR' and cp127>0"
               cnnConnection.Execute strSql, intI
            End If
         'end 2022/3/21
         End If
         'end 2019/10/30
         'end 2018/9/11

         
         cnnConnection.CommitTrans
         
On Error GoTo ErrHnd
         If iRow = iPrevRow1 Then SelectRow 0, MSHFlexGrid1, iPrevRow1
         .TextMatrix(iRow, 0) = "X"
         .RowHeight(iRow) = 0
         iPrevRow1 = 0 'Added by Morgan 2018/9/19
         lSelRows1 = lSelRows1 - 1
         lTotRows1 = lTotRows1 - 1
         lblCount(1) = lSelRows1 & " / " & lTotRows1
         DoEvents
      End If
   Next
   End With
   
   FormSave1 = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function FormSave2() As Boolean
   Dim iRow As Integer, idxCP09 As Integer, idxMemo As Integer
   Dim idxCP64 As Integer, idxPrice As Integer, idxCaseNo As Integer 'Added by Lydia 2016/08/22
      
On Error GoTo ErrHnd
   
   idxCP09 = GetFieldId("cp09", MSHFlexGrid2)
   idxMemo = GetFieldId("備註", MSHFlexGrid2)
    'Added  by Lydia 2016/08/22
   idxCP64 = GetFieldId("CP64", MSHFlexGrid2)
   idxCaseNo = GetFieldId("本所案號", MSHFlexGrid2)
   idxPrice = GetFieldId("報價", MSHFlexGrid2)
   
   With MSHFlexGrid2
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT
         strSql = "update letterprogress set lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & ",lp12='" & ChgSQL(.TextMatrix(iRow, idxMemo)) & "' where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
         cnnConnection.Execute strSql, intI
         
         'Added by Lydia 2016/08/22 直寄凡有報價確認,更新案件進度備註
         strExc(1) = "" & MSHFlexGrid2.TextMatrix(iRow, idxCaseNo)
         strExc(2) = "" & MSHFlexGrid2.TextMatrix(iRow, idxPrice)
         strExc(3) = "" & MSHFlexGrid2.TextMatrix(iRow, idxCP64)
         If strExc(2) <> "" Then
            strExc(2) = Mid(strExc(2), 1, InStr(strExc(2), ")") - 1) & "P);"
            strSql = "update caseprogress set cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " " & 業務報價確認備註 & strExc(2) & "'||CP64 " & _
                     "where cp09='" & MSHFlexGrid2.TextMatrix(iRow, idxCP09) & "' "
            'Added by Morgan 2019/5/20 排除有報價確認記錄者(轉定稿時已經有更新備註) Ex:CFP-029943
            strSql = strSql & " and not exists(select * from lettercache where lc01='" & MSHFlexGrid2.TextMatrix(iRow, idxCP09) & "')"
            'end 2019/5/20
            cnnConnection.Execute strSql, intI
         End If
         'end 2016/08/22
         
         cnnConnection.CommitTrans
         
On Error GoTo ErrHnd
         If iRow = iPrevRow2 Then SelectRow 0, MSHFlexGrid2, iPrevRow2
         .TextMatrix(iRow, 0) = "X"
         iPrevRow2 = 0 'Added by Morgan 2018/9/19
         .RowHeight(iRow) = 0
         lSelRows2 = lSelRows2 - 1
         lTotRows2 = lTotRows2 - 1
         lblCount(2) = lSelRows2 & " / " & lTotRows2
         DoEvents
      End If
   Next
   End With
   
   FormSave2 = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm210144 = Nothing
End Sub

Private Sub KillTemp()
On Error GoTo ErrHnd
   WebBrowser1.Navigate "about:blank": DoEvents '要先釋放預覽文件,否則會無法刪除
   
   'Modified by Morgan 2020/2/18
   'If Dir(m_AttachPath & "\.") <> "" Then
   '   Kill m_AttachPath & "\*.*"
   'End If
   PUB_ClearTempFolder m_AttachPath
   'end 2020/2/18
   Exit Sub
   
ErrHnd:
   Resume Next
End Sub

Private Sub SetCombo1()
   Dim ii As Integer
   Dim stDef As String 'Add by Amy 2019/11/05 預設
   Dim strQ As String, strMCTF As String 'Add by Amy 2020/02/19
   
   'Modify By Sindy 2022/5/25 設定屬智權人員作業的下拉選單(共用模組)
   'Modify by Amy 2023/02/10 +Me.Name
   Call PUB_SetCombo1Sales(Combo1, , Me.Name)
   
'   Combo1.Clear
'   Combo1.AddItem strUserNum & " " & strUserName
   
   'Add by Amy 2019/11/05 內商程序人員操作時,增加 P2002 內商程序,並預設在此欄位值
   'Modify by Amy 2020/02/19 +MCTF
   'Modify by Sindy 2021/11/24 + Or Pub_StrUserSt03 = "F11"
   If Left(Pub_StrUserSt03, 2) = "P2" Or Pub_StrUserSt03 = "F11" Then
        strMCTF = GetMCTF0XAllCode(strUserNum)
        If strSrvDate(1) >= T商標電子化第2階段啟用日 And strMCTF <> MsgText(601) Then
             strQ = "And st01 in ('" & strMCTF & "') "
        ElseIf Pub_StrUserSt03 = "P22" Then
             strQ = "And st01='P2002' "
        End If
        If strQ <> MsgText(601) Then
             strQ = "Select st01,st02 From Staff Where 1=1 " & strQ & " Order by st01 "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strQ)
             If intI = 1 Then
                 RsTemp.MoveFirst
                 Do While Not RsTemp.EOF
                     For ii = 0 To Combo1.ListCount - 1
                        If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
                           'Add By Sindy 2020/11/20 抓第一筆為預設值
                           If stDef = "" Then
                           '2020/11/20 END
                              stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                           End If
                           Exit For
                        End If
                     Next
                     If ii = Combo1.ListCount Then
                        Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                        'Add By Sindy 2020/11/20 抓第一筆為預設值
                        If stDef = "" Then
                        '2020/11/20 END
                           stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                        End If
                     End If
                     RsTemp.MoveNext
                 Loop
             End If
        End If
   End If
   'end 2020/02/19
   
'   '檢查當時是否需要為他人職代
'   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
'

'   'Modified by Lydia 2020/06/08 +增加特殊權限"AREA"
'   Call Pub_SetSAManageEmpCombo(strUserNum, Combo1, False, , , "AREA")
''end 2014/9/4
'
'   'Added by Morgan 2014/5/15
'   '專利處智權同仁代處理人
'   'Modify by Amy 2015/03/13 +特殊設定(總經理業務工作代理人員)
'   If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Or InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 Then
'      If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      Else
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='總經理員工編號' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      End If
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            For ii = 0 To Combo1.ListCount - 1
'               If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
'                  Exit For
'               End If
'            Next
'            If ii = Combo1.ListCount Then
'               Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'            End If
'
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   'end 2014/5/15
'   '帶人主管抓虛建編號
'   strSql = "select st01,st02 from staff where st01<'63001' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & strUserNum & ";')>0"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         For ii = 0 To Combo1.ListCount - 1
'            If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
'               Exit For
'            End If
'         Next
'         If ii = Combo1.ListCount Then
'            Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'         End If
'
'         RsTemp.MoveNext
'      Loop
'   End If
   
   If Pub_StrUserSt03 = "M51" Then
      Combo1.AddItem "      " & "全部"
   End If
   Combo1.ListIndex = 0
   'Modify By Sindy 2020/11/20 有預設值就帶
   'If Pub_StrUserSt03 = "P22" Then Combo1 = stDef 'Add by Amy 2019/11/05 內商程序人預設 P2002
   If stDef <> "" Then Combo1 = stDef 'Add by Amy 2019/11/05 內商程序人預設 P2002
'cancel by sonia 2024/9/27
'   'Add By Sindy 2023/5/16
'   If InStr(Pub_GetSpecMan("P1004管理人員"), strUserNum) > 0 Then
'      Combo1 = "P1004 " & GetPrjSalesNM("P1004")
'   End If
'   '2023/5/16 END
'end 2024/9/27
End Sub

Private Sub RePosForm(pFull As Boolean)
   Static lngLeft As Long
   
   If Forms(0).WindowState <> 1 Then
      'Modified by Mogrgan 2019/3/4 加pdf檔下拉選單
      If lngLeft = 0 Then lngLeft = WebBrowser1.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         cmdBrowser.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         cmdBrowser.Caption = "點我展開"
      End If
      WebBrowser1.Width = Me.Width - WebBrowser1.Left - 90
      WebBrowser1.Height = Me.Height - Frame4.Height - 390
      Frame4.Left = WebBrowser1.Left
      Frame4.Width = WebBrowser1.Width
      cmdBrowser.Width = Frame4.Width
      cboAtt.Width = Frame4.Width - cboAtt.Left
      'end 2019/3/4
      
      MSHFlexGrid1.Height = (Me.Height) * 0.52 - MSHFlexGrid1.Top - Frame1.Height
      Frame1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height - 120
      
      MSHFlexGrid2.Top = Frame1.Top + Frame1.Height - 60
      MSHFlexGrid2.Height = Me.Height - MSHFlexGrid2.Top - Frame2.Height - 280
      Frame2.Top = MSHFlexGrid2.Top + MSHFlexGrid2.Height - 120
   End If
End Sub
'Added by Morgan 2019/3/4
Private Sub lblAttCnt_Click()
   SendMessage cboAtt.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub


Private Sub MSHFlexGrid1_DblClick()
   If iPrevRow1 > 0 Then
      fnOpen1 1
      'Added by Morgan 2017/12/27
      'Removed by Morgan 2018/10/1 看進度備註,LP35先保留
      'strExc(1) = GetValue(iPrevRow1, "LP35", MSHFlexGrid1)
      'If strExc(1) <> "" Then
      '   MsgBox strExc(1), , "報價備註"
      'End If
      'end 2018/10/1
      'end 2017/12/27
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim nCol As Integer, nRow As Integer, iRow As Integer, iCol As Integer
   Dim stValue As String
   Dim stCP09 As String
      
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      '紀錄前次點選的收文號
      If iPrevRow1 > 0 Then
         stCP09 = GetValue(iPrevRow1, "cp09", MSHFlexGrid1)
      End If
      
      .col = nCol
      If m_blnColOrderAsc1 = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc1 = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc1 = False
      End If
               
      '重設排序後前次點選的位置
      If iPrevRow1 > 0 Then
         For iRow = 1 To .Rows - 1
            If stCP09 = GetValue(iRow, "cp09", MSHFlexGrid1) Then
               iPrevRow1 = iRow
               Exit For
            End If
         Next
      End If
   'Modified by Morgan 2018/9/19 已確認的也不可再點選
   'ElseIf nRow > 0 And .TextMatrix(nRow, 3) <> "" Then
   ElseIf nRow > 0 And .TextMatrix(nRow, 4) <> "" And .TextMatrix(nRow, 0) <> "X" Then
   'end 2018/9/19
      SelectRow nRow, MSHFlexGrid1, iPrevRow1
      
      .row = nRow
      .col = nCol
      If nCol < 3 Then
         ClickGrid MSHFlexGrid1, 1
      End If
      
      '有確認的點選備註欄可輸入
      If .TextMatrix(.row, 0) & .TextMatrix(.row, 1) & .TextMatrix(.row, 2) = "V" Then
         SetBox MSHFlexGrid1
      End If
   End If
   .Visible = True
   End With
End Sub


Private Sub SetBox(ByRef FlexGrid As MSHFlexGrid)
   Dim lngLeft As Long, lngTop As Long, iCol As Integer, ii As Integer

   iCol = GetFieldId("備註", FlexGrid)
   With FlexGrid
      If .col = iCol Then
         txtInput.FontName = .CellFontName
         txtInput.FontSize = .CellFontSize
         'txtInput.Alignment = .CellAlignment \ 5
         txtInput.Text = .TextMatrix(.row, .col)
         txtInput.Tag = txtInput.Text
         txtInput.Width = .ColWidth(.col)
         txtInput.Height = .RowHeight(.row)
         txtInput.Tag = txtInput.Text
         'Modified by Morgan 2022/1/18 2.0的TextBox無法顯示在上層,改放在Frame內控制
         'txtInput.Visible = True
         Frame3.Visible = True
         'end 2022/1/18
         txtInput.SetFocus
         
         lngLeft = .Left + 25
         lngTop = .Top + .RowHeight(0) + 25
         
         lngLeft = lngLeft + .ColPos(.col)
         
         For ii = .TopRow To .row - 1
            lngTop = lngTop + .RowHeight(ii)
         Next
         'Modified by Morgan 2022/1/18 2.0的TextBox無法顯示在上層,改放在Frame內控制
         'txtInput.Left = lngLeft: txtInput.Top = lngTop
         Frame3.Width = txtInput.Width - 10
         Frame3.Height = txtInput.Height - 10
         Frame3.Left = lngLeft: Frame3.Top = lngTop
         'end 2022/1/18
         m_InputRow = .row
         m_InputCol = .col
         Set m_InputGrid = FlexGrid
      End If
   End With
   
End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MSHFlexGrid1.ToolTipText = ""
If MSHFlexGrid1.MouseRow <> 0 And MSHFlexGrid1.MouseCol > 0 Then
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol) <> "" Then
      MSHFlexGrid1.ToolTipText = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol)
   End If
End If
End Sub

Private Sub MSHFlexGrid1_Scroll()
   'Modified by Morgan 2022/1/18 2.0的TextBox無法顯示在上層,改放在Frame內控制
   'If txtInput.Visible = True Then
   '   m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
   '   txtInput.Visible = False
   'End If
   If Frame3.Visible = True Then
      m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
      Frame3.Visible = False
   End If
   'end 2022/1/18
End Sub

Private Sub MSHFlexGrid2_DblClick()
   If iPrevRow2 > 0 Then
      fnOpen2 1
      'Added by Morgan 2017/12/27
      'Removed by Morgan 2018/10/9 看進度備註,LP35先保留
      'strExc(1) = GetValue(iPrevRow2, "LP35", MSHFlexGrid2)
      'If strExc(1) <> "" Then
      '   MsgBox strExc(1), , "報價備註"
      'End If
      'end 2018/10/9
      'end 2017/12/27
   End If
End Sub

Private Sub MSHFlexGrid2_Click()
  Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stCP09 As String
      
   With MSHFlexGrid2
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      If nCol = 0 Then
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) = "" Then
               stValue = "V"
               Exit For
            '已刪除資料標示為 X
            ElseIf .TextMatrix(iRow, 0) = "V" Then
               stValue = ""
               Exit For
            End If
         Next
         
         lSelRows2 = 0
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) <> "X" Then
               If .TextMatrix(iRow, 0) <> stValue Then
                  .TextMatrix(iRow, 0) = stValue
               End If
            End If
            If .TextMatrix(iRow, 0) = "V" Then
               lSelRows2 = lSelRows2 + 1
            End If
         Next
         
         lblCount(2) = lSelRows2 & " / " & lTotRows2
      Else
         
         '紀錄前次點選的收文號
         If iPrevRow2 > 0 Then
            stCP09 = GetValue(iPrevRow2, "cp09", MSHFlexGrid2)
         End If
         
         .col = nCol
         If m_blnColOrderAsc2 = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc2 = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc2 = False
         End If
                  
         '重設排序後前次點選的位置
         If iPrevRow2 > 0 Then
            For iRow = 1 To .Rows - 1
               If stCP09 = GetValue(iRow, "cp09", MSHFlexGrid2) Then
                  iPrevRow2 = iRow
                  Exit For
               End If
            Next
         End If
      End If
   'Modified by Morgan 2018/9/19 已確認的也不可再點選
   'ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" And .TextMatrix(nRow, 0) <> "X" Then
   'end 2018/9/19
      SelectRow nRow, MSHFlexGrid2, iPrevRow2
      
      .col = nCol
      .row = nRow
      If nCol = 0 Then
         ClickGrid MSHFlexGrid2, 2
      End If
      '有確認的點選備註欄可輸入
      If .TextMatrix(.row, 0) = "V" Then
         SetBox MSHFlexGrid2
      End If
   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(ByRef FlexGrid As MSHFlexGrid, Index As Integer)
   Dim iCol As Integer

   With FlexGrid
   If .Text = "V" Or .Text = "E" Then
      If Index = 1 Then
         lSelRows1 = lSelRows1 - 1
      Else
         lSelRows2 = lSelRows2 - 1
      End If
      .Text = ""
      
   '已刪除資料標示為 X
   ElseIf .Text = "" Then
      iCol = GetFieldId("Read", FlexGrid)
      'Modified by Morgan 2014/5/2 未發文待確認改不必開啟也可確認
      'If .TextMatrix(.row, iCol) = "" Then
      If Index = 2 And .TextMatrix(.row, iCol) = "" Then
         .Visible = True
         MsgBox "請開啟客戶函後再行確認!!", vbExclamation
         Exit Sub
      
      '分所非直寄的不可於發文室發文日當日確認
      'Modified by Morgan 2015/7/9
      ElseIf Index = 1 And m_bolBranch = True Then
         'Modified by Morgan 2019/10/1 修正,應該是非E化才要檢查
         If .TextMatrix(.row, m_idxECase1) = "" Then
            iCol = GetFieldId("cp127", FlexGrid)
            If Trim(.TextMatrix(.row, iCol)) = "" Then
               .Visible = True
               MsgBox "分所案件，北所發文室尚未發文，不可確認!!", vbExclamation
               Exit Sub
            ElseIf .TextMatrix(.row, iCol) = strSrvDate(1) Then
               .Visible = True
               MsgBox "分所案件，紙本應尚未送達不可確認!!", vbExclamation
               Exit Sub
            End If
         End If
      End If
      '非直寄
      If Index = 1 Then
         'Added by Morgan 2015/6/16
         If .TextMatrix(.row, m_idxECase1) <> "" Then
            iCol = GetFieldId("Read", FlexGrid)
            If .TextMatrix(.row, iCol) = "" Then
               .Visible = True
               MsgBox "請開啟客戶函後再行確認!!", vbExclamation, "E化客戶案件提醒"
               Exit Sub
            End If
         
            If .col <> 2 Then
               .Visible = True
               'Added by Morgan 2018/10/24
               If .TextMatrix(.row, m_idxECase1) = "E" Then
                  'e化客戶非實體不可勾選親送，而寄送表示確認要 Email 給客戶(由程序作業)
                  If .col = 0 And .TextMatrix(.row, m_idxPaper) = "N" Then
                     MsgBox "有實體文件(例如收據、證書…)才可勾選 ""親送"" ！", vbExclamation, "全E化客戶案件提醒"
                     Exit Sub
                  End If
               
               'end 2018/10/24
               'Modified by Morgan 2022/4/13 無實體才應勾選不寄
               'Else
               ElseIf .TextMatrix(.row, m_idxPaper) = "N" Then
               'end 2022/4/13
                  If MsgBox(.TextMatrix(.row, 4) & "為E化案件應勾選 ""不寄""，是否仍要勾選 """ & .TextMatrix(0, .col) & """？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
            'Added by Morgan 2018/10/24
            '全E化客戶非實體勾選不寄表示程序不要 Email 給客戶
            ElseIf .TextMatrix(.row, m_idxECase1) = "E" Then
               .Visible = True
               If MsgBox("勾選 ""不寄"" 則程序人員將不會 Email 給客戶！" & vbCrLf & vbCrLf & "是否確定不寄？", vbYesNo + vbQuestion + vbDefaultButton2, "全E化客戶案件提醒") = vbNo Then
                  Exit Sub
               End If
            'end 2018/10/24
            End If
         End If
         'end 2015/6/16
         
         If .TextMatrix(.row, 0) = "V" Then
            .TextMatrix(.row, 0) = ""
         ElseIf .TextMatrix(.row, 1) = "V" Then
            .TextMatrix(.row, 1) = ""
         'Added by Morgan 2022/7/4 +只E不寄紙本
         ElseIf .TextMatrix(.row, 1) = "E" Then
            .TextMatrix(.row, 1) = ""
         'end 2022/7/4
         ElseIf .TextMatrix(.row, 2) = "V" Then
            .TextMatrix(.row, 2) = ""
         Else
            lSelRows1 = lSelRows1 + 1
         End If
      '直寄
      Else
         lSelRows2 = lSelRows2 + 1
      End If
      .Text = "V"
   End If
   
   '非直寄
   If Index = 1 Then
      lblCount(1) = lSelRows1 & " / " & lTotRows1
   
      'Added by Morgan 2022/7/4
      '全E化特定客戶詢問是否寄送紙本信函(收據)
      'Removed by Morgan 2025/3/5 取消，已改用欄位設定
      'If .TextMatrix(.row, m_idxECase1) = "E" And .TextMatrix(.row, 1) = "V" And .TextMatrix(.row, m_idxPaper) = "Y" Then
      '   If GetValue(.row, "cp09", FlexGrid) < "C" Or GetValue(.row, "cp10", FlexGrid) = "1101" Then
      '      '預設不寄紙本: X41570 訊凱關係企業
      '      If Left(GetValue(.row, "tm23", FlexGrid), 6) = "X41570" Then
      '         If MsgBox("您已勾選 [寄送]，程序將會EMail給客戶。" & vbCrLf & vbCrLf & "此客戶有特殊要求，請再確認是否寄送 [紙本] 信函及收據？", vbYesNo + vbQuestion + vbDefaultButton2, "全E化客戶案件提醒") = vbNo Then
      '            .TextMatrix(.row, 1) = "E"
      '         End If
      '      End If
      '   End If
      'End If
      'end 2022/7/4
   '直寄
   Else
      lblCount(2) = lSelRows2 & " / " & lTotRows2
   End If
   
   End With
End Sub

Private Sub MSHFlexGrid2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MSHFlexGrid2.ToolTipText = ""
If MSHFlexGrid2.MouseRow <> 0 And MSHFlexGrid2.MouseCol > 0 Then
   If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.MouseRow, MSHFlexGrid2.MouseCol) <> "" Then
      MSHFlexGrid2.ToolTipText = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.MouseRow, MSHFlexGrid2.MouseCol)
   End If
End If
End Sub

Private Sub txtInput_Change()
   txtInput = PUB_StrToStr(txtInput, txtInput.MaxLength)
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
   OpenIme
End Sub

Private Sub txtInput_KeyPress(KeyAscii As MSForms.ReturnInteger)
   Dim iCol As Integer, iRow As Integer
   
   If KeyAscii = vbKeyReturn Then
      m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
      'Modified by Morgan 2022/1/18 2.0的TextBox無法顯示在上層,改放在Frame內控制
      'txtInput.Visible = False
      Frame3.Visible = False
      'end 2022/1/18
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
   End If
End Sub

Private Sub txtInput_LostFocus()
   'Modified by Morgan 2022/1/18 2.0的TextBox無法顯示在上層,改放在Frame內控制
   'If txtInput.Visible = True Then
   '   m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
   '   txtInput.Visible = False
   'End If
   If Frame3.Visible = True Then
      m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
      Frame3.Visible = False
   End If
   'end 2022/1/18
End Sub


Public Sub PubShowNextData()
   Dim StrTag As String
   Dim bolCancel As Boolean
   
   Me.Enabled = False
   Select Case cmdState
      Case 0, 1
         If (cmdState = 0 And iPrevRow1 > 0) Or (cmdState = 1 And iPrevRow2 > 0) Then
            If fnSaveParentForm(Me) = False Then
               Me.Enabled = True
               Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            If cmdState = 0 Then
               StrTag = Pub_RplStr(MSHFlexGrid1.TextMatrix(iPrevRow1, 4))
            Else
               StrTag = Pub_RplStr(MSHFlexGrid2.TextMatrix(iPrevRow2, 1))
            End If
            If UBound(Split(StrTag, "-")) = 1 Then
               StrTag = StrTag & "-0-00"
            End If
            frm100101_2.Show
            frm100101_2.Tag = StrTag
            frm100101_2.cmdOK(6).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
         End If
         
      
      Case 2 '下載 Added by Morgan 2025/5/6
         fnDownPdf
      Case 3 '更換 Added by Morgan 2025/5/6
         fnUpdatePdf
      Case 4, 5 '卷宗區
         
         'Added by Morgan 2021/7/23
         If PUB_CheckFormExist("frm100101_L") Then
            MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
            Me.Enabled = True
            Exit Sub
            'frm100101_L.SetParent Nothing
            'frm100101_L.cmdExit.Value = True
            'Me.ZOrder
         End If
         'end 2021/7/23
         
         If (cmdState = 4 And iPrevRow1 > 0) Or (cmdState = 5 And iPrevRow2 > 0) Then
            'Added by Morgan 2021/7/23
            If fnSaveParentForm(Me, True) = False Then
               Me.Enabled = True
               Exit Sub
            End If
            'end 2021/7/23
                  
            Screen.MousePointer = vbHourglass
            If cmdState = 4 Then
               StrTag = Pub_RplStr(MSHFlexGrid1.TextMatrix(iPrevRow1, 4))
            Else
               StrTag = Pub_RplStr(MSHFlexGrid2.TextMatrix(iPrevRow2, 1))
            End If
            If UBound(Split(StrTag, "-")) = 1 Then
               StrTag = StrTag & "-0-00"
            End If
            frm100101_L.m_strKey = StrTag
            'frm100101_L.Hide
            frm100101_L.SetParent Me
            If frm100101_L.QueryData = True Then
               frm100101_L.Show
               Me.Hide
            Else
               Unload frm100101_L
            End If
            Screen.MousePointer = vbDefault
         End If
   End Select
   Me.Enabled = True
End Sub
'Added by Morgan 2019/3/4
Private Sub SetAttList(Optional pItems As String)
   Dim arrItem() As String
   Dim ii As Integer, iAttCnt As Integer
   
   cboAtt.Clear: lblAttCnt = " PDF:(0)"
   If pItems <> "" Then
      arrItem = Split(pItems, ";")
      For ii = LBound(arrItem) To UBound(arrItem)
         If arrItem(ii) <> "" Then
            cboAtt.AddItem arrItem(ii)
            iAttCnt = iAttCnt + 1
         End If
      Next
      lblAttCnt = " PDF:(" & iAttCnt & ")"
   End If
End Sub

'Added by Morgan 2021/9/30 檢查是否有EMail通知信
Private Function ChkEMailRec(pLP01 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select lp38,lp39,lp40 from letterprogress where lp01='" & pLP01 & "' and lp39>0"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      ChkEMailRec = True
   End If
End Function

'Added by Morgan 2025/5/6
Private Sub fnDownPdf()
   Dim stFiles As String, stCP09 As String, stCP10 As String
   Dim arrFileName() As String
   Dim ii As Integer
   Dim hLocalFile As Long
   
   If iPrevRow1 = 0 Then
       MsgBox "請先點選欲下載的資料列！", vbInformation
       Exit Sub
   End If
   SetMouseBusy
   With MSHFlexGrid1
   If .TextMatrix(iPrevRow1, 4) <> "" Then
      stFiles = ""
      stCP09 = GetValue(iPrevRow1, "cp09", MSHFlexGrid1)
      stCP10 = GetValue(iPrevRow1, "cp10", MSHFlexGrid1)
      If m_AttachPath3 <> "" Then
         If PUB_GetAttachFile4Cust(stCP09, stFiles, m_AttachPath3, False, stCP10, , , True) = True Then
            'Modified by Morgan 2025/5/7 改開啟資料夾
            'arrFileName = Split(stFiles, ";")
            'For ii = LBound(arrFileName) To UBound(arrFileName)
            '   If arrFileName(ii) <> "" Then
            '      ShellExecute hLocalFile, "open", m_AttachPath3 & "\" & arrFileName(ii), vbNullString, vbNullString, 1
            '   End If
            'Next
            MsgBox "下載完成！", vbCritical
            ShellExecute hLocalFile, "open", m_AttachPath3, vbNullString, vbNullString, 1
            'end 2025/5/7
         Else
            MsgBox "下載失敗！", vbCritical
         End If
      End If
   End If
   End With
   SetMouseReady
End Sub

'Added by Morgan 2025/5/6 更換
Private Sub fnUpdatePdf()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stFile As String, stFullPath As String
   Dim stCPP01 As String, stCPP02 As String
   Dim fs, f
   
   SetMouseBusy
   stFile = Dir(m_AttachPath3 & "\*.pdf")
   If stFile <> "" Then
      Do While stFile <> ""
         intQ = InStr(stFile, ".")
         stCPP01 = Left(stFile, intQ - 1)
         stCPP02 = Mid(stFile, intQ + 1)
         stSQL = "select cp01,cp02,cp03,cp04 from casepaperpdf,caseprogress where cpp01='" & stCPP01 & "' and cpp02='" & stCPP02 & "' and cp09(+)=cpp01"
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            stFullPath = m_AttachPath3 & "\" & stFile
            If PUB_ChkFileOpening(stFullPath, , False) = True Then
               MsgBox stFullPath & vbCrLf & "檔案正在使用中，請先關閉才可執行更換！", vbExclamation
               GoTo EXITSUB
            End If
            '刪除檔案
            If DelAttFile_PDF(rsQuery("cp01") & "-" & rsQuery("cp02") & "-" & rsQuery("cp03") & "-" & rsQuery("cp04"), stCPP01, stCPP02) = False Then
               GoTo EXITSUB
            Else
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFullPath)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg stFullPath & MsgText(9221)
                  GoTo EXITSUB
               End If
              
               '存檔
               If SaveAttFile_PDF(stCPP01, stFullPath, stCPP02, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                  GoTo EXITSUB
               Else
                  Pub_SaveLog strUserNum, "抽換卷宗區附件：" & stCPP02, rsQuery("cp01"), rsQuery("cp02"), rsQuery("cp03"), rsQuery("cp04"), stCPP01
               End If
               If PUB_DelPCOrgFile(stFullPath) = False Then
                  GoTo EXITSUB
               End If
            End If
         Else
            MsgBox "卷宗區找不到可更換的檔案！" & vbCrLf & vbCrLf & "收文號:" & stCPP01 & " " & "檔名:" & stCPP02 & vbCrLf & "要更換的檔案:" & stFullPath, vbCritical
            GoTo EXITSUB
         End If
         stFile = Dir()
      Loop
      
      KillTemp
      MsgBox "更換完成！"
      intI = 0
      If m_AttachPath2 = "" Then
         fnOpen1 1
      Else
         Do While (Dir(m_AttachPath2 & "\*.pdf") <> "" And intI < 3)
            KillTemp
            Sleep 1000
            intI = intI + 1
         Loop
         
         If Dir(m_AttachPath2 & "\*.pdf") = "" Then
            fnOpen1 1
         End If
      End If
   Else
      MsgBox "無檔案可更換！", vbCritical
   End If
   
EXITSUB:
   Set rsQuery = Nothing
   Set fs = Nothing
   Set f = Nothing
   SetMouseReady
End Sub
