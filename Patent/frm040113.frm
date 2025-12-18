VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040113 
   BorderStyle     =   1  '單線固定
   Caption         =   "公文來函判發作業"
   ClientHeight    =   5724
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   8928
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   4410
      TabIndex        =   25
      Top             =   0
      Width           =   4515
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
         ItemData        =   "frm040113.frx":0000
         Left            =   900
         List            =   "frm040113.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   28
         Top             =   330
         Width           =   3645
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "點我展開"
         Height          =   345
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   26
         Top             =   0
         Width           =   4515
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
         TabIndex        =   27
         Top             =   330
         Width           =   930
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5115
      Left            =   4410
      TabIndex        =   5
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
      Location        =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   630
      Width           =   4410
      _ExtentX        =   7789
      _ExtentY        =   1736
      _Version        =   393216
      Cols            =   5
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|案件性質|案件名稱|收文日"
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
   Begin VB.Frame Frame3 
      Caption         =   "退回意見"
      Height          =   1065
      Left            =   0
      TabIndex        =   23
      Top             =   1620
      Width           =   4380
      Begin VB.TextBox txtLP37 
         Height          =   795
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   24
         Top             =   180
         Width           =   4200
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   345
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   4515
      Begin VB.CommandButton cmdIDS 
         Caption         =   "IDS資料確認"
         Height          =   315
         Left            =   2970
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0080FFFF&
         Caption         =   "退回程序"
         Height          =   315
         Index           =   4
         Left            =   90
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   480
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "專利相關案件(&R)"
         Height          =   315
         Index           =   5
         Left            =   1035
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   480
         Width           =   1905
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0FF&
         Caption         =   "交工程師確認"
         Height          =   315
         Index           =   2
         Left            =   1035
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "進度"
         Height          =   345
         Index           =   0
         Left            =   2340
         TabIndex        =   14
         Top             =   135
         Width           =   600
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0080FFFF&
         Caption         =   "內部收文"
         Height          =   315
         Index           =   1
         Left            =   90
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "判發"
         Height          =   345
         Index           =   0
         Left            =   3525
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   135
         Width           =   825
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   3060
         TabIndex        =   19
         Top             =   210
         Width           =   315
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   3960
      Width           =   4410
      _ExtentX        =   7789
      _ExtentY        =   2371
      _Version        =   393216
      Cols            =   5
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|案件性質|案件名稱|收文日"
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
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   0
      TabIndex        =   11
      Top             =   5160
      Width           =   4515
      Begin VB.CommandButton Command1 
         Caption         =   "進度(&Z)"
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   135
         Width           =   780
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0080FFFF&
         Caption         =   "分案"
         Height          =   315
         Index           =   3
         Left            =   80
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label lblCount2 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2475
         TabIndex        =   20
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "重整"
      Height          =   345
      Left            =   2760
      TabIndex        =   13
      Top             =   0
      Width           =   780
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   990
      TabIndex        =   1
      Top             =   30
      Width           =   1575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2778;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H000080FF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   30
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "無紙"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   2385
      TabIndex        =   29
      Top             =   420
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3285
      TabIndex        =   22
      Top             =   420
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "退回重送"
      Height          =   180
      Left            =   3510
      TabIndex        =   21
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "待分案分析"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "判發人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:雙擊預覽)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   420
      Width           =   1065
   End
End
Attribute VB_Name = "frm040113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; MSHFlexGrid1, MSHFlexGrid2改字型=新細明體-ExtB、Combo1
'Created by Morgan 2014/3/27
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim process_id As Long
Dim process_handle As Long

Dim iPrevRow As Integer '前次點選列
Dim lTotRows As Long, lSelRows As Long
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Dim m_AttachPath As String
Dim oFileSys As New FileSystemObject
Dim oFile As File

Public cmdState As Integer 'Added by Morgan 2014/7/28
Public strCP14 As String 'Added by Morgan 2014/10/15
Public bolMan As Boolean, m_EV02 As String, iPrevRow2 As Integer, lTotRows2 As Long, lSelRows2 As Long 'Add by Lydia 2015/01/20

Dim idxBK As Integer 'Added by Morgan 2016/1/19
Dim idxLP37 As Integer 'Added by Morgan 2019/1/9 退回意見
Dim m_PA09 As String     'add by sonia 2018/2/13
Const lngColor1 As Long = &HC0C0FF '淺紅(審查意見通知已交工程師確認) Added by Morgan 2016/1/20

Private Sub cboAtt_Click()
   Dim hLocalFile As Long
   Dim arrFileName() As String
   
   arrFileName = Split(cboAtt.List(cboAtt.ListIndex), Chr(9))
   WebBrowser1.Navigate m_AttachPath & "\" & arrFileName(0): DoEvents
   'ShellExecute hLocalFile, "open", m_AttachPath & "\" & arrFileName(0), vbNullString, vbNullString, 1
End Sub

Private Sub cmdIDS_Click()
   Dim nFrm As Form
   Set nFrm = Forms(0).GetForm("frm090401_1")
   If Not nFrm Is Nothing Then
      nFrm.m_CP09 = GetValue(iPrevRow, "cp43")
      nFrm.m_bConfirm = True
      nFrm.Show vbModal
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim iRow As Integer, bContinue As Boolean
   Dim iIdx As Integer
   Dim bolShowForm As Boolean
   Dim strCP09 As String
   'Modified by Lydia 2015/01/20
   'Added by Morgan 2014/7/28
'   If Index = 2 Then
'      CmdState = Index
'      PubShowNextData
'      Exit Sub
'   End If
   'end 2014/7/28
   'Add by Lydia 2015/01/20
   If Index <> 3 Then
      SetMouseBusy
      bContinue = False
      With MSHFlexGrid1
      iIdx = GetFieldId("cp10", Me.MSHFlexGrid1)
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "V" Then
            bContinue = True
            strCP09 = .TextMatrix(iRow, GetFieldId("cp09", Me.MSHFlexGrid1))
              
            'Added by Morgan 2019/1/9
            '退回程序
            If Index = 4 Then
               SetMouseReady
               With frm040113_2
               .intRow = iRow
               .strCP09 = strCP09
               'Modified by Lydia 2021/10/07 Label3=>lblFM2
               .lblFM2(0) = MSHFlexGrid1.TextMatrix(iRow, GetFieldId("本所案號", MSHFlexGrid1))
               .lblFM2(1) = MSHFlexGrid1.TextMatrix(iRow, GetFieldId("案件名稱", MSHFlexGrid1))
               .lblFM2(2) = MSHFlexGrid1.TextMatrix(iRow, GetFieldId("案件性質", MSHFlexGrid1))
               .lblFM2(3) = MSHFlexGrid1.TextMatrix(iRow, GetFieldId("收文日", MSHFlexGrid1))
               'end 2021/10/07
               End With
               frm040113_2.Show vbModal
               SetMouseBusy
            'end 2019/1/9
            '判發
            ElseIf Index = 0 Then
                'Added by Morgan 2016/1/20
                If .TextMatrix(iRow, idxBK) = "Y" Then
                   strExc(0) = "select cpp02" & _
                      " from casepaperpdf" & _
                      " where cpp01='" & strCP09 & "' and instr(upper(cpp02),'.INFO.')>0"
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 0 Then
                      MsgBox "卷宗區沒有 info 的電子檔不可判發！", vbCritical, .TextMatrix(iRow, 1)
                      GoTo EXITSUB
                   End If
                   
                'Added by Morgan 2020/12/30
                'IDS提申要檢查是否清單都已確認
                ElseIf GetValue(iRow, "IDS") = "Y" Then
                  strExc(1) = GetValue(iRow, "cp43")
                  If ChkIDS(strExc(1), True) = True Then
                     If MsgBox("IDS資料尚未全部確認，是否確定要判發！", vbYesNo + vbQuestion + vbDefaultButton2, .TextMatrix(iRow, 1)) = vbNo Then
                        GoTo EXITSUB
                     End If
                  End If
                  
                End If
                'end 2016/1/20
                
                'Exit For 'Removed by Morgan 2019/1/9
            
            'Added by Morgan 2016/1/18
            '交工程師確認
            ElseIf Index = 2 Then
               If .TextMatrix(iRow, iIdx) = "1202" Then
                   If .TextMatrix(iRow, idxBK) = "Y" Then
                      MsgBox "承辦人已經不是程序！", vbCritical, .TextMatrix(iRow, 1)
                      GoTo EXITSUB
                   End If
                   
                   bolShowForm = False
                   strCP09 = .TextMatrix(iRow, GetFieldId("cp09", MSHFlexGrid1))
                   .TextMatrix(iRow, GetFieldId("bcp14", MSHFlexGrid1)) = ""
                   strExc(0) = "select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) CaseNo,cu04,pa05,cpm03,c1.cp06,c2.cp14,st02,st03,st04" & _
                      " from caseprogress c1,caseprogress c2,staff,patent,customer,casepropertymap" & _
                      " where c1.cp09='" & strCP09 & "' and c2.cp09(+)=c1.cp43 and st01(+)=c2.cp14" & _
                      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
                      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10"
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      'Modified by Morgan 2016/8/30 改在職也可改承辦人--游經理
                      'If RsTemp("cp14") = "71011" Or RsTemp("st04") = "2" Or RsTemp("st03") = "P12" Or Left(RsTemp("st03"), 2) <> "P1" Then
                      '   bolShowForm = True
                      'End If
                      bolShowForm = True
                      'end 2016/8/30
                      
                      ShowEngForm RsTemp, strCP09, bolShowForm
                      
                      If strCP14 <> "" Then
                         .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid1)) = strCP14
                      Else
                         GoTo EXITSUB
                      End If
                      SetMouseBusy
                   End If
               Else
                   MsgBox "請點選審查意見通知函！", vbCritical, .TextMatrix(iRow, 1)
                   GoTo EXITSUB
               End If
            'end 2016/1/18
            
            'Added by Morgan 2014/5/27
            '內部收文
            ElseIf Index = 1 Then
               'Modified by Morgan 2014/12/25 +1227
               'Memo by Morgan 2024/4/17 案件性質若有調整時內部收文也要同步
               If .TextMatrix(iRow, iIdx) <> "1201" And .TextMatrix(iRow, iIdx) <> "1202" And .TextMatrix(iRow, iIdx) <> "1227" Then
                  MsgBox "非通知修正或審查意見通知函不可設內部收文！(" & .TextMatrix(iRow, 1) & ")", vbCritical
                  GoTo EXITSUB
               'Added by Morgan 2023/6/26
               ElseIf ChkXCust(strCP09) = True Then
                  GoTo EXITSUB
               'Added by Morgan 2014/10/14
               Else
                  bolShowForm = False
                  .TextMatrix(iRow, GetFieldId("bcp09", Me.MSHFlexGrid1)) = ""
                  .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid1)) = ""
                  '若承辦人為71011或已離職時彈視窗選新承辦人
                  'Modified by Morgan 2015/6/25 +承辦人為程序或非專利處也要能改
                  strExc(0) = "select cp09,cp14,st04,st03 from nextprogress,caseprogress,staff" & _
                     " where np01='" & strCP09 & "' and cp43(+)=np01 and cp10(+)=np07 and substr(cp09,1,1)='B' and st01(+)=cp14"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  '已收文
                  If intI = 1 Then
                     .TextMatrix(iRow, GetFieldId("bcp09", Me.MSHFlexGrid1)) = RsTemp("cp09")
                     'Modified by Lydia 2015/01/20
                     'If RsTemp("cp14") = "71011" Or RsTemp("st04") = "2" Then
                      'Modified by Morgan 2016/8/30 改在職也可改承辦人--游經理
                      'If RsTemp("cp14") = "71011" Or RsTemp("st04") = "2" Or RsTemp("st03") = "P12" Or Left(RsTemp("st03"), 2) <> "P1" Then bolShowForm = True
                      bolShowForm = True
                      'end 2016/8/30
                        'Modified by Morgan 2020/3/18 +pa09 Ex:P-121749
                        strExc(0) = "select cp09,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
                           ",cu04,pa05,decode(pa09,'000',cpm03,cpm04) cpm03,cp06,cp14,st02,pa09 from caseprogress,staff,patent,customer,casepropertymap" & _
                           " where cp09='" & RsTemp("cp09") & "' and st01(+)=cp14" & _
                           " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
                           " and cpm01(+)=cp01 and cpm02(+)=cp10"
                     'End If
                  '未收文
                  Else
                     strExc(0) = "select b.cp09,b.cp14,st04,st03 from caseprogress a,caseprogress b,staff" & _
                        " where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and st01(+)=b.cp14"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        .TextMatrix(iRow, GetFieldId("bcp09", Me.MSHFlexGrid1)) = "B"
                        .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid1)) = "" & RsTemp("cp14")
                        'Modified by Lydia 2015/01/20
                        'If RsTemp("cp14") = "71011" Or RsTemp("st04") = "2" Then
                        'Modified by Morgan 2016/8/30 改在職也可改承辦人--游經理
                        'If RsTemp("cp14") = "71011" Or RsTemp("st04") = "2" Or RsTemp("st03") = "P12" Or Left(RsTemp("st03"), 2) <> "P1" Then bolShowForm = True
                        bolShowForm = True
                        'end 2016/8/30
                           'modify by sonia 2018/2/13 +pa09
                           strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
                           ",cu04,pa05,decode(pa09,'000',cpm03,cpm04) cpm03,cp06,cp14,st02,pa09 from nextprogress,caseprogress,staff,patent,customer,casepropertymap" & _
                           " where np01='" & strCP09 & "' and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp09='" & RsTemp("cp09") & "' and st01(+)=cp14" & _
                           " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
                           " and cpm01(+)=np02 and cpm02(+)=np07"
                        'End If
                     End If
                  End If
                  
                  'Modifeid by Lydia 2015/01/20 內部收文可設定基數
                  'If bolShowForm = True Then
                     strExc(5) = RsTemp("cp14")
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        m_PA09 = "" & RsTemp("PA09")  'add by sonia 2018/2/13
                        
                        'Modified by Morgan 2016/1/19 改共用
                        'Modified by Morgan 2016/8/17 已收文帶出收文號
                        'ShowEngForm RsTemp, "B", bolShowForm
                        strExc(1) = .TextMatrix(iRow, GetFieldId("bcp09", Me.MSHFlexGrid1))
                        ShowEngForm RsTemp, strExc(1), bolShowForm
                        'end 2016/8/17
                        'end 2016/1/19
                        
                        If strCP14 <> "" Then
                           .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid1)) = strCP14
                        Else
                           GoTo EXITSUB
                        End If
                        SetMouseBusy
                     End If
                  'End If
               'end 2014/10/14
               End If
            End If
         End If
      Next
      End With
      If bContinue = False Then
         MsgBox "請先勾選(V)資料列！", vbInformation
         
      'Modified by Morgan 2016/1/19 +2
      ElseIf Index = 0 Or Index = 1 Or Index = 2 Then
         'FormSave Index
         'Add by Lydia 2015/01/20
          FormSave Index, MSHFlexGrid1
      End If
      
   Else
   'Add by Lydia 2015/01/20輸入舉發,舉發答辯,准,駁產生內部收文分析(B類941),若原承辦工程師已離職，請將分析這道掛至游經理確認內部收文系統，並由游經理更改承辦人後發ｅ－ｍａｉｌ通知承辦人
        SetMouseBusy2
        bContinue = False
        With MSHFlexGrid2
        For iRow = 1 To .Rows - 1
           If .TextMatrix(iRow, 0) = "V" Then
                bContinue = True: bolShowForm = True
                .TextMatrix(iRow, GetFieldId("bcp09", Me.MSHFlexGrid2)) = ""
                .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid2)) = ""
                strCP09 = .TextMatrix(iRow, GetFieldId("cp09", Me.MSHFlexGrid2))
                strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo,cu04,pa05,cpm03,cp06,cp14,st02 from caseprogress,staff,patent,customer,casepropertymap " & _
                            "where cp09='" & strCP09 & "' and st01(+)=cp14 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) " & _
                            "and cpm01(+)=cp01 and cpm02(+)=cp10 "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                   
                      With frm040113_1
                        .lblRecNo = strCP09
                        .lblCaseNo = RsTemp("CaseNo")
                        .lblAppName = RsTemp("cu04")
                        .lblCaseName = RsTemp("pa05")
                        .lblProperty = RsTemp("cpm03")
                        .lblOurDeadLine = ChangeWStringToTDateString("" & RsTemp("cp06"))
                        strExc(1) = RsTemp("cp14") & " " & RsTemp("st02")
                        strExc(0) = "select st01||' '||st02 from staff where st01<>'71011' and st03 in ('P10','P11') and st04='1' order by st01"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           Do While Not RsTemp.EOF
                              .cboCP14.AddItem RsTemp(0)
                              RsTemp.MoveNext
                           Loop
                        End If
                        .cboCP14 = strExc(1)
                        End With
                        SetMouseReady2
                        If bolShowForm = True Then frm040113_1.mInputKey = "1" 'Add by Lydia 2015/01/20預設選承辦人
                        frm040113_1.Show vbModal
                        If strCP14 <> "" Then
                           .TextMatrix(iRow, GetFieldId("bcp14", Me.MSHFlexGrid2)) = strCP14
                        Else
                           GoTo EXITSUB
                        End If
                        SetMouseBusy2
                   End If
           End If
        Next
        End With
        If bContinue = False Then
           MsgBox "請先勾選(V)資料列！", vbInformation
        Else
           FormSave Index, MSHFlexGrid2
        End If
   'end 2015/01/20
   End If
EXITSUB:

   SetMouseReady
   SetMouseReady2
End Sub
'Modified by Lydia 2015/01/21 + iPt, FGrid
'Public Sub PubShowNextData()
'   If iPrevRow = 0 Then Exit Sub
'   Select Case CmdState
'   Case 2
'      Me.Enabled = False
'      If fnSaveParentForm(Me) = False Then
'         Me.Enabled = True
'         Exit Sub
'      End If
'      Screen.MousePointer = vbHourglass
'      frm100101_2.Show
'      frm100101_2.Tag = Pub_RplStr(MSHFlexGrid1.TextMatrix(iPrevRow, 1))
'      frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
'      frm100101_2.StrMenu
'      Screen.MousePointer = vbDefault
'      Me.Enabled = True
'   End Select
'End Sub
Public Sub PubShowNextData(ByRef iPt As Integer, ByRef Fgrid As MSHFlexGrid, Optional Index As Integer)
   If iPt = 0 Then Exit Sub
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If Index = 5 Then
      frm100101_h.Show
      frm100101_h.KeyString = Pub_RplStr(Fgrid.TextMatrix(iPt, 1))
      frm100101_h.SearchKind = "本所案號"
      frm100101_h.cmdOK(3).Visible = False '下一筆按鈕隱藏
      frm100101_h.StrMenu
   Else
      frm100101_2.Show
      frm100101_2.Tag = Pub_RplStr(Fgrid.TextMatrix(iPt, 1))
      frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
      frm100101_2.StrMenu
   End If
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Sub fnOpen1()
   Dim stFiles As String
   Dim stFileName As String
   Dim arrFileName() As String
   Dim idx As Integer
   Dim bJoin As Boolean
   Dim stFileDescs As String
   Dim stCP09 As String
   Dim stCP10 As String, stCP43 As String 'Added by Morgan 2021/11/29

   If iPrevRow = 0 Then
      MsgBox "請先點選要預覽的資料列！", vbInformation
   Else
      SetMouseBusy
      With MSHFlexGrid1
      If .TextMatrix(iPrevRow, 1) <> "" Then
         WebBrowser1.Navigate "about:blank": DoEvents
         bJoin = True
         stCP09 = GetValue(iPrevRow, "cp09")
         stCP10 = GetValue(iPrevRow, "cp10") 'Added by Morgan 2021/11/29
         stCP43 = GetValue(iPrevRow, "cp43") 'Added by Morgan 2021/11/29
         stFiles = ""
         'Modified by Morgan 2019/1/23 +stFileDescs
         'Modified by Morgan 2021/11/29 +stCP10,stCP43
         If GetAttachFile(stCP09, stFiles, m_AttachPath, bJoin, , stFileDescs, stCP10, stCP43) = True Then
            If stFileDescs <> "" Then
               SetValue iPrevRow, "FDesc", stFileDescs
            Else
               stFileDescs = GetValue(iPrevRow, "FDesc")
            End If
            arrFileName = Split(stFiles, ";")
            For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
               If arrFileName(idx) <> "" Then
                  stFileName = m_AttachPath & "\" & arrFileName(idx)
                  WebBrowser1.Navigate stFileName
                  SetAttList stFileDescs
                  Exit For
               End If
            Next
         End If
      End If
      End With
      SetMouseReady
   End If
End Sub

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      Command5.Value = True
      Combo1.Tag = Combo1
   End If
End Sub
Private Sub Form_Activate()
   Static bDone As Boolean
   'Added by Morgan 2014/5/29
   'Removed by Morgan 2014/7/25 取消,因使用共同查詢會不斷的彈出 --游經理
   'If Me.WindowState = 0 Then Me.WindowState = 2
   If bDone = False Then
      Combo1_Click
      bDone = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   SetAttList 'Added by Morgan 2019/1/23
   'Modified by Morgan 2014/5/9 考慮多人系統共用問題改放員工編號資料夾
   'm_AttachPath = App.path & "\" & Pub_GetSpecMan("EDocPath")
   m_AttachPath = App.path & "\" & strUserNum
   KillTemp
   Me.WindowState = 2 'Added by Morgan 2014/5/15
   SetCombo1
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If Command4.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub
'Modified by Morgan 2019/1/19 加重試3次後彈訊息(檔案被鎖住時無法刪除)
Private Sub KillTemp()
   Dim iTimes As Integer
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.pdf"
   End If
   Exit Sub
   
ErrHnd:
   If iTimes < 2 Then
      iTimes = iTimes + 1
      Sleep 1000
      Resume
   Else
      'MsgBox "暫存檔無法清除！" & vbCrLf & vbCrLf & "請重新執行本作業，否則有可能載入的不是最新的定稿！", vbExclamation
   End If
   Err.Clear
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command4_Click()
   If Command4.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub RePosForm(pFull As Boolean)
   Static lngLeft As Long
   Dim a1 As Long
   
   If Forms(0).WindowState <> 1 Then
      'Modified by Mogrgan 2019/1/23 加pdf檔下拉選單
      Frame3.Visible = False
      If lngLeft = 0 Then lngLeft = WebBrowser1.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         Command4.Caption = "點我展開"
      End If
      WebBrowser1.Width = Me.Width - WebBrowser1.Left - 90
      WebBrowser1.Height = Me.Height - Frame4.Height - 390
      Frame4.Left = WebBrowser1.Left
      Frame4.Width = WebBrowser1.Width
      Command4.Width = Frame4.Width
      cboAtt.Width = Frame4.Width - cboAtt.Left
      'end 2019/1/23
      
      
      If bolMan = True Or Pub_StrUserSt03 = "M51" Then
          'Removed by Morgan 2025/2/20 取消分析及交工程師確認功能--游協理
          ''Add by Lydia 2015/01/20 游經理才可見
          'a1 = (Me.Height - MSHFlexGrid1.Top - Frame1.Height - 1500) / 2
          ''Modified by Morgan 2019/1/9 調整比例(原高度相同)
          'MSHFlexGrid1.Height = a1 * 6 / 5
          'MSHFlexGrid2.Height = a1 * 4 / 5
          'Frame1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height - 100
          'Label3.Top = Frame1.Top + Frame1.Height + 100
          'MSHFlexGrid2.Top = Label3.Top + 400
          'If pFull = True Then
          '   Frame2.Visible = False
          'Else
          '   Frame2.Visible = True
          '   Frame2.Top = MSHFlexGrid2.Top + MSHFlexGrid2.Height + 10
          'End If
          MSHFlexGrid1.Height = Me.Height - MSHFlexGrid1.Top - Frame1.Height - 300
          Frame1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height - 100
          Frame2.Visible = False
          cmdOK(2).Visible = False
          'end 2025/2/20
      Else
          MSHFlexGrid1.Height = Me.Height - MSHFlexGrid1.Top - Frame1.Height - 300
          Frame1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height - 100
          Frame2.Visible = False
          'Added by Morgan 2018/10/31
          cmdOK(1).Visible = False
          cmdOK(2).Visible = False
          'end 2018/10/31
          'Add by Amy 2019/12/04+商標(只顯示 進度/退回承序/判發 鈕)
          If Left(Pub_StrUserSt03, 2) = "P2" Then
            Command1(5).Visible = False '專利相關案件 鈕不顯示
            '退回程序 鈕調位置
            cmdOK(4).Top = 145
            cmdOK(4).Left = 1400
          End If
          'end 2019/12/04
          
      End If
      
      With MSHFlexGrid1
      If txtLP37 <> "" Then
         .Height = Frame1.Top - .Top - Frame3.Height
         Frame3.Top = .Top + .Height + 50
         Frame3.Visible = True
      Else
         .Height = Frame1.Top - .Top + 50
      End If
      End With
      
   End If
End Sub

Private Sub Command5_Click()
   WebBrowser1.Navigate "about:blank": DoEvents 'Added by Morgan 2019/1/19
   SetAttList 'Added by Morgan 2019/1/23
   SetMouseBusy
   QueryData
   SetMouseReady
   'Add by Lydia 2015/01/20
   SetMouseBusy2
   QueryData2
   SetMouseReady2
   KillTemp 'Added by Morgan 2019/1/19 要先刪除暫存否則退會重送的定稿可能會讀到舊的
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 1140, 800, 1000, 825)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
      
      iPrevRow = 0
      lTotRows = 0
      lSelRows = 0
      lblCount = lSelRows & " / " & lTotRows
   End If
   .FixedCols = 2
   .FormatString = "V|本所案號|案件性質|案件名稱|收文日"
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

Private Sub QueryData()
   Dim iRow As Integer
   Dim stCon As String
   Dim idx As Integer, idxCP09 As Integer, idxST03 As Integer, idxCP10 As Integer, idxECase As String
   
   If Trim(Left(Combo1.Text, 6)) <> "" Then
      stCon = " and lp04='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   End If
   
   SetGrid True
   
   'modify by sonia 2016/7/13 剔除FMP案故加入substr(cp12,1,1)<>'F'條件
   'Modified by Morgan 2018/9/26 收文日改先抓櫃台收文日(cp119)
   'Modified by Morgan 2018/10/9 +判斷已發文(CFP案C類來函要確認後才會上發文日)
   'Modified by Morgan 2019/1/10 +退回未上傳不要顯示
   'Modified by Morgan 2019/1/23 +FDesc
   'Modified by Morgan 2019/4/23 +ECase E化
   'Modified by Morgan 2019/8/5 假發文的程序不必判發(CFP已提申例外) Ex:P-121392--游經理
   'Modify by Amy 2019/11/05 增加Trademark,ServicePractice
   'Modified by Morgan 2020/1/8 +LP43判斷(因CFP要工程師判發的已提申不一定有客戶函)
   'Modified by Morgan 2021/3/26 CFP已提申不必排除國外部案件 Ex:CFP-032254
   'Modified by Morgan 2021/5/28 +LP50交工程師確認日期
   'Modified by Lydia 2024/02/19 人工上傳檔案的副檔名統一為小寫; substr(cpp02(+),-8)='.CUS.PDF'>> upper(substr(cpp02(+),-8))='.CUS.PDF' ; ex.CFP-32576的CB3005255
   strExc(0) = "select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號" & _
      ",Decode(Decode(pa01,null,Decode(tm01,null,sp09,tm10),pa09), '000',cpm03,cpm04) 案件性質" & _
      ",Decode(pa01,null,Decode(tm01,null,sp05,tm05),pa05) 案件名稱,sqldatet(nvl(c1.cp119,c1.cp05)) 收文日,c1.cp09" & _
      ",c1.cp01||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||c1.cp04)||'.'||c1.cp10||'.PDF' DocName" & _
      ",c1.cp10,'' BCP09,'' BCP14,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp06,st03" & _
      ",'' BK,lp37,'' FDesc,decode(c1.cp154,'QPGMR','Y') ECase,decode(c1.cp01||c1.cp10||c2.cp10,'CFP1909214','Y') IDS,c1.cp43" & _
      " From letterprogress,caseprogress c1, patent, casepropertymap,staff,casepaperpdf,TradeMark,ServicePractice,caseprogress c2" & _
      " where lp05=0 and NVL(LP43,LP10)='Y' and lp03>0" & stCon & " and c1.cp09(+)=lp01" & _
      " and ((substr(c1.cp12,1,1)<>'F' and nvl(c1.cp27,lp50)>19221111) or c1.cp01||c1.cp10='CFP1909')" & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 " & _
      " and tm01(+)=c1.cp01 And tm02(+)=c1.cp02 And tm03(+)=c1.cp03 And tm04(+)=c1.cp04 " & _
      " and sp01(+)=c1.cp01 And sp02(+)=c1.cp02 And sp03(+)=c1.cp03 And sp04(+)=c1.cp04 " & _
      " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 and st01(+)=c1.cp14" & _
      " and cpp01(+)=lp01 and upper(substr(cpp02(+),-8))='.CUS.PDF' and cpp10(+)<>'D'" & _
      " and (lp36 is null or cpp02 is not null) and c2.cp09(+)=c1.cp43" & _
      " order by c1.cp05,c1.cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      lTotRows = RsTemp.RecordCount
      lblCount = lSelRows & " / " & lTotRows
      idx = GetFieldId("案件性質", MSHFlexGrid1)
      idxCP09 = GetFieldId("CP09", MSHFlexGrid1)
      idxST03 = GetFieldId("ST03", MSHFlexGrid1)
      idxCP10 = GetFieldId("CP10", MSHFlexGrid1)
      idxBK = GetFieldId("BK", MSHFlexGrid1)
      idxLP37 = GetFieldId("LP37", MSHFlexGrid1) 'Added by Morgan 2019/1/9
      idxECase = GetFieldId("ECase", MSHFlexGrid1) 'Added by Morgan 2019/4/23
      SetGrid
      
      'Added by Morgan 2016/1/19
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, idxCP09), "1")
         'Add byAmy 2019/12/04 +if 判斷非商標
         If Left(Pub_StrUserSt03, 2) <> "P2" Then
            '審查意見通知已交工程師(非程序)確認時該列底色呈現淡紅色
            If .TextMatrix(iRow, idxCP10) = "1202" And .TextMatrix(iRow, idxST03) <> "P12" And Left(.TextMatrix(iRow, idxST03), 1) <> "F" Then
               .TextMatrix(iRow, idxBK) = "Y"
               .row = iRow
               'Modified by Morgan 2018/8/14
               'SetRowBK lngColor1
               SetRowBK cmdOK(2).BackColor
               'end 2018/8/14
            End If
         End If
         
         'Added by Morgan 2019/1/9
         '退回
         If .TextMatrix(iRow, idxLP37) <> "" Then
            .row = iRow
            SetRowBK Label5.BackColor
         End If
         'end 2019/1/9
         
         'Added by Morgan 2019/4/23
         'E化案件
         If .TextMatrix(iRow, idxECase) = "Y" Then
            .row = iRow
            .col = 0
            .CellBackColor = lblColor(1).BackColor
            lblColor(1).Visible = True
            lblColorDesc(1).Visible = True
         End If
         'end 2019/4/23
      Next
      'end 2016/1/19
      
      .col = 1: .row = 1
      SelectRow 1
      .Visible = True
      End With
   Else
      MsgBox "無待判發資料！", vbExclamation
   End If
   'WebBrowser1.Navigate "about:blank" 'Removed by Morgan 2019/1/23
End Sub

'Added by Morgan 2016/1/20
Private Sub SetRowBK(Optional pCellBackColor As Long = 0)
   Dim ii As Integer
   With MSHFlexGrid1
   If pCellBackColor = 0 Then pCellBackColor = .BackColor
   For ii = .FixedCols To .Cols - 1
      .col = ii
      .CellBackColor = pCellBackColor
   Next
   End With
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

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Sub ClosePDF()
    Dim lRtn As Long
    If process_handle <> 0 Then
      lRtn = TerminateProcess(process_handle, 0&)
      lRtn = CloseHandle(process_handle)
      process_handle = 0
    End If
End Sub
'Add by Lydia 2015/01/22 + UpdFlexGrid
'Private Function FormSave(pIdx As Integer) As Boolean
Private Function FormSave(pIdx As Integer, ByRef UpdFlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer, idxCP09 As Integer, idxCP10 As Integer
   'Added by Morgan 2014/10/16
   Dim idxBCP09 As Integer, idxBCP14 As Integer, idxCP06 As Integer
   Dim idxCP01 As Integer, idxCP02 As Integer, idxCP03 As Integer, idxCP04 As Integer
   Dim idxCaseNo As Integer, idxCaseName As Integer, idxCP10C As Integer
   Dim strBCP09 As String, strBCP13 As String, strBCP12 As String
   Dim bolMail As Boolean, strSub As String, strContent As String, strTo As String, stAttFileName As String, strCP09 As String
   Dim strBCP10 As String  'add by sonia 2018/2/13
   
On Error GoTo ErrHnd
   
   idxCP09 = GetFieldId("cp09", UpdFlexGrid)
   'Added by Morgan 2014/10/16
   idxCP10 = GetFieldId("cp10", UpdFlexGrid)
   idxBCP09 = GetFieldId("bcp09", UpdFlexGrid)
   idxBCP14 = GetFieldId("bcp14", UpdFlexGrid)
   idxCP01 = GetFieldId("cp01", UpdFlexGrid)
   idxCP02 = GetFieldId("cp02", UpdFlexGrid)
   idxCP03 = GetFieldId("cp03", UpdFlexGrid)
   idxCP04 = GetFieldId("cp04", UpdFlexGrid)
   idxCP06 = GetFieldId("cp06", UpdFlexGrid)
   'end 2014/10/16
   
   'Added by Morgan 2016/1/20
   idxCaseNo = GetFieldId("本所案號", UpdFlexGrid)
   idxCaseName = GetFieldId("案件名稱", UpdFlexGrid)
   idxCP10C = GetFieldId("案件性質", UpdFlexGrid)
   'end 2016/1/20
   
   With UpdFlexGrid
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT
         
         '判發
         If pIdx = 0 Then
         
            'Added by Morggan 2016/3/11
            '交工程師確認的要更新發文日,齊備日,完稿日,會稿日,會稿完成日(專利用)
            If .TextMatrix(iRow, idxBK) = "Y" Then
               strSql = "update engineerprogress set ep06=nvl(ep06," & strSrvDate(1) & "),ep07=nvl(ep07," & strSrvDate(1) & "),ep08=nvl(ep08," & strSrvDate(1) & "),ep09=nvl(ep09," & strSrvDate(1) & "),ep34=nvl(ep34,'N') where ep02='" & .TextMatrix(iRow, idxCP09) & "'"
               cnnConnection.Execute strSql, intI
               
               strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp27 is null"
               cnnConnection.Execute strSql, intI
            End If
            'end 2016/3/11
            
            '更新判發人,判發日
            strSql = "update letterprogress set " & "lp04='" & strUserNum & "',lp05=" & strSrvDate(1) & " where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2022/2/16
            '更新全E化案件發文室發文日(智權人員寄發文件會看到，要和判發日期相同才能正確預估程序預定會EMail的日期)
            strSql = "update caseprogress set (cp127,cp128)=(select lp05,lp17 from letterprogress where lp01=cp09)" & _
               " where cp09='" & .TextMatrix(iRow, idxCP09) & "' and cp154='QPGMR' and cp127>0"
            cnnConnection.Execute strSql, intI
            If intI = 1 Then
               strSql = "update caseprogress set cp154='QPGMR' where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
               cnnConnection.Execute strSql, intI
            End If
            'end 2022/2/16
            
            'Add by Amy 2019/12/04 +if 非商標
            If Left(Pub_StrUserSt03, 2) <> "P2" Then
                'Added by Morgan 2014/10/22
                '通知修正若要函知時刪除內部收文之修正並恢復下一程序期限
                If .TextMatrix(iRow, idxCP10) = "1201" Then
                   strSql = "delete caseprogress where cp43='" & .TextMatrix(iRow, idxCP09) & "' and cp10='204' and substr(cp09,1,1)='B'"
                   cnnConnection.Execute strSql, intI
                   strSql = "update nextprogress set np06='' where np01='" & .TextMatrix(iRow, idxCP09) & "' and np07='204' and np06='Y'"
                   cnnConnection.Execute strSql, intI
                End If
                'end 2014/10/22
                
                'Added by Morgan 2018/9/27
                If .TextMatrix(iRow, idxCP10) = "1909" Then
                  'Added by Morgan 2020/1/8 加判斷有客戶函才EMail(因不通知客戶改也要能判發)
                  strSql = "update letterprogress set lp04=lp04 where lp01='" & .TextMatrix(iRow, idxCP09) & "' and lp10='Y'"
                  cnnConnection.Execute strSql, intI
                  If intI = 1 Then
                  'end 2020/1/8
                     strExc(0) = .TextMatrix(iRow, idxCaseNo) & "(" & .TextMatrix(iRow, idxCP10C) & ")客戶函已判發!!"
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                       " select '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                       ",'" & ChgSQL(strExc(0)) & "','如旨' from caseprogress where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
                     cnnConnection.Execute strSql, intI
                  End If 'Added by Morgan 2020/1/8
                End If
                'end 2018/9/27
                
                'Added by Morgan 2024/8/15
                'X15833160欣興電子P案收到官方來函時,判發後系統自動發MAIL提醒智權同仁收文分析
                If .TextMatrix(iRow, idxCP01) = "P" And Left(.TextMatrix(iRow, idxCP09), 1) = "C" And InStr("1002、1202、1203、1209、1225、1226、1227、1802、1810", .TextMatrix(iRow, idxCP10)) > 0 Then
                     strExc(0) = .TextMatrix(iRow, idxCaseNo) & "「" & .TextMatrix(iRow, idxCaseName) & "」->" & .TextMatrix(iRow, idxCP10C) & "，請儘速收文分析！"
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                       " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                       ",'" & ChgSQL(strExc(0)) & "','如旨' from caseprogress,patent where cp09='" & .TextMatrix(iRow, idxCP09) & "'" & _
                       " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa26 like 'X1583316%'"
                     cnnConnection.Execute strSql, intI
                End If
                'end 2024/8/15
                
            End If 'end 非商標
                  
         '內部收文(專利用)
         ElseIf pIdx = 1 Then
         
               'Modified by Morgan 2016/2/18 因為有可能取消內部收文,lp11改不清除以便恢復資料
               strSql = "update letterprogress set lp04='" & strUserNum & "',lp05=" & strSrvDate(1) & ",lp10='N' where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
               cnnConnection.Execute strSql, intI
               
               PUB_DelFtpFile2 .TextMatrix(iRow, idxCP09), " and instr(upper(cpp02),'.CUS.PDF')>0" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
               
               'Added by Morgan 2014/6/11 客戶函也要刪除--玲玲
               strSql = "delete casepaperpdf where cpp01='" & .TextMatrix(iRow, idxCP09) & "' and instr(upper(cpp02),'.CUS.PDF')>0"
               cnnConnection.Execute strSql, intI
         
            'Added by Morgan 2014/10/16
            If .TextMatrix(iRow, idxBCP09) <> "" Then
               '新增內部收文
               If .TextMatrix(iRow, idxBCP09) = "B" Then
                  strBCP09 = AutoNo("B", 6)
                  strBCP13 = PUB_GetAKindSalesNo(.TextMatrix(iRow, idxCP01), .TextMatrix(iRow, idxCP02), .TextMatrix(iRow, idxCP03), .TextMatrix(iRow, idxCP04))
                  strBCP12 = GetSalesArea(strBCP13)
                  'modify by sonia 2018/2/13 台灣案固定204,非台灣案依下一程序管制之案件性質
                  'strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp20,cp26,cp32,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43)" & _
                  '   " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",cp06,cp07,cp08,'" & strBCP09 & "','204','" & strBCP12 & "','" & strBCP13 & "'" & _
                  '   ",'" & .TextMatrix(iRow, idxBCP14) & "',0,0,0,'N','N','N',cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp09" & _
                  '   " from caseprogress where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
                  If m_PA09 = "000" Then
                     strBCP10 = "204"
                  Else
                     strExc(0) = " select np07 from nextprogress where np01='" & .TextMatrix(iRow, idxCP09) & "' and np07 in ('204','205') and np06 is null "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strBCP10 = "" & RsTemp.Fields("np07")
                     End If
                  End If
                  
                  strExc(3) = "cp06"
                  'Added by Morgan 2023/7/10
                  '專利國內部P案之B類收文之案件性質203主動修正、204  修正、205申復、206補充說明,請將本所期限調整為收文日起算7個工作日--郭
                  'Removed by Morgan 2023/8/1 取消-郭
                  'strExc(1) = .TextMatrix(iRow, idxCP06)
                  'strExc(2) = PUB_GetPBRecCP06(strExc(1), .TextMatrix(iRow, idxCP01), strBCP10)
                  'If strExc(1) <> strExc(2) Then strExc(3) = strExc(2)
                  'end 2023/7/10
                  
                  strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp20,cp26,cp32,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43)" & _
                     " select cp01,cp02,cp03,cp04," & strSrvDate(1) & "," & strExc(3) & ",cp07,cp08,'" & strBCP09 & "','" & strBCP10 & "','" & strBCP12 & "','" & strBCP13 & "'" & _
                     ",'" & .TextMatrix(iRow, idxBCP14) & "',0,0,0,'N','N','N',cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp09" & _
                     " from caseprogress where cp09='" & .TextMatrix(iRow, idxCP09) & "'"
                  cnnConnection.Execute strSql, intI
                  'end 2018/2/13
                  
                  'Modified by Morgan 2023/9/15 +np24
                  strSql = "update nextprogress set np06='Y',np24='" & strBCP09 & "' where np01='" & .TextMatrix(iRow, idxCP09) & "' and np07 in ('204','205') and np06 is null"
                  cnnConnection.Execute strSql, intI
                   'Added by Morgan 2019/12/20
                  strExc(1) = ""
                  ClsPDGetCaseProperty .TextMatrix(iRow, idxCP01), strBCP10, strExc(1), IIf(m_PA09 = "000", False, True)
                  Call PUB_UpdRelationCaseFixEP(.TextMatrix(iRow, idxCP01), .TextMatrix(iRow, idxCP02), .TextMatrix(iRow, idxCP03), .TextMatrix(iRow, idxCP04), strBCP10, strExc(1))
                  'end 2019/12/20

               Else
                  strBCP09 = .TextMatrix(iRow, idxBCP09)
                  '更新內部收文承辦人
                  If .TextMatrix(iRow, idxBCP14) <> "" Then
                     '更新承辦人
                     strSql = "update caseprogress set cp14='" & .TextMatrix(iRow, idxBCP14) & "' where cp09='" & strBCP09 & "'"
                     cnnConnection.Execute strSql, intI
                  End If
               End If
               'Add by Lydia 2015/01/20 開放基數(計件值)修改
               If Val(m_EV02) <> 0 Then
                    strExc(0) = " select * from ExValue where EV01='" & strBCP09 & "' "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                       strSql = "update Exvalue set ev02=" & Val(m_EV02) & "  where EV01='" & strBCP09 & "' "
                    Else
                       strSql = "insert into ExValue(EV01,EV02) values ('" & strBCP09 & "'," & Val(m_EV02) & ") "
                    End If
                    Pub_SeekTbLog strSql 'Added by Morgan 2016/2/16
                    cnnConnection.Execute strSql, intI
               End If
               
               
               '發EMail給智權同仁及承辦工程師
               'Modified by Morgan 2016/8/17 +cpm03
               ''Modified by Morgan 2024/4/17 改共用(內部收文也要用)
               'strExc(0) = "select nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CuName,cp14||' '||st02 ProName,cp13,cp14,decode(pa09,'000',cpm03,cpm04) cpm03" & _
               '   " from caseprogress,patent,customer,staff,casepropertymap" & _
               '   " where cp09='" & strBCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
               '   " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)  and st01(+)=cp14 and cpm01(+)=cp01 and cpm02(+)=cp10"
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               'If intI = 1 Then
               '   'Modified by Morgan 2016/8/17
               '   'strSub = .TextMatrix(iRow, 1) & "「" & .TextMatrix(iRow, 3) & "」 --> " & .TextMatrix(iRow, 2) & "(已內部收文修正)"
               '   strSub = .TextMatrix(iRow, idxCaseNo) & "「" & .TextMatrix(iRow, idxCaseName) & "」 --> " & .TextMatrix(iRow, idxCP10C) & "(已內部收文" & RsTemp("cpm03") & ")"
               '   'end 2016/8/17
               '   strContent = "本所案號：" & .TextMatrix(iRow, idxCaseNo)
               '   strContent = strContent & vbCrLf & "申請人　：" & RsTemp("CuName")
               '   strContent = strContent & vbCrLf & "案件名稱：" & .TextMatrix(iRow, idxCaseName)
               '   strContent = strContent & vbCrLf & "案件性質：" & .TextMatrix(iRow, idxCP10C)
               '   strContent = strContent & vbCrLf & "本所期限：" & ChangeWStringToWDateString(.TextMatrix(iRow, idxCP06))
               '   strContent = strContent & vbCrLf & "來函內容：請至卷宗區參看官方來函"
               '   strContent = strContent & vbCrLf & "修正承辦人：" & RsTemp("ProName")
               '
               '   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               '      " values( '" & strUserNum & "','" & RsTemp("cp13") & ";" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd')" & _
               '      ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "')"
               '   cnnConnection.Execute strSql, intI
               'End If
               PUB_PBCPInform strBCP09
               'end 2024/4/17
            End If
            'end 2014/10/16
            
         'Added by Morgan 2016/1/20
         '交工程師確認(專利用)
         ElseIf pIdx = 2 Then
            strCP09 = .TextMatrix(iRow, idxCP09)
            'Modified by Morgan 2016/3/11 發文日要先拿掉等判發再連會完日等一起上否則批次會發通知給工程師
            strSql = "update caseprogress set cp14='" & .TextMatrix(iRow, idxBCP14) & "',cp27=null where cp09='" & strCP09 & "'"
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2021/5/28
            strSql = "update letterprogress set " & "lp04='" & strUserNum & "',lp50=" & strSrvDate(1) & " where lp01='" & .TextMatrix(iRow, idxCP09) & "'"
            cnnConnection.Execute strSql, intI
            'end 2021/5/28
            
            If Val(m_EV02) <> 0 Then
                 strExc(0) = " select * from ExValue where EV01='" & strCP09 & "' "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 1 Then
                    strSql = "update Exvalue set ev02=" & Val(m_EV02) & "  where EV01='" & strCP09 & "' "
                 Else
                    strSql = "insert into ExValue(EV01,EV02) values ('" & strCP09 & "'," & Val(m_EV02) & ") "
                 End If
                 cnnConnection.Execute strSql, intI
            End If
            
            'Added by Morgan 2021/2/23 上齊備，預設游經理判發，不會稿
            'Modified by Morgan 2025/2/21 73022->Left(pub_PMan, 5)
            pub_PMan = Pub_GetSpecMan("專利處特定編號")
            strSql = "update engineerprogress set ep06=nvl(ep06," & strSrvDate(1) & "),ep34=nvl(ep34,'N'),ep40='" & Left(pub_PMan, 5) & "' where ep02='" & strCP09 & "'"
            'end 2025/2/21
            cnnConnection.Execute strSql, intI
            'end 2021/2/23
               
            strExc(0) = "select nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CuName,cp14||' '||st02 ProName" & _
               ",cp01||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04)||'.PDF' Doc2" & _
               ",cp05,cp14,cp64 from caseprogress,patent,customer,staff" & _
               " where cp09='" & strCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
               " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)  and st01(+)=cp14"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stAttFileName = RsTemp("Doc2")
               strSub = "已收到 " & .TextMatrix(iRow, idxCaseNo) & " 通知" & .TextMatrix(iRow, idxCP10C) & "(如附件),請立即確認是否屬本所失誤..."
               strContent = "已收到 " & .TextMatrix(iRow, idxCaseNo) & " 通知" & .TextMatrix(iRow, idxCP10C) & "(如附件),請立即確認是否屬本所失誤，若是請提供通知客戶函的補充內容，不論是否需補充內容均請回覆(註:以此封郵件直接回覆)。" & vbCrLf
               strContent = strContent & vbCrLf & "附件：" & RsTemp("Doc2")
               strContent = strContent & vbCrLf & "本所案號：" & .TextMatrix(iRow, idxCaseNo)
               strContent = strContent & vbCrLf & "案件名稱：" & .TextMatrix(iRow, idxCaseName)
               strContent = strContent & vbCrLf & "申請人　：" & RsTemp("CuName")
               strContent = strContent & vbCrLf & "來函日期：" & ChangeWStringToWDateString(RsTemp("cp05"))
               strContent = strContent & vbCrLf & "來函性質：" & .TextMatrix(iRow, idxCP10C)
               strContent = strContent & vbCrLf & "進度備註：" & RsTemp("cp64")
               strTo = RsTemp("cp14")
               bolMail = True
            End If
         'end 2016/1/20
         
         '分案(專利用)
         ElseIf pIdx = 3 Then
             strBCP09 = .TextMatrix(iRow, idxCP09)
             '更新內部收文承辦人
             If .TextMatrix(iRow, idxBCP14) <> "" Then
                strSql = "update caseprogress set cp14='" & .TextMatrix(iRow, idxBCP14) & "' where cp09='" & strBCP09 & "'"
                cnnConnection.Execute strSql, intI
             End If
             '開放基數(計件值)修改
             If Val(m_EV02) <> 0 Then
                 strExc(0) = " select * from ExValue where EV01='" & strBCP09 & "' "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 1 Then
                    strSql = "update Exvalue set ev02=" & Val(m_EV02) & "  where EV01='" & strBCP09 & "' "
                 Else
                    strSql = "insert into ExValue(EV01,EV02) values ('" & strBCP09 & "'," & Val(m_EV02) & ") "
                 End If
                 cnnConnection.Execute strSql, intI
             End If
             
            '發EMail給承辦工程師
            strExc(0) = "select nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CuName,cp14||' '||st02 ProName,cp13,cp14" & _
               " from caseprogress,patent,customer,staff" & _
               " where cp09='" & strBCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
               " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)  and st01(+)=cp14"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strSub = .TextMatrix(iRow, idxCaseNo) & "「" & .TextMatrix(iRow, idxCaseName) & "」 --> " & .TextMatrix(iRow, idxCP10C) & "(已修正承辦人)"
               strContent = "本所案號：" & .TextMatrix(iRow, idxCaseNo)
               strContent = strContent & vbCrLf & "申請人　：" & RsTemp("CuName")
               strContent = strContent & vbCrLf & "案件名稱：" & .TextMatrix(iRow, idxCaseName)
               strContent = strContent & vbCrLf & "案件性質：" & .TextMatrix(iRow, idxCP10C)
               strContent = strContent & vbCrLf & "本所期限：" & ChangeWStringToWDateString(.TextMatrix(iRow, idxCP06))
               strContent = strContent & vbCrLf & "來函內容：請至卷宗區參看官方來函"
               strContent = strContent & vbCrLf & "修正承辦人：" & RsTemp("ProName")
               
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " values( '" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "')"
               cnnConnection.Execute strSql, intI
            End If
            
         End If
         cnnConnection.CommitTrans
         
On Error GoTo ErrHnd
        
        'Added by Morgan 2016/1/19
        '交工程師確認(專利用)
        If pIdx = 2 Then
                        
            .TextMatrix(iRow, 0) = ""
            .TextMatrix(iRow, idxBK) = "Y"
            .row = iRow
            'Modified by Morgan 2018/8/14
            'SetRowBK lngColor1
            SetRowBK cmdOK(2).BackColor
            'end 2018/8/14
            
            If iRow = iPrevRow Then SelectRow 0
            
            If bolMail And strTo <> "" Then
               If GetAttachFile(strCP09, "", m_AttachPath, True, stAttFileName) = True Then
                  PUB_SendMail strUserNum, strTo, "", strSub, strContent, , m_AttachPath & "\" & stAttFileName
               End If
            End If
        'end 2016/1/19
        
        'Add by Lydia 2015/01/20 分案(專利用)
        ElseIf pIdx = 3 Then
            
             If iRow = iPrevRow2 Then SelectRow2 0
            .TextMatrix(iRow, 0) = "X"
            .RowHeight(iRow) = 0
            lSelRows2 = lSelRows2 - 1
            lTotRows2 = lTotRows2 - 1
            lblCount2 = lSelRows2 & " / " & lTotRows2
            DoEvents
        'end 2015/01/20
        
        '判發,內部收文
        Else
            'Modified by Morgan 2019/1/9
            'If iRow = iPrevRow Then SelectRow 0
            '.TextMatrix(iRow, 0) = "X"
            '.RowHeight(iRow) = 0
            'lSelRows = lSelRows - 1
            'lTotRows = lTotRows - 1
            'lblCount = lSelRows & " / " & lTotRows
            'DoEvents
            UpdateGrid1 iRow
            'end 2019/1/9
        End If
      End If
   Next
   End With
   
   FormSave = True
   PUB_SendMailCache 'Added by Morgan 2014/10/23
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub SelectRow(pRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With MSHFlexGrid1
   nCol = .col
   If iPrevRow > 0 Then
      If iPrevRow <> pRow Then
         .row = iPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         
         If .TextMatrix(.row, idxBK) = "Y" Then
            SetRowBK lngColor1
         'Added by Morgan 2019/1/9
         ElseIf .TextMatrix(.row, idxLP37) <> "" Then
            SetRowBK Label5.BackColor
         'end 2019/1/9
         Else
            SetRowBK
         End If
      End If
   End If
   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      'Modified by Morgan 2016/3/11 固定欄位後1欄顯示底色(只有一筆時才能顯示已交工程師的狀態)
      For iCol = .FixedCols + 1 To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   iPrevRow = pRow
   SetLP37 pRow 'Added by Morgan 2019/1/9
   
   'Added by Morgan 2020/12/30
   If GetValue(iPrevRow, "IDS") = "Y" Then
      strExc(1) = GetValue(iPrevRow, "CP43")
      If ChkIDS(strExc(1)) = True Then
         cmdIDS.Visible = True
      End If
   Else
      cmdIDS.Visible = False
   End If
   'end 2020/12/30
   End With
End Sub

Private Sub SetCombo1()
   Dim ii As Integer
   Dim arrNum() As String
   Combo1.Clear
   
   'Add by Amy 2019/12/04 +商標
   If Left(Pub_StrUserSt03, 2) = "P2" Then
        Combo1.AddItem strUserNum & " " & strUserName, 0
   Else
      'Modified by Morgan 2015/4/21 改用權限控制,目前為游經理及王副總可執行
      'Modified by Morgan 2016/6/17 配合非臺灣案有其他判發人再改成可看自己及當時請假之被代理人
    
       
       'Added by Morgan 2016/9/8
       '王副總固定可看游經理的 --游經理
       'Modified by Morgan 2017/3/14
       '林柄佑也可看游經理的 --游經理
       'Modified by Morgan 2022/4/27
       '取消林柄佑改李柏翰 --游經理
       'If strUserNum = "71011" Or strUserNum = "82026" Then
'cancel by sonia 2024/4/23 移至下方
'       'modify by sonia 2024/4/9 游經理要求再加回林柄佑
'       If strUserNum = "71011" Or strUserNum = "99050" Or strUserNum = "82099" Then
'          Combo1.AddItem "73022 " & GetPrjSalesNM("73022")
'       End If
'end 2024/4/23
       'end 2016/9/8
       
       'Modified by Morgan 2018/7/27 CFP電子化(71011也有要判發的定稿,改都預設看自己的)
       If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
          Combo1.AddItem strUserNum & " " & strUserName, 0
       Else
          Combo1.AddItem strUserNum & " " & strUserName
       End If
       'end 2018/7/27
       
       'Added by Morgan 2017/12/25 郭固定可看玲玲的(原設案件職代但分信規則有衝突)
       'Added by Morgan 2018/11/13 +郭也可看王副總的(CFP核准函)
       If strUserNum = "79075" Then
          'Modified by Morgan 2022/5/20 改可以看所有P程序--郭
          'Combo1.AddItem "81002 " & GetPrjSalesNM("81002")
          strExc(0) = "select st01,st02 from staff where st03='P12' and st04='1' and st01<'F' and substr(st01,-2)<'9' and st01<>'" & strUserNum & "'"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
            Do While Not RsTemp.EOF
               Combo1.AddItem RsTemp("st01") & " " & RsTemp("st02")
               RsTemp.MoveNext
            Loop
          End If
          'end 2022/5/20
          'Combo1.AddItem "71011 " & GetPrjSalesNM("71011")   'cancel by sonia  2023/6/5 71011已退休
       'Added by Morgan 2019/6/24 CFP程序人員的判發選單都可選擇副總--郭
       ElseIf Pub_strUserST05 = "83" Or Pub_strUserST05 = "85" Then 'CFP
          'Combo1.AddItem "71011 " & GetPrjSalesNM("71011")   'cancel by sonia  2023/6/5 71011已退休
       End If
       
       'Added by Morgan 2022/5/20 余彥葶可以看到陳悅軒的公文(蕭茹曣判發) --郭
       If strUserNum = "A2023" Then
         Combo1.AddItem "A3014 " & GetPrjSalesNM("A3014")
       End If
       'end 2022/5/20
       'end 2017/12/25
       
       'Added by Morgan 2022/12/6
       '李柏翰的資料，王副總、郭雅娟、CFP程序都可以看--秀玲
       If strUserNum = "71011" Or strUserNum = "79075" Or Pub_strUserST05 = "83" Or Pub_strUserST05 = "85" Then
            Combo1.AddItem "99050 " & GetPrjSalesNM("99050")
       End If
       'end 2022/12/6
   End If
   
   If Pub_StrUserSt03 = "M51" Then
      Combo1.AddItem "      " & "全部", 0
   End If
   
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
   'add by sonia 2024/4/23 由上方移下來，否則82099會被上一行之Pub_SetForOthersEmpCombo模組取消73022
   'modify by sonia 2024/4/9 游經理要求再加回林柄佑  '2024/4/23 modify by sonia 82026改為82099
   'Modified by Morgan 2025/2/21 +P10
   'Modified by Morgan 2025/6/26 +79075
   If strUserNum = "99050" Or PUB_GetST03(strUserNum) = "P10" Or strUserNum = "79075" Then
      'Modified by Morgan 2025/2/21 73022->pub_PMan
      'Combo1.AddItem "73022 " & GetPrjSalesNM("73022")
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      arrNum() = Split(pub_PMan, ";")
      For ii = LBound(arrNum) To UBound(arrNum)
         If arrNum(ii) <> "" And arrNum(ii) <> strUserNum Then
            Combo1.AddItem arrNum(ii) & " " & GetPrjSalesNM(arrNum(ii))
         End If
      Next
      'end 2025/2/21
   End If
   'end 2024/4/23
   Combo1.ListIndex = 0
   
   'Add by Lydia 2015/01/20 判斷是否為游經理的權限
   bolMan = False
   For intI = 0 To Combo1.ListCount - 1
       'Modified by Morgan 2025/2/21
       'If InStr(Combo1.List(intI), "73022") > 0 Then
       pub_PMan = Pub_GetSpecMan("專利處特定編號")
       If InStr(pub_PMan, Left(Combo1.List(intI), 5)) > 0 Then
       'end 2025/2/21
          bolMan = True
          Exit For
       End If
   Next intI
   
'Combo1.AddItem "73022 游登銘"
Combo1.ListIndex = 0
'Combo1.Enabled = False
'bolMan = True
'end 2015/4/21
'end 2016/6/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm040113 = Nothing
End Sub

Private Sub lblAttCnt_Click()
   SendMessage cboAtt.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

Private Sub MSHFlexGrid1_DblClick()
   Dim nCol As Integer
   
   If MSHFlexGrid1.MouseRow > 0 Then
      nCol = MSHFlexGrid1.MouseCol
      
      fnOpen1
      
      'Added by Morgan 2016/7/21
      '開啟文件時自動勾選--郭雅娟,游經理
      If nCol <> 0 Then
         MSHFlexGrid1.col = 0
         If MSHFlexGrid1.Text = "" Then
            ClickGrid MSHFlexGrid1
         End If
      End If
      'end 2016/7/21
   End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stCP09 As String
      
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
   
'Removed by Morgan 2015/4/13 取消全選--游經理
'      If nCol = 0 Then
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) = "" Then
'               stValue = "V"
'               Exit For
'            '已刪除資料標示為 X
'            ElseIf .TextMatrix(iRow, 0) = "V" Then
'               stValue = ""
'               Exit For
'            End If
'         Next
'
'         lSelRows = 0
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) <> "X" Then
'               If .TextMatrix(iRow, 0) <> stValue Then
'                  .TextMatrix(iRow, 0) = stValue
'               End If
'            End If
'            If .TextMatrix(iRow, 0) = "V" Then
'               lSelRows = lSelRows + 1
'            End If
'         Next
'
'         lblCount = lSelRows & " / " & lTotRows
'      Else
         
         '紀錄前次點選的收文號
         If iPrevRow > 0 Then
            stCP09 = GetValue(iPrevRow, "cp09")
         End If
         
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
                  
         '重設排序後前次點選的位置
         If iPrevRow > 0 Then
            For iRow = 1 To .Rows - 1
               If stCP09 = GetValue(iRow, "cp09") Then
                  iPrevRow = iRow
                  Exit For
               End If
            Next
         End If
         
'      End If'Removed by Morgan 2015/4/13 取消全選--游經理
      
   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid MSHFlexGrid1
      End If
      SelectRow nRow
   End If
   
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(FlexGrid As MSHFlexGrid)
   With FlexGrid
   If .Text = "V" Then
      lSelRows = lSelRows - 1
      .Text = ""
   '已刪除資料標示為 X
   ElseIf .Text = "" Then
      lSelRows = lSelRows + 1
      .Text = "V"
   End If
   lblCount = lSelRows & " / " & lTotRows
   End With
End Sub
'Modified by Morgan 2019/1/123 +pFileDescs
'Modified by Morgan 2021/11/29 +pCP10,pCP43
Private Function GetAttachFile(ByVal strCPP01 As String, Optional ByRef pFileName As String, Optional ByRef pSavePath As String, Optional pJoinDoc As Boolean, Optional pSaveName As String, Optional pFileDescs As String, Optional pCP10 As String, Optional pCP43 As String) As Boolean
   Dim stAttPath As String
   Dim stTempFile As String
   Dim stSort As String
   
'Removed by Morgan 2015/3/24
'   Dim lngSize As Long
'   Dim iFileNo As Integer
'   Dim bytes() As Byte

On Error GoTo ErrHnd
   
   '建立暫存資料夾
   If Dir(pSavePath, vbDirectory) = "" Then
      MkDir pSavePath
   End If
   
   If pSaveName <> "" Then
      stTempFile = pSaveName
   Else
      stTempFile = "$" & strCPP01 & ".pdf"
      If pJoinDoc = True Then
         If Dir(pSavePath & "\" & stTempFile) <> "" Then
            pFileName = stTempFile
            GetAttachFile = True
         End If
      End If
   End If
   
   If GetAttachFile = False Then
      
      'Modified by Morgan 2015/3/23 上傳檔案改呼叫共用函數(要改為FTP方式)
      'Modified by Morgan 2019/1/23 +FDesc
      'Modified by Morgan 2019/6/27 +Sort(原來只控制客戶函先)
      stSort = PUB_GetAttSort("P")
      strExc(0) = "select cpp01,cpp02,'('||Round(cpp03 / 1024, 2)||' KB) '||GETFILEDESC(cpp02,CP01,CP10) as FDesc," & stSort & " Sort" & _
         " from caseprogress,casepaperpdf" & _
         " where cp09='" & strCPP01 & "' and cpp01(+)=cp09" & IIf(pFileName <> "", " and upper(cpp02)=upper('" & ChgSQL(pFileName) & "')", "") & _
         " and CPP10<>'D' and upper(substr(cpp02,-4))='.PDF'"
      
      'Added by Morgan 2021/11/29 轉公文1998還要抓相關收文號檔案
      If pCP10 = "1998" And pCP43 <> "" Then
         strExc(0) = strExc(0) & " union select cpp01,cpp02,'('||Round(cpp03 / 1024, 2)||' KB) '||GETFILEDESC(cpp02,CP01,CP10) as FDesc," & stSort & " Sort" & _
            " from caseprogress,casepaperpdf" & _
            " where cp09='" & pCP43 & "' and cpp01(+)=cp09" & IIf(pFileName <> "", " and upper(cpp02)=upper('" & ChgSQL(pFileName) & "')", "") & _
            " and CPP10<>'D' and upper(substr(cpp02,-4))='.PDF'"
      End If
      strExc(0) = strExc(0) & " order by Sort"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            'Modified by Morgan 2015/6/18 檔名含空白無法合併
            'stAttPath = pSavePath & "\" & .Fields("cpp02")
            stAttPath = pSavePath & "\" & Replace(.Fields("cpp02"), " ", "_")
            'end 2015/6/18
            'If Dir(stAttPath) = "" Then '檔案存在時不必再下載 'Removed by Morgan 2015/12/23 還是要下載,舉發案不同收文號檔名會重複
            
'Modified by Morgan 2015/3/24 上傳檔案改呼叫共用函數(要改為FTP方式)
'               lngSize = Val(.Fields("cpp03").Value)
'               ReDim bytes(lngSize)
'               If lngSize > 0 Then
'                  bytes() = .Fields("cpp04").GetChunk(lngSize)
'               End If
'
'               iFileNo = FreeFile
'               Open stAttPath For Binary Access Write As #iFileNo
'               If lngSize > 0 Then Put #iFileNo, , bytes()
'               Close #iFileNo
               'Modified by Morgan 2015/6/18 檔名含空白無法合併
               'If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), pSavePath) = False Then
               If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), stAttPath, True) = False Then
               'end 2015/6/18
                  Exit Function
               End If
'end 2015/3/24
               
            'End If
            'Modified by Morgan 2015/6/18 檔名含空白無法合併
            'pFileName = IIf(pFileName = "", "", pFileName & ";") & .Fields("cpp02")
            pFileName = IIf(pFileName = "", "", pFileName & ";") & Replace(.Fields("cpp02"), " ", "_")
            'end 2015/6/18
            
            pFileDescs = pFileDescs & Replace(.Fields("cpp02"), " ", "_") & Chr(9) & .Fields("FDesc") & ";"
            .MoveNext
         Loop
         End With
         GetAttachFile = True
      End If
   End If
   If InStr(pFileName, ";") > 0 Then
      '開啟也要合併，否則開啟後再預覽時檔案會被開啟中的鎖住
      If Dir(pSavePath & "\" & stTempFile) = "" Then
         If JoinPdf(pFileName, stTempFile, pSavePath) = True Then
            If pJoinDoc = True Then
               pFileName = stTempFile
            End If
         End If
      End If
   End If
   Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
'Removed by Morgan 2015/3/24
'   If iFileNo > 0 Then Close #iFileNo
End Function

'pdf合併(檔名以分號區隔),目的預設與來源相同
Private Function JoinPdf(pFromFiles As String, pToFileName As String, pFromPath As String) As Boolean
   Dim strCmd As String, iTimes As Integer
   Dim arrFiles() As String
   Dim stTempName As String, idx As Integer
   Dim stNewFiles As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim bolRetry As Boolean
   Dim bolPFeeForm As Boolean
   
On Error GoTo ErrHnd
   
   '切換至來源目錄
      If pFromPath <> "." Then ChDir pFromPath
      
'Modified by Morgan 2014/5/6 卷宗區的檔案不會有中文
'   '多檔先合併
'   If InStr(pFromFiles, ";") > 0 Then
'      '中文檔名無法合併要將檔案依順序重新命名為 pToFileName.001, pToFileName.002, pToFileName.003..
'      arrFiles = Split(pFromFiles, ";")
'      stNewFiles = ""
'      For idx = LBound(arrFiles) To UBound(arrFiles)
'         If oFileSys.FileExists(arrFiles(idx)) = True Then
'            Set oFile = oFileSys.GetFile(arrFiles(idx))
'            stTempName = pToFileName & "." & Format(idx + 1, "000") '序號需索引一致後面更名才會對到
'            oFile.Name = stTempName
'            stNewFiles = IIf(stNewFiles <> "", stNewFiles & ";", "") & stTempName
'         Else
'            MsgBox "找不到pdf附件(" & arrFiles(idx) & ")，作業取消！"
'            GoTo ErrHnd
'         End If
'      Next
'   Else
      stNewFiles = pFromFiles
'   End If
   
   '1個檔案用更名
   If InStr(stNewFiles, ";") = 0 Then
      Set oFile = oFileSys.GetFile(stNewFiles)
      bolRetry = False
      'Modified by Morgan 2019/3/4 改用複製(考慮EMail檢查附件及下拉選單問題)
      'oFile.Name = pToFileName
      oFile.Copy pToFileName
   '合併
   Else
      stNewFiles = Replace(stNewFiles, ";", " ")
      '刪舊檔
      If Dir(".\" & pToFileName) <> "" Then Kill ".\" & pToFileName
      
      '特殊路徑無法存檔故先就地合併
      'Modified by Morgan 2014/5/9 合併程式改放執行檔路徑
      'strCmd = "pdftk.exe " & stNewFiles & " cat output .\" & pToFileName
      strCmd = pub_PdftkEXE & " " & stNewFiles & " cat output .\" & pToFileName
      process_id = SHELL(strCmd, vbHide)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
      If process_handle <> 0 Then
         For intI = 1 To 10
            If PUB_CheckIsRunning(pub_PdftkName) = True Then
               Sleep 1000
            Else
               Exit For
            End If
         Next
         If intI > 10 Then
            TerminateProcess process_handle, 0&
            CloseHandle process_handle
            GoTo ExitPoint
         Else
            CloseHandle process_handle
         End If
      Else
         GoTo ExitPoint
      End If
   End If
      
   JoinPdf = True
   
'Modified by Morgan 2014/5/6 卷宗區的檔案不會有中文
'   '檔名還原(直接開啟要用)
'   For idx = LBound(arrFiles) To UBound(arrFiles)
'      stTempName = pToFileName & "." & Format(idx + 1, "000")
'      If oFileSys.FileExists(stTempName) = True Then
'         Set oFile = oFileSys.GetFile(stTempName)
'         bolRetry = False
'         oFile.Name = arrFiles(idx)
'      End If
'   Next
   
ErrHnd:
   If Err.Number <> 0 Then
      '檔案可能不會馬上釋放,會無法更名
      If Err.Number = 70 And bolRetry = False Then
         Sleep 1000
         bolRetry = True
         Resume
      Else
         MsgBox Err.Description, vbCritical
      End If
   End If
   
ExitPoint:
   
   ChDir App.path '目錄切回
End Function

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub
'Add by Lydia 2015/01/20輸入舉發,舉發答辯,准,駁產生內部收文分析(B類941),若原承辦工程師已離職，請將分析這道掛至游經理確認內部收文系統，並由游經理更改承辦人後發ｅ－ｍａｉｌ通知承辦人
Private Sub SetMouseBusy2()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid2.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady2()
   Screen.MousePointer = vbDefault
   MSHFlexGrid2.MousePointer = vbDefault
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stCP09 As String
      
   With MSHFlexGrid2
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
   
'Removed by Morgan 2015/4/13 取消全選--游經理
'      If nCol = 0 Then
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) = "" Then
'               stValue = "V"
'               Exit For
'            '已刪除資料標示為 X
'            ElseIf .TextMatrix(iRow, 0) = "V" Then
'               stValue = ""
'               Exit For
'            End If
'         Next
'
'         lSelRows2 = 0
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) <> "X" Then
'               If .TextMatrix(iRow, 0) <> stValue Then
'                  .TextMatrix(iRow, 0) = stValue
'               End If
'            End If
'            If .TextMatrix(iRow, 0) = "V" Then
'               lSelRows2 = lSelRows2 + 1
'            End If
'         Next
'
'         lblCount2 = lSelRows2 & " / " & lTotRows2
'      Else
         
         '紀錄前次點選的收文號
         If iPrevRow2 > 0 Then
            stCP09 = GetValue(iPrevRow2, "cp09")
         End If
         
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
                  
         '重設排序後前次點選的位置
         If iPrevRow2 > 0 Then
            For iRow = 1 To .Rows - 1
               If stCP09 = GetValue(iRow, "cp09") Then
                  iPrevRow2 = iRow
                  Exit For
               End If
            Next
         End If
         
'      End If'Removed by Morgan 2015/4/13 取消全選--游經理
      
   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid2 MSHFlexGrid2
      End If
      SelectRow2 nRow
   End If
   
   .Visible = True
   End With
End Sub

Private Sub SetGrid2(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 1140, 800, 1000, 825)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid2
   If pReset = True Then
      .Clear
      .Rows = 2
      
      iPrevRow2 = 0
      lTotRows2 = 0
      lSelRows2 = 0
      lblCount2 = lSelRows2 & " / " & lTotRows2
   End If
   .FixedCols = 2
   .FormatString = "V|本所案號|案件性質|案件名稱|收文日"
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

Private Sub QueryData2()
   Dim iRow As Integer
   Dim stCon As String
   
   stCon = ChangeTStringToWString(PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -7))
   SetGrid2 True
   'Modified by Lydia 2015/04/24 輸入舉發,舉發答辯,准,駁產生內部收文分析,若原承辦工程師已離職或承辦人掛為程序人員也請將分析這道掛至游經理確認內部收文系統，並由游經理更改承辦人後發ｅ－ｍａｉｌ通知承辦人。
   '+ST03='P12'
   'Modified by Morgan 2015/6/22 +73017換部門,加判斷非專利處 substr(st03,1,2)<>'P1'
   strExc(0) = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,cpm03 案件性質,pa05 案件名稱,sqldatet(cp05) 收文日," & _
               "cp09,cp10,cp14,'' Bcp09,'' Bcp14,cp01,cp02,cp03,cp04,cp06,cu04,st02 " & _
               "From caseprogress, patent,casepropertymap ,staff,customer " & _
               "where cp05>='" & stCon & "' and cp01='P' and cp10 ='941' and substr(cp09,1,1)='B' " & _
               "and (st04='2' or cp14='71011' or st03='P12' or substr(st03,1,2)<>'P1') and cp27||cp57 is null and pa01(+)=cp01 and pa02(+)=cp02 " & _
               "and pa03(+)=cp03 and pa04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and cp14=st01(+) " & _
               "and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) order by cp05,cp09 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid2
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      lTotRows2 = RsTemp.RecordCount
      lblCount2 = lSelRows2 & " / " & lTotRows2
      SetGrid2
      .col = 1: .row = 1
      SelectRow2 1
      .Visible = True

      End With
   End If
   'WebBrowser1.Navigate "about:blank" 'Removed by Morgan 2019/1/23
End Sub

Private Sub SelectRow2(pRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With MSHFlexGrid2
   nCol = .col
   If iPrevRow2 > 0 Then
      If iPrevRow2 <> pRow Then
         .row = iPrevRow2
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
   iPrevRow2 = pRow
   End With
End Sub

Private Sub ClickGrid2(FlexGrid As MSHFlexGrid)
   With FlexGrid
   If .Text = "V" Then
      lSelRows2 = lSelRows2 - 1
      .Text = ""
   '已刪除資料標示為 X
   ElseIf .Text = "" Then
      lSelRows2 = lSelRows2 + 1
      .Text = "V"
   End If
   lblCount2 = lSelRows2 & " / " & lTotRows2
   End With
End Sub
Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      PubShowNextData iPrevRow, MSHFlexGrid1
   ElseIf Index = 1 Then
      PubShowNextData iPrevRow2, MSHFlexGrid2
   'Added by Morgan 2018/12/20
   ElseIf Index = 5 Then
      PubShowNextData iPrevRow, MSHFlexGrid1, Index
   'end 2018/12/20
   End If
End Sub
'end 2015/01/20

Private Sub ShowEngForm(rsQuery As ADODB.Recordset, pRecNo As String, pSetEng As Boolean)
   Dim strTmp As String, strSql As String, intQ As Integer
   
   With frm040113_1
   If Left(pRecNo, 1) = "C" Then
      .Caption = "C類來函分案"
   End If
   .lblRecNo = pRecNo
   .lblCaseNo = rsQuery("CaseNo")
   .lblAppName = rsQuery("cu04")
   .lblCaseName = rsQuery("pa05")
   .lblProperty = rsQuery("cpm03")
   .lblOurDeadLine = ChangeWStringToTDateString("" & rsQuery("cp06"))
   strTmp = rsQuery("cp14") & " " & rsQuery("st02")
   strSql = "select st01||' '||st02 from staff where st01<>'71011' and st03 in ('P10','P11') and st04='1'"
   'Added by Morgan 2022/11/16 82026換部門
   strSql = strSql & " union select st01||' '||st02 from staff where st01='" & rsQuery("cp14") & "' and st04='1'"
   'end 2022/11/16
   strSql = strSql & " order by 1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      Do While Not rsQuery.EOF
         .cboCP14.AddItem rsQuery(0)
         rsQuery.MoveNext
      Loop
   End If
   .cboCP14 = strTmp
   End With
   SetMouseReady
   If pSetEng Then
      frm040113_1.mInputKey = "1"
   End If
   frm040113_1.Show vbModal
End Sub

'Added by Morgan 2019/1/9
Public Sub UpdateGrid1(pRow As Integer)
   If pRow = iPrevRow Then SelectRow 0
   With MSHFlexGrid1
   .TextMatrix(pRow, 0) = "X"
   .RowHeight(pRow) = 0
   End With
   lSelRows = lSelRows - 1
   lTotRows = lTotRows - 1
   lblCount = lSelRows & " / " & lTotRows
   DoEvents
End Sub

'Added by Morgan 2019/1/9
'帶出退回意見
Private Sub SetLP37(Optional pRow As Integer)
   Dim iCol As Integer
   
   Frame3.Visible = False
   With MSHFlexGrid1
   If pRow = 0 Then
      txtLP37 = ""
   Else
      txtLP37 = .TextMatrix(pRow, idxLP37)
   End If
   If txtLP37 <> "" Then
      .Height = Frame1.Top - .Top - Frame3.Height
      Frame3.Top = .Top + .Height + 50
      Frame3.Visible = True
   Else
      .Height = Frame1.Top - .Top + 50
   End If
   End With
End Sub

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

Private Function ChkIDS(pCP09 As String, Optional pComfirm As Boolean) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from IDSList where il01='" & pCP09 & "'" & IIf(pComfirm = True, " and il08 is null", "")
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ <> 0 Then
      ChkIDS = True
   End If
End Function

'Added by Morgan 2023/6/26
'力成(X82532)案件都要通知客戶不可內部收文
Private Function ChkXCust(pCP09 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rstQ As ADODB.Recordset
   
   stSQL = "select pa26,cu04 from caseprogress,patent,customer" & _
      " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and substr(pa26,1,6)='X82532'" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   intQ = 1
   Set rstQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      MsgBox rstQ("cu04") & "(" & rstQ("pa26") & ")案件，OA一律要通知，即使OA是單純的通知文字誤繕修正，也要先直寄OA，載明不收費即可。", vbExclamation, "不可內部收文!!"
      ChkXCust = True
   End If
End Function
