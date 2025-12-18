VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_2_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護_申請書"
   ClientHeight    =   5040
   ClientLeft      =   3980
   ClientTop       =   2270
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8190
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm090201_2_5.frx":0000
      Left            =   1110
      List            =   "frm090201_2_5.frx":0002
      TabIndex        =   57
      Text            =   "Combo1"
      Top             =   90
      Width           =   2385
   End
   Begin VB.Frame FrameFee 
      Height          =   1245
      Left            =   1140
      TabIndex        =   36
      Top             =   5250
      Width           =   6345
      Begin VB.TextBox txtCP137 
         Height          =   270
         Left            =   1890
         TabIndex        =   46
         Top             =   390
         Width           =   420
      End
      Begin VB.TextBox txtDecreaseFee 
         Height          =   270
         Left            =   4080
         TabIndex        =   45
         Top             =   930
         Width           =   840
      End
      Begin VB.TextBox txtAddFee 
         Height          =   270
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   930
         Width           =   840
      End
      Begin VB.TextBox txtCount 
         Height          =   270
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   660
         Width           =   420
      End
      Begin VB.TextBox txtCP136 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5100
         TabIndex        =   42
         Top             =   420
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtCP138 
         Height          =   270
         Left            =   3660
         TabIndex        =   41
         Top             =   390
         Width           =   420
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   40
         Top             =   108
         Width           =   420
      End
      Begin VB.TextBox txtCP135 
         Height          =   270
         Left            =   5100
         TabIndex        =   39
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtAddItem 
         Height          =   270
         Left            =   3660
         TabIndex        =   38
         Top             =   120
         Width           =   420
      End
      Begin VB.TextBox txtCP135_tmp 
         Height          =   270
         Left            =   5820
         TabIndex        =   37
         Top             =   930
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪除已審項數"
         ForeColor       =   &H00000040&
         Height          =   180
         Left            =   2580
         TabIndex        =   55
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次應退還規費"
         Height          =   180
         Left            =   2790
         TabIndex        =   54
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次應加收規費"
         Height          =   180
         Left            =   600
         TabIndex        =   53
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label13 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "修正後之請求項總項數"
         Height          =   180
         Left            =   90
         TabIndex        =   52
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪除未審項數"
         ForeColor       =   &H00000040&
         Height          =   180
         Left            =   810
         TabIndex        =   51
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblAddItem 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "新增請求項項數"
         Height          =   180
         Left            =   2400
         TabIndex        =   50
         Top             =   144
         Width           =   1260
      End
      Begin VB.Label lblPage 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "總頁數"
         Height          =   180
         Left            =   4560
         TabIndex        =   49
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "原請求項項數"
         Height          =   180
         Left            =   825
         TabIndex        =   48
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label LabelCP136 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "計算項數"
         Height          =   180
         Left            =   4380
         TabIndex        =   47
         Top             =   465
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   5010
      MaxLength       =   11
      TabIndex        =   31
      Top             =   705
      Width           =   1260
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   3030
      MaxLength       =   15
      TabIndex        =   30
      Text            =   "一(二)"
      Top             =   705
      Width           =   1485
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1890
      MaxLength       =   7
      TabIndex        =   29
      Top             =   435
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請書(&A)"
      Height          =   345
      Index           =   1
      Left            =   5640
      TabIndex        =   28
      Top             =   60
      Width           =   1125
   End
   Begin VB.Frame Frame204 
      Appearance      =   0  '平面
      Caption         =   "附送書件"
      ForeColor       =   &H00C00000&
      Height          =   3945
      Left            =   60
      TabIndex        =   1
      Top             =   1050
      Width           =   8055
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   5100
         TabIndex        =   27
         Text            =   "新型專利技術報告意見說明書"
         Top             =   1170
         Width           =   2835
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "申復書"
         Height          =   195
         Index           =   20
         Left            =   3240
         TabIndex        =   26
         Tag             =   ".EX.pdf"
         Top             =   780
         Width           =   3375
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "委任書"
         Height          =   195
         Index           =   19
         Left            =   3240
         TabIndex        =   25
         Tag             =   ".POA.pdf"
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "基本資料表"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   24
         Tag             =   ".CONTACT.pdf"
         Top             =   300
         Width           =   1305
      End
      Begin VB.CheckBox chkAtt2 
         Caption         =   "文件描述"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4095
         TabIndex        =   23
         Top             =   1230
         Width           =   1110
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "其他"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   22
         Top             =   1230
         Width           =   690
      End
      Begin VB.CheckBox chkAtt2 
         Caption         =   "文件檔名"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4095
         TabIndex        =   21
         Tag             =   ".ATT.pdf"
         Top             =   1470
         Width           =   1110
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之新型摘要"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   20
         Tag             =   "2.fix_ABSTRACT_u.pdf"
         Top             =   3225
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之新型圖式"
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   19
         Tag             =   "2.FIG.fix.pdf"
         Top             =   3015
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之設計圖式"
         Height          =   195
         Index           =   17
         Left            =   3960
         TabIndex        =   18
         Tag             =   "3.FIG.fix.pdf"
         Top             =   2610
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之新型申請專利範圍"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   17
         Tag             =   "2.fix_CLAIMS_u.pdf"
         Top             =   3675
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之新型說明書"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   14
         Left            =   150
         TabIndex        =   16
         Tag             =   "2.fix_DESCRIPTION_u.pdf"
         Top             =   3450
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之設計說明書"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   18
         Left            =   3960
         TabIndex        =   15
         Tag             =   "3.fix_u.pdf"
         Top             =   2820
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之新型申請專利範圍"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   14
         Tag             =   "2.FIX_CLAIMS.pdf"
         Top             =   2790
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之序列表"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Tag             =   "1.FIX.SEQ.pdf"
         Top             =   765
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之新型摘要"
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   12
         Tag             =   "2.FIX_ABSTRACT.pdf"
         Top             =   2355
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之新型說明書"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   11
         Tag             =   "2.FIX_DESCRIPTION.pdf"
         Top             =   2580
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之發明摘要"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   10
         Tag             =   "1.fix_ABSTRACT_u.pdf"
         Top             =   1410
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之發明說明書"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   9
         Tag             =   "1.fix_DESCRIPTION_u.pdf"
         Top             =   1635
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之序列表"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Tag             =   "1.fix_SEQ_u.pdf"
         Top             =   1860
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正部分劃線之發明申請專利範圍"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   7
         Tag             =   "1.fix_CLAIMS_u.pdf"
         Top             =   2070
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之發明摘要"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Tag             =   "1.fix_ABSTRACT.pdf"
         Top             =   330
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之發明說明書"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Tag             =   "1.fix_DESCRIPTION.pdf"
         Top             =   540
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之發明申請專利範圍"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Tag             =   "1.fix_CLAIMS.pdf"
         Top             =   975
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之發明圖式"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   3
         Tag             =   "1.FIG.fix.pdf"
         Top             =   1200
         Width           =   3045
      End
      Begin VB.CheckBox chk1Tab1 
         Caption         =   "修正後之設計說明書"
         Height          =   195
         Index           =   16
         Left            =   3960
         TabIndex        =   2
         Tag             =   "3.fix.pdf"
         Top             =   2385
         Width           =   3045
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         X1              =   90
         X2              =   3270
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         X1              =   3930
         X2              =   7110
         Y1              =   2310
         Y2              =   2310
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(U)"
      Height          =   345
      Index           =   0
      Left            =   6870
      TabIndex        =   0
      Top             =   60
      Width           =   1125
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   6660
      TabIndex        =   58
      Top             =   690
      Visible         =   0   'False
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別:"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   150
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發文字號：（　　）智專                                     字第                               號"
      Height          =   180
      Left            =   945
      TabIndex        =   35
      Top             =   750
      Width           =   5580
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Left            =   945
      TabIndex        =   34
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "辦理依據:"
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "(無發文日期, 辦理依據整行不顯示)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2940
      TabIndex        =   32
      Top             =   480
      Width           =   2730
   End
End
Attribute VB_Name = "frm090201_2_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sindy 2021/5/10 Form2.0已修改; lstNameAgent
'Create By Sindy 2019/8/13
Option Explicit

Public cmdState As Integer '紀錄作用按鍵
Dim pa() As String, cp() As String
Public m_CaseNo As String
Public m_CP09 As String
Dim bolHad404 As Boolean, strHad404CP09 As String
Dim m_IPOSendDt As String, m_IPOSendData1 As String, m_IPOSendData2 As String
Dim m_AgentName As String 'Add By Sindy 2021/5/10


Private Sub chk1Tab1_Click(Index As Integer)
Dim iChecked As Single
   
   If Val(chk1Tab1(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   Select Case Index
   '摘要
   Case 0, 5, 9, 13:
      If chk1Tab1(0).Enabled = True Then chk1Tab1(0).Value = iChecked
      If chk1Tab1(5).Enabled = True Then chk1Tab1(5).Value = iChecked
      If chk1Tab1(9).Enabled = True Then chk1Tab1(9).Value = iChecked
      If chk1Tab1(13).Enabled = True Then chk1Tab1(13).Value = iChecked
   '說明書
   Case 1, 6, 10, 14, 16, 18:
      If chk1Tab1(1).Enabled = True Then chk1Tab1(1).Value = iChecked
      If chk1Tab1(6).Enabled = True Then chk1Tab1(6).Value = iChecked
      If chk1Tab1(10).Enabled = True Then chk1Tab1(10).Value = iChecked
      If chk1Tab1(14).Enabled = True Then chk1Tab1(14).Value = iChecked
      If chk1Tab1(16).Enabled = True Then chk1Tab1(16).Value = iChecked
      If chk1Tab1(18).Enabled = True Then chk1Tab1(18).Value = iChecked
   '發明序列表
   Case 2, 7:
      If chk1Tab1(2).Enabled = True Then chk1Tab1(2).Value = iChecked
      If chk1Tab1(7).Enabled = True Then chk1Tab1(7).Value = iChecked
   '專利範圍
   Case 3, 8, 11, 15:
      If chk1Tab1(3).Enabled = True Then chk1Tab1(3).Value = iChecked
      If chk1Tab1(8).Enabled = True Then chk1Tab1(8).Value = iChecked
      If chk1Tab1(11).Enabled = True Then chk1Tab1(11).Value = iChecked
      If chk1Tab1(15).Enabled = True Then chk1Tab1(15).Value = iChecked
   End Select
End Sub

Private Sub chkAtt_Click(Index As Integer)
   If chkAtt(2).Value = 1 Then
      chkAtt2(0).Value = 1
      chkAtt2(1).Value = 1
   Else
      chkAtt2(0).Value = 0
      chkAtt2(1).Value = 0
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
Dim m_bolShowEng As Boolean
Dim stContent As String
Dim ET03 As String
Dim ii As Integer
Dim Cancel As Boolean

Select Case cmdState
Case 0 '回前畫面
'   If UCase(m_PrevForm.Name) = UCase("frm090202_2") Then
'      frm090202_2.Show
'   Else
'      frm090201_2.Show
'   End If
   Unload Me
Case 1 '申請書
   
'   Cancel = False
'   lstNameAgent_Validate Cancel
'   If Cancel = True Then
'      lstNameAgent.SetFocus
'      Exit Sub
'   End If
   
   Screen.MousePointer = vbHourglass
   
   '出名代理人
   If cp(110) = "" Then
      '台灣專利案若以專利商標出名則提醒
      'Modified by Morgan 2020/3/11
      'If pa(1) = "P" And pa(9) = "000" And pa(161) = "Y" And InStr(NewCasePtyList, cp(10)) > 0 Then
'Removed by Morgan 2020/3/13 取消--郭
'      If pa(1) = "P" And pa(9) = "000" And pa(161) = "T" And InStr(NewCasePtyList, cp(10)) > 0 Then
'         cp(110) = "94007"
'      End If
'end 2020/3/13
      '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
      lstNameAgent.Clear
      If pa(9) = "000" And Len(cp(10)) <> 4 Then
         PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True
         'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
         lstNameAgent.Height = 1100
         lstNameAgent.Width = 1300
      End If
   End If
   
   'If cp(10) = "203" Or cp(10) = "204" Then
   If Left(Trim(Combo1.Text), 2) = "01" Then
      ET03 = "15" '修正申請書
   Else
      ET03 = "20" '申復申請書
   End If
   
   If chkAtt(0).Value = 1 Then
      '申請書
      StartLetter2 "01", ET03
      NowPrint cp(9), "01", ET03, False, strUserNum, , , True, stContent
      '基本資料
      StartLetterPA_EData "01", "14", cp(9), pa, cp, True, True, , , m_bolShowEng
      NowPrint cp(9), "01", "14", True, strUserNum, , stContent, , , , , True, , , , , , , , True
   Else
      '申請書
      StartLetter2 "01", ET03
      NowPrint cp(9), "01", ET03, True, strUserNum, , , , , , , True, , , , , , , , True
   End If
   
   Screen.MousePointer = vbDefault
   
   Unload Me
   MsgBox "資料已產生完畢!!!"
Case Else
End Select
End Sub

Private Sub Form_Load()
Dim chk As CheckBox
   
   MoveFormToCenter Me
   
   ReDim pa(1 To TF_PA) As String
   ReDim cp(TF_CP)
   '專利基本檔
   pa(1) = SystemNumber(m_CaseNo, 1)
   pa(2) = SystemNumber(m_CaseNo, 2)
   pa(3) = SystemNumber(m_CaseNo, 3)
   pa(4) = SystemNumber(m_CaseNo, 4)
   Call ClsPDReadPatentDatabase(pa(), 國內)
   m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
   '進度檔
   cp(9) = m_CP09
   Call PUB_ReadCaseProgressDatabase(cp(), 國內)
   
   '來函文號:帶相關的總收文號,若已收到延期受理的函請帶延期受理的來文字號
   bolHad404 = False
   strExc(0) = "select cp09,cp118,cp10,cp158,cp05 From caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='404' and cp43='" & cp(9) & "'" & _
               " Union" & _
               " select cp09,cp118,cp10,cp158,cp05 From caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='404' and cp43 in(select cp43 from caseprogress where cp09='" & cp(9) & "')" & _
               " order by cp158 desc,cp05 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         '已延過期且為紙本送件,則走未延過期的申請書流程
         If "" & RsTemp.Fields("cp118") <> "" Then
            bolHad404 = True
            strHad404CP09 = RsTemp.Fields(0)
            strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument" & _
                        " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                        " AND ed11(+)=cp09 AND cp43='" & strHad404CP09 & "' and cp10='1004'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp("ED08")) Then
                  m_IPOSendDt = RsTemp("ED08") - 19110000
                  If Trim("" & RsTemp("cp08")) <> "" Then
                     strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                     m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                     strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                     m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
                  End If
               End If
            End If
         End If
      End If
   End If
   If m_IPOSendDt = "" Then
      strExc(0) = "SELECT cp08,ed08,cp09 FROM caseprogress,edocument,(SELECT cp43 FROM caseprogress" & _
                  " where CP09='" & cp(9) & "' AND cp43 IS NOT NULL) A" & _
                  " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " AND CP09=A.cp43 AND ed11(+)=A.cp43" & _
                  " ORDER BY CP05 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            m_IPOSendDt = RsTemp("ED08") - 19110000
            If Trim("" & RsTemp("cp08")) <> "" Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
               m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
            End If
         ElseIf Trim("" & RsTemp("cp08")) <> "" Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
               m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
         End If
      End If
      If m_IPOSendDt = "" Then
         strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp09 ORDER BY CP05 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp("ED08")) Then
               m_IPOSendDt = RsTemp("ED08") - 19110000
               If Trim("" & RsTemp("cp08")) <> "" Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
               End If
            ElseIf Trim("" & RsTemp("cp08")) <> "" Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
            End If
         End If
      End If
   End If
   Text10.Text = m_IPOSendDt
   If m_IPOSendDt <> "" Then
      Label6 = Replace(Label6.Caption, "（　　）", "（ " & Left(Text10, 3) & " ）")
   End If
   Text7.Text = m_IPOSendData1
   Text8.Text = m_IPOSendData2
      
   Combo1.Clear
   Combo1.AddItem "01 修正申請書"
   'Modify By Sindy 2022/11/10 玲玲說再審查不要出現申復申請書
   If cp(10) <> "107" Then
   '2022/11/10 END
      Combo1.AddItem "02 申復申請書"
   End If
   If cp(10) = "203" Or cp(10) = "204" Then
      Combo1.ListIndex = 0
   ElseIf cp(10) = "205" Then
      Combo1.ListIndex = 1
      chk1Tab1(20).Value = 1
   End If
   
   'Add By Sindy 2019/8/29 為延期再審
   If cp(10) = "107" Then
      chk1Tab1(20).Caption = "再審查理由書"
      chk1Tab1(20).Tag = ".RE.pdf"
   End If
   '2019/8/29 END
   
   For Each chk In chk1Tab1
      If Left(chk.Tag, 1) = pa(8) Or Left(chk.Tag, 1) = "." Then
         chk.Enabled = True
      Else
         chk.Enabled = False
      End If
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_2_5 = Nothing
End Sub

'申請書
'Optional ByVal bolAttachments As Boolean = True 是否含附送書件
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String, _
   Optional ByVal bolAttachments As Boolean = True, Optional ByRef m_bolShowEng As Boolean = False) As Boolean

Dim strTxt(110) As String, strTmp As String, strTmp1 As String, strTmp2 As String
Dim ii As Integer, jj As Integer
Dim chk As CheckBox
   
   ii = 0
   EndLetter ET01, cp(9), ET03, strUserNum
   
   '本所案號
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '辦理依據
   If m_IPOSendDt <> "" Or m_IPOSendData1 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(m_IPOSendDt) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','智專字','" & m_IPOSendData1 & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','發文號','" & m_IPOSendData2 & "')"
   End If
   
   If pa(26) <> "" Then
      If GetPrjNationNumber1(pa(26)) > "010" Then m_bolShowEng = True
   End If
   If pa(27) <> "" Then
      If GetPrjNationNumber1(pa(27)) > "010" Then m_bolShowEng = True
   End If
   If pa(28) <> "" Then
      If GetPrjNationNumber1(pa(28)) > "010" Then m_bolShowEng = True
   End If
   If pa(29) <> "" Then
      If GetPrjNationNumber1(pa(29)) > "010" Then m_bolShowEng = True
   End If
   If pa(30) <> "" Then
      If GetPrjNationNumber1(pa(30)) > "010" Then m_bolShowEng = True
   End If
   '顯示英文
   If m_bolShowEng = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','顯示英文','♀')"
   End If
   
   '申請人
   Call PUB_GetApplPA_EData(ET01, ET03, cp(9), pa(), , , , m_bolShowEng)
   
   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03" '取消order by OA03:依存入的順序
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/8 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, cp(9), ET03, ii, strTxt())
   
   If (Val(txtAddItem) = 0 And Val(txtCP137) = 0) Or _
      (Val(txtAddFee) = 0 And Val(txtDecreaseFee) = 0) Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','規費不變','♀')"
   Else
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','規費有變動','♀')"
      '規費變動
      If Val(txtAddItem) > 0 Or Val(txtCP137) > 0 Or Val(txtCP138) > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','原請求項數','" & Val(txtItem) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','新增項數','" & Val(txtAddItem) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','刪除項數','" & IIf(Val(txtCP137) + Val(txtCP138) = 0, 0, Val(txtCP137) + Val(txtCP138)) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','修正後總項數','" & Val(txtCount) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','本次應加規費','" & Val(txtAddFee) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','本次應退規費','" & Val(txtDecreaseFee) & "')"
      End If
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','備註','否')"
   
   '附件-基本資料表
   If chkAtt(0).Value = 0 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','未變更本案基本資料')"
   Else
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & chkAtt(0).Tag & "')"
   End If
   
   '*******************************************************************************
   '附送書件
   '*******************************************************************************
   strTmp = ""
   For Each chk In chk1Tab1
      If chk.Value = 1 Then
         strTmp1 = "": strTmp2 = ""
         strTmp1 = chk.Caption
         If strTmp = "" Then
            strTmp1 = "　【" & strTmp1 & "】"
            If Len(strTmp1) < 14 Then
               strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
            End If
         Else
            strTmp1 = "　　【" & strTmp1 & "】"
            If Len(strTmp1) < 15 Then
               strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
            End If
         End If
         If chk.Tag <> "" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
               strTmp2 = Mid(chk.Tag, 2)
            Else
               strTmp2 = Trim(chk.Tag)
            End If
            If strTmp2 <> "" Then
               strTmp2 = m_CaseNo & strTmp2
            End If
         End If
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
      End If
   Next
   If strTmp <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','附送書件','　" & strTmp & "')"
   End If
   '*******************************************************************************
   
   '其他
   If chkAtt(2).Value = 1 Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
      If chkAtt2(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','" & Text1 & "')"
      End If
      If chkAtt2(1).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & m_CaseNo & chkAtt2(1).Tag & "')"
      End If
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

