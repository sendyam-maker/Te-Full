VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090126_New 
   BorderStyle     =   1  '單線固定
   Caption         =   "查名單輸入(網中)"
   ClientHeight    =   7380
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7068
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7068
   Begin VB.CommandButton cmdPicClear 
      Caption         =   "清除圖片"
      Height          =   380
      Left            =   3000
      TabIndex        =   20
      Top             =   5088
      Width           =   900
   End
   Begin VB.Frame Frame3 
      Height          =   1356
      Left            =   48
      TabIndex        =   46
      Top             =   5544
      Width           =   3732
      Begin VB.TextBox txtDB 
         Height          =   300
         Index           =   5
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   22
         Top             =   240
         Width           =   860
      End
      Begin VB.TextBox txtDB 
         Height          =   300
         Index           =   4
         Left            =   1128
         MaxLength       =   7
         TabIndex        =   21
         Top             =   240
         Width           =   860
      End
      Begin VB.CheckBox ChkS3 
         Caption         =   "僅查本所代理"
         Height          =   228
         Index           =   1
         Left            =   1464
         TabIndex        =   23
         Top             =   612
         Width           =   1524
      End
      Begin VB.CheckBox ChkS3 
         Caption         =   "是"
         Height          =   228
         Index           =   2
         Left            =   2304
         TabIndex        =   24
         Top             =   1008
         Width           =   588
      End
      Begin VB.Label Label2 
         Caption         =   "～"
         Height          =   228
         Left            =   2016
         TabIndex        =   51
         Top             =   288
         Width           =   228
      End
      Begin VB.Label Label1 
         Caption         =   "查詢區間："
         Height          =   252
         Index           =   12
         Left            =   96
         TabIndex        =   49
         Top             =   288
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "查詢資料範圍："
         Height          =   252
         Index           =   8
         Left            =   96
         TabIndex        =   48
         Top             =   624
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "是否包含無效或核駁資料："
         Height          =   252
         Index           =   9
         Left            =   96
         TabIndex        =   47
         Top             =   1008
         Width           =   2160
      End
   End
   Begin VB.TextBox txtDB 
      BorderStyle     =   0  '沒有框線
      Height          =   560
      Index           =   3
      Left            =   1356
      Locked          =   -1  'True
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   45
      Top             =   2088
      Width           =   5600
   End
   Begin VB.PictureBox G_SeekPicColor 
      Height          =   2232
      Left            =   7464
      ScaleHeight     =   182
      ScaleMode       =   3  '像素
      ScaleWidth      =   235
      TabIndex        =   41
      Top             =   4968
      Width           =   2868
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   380
      Left            =   72
      TabIndex        =   39
      Top             =   3720
      Width           =   4164
      Begin VB.TextBox txtOther 
         Height          =   300
         Index           =   3
         Left            =   3576
         MaxLength       =   2
         TabIndex        =   14
         Top             =   48
         Width           =   420
      End
      Begin VB.TextBox txtOther 
         Height          =   300
         Index           =   2
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   13
         Top             =   48
         Width           =   300
      End
      Begin VB.TextBox txtOther 
         Height          =   300
         Index           =   1
         Left            =   2352
         MaxLength       =   6
         TabIndex        =   12
         Top             =   48
         Width           =   852
      End
      Begin VB.TextBox txtOther 
         Height          =   300
         Index           =   0
         Left            =   1848
         MaxLength       =   2
         TabIndex        =   11
         Top             =   48
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已收文本所案號："
         Height          =   180
         Index           =   10
         Left            =   336
         TabIndex        =   40
         Top             =   96
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdGrp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "輸入"
      Height          =   330
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   2304
      Width           =   820
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "離開(&X)"
      Height          =   420
      Index           =   1
      Left            =   5976
      TabIndex        =   38
      Top             =   132
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "送出(&P)"
      Height          =   420
      Index           =   0
      Left            =   5040
      TabIndex        =   37
      Top             =   132
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "檢索方式"
      ForeColor       =   &H00FF00FF&
      Height          =   420
      Left            =   96
      TabIndex        =   36
      Top             =   4176
      Width           =   6900
      Begin VB.CheckBox ChkS4 
         Caption         =   "文字及圖形檢索"
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   2
         Left            =   4584
         TabIndex        =   17
         Top             =   144
         Width           =   1548
      End
      Begin VB.CheckBox ChkS4 
         Caption         =   "圖形檢索"
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   1
         Left            =   2988
         TabIndex        =   16
         Top             =   144
         Width           =   1068
      End
      Begin VB.CheckBox ChkS4 
         Caption         =   "文字檢索"
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   0
         Left            =   1392
         TabIndex        =   15
         Top             =   144
         Width           =   1068
      End
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "選擇圖片(&E)"
      Height          =   380
      Left            =   1392
      TabIndex        =   19
      Top             =   5088
      Width           =   1164
   End
   Begin VB.PictureBox tmpPic 
      Height          =   2232
      Left            =   3984
      ScaleHeight     =   182
      ScaleMode       =   3  '像素
      ScaleWidth      =   235
      TabIndex        =   35
      Top             =   5088
      Width           =   2868
      Begin VB.Image tmpImg 
         Height          =   1776
         Left            =   480
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1896
      End
   End
   Begin VB.CheckBox ChkS2 
      Caption         =   "證明標章"
      ForeColor       =   &H00FF0000&
      Height          =   250
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   840
      Width           =   996
   End
   Begin VB.CheckBox ChkS2 
      Caption         =   "團體標章"
      ForeColor       =   &H00FF0000&
      Height          =   250
      Index           =   0
      Left            =   4728
      TabIndex        =   4
      Top             =   840
      Width           =   996
   End
   Begin VB.CheckBox ChkS3 
      Caption         =   "是"
      Height          =   250
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   492
   End
   Begin VB.CheckBox ChkS1 
      Caption         =   "保留組群"
      ForeColor       =   &H00FF00FF&
      Height          =   250
      Index           =   1
      Left            =   3792
      TabIndex        =   2
      Top             =   456
      Width           =   996
   End
   Begin VB.CheckBox ChkS1 
      Caption         =   "保留資料"
      ForeColor       =   &H00FF00FF&
      Height          =   250
      Index           =   0
      Left            =   2664
      TabIndex        =   1
      Top             =   456
      Width           =   996
   End
   Begin VB.TextBox txtDB 
      Height          =   560
      Index           =   2
      Left            =   1356
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   7
      Top             =   1478
      Width           =   5600
   End
   Begin VB.TextBox txtDB 
      Height          =   300
      Index           =   1
      Left            =   1356
      MaxLength       =   30
      TabIndex        =   6
      Text            =   "所有欄位長度和順序都已設好"
      Top             =   1128
      Width           =   5600
   End
   Begin VB.TextBox txtDB 
      Height          =   300
      Index           =   0
      Left            =   1056
      MaxLength       =   6
      TabIndex        =   0
      Top             =   432
      Width           =   708
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   324
      Left            =   5040
      TabIndex        =   52
      Top             =   3696
      Visible         =   0   'False
      Width           =   1380
      VariousPropertyBits=   746604571
      Size            =   "2434;572"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "3519組群："
      Height          =   204
      Index           =   13
      Left            =   360
      TabIndex        =   50
      Top             =   2088
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "智權備註："
      Height          =   252
      Index           =   11
      Left            =   384
      TabIndex        =   44
      Top             =   3264
      Width           =   900
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   444
      Index           =   1
      Left            =   1356
      TabIndex        =   10
      Top             =   3192
      Width           =   5604
      VariousPropertyBits=   -1467989989
      MaxLength       =   50
      ScrollBars      =   2
      Size            =   "9885;783"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   240
      Left            =   96
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7008
      Width           =   3144
      VariousPropertyBits=   671105055
      Size            =   "5546;423"
      Value           =   "Create ID: "
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAutoNo 
      Height          =   252
      Left            =   1080
      TabIndex        =   42
      Top             =   132
      Width           =   1356
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   384
      Index           =   2
      Left            =   1368
      TabIndex        =   18
      Top             =   4656
      Width           =   5496
      VariousPropertyBits=   -1467989989
      MaxLength       =   50
      ScrollBars      =   2
      Size            =   "9694;670"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "不同代號請用"".""或"",""分隔"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   2424
      TabIndex        =   34
      Top             =   840
      Width           =   2172
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   444
      Index           =   0
      Left            =   1356
      TabIndex        =   9
      Top             =   2698
      Width           =   5604
      VariousPropertyBits=   -1467989989
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "9885;783"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   252
      Left            =   1800
      TabIndex        =   33
      Top             =   456
      Width           =   804
      Size            =   "1418;444"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "圖形："
      Height          =   252
      Index           =   7
      Left            =   744
      TabIndex        =   32
      Top             =   5136
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "文字："
      Height          =   252
      Index           =   6
      Left            =   744
      TabIndex        =   31
      Top             =   4704
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "客戶名稱："
      Height          =   252
      Index           =   5
      Left            =   384
      TabIndex        =   30
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "組群："
      Height          =   204
      Index           =   4
      Left            =   696
      TabIndex        =   29
      Top             =   1536
      Width           =   588
   End
   Begin VB.Label Label1 
      Caption         =   "類別(2碼)："
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   28
      Top             =   1176
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "是否進行全類檢索："
      Height          =   252
      Index           =   2
      Left            =   96
      TabIndex        =   27
      Top             =   864
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "委查人："
      Height          =   252
      Index           =   1
      Left            =   96
      TabIndex        =   26
      Top             =   456
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "查名單號："
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   0
      Left            =   96
      TabIndex        =   25
      Top             =   132
      Width           =   900
   End
End
Attribute VB_Name = "frm090126_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/06/11 Form2.0 ; lblFM2, txtFM2, textCUID
Option Explicit
Dim m_PrevForm As Form '前一畫面
Dim mApNoList As String
Dim iLr As Integer '記錄mApNoList的第幾筆資料
Dim strTemp As String, strTemp1 As String
Dim bolPack As Boolean      '是否送出
Dim stKeyUser As String '從接洽單傳委查人
Dim m_AttachPath As String  '預設資料夾
Dim bolRe As Boolean        '送出後,保留目前輸入資料
Dim tmpReNo As String       '保留送出單號
Dim haveKey As String       '已上傳附件
Dim iList As String         '產生的查名單List
Dim strCP09 As String       '已收文案件,申請的總收文號
Dim strCP10 As String       '已收文案件,進度檔的案件性質
Dim bolChkDept As Boolean '詢問是否代填查名單
Dim nfrm090132 As Form
Dim oText As Control

Public Sub SetParent(ByRef fm As Form, Optional ByVal pAppNo As String, Optional ByVal pUserNo As String)
   Set m_PrevForm = fm
   mApNoList = pAppNo
   
   If pUserNo <> "" Then '從接洽單傳委查人
      stKeyUser = pUserNo
   End If
   
End Sub

Private Sub ChkS1_Click(Index As Integer)
   If ChkS1(Index).Value = 1 Then
      For Each oText In ChkS1
         If oText.Index <> Index Then
            oText.Value = 0
         End If
      Next
   End If
End Sub

Private Sub ChkS2_Click(Index As Integer)
   If ChkS2(Index).Value = 1 Then
      For Each oText In ChkS2
         If oText.Index <> Index Then
            oText.Value = 0
         End If
      Next
      ChkS3(0).Value = 0
      txtDB(1) = ""
   End If
End Sub

Private Sub ChkS3_Click(Index As Integer)
   ChkS2(0).Value = 0
   ChkS2(1).Value = 0
End Sub

Private Sub ChkS4_Click(Index As Integer)
   If ChkS4(Index).Value = 1 Then
      For Each oText In ChkS4
         If oText.Index <> Index Then
            oText.Value = 0
         End If
      Next
   End If
End Sub

Private Sub cmdGrp_Click()
   If Not nfrm090132 Is Nothing Then
      nfrm090132.SetParent IIf(mApNoList <> "", "Q", "M"), Me, txtDB(3), IIf(ChkS4(1).Value = "1", "P", "W")
      nfrm090132.Show vbModal
   End If
End Sub

Public Sub SetData(ByVal pInputVal As String)
   txtDB(3).Text = pInputVal
   If lblAutoNo.Caption = "" And Trim(txtDB(1) & txtDB(2) & txtDB(3)) <> "" Then
      '改用最大流水號+1碼,(保留)舊Code
      'lblAutoNo.Caption = AutoNo("HH", 5)  '輸入組群後,給予查名單號
      lblAutoNo.Caption = Pub_GetAutoTMA01
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
   Case 0 '送出
   
      If TxtValidate = False Then Exit Sub
      
      If ChkS4(1).Value = 1 Or ChkS4(2).Value = 1 Then
         If PUB_TMQAppFileChk(lblAutoNo.Caption, "1", TMQ_附件F04) = False Then
            MsgBox "未選擇圖片，請檢查！", vbExclamation, "輸入檢查"
            Exit Sub
         End If
      End If
      
      '1234
      If Index = 0 And strUserNum = "A3034" And Trim(txtFM2(1)) = "" Then
         If MsgBox("是否要輸入智權備註？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            Exit Sub
         End If
      End If
      
      '保留文字查名
      If ChkS4(0).Value = 1 Then
          Set tmpImg.Picture = Nothing
      '保留圖形查名
      ElseIf ChkS4(1).Value = 1 Then
          txtFM2(2).Text = ""
      End If
      
      '檢查是否有重複申請
      strExc(0) = "select tma01 from TMQAppForm where tma13 is null and tma01<>'" & lblAutoNo.Caption & "' and tma03='" & txtDB(0) & "' and tma22='" & txtDB(1) & "' and tma23='" & txtDB(2) & "' and tma24='" & txtDB(3) & "' and tma18='" & ChgSQL(txtFM2(0)) & "' "
      If Trim(txtFM2(2)) <> "" Then strExc(0) = strExc(0) & " and tma26='" & ChgSQL(txtFM2(2)) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("委查單號:" & RsTemp.Fields("tma01") & " 與目前查名有相同客戶名稱、組群" & IIf(Trim(txtFM2(2)) <> "", "和查名文字", "") & "，請確認是否繼續?", vbCritical + vbYesNo + vbDefaultButton2, "輸入檢查") = vbNo Then
            Exit Sub
         End If
      End If
      
      Screen.MousePointer = vbHourglass
      If FormSave() = True Then
         bolPack = True
         MsgBox "已產生查名單：" & lblAutoNo.Caption, vbInformation + vbOKOnly
         If ChkS1(0).Value = 1 Or ChkS1(1).Value = 1 Then bolRe = True
         FormReset
         txtDB(1).SetFocus
         txtDB_GotFocus 1
      Else
         cmdOK(0).Enabled = False
      End If
      Screen.MousePointer = vbDefault
   Case 1  '離開
      Unload Me
   Case Else
   End Select
   
   Exit Sub

End Sub

Private Sub CmdPic_Click()
  If isCheckInput = False Then
     Exit Sub
  End If
  
  Call Do_Picture
End Sub

Private Sub cmdPicClear_Click()
   If tmpImg.Picture <> 0 Then
      If MsgBox("是否清除圖片？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      If lblAutoNo.Caption <> "" Then
         If PUB_TMQAppFileDel(lblAutoNo.Caption, "1", TMQ_附件F04) = False Then
            Exit Sub
         End If
      End If
   End If
   Set tmpImg.Picture = Nothing
   Set G_SeekPicColor.Picture = LoadPicture("")
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()
Dim tmpBol As Boolean
Dim LrArray As Variant

   Me.Width = 7140
   MoveFormToCenter Me
   
   FormReset
   textCUID.BackColor = &H8000000F
   
   '預設委查人
   If stKeyUser = "" Then
      stKeyUser = strUserNum
   End If
   txtDB(0) = stKeyUser
   Call txtDB_Validate(0, tmpBol)
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   strExc(0) = "select * from tmqappsumr"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "請先補上查名人員統計表！", vbExclamation + vbOKOnly, "測試階段"
      Unload Me
      Exit Sub
   End If
   
   iList = ""
   '從查覆區來
   If mApNoList <> "" Then
      ChkS1(0).Visible = False
      ChkS1(1).Visible = False
      LrArray = Split(mApNoList, ",")
      iLr = 0
      If ShowRecData(LrArray(iLr)) = False Then
         MsgBox "查無資料!", vbCritical, "查名單查詢"
         Unload Me
         Exit Sub
      End If
      cmdOK(0).Enabled = False
   End If

   Set nfrm090132 = Forms(0).GetForm("frm090132")
   If Not nfrm090132 Is Nothing And (cmdOK(0).Enabled = True Or cmdGrp.Caption = "顯示") Then
      cmdGrp.Visible = True
   Else
      cmdGrp.Visible = False
   End If
   
   ChkS2(1).Visible = False '2024/10/4 關閉證明標章「9999」代碼
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '刪除無用的查名附件和釋放最新申請編號
    Call ReturnApp(lblAutoNo.Caption, bolPack)
  
    If TypeName(m_PrevForm) <> "Nothing" Then
        Select Case UCase(TypeName(m_PrevForm))
            Case "FRM090127_NEW", "FRM090127_1" '待查區/查覆區/覆核區
                If InStr(m_PrevForm.Caption, "查覆區") > 0 Then
                   m_PrevForm.txtField(6).Text = "0" '預設全部,回查覆區後,是否要分狀態
                Else
                   m_PrevForm.txtField(6).Text = "1"
                End If
                If m_PrevForm.QueryData = False Then
                End If
            Case "FRM090801", "FRM090801_NEW"
                If iList <> "" Then
                   PubShowNextData
                End If
        End Select
        iList = ""
        m_PrevForm.Show
    End If
    
   Set nfrm090132 = Nothing
   
   Set frm090126_New = Nothing
End Sub

'清空欄位
Private Sub FormReset()
   
   lblAutoNo.Caption = ""
   strExc(1) = txtDB(1).Text  '類別
   strExc(2) = txtDB(2).Text  '組群
   strExc(3) = txtFM2(0).Text '客戶名稱
   For Each oText In txtDB
     If oText.Index > 0 Then
        oText.Text = ""
     End If
   Next
   '清除:本所案號
   For Each oText In txtOther
      oText.Text = ""
   Next
   cmdOK(0).Enabled = True: cmdOK(1).Enabled = True
   bolPack = False
   Call UpdateCUID(0)
   
   If (ChkS1(0).Value = 1 Or ChkS1(1).Value = 1) And bolRe = True Then
      tmpReNo = lblAutoNo.Caption
      '(預設)保留客戶名稱
      txtFM2(0).Text = strExc(3)
      
      '保留組群>>不保留資料
      If ChkS1(1).Value = 1 Then
         txtDB(1).Text = strExc(1) '保留類別
         txtDB(2).Text = strExc(2) '保留組群
         '3519組群預設清空,不保留
         '其他:清空檢索方式和內容
         txtFM2(1).Text = ""
         txtFM2(2).Text = ""
         For Each oText In ChkS2
            oText.Value = 0
         Next
         For Each oText In ChkS3
            oText.Value = 0
         Next
         For Each oText In ChkS4
            oText.Value = 0
         Next
         haveKey = ""
         Clipboard.Clear
         Set tmpImg.Picture = Nothing
         Set G_SeekPicColor.Picture = LoadPicture("")
      End If
      '保留資料
      If ChkS1(0).Value = 1 Then
         For Each oText In txtDB
            If oText.Index > 0 Then
               oText.Text = ""
            End If
         Next
         For Each oText In ChkS3
            oText.Value = 0
         Next
      End If
   Else
      bolRe = False: tmpReNo = ""
      For Each oText In txtDB
         If oText.Index > 0 Then
            oText.Text = ""
         End If
      Next
      For Each oText In txtFM2
         oText.Text = ""
      Next
      For Each oText In ChkS2
         oText.Value = 0
      Next
      For Each oText In ChkS3
         oText.Value = 0
      Next
      For Each oText In ChkS4
         oText.Value = 0
      Next
      haveKey = ""
      Clipboard.Clear
      Set tmpImg.Picture = Nothing
      Set G_SeekPicColor.Picture = LoadPicture("")
   End If
   
   txtEna True
   
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtDB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 txtDB(Index)
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp1 As String, strTmp2 As String

   Select Case Index
      Case 0  '委查人
         lblFM2.Caption = ""
         lblFM2.Tag = ""
         If Trim(txtDB(Index)) <> "" Then
            'Added by Lydyia 2025/04/15 排除網中測試資料
            If txtDB(0) = "TMSEARCH" Then
            Else
            'end 2025/04/15
               strTmp1 = GetStaffName(txtDB(Index), False, , strTmp2)
               If strTmp1 = "" Then
                  MsgBox "請輸入正確員工編號！", vbExclamation, "輸入檢查"
                  GoTo EXITSUB
               End If
               lblFM2.Caption = strTmp1
               lblFM2.Tag = strTmp2
            End If
         Else
            MsgBox "請輸入正確員工編號！", vbExclamation, "輸入檢查"
            GoTo EXITSUB
         End If
      Case 1, 2  '類別, 組群
         If txtDB(Index) <> Empty Then
            txtDB(Index).Text = Replace(txtDB(Index).Text, ".", ",") '組群間隔置換為","
            strExc(4) = PUB_RepToOneSpace(PUB_StringFilter(txtDB(Index).Text))   '清除字串中的enter & 清除連續空白
            txtDB(Index).Text = IIf(Right(strExc(4), 1) = ",", Mid(strExc(4), 1, Len(strExc(4)) - 1), strExc(4))
            If Index = 2 And InStr(txtDB(Index), "3519") > 0 Then
                '參考---Modified by Lydia 2024/07/18 避免智權人員誤解，直接拿掉組群已輸入的3519---嘉雯
                txtDB(Index).Text = Replace(Replace(Replace(txtDB(Index), ",3519", ""), "3519,", ""), "3519", "")
            End If
            If Pub_ChkTMQCisExist(Me.Name, txtDB(Index), "" & Index, IIf(ChkS4(0).Value = 1 Or ChkS4(2).Value = 1, "W", "P")) = False Then
               GoTo EXITSUB
            End If
            
            '輸入組群後,給予查名單號
            If lblAutoNo.Caption = "" Then
               '改用最大流水號+1碼,(保留)舊Code
               'lblAutoNo.Caption = AutoNo("HH", 5)
               lblAutoNo.Caption = Pub_GetAutoTMA01
               '保留上次輸入資料
               If (ChkS1(0).Value = 1 Or ChkS1(0).Value = 1) And tmpReNo <> "" And haveKey <> "" Then
                  If AttachFileRedo(tmpReNo, haveKey) = False Then
                  End If
               End If
            End If
         End If
      Case 4, 5 '查詢區間
         If txtDB(Index).Text <> "" Then
            If CheckIsTaiwanDate(txtDB(Index).Text) = False Then
               MsgBox "請輸入民國年月日！", vbCritical
               GoTo EXITSUB
            End If
         End If
   End Select
   Exit Sub
   
EXITSUB:
   txtDB(Index).SetFocus
   txtDB_GotFocus Index
   Cancel = True
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtFM2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtFM2(Index)
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer, Optional ByRef rsSrcTmp As ADODB.Recordset)
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   Else
      If IsNull(rsSrcTmp.Fields("TMA03")) = False Then
         If IsEmptyText(rsSrcTmp.Fields("TMA03")) = False Then
            strCName = GetStaffName(rsSrcTmp.Fields("TMA03"), True)
         End If
      End If
      If IsNull(rsSrcTmp.Fields("TMA04")) = False Then
         If IsEmptyText(rsSrcTmp.Fields("TMA04")) = False Then
            strCDate = ChangeWStringToTDateString(Format(rsSrcTmp.Fields("TMA04"), "yyyymmdd"))
            strCTime = Format(rsSrcTmp.Fields("TMA04"), "hh:mm:dd")
         End If
      End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime
              
End Sub

'控制讀寫
Private Sub txtEna(bolUpd As Boolean)
   Dim bolP As Boolean
   
   If bolUpd = True Then
      bolP = False
   Else
      bolP = True
   End If
   
   For Each oText In txtDB
      If oText.Index <> 3 Then  '排除:3519組群
         oText.Locked = bolP
      End If
   Next
   For Each oText In txtFM2
      oText.Locked = bolP
   Next
   For Each oText In txtOther
      oText.Locked = bolP
   Next

   cmdPic.Enabled = bolUpd
   cmdPicClear.Enabled = bolUpd
   Frame2.Enabled = bolUpd '已收文
   Frame1.Enabled = bolUpd '檢索方式
   Frame3.Enabled = bolUpd '查詢區間
   ChkS3(0).Enabled = bolUpd
   ChkS2(0).Enabled = bolUpd
   ChkS2(1).Enabled = bolUpd
   If bolUpd = False Then
      ChkS3(0).BackColor = &H80000005
      ChkS2(0).BackColor = &H80000005
      ChkS2(1).BackColor = &H80000005
   End If
End Sub

'載入資料
Private Function ShowRecData(ByVal iNo As String) As Boolean
Dim rsR As New ADODB.Recordset
Dim intQ As Integer

   ShowRecData = False
   intQ = 1
   If rsR.State <> adStateClosed Then rsR.Close
   strSql = "select a.*,s1.st02 from tmqappform a,staff s1 where tma03=st01(+) and TMA01='" & iNo & "' "
   Set rsR = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      lblAutoNo.Caption = iNo
      txtDB(0) = "" & rsR.Fields("TMA03")
      lblFM2 = "" & rsR.Fields("ST02")
      txtDB(1) = "" & rsR.Fields("TMA22") '類別
      txtDB(2) = "" & rsR.Fields("TMA23") '組群
      txtDB(3) = "" & rsR.Fields("TMA24") '3519組群
      txtFM2(0) = "" & rsR.Fields("TMA18") '客戶名稱
      txtFM2(1) = "" & rsR.Fields("TMA33") '智權備註
      '已收文本所案號=委查人輸入
      If "" & rsR.Fields("TMA35") <> "" Then
         strExc(0) = "" & rsR.Fields("TMA35")
         Call ChgCaseNo(strExc(0), strExc)
         txtOther(0) = strExc(1)
         txtOther(1) = strExc(2)
         txtOther(2) = strExc(3)
         txtOther(3) = strExc(4)
      End If
      If Val("" & rsR.Fields("TMA25")) > 0 Then '檢索方式
         ChkS4(Val("" & rsR.Fields("TMA25")) - 1).Value = 1
      End If
      txtFM2(2) = "" & rsR.Fields("TMA26") '文字
      If "" & rsR.Fields("TMA27") = "Y" Then '圖形
         If AttachFileGet(lblAutoNo.Caption) = False Then
         End If
      End If
      If "" & rsR.Fields("TMA20") = "1" Then '團體標章
         ChkS2(0).Value = 1
      End If
      If "" & rsR.Fields("TMA20") = "2" Then '證明標章
         ChkS2(1).Value = 1
      End If
      If "" & rsR.Fields("TMA21") = "Y" Then '是否進行全類檢索
         ChkS3(0).Value = 1
      End If
      If "" & rsR.Fields("TMA29") = "2" Then '僅查本所代理
         ChkS3(1).Value = 1
      End If
      If "" & rsR.Fields("TMA30") = "Y" Then '是否包含無效或核駁資料
         ChkS3(2).Value = 1
      End If
      '查詢區間
      txtDB(4) = TransDate("" & rsR.Fields("TMA31"), 1)
      txtDB(5) = TransDate("" & rsR.Fields("TMA32"), 1)
      
      txtEna False
      Call UpdateCUID(1, rsR)
      ShowRecData = True
   End If
   Set rsR = Nothing

End Function

Private Function AttachFileGet(ByVal mTMF01 As String, Optional ByRef strRfilePath As String = "") As Boolean
Dim outType As String
Dim stTempFile As String

On Error GoTo ErrHnd
   
   AttachFileGet = False
   '開啟時,無法刪除,預設下次開啟表單執行刪檔
   outType = "JPG"   '----網中系統限制圖片只能為JPG檔
   Call PUB_KillTempFile(strUserNum & "\H*." & outType)
   Call PUB_KillTempFile(strUserNum & "\H*." & LCase(outType))
   
   strSql = "select * from TMQAPPFile where TMF01='" & mTMF01 & "' AND TMF02='1' order by TMF03 " 'TradeMarkQuery轉入資料附件>0 ，直接新增只有1個附件=00 (AND TMF03='" & TMQ_附件F04 & "')
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Right(UCase("" & RsTemp.Fields("TMF10")), 4) <> "." & outType Then
         MsgBox "附件非JPG檔，無法載入圖片！", vbCritical
      Else
         stTempFile = m_AttachPath & "\" & mTMF01 & RsTemp.Fields("TMF02") & RsTemp.Fields("TMF03") & "." & LCase(Trim(outType))
         
         strRfilePath = stTempFile
         If PUB_TMQAppFileGet(m_AttachPath, stTempFile, mTMF01, "" & RsTemp.Fields("TMF02"), "" & RsTemp.Fields("TMF03")) = False Then
            MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
            Exit Function
         End If
          
         '預設載入
         Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stTempFile))
         '固定PictureBox中的image,載入圖片後調整圖片大小
         Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpPic, tmpImg)
         AttachFileGet = True
      End If
   End If
   
   Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

'複製前一張單據的查名附件
Private Function AttachFileRedo(ByVal pNo As String, ByVal pKey03 As String) As Boolean
Dim strTempFile As String

   AttachFileRedo = False
   If AttachFileGet(pNo, strTempFile) = True Then
      If PUB_TMQAppFileSave(False, lblAutoNo.Caption, "1", pKey03, strTempFile) = False Then
         Exit Function
      End If
   End If
   AttachFileRedo = True

End Function

'插入圖片
Private Sub Do_Picture()

   frmPic001.oCP01 = lblAutoNo.Caption
   frmPic001.oCP02 = "0"
   frmPic001.oCP03 = "1"
   frmPic001.oCP04 = TMQ_附件F04
   Set frmPic001.oPic = G_SeekPicColor
   Set frmPic001.oImg = tmpImg
   Set frmPic001.UpForm = Me
   frmPic001.oRtPic = False
   frmPic001.m_TMQ = "A"  '與原查名單區別
   frmPic001.cmdOK(4).Visible = False
   frmPic001.cmdOK(5).Visible = False
   frmPic001.cmdOK(6).Visible = False
   frmPic001.cmdOK(7).Visible = False
   frmPic001.cmdOK(2).Caption = "存檔(&O)"
   frmPic001.cmdOK(3).Caption = "離開(&X)"
   frmPic001.Label11.Caption = "選擇圖片"
   frmPic001.cmdOK(0).Left = frmPic001.cmdOK(0).Left - 250
   frmPic001.cmdOK(1).Left = frmPic001.cmdOK(1).Left - 250
   frmPic001.cmdOK(2).Left = frmPic001.cmdOK(2).Left - 250
   frmPic001.cmdOK(3).Left = frmPic001.cmdOK(3).Left - 250
   frmPic001.Width = 3800
   MoveFormToCenter frmPic001
   frmPic001.SetSeekCmdok
   Unload frmpic002
   frmPic001.Show vbModal
   
   '重置圖片
   If AttachFileGet(lblAutoNo.Caption) Then
   End If
          
End Sub

Private Function isCheckInput() As Boolean
Dim bMsg As Boolean

isCheckInput = True

If Trim(lblAutoNo.Caption = "") And (ChkS2(0).Value = 1 Or ChkS2(1).Value = 1) Then
   '改用最大流水號+1碼,(保留)舊Code
   'lblAutoNo.Caption = AutoNo("HH", 5)
   lblAutoNo.Caption = Pub_GetAutoTMA01
End If

If lblAutoNo.Caption = "" And Trim(txtDB(1) & txtDB(2) & txtDB(3)) = "" Then
   MsgBox "請先輸入" & IIf(ChkS3(0).Value = 1 And Trim(txtDB(1)) = "", "類別", "組群") & "!", vbExclamation
   txtDB(IIf(ChkS3(0).Value = 1 And Trim(txtDB(1)) = "", 1, 2)).SetFocus
   txtDB_GotFocus IIf(Trim(txtDB(1)) = "", 1, 2)
   isCheckInput = False
   Exit Function
Else
   If PUB_TMQAppFileChk(lblAutoNo.Caption, "1", TMQ_附件F04) Then bMsg = True
   
   If bMsg = True Then
      If MsgBox("已有查名內容,是否覆蓋原有內容?", vbCritical + vbYesNo, "存檔") = vbYes Then
         Exit Function
      Else
         isCheckInput = False
      End If
   End If
End If

End Function

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
   If txtFM2(Index) <> "" Then
      txtFM2(Index).Text = PUB_RepToOneSpace(PUB_StringFilter(txtFM2(Index).Text))
   End If
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   TextInverse txtOther(Index)
End Sub

Private Sub txtOther_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)

   Select Case Index
      Case 0
         If Trim(txtOther(Index)) <> "" And Trim(txtOther(Index)) <> "T" And Trim(txtOther(Index)) <> "TS" Then
            MsgBox "請輸入T, TS案！", vbExclamation, "輸入檢查"
            GoTo EXITSUB
         End If
      Case 1
         If Trim(txtOther(Index)) <> "" And Len(Trim(txtOther(Index))) <> 6 Then
            MsgBox "請輸入6碼案號！", vbExclamation, "輸入檢查"
            GoTo EXITSUB
         End If
      Case 2
         If Trim(txtOther(0)) <> "" And Trim(txtOther(1)) <> "" And Trim(txtOther(Index)) = "" Then
            txtOther(Index) = "0"
         End If
      Case 3
         If Trim(txtOther(0)) <> "" And Trim(txtOther(1)) <> "" And Trim(txtOther(Index)) = "" Then
            txtOther(Index) = "00"
         End If
   End Select
   Exit Sub
   
EXITSUB:
   txtOther(Index).SetFocus
   txtOther_GotFocus Index
   Cancel = True
   
End Sub

Private Function TxtValidate() As Boolean
Dim tmpBol As Boolean
Dim intP As Integer

   TxtValidate = False
 
    If ChkS4(0).Value = 0 And ChkS4(1).Value = 0 And ChkS4(2).Value = 0 Then
       MsgBox "文字或圖形最少選一種！", vbExclamation, "輸入檢查"
       Exit Function
    Else
       If ChkS4(0).Value = 1 Or ChkS4(2).Value = 1 Then
          If Trim(txtFM2(2)) = "" Then
             MsgBox IIf(ChkS4(0).Value = 1, "文字", "文字及圖形") & "檢索請輸入文字！", vbExclamation, "輸入檢查"
             txtFM2(2).SetFocus
             txtFM2_GotFocus 2
             Exit Function
          End If
          If ChkS4(0).Value = 1 And tmpImg.Picture <> 0 Then
             MsgBox "文字檢索不可輸入圖形！", vbExclamation, "輸入檢查"
             cmdPicClear.SetFocus
             Exit Function
          End If
       ElseIf ChkS4(1).Value = 1 Then
          If Trim(txtFM2(2)) <> "" Then
             MsgBox "圖形檢索不可輸入文字！", vbExclamation, "輸入檢查"
             txtFM2(2).SetFocus
             txtFM2_GotFocus 2
             Exit Function
          End If
       End If
    End If
    
    For Each oText In txtDB
      If Trim(oText.Text) = "" Then
         If oText.Index = 0 Then
            MsgBox "委查人不可以空白！", vbExclamation, "輸入檢查"
            txtDB(0).SetFocus
            txtDB_GotFocus 0
            Exit Function
         End If
      Else
         Call txtDB_Validate(oText.Index, tmpBol)
         If tmpBol = True Then
            Exit Function
         End If
      End If
    Next
    
    If ChkS2(0).Value = 1 Or ChkS2(1).Value = 1 Then
       If ChkS3(0).Value = 1 Then
          MsgBox IIf(ChkS2(0).Value = 1, "團體標章", "證明標章") & "不可進行全類檢索！", vbExclamation, "輸入檢查"
          Exit Function
       End If
       If Trim(txtDB(1) & txtDB(2) & txtDB(3)) <> "" Then
          MsgBox IIf(ChkS2(0).Value = 1, "團體標章", "證明標章") & "不可輸入類別或組群！", vbExclamation, "輸入檢查"
          Exit Function
       End If
    Else
       If ChkS3(0).Value = 1 Then
         If Trim(txtDB(1)) = "" Then
            MsgBox "類別不可以空白！", vbExclamation, "輸入檢查"
            Exit Function
         End If
         If Trim(txtDB(2) & txtDB(3)) <> "" Then
            MsgBox "進行全類檢索不可輸入組群！", vbExclamation, "輸入檢查"
            Exit Function
         End If
       Else
         If Trim(txtDB(1)) <> "" Then
            MsgBox "輸入類別，請設定為進行全類檢索！", vbExclamation, "輸入檢查"
            Exit Function
         End If
         If Trim(txtDB(2) & txtDB(3)) = "" Then
            MsgBox "組群不可以空白！", vbExclamation, "輸入檢查"
            Exit Function
         End If
       End If
    End If
    If Trim(txtDB(4)) <> "" And Trim(txtDB(5)) <> "" Then
       If Val(txtDB(4)) > Val(txtDB(5)) Then
           MsgBox "查詢區間起值不可大於迄值！", vbExclamation, "輸入檢查"
           Exit Function
       ElseIf Val(txtDB(4)) > Val(txtDB(5)) Then
           If MsgBox("查詢區間起值等於迄值，是否回到畫面重新輸入？", vbExclamation + vbYesNo + vbDefaultButton1, "輸入檢查") = vbYes Then
              Exit Function
           End If
       End If
    End If
    
    '檢查符號
    For intI = 1 To 2
       For intP = 1 To Len(txtDB(intI))
         If (Asc(Mid(txtDB(intI), intP, 1)) > Asc("9") Or Asc(Mid(txtDB(intI), intP, 1)) < Asc("0")) And Asc(Mid(txtDB(intI), intP, 1)) <> Asc(",") Then
            MsgBox "輸入錯誤，請輸入數字或是 , (號) ！", vbExclamation, "輸入檢查"
            txtDB(intI).SetFocus
            txtDB_GotFocus intI
            Exit Function
         End If
       Next intP
    Next intI
    
    '商申人員詢問是否代填查名單
    If bolChkDept = False And lblFM2.Tag = "P21" Then
       If MsgBox("委查人:" & lblFM2.Caption & "，請問資料是否正確？", vbCritical + vbYesNo + vbDefaultButton2, "輸入檢查") = vbNo Then
          txtDB(0).SetFocus
          txtDB_GotFocus 0
          Exit Function
       End If
    End If
    bolChkDept = True
    
    If Trim(txtFM2(0)) = "" Then
       MsgBox "客戶名稱不可以空白！", vbExclamation, "輸入檢查"
       txtFM2(0).SetFocus
       txtFM2_GotFocus 0
       Exit Function
    End If
    
    strCP09 = ""
    If Trim(txtOther(0) & txtOther(1) & txtOther(2) & txtOther(3)) <> "" Then
       txtOther(2).Text = Mid(txtOther(2).Text & "0", 1, 1)
       txtOther(3).Text = Mid(txtOther(3).Text & "00", 1, 2)
       strExc(1) = "select cp09,cp10,cp27,cp57 from caseprogress where cp01='" & txtOther(0) & "' and cp02='" & txtOther(1) & "' and cp03='" & txtOther(2) & "' and cp04='" & txtOther(3) & "' and cp57 is null "
       If txtOther(0) = "T" Then
          strExc(1) = strExc(1) & "and instr('" & TMQ_T案 & "', cp10) > 0 "
       ElseIf txtOther(0) = "TS" Then
          strExc(1) = strExc(1) & "and instr('" & TMQ_TS案 & "', cp10) > 0 "
       End If
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
       If intI = 1 Then
          If Not IsNull(RsTemp.Fields("cp27")) Then
             MsgBox "申請進度已發文不可輸入，請查明!", vbCritical, "已收文本所案件"
             Exit Function
          End If
          strCP09 = RsTemp.Fields("cp09")
          strCP10 = RsTemp.Fields("cp10")
       Else
          MsgBox "本所案號無申請/查名的案件進度，請查明!", vbCritical
          Exit Function
       End If
    End If
    txtFM2(0).Text = PUB_RepToOneSpace(PUB_StringFilter(txtFM2(0).Text))  '清除字串中的enter & 清除連續空白
    txtFM2(1).Text = PUB_RepToOneSpace(PUB_StringFilter(txtFM2(1).Text))
    txtFM2(2).Text = PUB_RepToOneSpace(PUB_StringFilter(txtFM2(2).Text))
    
   TxtValidate = True
   
   Exit Function
  
End Function

'存檔
Private Function FormSave() As Boolean
Dim midSql As String, exSQL As String
Dim tmpArr As Variant, tmpTitle As String
Dim inC As Integer '中文Keyword
Dim inE As Integer '英文Keyword
Dim tmpClass As String, inputTime As String, cntX As Integer
Dim chkAllStatus As String '內商查名單分單狀態：若查名中心聯絡開始不分單將狀態改為N，恢復分單將狀態改為Y
Dim strTMA09 As String, strTMA10 As String, strTMA11 As String, strTMA12 As String, strTMA25 'TMA09分發日期,TMA10查覆(查名)人員(商申組),TMA11查覆期限,TMA12送出期限,TMA25檢索方式

On Error GoTo ErrHnd
   
    If Trim(lblAutoNo.Caption) = "" Then
       '改用最大流水號+1碼,(保留)舊Code
       'lblAutoNo.Caption = AutoNo("HH", 5)
       lblAutoNo.Caption = Pub_GetAutoTMA01
    End If
   
    '判斷文字查詢的筆數
    If Trim(txtFM2(2)) <> "" Then
       Call PUB_CountTxtNEC(inE, inC, txtFM2(2))
    End If
    
    haveKey = ""
    strExc(1) = "select tmf03 from tmqappfile where tmf01='" & lblAutoNo.Caption & "' and tmf02='1' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
    If intI = 1 Then
      haveKey = "" & RsTemp.Fields("tmf03")
    End If
    chkAllStatus = Pub_GetSpecMan("內商查名單分單狀態")

    inputTime = Left(Format(ServerTime, "000000"), 4)
    '類別/組群+3519組群
    tmpClass = Mid(IIf(txtDB(1) <> "", "," & txtDB(1), IIf(txtDB(2) <> "", "," & txtDB(2), "") & IIf(txtDB(3) <> "", "," & txtDB(3), "")), 2)
    If tmpClass <> "" Then
       tmpArr = Split(tmpClass, ",")
       cntX = UBound(tmpArr) + 1
    Else
       cntX = 1
    End If
    
    '分發日期
    If Val(inputTime) > 1800 Or UCase(chkAllStatus) = "N" Then
       strTMA09 = ""
    Else
       strTMA09 = strSrvDate(1)
    End If
    '檢索方式
    strTMA25 = IIf(ChkS4(0).Value = 1, "1", IIf(ChkS4(1).Value = 1, "2", IIf(ChkS4(2).Value = 1, "3", "")))
    '查覆(查名)人員(商申組)
    If strTMA09 <> "" Then
       'Added by Lydia 2025/04/21 開放測試模式，指定查名人員
       If strSrvDate(1) <= "20991231" Then
           strTMA10 = "A4022"
       Else
       'end 2025/04/21
           strTMA10 = PUB_GetTMAUserPos(strTMA25)
       End If
       
       '送出期限TMA12/查覆期限TMA11：團體標章和證明標章仍舊由查名人負責，設定為19221111
       If ChkS2(0).Value = 1 Or ChkS2(1).Value = 1 Then
          strTMA12 = "19221111"
          strTMA11 = PUB_GetNewTMADate(IIf(ChkS4(0).Value = 1, "4", "5"), strTMA09, inputTime, chkAllStatus)
       Else
          '只能預設送出期限，查覆期限>>1.TMA07回寫日期Trigger觸發計算 2.遇見不發單時間改成批次
          strTMA12 = PUB_GetNewTMADate(strTMA25, strTMA09, inputTime, chkAllStatus)
       End If
    End If
   
   cnnConnection.BeginTrans
      cnnConnection.Execute "delete from TMQAppForm where TMA01='" & lblAutoNo.Caption & "' and tma03='AAAAAA' " '先刪除保留記錄
      tmpTitle = ">>查名單"
   'TMA01查名單號,TMA02資料來源,TMA03 CREATE ID,TMA04 CREATE DATE,TMA08委查人(員工代號),TMA09分發日期,TMA10查覆(查名)人員(商申組)
       exSQL = "Insert Into TMQAppForm (TMA01,TMA02,TMA03,TMA04,TMA08,TMA09,TMA10"
       midSql = " VALUES ('" & lblAutoNo.Caption & "','1','" & strUserNum & "', sysdate, '" & txtDB(0) & "', '" & strTMA09 & "', '" & strTMA10 & "'"
       '1234 特定建檔日期
       'midSql = " VALUES ('" & lblAutoNo.Caption & "','1','" & strUserNum & "', to_date('20241111'||' '||to_char(sysdate,'HH24MISS'),'yyyymmdd HH24MISS'), '" & txtDB(0) & "', '" & strTMA09 & "', '" & strTMA10 & "'"
   'TMA11查覆期限,TMA12送出期限,TMA18客戶名稱,TMA20：1-團體標章/2-證明標章
       exSQL = exSQL & ",TMA11,TMA12,TMA18,TMA20"
       midSql = midSql & ",'" & strTMA11 & "', '" & strTMA12 & "','" & ChgSQL(txtFM2(0)) & "','" & IIf(ChkS2(0).Value = 1, "1", IIf(ChkS2(1).Value = 1, "2", "")) & "' "
   'TMA21是否進行全類檢索,TMA22類別,TMA23組群(非3519),TMA24商品服務名稱組群(3519組群),TMA25檢索方式
       exSQL = exSQL & ",TMA21,TMA22,TMA23,TMA24,TMA25"
       midSql = midSql & ",'" & IIf(ChkS3(0).Value = 1, "Y", "") & "', '" & ChgSQL(txtDB(1)) & "', '" & ChgSQL(txtDB(2)) & "', '" & ChgSQL(txtDB(3)) & "','" & strTMA25 & "' "
   'TMA26文字,TMA27圖形,TMA29查詢資料範圍,TMA30是否包含無效或核駁資料
       exSQL = exSQL & ",TMA26,TMA27,TMA29,TMA30"
       midSql = midSql & ", '" & ChgSQL(txtFM2(2)) & "', '" & IIf(haveKey <> "", "Y", "") & "', '" & IIf(ChkS3(1).Value = 1, "2", "1") & "', '" & IIf(ChkS3(2).Value = 1, "Y", "") & "' "
   'TMA31查詢區間-起始日期,TMA32查詢區間-終止日期,TMA33智權備註,TMA34=新案收文號,TMA35=已收文本所案號
       exSQL = exSQL & ",TMA31,TMA32,TMA33,TMA34,TMA35"
       midSql = midSql & ", '" & DBDATE(txtDB(4)) & "', '" & DBDATE(txtDB(5)) & "', '" & ChgSQL(txtFM2(1)) & "', '" & strCP09 & "','" & IIf(strCP09 <> "", txtOther(0) & txtOther(1) & IIf(txtOther(2) = "", "0", txtOther(2)) & IIf(txtOther(3) = "", "00", txtOther(3)), "") & "' "
   'TMA36委查中文筆數,TMA37委查英文筆數,TMA38委查圖形筆數
       exSQL = exSQL & ",TMA36,TMA37,TMA38"
       midSql = midSql & ", " & inC * cntX & ", " & inE * cntX & ", " & IIf(haveKey = "", 0, cntX) & " "
       
       strSql = exSQL & ")" & midSql & ")"
       cnnConnection.Execute strSql
       If strTMA10 <> "" Then
          Call PUB_TMAtoTake("1", strTMA10, strTMA25, "1", False)
       End If
       '已收文案件新增TS.Menu 至卷宗區
       If strCP09 <> "" Then
          tmpTitle = ">>卷宗區"
          midSql = Trim(txtOther(0)) & CStr(Val(txtOther(1))) & IIf(txtOther(2) <> "0" Or txtOther(3) <> "00", "-" & txtOther(2), "") & IIf(txtOther(3) <> "00", "-" & txtOther(3), "")
          midSql = midSql & "." & strCP10 & "." & lblAutoNo.Caption & "." & TMQ_查名作業 & ".menu"
          strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                  " values('" & strCP09 & "','" & midSql & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'Y')"
          cnnConnection.Execute strSql
          '增加查名代號記錄=>查名單收文對照檔
          midSql = PUB_GetTMQCaseMapNo(strCP09)
          strSql = "insert into tmqcasemap(tqc01,tqc02,tqc03,tqc04,tqc05,tqc06)" & _
                  " values('" & midSql & "','" & strCP09 & "','" & lblAutoNo.Caption & "','" & strUserNum & "'," & strSrvDate(1) & "," & Left(Format(ServerTime, "000000"), 4) & ") "
          cnnConnection.Execute strSql
          '追加未完成的查名單, 取消原承辦期限和查名齊備日
          strSql = "update CaseProgress SET CP48=null,CP143=0 WHERE CP09='" & strCP09 & "' and cp158=0 "
          cnnConnection.Execute strSql
       End If
    cnnConnection.CommitTrans
 
   FormSave = True
   Exit Function

ErrHnd:
   If Err.Number <> 0 Then
      MsgBox "存檔失敗：" & vbCrLf & Err.Description, "查名單" & tmpTitle
      cnnConnection.RollbackTrans
   End If

End Function

'刪除無用的查名附件
Private Sub ReturnApp(ByVal tmpNo As String, ByVal bolOK As Boolean)
Dim rsR1 As New ADODB.Recordset

    If tmpNo <> "" Or bolOK = False Then
        If rsR1.State <> adStateClosed Then rsR1.Close
        Set rsR1 = Nothing
        rsR1.CursorLocation = adUseClient
        '(保留)
        'rsR1.Open "select count(*) from TMQAppForm where TMA01='" & tmpNo & "' ", cnnConnection
        'If rsR1(0) = 0 Then
        '   '離開
        '   If PUB_TMQAppFileDel(tmpNo, "1", TMQ_附件F04) = False Then
        '      Exit Sub
        '   End If
        'End If
        rsR1.Open "select count(*) from TMQAppForm where TMA01='" & tmpNo & "' and TMA03='AAAAAA' ", cnnConnection
        If rsR1(0) = 1 Then
           '離開
           If PUB_TMQAppFileDel(tmpNo, "1", TMQ_附件F04) = False Then
              Exit Sub
           End If
           cnnConnection.Execute "delete from tmqappform where TMA01='" & tmpNo & "'"
        End If
        '-----------------
    End If
    Set rsR1 = Nothing
End Sub

'回傳查名內容
Public Sub PubShowNextData()
Dim intA As Integer, intB As Integer
Dim mLoad As Boolean
Dim sPath As String
Dim APKind As String
Dim rsR As New ADODB.Recordset
Dim mNo As String

If iList <> "" Then
    Me.Enabled = False: mLoad = False
    Screen.MousePointer = vbHourglass
    APKind = "H"
    
    m_PrevForm.Show
    mLoad = False
    '以最後一個編號的查詢內容為主
      strExc(0) = "select a.*,tmf01,tmf02,tmf03,tmf09,tmf10 from tmqappform a, tmqappfile f where tma01 in (" & GetAddStr(iList) & ") and tma01=tmf01(+) and ('1'=tmf02(+) or '2'=tmf02(+)) order by tma01 desc"
      intI = 1
      Set rsR = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         rsR.MoveFirst
         APKind = rsR.Fields("tqa01")
         mNo = rsR.Fields("tma01") & " " & rsR.Fields("tma25")
         txtUnicode = ""
         Do While Not rsR.EOF
            If mNo <> rsR.Fields("tma01") & " " & rsR.Fields("tma25") Then Exit Do
            
            If "" & rsR.Fields("tma25") <> "2" Then
               txtUnicode = txtUnicode & rsR.Fields("tma26") & " "
            Else
               mLoad = True
               APKind = "" & rsR.Fields("tmf01") & rsR.Fields("tmf02") & rsR.Fields("tmf03")
            End If

            rsR.MoveNext
         Loop
      End If
      If txtUnicode <> "" Then
        m_PrevForm.opt1(0).Value = True
        'm_PrevForm.PicText = txtUnicode 'Mark by Lydia 2024/10/25 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
      ElseIf mLoad = True Then
        sPath = Dir(m_AttachPath & "\" & APKind & "*.*")
        If sPath = "" Then
           mLoad = AttachFileGet(Mid(mNo, 1, InStr(mNo, " ") - 1))
        Else
           sPath = m_AttachPath & "\" & sPath
        End If
        If mLoad = True Then
           m_PrevForm.opt1(1).Value = True
           m_PrevForm.optColor(0).Value = True
           Call m_PrevForm.PicToObj(sPath)
        End If
      End If

    m_PrevForm.cmdTMQ.Tag = iList
    m_PrevForm.Combo1(0).Text = "000" & " " & GetPrjNationName("000")
   '設定案件性質
    Call m_PrevForm.Text1_LostFocus(6)
    Call m_PrevForm.QueryTMQ
    If m_PrevForm.Text1(6) = "T" Then 'TS案無商標種類
       m_PrevForm.Combo6.ListIndex = 0 '接洽單的商標種類
    End If
    m_PrevForm.bolExternalCall = False '還原預設值
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Me.Hide
End If
End Sub
