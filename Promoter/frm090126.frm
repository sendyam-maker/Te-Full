VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090126 
   BorderStyle     =   1  '單線固定
   Caption         =   "查名單輸入"
   ClientHeight    =   7320
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10524
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10524
   Begin VB.CommandButton cmdProcOld 
      Caption         =   "處理舊附件"
      Height          =   372
      Left            =   6072
      TabIndex        =   55
      Top             =   2304
      Visible         =   0   'False
      Width           =   1524
   End
   Begin VB.CommandButton cmdChk1131220 
      Caption         =   "檢查1131220"
      Height          =   372
      Left            =   6048
      TabIndex        =   54
      Top             =   1416
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.CommandButton cmdGrp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "輸入"
      Height          =   310
      Left            =   264
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   2304
      Width           =   840
   End
   Begin VB.CommandButton cmdApp 
      BackColor       =   &H00FFC0FF&
      Caption         =   "下載檔案"
      Height          =   348
      Index           =   2
      Left            =   6024
      Style           =   1  '圖片外觀
      TabIndex        =   51
      Top             =   3792
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdApp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "匯入指定範圍"
      Height          =   348
      Index           =   0
      Left            =   6096
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   2952
      Width           =   1500
   End
   Begin VB.CommandButton cmdTrans 
      BackColor       =   &H00FFFFC0&
      Caption         =   "轉換文字"
      Height          =   405
      Left            =   9528
      Style           =   1  '圖片外觀
      TabIndex        =   49
      Top             =   216
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox tmpKeyPic1 
      Height          =   855
      Left            =   2220
      ScaleHeight     =   67
      ScaleMode       =   3  '像素
      ScaleWidth      =   270
      TabIndex        =   44
      Top             =   3960
      Width           =   3285
      Begin VB.Image tmpKeyImg1 
         Height          =   1770
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   1890
      End
   End
   Begin VB.CheckBox ChkS2 
      Caption         =   "保留組群"
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   4320
      TabIndex        =   43
      Top             =   800
      Width           =   1095
   End
   Begin VB.CheckBox ChkS1 
      Caption         =   "保留資料"
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   3120
      TabIndex        =   42
      Top             =   800
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   300
      Left            =   60
      TabIndex        =   38
      Top             =   2730
      Width           =   4575
      Begin VB.CheckBox chk1 
         Caption         =   "已收文"
         Height          =   225
         Left            =   0
         TabIndex        =   4
         Top             =   10
         Width           =   885
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   6
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   5
         Top             =   15
         Width           =   525
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   7
         Left            =   2460
         MaxLength       =   6
         TabIndex        =   6
         Top             =   15
         Width           =   765
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   8
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   7
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   9
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   8
         Top             =   15
         Width           =   405
      End
      Begin VB.Label Label4 
         Caption         =   "本所案號:"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   5
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5280
      Width           =   1665
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4680
      Width           =   1665
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "文字說明(&H)"
      Height          =   465
      Left            =   720
      TabIndex        =   35
      Top             =   4650
      Width           =   1125
   End
   Begin VB.PictureBox tmpKeyPic2 
      Height          =   855
      Left            =   2220
      ScaleHeight     =   67
      ScaleMode       =   3  '像素
      ScaleWidth      =   270
      TabIndex        =   33
      Top             =   5850
      Width           =   3285
      Begin VB.Image tmpKeyImg2 
         Height          =   1770
         Left            =   0
         Top             =   0
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "選擇PDF"
      Height          =   465
      Index           =   0
      Left            =   720
      TabIndex        =   32
      Top             =   4170
      Width           =   1125
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "插入PDF 2"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   31
      Top             =   5550
      Width           =   1185
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "插入PDF 1"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   30
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3330
      Width           =   1695
      Begin VB.OptionButton opt1 
         Caption         =   "文字"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   765
      End
      Begin VB.OptionButton opt1 
         Caption         =   "圖形"
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   10
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "插入影像2"
      Height          =   255
      Index           =   2
      Left            =   4305
      TabIndex        =   28
      Top             =   5550
      Width           =   1200
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "插入影像1"
      Height          =   255
      Index           =   1
      Left            =   4305
      TabIndex        =   27
      Top             =   3720
      Width           =   1200
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   6720
      ScaleHeight     =   247
      ScaleMode       =   3  '像素
      ScaleWidth      =   264
      TabIndex        =   19
      Top             =   4200
      Width           =   3210
   End
   Begin VB.TextBox txt1 
      Height          =   560
      Index           =   1
      Left            =   1080
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   1
      Top             =   1050
      Width           =   4425
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   450
      Width           =   1005
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "選擇圖片"
      Height          =   465
      Left            =   720
      TabIndex        =   13
      Top             =   3690
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4650
      TabIndex        =   15
      Top             =   30
      Width           =   825
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "送出(&P)"
      Height          =   435
      Index           =   0
      Left            =   3840
      TabIndex        =   14
      Top             =   30
      Width           =   825
   End
   Begin VB.PictureBox tmpPic 
      Height          =   3795
      Left            =   6570
      ScaleHeight     =   312
      ScaleMode       =   3  '像素
      ScaleWidth      =   295
      TabIndex        =   18
      Top             =   810
      Width           =   3585
      Begin VB.Image tmpInsPDF 
         Height          =   372
         Left            =   0
         Picture         =   "frm090126.frx":0000
         Top             =   0
         Visible         =   0   'False
         Width           =   1296
      End
      Begin VB.Image tmpImg 
         Height          =   1770
         Left            =   1425
         Stretch         =   -1  'True
         Top             =   1095
         Width           =   1890
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSA 
      Caption         =   "3519組群："
      Height          =   372
      Index           =   5
      Left            =   96
      TabIndex        =   53
      Top             =   2064
      Width           =   996
   End
   Begin MSForms.TextBox textService 
      Height          =   564
      Left            =   1152
      TabIndex        =   3
      Top             =   2040
      Width           =   4356
      VariousPropertyBits=   -1467989985
      MaxLength       =   80
      Size            =   "7683;995"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCName 
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   1660
      Width           =   4425
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "7805;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblname 
      Height          =   255
      Left            =   2160
      TabIndex        =   48
      Top             =   480
      Width           =   855
      Caption         =   "lblname"
      Size            =   "1508;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   60
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6930
      Width           =   5400
      VariousPropertyBits=   671105055
      Size            =   "9525;529"
      Value           =   "Create ID:            Create Date:   "
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSA 
      Caption         =   "文字貼上：請使用Ctrl+V 文字複製：請使用Ctrl+C"
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   3690
      Width           =   1935
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   495
      Index           =   0
      Left            =   6600
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   3285
      VariousPropertyBits=   -1400879077
      MaxLength       =   50
      Size            =   "5794;873"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000A&
      Caption         =   "證明標章：9999; 113/10/4 隱藏"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   4890
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      Caption         =   "團體標章：8888"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   4650
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "不同組群請用"".""或"",""分隔"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   34
      Top             =   840
      Width           =   2175
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   540
      Index           =   2
      Left            =   2220
      TabIndex        =   12
      Top             =   4920
      Width           =   3285
      VariousPropertyBits=   -1467987941
      MaxLength       =   50
      Size            =   "5794;952"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   540
      Index           =   1
      Left            =   2220
      TabIndex        =   11
      Top             =   3120
      Width           =   3285
      VariousPropertyBits=   -1467987941
      MaxLength       =   50
      Size            =   "5794;952"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSA 
      Caption         =   "2"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblSA 
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblSA 
      Caption         =   "指定商品/服務："
      Height          =   252
      Index           =   1
      Left            =   5952
      TabIndex        =   24
      Top             =   804
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label lblSA 
      Caption         =   "客戶名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label lblNo 
      AutoSize        =   -1  'True
      Caption         =   "申請編號："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   900
   End
   Begin VB.Label lblAutoNo 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   1080
      TabIndex        =   21
      Top             =   240
      Width           =   1605
   End
   Begin VB.Label Label5 
      Caption         =   "注意：                          圖形3519必須輸內容；文字及圖形均不再限組群數(文字限同類)，系統會自動分割成多張查名單。"
      ForeColor       =   &H000000FF&
      Height          =   1470
      Left            =   30
      TabIndex        =   20
      Top             =   5130
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "組群："
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "委查人："
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "frm090126"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/10/04 關閉證明標章「9999」代碼：為因應證明標章審查基準修訂及實務審查方式改變，檢索證明標章相同或近似前案，需依照一般商標作業方式進行，此已通告智權部。
'Memo by Lydia 2021/10/01 改成Form2.0 ; textCUID、lblname、txt1(2)=>textCName、txt1(3)=>textService
    'txtUnicode(index)用「Form20上線日」日期控制不用經過二進位處理
    'txt1(4),txt1(5)設locked=true , 才能排除Unicode檢查(二進位處理用)
'end 2021/10/01
'CREATE BY Lydia 2015/06/26 查名單作業(查名單電子化)
Option Explicit
Dim m_PrevForm As Form '前一畫面
Public mApNoList As String
Dim iLr As Integer '記錄mApNoList的第幾筆資料

Dim iCount As Integer
Public IsWmf As Boolean
Dim file_num1 As Integer, file_num2 As Integer '文字1,2用的二進位檔
Dim oText As Control
 
Dim strTemp As String, strTemp1 As String
Dim bolPack As Boolean      '是否送出

Dim NowTMQno As String      '查名單(分單單號)
Dim m_AttachPath As String  '預設資料夾
'Public strLoadPath As String '讀取前次設定路徑 'Remove by Lydia 2016/05/26
Dim tmpS3519 As String      '保留指定商品與服務
Dim bolRe As Boolean        '送出後,保留目前輸入資料
Dim tmpReNo As String       '保留送出單號
Dim haveKey As String       '附件ListIndex
Dim iList As String         '產生的查名單List
Dim strCP09 As String       '已收文案件,申請的總收文號 Added by Lydia 2016/04/06
Dim strCP10 As String       '已收文案件,進度檔的案件性質 'Added by Lydia 2016/04/25
Public stKeyUser As String 'Added by Lydia 2016/04/28 從接洽單傳委查人
Dim bolChkDept As Boolean 'Added by Lydia 2017/01/20 詢問是否代填查名單
Dim cInP As Integer  'Added by Lydia 2018/04/01 文字查名單最大組群個數(分割原則)

Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/22 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim nFrm As Form 'Added by Lydia 2024/07/17
Dim SyxMsg As String 'Added by Lydia 2018/11/22 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub ChkS1_Click()
   If ChkS1.Value = 1 Then
      ChkS2.Value = 0
   End If
End Sub

Private Sub ChkS2_Click()
   If ChkS2.Value = 1 Then
      ChkS1.Value = 0
   End If
End Sub

Private Sub cmdChk1131220_Click()
Dim strR1 As String, intR As Integer
Dim rsRd As New ADODB.Recordset

''Memo by Lydia 2024/12/20 自112/9/21停止批次刪除作業,113/12/20恢復批次刪除作業--每日批次執行發生FTP連線失敗,下方程式執行結果與正式DB殘留的TMQFile一樣
'   strExc(0) = App.path & "\" & strUserNum
'   If Dir(strExc(0), vbDirectory) = "" Then
'      MkDir strExc(0)
'   End If
'
'   strExc(3) = "": strExc(5) = ""
'   strR1 = "select tmq01,tmq20,tmq21,c.tqf12 as ntqf12,b.* from trademarkquery a,tmqfile_1131218  b,tmqfile c " & _
'           "where tmq20 like '113/12/20%' " & _
'           "and tmq01=b.tqf02(+) and nvl(b.tqf12,'N') <>'N' and b.tqf01=c.tqf01(+) and b.tqf02=c.tqf02(+) and b.tqf03=c.tqf03(+) and b.tqf04=c.tqf04(+) " & _
'           "order by b.tqf01,b.tqf02,b.tqf03,b.tqf04 "
'   intR = 1
'   Set rsRd = ClsLawReadRstMsg(intR, strR1)
'   If intR = 1 Then
'      rsRd.MoveFirst
'      Do While Not rsRd.EOF
'         If strExc(5) <> "" & rsRd.Fields("tmq01") Then
'            Sleep 100
'            Call PUB_KillTempFile(strUserNum & "\*.*")
'            strExc(5) = "" & rsRd.Fields("tmq01")
'         End If
'         If PUB_GetFtpFile("" & rsRd.Fields("tqf12"), strExc(0) & "\" & rsRd.Fields("tqf02") & rsRd.Fields("tqf03") & rsRd.Fields("tqf04") & ".TS.pdf", "TMQFILE") = True Then
'            strExc(3) = strExc(3) & vbCrLf & rsRd.Fields("TQF02") & "-" & rsRd.Fields("TQF03") & "-" & rsRd.Fields("TQF04") & ":" & rsRd.Fields("TQF12")
'         End If
'         rsRd.MoveNext
'      Loop
'      If strExc(3) <> "" Then
'         PUB_SendMail strUserNum, "A3034", "", "尚有TMQFile", strExc(3)
'      End If
'      MsgBox "OK"
'   End If

End Sub

Private Sub cmdHelp_Click()
   strExc(0) = ""
   strExc(0) = strExc(0) & "1. 可以從它處複製文字，然後按Ctrl+V貼上。" & vbCrLf
   strExc(0) = strExc(0) & "2. 若不知道文字的輸入碼則可以在對應的欄位輸入'中文'、'English'、'日文'或阿拉伯數字，" & vbCrLf
   strExc(0) = strExc(0) & "   依正確文字內容自行組合(ex. 中文日文 )，然後插入正確文字的JPG或PDF檔。" & vbCrLf
   strExc(0) = strExc(0) & "3. 若遇見輸入法切換後，無法輸入的情況，請關閉查名單輸入畫面，再重新呼叫一次。" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   MsgBox strExc(0), vbOKOnly, "文字說明"
End Sub

Private Sub cmdKey_Click(Index As Integer)
Dim tmpI As Integer
  Select Case Index
       Case 0: tmpI = Val(TMQ_AkindPic)
       Case 1, 3: tmpI = Val(TMQ_AkindWord1)
       Case 2, 4: tmpI = Val(TMQ_AkindWord2)
  End Select
  If isCheckInput(tmpI) = False Then
     Exit Sub
  End If
  Select Case Index
     Case 1, 2   '影像
          Call Do_Picture(Index)
     Case 3, 4   'PDF
          Call AttachFileAdd(Index - 2)
     Case 0   '圖形查詢-PDF
          Call AttachFileAdd(0)
  End Select

End Sub
Private Function TxtValidate() As Boolean
Dim tmpBol As Boolean
Dim Myi As Integer

  TxtValidate = False
 
    If opt1(0).Value = False And opt1(1).Value = False Then
       MsgBox "文字或圖形最少選一種！", vbExclamation, "操作錯誤！"
       Exit Function
    End If
    If Trim(txt1(0)) = "" Then
       MsgBox "委查人不可以空白！", vbExclamation, "操作錯誤！"
       txt1(0).SetFocus
       Exit Function
    End If
    'Added by Lydia 2021/01/18 文字1和文字2不可相同; ex.HB0010500~HB0010522的文字1和文字2相同,雖然文字2有附圖不等於不同字, 資料已做處理將文字2的附件歸入文字1,並且更新查名筆數和刪除明細列。
    If opt1(0).Value = True Then
        If (txtUnicode(1) = "" And txtUnicode(2) = "") Or (txtUnicode(1) = "" And txtUnicode(2) <> "") Then
             MsgBox "請輸入文字1！", vbExclamation, "操作錯誤！"
             txtUnicode(1).SetFocus
             Exit Function
        End If
        If Trim(txtUnicode(1)) = Trim(txtUnicode(2)) Then
             MsgBox "文字1和文字2不可相同！", vbExclamation, "操作錯誤！"
             txtUnicode(1).SetFocus
             Exit Function
        End If
    End If
    'end 2021/01/18
    
    'Added by Lydia 2017/01/20 商申人員詢問是否代填查名單
    If bolChkDept = False And lblname.Tag = "P21" Then
       'Modified by Lydia 2017/06/08 修改條件
       'If MsgBox("委查人:" & lblName.Caption & "，請問是否為代填查名單？", vbCritical + vbYesNo + vbDefaultButton2, "商申人員代填查名單") = vbYes Then
       If MsgBox("委查人:" & lblname.Caption & "，請問資料是否正確？", vbCritical + vbYesNo + vbDefaultButton2, "商申人員代填查名單") = vbNo Then
          txt1(0).SetFocus
          Exit Function
       End If
    End If
    bolChkDept = True
    
    If Trim(textCName) = "" Then
       MsgBox "客戶名稱不可以空白！", vbExclamation, "操作錯誤！"
       textCName.SetFocus
       Exit Function
    End If
    If Trim(txt1(0)) <> "" And lblname.Caption = "" Then
       MsgBox "委查人輸入錯誤！", vbExclamation, "操作錯誤！"
       txt1(0).SetFocus
       Exit Function
    End If
    'Modified by Lydia 2024/07/17
    'If Trim(txt1(1)) = "" Then
    '   MsgBox "組群不可空白！", vbExclamation, "操作錯誤！"
    If Trim(txt1(1)) = "" And Trim(textService) = "" Then
       MsgBox "請輸入組群或3519組群！", vbExclamation, "操作錯誤！"
    'end 2024/07/17
       txt1(1).SetFocus
       Exit Function
    Else
       txt1_Validate 1, tmpBol
       If tmpBol = True Then
           Exit Function
       End If
    End If
    
    strCP09 = ""
    If Trim(txt1(6) & txt1(7) & txt1(8) & txt1(9)) <> "" Then
        If chk1.Value = False Then
           MsgBox "未勾選已收文!!", vbCritical
           Exit Function
        End If
        txt1(6) = IIf(txt1(6) = "", "T", txt1(6))
        txt1(7) = IIf(txt1(7) = "", "000000", txt1(7))
        txt1(8).Text = Mid(txt1(8).Text & "0", 1, 1)
        txt1(9).Text = Mid(txt1(9).Text & "00", 1, 2)
        If ClsPDCheckCaseCodeIsExist(txt1(6).Text, txt1(7).Text, txt1(8).Text, txt1(9).Text) = False Then
           txt1(9).SetFocus
           Exit Function
        'Added by Lydia 2016/04/06 判斷案件進度是否已發文
        Else
           'Modified by Lydia 2016/04/25 +TS
           'strExc(1) = "select cp09,cp10,cp27,cp57 from caseprogress where cp01='" & txt1(6) & "' and cp02='" & txt1(7) & "' and cp03='" & txt1(8) & "' and cp04='" & txt1(9) & "' and cp10='101' and cp57 is null "
           strExc(1) = "select cp09,cp10,cp27,cp57 from caseprogress where cp01='" & txt1(6) & "' and cp02='" & txt1(7) & "' and cp03='" & txt1(8) & "' and cp04='" & txt1(9) & "' and cp57 is null "
           If txt1(6) = "T" Then
              'Modified by Lydia 2021/11/19 增加737智財協作之T案
              'strExc(1) = strExc(1) & "and cp10='" & TMQ_T案 & "'"
              strExc(1) = strExc(1) & "and instr('" & TMQ_T案 & "', cp10) > 0 "
           ElseIf txt1(6) = "TS" Then
              'Modified by Lydia 2021/11/19
              'strExc(1) = strExc(1) & "and cp10='" & TMQ_TS案 & "'"
              strExc(1) = strExc(1) & "and instr('" & TMQ_TS案 & "', cp10) > 0 "
           End If
           
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
           If intI = 1 Then
              If Not IsNull(RsTemp.Fields("cp27")) Then
                 MsgBox "申請進度已發文不可變更，請查明!", vbCritical
                 Exit Function
              End If
              strCP09 = RsTemp.Fields("cp09")
              strCP10 = RsTemp.Fields("cp10")  'Added by Lydia 2016/04/25
           Else
              MsgBox "本所案號無申請/查名的案件進度，請查明!", vbCritical
              Exit Function
           End If
        End If
    End If
    txt1(1).Text = PUB_RepToOneSpace(PUB_StringFilter(txt1(1).Text))  '清除字串中的enter & 清除連續空白
    txtUnicode(1).Text = PUB_RepToOneSpace(PUB_StringFilter(txtUnicode(1).Text))
    txtUnicode(2).Text = PUB_RepToOneSpace(PUB_StringFilter(txtUnicode(2).Text))
    txt1(4).Text = txtUnicode(1): txt1(5).Text = txtUnicode(2)
     '檢查符號
     For Myi = 1 To Len(txt1(1))
        If (Asc(Mid(txt1(1), Myi, 1)) > Asc("9") Or Asc(Mid(txt1(1), Myi, 1)) < Asc("0")) And Asc(Mid(txt1(1), Myi, 1)) <> Asc(",") Then
            MsgBox "輸入錯誤，請輸入數字或是 , (號) ！", vbExclamation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Function
        End If
     Next Myi
    'Modified by Lydia 2024/07/17 改成3519組群輸入
    'If InStr(1, txt1(1), "3519") > 0 Then
    '    If Trim(textService) = "" Then
    '        MsgBox "指定商品/服務不可以空白！", vbExclamation, "操作錯誤！"
    '        textService.SetFocus
    '        Exit Function
   '     End If
    'End If
    If InStr(txt1(1), "3519") > 0 Then
       'Modified by Lydia 2024/07/18 避免智權人員誤解，直接拿掉組群已輸入的3519---嘉雯
       'MsgBox "3519組群請點選輸入按鈕 ！", vbExclamation, "操作錯誤！"
       'txt1(1).SetFocus
       'txt1_GotFocus 1
       'Exit Function
       txt1(1).Text = Replace(Replace(Replace(txt1(1), ",3519", ""), "3519,", ""), "3519", "")
    End If
    'end 2024/07/17
    
   'Added by Lydia 2024/10/04 關閉證明標章「9999」代碼
   If InStr(txt1(1), "9999") > 0 Then
      MsgBox "不可輸入組群代碼9999!!!", vbExclamation + vbOKOnly, "輸入檢查"
      txt1(1).SetFocus
      Exit Function
   End If
   'end 2024/10/04
   
   'Added by Lydia 2021/10/01 用日期控制不用經過二進位處理存入TQA07-TQA08，直接存入TQA13-TQA14
   If strSrvDate(1) >= Form20上線日 Then
      '保留: 需要檢查?的存在？
      'If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      '    Exit Function
      'End If
   Else
      '客戶名稱已改成Form 2.0要檢查, 但未上線前的文字1-2要先排除檢查才能照原本的方式
      If PUB_ChkUniText(Me, , True, "TextBox", "txtUnicode") = False Then
          Exit Function
      End If
   End If
   'end 2021/10/01
   
  TxtValidate = True
End Function
Private Sub cmdok_Click(Index As Integer)
Dim rsRW As New ADODB.Recordset

On Error GoTo ErrHand
    Select Case Index
    Case 0 '送出
    
        If TxtValidate = False Then Exit Sub
       
        If rsRW.State <> adStateClosed Then rsRW.Close
        Set rsRW = Nothing
        rsRW.CursorLocation = adUseClient
        strExc(0) = "select count(*) from TMQFILE where TQF01='" & lblAutoNo.Caption & "' and TQF02='" & TMQ_附件F02 & "' " & _
                    "AND TQF03" & IIf(opt1(0).Value = True, "> '" & TMQ_AkindPic & "'", "='" & TMQ_AkindPic & "'") & " AND TQF04='" & TMQ_附件F04 & "'"
        rsRW.Open strExc(0), cnnConnection
        If rsRW(0) = 0 Then
           If opt1(1).Value = True Or (opt1(0).Value = True And txtUnicode(1).Text = "" And txtUnicode(2).Text = "") Then
               MsgBox "未輸入查名內容,請檢查！", vbExclamation, "輸入錯誤！"
               Exit Sub
           End If
        End If
        
        '保留文字查名
        If opt1(0).Value = True Then
              Set tmpImg.Picture = Nothing
        '保留圖形查名
        ElseIf opt1(1).Value = True Then
              txtUnicode(1).Text = ""
              txtUnicode(2).Text = ""
              Set tmpKeyImg1.Picture = Nothing
              Set tmpKeyImg2.Picture = Nothing
        End If
        
        '檢查是否有重複申請
        intI = 1
        strExc(0) = "select tqa01 from tmqapp where tqa20 is null and tqa01 <> '" & lblAutoNo.Caption & "' and tqa02='" & txt1(0) & "' and tqa03='" & txt1(1) & "' and tqa04='" & textCName & "'"
        If opt1(0).Value = True Then
           strExc(1) = strExc(0)
           If txtUnicode(1) <> "" And txt1(4).Text = txtUnicode(1).Text Then
              strExc(0) = strExc(0) & " and tqa13='" & txt1(4) & "'"
           End If
           If txtUnicode(2) <> "" And txt1(5).Text = txtUnicode(2).Text Then
              strExc(0) = strExc(0) & " and tqa14='" & txt1(5) & "'"
           End If
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              'Modified by Lydia 2016/06/17
              'If MsgBox(RsTemp(0) & "與目前查名有相同客戶名稱、組群" & IIf(strExc(1) <> strExc(0), "和查名文字", "") & ",請確認是否繼續?", vbCritical + vbYesNo) = vbNo Then
              strExc(3) = GetTQA01toTMQ01(RsTemp(0))
              If MsgBox("委查單號:" & strExc(3) & " 與目前查名有相同客戶名稱、組群" & IIf(strExc(1) <> strExc(0), "和查名文字", "") & ",請確認是否繼續?", vbCritical + vbYesNo) = vbNo Then
                  Exit Sub
              End If
           End If
        Else
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              'Modified by Lydia 2016/06/17
              'If MsgBox(RsTemp(0) & "與目前查名有相同客戶名稱、組群,請確認是否繼續?", vbCritical + vbYesNo) = vbNo Then
              strExc(3) = GetTQA01toTMQ01(RsTemp(0))
              If MsgBox("委查單號:" & strExc(3) & " 與目前查名有相同客戶名稱、組群,請確認是否繼續?", vbCritical + vbYesNo) = vbNo Then
                  Exit Sub
              End If
           End If
        End If
        
        If FormSave() = True Then
           '查名單(分單)
           If ProcessPack() = True Then
              bolPack = True
              'Added by Lydia 2024/05/31
              If strSrvDate(1) >= 查名單網中系統平行測試 Then
                 If ProcPackNew(lblAutoNo.Caption) = False Then
                 End If
              End If
              'end 2024/05/31
              
              MsgBox "本次作業共分割成 " & Format(Len(NowTMQno) \ 10, "0") & " 張委查單。", vbInformation, "查名單分單完成"
              
              iList = IIf(iList <> "", iList & ",", "") & lblAutoNo.Caption
              If textService.Text <> "" Then tmpS3519 = textService.Text
           End If

           If (ChkS1.Value = 1 Or ChkS2.Value = 1) Then bolRe = True
           FormReset
           txt1(1).SetFocus
        End If

    Case 1  '離開
        Unload Me
   
    Case Else
    End Select
    
    Exit Sub
ErrHand:
   DataErrorMessage (3)
End Sub
'刪除無用的查名附件和釋放最新申請編號
Private Sub ReturnApp(ByVal tmpNo As String, ByVal bolOK As Boolean)
Dim rsR1 As New ADODB.Recordset

    If tmpNo <> "" Or bolOK = False Then
        If rsR1.State <> adStateClosed Then rsR1.Close
        Set rsR1 = Nothing
        rsR1.CursorLocation = adUseClient
        rsR1.Open "select count(*) from TMQApp where TQA01='" & tmpNo & "'", cnnConnection
        If rsR1(0) = 0 Then
           '離開
            cnnConnection.BeginTrans
              strExc(1) = AccAutoNo("HM", 4, GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2)))
              '若為最新申請單,單號釋回
              If Val(Right(strExc(1), 4)) - 1 = Val(Right(tmpNo, 4)) Then
                strSql = "update acc1r0 set a1r04=a1r04-1 where a1r01='HM' and a1r02=" & Year(ChangeWStringToWDateString(strSrvDate(1))) & " and a1r03=" & Month(ChangeWStringToWDateString(strSrvDate(1)))
                cnnConnection.Execute strSql, intI
              End If
                '刪除不用的圖檔或附件
                'Added by Lydia 2016/06/23
                'Remove by Lydia 2016/07/07
                'If strSrvDate(1) < TMQFileFTP Then
                '    strSql = " DELETE FROM TMQFILE WHERE TQF01='" & tmpNo & "' and TQF02='" & TMQ_附件F02 & "' and TQF04='" & TMQ_附件F04 & "'"
                '    cnnConnection.Execute strSql, intI
                'Else
                    If PUB_TMQAFileDel(tmpNo, TMQ_附件F02, "", TMQ_附件F04) Then
                    End If
                'End If
            cnnConnection.CommitTrans
        End If
    End If
        
End Sub

Private Sub CmdPic_Click()

  If isCheckInput(0) = False Then
     Exit Sub
  End If
  
  Call Do_Picture(0)
End Sub
'載入資料
Private Function ShowRecData(ByVal iNo As String) As Boolean
Dim rsR As New ADODB.Recordset
Dim iQ As Integer
   ShowRecData = False
   iQ = 1
   If rsR.State <> adStateClosed Then rsR.Close
   strSql = "select a.*,s1.st02 from tmqapp a,staff s1 where tqa02=st01(+) and TQA01='" & iNo & "' "
   Set rsR = ClsLawReadRstMsg(iQ, strSql)
   If iQ = 1 Then
      lblAutoNo.Caption = iNo
      txt1(0) = "" & rsR.Fields("TQA02")
      txt1(1) = "" & rsR.Fields("TQA03")
      textCName = "" & rsR.Fields("TQA04")
      textService = "" & rsR.Fields("TQA05")
      'Added by Lydia 2024/07/17
      cmdGrp.Caption = "輸入"
      If "" & rsR.Fields("TQA11") >= "20240625" Or Mid("" & rsR.Fields("TQA05"), 1, 4) = "3519" Then '3519組群輸入啟用日
         lblSA(5) = "3519組群："
         If "" & rsR.Fields("TQA05") <> "" Then
            cmdGrp.Caption = "顯示"
         End If
      Else
         lblSA(5) = "指定商品　/服務："
      End If
      'end 2024/07/17
      '已收文
      If "" & rsR.Fields("TQA15") = "Y" Then
          chk1.Value = 1
      Else
          chk1.Value = 0
      End If
      '本所案號
      txt1(6) = "" & rsR.Fields("TQA16")
      txt1(7) = "" & rsR.Fields("TQA17")
      txt1(8) = "" & rsR.Fields("TQA18")
      txt1(9) = "" & rsR.Fields("TQA19")
      
      '文字查名
      If rsR.Fields("TQA06") = "1" Then
         opt1(0).Value = True
         If Not IsNull(rsR.Fields("TQA07")) Then
            txtUnicode(1).Text = rsR.Fields("TQA07")  'Unicode字
         ElseIf Not IsNull(rsR.Fields("TQA13")) Then
            txtUnicode(1).Text = rsR.Fields("TQA13")  'Big5字
         'Modified by Lydia 2016/03/28 預設顯示附件
         'ElseIf AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord1) = False Then
         End If
         If AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord1) = False Then
         End If
         If Not IsNull(rsR.Fields("TQA08")) Then
            txtUnicode(2).Text = rsR.Fields("TQA08")
         ElseIf Not IsNull(rsR.Fields("TQA14")) Then
            txtUnicode(2).Text = rsR.Fields("TQA14")
         'Modified by Lydia 2016/03/28 預設顯示附件
         'ElseIf AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord2) = False Then
         End If
         If AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord2) = False Then
         End If
      '圖形查名
      Else
         opt1(1).Value = True
         If AttachFileGet(lblAutoNo.Caption, TMQ_AkindPic) = False Then
         End If
      End If
      txtEna False
      Call UpdateCUID(1, rsR) 'Added by Lydia 2016/06/15
      ShowRecData = True
   End If
   Set rsR = Nothing

End Function
Private Function AttachFileGet(ByVal mTQF01 As String, ByVal mTQFkind As String, Optional ByRef bolR As Boolean = False, Optional ByRef strRfilePath As String = "", Optional ByRef strRF As String = "") As Boolean
Dim adoRst As New ADODB.Recordset
Dim outType As String
Dim stTempFile As String
Dim fileN As Integer
Dim bytes() As Byte

On Error GoTo ErrHnd
   
    AttachFileGet = False
    '開啟時,無法刪除,預設下次開啟表單執行刪檔
    If Dir(m_AttachPath & "\HM*.jpg") <> "" Then
        Kill m_AttachPath & "\HM*.jpg"
    End If
    If Dir(m_AttachPath & "\HM*.pdf") <> "" Then
        Kill m_AttachPath & "\HM*.pdf"
    End If
    If Dir(m_AttachPath & "\HM*.JPG") <> "" Then
        Kill m_AttachPath & "\HM*.JPG"
    End If
    If Dir(m_AttachPath & "\HM*.PDF") <> "" Then
        Kill m_AttachPath & "\HM*.PDF"
    End If
    If adoRst.State <> adStateClosed Then adoRst.Close
    Set adoRst = Nothing
    adoRst.CursorLocation = adUseClient
    adoRst.Open "select * from TMQFile where TQF01='" & mTQF01 & "' AND TQF02='" & TMQ_附件F02 & "' AND TQF03='" & mTQFkind & "' AND TQF04='" & TMQ_附件F04 & "'", cnnConnection, adOpenStatic, adLockOptimistic
    If adoRst.RecordCount > 0 Then

       outType = "" & adoRst.Fields("TQF05")
       'Modified by Lydia 2016/06/23
       'stTempFile = m_AttachPath & "\" & mTQF01 & "_" & mTQFkind & "." & LCase(Trim(outType))
       stTempFile = m_AttachPath & "\" & mTQF01 & TMQ_附件F02 & mTQFkind & TMQ_附件F04 & "." & LCase(Trim(outType))
       
       strRfilePath = stTempFile
       strRF = outType
       'Modified by Lydia 2016/06/23 改放在FTP
       'Remove by Lydia 2016/07/07
       'If strSrvDate(1) < TMQFileFTP Then
       '     ReDim bytes(Val(adoRst.Fields("TQF06").Value))
       '     bytes() = adoRst.Fields("TQF07").GetChunk(Val(adoRst.Fields("TQF06").Value))
       '     fileN = FreeFile
       '     Open stTempFile For Binary Access Write As #fileN
       '     Put #fileN, , bytes()
       '     Close #fileN
       'Else
            If PUB_TMQGetAFile(m_AttachPath, stTempFile, mTQF01, TMQ_附件F02, mTQFkind, TMQ_附件F04, outType) = False Then
               MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
               Exit Function
            End If
       'End If
        
       If bolR = False Then
            If InStr(UCase(outType), "PDF") = 0 Then
               Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stTempFile))
               Select Case mTQFkind
                  Case TMQ_AkindPic ' "0"
                      '固定PictureBox中的image,載入圖片後調整圖片大小
                      Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpPic, tmpImg)
                  Case TMQ_AkindWord1 ' "1"
                     'Modified by Lydia 2016/03/28 固定縮放
                     'Set tmpKeyImg1.Picture = pvGetStdPicture(Trim(stTempFile))
                      Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpKeyPic1, tmpKeyImg1)
                  Case TMQ_AkindWord2 '"2"
                      'Modified by Lydia 2016/03/28 固定縮放
                     'Set tmpKeyImg2.Picture = pvGetStdPicture(Trim(stTempFile))
                      Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpKeyPic2, tmpKeyImg2)
               End Select
            Else
               Select Case mTQFkind
                  Case TMQ_AkindPic '"0"
                     Set tmpImg.Picture = tmpInsPDF.Picture
                     tmpImg.Stretch = False
                     tmpImg.Top = 0
                     tmpImg.Left = 0
                  Case TMQ_AkindWord1 ' "1"
                     Set tmpKeyImg1.Picture = tmpInsPDF.Picture
                     'Added by Lydia 2016/03/28
                     tmpKeyImg1.Stretch = False
                     tmpKeyImg1.Top = 0
                     tmpKeyImg1.Left = 0
                  Case TMQ_AkindWord2 '"2"
                     Set tmpKeyImg2.Picture = tmpInsPDF.Picture
                     'Added by Lydia 2016/03/28
                     tmpKeyImg2.Stretch = False
                     tmpKeyImg2.Top = 0
                     tmpKeyImg2.Left = 0
               End Select
            End If
       End If
    Else
       Exit Function
    End If
   
    AttachFileGet = True
    Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
   If fileN > 0 Then Close #fileN
End Function

'Added by Lydia 2025/03/25 處理匯入舊版查名單資料(H11300001~H11300844)的附件---2025/03/25 處理完成
Private Sub cmdProcOld_Click()
'1.刪除文字查詢的附件>>152筆
'2.將圖形查名附件的尾碼2碼從01.JPG改為00.JPG>>329筆
  '若為PDF檔則刪除單據資料(11筆H11300030,H11300032,H11300033,H11300096,H11300097,H11300109,H11300110,H11300111,H11300112,H11300823,H11300844,)。
Dim intA As Integer, strA1 As String
Dim rsAD As New ADODB.Recordset

   strExc(1) = "1.刪除文字查詢的附件"
   MsgBox strExc(1)
   strA1 = "select TMA71,b.* from tmqappform a,tmqappfile b where nvl(TMA71,'N') <> 'N' and tma01=tmf01(+) and tma25<>'2' and tmf01 is not null order by tmf01"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         If PUB_TMQAppFileDel(rsAD.Fields("tmf01"), rsAD.Fields("tmf02"), rsAD.Fields("tmf03")) = False Then
            MsgBox strExc(1) & "：刪檔失敗" & rsAD.Fields("tmf03")
            Exit Sub
         End If
         rsAD.MoveNext
      Loop
   End If

   strExc(1) = "2.刪除圖形查名附件為PDF"
   MsgBox strExc(1)
   strA1 = "select TMA71,b.* from tmqappform a,tmqappfile b where nvl(TMA71,'N') <> 'N' and tma01=tmf01(+) and tma25='2' and tmf10 like '%.PDF' order by tmf01"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         If PUB_TMQAppFileDel(rsAD.Fields("tmf01"), rsAD.Fields("tmf02"), rsAD.Fields("tmf03")) = False Then
            MsgBox strExc(1) & "：刪檔失敗" & rsAD.Fields("tmf03")
            Exit Sub
         Else
            strSql = "delete from tmqappform where tma01='" & rsAD.Fields("tmf01") & "' "
            cnnConnection.Execute strSql
         End If
         rsAD.MoveNext
      Loop
   End If
   
   strExc(1) = "3.將圖形查名附件的尾碼2碼從01.JPG改為00.JPG"
   MsgBox strExc(1)
   strA1 = "select TMA71,b.* from tmqappform a,tmqappfile b where nvl(TMA71,'N') <> 'N' and tma01=tmf01(+) and tma25='2' and tmf10 like '%01.JPG' order by tmf01"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      strExc(0) = PUB_GetFtpTableDir("TMQAPPFILE")
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         'FTP檔案更名
         If PUB_FtpRenFile(strExc(0) & "/" & rsAD.Fields("tmf09"), strExc(0) & "/" & Mid(rsAD.Fields("tmf09"), 1, 16), Replace(rsAD.Fields("tmf10"), "01.JPG", "00.JPG")) = False Then
            MsgBox strExc(1) & "：更名失敗" & rsAD.Fields("tmf03")
         Else
            strSql = "update tmqappfile set tmf03='00', tmf09=replace(tmf09,'01.JPG','00.JPG'), tmf10=replace(tmf10,'01.JPG','00.JPG') where tmf01='" & rsAD.Fields("tmf01") & "' and tmf02='" & rsAD.Fields("tmf02") & "' and tmf03='" & rsAD.Fields("tmf03") & "' "
            cnnConnection.Execute strSql
         End If
         rsAD.MoveNext
      Loop
   End If
   
   MsgBox "OK!"
   
   Set rsAD = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Me.ActiveControl <> txt1(0) And KeyCode = vbKeyReturn Then
     KeyCode = 0
  End If
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
Dim LrArray As Variant
    
'Modified by Lydai 2021/10/01 Top = 2760 => 3120
tmpPic.Left = 1920: tmpPic.Top = 3120

Me.Width = 5780

MoveFormToCenter Me
m_AttachPath = App.path & "\" & strUserNum
If Dir(m_AttachPath, vbDirectory) = "" Then
   MkDir m_AttachPath
End If
'Remove by Lydia 2016/05/26
''讀取前次設定路徑
'strLoadPath = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
'If strLoadPath = "" Then
'   strLoadPath = PUB_Getdesktop
'End If

iList = ""
ChkS1.Value = 0
ChkS2.Value = 0
FormReset

'Added by Lydia 2024/09/25
If strSrvDate(1) >= 查名單網中系統平行測試 Then
   strExc(0) = "select * from tmqappsumr"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "請先補上查名人員統計表！", vbExclamation + vbOKOnly, "測試階段"
      Unload Me
      Exit Sub
   End If
End If
'end 2024/09/25

'從查覆區來
If mApNoList <> "" Then
   ChkS1.Visible = False
   ChkS2.Visible = False
   LrArray = Split(mApNoList, ",")
   iLr = 0
   If ShowRecData(LrArray(iLr)) = False Then
      MsgBox "查無資料!", vbCritical, "查名單查詢"
      Unload Me
      Exit Sub
   End If
   cmdok(0).Enabled = False
   'textService.MaxLength = 80 'Added by Lydia 2021/10/19 原本80個字 'Mark by Lydia 2024/07/17
Else
    'Added by Lydia 2021/10/19 3519組群圖形查名單之指定商品/服務欄位,限制字元數為40個字元(含)以下(原本80個字),以限制委查人委查特定零售服務的個數
    'textService.MaxLength = 40 'Mark by Lydia 2024/07/17
End If

If txt1(0) = "" Then
   'Modified by Lydia 2016/04/28
   If stKeyUser <> "" Then
      txt1(0) = stKeyUser
   Else
      txt1(0) = strUserNum
   End If
End If
textCUID.BackColor = &H8000000F

'Modified by Lydia 2017/01/20 委查人的部門
'ClsPDGetStaff txt1(0).Text, strTemp, strTemp1
strTemp = GetStaffName(txt1(0).Text, False, , strTemp1)
lblname.Caption = strTemp
lblname.Tag = strTemp1 'Added by Lydia 2017/01/20

'Added by Lydia 2018/04/01 文字查名單最大組群個數(分割原則)
If strSrvDate(1) <= "20180402" Then
     cInP = 6
Else
     cInP = 3
End If

'Mark by Lydia 2022/07/07 已執行過,先保留
'If Pub_StrUserSt03 = "M51" Then cmdTrans.Visible = True 'Added by Lydia 2021/10/01

'Added by Lydia 2024/05/09
If Pub_StrUserSt03 = "M51" Then
   'cmdApp(0).Visible = True
   'cmdApp(1).Visible = True
   'cmdApp(2).Visible = True
   cmdChk1131220.Visible = True
End If

'Added by Lydia 2024/07/17
Set nFrm = Forms(0).GetForm("frm090132")
If Not nFrm Is Nothing And (cmdok(0).Enabled = True Or cmdGrp.Caption = "顯示") Then
   cmdGrp.Visible = True
Else
   cmdGrp.Visible = False
End If
'end 2024/07/17


End Sub

Private Sub Form_Unload(Cancel As Integer)
    '刪除無用的查名附件和釋放最新申請編號
    Call ReturnApp(lblAutoNo.Caption, bolPack)

    Set frm090126 = Nothing
    bolChkDept = False 'Added by Lydia 2017/06/08
    
    If TypeName(m_PrevForm) <> "Nothing" Then
        Select Case UCase(TypeName(m_PrevForm))
            'Modified by Lydia 2025/04/30 +"FRM090127_1"
            Case "FRM090127", "FRM090127_1" '待查區/查覆區/覆核區
                If InStr(m_PrevForm.Caption, "查覆區") > 0 Then
                   m_PrevForm.txtField(6).Text = "0" '預設全部,回查覆區後,是否要分狀態
                Else
                   m_PrevForm.txtField(6).Text = "1"
                End If
                If m_PrevForm.QueryData = False Then
                End If
            Case "FRM090801", "FRM090801_NEW" '國內接洽單 Add By Sindy 2022/9/16 +, "FRM090801_NEW"
                If iList <> "" Then
                  'Modified by Lydia 2016/04/18 回傳查名內容
                  ' m_PrevForm.CmdTMQ.Tag = iList
                  ' Call m_PrevForm.QueryTMQ
                   PubShowNextData
                End If
        End Select
        iList = ""
        m_PrevForm.Show
    End If
    mApNoList = ""
    
    Set nFrm = Nothing 'Added by Lydia 2024/11/08
    
    Set m_PrevForm = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer, Optional ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   Else
        If IsNull(rsSrcTmp.Fields("TQA10")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("TQA10")) = False Then
              strCName = GetStaffName(rsSrcTmp.Fields("TQA10"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("TQA11")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("TQA11")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("TQA11"))
              strCDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("TQA12")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("TQA12")) = False Then
              strTemp = rsSrcTmp.Fields("TQA12")
              strCTime = Format(strTemp, "00:00")
           End If
        End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime
              
End Sub

Private Sub Opt1_Click(Index As Integer)
If opt1(0).Value = True Then
   Call SetInput(1)
   If txtUnicode(1).Enabled = True And mApNoList = "" Then txtUnicode(1).SetFocus
ElseIf opt1(1).Value = True Then
   Call SetInput(2)
   If cmdPic.Enabled = True And mApNoList = "" Then cmdPic.SetFocus
End If

End Sub

'檢查委查組群是否有重覆
'Modified by Lydia 2021/10/01 TextBox => Control
Private Function Check_ClassDouble(ByRef textGrp As Control) As Boolean
Dim StrArray As Variant
Dim i As Integer
Dim j As Integer
Dim strGrp As String
Dim rsMe As New ADODB.Recordset

   StrArray = ""
   If Len(textGrp) <> 0 Then
      StrArray = Split(textGrp, ",")
      strGrp = "-"
      For i = 0 To UBound(StrArray)
         If StrArray(i) <> "" And (Len(StrArray(i)) <> 4 Or IsNumeric(StrArray(i)) = False) Then
            Check_ClassDouble = True
            MsgBox "委查組群格式輸入錯誤!!!", vbCritical
            Exit Function
         End If
         For j = i + 1 To UBound(StrArray)
            If StrArray(i) = StrArray(j) Then
               Check_ClassDouble = True
               MsgBox "委查組群重覆輸入" & StrArray(i) & "，請查明再輸!", vbCritical
               Exit Function
            End If
         Next j
         If strGrp = "-" Then
            strGrp = Mid(StrArray(i), 1, 2)
         End If
         '文字限同類
         If opt1(0).Value = True And strGrp <> Mid(StrArray(i), 1, 2) Then
               Check_ClassDouble = True
               MsgBox "委查組群必須同一類，請查明再輸!", vbCritical
               Exit Function
         End If
         '檢查不可存在於組群刪除資料檔
         If rsMe.State <> adStateClosed Then rsMe.Close
         Set rsMe = Nothing
         rsMe.CursorLocation = adUseClient
         rsMe.Open "Select * From ClassDelete Where CD01='" & StrArray(i) & "'", cnnConnection, adOpenStatic, adLockReadOnly
         If rsMe.RecordCount > 0 Then
            Check_ClassDouble = True
            MsgBox StrArray(i) & "為已刪除的組群，輸入錯誤!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
         If rsMe.State <> adStateClosed Then rsMe.Close
         Set rsMe = Nothing
         
          strExc(1) = ""
          If opt1(0).Value = True Then
              strExc(1) = "select tmqc01 from tmqclass where length(tmqc01)=2 and tmqc01=" & CNULL(Mid("" & StrArray(i), 1, 2))
          ElseIf opt1(1).Value = True Then
              strExc(1) = "select tmqc01 from tmqclass where tmqc01=" & CNULL("" & StrArray(i))
          End If
          If strExc(1) <> "" Then
              intI = 1
              Set rsMe = ClsLawReadRstMsg(intI, strExc(1))
              If intI = 0 Then
                  MsgBox "組群 " & IIf(opt1(0).Value = True, Mid(StrArray(i), 1, 2), StrArray(i)) & " 查無資料!", vbCritical
                  Check_ClassDouble = True
                  Exit Function
              End If
          End If
      Next i
      If UBound(StrArray) = 0 Then
         If (Len(StrArray(0)) < 1 Or Len(StrArray(0)) > 4) Or IsNumeric(StrArray(0)) = False Then
            Check_ClassDouble = True
            MsgBox "委查組群格式輸入錯誤!!!", vbCritical
            Exit Function
         End If
      End If
   End If
End Function

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
'Mark by Lydia 2016/10/28 受win7輸入法影響,不切換輸入法
'If Index < 2 Or Index > 3 Then
'   CloseIme
'End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
       Case 0, 6, 7, 8, 9
           KeyAscii = UpperCase(KeyAscii)
       Case Else
    End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim tmpArr As Variant
Dim i As Integer
Dim tmpGrp As String
Dim rsMe As New ADODB.Recordset
Dim bolSpec As Boolean

   Select Case Index
      Case 0
         If txt1(Index) <> Empty Then
            'Modified by Lydia 2017/01/20 委查人的部門
            'If ClsPDGetStaff(txt1(0).Text, strTemp, strTemp1) Then
            strTemp = GetStaffName(txt1(0).Text, False, , strTemp1)
            lblname.Tag = ""
            If strTemp <> "" Then
               lblname.Caption = strTemp
               lblname.Tag = strTemp1
            Else
               txt1(0).SetFocus
               Cancel = True
            End If
         End If
      Case 1
         'Modified by Lydia 2024/07/17 + And txt1(1).Locked = False
         If txt1(Index) <> Empty And txt1(1).Locked = False Then
            txt1(1).Text = Replace(txt1(1).Text, ".", ",") '組群間隔置換為","
            strExc(4) = PUB_RepToOneSpace(PUB_StringFilter(txt1(1).Text))   '清除字串中的enter & 清除連續空白
            txt1(1).Text = IIf(Right(strExc(4), 1) = ",", Mid(strExc(4), 1, Len(strExc(4)) - 1), strExc(4))
            'Added by Lydia 2024/07/17 改成3519組群輸入
            If InStr(txt1(1), "3519") > 0 Then
               'Modified by Lydia 2024/07/18 避免智權人員誤解，直接拿掉組群已輸入的3519---嘉雯
               'MsgBox "3519組群請點選輸入按鈕 ！", vbExclamation, "操作錯誤！"
               'Txt1(1).SetFocus
               'txt1_GotFocus 1
               'Exit Sub
               txt1(1).Text = Replace(Replace(Replace(txt1(1), ",3519", ""), "3519,", ""), "3519", "")
            End If
            'end 2024/07/17
            'Added by Lydia 2024/10/04 關閉證明標章「9999」代碼
            If InStr(txt1(1), "9999") > 0 Then
               MsgBox "不可輸入組群代碼9999!!!", vbExclamation + vbOKOnly, "輸入檢查"
                txt1(1).SetFocus
                Cancel = True
                Exit Sub
            End If
            'end 2024/10/04
            'Modified by Lydia 2024/07/17
            'If Check_ClassDouble(Txt1(1)) Then
            If Pub_ChkTMQCisExist(Me.Name, txt1(1), "2", IIf(opt1(0).Value = True, "W", "P")) = False Then
                txt1(1).SetFocus
                Cancel = True
                Exit Sub
            Else
                '指定商品/服務,限組群3519
                'If InStr(txt1(1).Text, "3519") > 0 Then bolSpec = True 'Mark by Lydia 2024/07/17 改成3519組群輸入
            End If
            txt1(1).Tag = txt1(1).Text
            'Mark by Lydia 2024/07/17 改成3519組群輸入
            'If bolSpec = True Then
            '   lblSA(1).Visible = True: textService.Visible = True
            '   If textService.Text = "" And tmpS3519 <> "" Then textService.Text = tmpS3519
            'Else
            '   lblSA(1).Visible = False: textService.Visible = False
            '   textService.Text = ""
            'End If
            'end 2024/07/17
            
            '輸入組群後,給予申請編號
            If lblAutoNo.Caption = "" Then
               'Modified by Lydia 2024/07/17 改成模組
                'strExc(1) = AccAutoNo("HM", 4, GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2)))  '取得自動編號
                'lblAutoNo.Caption = strExc(1)
                'strExc(0) = AccSaveAutoNo("HM", Right(strExc(1), 4), GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2))) '回寫acc1r0
                ''保留上次輸入資料
                'If (ChkS1.Value = 1 Or ChkS2.Value = 1) And tmpReNo <> "" And haveKey <> "" Then
                '   If AttachFileRedo(tmpReNo, haveKey) = False Then
                '   End If
                'End If
                Call GetKeyNo
                'end 2024/07/17
            End If
         End If
      'Added by Lydia 2016/07/18 指定商品/服務從單行60字改為兩行80字
      Case 3
           txt1(Index).Text = PUB_RepToOneSpace(PUB_StringFilter(txt1(Index).Text))
      Case Else
   End Select
   'Added by Lydia 2016/04/22 檢查長度
   If Not CheckLengthIsOK(txt1(Index), txt1(Index).MaxLength) Then
      txt1(Index).SetFocus
      Cancel = True
   End If
End Sub
'複製前一張單據的查名附件
Private Function AttachFileRedo(ByVal RAno As String, ByVal RKey As String) As Boolean
Dim aj As Integer
Dim ArrJ As Variant
Dim strJ As String, strT As String

AttachFileRedo = False
    ArrJ = Split(RKey, ",")
    For aj = 0 To UBound(ArrJ)
       If ArrJ(aj) <> "" Then
          If AttachFileGet(RAno, ArrJ(aj), True, strJ, strT) Then
             If PUB_TMQAFileSave(lblAutoNo, TMQ_附件F02, Trim(ArrJ(aj)), TMQ_附件F04, strT, strJ) = False Then
                Exit Function
             End If
          End If
       End If
    Next aj
    
AttachFileRedo = True

End Function

'清空欄位
Private Sub FormReset()
Dim tmpStr As String
    
    If (ChkS1.Value = 1 Or ChkS2.Value = 1) And bolRe = True Then
        tmpReNo = lblAutoNo.Caption
        
        lblAutoNo.Caption = ""
        strExc(1) = txt1(1).Text
        strExc(2) = textCName.Text
        tmpStr = textService.Text
        For Each oText In txt1
          If oText.Index > 0 Then
             oText.Text = ""
          End If
        Next
        'Added by Lyda 2021/10/01 預設清空
        textCName.Text = ""
        textService.Text = ""
        'end 2021/10/01
        
        cmdok(0).Enabled = True: cmdok(1).Enabled = True
        '不保留Unicode
        If ChkS1.Value = 0 Then
           txtUnicode(1).Text = "": txtUnicode(2).Text = ""
        End If
        
        chk1.Value = 0
        
        '保留指定商品/服務
        'If tmpStr <> "" Then textService.Text = tmpStr 'Mark by Lydia 2024/10/08 不保留3519組群
        '保留客戶名稱
        textCName.Text = strExc(2)
        
        '保留資料
        If ChkS1.Value = 1 Then

        End If
        
        '保留組群
        If ChkS2.Value = 1 Then
           txt1(1).Text = strExc(1)
           haveKey = ""
            Clipboard.Clear
            Set tmpImg.Picture = Nothing
            Set tmpKeyImg1.Picture = Nothing
            Set tmpKeyImg2.Picture = Nothing
            Set G_SeekPicColor.Picture = LoadPicture("")
        End If
        Call UpdateCUID(0)
        bolPack = False
        txtEna True
               
    Else
        bolRe = False: tmpReNo = ""
        lblAutoNo.Caption = ""
        tmpStr = textService.Text
        For Each oText In txt1
          If oText.Index > 0 Then
             oText.Text = ""
          End If
        Next
        'Added by Lyda 2021/10/01 預設清空
        textCName.Text = ""
        textService.Text = ""
        'end 2021/10/01
        
        cmdok(0).Enabled = True: cmdok(1).Enabled = True
        txtUnicode(1).Text = "": txtUnicode(2).Text = ""
        '保留指定商品/服務
        If tmpStr <> "" Then textService.Text = tmpStr
        
        opt1(0).Value = False: opt1(1).Value = False
        chk1.Value = 0
        Call SetInput(0)
        Call UpdateCUID(0)
        bolPack = False
        txtEna True
        Clipboard.Clear
        Set tmpImg.Picture = Nothing
        Set tmpKeyImg1.Picture = Nothing
        Set tmpKeyImg2.Picture = Nothing
        Set G_SeekPicColor.Picture = LoadPicture("")
    End If
End Sub

'控制讀寫
Private Sub txtEna(bolUpd As Boolean)
   Dim bolP As Boolean
   
   If bolUpd = True Then
      bolP = False
   Else
      bolP = True
   End If
   
   For Each oText In txt1
      oText.Locked = bolP
   Next
   
   txtUnicode(1).Locked = bolP
   txtUnicode(2).Locked = bolP
   cmdKey(1).Enabled = bolUpd:  cmdKey(2).Enabled = bolUpd
   cmdKey(3).Enabled = bolUpd:  cmdKey(4).Enabled = bolUpd
   cmdPic.Enabled = bolUpd: cmdKey(0).Enabled = bolUpd
   Frame2.Enabled = bolUpd '查名類別
   Frame3.Enabled = bolUpd '已收文
End Sub
'設定可輸入項目
Private Sub SetInput(ByVal aKind As String)
    'Mark by Lydia 2024/07/17 改成3519組群輸入
    'If InStr(1, txt1(1), "3519") = 0 Then
    '   lblSA(1).Visible = False: textService.Visible = False
    'Else
    '   lblSA(1).Visible = True: textService.Visible = True
    'End If
    'end 2024/07/17
    Select Case aKind
       Case "0" '未選擇
        cmdKey(1).Visible = False: cmdKey(2).Visible = False
        cmdKey(3).Visible = False: cmdKey(4).Visible = False
        tmpKeyPic1.Visible = False: tmpKeyPic2.Visible = False
        txtUnicode(1).Visible = False: txtUnicode(2).Visible = False
        cmdPic.Visible = False: cmdKey(0).Visible = False
        tmpPic.Visible = False
        cmdHelp.Visible = False
        lblSA(2).Visible = False
        lblSA(3).Visible = False
        lblSA(4).Visible = False 'Added by Lydia 2016/04/27
       Case "1" '文字查名
        '文字1
        tmpKeyPic1.Visible = True: txtUnicode(1).Visible = True
        lblSA(2).Visible = True: cmdKey(1).Visible = True: cmdKey(3).Visible = True
        '文字2
        tmpKeyPic2.Visible = True: txtUnicode(2).Visible = True
        lblSA(3).Visible = True: cmdKey(2).Visible = True: cmdKey(4).Visible = True
        lblSA(4).Visible = True 'Added by Lydia 2016/04/27
        cmdHelp.Visible = True: cmdHelp.Top = cmdKey(0).Top
        cmdPic.Visible = False: cmdKey(0).Visible = False
        tmpPic.Visible = False
        
       Case "2"  '圖形查名
        cmdKey(1).Visible = False: cmdKey(2).Visible = False
        cmdKey(3).Visible = False: cmdKey(4).Visible = False
        tmpKeyPic1.Visible = False: tmpKeyPic2.Visible = False
        txtUnicode(1).Visible = False: txtUnicode(2).Visible = False
        cmdPic.Visible = True: cmdKey(0).Visible = True
        tmpPic.Visible = True
        cmdHelp.Visible = False
        lblSA(2).Visible = False
        lblSA(3).Visible = False
        lblSA(4).Visible = False 'Added by Lydia 2016/04/27
    End Select
    
    'Added by Lydia 2024/09/26 檔案限制為JPG檔
    If strSrvDate(1) >= 查名單網中系統平行測試 Then
       cmdKey(0).Visible = False
       cmdKey(3).Visible = False
       cmdKey(4).Visible = False
    End If
End Sub
'插入圖片
Private Sub Do_Picture(ByVal tInx As Integer)
Dim mStr As String
   
    frmPic001.oCP01 = lblAutoNo.Caption
    frmPic001.oCP02 = "0"
    frmPic001.oCP03 = CStr(tInx) '0~2
    frmPic001.oCP04 = "0"
    
     '重新存檔才覆蓋
    If tInx = 0 Then
        Set frmPic001.oPic = G_SeekPicColor
        Set frmPic001.oImg = tmpImg
        mStr = "選擇圖片"
    Else
        If tInx = 1 Then
            Set frmPic001.oPic = G_SeekPicColor
            Set frmPic001.oImg = tmpKeyImg1
        Else
            Set frmPic001.oPic = G_SeekPicColor
            Set frmPic001.oImg = tmpKeyImg2
        End If
        mStr = "插入影像" & CStr(tInx)
    End If

    Set frmPic001.UpForm = Me
    frmPic001.oRtPic = False
    frmPic001.m_TMQ = tInx
    frmPic001.cmdok(4).Visible = False
    frmPic001.cmdok(5).Visible = False
    frmPic001.cmdok(6).Visible = False
    frmPic001.cmdok(7).Visible = False
    frmPic001.cmdok(2).Caption = "存檔(&O)"
    frmPic001.cmdok(3).Caption = "離開(&X)"
    frmPic001.Label11.Caption = mStr
    frmPic001.cmdok(0).Left = frmPic001.cmdok(0).Left - 250
    frmPic001.cmdok(1).Left = frmPic001.cmdok(1).Left - 250
    frmPic001.cmdok(2).Left = frmPic001.cmdok(2).Left - 250
    frmPic001.cmdok(3).Left = frmPic001.cmdok(3).Left - 250
    frmPic001.Width = 3800
    MoveFormToCenter frmPic001
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
    Unload frmpic002
    frmPic001.Show vbModal
    
    '重置圖片
    Select Case tInx
        Case 0
           If AttachFileGet(lblAutoNo.Caption, TMQ_AkindPic) Then
           End If
        Case 1
           If AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord1) Then
           End If
        Case 2
           If AttachFileGet(lblAutoNo.Caption, TMQ_AkindWord2) Then
           End If
    End Select
End Sub
'存檔
Private Function FormSave() As Boolean
Dim rsWrite As New ADODB.Recordset
Dim btHead(1) As Byte, btHead2(1) As Byte
Dim btTemp() As Byte, btTemp2() As Byte
Dim lLength As Long
Dim strFileName As String
Dim mSQL As String

On Error GoTo ErrHnd
    
   'Added by Lydia 2021/10/01 用日期控制不用經過二進位處理存入TQA07-TQA08，直接存入TQA13-TQA14
   If strSrvDate(1) >= Form20上線日 Then

   Else
   'end 2021/10/01
        '先將Unicode存成二進位檔
         UnicodeSave
         
         '載入二進位檔
         If txtUnicode(1).Text <> "" Then
            file_num1 = FreeFile
            strFileName = m_AttachPath & "\unicode1.txt"
            lLength = FileLen(strFileName) - 2
            ReDim btTemp(lLength) As Byte
            
            Open strFileName For Binary As #file_num1
            Get #file_num1, , btHead
            Get #file_num1, , btTemp
            Close #file_num1
         End If
         If txtUnicode(2).Text <> "" Then
            file_num2 = FreeFile
            strFileName = m_AttachPath & "\unicode2.txt"
            lLength = FileLen(strFileName) - 2
            ReDim btTemp2(lLength) As Byte
            
            Open strFileName For Binary As #file_num2
    
            Get #file_num2, , btHead2
            Get #file_num2, , btTemp2
            Close #file_num2
         End If
   End If 'Added by Lydia 2021/10/01

     cnnConnection.BeginTrans
       If rsWrite.State <> adStateClosed Then rsWrite.Close
       Set rsWrite = Nothing
        rsWrite.CursorLocation = adUseClient
        rsWrite.Open "select * from TMQApp where 1=0", cnnConnection, adOpenStatic, adLockOptimistic
        
        rsWrite.AddNew
        rsWrite.Fields(0).Value = lblAutoNo.Caption
        rsWrite.Fields(1).Value = Trim(txt1(0).Text)
        rsWrite.Fields(2).Value = Trim(txt1(1).Text)
        rsWrite.Fields(3).Value = Trim(textCName)
        rsWrite.Fields(4).Value = Trim(textService)
        rsWrite.Fields(5).Value = IIf(opt1(0).Value = True, "1", "2")
        'Added by Lydia 2021/10/01 用日期控制不用經過二進位處理存入TQA07-TQA08，直接存入TQA13-TQA14
        If strSrvDate(1) >= Form20上線日 Then
            rsWrite.Fields(12).Value = Trim(txtUnicode(1).Text)
            rsWrite.Fields(13).Value = Trim(txtUnicode(2).Text)
        Else
        'end 2021/10/01
            If txtUnicode(1).Text <> "" Then
              If txt1(4).Text <> txtUnicode(1).Text Then
                 rsWrite.Fields(6).Value = btTemp
              Else
                 rsWrite.Fields(12).Value = Trim(txt1(4).Text)
              End If
            End If
            If txtUnicode(2).Text <> "" Then
              If txt1(5).Text <> txtUnicode(2).Text Then
                 rsWrite.Fields(7).Value = btTemp2
              Else
                 rsWrite.Fields(13).Value = Trim(txt1(5).Text)
              End If
            End If
        End If 'Added by Lydia 2021/10/01
        rsWrite.Fields(8).Value = Null  '查覆完成日期
        rsWrite.Fields(9).Value = strUserNum
        rsWrite.Fields(10).Value = strSrvDate(1)
        rsWrite.Fields(11).Value = Left(Format(ServerTime, "000000"), 4)
        If chk1.Value = 1 Then rsWrite.Fields(14).Value = "Y"
        rsWrite.Fields(15).Value = Trim(txt1(6).Text)
        rsWrite.Fields(16).Value = Trim(txt1(7).Text)
        rsWrite.Fields(17).Value = Trim(txt1(8).Text)
        rsWrite.Fields(18).Value = Trim(txt1(9).Text)
        rsWrite.UPDATE
        'Move by Lydia 2016/06/23 從下方移過來

        '刪除不用的圖檔或附件
        'Added by Lydia 2016/06/23
        'Remove by Lydia 2016/07/07
        'If strSrvDate(1) < TMQFileFTP Then
        '    mSQL = " DELETE FROM TMQFILE WHERE TQF01='" & lblAutoNo.Caption & "' and TQF02='" & TMQ_附件F02 & "' and "
        '    mSQL = mSQL & IIf(Opt1(0).Value = True, "TQF03='" & TMQ_AkindPic & "'", "TQF03 <> '" & TMQ_AkindPic & "'") & " and TQF04='" & TMQ_附件F04 & "'"
        '     cnnConnection.Execute mSQL, intI 'Move by Lydia 2016/06/23 從下方移過來
        'Else
            If PUB_TMQAFileDel(lblAutoNo.Caption, TMQ_附件F02, IIf(opt1(0).Value = True, TMQ_AkindPic, " <> " & CNULL(TMQ_AkindPic)), TMQ_附件F04) Then
            End If
        'End If
        cnnConnection.CommitTrans
                
 Screen.MousePointer = vbDefault
 
 Call UpdateCUID(1, rsWrite)
 FormSave = True
 Exit Function
 
ErrHnd:
   DataErrorMessage (3)
   
End Function
'查名單送出=>分派查名人
Private Function ProcessPack() As Boolean
Dim midSql As String, exSQL As String
Dim tmpArr As Variant
Dim exsqlArr As Variant
Dim i As Integer, tmpWorkday As Integer
Dim inX As Integer, Inputtm As String, InputWDay As String
Dim tTmq10 As String, tTmq01 As String
Dim inC As Integer '中文Keyword
Dim inE As Integer '英文Keyword
Dim tmpClass As String, tmpQC02  As String
Dim strTmq05 As String 'Added by Lydia 2017/01/20 收件分發日期
Dim strErrNoList As String  'Added by Lydia 2017/06/23 記錄未正確分查名人的單號
Dim chkAllStatus As String 'Added by Lydia 2018/05/25 內商查名單分單狀態：若查名中心聯絡開始不分單將狀態改為N，恢復分單將狀態改為Y
Dim iRound As Integer 'Added by Lydia 2024/07/17

    '判斷文字查詢的筆數
    haveKey = ""
    If opt1(0).Value = True Then
        For i = 1 To 2
            If PUB_TMQFileIsExist(lblAutoNo.Caption, TMQ_附件F02, Format(i, "0"), TMQ_附件F04) Then
                If txtUnicode(i).Text = "" Then inC = inC + 1
                haveKey = haveKey & Format(i, "0") & ","
            End If
            '文字可能以替代字輸入,所以筆數依文字欄位來判斷
            Call PUB_CountTxtNEC(inE, inC, txtUnicode(i))
        Next i
    Else
        haveKey = "0,"
    End If
    NowTMQno = ""
    
    strErrNoList = "" 'Added by Lydia 2017/06/23
    chkAllStatus = Pub_GetSpecMan("內商查名單分單狀態") 'Added by Lydia 2018/05/25
    'Modified by Lydia 2025/05/21 若一般組群前2碼非35類，一般組群不可與3519組群合併，分成2張單---嘉雯
    'For iRound = 1 To 1   'Added by Lydia 2024/07/17 改成3519組群輸入:同時輸入組群和「3519指定商品/服務」，3519組群可與其他組群同時填在一張單,由系統以3個為一組分單
    For iRound = 1 To IIf(txt1(1) <> "" And textService <> "" And Left(txt1(1), 2) <> "35", 2, 1)
       inX = 1 '明細流水號(TQD04)
       tmpArr = Empty: exsqlArr = Empty
       tmpClass = ""
       'Modified by Lydia 2025/05/21 若一般組群前2碼非35類，一般組群不可與3519組群合併，分成2張單
       'strExc(1) = IIf(Trim(txt1(1)) = "", "", "," & Trim(txt1(1))) & IIf(Trim(textService) = "", "", "," & Trim(textService))
       'tmpArr = Split(Mid(strExc(1), 2), ",")
       'end 2024/06/26
       If iRound = 1 And Trim(txt1(1)) <> "" Then
          If Left(Trim(txt1(1)), 2) <> "35" Then
             strExc(1) = Trim(txt1(1))
          Else
             strExc(1) = Trim(txt1(1)) & IIf(textService <> "", "," & Trim(textService), "")
          End If
       Else
          strExc(1) = textService
       End If
       tmpArr = Split(strExc(1), ",")
       'end 2025/05/21

              
       If Inputtm = "" Then 'Added by Lydia 2024/07/17
          Inputtm = Left(Format(ServerTime, "000000"), 4) 'Move by Lydia 2017/01/20
       End If
       For i = 0 To UBound(tmpArr)
          If tmpArr(i) <> "" Then
             'Move by Lydia 2017/01/20
             'Inputtm = Left(Format(ServerTime, "000000"), 4)
             'Added by Lydia 2017/01/20 收件分發日期
             'Modified by Lydia 2018/05/25 增加控制內商查名單分單是否可分派查名人員
             'If Val(Inputtm) > 1800 Then
             If Val(Inputtm) > 1800 Or UCase(chkAllStatus) = "N" Then
                 strTmq05 = ""
             Else
                 strTmq05 = strSrvDate(1)
             End If
             'end 2017/01/20
             'Modified by Lydia 2017/11/07 改成共用模組
'             '預設工作天
'             If Opt1(0).Value = True Then
'                tmpWorkday = 2 '文字
'             Else
'                tmpWorkday = 3 '圖形
'                '本數決定工作天數
'                tmpClass = PUB_GetTMQClass(2, 0, 0, 1, tmpArr(i), , tmpQC02)
'                If Val(tmpQC02) >= 14 Then tmpWorkday = 4
'             End If
'             '大於13:30,期限+1天
'             If Val(Inputtm) > 1330 Then
'                InputWDay = CompWorkDay(tmpWorkday + 1, strSrvDate(1), 0)
'             Else
'                InputWDay = CompWorkDay(tmpWorkday, strSrvDate(1), 0)
'             End If
             'Modified by Lydia 2018/05/25 +狀態
             InputWDay = PUB_GetNewTmq06(IIf(opt1(0).Value = True, 1, 2), Trim(tmpArr(i)), strSrvDate(1), Inputtm, chkAllStatus)
             'end 2017/11/07
             
             '文字6個類組一張查名單
             If opt1(0).Value = True Then
                'Modified by Lydia 2018/04/01 改成變數
                'If (i + 1) Mod 6 = 1 Then
                If (i + 1) Mod cInP = 1 Then
                   tTmq01 = AccAutoNo("H", 4, GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2)))
                   strExc(1) = AccSaveAutoNo("H", Right(tTmq01, 4), GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2))) '自動編號 -回寫acc1r0
                   NowTMQno = NowTMQno & tTmq01 & ","
                End If
                'Modified by Lydia 2018/04/01 改成變數
                'If inX > 6 Then
                If inX > cInP Then
                   inX = 1: tmpClass = tmpArr(i)
                Else
                   tmpClass = tmpClass & "," & tmpArr(i) '多筆組群
                End If
             Else '圖形1個類組一張查名單
                tTmq01 = AccAutoNo("H", 4, GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2))) '取得自動編號
                strExc(1) = AccSaveAutoNo("H", Right(tTmq01, 4), GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2))) '自動編號 -回寫acc1r0
                NowTMQno = NowTMQno & tTmq01 & ","
                inX = 1: tmpClass = tmpArr(i)
             End If
             
             '新增查名單(分單)
             '因為文字查詢筆數決定查名排班分類,所以查名單主檔在讀取的最後一筆寫入
             'Modified by Lydia 2018/04/01 改成變數
             'If (inX = 6 Or i = UBound(tmpArr)) Or opt1(1).Value = True Then
             If (inX = cInP Or i = UBound(tmpArr)) Or opt1(1).Value = True Then
                 If Left(tmpClass, 1) = "," Then tmpClass = Mid(tmpClass, 2, Len(tmpClass) - 1)
                 
                 'TMQ01~TMQ06
                 'Added by Lydia 2016/04/06 已收文案件輸入案號後，視做先收文後查名(TMQ21)
                 'Modified by Lydia 2017/01/20 若查名單是在下午六點以後輸入，系統不分發tmq05=null
                 'midSql = "INSERT INTO trademarkquery(TMQ01,TMQ02,TMQ03,TMQ04,TMQ05,TMQ06,TMQ07,TMQ08,TMQ09,TMQ10,TMQ12,TMQ13,TMQ14,TMQ18,TMQ21) " & _
                        "VALUES ('" & tTmq01 & "','" & Trim(txt1(0).Text) & "','" & tmpClass & "'," & strSrvDate(1) & "," & strSrvDate(1) & "," & InputWDay
                 midSql = "INSERT INTO trademarkquery(TMQ01,TMQ02,TMQ03,TMQ04,TMQ05,TMQ06,TMQ07,TMQ08,TMQ09,TMQ10,TMQ12,TMQ13,TMQ14,TMQ18,TMQ21) " & _
                        "VALUES ('" & tTmq01 & "','" & Trim(txt1(0).Text) & "','" & tmpClass & "'," & strSrvDate(1) & "," & CNULL(strTmq05) & "," & InputWDay
                 'TMQ07~TMQ09 查詢筆數
                 If opt1(0).Value = True Then
                    midSql = midSql & "," & inC * inX & "," & inE * inX & ",null"
                 Else
                    midSql = midSql & ",null,null,1"
                 End If
                 '分派查名人
                 'Modified by Lydia 2017/01/20 若查名單是在下午六點以後輸入，系統不分發，即不顯示查名人。俟系統在凌晨根據查名人員出缺勤狀況，再分發查名單，即此時可顯示查名人。
                 'tTmq10 = PUB_GetTMQUserPos(True, "1", inC * inX, inE * inX, IIf(Opt1(1).Value = True, 1, 0), IIf(Opt1(1).Value = True, tmpClass, "X111"))
                 If strTmq05 = "" Then
                     tTmq10 = ""
                 Else
                     tTmq10 = PUB_GetTMQUserPos(True, "1", inC * inX, inE * inX, IIf(opt1(1).Value = True, 1, 0), IIf(opt1(1).Value = True, tmpClass, "X111"))
                 End If
                 
                 'Added by Lydia 2016/04/06 +TMQ21
                 midSql = midSql & ",'" & tTmq10 & "','" & strUserNum & "'," & strSrvDate(1) & "," & Left(Format(ServerTime, "000000"), 4) & ",'" & lblAutoNo.Caption & "'," & CNULL(strCP09) & ")"
                 exSQL = exSQL & midSql & ";"
                 '更新當日拿單量
                 'Modified by Lydia 2017/06/23 若未分查名人,則不更新
                 If tTmq10 <> "" Then Call PUB_TMQtake(1, tTmq10, inC * inX, inE * inX, IIf(opt1(1).Value = True, 1, 0), IIf(opt1(1).Value = True, tmpClass, "X111"))
                 
                 'Added by Lydia 2017/06/23 記錄未正確分查名人的單號
                 If strTmq05 <> "" And tTmq10 = "" Then strErrNoList = strErrNoList & tTmq01 & ","
                 
                 'Added by Lydia 2016/04/06 已收文案件新增TS.menu 至卷宗區
                 If strCP09 <> "" Then
                    midSql = Trim(txt1(6)) & CStr(Val(txt1(7))) & IIf(txt1(8) <> "0" Or txt1(9) <> "00", "-" & txt1(8), "") & IIf(txt1(9) <> "00", "-" & txt1(9), "")
                    'Modified by Lydia 2016/04/25
                    'midSql = midSql & ".101." & tTmq01 & "." & TMQ_查名作業 & ".menu"
                    midSql = midSql & "." & strCP10 & "." & tTmq01 & "." & TMQ_查名作業 & ".menu"

                    midSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                             " values('" & strCP09 & "','" & midSql & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'Y')"
                    exSQL = exSQL & midSql & ";"
                    '增加查名代號記錄
                    midSql = PUB_GetTMQCaseMapNo(strCP09)
                    midSql = "insert into tmqcasemap(tqc01,tqc02,tqc03,tqc04,tqc05,tqc06)" & _
                             " values('" & midSql & "','" & strCP09 & "','" & tTmq01 & "','" & strUserNum & "'," & strSrvDate(1) & "," & Left(Format(ServerTime, "000000"), 4) & ") "
                    exSQL = exSQL & midSql & ";"
                    'Added by Lydia 2019/04/18 追加未完成的查名單, 取消原承辦期限和查名齊備日
                    midSql = "update CaseProgress SET CP48=null,CP143=0 WHERE CP09='" & strCP09 & "' and cp158=0 "
                    exSQL = exSQL & midSql & ";"
                    'end 2019/04/18
                 End If
             End If
             
             '新增查名單明細檔
             If opt1(0).Value = True Then
                If txtUnicode(1).Text <> "" Or InStr(haveKey, "1") > 0 Then '文字有輸入、插入圖片和PDF
                   midSql = "INSERT INTO TMQDetail(TQD01,TQD02,TQD03,TQD04,TQD05) VALUES ('" & lblAutoNo.Caption & "','" & tTmq01 & "','1'," & inX & ",'" & tmpArr(i) & "') "
                   exSQL = exSQL & midSql & ";"
                End If
                If txtUnicode(2).Text <> "" Or InStr(haveKey, "2") > 0 Then
                   midSql = "INSERT INTO TMQDetail(TQD01,TQD02,TQD03,TQD04,TQD05) VALUES ('" & lblAutoNo.Caption & "','" & tTmq01 & "','2'," & inX & ",'" & tmpArr(i) & "') "
                   exSQL = exSQL & midSql & ";"
                End If
             Else
                midSql = "INSERT INTO TMQDetail(TQD01,TQD02,TQD03,TQD04,TQD05) VALUES ('" & lblAutoNo.Caption & "','" & tTmq01 & "','0'," & inX & ",'" & tmpArr(i) & "') "
                exSQL = exSQL & midSql & ";"
             End If
    
             inX = inX + 1
          End If
       Next i
    Next iRound 'Added by Lydia 2024/07/17
    
On Error GoTo ErrHnd2
    cnnConnection.BeginTrans
       exSQL = Mid(exSQL, 1, Len(exSQL) - 1)
       exsqlArr = Split(exSQL, ";")
       inX = 1 '明細流水號(TQD04)
       For i = 0 To UBound(exsqlArr)
           cnnConnection.Execute exsqlArr(i), intI
       Next i
       'Added by Lydia 2016/06/29 + 是否送出
       cnnConnection.Execute "update TMQFILE set TQF13='Y' where TQF01='" & lblAutoNo.Caption & "' and TQF02='" & TMQ_附件F02 & "' and TQF04='" & TMQ_附件F04 & "' ", intI
    cnnConnection.CommitTrans
    
 Screen.MousePointer = vbDefault
 
    'Added by Lydia 2017/06/23
    If strErrNoList <> "" Then
        'Modified by Lydia 2025/08/22 因為8/21下午嘉雯為了先消耗查名人員現有查名單,所以將人員狀態全部設定N
               '若是內商查名單分單狀態=N，不會進入strErrNoList，所以調整通知直接給嘉雯
        'MsgBox "查名單沒有分派到查名人，請與電腦中心聯繫！", vbCritical
        'strExc(0) = "查名單號如下列：" & vbCrLf & strErrNoList & vbCrLf & String(30, "-") & vbCrLf & _
                    "Step 1　確認人員清單:請與嘉雯聯繫，確認查名人員維護是否正確；" & vbCrLf & vbCrLf & _
                    "Step 2　重新分查名人:請開啟商標委查作業的查名單維護，可以輸入委查日期或委查人等特定條件，或者不輸入條件來重新分查名人。" & vbCrLf
        'PUB_SendMail strUserNum, "A3034;83002", "", "查名單沒有分派到查名人", vbCrLf & strExc(0)
        strExc(1) = Pub_GetSpecMan("內商查名主管")
        strExc(0) = "查名單號如下列：" & vbCrLf & strErrNoList & vbCrLf & String(30, "-") & vbCrLf & _
                    "請確認查名人員維護是否正確。"
        PUB_SendMail strUserNum, strExc(1), "", "查名單沒有分派到查名人", vbCrLf & strExc(0)
    End If
    'end 2017/06/23
 ProcessPack = True
 Exit Function

ErrHnd2:
   MsgBox "查名單送出分割錯誤"
   cnnConnection.RollbackTrans
End Function
'寫入Unicode文字(暫存本機二進位檔)
Private Sub UnicodeSave()
Dim btHead(1) As Byte
Dim btTemp() As Byte
Dim p As String

    '先刪檔
    If Dir(m_AttachPath & "\unicode1.txt") <> "" Then Kill m_AttachPath & "\unicode1.txt"
    If Dir(m_AttachPath & "\unicode2.txt") <> "" Then Kill m_AttachPath & "\unicode2.txt"

    
    btHead(0) = 255
    btHead(1) = 254
    
    If txtUnicode(1).Text <> "" Then
        btTemp = txtUnicode(1).Text
        Open m_AttachPath & "\unicode1.txt" For Binary As #1
        Put #1, , btHead
        Put #1, , btTemp
        Close #1
    End If
    
    If txtUnicode(2).Text <> "" Then
        btTemp = txtUnicode(2).Text
        Open m_AttachPath & "\unicode2.txt" For Binary As #2
        Put #2, , btHead
        Put #2, , btTemp
        Close #2
    End If
End Sub

'選擇檔案
Private Sub AttachFileAdd(fID As Integer)
Dim stFileName As String
Dim sFile
Dim ii As Integer
Dim fs, f, s
Dim strFile As String
   
On Error GoTo ErrHnd

   stFileName = "*.PDF"
   
    With CommonDialog1
       .CancelError = True
       .FileName = stFileName
       .Filter = "All Files (*.PDF)|*.PDF"
       'Modified by Lydia 2016/05/26
       '.InitDir = strLoadPath
        If GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "") <> "" Then
           .InitDir = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
        Else
           .InitDir = PUB_Getdesktop
        End If
       .MaxFileSize = 3000
       .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
       .ShowOpen
        If .FileName <> "" Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modified by Lydia 2016/05/26 記錄路徑只到資料夾位置
            'SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", sFile(0)
            'strLoadPath = sFile(0)
            SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
            For ii = 0 To UBound(sFile)
              If InStr(CStr(sFile(ii)), "#") > 0 Then
                 MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                 Exit Sub
              End If
              If UCase(Right(CStr(sFile(ii)), 4)) <> UCase(".pdf") Then
                 MsgBox "只能插入PDF檔！"
                 Exit Sub
              End If
                               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               ElseIf f.Size > 5242880 Then
                     If MsgBox("檔案過大（容量超過5MB），確認是否要插入檔案？", vbYesNo, "警告") = vbNo Then
                        Exit Sub
                     End If
               End If
            Next

            If PUB_TMQAFileSave(lblAutoNo.Caption, TMQ_附件F02, Trim(fID), TMQ_附件F04, "PDF", stFileName) = True Then
                If fID = 0 Then
                   tmpImg.Top = 0: tmpImg.Left = 0
                   tmpImg.Height = tmpInsPDF.Height
                   tmpImg.Width = tmpInsPDF.Width
                   Set tmpImg.Picture = tmpInsPDF.Picture
                ElseIf fID = 1 Then
                      'Added by Lydia 2016/03/28
                      tmpKeyImg1.Top = 0: tmpKeyImg1.Left = 0
                      tmpKeyImg1.Height = tmpInsPDF.Height
                      tmpKeyImg1.Width = tmpInsPDF.Width
                      'end 2016/03/28
                      Set tmpKeyImg1.Picture = tmpInsPDF.Picture
                   Else
                      'Added by Lydia 2016/03/28
                      tmpKeyImg2.Top = 0: tmpKeyImg2.Left = 0
                      tmpKeyImg2.Height = tmpInsPDF.Height
                      tmpKeyImg2.Width = tmpInsPDF.Width
                      'end 2016/03/28
                      Set tmpKeyImg2.Picture = tmpInsPDF.Picture
                End If
            End If
        End If
    End With
    
    Exit Sub
  
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub
Private Function isCheckInput(cID As Integer) As Boolean
Dim bMsg As Boolean

isCheckInput = True
'Modified by Lydia 2024/07/19
'If lblAutoNo.Caption = "" Or TTxt1(1) = "" Then
'   MsgBox "請先輸入組群!", vbExclamation
If lblAutoNo.Caption = "" Or Trim(txt1(1) & textService) = "" Then
   MsgBox "請先輸入組群或3519組群!", vbExclamation
'end 2024/07/19
   txt1(0).SetFocus
   isCheckInput = False
   Exit Function
Else

   If PUB_TMQFileIsExist(lblAutoNo.Caption, TMQ_附件F02, Trim(cID), TMQ_附件F04) Then bMsg = True
   
   If bMsg = True Then
      If MsgBox("已有查名內容,是否覆蓋原有內容?", vbCritical + vbYesNo, "存檔") = vbYes Then
         Exit Function
      Else
         isCheckInput = False
      End If
   End If
End If

End Function

Private Sub txtUnicode_GotFocus(Index As Integer)
   txtUnicode(Index).SelStart = 0
   txtUnicode(Index).SelLength = Len(txtUnicode(Index))
   'Mark by Lydia 2016/10/28 受win7輸入法影響,不切換輸入法
   'CloseIme
End Sub
'Added by Lydia 2016/04/18 回傳查名內容
Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim mLoad As Boolean
Dim sPath As String
Dim APKind As String
Dim rsR As New ADODB.Recordset
Dim mNo As String
If iList <> "" Then
    Me.Enabled = False: mLoad = False
    Screen.MousePointer = vbHourglass
    APKind = "HM" 'm_TMQApp '先代入申請號
    
    m_PrevForm.Show
    mLoad = False
    '以第一個申請編號的查詢內容為主
       strExc(1) = GetAddStr(iList)
       strExc(0) = "select a.*,tqf03,tqf05 from tmqapp a ,tmqfile f " & _
                      "where tqa01 in (" & strExc(1) & ") and tqa01=tqf01(+) and tqf02(+)='" & TMQ_附件F02 & "' and tqf04(+)='" & TMQ_附件F04 & "'"
                      
       strExc(0) = strExc(0) & " order by tqa01 desc,tqf03 "
        intI = 1
        Set rsR = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           rsR.MoveFirst
           APKind = rsR.Fields("tqa01")
           mNo = rsR.Fields("tqa01") & " " & rsR.Fields("tqa06")
           txtUnicode(0) = ""
           Do While Not rsR.EOF
              If mNo <> rsR.Fields("tqa01") & " " & rsR.Fields("tqa06") Then Exit Do
              
              If rsR.Fields("tqf03") = TMQ_AkindPic Then '圖形
                 If InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 Then
                    mLoad = True
                    'Modified by Lydia 2016/06/23
                    'APKind = APKind & "_" & TMQ_AkindPic
                    APKind = APKind & TMQ_附件F02 & TMQ_AkindPic & TMQ_附件F04
                 End If
              Else
                 '文字
                 If rsR.AbsolutePosition = 1 Then
                    If Len(rsR.Fields("tqa07")) > 0 Then
                       txtUnicode(0) = txtUnicode(0) & rsR.Fields("tqa07") & " "
                    ElseIf Len(rsR.Fields("tqa13")) > 0 Then
                       txtUnicode(0) = txtUnicode(0) & rsR.Fields("tqa13") & " "
                    End If
                    If Len(rsR.Fields("tqa08")) > 0 Then
                       txtUnicode(0) = txtUnicode(0) & rsR.Fields("tqa08") & " "
                    ElseIf Len(rsR.Fields("tqa14")) > 0 Then
                       txtUnicode(0) = txtUnicode(0) & rsR.Fields("tqa14") & " "
                    End If
                 End If
                 
                 If rsR.Fields("tqf03") = TMQ_AkindWord1 And InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 Then
                       mLoad = True
                       'Modified by Lydia 2016/06/23
                       'APKind = APKind & "_" & TMQ_AkindWord1
                       APKind = APKind & TMQ_附件F02 & TMQ_AkindWord1 & TMQ_附件F04
                 ElseIf rsR.Fields("tqf03") = TMQ_AkindWord2 And InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 And mLoad = False Then
                       mLoad = True
                       'Modified by Lydia 2016/06/23
                       'APKind = APKind & "_" & TMQ_AkindWord2
                       APKind = APKind & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04
                 End If
              End If
              rsR.MoveNext
           Loop
        End If
        If txtUnicode(0) <> "" Then
          m_PrevForm.opt1(0).Value = True
          'm_PrevForm.PicText = txtUnicode(0) 'Mark by Lydia 2024/10/25 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
        ElseIf mLoad = True Then
          sPath = Dir(m_AttachPath & "\" & APKind & "*.*")
          If sPath = "" Then
             mLoad = AttachFileGet(Left(mNo, 10), Right(APKind, 1), , sPath)
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
    If m_PrevForm.Text1(6) = "T" Then 'Added by Lydia 2017/10/19 TS案無商標種類
       m_PrevForm.Combo6.ListIndex = 0 'Added by Lydia 2016/05/30 接洽單的商標種類
    End If 'end 2017/10/19
    m_PrevForm.bolExternalCall = False '還原預設值
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Me.Hide
End If
End Sub

'Added by Lydia 2016/06/17 抓委查單號
Private Function GetTQA01toTMQ01(ByVal pNo As String) As String
Dim rsR1 As New ADODB.Recordset
Dim inR1 As Integer
Dim strR1 As String
    
    strR1 = "select tmq01 from trademarkquery where tmq18='" & pNo & "' order by tmq01"
    inR1 = 1
    Set rsR1 = ClsLawReadRstMsg(inR1, strR1)
    If inR1 = 1 Then
       strR1 = rsR1.GetString(adClipString, , , ",")
       strR1 = Left(strR1, Len(strR1) - 1)
       GetTQA01toTMQ01 = strR1
    End If
    
End Function

'Added by Lydia 2018/11/22
'Remove by Lydia 2021/10/01
'Private Sub txtUnicode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If SyxMsg <> "txtUnicode_" & Format(Index, "00") Then '避免連續產生訊息
'        bolMsgRight = False
'        SyxMsg = "txtUnicode_" & Format(Index, "00")
'    End If
'    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
'End Sub

'Added by Lydia 2021/10/01 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtUnicode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then Forms(0).PopupMenu2 txtUnicode(Index)
End Sub

'Added by Lydia 2021/10/01
Private Sub textCName_GotFocus()
    TextInverse textCName
End Sub

'Added by Lydia 2021/10/01
Private Sub textCName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textCName
End Sub

'Added by Lydia 2021/10/01
Private Sub textService_GotFocus()
    TextInverse textService
End Sub

'Added by Lydia 2021/10/01
Private Sub textService_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textService
   
End Sub

'Added by Lydia 2021/10/01 上線後執行,將舊資料的TQA07-TQA08改存入TQA13-TQA14
Private Sub CmdTrans_Click()
    opt1(0).Value = True
    cmdTrans.Enabled = False
    If strSrvDate(1) >= Form20上線日 Then
        strSql = "select * from tmqapp where tqa07||tqa08 is not null order by tqa11 desc,tqa12 desc "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            Do While Not RsTemp.EOF
                 cnnConnection.BeginTrans
                 On Error GoTo ErrHandle
                 txtUnicode(1).Text = "": txtUnicode(2).Text = ""
                 If Len(RsTemp.Fields("tqa07")) > 0 Then
                     txtUnicode(1) = RsTemp.Fields("tqa07")
                     strExc(1) = "update tmqapp set tqa13=" & CNULL(ChgSQL(txtUnicode(1))) & ",tqa07=null  where tqa01=" & CNULL(RsTemp.Fields("tqa01"))
                     cnnConnection.Execute strExc(1)
                 End If
                 If Len(RsTemp.Fields("tqa08")) > 0 Then
                     txtUnicode(2) = RsTemp.Fields("tqa08")
                     strExc(1) = "update tmqapp set tqa14=" & CNULL(ChgSQL(txtUnicode(2))) & ",tqa08=null  where tqa01=" & CNULL(RsTemp.Fields("tqa01"))
                     cnnConnection.Execute strExc(1)
                 End If
                 cnnConnection.CommitTrans
                 RsTemp.MoveNext
            Loop
        End If
    End If
    
    MsgBox "OK !"
    
    cmdTrans.Enabled = True
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
       MsgBox Err.Desc
       cnnConnection.RollbackTrans
       cmdTrans.Enabled = True
    End If
End Sub

'Added by Lydia 2024/05/09 查名單-網中：指定範圍的(原)委查單匯入(新)商標查詢輸入(網中)
Private Sub cmdApp_Click(Index As Integer)
'Memo by Lydia 2024/08/02 Table調整為54欄位
'Memo by Lydia 2025/04/24 Table調整為71欄位
Dim strKey As String
Dim strDate1 As String, StrDate2 As String
Dim rsRead As New ADODB.Recordset, intR As Integer
Dim strForCh As String, strForEn As String, strForJp As String, strSerKey As String, bolReCall As Boolean
Dim strPath As String, strFilePath As String
Dim tmpArrB As Variant, tmpArrA As Variant, intA As Integer, intB As Integer 'Added by Lydia 2024/05/27
Dim inC As Integer, inE As Integer, inP As Integer, cntX As Integer, bolCount As Boolean 'Added by Lydia 2025/04/4


   '根據(112年嘉雯提供文字查名結果)查名記錄final: 取1120904-1120928期間資料匯入
   'Mark by Lydia 2025/04/24 改成其他指定範圍
   'strDate1 = "20230904"
   'StrDate2 = "20230928"
   'If strDate1 = "" Or StrDate2 = "" Or StrDate2 < strDate1 Then
   '   MsgBox "無法執行！", vbExclamation
   'End If
   'end 2025/04/24
   
   strPath = App.path & "\" & strUserNum
   If Dir(strPath, vbDirectory) = "" Then
      MkDir strPath
   Else
      Call PUB_KillAnyFile(strPath)
   End If
   
   If Index = 2 Then  '下載檔案
      strExc(1) = "H11300004302.PDF"
      If PUB_TMQAppFileGet(strPath, strExc(1), Mid(strExc(1), 1, 9), Mid(strExc(1), 10, 1), Mid(strExc(1), 11, 2)) = True Then
         MsgBox "下載完成：" & vbCrLf & strExc(1)
      End If
      Exit Sub
   End If
   'Mark by Lydia 2025/04/24 改成其他指定範圍
   'strSql = " select tqa01,tqa02,tqa04,tqa06,tqa13,tqa14,tmq01,tmq10,tmq03,tmq24,tmq07,tmq08,tmq09" & _
            " From trademarkquery, tmqapp" & _
            " where tmq05>=" & strDate1 & " and tmq05<=" & StrDate2 & " and tmq18=tqa01(+) and tqa20 is null "
   strSql = " select tqa01,tqa02,tqa04,tqa06,tqa13,tqa14,tmq01,tmq10,tmq03,tmq24,tmq07,tmq08,tmq09" & _
            " From trademarkquery, tmqapp" & _
            " where tmq05>=20250407 and tmq05<=20250408 and tqa06='1' and tmq18=tqa01(+)" & _
            " and tqa20 is null and instr(tqa03,'8888')=0 and instr(tqa03,'9999')=0"
   bolCount = False  '是否寫入委查筆數
   'end 2025/04/24
   strSql = strSql & " order by tmq01"
   intR = 0
   Set rsRead = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      Screen.MousePointer = vbHourglass
      rsRead.MoveFirst
      Do While Not rsRead.EOF
         strForCh = "": strForEn = "": strForJp = ""
         bolReCall = False
         strSerKey = "" & rsRead.Fields("tqa13") '文字1
JumpWord2:
         'Mark by Lydia 2025/04/24 改成其他指定範圍
         'If "" & rsRead.Fields("tqa06") = "1" Then '文字
         '   strExc(0) = "select '1' as ord1,中文,英文,日文,備註,st01 from lydia_tmq2word where st01='" & rsRead.Fields("tmq10") & "' and 文字商標='" & ChgSQL(strSerKey) & "' "
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 0 Then  '抓其他人的分析
         '      strExc(0) = "select '2' as ord1,中文,英文,日文,備註,st01 from lydia_tmq2word where 文字商標='" & ChgSQL(strSerKey) & "' "
         '      intI = 1
        '       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        '    End If
        '    If intI = 1 Then
        '       RsTemp.MoveFirst
        '       Do While Not RsTemp.EOF
        '          If "" & RsTemp.Fields("中文") = "Y" Or Val("" & rsRead.Fields("tmq07")) > 0 Then
        '             strForCh = Pub_GetTwoListChk(strForCh, "" & RsTemp.Fields("備註"))
        '          End If
        '          If "" & RsTemp.Fields("英文") = "Y" Or Val("" & rsRead.Fields("tmq08")) > 0 Then
        '             strForEn = Pub_GetTwoListChk(strForEn, "" & RsTemp.Fields("備註"))
        '          End If
        '          If "" & RsTemp.Fields("日文") = "Y" Then
         '            strForJp = Pub_GetTwoListChk(strForJp, "" & RsTemp.Fields("備註"))
          '        End If
          '        RsTemp.MoveNext
          '     Loop
          '  End If
         'End If
         ''取得已存在的查覆結果
         'strExc(0) = "select nvl(min(tqd06),'7') tqd06 from tmqdetail where tqd02='" & rsRead.Fields("tmq01") & "' "
         'If "" & rsRead.Fields("tqa06") = "2" Then
         '   strExc(0) = strExc(0) & " and tqd03='0' "
         'Else
         '   If bolReCall = False Then
         '      strExc(0) = strExc(0) & " and tqd03='1' "
         '   Else
         '      strExc(0) = strExc(0) & " and tqd03='2' "
         '   End If
         'End If
         'intI = 1
         'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'If intI = 1 Then
         'end 2025/04/24
            'Modified by Lydia 2025/04/01 改用最大流水號+1碼,(保留)舊Code
            'strKey = AutoNo("HH", 5)
            strKey = Pub_GetAutoTMA01
            strSql = "insert into TMQAPPFORM(TMA01,TMA02,TMA03,TMA04,TMA08,TMA18,TMA23,TMA24,TMA25,TMA26,TMA27,TMA28,TMA29,TMA71,TMA10) values "
            'TMA01查名單號,TMA02資料來源,TMA03~TMA04 CREATE ID,DATE
            strSql = strSql & "('" & strKey & "', '1', 'QPGMR', sysdate "
            'TMA08委查人(員工代號),TMA18客戶名稱,TMA23組群,TMA24組群3519
            If InStr("" & rsRead.Fields("tmq03"), "3519") = 0 Then
               strExc(3) = "" & rsRead.Fields("tmq03")
               strExc(4) = ""
            Else
               If InStr("" & rsRead.Fields("tmq03"), "3519") = 1 Then
                  strExc(3) = ""
                  strExc(4) = "" & rsRead.Fields("tmq03")
               Else
                  strExc(3) = Mid("" & rsRead.Fields("tmq03"), 1, InStr("" & rsRead.Fields("tmq03"), "3519") - 2)
                  strExc(4) = Mid("" & rsRead.Fields("tmq03"), InStr("" & rsRead.Fields("tmq03"), "3519"))
               End If
            End If
            strSql = strSql & ", '" & rsRead.Fields("tqa02") & "','" & ChgSQL("" & rsRead.Fields("tqa04")) & "', '" & strExc(3) & "', '" & strExc(4) & "' "
            'TMA25檢索方式,TMA26文字,TMA27圖形,TMA28查名路徑
            '查名路徑:拿掉、其他不動
            strExc(1) = Replace(Replace("" & rsRead.Fields("tmq24"), " 、 ", ","), "、", ",")
            'Added by Lydia 2024/05/27
            strExc(1) = Replace(Replace(strExc(1), "(", ""), ")", "") '拿掉()
            strExc(1) = Replace(Replace(strExc(1), ", ", ","), " ,", ",") '拿掉,加空白
            strExc(1) = Replace(strExc(1), " ", ",") '空白改成,
            If InStr(strExc(1), "/") > 0 Then
               strExc(2) = "": strExc(3) = ""
               tmpArrA = Empty
               tmpArrA = Split(strExc(1), ",")
               For intA = 0 To UBound(tmpArrA)
                  If Trim(tmpArrA(intA)) <> "" Then
                     strExc(3) = ""
                     tmpArrB = Empty
                     tmpArrB = Split(tmpArrA(intA), "/")
                     If UBound(tmpArrB) = 0 Then
                        strExc(2) = strExc(2) & "," & Trim(tmpArrB(0))
                     Else
                        For intB = 0 To UBound(tmpArrB)
                           If Len(Trim(tmpArrB(intB))) = 5 Then
                              strExc(3) = Mid(Trim(tmpArrB(intB)), 1, 3)
                              strExc(2) = strExc(2) & "," & Trim(tmpArrB(intB))
                           ElseIf Len(Trim(tmpArrB(intB))) = 7 Then
                              strExc(3) = Mid(Trim(tmpArrB(intB)), 1, 5)
                              strExc(2) = strExc(2) & "," & Trim(tmpArrB(intB))
                           ElseIf Len(Trim(tmpArrB(intB))) <= 2 And strExc(3) <> "" Then
                              strExc(2) = strExc(2) & "," & strExc(3) & Trim(tmpArrB(intB))
                           End If
                        Next intB
                     End If
                     strExc(1) = Mid(strExc(2), 2)
                  End If
               Next intA
            End If
            'end 2024/05/27
            'Added by Lydia 2025/04/24 重新計算筆數
            inC = 0: inE = 0: inP = 0: cntX = 0
            If Val("" & rsRead.Fields("tqa06")) = 2 Then
               inP = 1
            Else
               Call PUB_CountTxtNEC(inE, inC, strSerKey)
            End If
            strExc(1) = Mid(IIf(strExc(3) <> "", "," & strExc(3), "") & IIf(strExc(4) <> "", "," & strExc(4), ""), 2)
            tmpArrA = Empty
            tmpArrA = Split(strExc(1), ",")
            cntX = UBound(tmpArrA) + 1
            'end 2025/04/24
            strSql = strSql & ", '" & rsRead.Fields("tqa06") & "', '" & ChgSQL(strSerKey) & "',  '" & IIf("" & rsRead.Fields("tqa06") = "2", "Y", "") & "', '" & ChgSQL(strExc(1)) & "' "
            'TMA29查詢資料範圍,TMA45~TMA47 AI檢索中文,英文,日文
            'Modified by Lydia 2024/09/23 改成空白=>AI拆字
            'strSql = strSql & ", '1', '" & ChgSQL(StrToStr(strForCh, 200)) & "', '" & ChgSQL(StrToStr(strForEn, 200)) & "', '" & ChgSQL(StrToStr(strForJp, 200)) & "' "
            strSql = strSql & ", '1' "
            'TMA71(原)委查單號 ,TMA10查覆(查名)人員(商申組)
            strSql = strSql & ", '" & rsRead.Fields("tmq01") & "', '" & rsRead.Fields("tmq10") & "') "
            cnnConnection.Execute "delete from TMQAppForm where TMA01='" & strKey & "' and tma03='AAAAAA' " 'Added by Lydia 2025/04/01 先刪除保留記錄
            cnnConnection.Execute strSql
            'Added by Lydia 2025/04/24
            If bolCount = True Then
               strSql = "Update TMQAppForm set tma36=" & inC & ", tma37=" & inE & ", tma38=" & inP & " Where TMA01='" & strKey & "' "
               cnnConnection.Execute strSql
            End If
            'end 2025/04/24
            
            '抓附件PUB_TMQAFileSave (tqf02='0'):
            If "" & rsRead.Fields("tqa06") = "2" Then 'Added by Lydia 2025/04/24 圖形查名才需要上傳附件
               strExc(0) = "select tqf01,tqf02,tqf03,tqf04,tqf05,tqf12 from tmqfile where tqf01='" & rsRead.Fields("tqa01") & "' and tqf02='0' " & _
                           " order by 1,2,3,4"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  intR = 0
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     strFilePath = strPath & "\" & RsTemp.Fields("tqf01") & RsTemp.Fields("tqf02") & RsTemp.Fields("tqf03") & RsTemp.Fields("tqf04") & "." & RsTemp.Fields("tqf05")
                     If PUB_TMQGetAFile("", strFilePath, RsTemp.Fields("tqf01"), RsTemp.Fields("tqf02"), RsTemp.Fields("tqf03"), RsTemp.Fields("tqf04"), RsTemp.Fields("tqf05")) = True Then
                        If PUB_TMQAppFileSave(False, strKey, IIf(Len("" & RsTemp.Fields("tqf02")) = 9, "3", "1"), "00", strFilePath) = False Then
                           MsgBox "上傳檔案失敗!" & vbCrLf & "查名單號:" & strKey & vbCrLf & strFilePath
                           Exit Do
                        Else
                           PUB_DelPCOrgFile strFilePath
                        End If
                     End If
                     RsTemp.MoveNext
                  Loop
               End If
            End If 'Added by Lydia 2025/04/24
         'End If 'Mark by Lydia 2025/04/24 改成其他指定範圍
         If "" & rsRead.Fields("tqa06") = "1" And bolReCall = False And "" & rsRead.Fields("tqa14") <> "" Then
            strSerKey = "" & rsRead.Fields("tqa14")  '文字2
            strForCh = "": strForEn = "": strForJp = ""
            bolReCall = True
            GoTo JumpWord2
         End If
         rsRead.MoveNext
      Loop
      Screen.MousePointer = vbDefault
      MsgBox "執行完畢！"
   End If
End Sub

'Added by Lydia 2024/05/31 查名單網中系統：將舊查名單轉換成新查名單
Private Function ProcPackNew(ByVal pKeyNo As String) As Boolean
'Memo by Lydia 2024/08/12 Table調整為54欄位>>文字查名測試OK,組群區分3519輸入
Dim intQ As Integer, strQuery As String, strNewKey As String
Dim rsQD As New ADODB.Recordset
Dim bolConn As Boolean, strSerKey As String, bolReCall As Boolean
Dim strFileName As String
Dim stDate1 As String, stDate2 As String, tmpArr As Variant
Dim inC As Integer, inE As Integer, inP As Integer, cntX As Integer

   ProcPackNew = False
   If pKeyNo = "" Then Exit Function
   
On Error GoTo ErrHandle

   '排除:團體標章和證明標章仍由查名人員進行查名
   'Modified by Lydia 2024/10/04　關閉證明標章「9999」代碼；拿掉 and instr(tmq03,'9999')=0
   strQuery = " select tqa01,tqa02,tqa04,tqa06,tqa13,tqa14,tmq01,tmq12,tmq10,tmq03,tmq04,tmq05,tmq24,tmq07,tmq08,tmq09,tmq14,tmq21" & _
            " From trademarkquery, tmqapp where tqa01='" & pKeyNo & "' and tmq18=tqa01(+) and instr(tmq03,'8888')=0 "
   strQuery = strQuery & " order by tmq01"
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      rsQD.MoveFirst
      bolConn = True
      cnnConnection.BeginTrans
      Do While Not rsQD.EOF
         bolReCall = False
         strSerKey = "" & rsQD.Fields("tqa13") '文字1
         If "" & rsQD.Fields("tmq05") <> "" Then
            stDate1 = PUB_GetNewTMADate(IIf(Val("" & rsQD.Fields("tmq09")) > 0, "2", "1"), "" & rsQD.Fields("tmq04"), "" & rsQD.Fields("tmq14"), Pub_GetSpecMan("內商查名單分單狀態"))
            stDate2 = "" '查覆期限要等到網中回寫後才能計算(Trigger)
         Else
            stDate1 = "" '送出(網中)期限
            stDate2 = "" '查覆期限
            '每日批次StrMenu62已寫隔日分發的處理
         End If
JumpWord2:  '文字2另外開一張單
         'Modified by Lydia 2025/04/01 改用最大流水號+1碼,(保留)舊Code
         'strNewKey = AutoNo("HH", 5)
         strNewKey = Pub_GetAutoTMA01
         strSql = "insert into TMQAPPFORM(TMA01,TMA71,TMA34,TMA35,TMA02,TMA03,TMA04,TMA08,TMA18,TMA23,TMA24,TMA25,TMA26,TMA27,TMA28,TMA29,TMA09,TMA10,TMA11,TMA12,TMA36,TMA37,TMA38) values "
                   
         If txt1(6) <> "" And txt1(7) <> "" Then
            strExc(1) = txt1(6) & txt1(7) & IIf(txt1(8) = "", "0", txt1(8)) & IIf(txt1(9) = "", "00", txt1(9))
         Else
            strExc(1) = ""
         End If
         
         'TMA01查名單號,TMA71(原)委查單號,TMA34=新案收文號,TMA35=已收文本所案號,TMA02資料來源,TMA03~TMA04 CREATE ID,DATE
         strSql = strSql & "('" & strNewKey & "', '" & rsQD.Fields("tmq01") & "','" & rsQD.Fields("tmq21") & "','" & strExc(1) & "', '1', '" & rsQD.Fields("tmq12") & "', sysdate "
         'TMA08委查人(員工代號),TMA18客戶名稱,TMA23組群,TMA24組群3519
         If InStr("" & rsQD.Fields("tmq03"), "3519") = 0 Then
            strExc(3) = "" & rsQD.Fields("tmq03")
            strExc(4) = ""
         Else
            If InStr("" & rsQD.Fields("tmq03"), "3519") = 1 Then
               strExc(3) = ""
               strExc(4) = "" & rsQD.Fields("tmq03")
            Else
               strExc(3) = Mid("" & rsQD.Fields("tmq03"), 1, InStr("" & rsQD.Fields("tmq03"), "3519") - 2)
               strExc(4) = Mid("" & rsQD.Fields("tmq03"), InStr("" & rsQD.Fields("tmq03"), "3519"))
            End If
         End If
         '重新計算筆數
         inC = 0: inE = 0: inP = 0: cntX = 0
         If Val("" & rsQD.Fields("tqa06")) = 2 Then
            inP = 1
         Else
            Call PUB_CountTxtNEC(inE, inC, strSerKey)
         End If
         strExc(1) = Mid(IIf(strExc(3) <> "", "," & strExc(3), "") & IIf(strExc(4) <> "", "," & strExc(4), ""), 2)
         tmpArr = Empty
         tmpArr = Split(strExc(1), ",")
         cntX = UBound(tmpArr) + 1
         
         strSql = strSql & ", '" & rsQD.Fields("tqa02") & "', '" & ChgSQL("" & rsQD.Fields("tqa04")) & "', '" & strExc(3) & "', '" & strExc(4) & "' "
         'TMA25檢索方式,TMA26文字,TMA27圖形,TMA28查名路徑=>查名人員後補,TMA29查詢資料範圍
         strSql = strSql & ",'" & rsQD.Fields("tqa06") & "', '" & ChgSQL(strSerKey) & "',  '" & IIf("" & rsQD.Fields("tqa06") = "2", "Y", "") & "', NULL, '1' "
         'TMA09分發日期,TMA10查覆(查名)人員(商申組),TMA11查覆期限,TMA12送出期限
         strSql = strSql & ", '" & rsQD.Fields("tmq05") & "', '" & rsQD.Fields("tmq10") & "','" & stDate2 & "', '" & stDate1 & "' "
         'TMA36~TMA38 委查中文,英文,圖形筆數
         strSql = strSql & ", '" & inC * cntX & "', '" & inE * cntX & "', '" & inP * cntX & "') "
         cnnConnection.Execute "delete from TMQAppForm where TMA01='" & strNewKey & "' and tma03='AAAAAA' " 'Added by Lydia 2025/04/01 先刪除保留記錄
         '---1234每日批次-隔日分發的處理
         cnnConnection.Execute strSql
         
         '附件
         strQuery = "select tqf01,tqf02,tqf03,tqf04,tqf05,tqf12 from tmqfile where tqf01='" & rsQD.Fields("tqa01") & "' order by 1,2,3,4 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strQuery)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strFileName = RsTemp.Fields("tqf01") & RsTemp.Fields("tqf02") & RsTemp.Fields("tqf03") & RsTemp.Fields("tqf04") & "." & RsTemp.Fields("tqf05")
               PUB_DelPCOrgFile m_AttachPath & "\" & strFileName
               If PUB_TMQGetAFile(m_AttachPath, strFileName, RsTemp.Fields("tqf01"), RsTemp.Fields("tqf02"), RsTemp.Fields("tqf03"), RsTemp.Fields("tqf04"), RsTemp.Fields("tqf05")) = True Then
                  If PUB_TMQAppFileSave(False, strNewKey, "1", "00", strFileName, False) = False Then
                     MsgBox "上傳檔案失敗!" & vbCrLf & "查名單號:" & strNewKey & vbCrLf & strFileName, vbCritical, "轉換網中資料格式"
                     Exit Do
                  Else
                     PUB_DelPCOrgFile strFileName
                  End If
               Else
                   MsgBox "無法下載檔案：" & vbCrLf & strFileName, vbCritical, "轉換網中資料格式"
               End If
               RsTemp.MoveNext
            Loop
         End If
         
         If "" & rsQD.Fields("tqa06") = "1" And bolReCall = False And "" & rsQD.Fields("tqa14") <> "" Then
            strSerKey = "" & rsQD.Fields("tqa14")  '文字2
            bolReCall = True
            GoTo JumpWord2
         End If
         rsQD.MoveNext
      Loop
      cnnConnection.CommitTrans
      bolConn = False
   End If
   Set rsQD = Nothing
   ProcPackNew = True
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      If bolConn = True Then cnnConnection.RollbackTrans
      
      MsgBox "發生錯誤" & vbCrLf & Err.Description, vbCritical, "轉換網中資料格式"
   End If
End Function

'Added by Lydia 2024/07/17
Private Sub cmdGrp_Click()

   If Not nFrm Is Nothing Then
      nFrm.SetParent IIf(mApNoList <> "", "Q", "M"), Me, textService, IIf(opt1(0).Value = "1", "W", "P")
      nFrm.Show vbModal
   End If
End Sub

'Added by Lydia 2024/07/17
Public Sub SetData(ByVal pInputVal As String)
   textService.Text = pInputVal
   If lblAutoNo.Caption = "" And (txt1(1) <> "" Or textService <> "") Then
      Call GetKeyNo
   End If
End Sub

'Added by Lydia 2024/07/17
Private Sub GetKeyNo()
Dim strTmp1 As String, strTmp2 As String

   strTmp2 = AccAutoNo("HM", 4, GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2)))  '取得自動編號
   lblAutoNo.Caption = strTmp2
   strTmp1 = AccSaveAutoNo("HM", Right(strTmp2, 4), GetTaiwanThisYear, Val(Mid(strSrvDate(1), 5, 2))) '回寫acc1r0
   '保留上次輸入資料
   If (ChkS1.Value = 1 Or ChkS2.Value = 1) And tmpReNo <> "" And haveKey <> "" Then
      If AttachFileRedo(tmpReNo, haveKey) = False Then
      End If
   End If
End Sub


