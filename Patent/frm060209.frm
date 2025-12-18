VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060209 
   BorderStyle     =   1  '單線固定
   Caption         =   "行事曆提醒通知"
   ClientHeight    =   5484
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8412
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   8412
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   60
      TabIndex        =   11
      Top             =   480
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   8700
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "查詢"
      TabPicture(0)   =   "frm060209.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCnt"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Txt1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Txt1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Cmb2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "GRD1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Txt1(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Combo1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "明細資料"
      TabPicture(1)   =   "frm060209.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCP"
      Tab(1).Control(1)=   "txtSC(2)"
      Tab(1).Control(2)=   "txtSC(1)"
      Tab(1).Control(3)=   "txtSC(5)"
      Tab(1).Control(4)=   "txtSC(6)"
      Tab(1).Control(5)=   "txtSC(7)"
      Tab(1).Control(6)=   "txtSC(8)"
      Tab(1).Control(7)=   "txtSC(10)"
      Tab(1).Control(8)=   "txtSC(3)"
      Tab(1).Control(9)=   "txtSC(9)"
      Tab(1).Control(10)=   "txtSC(4)"
      Tab(1).Control(11)=   "lstUsers(1)"
      Tab(1).Control(12)=   "lstUsers(0)"
      Tab(1).Control(13)=   "lstSC04"
      Tab(1).Control(14)=   "Line1"
      Tab(1).Control(15)=   "textCUID"
      Tab(1).Control(16)=   "Cmb1"
      Tab(1).Control(17)=   "lblFC(1)"
      Tab(1).Control(18)=   "lblFC(0)"
      Tab(1).Control(19)=   "Label1(10)"
      Tab(1).Control(20)=   "Label1(13)"
      Tab(1).Control(21)=   "Label1(2)"
      Tab(1).Control(22)=   "Label1(0)"
      Tab(1).Control(23)=   "Label1(1)"
      Tab(1).Control(24)=   "Label1(3)"
      Tab(1).Control(25)=   "Label1(4)"
      Tab(1).Control(26)=   "Label1(6)"
      Tab(1).Control(27)=   "Label1(11)"
      Tab(1).ControlCount=   28
      Begin VB.CommandButton cmdCP 
         Caption         =   "進度檔"
         Height          =   375
         Left            =   -68800
         Style           =   1  '圖片外觀
         TabIndex        =   47
         Top             =   432
         Width           =   855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   372
         Left            =   4536
         TabIndex        =   45
         Top             =   792
         Width           =   3588
         Begin VB.TextBox txtCode 
            Height          =   270
            Index           =   3
            Left            =   2712
            MaxLength       =   2
            TabIndex        =   8
            Top             =   0
            Width           =   444
         End
         Begin VB.TextBox txtCode 
            Height          =   270
            Index           =   2
            Left            =   2328
            MaxLength       =   1
            TabIndex        =   7
            Top             =   0
            Width           =   324
         End
         Begin VB.TextBox txtCode 
            Height          =   270
            Index           =   1
            Left            =   1488
            MaxLength       =   6
            TabIndex        =   6
            Top             =   0
            Width           =   780
         End
         Begin VB.TextBox txtCode 
            Height          =   270
            Index           =   0
            Left            =   984
            MaxLength       =   3
            TabIndex        =   5
            Top             =   0
            Width           =   444
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   1320
            X2              =   2784
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   204
            Index           =   4
            Left            =   24
            TabIndex        =   46
            Top             =   48
            Width           =   972
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         ItemData        =   "frm060209.frx":0038
         Left            =   5040
         List            =   "frm060209.frx":003A
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   472
         Width           =   2835
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3336
         Left            =   96
         TabIndex        =   23
         Top             =   1248
         Width           =   8052
         _ExtentX        =   14203
         _ExtentY        =   5906
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "細明體-ExtB"
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
      Begin VB.ComboBox Cmb2 
         Height          =   276
         Left            =   3240
         TabIndex        =   4
         Text            =   "Cmb2"
         Top             =   795
         Width           =   1215
      End
      Begin VB.TextBox Txt1 
         Height          =   270
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   795
         Width           =   900
      End
      Begin VB.TextBox Txt1 
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   795
         Width           =   900
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   2
         Left            =   -70440
         TabIndex        =   14
         Top             =   480
         Width           =   400
         VariousPropertyBits=   679493659
         MaxLength       =   4
         Size            =   "706;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   13
         Top             =   480
         Width           =   1005
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1773;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   5
         Left            =   -73920
         TabIndex        =   15
         Top             =   810
         Width           =   555
         VariousPropertyBits=   679493659
         MaxLength       =   3
         Size            =   "979;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   6
         Left            =   -73320
         TabIndex        =   16
         Top             =   810
         Width           =   795
         VariousPropertyBits=   679493659
         MaxLength       =   6
         Size            =   "1402;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   7
         Left            =   -72480
         TabIndex        =   17
         Top             =   810
         Width           =   345
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "609;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   8
         Left            =   -72120
         TabIndex        =   18
         Top             =   810
         Width           =   555
         VariousPropertyBits=   679493659
         MaxLength       =   2
         Size            =   "979;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   10
         Left            =   -70440
         TabIndex        =   19
         Top             =   810
         Width           =   400
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "706;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   3
         Left            =   -74880
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3060
         Visible         =   0   'False
         Width           =   400
         VariousPropertyBits=   679493659
         Size            =   "706;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   9
         Left            =   -69060
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   400
         VariousPropertyBits=   679493659
         Size            =   "706;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   4
         Left            =   -74880
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2310
         Visible         =   0   'False
         Width           =   400
         VariousPropertyBits=   679493659
         MaxLength       =   300
         Size            =   "706;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   285
         Index           =   1
         Left            =   -70380
         TabIndex        =   44
         Top             =   2760
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;503"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   43
         Top             =   2760
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;424"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstSC04 
         Height          =   285
         Left            =   -73920
         TabIndex        =   42
         Top             =   1830
         Width           =   5685
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "10028;422"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   -73560
         X2              =   -71970
         Y1              =   930
         Y2              =   930
      End
      Begin MSForms.Label lblName 
         Height          =   225
         Left            =   2100
         TabIndex        =   41
         Top             =   510
         Width           =   1305
         Caption         =   "Form2.0"
         Size            =   "2302;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   -74880
         TabIndex        =   40
         Top             =   4500
         Width           =   6675
         VariousPropertyBits=   679495707
         Size            =   "11774;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Cmb1 
         Height          =   330
         Left            =   -73920
         TabIndex        =   39
         Top             =   1170
         Width           =   5655
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "9975;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFC 
         Height          =   225
         Index           =   1
         Left            =   -72990
         TabIndex        =   38
         Top             =   1560
         Width           =   6135
         Caption         =   "Form2.0"
         Size            =   "10821;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFC 
         Height          =   225
         Index           =   0
         Left            =   -73920
         TabIndex        =   37
         Top             =   1560
         Width           =   885
         Caption         =   "Form2.0"
         Size            =   "1561;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCnt 
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   6240
         TabIndex        =   36
         Top             =   4680
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "提醒人員："
         Height          =   180
         Index           =   10
         Left            =   -74865
         TabIndex        =   35
         Top             =   2790
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流  水  號：              (自動編號)"
         Height          =   180
         Index           =   13
         Left            =   -71400
         TabIndex        =   34
         Top             =   525
         Width           =   2500
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   180
         Index           =   2
         Left            =   -74865
         TabIndex        =   33
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "管制日期："
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   32
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   31
         Top             =   855
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "週　　期：          (1.單次 2.每週 3.每月 4.每3個月 5. 每年)"
         Height          =   180
         Index           =   3
         Left            =   -71400
         TabIndex        =   30
         Top             =   840
         Width           =   4440
      End
      Begin VB.Label Label1 
         Caption         =   "FC代理人："
         Height          =   180
         Index           =   4
         Left            =   -74865
         TabIndex        =   29
         Top             =   1582
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "事　　由："
         Height          =   180
         Index           =   6
         Left            =   -74865
         TabIndex        =   28
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "可解除人員："
         Height          =   180
         Index           =   11
         Left            =   -71520
         TabIndex        =   27
         Top             =   2790
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "顏色符號說明："
         Height          =   180
         Index           =   3
         Left            =   3720
         TabIndex        =   22
         Top             =   532
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "(                             )"
         Height          =   180
         Index           =   2
         Left            =   3120
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "管制期限："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "員工編號："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   525
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm060209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Lydia 2015/12/28 國外部行事曆提醒通知
'Memo by Lydia 2020/01/15 更名為「行事曆提醒通知」
'Memo By Lydia 2021/04/20 Form2.0已修改(lblFC、Cmb1、lstUsers、textCUID、lblName、lstSC04、txtSC(index)) 、GRD1改字型=新細明體-ExtB
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Public m_Role As String 'Added by Lydia 2020/01/15 業務拓展小組成員F41

'Modified by Lydia 2021/04/20
'Dim oText As TextBox
Dim oText As Control
Dim oLabel As LABEL
Dim idx As Integer
Dim mSC11 As String '建檔人
Dim mSC18 As String '解除期限日期
'Dim mESeqNo As String '暫存TB編號 'Remove by Lydia 2020/09/14
Dim stNumList1(1 To 5) As String
Dim mSC(1 To 11) As Integer 'grid的行位置
Dim SWPRow As Integer '選取列
Dim colSC18 As Integer
Dim bolDbClick As Boolean 'Added by Lydia 2021/04/20
Dim colSC20 As Integer, colPA177 As Integer 'Added by Lydia 2023/07/28
Dim m_PrevForm As Form, strCode(1 To 4) As String, bolPreClose As Boolean 'Added by Lydia 2025/09/10

'Added by Lydia 2025/09/10
Public Sub SetParent(ByRef pFrm As Form, Optional ByVal pCaseNo As String)
   
   Set m_PrevForm = pFrm
   If pCaseNo <> "" Then
      Call ChgCaseNo(pCaseNo, strCode)
   End If
End Sub

Private Function DetailCancel(ByVal inDx As Integer) As Boolean
Dim Sdate As String
Dim SNo  As Integer
Dim bMail As Boolean
Dim m_InputDate As String 'Added by Lydia 2023/07/28 輸入勘誤日期

   DetailCancel = False
   
On Error GoTo ErrorHand

 If GRD1.TextMatrix(inDx, mSC(1)) = "" Then Exit Function
   
   '檢查是否有異動
    txtSC(1) = ChangeWStringToTString(GRD1.TextMatrix(inDx, mSC(1)))
    txtSC(2) = GRD1.TextMatrix(inDx, mSC(2))
    If ShowRecord(0, False) = False Then
        MsgBox "查無資料，請確認行事曆資料是否有異動！", vbCritical, "解除管制"
        Exit Function
    End If
    
    If GRD1.TextMatrix(inDx, colSC18) <> "" Or mSC18 <> "" Then
        MsgBox "本記錄已解除管制!", vbCritical, "解除管制"
        Exit Function
    End If
    If JudgeRight(bMail, mSC11, Trim(txtSC(9))) Then
        'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：行事曆解除期限檢查
        If "" & GRD1.TextMatrix(inDx, colPA177) = "Y" And "" & GRD1.TextMatrix(inDx, colSC20) <> "" And InStr("" & txtSC(4), "請程序確認") > 0 And InStr("" & txtSC(4), "公報刊載日期") > 0 Then
           m_InputDate = PUB_ChkFCPlinkSC(DBDATE(txtSC(1)), txtSC(2))
           If m_InputDate = "" Then
              Exit Function
           End If
        End If
        'end 2023/07/28
        
        'Modified by Lydia 2016/03/01
        'strExc(1) = "請再次確定要解除" & txtSC(1) & "(流水號 " & txtSC(2) & _
                    " )" & vbCrLf & "事由：" & txtSC(4) & " ?"
        strExc(1) = "請再次確定要解除" & txtSC(1) & "(流水號 " & txtSC(2) & " )" & " ?" & vbCrLf & _
                   IIf(Trim(txtSC(5) & txtSC(6)) <> "", "本所案號: " & txtSC(5) & "-" & txtSC(6) & "-" & txtSC(7) & "-" & txtSC(8) & vbCrLf, "") & _
                   "事由:" & Replace(txtSC(4), vbCrLf, vbCrLf & "　　 ") & vbCrLf
        If MsgBox(strExc(1), vbInformation + vbYesNo, "解除管制") = vbYes Then
            strExc(2) = Mid(Right("000000" & ServerTime, 6), 1, 4)
            'Added by Lydia 2016/02/25  +案號顯示
            strExc(3) = IIf(Trim(txtSC(5) & txtSC(6)) <> "", " ，案號: " & txtSC(5) & Val(Trim(txtSC(6))) & IIf(txtSC(7) & txtSC(8) = "000", "", txtSC(7) & txtSC(8)), "")
            cnnConnection.BeginTrans
                If PUB_AddFCPStaffCalendar(IIf(txtSC(10) = "1", "", DBDATE(txtSC(1))), txtSC(10), txtSC(3), txtSC(4), txtSC(9), txtSC(10), txtSC(5), txtSC(6), txtSC(7), txtSC(8), Sdate, SNo, mSC11) Then
                   'Modified by Lydia 2016/02/25
                   'MsgBox "下次行事曆的管制日期: " & ChangeWStringToTString(Sdate) & "　流水號: " & sNO, vbInformation, "解除管制"
                   MsgBox "下次行事曆的管制日期: " & ChangeWStringToTString(Sdate) & "　流水號: " & SNo & strExc(3), vbInformation, "解除管制"
                Else
                   If txtSC(10) <> "1" Then MsgBox "下次行事曆新增失敗!", vbCritical, "解除管制"
                End If
               strSql = "UPDATE staff_calendar SET sc17='" & strUserNum & "',sc18=" & strSrvDate(1) & ",sc19=" & CNULL(strExc(2), True) & _
                        " where sc01=" & CNULL(DBDATE(txtSC(1)), True) & " and sc02=" & CNULL(txtSC(2), True)
               cnnConnection.Execute strSql, intI
               'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：解除行事曆=>當程序確認公報刊載日期後解除行事曆自動收文「通知資訊變更961」,發一封Email給承辦工程師
               If m_InputDate <> "" Then
                  strExc(0) = "select c2.cp09 as oldCP09,c2.cp10 as oldCP10,c1.cp09,c1.cp10,c1.cp12,c1.cp13,c1.cp14 from caseprogress c1, caseprogress c2 where c1.cp09='" & GRD1.TextMatrix(inDx, colSC20) & "' and c1.cp43=c2.cp09(+)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = txtSC(5): strExc(2) = txtSC(6):  strExc(3) = txtSC(7): strExc(4) = txtSC(8)
                     If PUB_GetFCPlinkMC("2A", m_InputDate, strExc, "" & RsTemp.Fields("CP09"), "" & RsTemp.Fields("oldCP10"), "" & RsTemp.Fields("CP10"), "" & RsTemp.Fields("CP12"), "" & RsTemp.Fields("CP13"), "" & RsTemp.Fields("CP14")) = True Then
                     End If
                     'Added by Lydia 2024/04/10 當程序解除行事曆期限時，系統會彈視窗輸入公告日，請自動將公報刊載日期一併掛在核准那道的承辦期限。 ----請參考frm06010602_3
                     strSql = "Update Caseprogress Set cp48='" & DBDATE(m_InputDate) & "' where cp09='" & GRD1.TextMatrix(inDx, colSC20) & "' and cp158=0 and cp159=0 "
                     cnnConnection.Execute strSql
                     'end 2024/04/10
                  End If
               End If
               'end 2023/07/28
            cnnConnection.CommitTrans

            mSC18 = strSrvDate(2)
            '解除人非輸入人員時,mail通知輸入人員
            'Modified by Lydia 2016/07/18 改成模組判斷
            'If bMail Then
              'Modified by Lydia 2016/02/25 +案號顯示
               'Modified by Lydia 2020/01/15 拿掉"國外部"
               'strExc(1) = "國外部行事曆：管制日期: " & txtSC(1) & " 流水號: " & txtSC(2) & strExc(3) & " ， 已被解除管制!"
               strExc(1) = "行事曆：管制日期: " & txtSC(1) & " 流水號: " & txtSC(2) & strExc(3) & " ， 已被解除管制!"
               'Modified by Lydia 2016/03/01 +行事曆內容
               strExc(4) = "本所案號: " & IIf(Trim(txtSC(5) & txtSC(6)) <> "", txtSC(5) & "-" & txtSC(6) & "-" & txtSC(7) & "-" & txtSC(8), "") & vbCrLf
               strExc(4) = strExc(4) & "案件名稱: " & Trim(IIf(Cmb1.Text <> "", Mid(Cmb1.Text, InStr(Cmb1.Text, ":") + 1), "")) & vbCrLf
               strExc(4) = strExc(4) & "事　　由: " & Replace(txtSC(4), vbCrLf, vbCrLf & "　　　　  ") & vbCrLf
            '   PUB_SendMail strUserNum, mSC11, "", strExc(1), strExc(4)
            'End If
            Call PUB_CancelFCPStaffCalendar(strUserNum, mSC11, strExc(1), strExc(4), txtSC(5), txtSC(6), txtSC(7), txtSC(8))
            'end 2016/07/18
            
            DetailCancel = True
            Call PUB_SendMailCache 'Added by Lydia 2023/08/25
        End If
    End If 'If JudgeRight
   
   Exit Function
   
ErrorHand:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String
   
   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

Private Sub Cmb2_Click()
Dim dType As Integer
Dim idx As Integer
    If Cmb2.Text <> "" Then
       Select Case Trim(Cmb2.Text)
           Case "一週"
                dType = 2:  idx = 7
           Case "二週"
                dType = 2:  idx = 14
           Case "三週"
                dType = 2:  idx = 21
           Case "一個月"
                dType = 1:  idx = 1
           Case "二個月"
                dType = 1:  idx = 2
           Case "三個月"
                dType = 1:  idx = 3
       End Select
        txt1(0) = ChangeWStringToTString(CompDate(dType, -idx, strSrvDate(1)))
        txt1(1) = ChangeWStringToTString(CompDate(dType, idx, strSrvDate(1)))
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim tmpB As Boolean
  
   If lblName = "" Then
      MsgBox "員工編號錯誤！", vbExclamation
      If txt1(2).Enabled = True Then
         txt1_GotFocus 2
         txt1(2).SetFocus
      End If
      Exit Sub
   ElseIf Pub_StrUserSt03 = "F22" And txt1(2) <> strUserNum Then
      If PUB_GetST03(txt1(2)) <> Pub_StrUserSt03 Then
         MsgBox "員工編號錯誤！", vbExclamation, "權限不足"
         Exit Sub
      End If
   End If

   tmpB = False
   For intI = 0 To 1
       txt1_Validate intI, tmpB
       If tmpB Then Exit Sub
   Next
   'Added by Lydia 2025/09/10
   If Trim(txtCode(0)) <> "" And Trim(txtCode(1)) <> "" Then
      If Trim(txtCode(2)) = "" Then txtCode(2) = "0"
      If Trim(txtCode(3)) = "" Then txtCode(3) = "00"
   End If
   'end 2025/09/10
   
  If QueryData(True) = False Then
  End If
End Sub

Private Function QueryData(Optional ByRef bolM As Boolean = True) As Boolean
Dim stVTB As String, strTmp As String
Dim rsRead As New ADODB.Recordset
Dim strS1 As String, strS2 As String
Dim inX As Integer
Dim stSQL As String, strTempName As String
Dim tmpArr As Variant
Dim strUsers As String
Dim stDept As String
Dim stNumList As String, stIdList
Dim ii As Integer, jj As Integer
Dim stUserID As String

QueryData = False
   
   Erase stNumList1
   
   'Modified by Lydia 2025/09/15 傳入本所案號，不限操作是否為提醒人員
   'stUserID = Txt1(2)
   stUserID = IIf(txt1(2) <> "", txt1(2), strUserNum)
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
   Else
      stDept = GetST15(stUserID)
   End If
   
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   'stNumList1(1):員工編號 (2)~(5):第2級期限管制人~第5級期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stIdList = Split(stNumList, ",")
   '去除重複編號
   stNumList = stIdList(LBound(stIdList))
   For ii = LBound(stIdList) + 1 To UBound(stIdList)
      For jj = LBound(stIdList) To ii - 1
         If stIdList(jj) = stIdList(ii) Then
            Exit For
         End If
      Next
      If ii = jj Then
         stNumList = stNumList & "," & stIdList(ii)
      End If
   Next
   stIdList = Split(stNumList, ",")

  
   '管制日期
   If txt1(0) <> "" And txt1(1) <> "" Then
      strS1 = strS1 & " and sc01>=" & DBDATE(txt1(0)) & " and sc01<=" & DBDATE(txt1(1))
   ElseIf txt1(0) <> "" Then
         strS1 = strS1 & " and sc01>=" & DBDATE(txt1(0))
       ElseIf txt1(1) <> "" Then
         strS1 = strS1 & " and sc01<=" & DBDATE(txt1(1))
   End If
   'Modified by Lydia 2016/06/08 當天之前未解除的行事曆一併出現
   If strS1 <> "" Then
       strS1 = " ((sc18 is null" & strS1 & ") or (sc18 is null and sc01<=" & strSrvDate(1) & "))"
   Else
       strS1 = " sc18 is null"
   End If
   
   '提醒人員
   If txt1(2) <> "" Then
      strS2 = strS2 & " and instr(sc03," & CNULL(txt1(2)) & ") > 0 "
   End If
   
   'Added by Lydia 2025/09/10 本所案號
   If Trim(txtCode(0)) <> "" And Trim(txtCode(1)) <> "" Then
      strS2 = strS2 & " and sc05='" & Trim(txtCode(0)) & "' and sc06='" & Trim(txtCode(1)) & "' and sc07='" & Trim(txtCode(2)) & "' and sc08='" & Trim(txtCode(3)) & "' "
   End If
   'end 2025/09/10
   
   'Modified by Lydia 2016/06/06
   'strTmp = "select sc01,sc02,sc03 from staff_calendar where sc18 is null " & strS1
   'Modified by Lydia 2020/09/14
   'strTmp = "select sc01,sc02,sc03 from staff_calendar where" & strS1
   strTmp = "select sc01,sc02 from staff_calendar where" & strS1
   stSQL = strTmp & strS2
   If txt1(2).Text <> "" Then 'Added by Lydia 2025/09/15 增加判斷
       tmpArr = Empty
       tmpArr = Split(stNumList, ",")
       stVTB = ""
       
      '登入者為2級且為3級以上主管
      '2級只看自己及部屬資料 ;'3級以上逾期資料
       For intI = 0 To UBound(tmpArr)
           If tmpArr(intI) <> "" Then
              If CNULL(txt1(2)) <> tmpArr(intI) Then
                 '2級主管
                 If InStr(stNumList1(2), tmpArr(intI)) > 0 Then
                    stVTB = stVTB & " Union " & strTmp & " and instr(sc03," & tmpArr(intI) & ") > 0 "
                    'Added by Lydia 2025/09/10 本所案號
                    If Trim(txtCode(0)) <> "" And Trim(txtCode(1)) <> "" Then
                       stVTB = stVTB & " and sc05='" & Trim(txtCode(0)) & "' and sc06='" & Trim(txtCode(1)) & "' and sc07='" & Trim(txtCode(2)) & "' and sc08='" & Trim(txtCode(3)) & "' "
                    End If
                    'end 2025/09/10
                 '3級主管只看逾期
                 ElseIf InStr(stNumList1(3) & stNumList1(4) & stNumList1(5), tmpArr(intI)) > 0 Then
                    stVTB = stVTB & " Union " & strTmp & " and sc01<" & strSrvDate(1) & " and instr(sc03," & tmpArr(intI) & ") > 0 "
                    'Added by Lydia 2025/09/10 本所案號
                    If Trim(txtCode(0)) <> "" And Trim(txtCode(1)) <> "" Then
                       stVTB = stVTB & " and sc05='" & Trim(txtCode(0)) & "' and sc06='" & Trim(txtCode(1)) & "' and sc07='" & Trim(txtCode(2)) & "' and sc08='" & Trim(txtCode(3)) & "' "
                    End If
                    'end 2025/09/10
                 End If
              End If
           End If
       Next
   End If 'Added by Lydia 2025/09/15
'Remove by Lydia 2020/09/14  改用DB的函數
'   stSQL = stSQL & stVTB & " order by 1,2,3 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'On Error GoTo ErrHnd1:
'   If intI = 1 Then
'       '提醒人員從員工編號轉姓名
'       Set rsRead = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
'
'       cnnConnection.BeginTrans
'       rsRead.MoveFirst
'       With rsRead
'          Do While Not .EOF
'             tmpArr = Empty: strUsers = ""
'             tmpArr = Split(.Fields("SC03"), ",")
'             For inX = 0 To UBound(tmpArr)
'                 If tmpArr(inX) <> "" Then
'                    'Modified by Lydia 2016/06/08
'                    'If ClsPDGetStaff(TmpArr(inX), strTempName) = True Then
'                    strTempName = GetStaffName(tmpArr(inX), True)
'                    If strTempName <> "" Then
'                      strUsers = strUsers & IIf(Len(strUsers) > 0, ",", "") & strTempName
'                    End If
'                 End If
'             Next
'             If strUsers <> "" Then
'                strExc(1) = "update rdatafactory set r004=" & CNULL(strUsers) & " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(rsRead.AbsolutePosition)
'                cnnConnection.Execute strExc(1), intI
'             End If
'             .MoveNext
'          Loop
'       End With
'       cnnConnection.CommitTrans
'On Error GoTo ErrorHand2:
'end 2020/09/14

       'Modified by Lydia 2016/06/08
       'strTmp = "select '' chk1,(sc01-19110000) 管制日期,sc02,r004 解除人員,sc04,decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508," & _
                "sc01,sc03,sc05,sc06,sc07,sc08,sc09,decode(sc10,'1','單次','2','每週','3','每月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n from staff_calendar,staff,RDataFactory where sc11=st01(+) " & _
                strS1 & " and to_char(sc01)=R001(+) and to_char(sc02)=R002(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "'"
       'Modified by Lydia 2016/06/28 +4.每3個月
       'Modified by Lydia 2020/09/14 改用DB的函數
       'strTmp = "select '' chk1,(sc01-19110000) 管制日期,sc02,r004 解除人員,sc04,decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508," & _
                "sc01,sc03,sc05,sc06,sc07,sc08,sc09,decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n from staff_calendar,staff,RDataFactory where sc11=st01(+) " & _
                "and to_char(sc01)=R001(+) and to_char(sc02)=R002(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "'"
       'Modified by Lydia 2023/07/28 外專-FCP專利連結案管制：抓專利連結通知PA177,來源收文號SC20
       'strTmp = "select '' chk1,(sc01-19110000) 管制日期,sc02,getstaffnamelist(sc03) 解除人員,sc04,decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508," & _
                "sc01,sc03,sc05,sc06,sc07,sc08,sc09,decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n from staff_calendar,staff " & _
                "where sc11=st01(+) and (sc01,sc02) in (" & stSQL & stVTB & ") "
       'end 2020/09/14
       strTmp = "select '' chk1,(sc01-19110000) 管制日期,sc02,getstaffnamelist(sc03) 解除人員,sc04,decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508," & _
                "sc01,sc03,sc05,sc06,sc07,sc08,sc09,decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n,SC20,PA177 from staff_calendar,staff,patent " & _
                "where sc11=st01(+) and (sc01,sc02) in (" & stSQL & stVTB & ") and sc05=pa01(+) and sc06=pa02(+) and sc07=pa03(+) and sc08=pa04(+) "
       strTmp = strTmp & " order by sc01,sc02 "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
       'Added by Lydia 2020/09/14
       If intI = 0 Then
            If bolM = True Then MsgBox "查無資料!!"
            LblCnt.Caption = "共 0 筆"
            GRD1.Clear
            Call SetGrd
       Else
       'end 2020/09/14
            GRD1.FixedCols = 0
            Set GRD1.Recordset = RsTemp
            Call SetGrd(RsTemp.RecordCount + 1)
            GRD1.FixedCols = 3
            QueryData = True
            LblCnt.Caption = "共 " & RsTemp.RecordCount & " 筆"
            Call SetColor
       End If 'Added by Lydia 2020/09/14
'Remove by Lydia 2020/09/14
'   Else
'       If bolM = True Then MsgBox "查無資料!!"
'       lblCnt.Caption = "共 0 筆"
'       GRD1.Clear
'       Call SetGrd
'   End If
'end 2020/09/14

   Exit Function
   
'Remove by Lydia 2020/09/14
'ErrHnd1:
'   If Err.Number > 0 Then
'      cnnConnection.RollbackTrans
'      MsgBox Err.Description
'      Exit Function
'   End If
'end 2020/09/14
ErrorHand2:
   If Err.Number > 0 Then
      MsgBox Err.Description
      Exit Function
   End If
   
End Function
'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = GRD1.Recordset.RecordCount
End Sub

Private Sub SetColor(Optional sHide As String = "N")
   Dim lngToday As Long, stType As String
   Dim ii As Integer, jj As Integer, dblCnt As Double
   

   With GRD1
   If .Rows > 1 Then
      .Visible = False
      lngToday = Val(strSrvDate(2))
      For ii = 1 To .Rows - 1
         .RowHeight(ii) = 255
         .row = ii
         '逾管控期限
         If .TextMatrix(ii, 1) < lngToday Then
            .TextMatrix(ii, 1) = "*" & .TextMatrix(ii, 1)
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         ElseIf .TextMatrix(ii, 1) = lngToday Then
            .TextMatrix(ii, 1) = "v" & .TextMatrix(ii, 1)
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   
   PUB_AddExcuteLog Me.Name 'Added by Lydia 2016/01/27
   
   Combo1.Clear
   '符號加在管制日期
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v): 表示當日期限"
   Combo1.AddItem "藍色: 表示點選資料"
   Combo1.ListIndex = 0
   Cmb2.Clear
   Cmb2.AddItem "一週"
   Cmb2.AddItem "二週"
   Cmb2.AddItem "三週"
   Cmb2.AddItem "一個月"
   Cmb2.AddItem "二個月"
   Cmb2.AddItem "三個月"
   Cmb2.Text = ""
   
   textCUID.BackColor = &H8000000F
   ClearField

   'Added by Lydia 2025/09/10
   For Each oText In txtCode
      oText.Text = Empty
   Next
   Frame1.BackColor = &H8000000F
   If m_Role <> "" Then
      Frame1.Visible = False
      cmdCP.Visible = False
   End If
   'end 2025/09/10

   'Added by Lydia 2025/09/10
   If strCode(1) <> "" And strCode(2) <> "" Then
      txtCode(0) = strCode(1)
      txtCode(1) = strCode(2)
      txtCode(2) = strCode(3)
      txtCode(3) = strCode(4)
   'Added by Lydia 2025/09/11 從其他畫面進入，限制案號
      Frame1.Enabled = False
      'Added by Lydia 2025/09/15
      Label2(0).Visible = False: Label2(1).Visible = False: Label2(2).Visible = False
      lblName.Visible = False: Cmb2.Visible = False
      txt1(0).Visible = False: txt1(1).Visible = False: txt1(2).Visible = False
      'end 2025/09/15
   Else
      Frame1.Enabled = True
      'Move by Lydia 2025/09/15 從上面移下來
      '預設為操作者；
      txt1(2) = strUserNum
      lblName.Caption = strUserName
      '管制日期預設抓系統日前、後2個工作天
      txt1(0) = ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1))
      txt1(1) = ChangeWStringToTString(CompWorkDay(3, strSrvDate(1)))
      'Added by Lydia 2025/09/15
      Label2(0).Visible = True: Label2(1).Visible = True: Label2(2).Visible = True
      lblName.Visible = True: Cmb2.Visible = True
      txt1(0).Visible = True: txt1(1).Visible = True: txt1(2).Visible = True
      'end 2025/09/15
   'end 2025/09/11
   End If
   'end 2025/09/10
   
   If QueryData(False) = False Then
   End If
   Me.SSTab1.Tab = 0
      
   'Added by Lydia 2021/04/20 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstSC04.Height = 800
   lstUsers(0).Height = 1500
   lstUsers(0).Width = 1180
   'lstUsers(0).ScrollBars = 2 '垂直捲軸
   lstUsers(1).Height = 1000
   lstUsers(1).Width = 1180

End Sub
Private Sub SetGrd(Optional ByRef iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Modified by Lydia 2023/07/28 +SC20,PA177
   arrGridHeadText = Array("解除", "管制日期", "流水號", "提醒人員", "事由", "本所案號", "SC01", "SC03", "SC05", "SC06", "SC07", "SC08", "SC09", "週期", "已解除", "SC11", "輸入人員", "SC20", "PA177")
   'Modified by Lydia 2016/06/28
   'arrGridHeadWidth = Array(500, 840, 640, 1500, 2000, 1320, 0, 0, 0, 0, 0, 0, 0, 500, 0, 0, 840)
   'Modified by Lydia 2023/07/28
   'arrGridHeadWidth = Array(500, 840, 600, 1500, 2000, 1320, 0, 0, 0, 0, 0, 0, 0, 720, 0, 0, 840)
   arrGridHeadWidth = Array(500, 840, 600, 1500, 2000, 1320, 0, 0, 0, 0, 0, 0, 0, 720, 0, 0, 840, 0, 0)
   
   With GRD1

       .Visible = False
       .Cols = UBound(arrGridHeadText) + 1
       .Rows = iR
       For iRow = 0 To .Cols - 1
          .row = 0
          .col = iRow
          .Text = arrGridHeadText(iRow)
          'Mark by Lydia 2023/07/28 改用模組
          'If Left(arrGridHeadText(iRow), 2) = "SC" Then
          '   mSC(Val(Right(arrGridHeadText(iRow), 2))) = iRow
          'ElseIf arrGridHeadText(iRow) = "流水號" Then
          '      mSC(2) = iRow
          'ElseIf arrGridHeadText(iRow) = "事由" Then
          '      mSC(4) = iRow
          'ElseIf arrGridHeadText(iRow) = "週期" Then
          '      mSC(10) = iRow
          'ElseIf arrGridHeadText(iRow) = "已解除" Then
          '      colSC18 = iRow
          'End If
          'end 2023/07/28
          .ColWidth(iRow) = arrGridHeadWidth(iRow)
          .CellAlignment = flexAlignCenterCenter
       Next
       'Added by Lydia 2023/07/28
       If colSC20 = 0 Then
          For intI = 1 To UBound(mSC)
            If intI = 2 Then
               mSC(intI) = PUB_MGridGetId("流水號", GRD1)
            ElseIf intI = 4 Then
               mSC(intI) = PUB_MGridGetId("事由", GRD1)
            ElseIf intI = 10 Then
               mSC(intI) = PUB_MGridGetId("週期", GRD1)
            Else
               mSC(intI) = PUB_MGridGetId("SC" & Format(intI, "00"), GRD1)
            End If
          Next intI
          colSC18 = PUB_MGridGetId("已解除", GRD1)
          colSC20 = PUB_MGridGetId("SC20", GRD1)
          colPA177 = PUB_MGridGetId("PA177", GRD1)
       End If
       'end 2023/07/28
       
       For intI = 1 To iR - 1
         .row = intI
         For iRow = 0 To .Cols - 1
           .col = iRow
           If iRow < 3 Then
              .CellBackColor = QBColor(15)
           End If
           '流水號
           If iRow = 2 Then
              .CellAlignment = flexAlignRightCenter
           '週期，解除
           ElseIf iRow = 0 Or iRow = 13 Or iRow = 14 Then
              .CellAlignment = flexAlignCenterCenter
           End If

         Next iRow
       Next intI


       .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2017/1/12
   If Not bolUnloading Then
      'Modified by Lydia 2025/09/10 排除另外呼叫 And TypeName(m_PrevForm) = "Nothing"
      If m_Role <> "F41" And TypeName(m_PrevForm) = "Nothing" Then 'Added by Lydia 2020/01/15 排除業務拓展小組的身份進入
            '自動執行:程序大項工作期限通知
            If Pub_StrUserSt03 = "M51" Or _
               Pub_GetSpecMan("外專告准程序") = strUserNum Or _
               Pub_GetSpecMan("外專告准程序主管") = strUserNum Then
               strSql = "select * from executelog where el01='frm060210' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI <> 1 Then
                  pub_bolInformCheck = True
                  Load frm060210
                  frm060210.cmdQuery_Click
                  pub_bolInformCheck = False
               End If
            End If
      End If
   End If
   '2017/1/12 END
   
   'Added by Lydia 2025/09/10
   If TypeName(m_PrevForm) <> "Nothing" And bolPreClose = False Then
      m_PrevForm.Show
   End If
   'end 2025/09/10
   
   Set frm060209 = Nothing
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal bolMsg As Boolean = True) As Boolean
   
   Dim adoRst As New ADODB.Recordset
   Dim stCon As String
   
   If p_iWay = -1 Or p_iWay = 1 Or p_iWay = 0 Then
      If txtSC(1) <> "" Then stCon = stCon & " and SC01=" & CNULL(DBDATE(txtSC(1)), True)
      If txtSC(2) <> "" Then stCon = stCon & " and SC02=" & CNULL(txtSC(2), True)
   End If
   
   strExc(0) = "SELECT * FROM Staff_Calendar WHERE rownum<2 "
   Select Case p_iWay
      '尋找
      Case 0: strExc(0) = strExc(0) & stCon
      '首筆
      Case -2: strExc(0) = strExc(0) & stCon & " order by SC01,SC02"
      '前一筆
      Case -1
         If stCon <> "" Then
            strExc(0) = strExc(0) & " and SC01||lpad(SC02,4,'0') <'" & DBDATE(txtSC(1)) & Format(txtSC(2), "0000") & "' order by SC01||lpad(SC02,4,'0') DESC"
         End If
      '後一筆
      Case 1
         If stCon <> "" Then
            strExc(0) = strExc(0) & " and SC01||lpad(SC02,4,'0') >'" & DBDATE(txtSC(1)) & Format(txtSC(2), "0000") & "' order by SC01||lpad(SC02,4,'0') ASC"
         End If
      '末筆
      Case 2
         strExc(0) = "SELECT * FROM Staff_Calendar where 1=1" & stCon & " order by SC01 DESC,SC02 DESC"
   End Select
      
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         If bolMsg = True Then MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         If bolMsg = True Then MsgBox "已經是最後筆！", vbInformation
      Else
         If bolMsg = True Then MsgBox "查無資料！", vbInformation
      End If
   End If
   
   
   SetCtrlReadOnly True
   Set adoRst = Nothing

End Function
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
Dim CUID(1 To 6) As String
Dim tmpArr As Variant, tmpBol As Boolean
   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtSC
            idx = oText.Index
            If idx = 1 Then
               oText.Text = ChangeWStringToTString(p_Rst.Fields(idx - 1))
            Else
                'Modified by Lydia 2017/07/18 + chgsql 去除單引號
               oText.Text = ChgSQL("" & p_Rst.Fields(idx - 1))
            End If
            oText.Tag = oText.Text
         Next
         If txtSC(5).Text <> "" And txtSC(6).Text <> "" Then
            If GetPdata(txtSC(5).Text, txtSC(6).Text, txtSC(7).Text, txtSC(8).Text, False) Then
            End If
         End If
         If txtSC(4) <> "" Then
            lstSC04.Clear
            tmpArr = Empty
            tmpArr = Split(txtSC(4), vbCrLf)
            For idx = 0 To UBound(tmpArr)
               If tmpArr(idx) <> "" Then
                  lstSC04.AddItem Trim(tmpArr(idx))
               End If
            Next
         End If
         
         CUID(1) = "" & .Fields("SC11")
         CUID(2) = "" & .Fields("SC12")
         CUID(3) = "" & .Fields("SC13")
         CUID(4) = "" & .Fields("SC14")
         CUID(5) = "" & .Fields("SC15")
         CUID(6) = "" & .Fields("SC16")
         mSC11 = "" & .Fields("SC11")
         mSC18 = "" & .Fields("SC18")
                  
         'Added by Lydia 2025/09/10
         If m_Role = "" And Trim(txtSC(5)) <> "" And Trim(txtSC(6)) <> "" Then
            cmdCP.Visible = True
         Else
            cmdCP.Visible = False
         End If
         'end 2025/09/10
         
         If txtSC(3) <> "" Then
            SetlstUsers 0, txtSC(3)
         End If
         If txtSC(9) <> "" Then
            SetlstUsers 1, txtSC(9)
         End If
      End If
   End With
   UpdateCUID CUID, textCUID

End Sub
'判斷是否能修改或解除
Private Function JudgeRight(ByRef bolMsg As Boolean, ByVal iSC11 As String, ByVal iSC09 As String) As Boolean
Dim tmpArr As Variant
Dim idR As Integer
Dim strAr As String
    
    JudgeRight = False
    If iSC11 <> "" Then
       bolMsg = True
       
       If Pub_StrUserSt03 = "M51" Then
            JudgeRight = True
       '建檔人=解除人員
       ElseIf UCase(iSC11) = UCase(strUserNum) Then
            JudgeRight = True: bolMsg = False
       '可解除人員
       ElseIf InStr(UCase(iSC09), UCase(strUserNum)) > 0 Then
            JudgeRight = True
       '可解除人員的主管
       Else
            strAr = UCase(iSC09) & "," & UCase(iSC11)
            tmpArr = Empty
            tmpArr = Split(strAr, ",")
            For idR = 0 To UBound(tmpArr)
                If tmpArr(idR) <> "" Then
                   If PUB_GetST52(tmpArr(idR), strUserNum) = True Then
                      JudgeRight = True
                      Exit For
                   End If
                End If
            Next
       End If
    End If

End Function
Private Sub ClearField()

   For Each oText In txtSC
      oText.Text = Empty
      oText.Tag = Empty
   Next

   textCUID = ""
   lstSC04.Clear
   'Cmb1.Text = "" 'Remove by Lydia 2021/04/20
   Cmb1.Clear
   mSC11 = ""
   lstUsers(0).Clear
   lstUsers(1).Clear
   lblFC(0) = "": lblFC(1) = ""
   'Added by Lydia 2021/04/21
   lstUsers(0).Tag = ""
   lstUsers(1).Tag = ""

End Sub

Private Function ComposeList(oList As ListBox) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         iPos = InStr(oList.List(intI), Chr(1))
         If iPos > 0 Then
            stItem = Left(oList.List(intI), iPos - 1)
         Else
            stItem = oList.List(intI)
         End If
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & vbCrLf & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Sub GRD1_DblClick()
   If GRD1.MouseRow > 0 And GRD1.TextMatrix(GRD1.row, mSC(1)) <> "" Then
      txtSC(1) = ChangeWStringToTString(GRD1.TextMatrix(GRD1.row, mSC(1)))
      txtSC(2) = GRD1.TextMatrix(GRD1.row, mSC(2))
      If ShowRecord(0, True) Then
         bolDbClick = True 'Added by Lydia 2021/04/20
         SSTab1.Tab = 1
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   SWPRow = GRD1.MouseRow
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "解除" Then
      If InStr("流水號,週期", Me.GRD1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub grd1_SelChange()
   Dim ii As Integer, ColorL As Long, ColorN As Long
   Dim pRow As Integer, pCol As Integer
   Dim tmpBol As Boolean

   With GRD1
      If .MouseRow > 0 Then
         If .MouseCol = 0 Then
            pRow = .MouseRow
            pCol = .MouseCol
            .row = pRow
            .col = pCol
            If .Text = "" And .TextMatrix(pRow, colSC18) = "" Then
               If DetailCancel(pRow) Then
                  .RowHeight(pRow) = 0
               End If
            End If
         End If
      End If
   End With
   
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      If bolDbClick = False Then 'Added by Lydia 2021/04/20 避免重複設置資料
            '若有資料
            If Me.GRD1.Rows > 1 And SWPRow > 0 And Me.GRD1.TextMatrix(Val("0" & SWPRow), mSC(1)) <> "" And Me.GRD1.RowHeight(SWPRow) > 0 Then
               '若點選的那筆無資料, 則退出函式
               If Me.GRD1.TextMatrix(Val("0" & SWPRow), mSC(1)) = "" Then SSTab1.Tab = 0: Exit Sub
                  txtSC(1) = ChangeWStringToTString(GRD1.TextMatrix(GRD1.row, mSC(1)))
                  txtSC(2) = GRD1.TextMatrix(GRD1.row, mSC(2))
                  If ShowRecord(0, True) Then
                     SSTab1.Tab = 1
                  End If
            Else
               ClearField
            End If
      End If 'Added by Lydia 2021/04/20
'Added by Lydia 2016/03/01
      cmdSearch.Enabled = False
   Else
      cmdSearch.Enabled = True
'end 2016/03/01
      bolDbClick = False 'Added by Lydia 2021/04/20
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
      Case 2
         KeyAscii = UpperCase(KeyAscii)
      Case Else
         KeyAscii = Pub_NumAscii(KeyAscii)
  End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   Select Case Index

      Case 0, 1
         If txt1(Index) = "" Then
            MsgBox "管制日期不可空白!", vbCritical, "輸入錯誤"
            GoTo JumpCancel
         Else
            If CheckIsTaiwanDate(txt1(Index)) = False Then
                GoTo JumpCancel
            Else
                If txt1(0) <> "" And txt1(1) <> "" And txt1(0) > txt1(1) Then
                   MsgBox "管制日期止不可小於管制日期起!", vbCritical, "輸入錯誤"
                   GoTo JumpCancel
                End If
            End If
         End If

      Case 2
         If txt1(Index) <> "" Then
            If Len(txt1(Index)) = 5 Then
               'Modified by Lydia 2016/08/16 遇到離職人員不彈訊息
               'If ClsPDGetStaff(Txt1(Index), strExc(1)) = True Then
               '   lblName = strExc(1)
               'End If
               lblName = GetStaffName(txt1(Index), True)
            End If
            If lblName = "" Then
               MsgBox "員工編號輸入錯誤！", vbExclamation
               Cancel = True
            End If
         Else
            lblName = ""
         End If
   End Select
   
   If Cancel = False Then
      If txt1(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txt1(Index), iLen) Then
            GoTo JumpCancel
         End If
      End If
   End If
   Exit Sub
   
JumpCancel:
   txt1_GotFocus Index
   Cancel = True
End Sub

Private Sub txtSC_GotFocus(Index As Integer)
   TextInverse txtSC(Index)
End Sub


Private Function GetPdata(ByVal Cc01 As String, Cc02 As String, Optional ByVal Cc03 As String, Optional ByVal Cc04 As String, Optional ByVal bolMsg As Boolean = True) As Boolean
Dim inX As Integer
Dim Str01 As String, Str02 As String, Str03 As String, sDate01, sDate02 As String

GetPdata = False
Cmb1.Clear
lblFC(0) = "": lblFC(1) = ""

If Cc03 = "" Then Cc03 = "0"
If Cc04 = "" Then Cc04 = "00"

Dim strSql As String, intCaseKind As Integer


If ClsPDGetSystemKind(Cc01, intCaseKind) Then

   Select Case intCaseKind
      Case 專利
         strSql = "select pa05,pa06,pa07,pa108,pa136,pa75,NVL(FA05,NVL(FA04,FA06)) from patent,fagent " & _
            "where pa01=" & CNULL(Cc01) & " and pa02=" & CNULL(Cc02) & " and pa03=" & CNULL(Cc03) & " and pa04=" & CNULL(Cc04) & _
            " and substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) "
      Case 商標
         strSql = "select tm05,tm06,tm07,tm57,tm73,tm44,NVL(FA05,NVL(FA04,FA06)) from trademark,fagent " & _
            "where tm01=" & CNULL(Cc01) & " and tm02=" & CNULL(Cc02) & " and tm03=" & CNULL(Cc03) & " and tm04=" & CNULL(Cc04) & _
            " and substr(TM44,1,8)=FA01(+) And substr(TM44,9,1)=FA02(+) "
      Case 法務
         strSql = "select lc05,lc06,lc07,lc34,lc36,lc22,NVL(FA05,NVL(FA04,FA06)) from lawcase,FAGENT " & _
            "where lc01=" & CNULL(Cc01) & " and lc02=" & CNULL(Cc02) & " and lc03=" & CNULL(Cc03) & " and lc04=" & CNULL(Cc04) & _
            " and substr(LC22,1,8)=FA01(+) And substr(LC22,9,1)=FA02(+) "
      Case 顧問
         strSql = "select hc06,'','',hc19,hc20,'','' from hirecase " & _
            "where hc01=" & CNULL(Cc01) & " and hc02=" & CNULL(Cc02) & " and hc03=" & CNULL(Cc03) & " and hc04=" & CNULL(Cc04)
      Case Else
         strSql = "select sp05,sp06,sp07,sp61,sp68,sp26,NVL(FA05,NVL(FA04,FA06)) from servicepractice,FAGENT " & _
            "where sp01=" & CNULL(Cc01) & " and sp02=" & CNULL(Cc02) & " and sp03=" & CNULL(Cc03) & " and sp04=" & CNULL(Cc04) & _
            " and substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) "
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       txtSC(7) = Cc03: txtSC(8) = Cc04
       Cmb1.AddItem "中 : " & Trim(RsTemp(0))
       Cmb1.AddItem "英 : " & Trim(RsTemp(1))
       'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
       Cmb1.AddItem "外 : " & Trim(RsTemp(2))
       Cmb1.Text = "中 : " & Trim(RsTemp(0))
       lblFC(0) = ChangeCustomerS("" & RsTemp(5))
       lblFC(1) = "" & Trim(RsTemp(6))
       GetPdata = True
   Else
       If bolMsg = True Then ShowMsg MsgText(9141)
   End If
End If

End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtSC
      oText.Locked = bLocked
   Next
End Sub

' 更新 Create 及 Update 的人
'Modified by Lydia 2021/04/20
'Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As TextBox)
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Control)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   Dim arrName 'Added by Lydia 2021/04/23
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Lydia 2021/04/23
   
   If p_stNums <> "" Then
      'Modified by Lydia 2021/04/23 改寫法
'      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         arrID = Split(p_stNums, ",")
'         With RsTemp
'         '照原順序排
'         For intI = UBound(arrID) To LBound(arrID) Step -1
'            .MoveFirst
'            Do While Not .EOF
'               If .Fields("st01") = arrID(intI) Then
'                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
'                  '員工編號已可非數字需做轉換
'                  lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'                  .MoveLast
'               End If
'               .MoveNext
'            Loop
'         Next
'         End With
'      End If
      strExc(0) = "select getstaffnamelist('" & p_stNums & "') from dual"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          arrID = Split(p_stNums, ",")
          arrName = Split("" & RsTemp.Fields(0), ",")
          For intI = UBound(arrID) To LBound(arrID) Step -1
               lstUsers(p_idx).AddItem arrName(intI), 0
               'Form 2.0的Listbox沒有ItemData,改放在.Tag; 讀取用PUB_GetItemData
               lstUsers(p_idx).Tag = arrID(intI) & IIf(lstUsers(p_idx).Tag <> "", ",", "") & lstUsers(p_idx).Tag
          Next intI
      End If
'-------'Modified by Lydia 2021/04/23 改寫法
   End If
End Sub

'Added by Lydia 2025/09/10 共同查詢->進度檔
Private Sub cmdCP_Click()
Dim StrTag As String

   If Trim(txtSC(5)) = "" Or Trim(txtSC(6)) = "" Then Exit Sub
   StrTag = txtSC(5) & "-" & txtSC(6) & "-" & txtSC(7) & "-" & txtSC(8)

   If strCode(1) = "" And strCode(1) = "" Then
      If PUB_CheckFormExist("frm100101_2") Then
         MsgBox "請先關閉共同查詢之〔案件資料及案件進度查詢〕！", vbCritical + vbOKOnly
         Exit Sub
      End If
      
   End If
   Screen.MousePointer = vbHourglass
   '看完整進度
   If strCode(1) = "" And strCode(2) = "" Then
      frm100101_2.SetParent Me
   End If
   frm100101_2.Show
   frm100101_2.Tag = StrTag
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   
   bolPreClose = False
   If TypeName(m_PrevForm) = "frm090202_2" Then
      bolPreClose = True
      Unload frm090202_2
      Unload Me
   ElseIf TypeName(m_PrevForm) = "frm100101_2" Then
      bolPreClose = True
      Unload Me
   End If

End Sub

'Added by Lydia 2025/09/10
Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If Trim(txtCode(Index)) <> "" And txtCode(Index) <> "FCP" And txtCode(Index) <> "P" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            txtCode(Index).SetFocus
            Cancel = True
        End If
    End If
End Sub

