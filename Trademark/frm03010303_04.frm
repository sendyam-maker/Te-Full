VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010303_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品及服務資料輸入"
   ClientHeight    =   5720
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9450
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0C0FF&
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6900
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   28
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   290
      Left            =   8220
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   2730
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "從右邊本所案號複製過來"
      Height          =   290
      Left            =   5100
      TabIndex        =   26
      Top             =   420
      Width           =   2385
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部複製(&A)"
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   25
      Top             =   0
      Width           =   1185
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "類別複製(&C)"
      Height          =   375
      Index           =   4
      Left            =   2460
      TabIndex        =   24
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   0
      Left            =   7500
      MaxLength       =   3
      TabIndex        =   6
      Top             =   420
      Width           =   405
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   1
      Left            =   7995
      MaxLength       =   6
      TabIndex        =   7
      Top             =   420
      Width           =   645
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   2
      Left            =   8745
      MaxLength       =   1
      TabIndex        =   8
      Top             =   420
      Width           =   180
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   3
      Left            =   9030
      MaxLength       =   2
      TabIndex        =   9
      Top             =   420
      Width           =   285
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "CFT 定稿列印(&P)"
      Height          =   375
      Index           =   2
      Left            =   5220
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2505
      Left            =   90
      TabIndex        =   11
      Top             =   3180
      Width           =   9105
      _ExtentX        =   16051
      _ExtentY        =   4427
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "中文-1"
      TabPicture(0)   =   "frm03010303_04.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "中文-2"
      TabPicture(1)   =   "frm03010303_04.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "英文-1"
      TabPicture(2)   =   "frm03010303_04.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "英文-2"
      TabPicture(3)   =   "frm03010303_04.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "日文-1"
      TabPicture(4)   =   "frm03010303_04.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "日文-2"
      TabPicture(5)   =   "frm03010303_04.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt1(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "附件區"
      TabPicture(6)   =   "frm03010303_04.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "lblAtt"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "CommonDialog1"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "MSHFlexGrid1"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cmdSaveAtt"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdRemAtt"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "cmdAddAtt"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "cmdOpenAtt"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).ControlCount=   7
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   390
         Width           =   615
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "新增"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         TabIndex        =   20
         Top             =   390
         Width           =   615
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2055
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Width           =   615
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   390
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   750
         Width           =   8790
         _ExtentX        =   15522
         _ExtentY        =   2999
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V| 檔案名稱| 最後修改時間|上傳時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2820
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblAtt 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "本頁籤不分類別且操作會直接更新!!!"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   3480
         TabIndex        =   23
         Top             =   450
         Width           =   3150
      End
      Begin MSForms.TextBox txt1 
         Height          =   2115
         Index           =   5
         Left            =   -74960
         TabIndex        =   5
         Top             =   330
         Width           =   9015
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15901;3731"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2115
         Index           =   4
         Left            =   -74960
         TabIndex        =   4
         Top             =   330
         Width           =   9015
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15901;3731"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2115
         Index           =   3
         Left            =   -74960
         TabIndex        =   3
         Top             =   330
         Width           =   9015
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15901;3731"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2115
         Index           =   2
         Left            =   -74960
         TabIndex        =   2
         Top             =   330
         Width           =   9015
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15901;3731"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2115
         Index           =   1
         Left            =   -74960
         TabIndex        =   1
         Top             =   360
         Width           =   9015
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15901;3731"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2025
         Index           =   0
         Left            =   -74955
         TabIndex        =   0
         Top             =   360
         Width           =   8940
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15769;3572"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1935
      Left            =   90
      TabIndex        =   14
      Top             =   720
      Width           =   9105
      _ExtentX        =   16051
      _ExtentY        =   3404
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   3
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "延展欄為N,代表此類別未延展。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2550
      TabIndex        =   17
      Top             =   510
      Width           =   2505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   $"frm03010303_04.frx":00C4
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   2430
   End
   Begin VB.Line Line1 
      X1              =   9270
      X2              =   7740
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label2 
      Caption         =   $"frm03010303_04.frx":00FA
      Height          =   405
      Left            =   90
      TabIndex        =   15
      Top             =   2730
      Width           =   8055
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   960
      TabIndex        =   13
      Top             =   510
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   30
      TabIndex        =   12
      Top             =   510
      Width           =   900
   End
End
Attribute VB_Name = "frm03010303_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/13 改成Form2.0 ; txt1(index)、grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'add by nick 2004/09/16 新增加功能
Option Explicit

Public UpForm As Form
Public TGKey As String
Public AllClass As String
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim NowRow As Integer
Dim IsSave As Boolean
Public ChkCht As Boolean
Public ChkEng As Boolean
Public ChkJpn As Boolean
Public PubMsg As String
'Added by Morgan 2023/2/15
Dim m_SaveFolder As String
Dim m_AttachPath As String
'Add By Sindy 2024/12/20
Dim m_CompareGoods As String
Public m_TM08 As String
Public m_TM15 As String
'2024/12/20 END


Private Sub cmd_Click()
If IsSave = True Then IsSave = False

'Added by Lydia 2021/09/22 (類別輸入)檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Sub
End If
'end 2021/09/22
            
With grd1
    .row = NowRow
    .col = 3 '2
    .Text = txt1(0)
    .col = 4 '3
    .Text = txt1(1)
    .col = 5 '4
    .Text = txt1(2)
    'add by nickc 2008/03/28 加欄位
    .col = 6 '5
     .Text = txt1(3)
    .col = 7 '6
    .Text = txt1(4).Text
    .col = 8 '7
     .Text = txt1(5).Text
End With
End Sub

Private Sub cmd2_Click()
   Dim oCounts As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'add by nickc 2006/06/21
   Dim IsHaveClass As Boolean
If Trim(txt2(0).Text) <> "" And Trim(txt2(1).Text) <> "" Then
    'add by nickc 2006/06/21
    MsgBox "相同類別資料才可以複製！", vbInformation, "注意！"
    IsHaveClass = False
    With AdoRecordSet3
        CheckOC3
        .CursorLocation = adUseClient
        .Open "select count(*) from Tmgoods where tg01='" & Trim(txt2(0).Text) & "' and tg02='" & Trim(txt2(1).Text) & "' and tg03='" & IIf(Trim(txt2(2).Text) = "", "0", Trim(txt2(2).Text)) & "' and tg04='" & IIf(Trim(txt2(3).Text) = "", "00", Trim(txt2(3).Text)) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
        If .Fields(0) = 0 Then
            strTit = "資料檢核"
            strMsg = txt2(0).Text & "-" & txt2(1).Text & "-" & IIf(Trim(txt2(2).Text) = "", "0", Trim(txt2(2).Text)) & "-" & IIf(Trim(txt2(3).Text) = "", "0", Trim(txt2(3).Text)) & "尚未建立商品及服務！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txt1(0).SetFocus
            Exit Sub
        End If
        For oCounts = 1 To grd1.Rows - 1
            grd1.row = oCounts
            ChgData oCounts
            grd1.col = 1
            CheckOC3
            .CursorLocation = adUseClient
            'edit by nickc 2006/06/21
            '.Open "select * from Tmgoods where tg01='" & Trim(txt2(0).Text) & "' and tg02='" & Trim(txt2(1).Text) & "' and tg03='" & IIf(Trim(txt2(2).Text) = "", "0", Trim(txt2(2).Text)) & "' and tg04='" & IIf(Trim(txt2(3).Text) = "", "00", Trim(txt2(3).Text)) & "' and tg05='" & Grd1.Text & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            'edit by nickc 2008/03/26 解決超過 4000的問題
            '.Open "select * from Tmgoods where tg01='" & Trim(txt2(0).Text) & "' and tg02='" & Trim(txt2(1).Text) & "' and tg03='" & IIf(Trim(txt2(2).Text) = "", "0", Trim(txt2(2).Text)) & "' and tg04='" & IIf(Trim(txt2(3).Text) = "", "00", Trim(txt2(3).Text)) & "' and tg05='" & grd1.Text & "' and length(tg06||tg07||tg08)>0 ", cnnConnection, adOpenStatic, adLockReadOnly
            .Open "select * from Tmgoods where tg01='" & Trim(txt2(0).Text) & "' and tg02='" & Trim(txt2(1).Text) & "' and tg03='" & IIf(Trim(txt2(2).Text) = "", "0", Trim(txt2(2).Text)) & "' and tg04='" & IIf(Trim(txt2(3).Text) = "", "00", Trim(txt2(3).Text)) & "' and tg05='" & grd1.Text & "' and nvl(length(tg06),0)+nvl(length(tg07),0)+nvl(length(tg08),0)>0 ", cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 Then
                'add by nickc 2006/06/21
                IsHaveClass = True
                grd1.col = 3 '2
                grd1.Text = CheckStr(.Fields("TG06").Value)
                'edit by nickc 2008/03/28 加欄位
                'grd1.col = 3
                'grd1.Text = CheckStr(.Fields("TG07").Value)
                'grd1.col = 4
                'grd1.Text = CheckStr(.Fields("TG08").Value)
                grd1.col = 4 '3
                grd1.Text = CheckStr(.Fields("TG15").Value)
                grd1.col = 5 '4
                grd1.Text = CheckStr(.Fields("TG07").Value)
                grd1.col = 6 '5
                grd1.Text = CheckStr(.Fields("TG16").Value)
                grd1.col = 7 '6
                grd1.Text = CheckStr(.Fields("TG08").Value)
                grd1.col = 8 '7
                grd1.Text = CheckStr(.Fields("TG17").Value)
            End If
        Next oCounts
        CheckOC3
        ChgData 1
        If IsHaveClass = True Then
            strTit = "恭喜"
            strMsg = "複製完成！"
        Else
            strTit = "錯誤"
            strMsg = "未找到相同類別！"
        End If
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
    End With
End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim oStrSQL As String
   Dim i930922 As Integer
   Dim IsHaveAsk As Boolean
   Dim m_copytext As String
Select Case Index
Case 0
    m_copytext = ""  'Added by Lydia 2021/09/22
    '存檔
    With grd1
        For i930922 = 1 To .Rows - 1
            ChgData (i930922)
            'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
            If PUB_ChkUniText(Me, , True, "TextBox") = False Then
                 Exit Sub
            End If
            'end 2021/08/20
            
            'edit by nick 2004/10/15 改成中文或英文或日文有輸其一即可
            'If Trim(txt1(0).Text) = "" Then
            'edit by nickc 2008/03/28 加欄位
            'If Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" And Trim(txt1(2).Text) = "" Then
            '2010/11/19 MODIFY BY SONIA 阿蓮說CFT不管制(CFT-010455由17改19類)
            'If Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" And Trim(txt1(2).Text) = "" And Trim(txt1(3).Text) = "" And Trim(txt1(4).Text) = "" And Trim(txt1(5).Text) = "" Then
            If grd1.TextMatrix(i930922, 2) <> "N" Then 'N.不延展商品 Add By Sindy 2014/5/23 +if
            '2014/5/23 END
               If m_TM01 <> "CFT" And Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" And Trim(txt1(2).Text) = "" And Trim(txt1(3).Text) = "" And Trim(txt1(4).Text) = "" And Trim(txt1(5).Text) = "" Then
                    strTit = "資料檢核"
                    'edit by nick 2004/10/15
                    'strMsg = "中文一定要輸入！"
                    strMsg = "中文、英文、日文最少輸一種！"
                    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                    txt1(0).SetFocus
                    Exit Sub
               End If
               'edit by nickc 2008/03/28 加欄位
               'If ChkCht = True And Trim(txt1(0).Text) = "" Then
               If ChkCht = True And Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" Then
                    strTit = "資料檢核"
                    nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
                    txt1(0).SetFocus
                    Exit Sub
               End If
               'edit by nickc 2008/03/28 加欄位
               'If ChkEng = True And Trim(txt1(1).Text) = "" Then
               If ChkEng = True And Trim(txt1(2).Text) = "" And Trim(txt1(3).Text) = "" Then
                    strTit = "資料檢核"
                    nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
                    'edit by nickc 2008/03/28 加欄位
                    'txt1(1).SetFocus
                    txt1(2).SetFocus
                    Exit Sub
               End If
               'edit by nickc 2008/03/28 加欄位
               'If ChkJpn = True And Trim(txt1(2).Text) = "" Then
               If ChkJpn = True And Trim(txt1(4).Text) = "" And Trim(txt1(5).Text) = "" Then
                    strTit = "資料檢核"
                    nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
                    'edit by nickc 2008/03/28 加欄位
                    'txt1(2).SetFocus
                    txt1(4).SetFocus
                    Exit Sub
               End If
            End If
            'add by nickc 2006/06/13檢查有無問號
            'Modified by Lydia 2021/09/22 記錄有?的類別
'            If InStr(1, txt1(0), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(1), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(2), "?") <> 0 Then IsHaveAsk = True
'            'add by nickc 2008/03/28 加欄位
'            If InStr(1, txt1(3), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(4), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(5), "?") <> 0 Then IsHaveAsk = True
            If InStr(1, txt1(0), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-中文1"
                IsHaveAsk = True
            End If
            If InStr(1, txt1(1), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-中文2"
                IsHaveAsk = True
            End If
            If InStr(1, txt1(2), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-英文1"
                IsHaveAsk = True
            End If
            If InStr(1, txt1(3), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-英文2"
                IsHaveAsk = True
            End If
            If InStr(1, txt1(4), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-日文1"
                IsHaveAsk = True
            End If
            If InStr(1, txt1(5), "?") <> 0 Then
                m_copytext = m_copytext & "、" & grd1.TextMatrix(i930922, 1) & " 類-日文2"
                IsHaveAsk = True
            End If
            'end 2021/09/22
            
            'edit by nickc 2007/03/27 放大
            'If CheckLengthIsOK(txt1(0), 2000) = False Then SSTab1.Tab = 0: txt1_GotFocus 0: Exit Sub
            'If CheckLengthIsOK(txt1(1), 2000) = False Then SSTab1.Tab = 1: txt1_GotFocus 1: Exit Sub
            'If CheckLengthIsOK(txt1(2), 2000) = False Then SSTab1.Tab = 2: txt1_GotFocus 2: Exit Sub
            If CheckLengthIsOK(txt1(0), 4000) = False Then SSTab1.Tab = 0: txt1_GotFocus 0: Exit Sub
            If CheckLengthIsOK(txt1(1), 4000) = False Then SSTab1.Tab = 1: txt1_GotFocus 1: Exit Sub
            If CheckLengthIsOK(txt1(2), 4000) = False Then SSTab1.Tab = 2: txt1_GotFocus 2: Exit Sub
            'add by nickc 2008/03/28 加欄位
            If CheckLengthIsOK(txt1(3), 4000) = False Then SSTab1.Tab = 3: txt1_GotFocus 3: Exit Sub
            If CheckLengthIsOK(txt1(4), 4000) = False Then SSTab1.Tab = 4: txt1_GotFocus 4: Exit Sub
            If CheckLengthIsOK(txt1(5), 4000) = False Then SSTab1.Tab = 5: txt1_GotFocus 5: Exit Sub
        Next i930922
        'add by nickc 2006/06/13若有問號，問一下要不要繼續
        If IsHaveAsk = True Then
            'Modififed by Lydia 2021/09/22 記錄有?的類別
            'If MsgBox("輸入的名稱含有問號，" & vbCrLf & "　　　　若是正常請按　是　繼續" & vbCrLf & "　　　　若不正常請按　否　修正！", vbYesNo) = vbNo Then
            If MsgBox("以下類別：" & Mid(m_copytext, 2) & vbCrLf & vbCrLf & "輸入的名稱含有問號，" & vbCrLf & "　　　　若是正常請按　是　繼續" & vbCrLf & "　　　　若不正常請按　否　修正！", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    'Save
        Dim IsCanUpdate As Boolean
        Me.Enabled = False
        grd1.Visible = False
        Screen.MousePointer = vbHourglass
        grd1.MousePointer = flexHourglass
        On Error GoTo ShowErr
        cnnConnection.BeginTrans
        
        For i930922 = 1 To .Rows - 1
            ChgData (i930922)
            .col = 1
            oStrSQL = "select * from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
                'add by nickc 2006/06/13
                IsCanUpdate = False
                If CheckStr(AdoRecordSet3.Fields("tg06")) <> txt1(0) Then
                    IsCanUpdate = True
                End If
                'edit by nickc 2008/03/28 加欄位
                'If CheckStr(AdoRecordSet3.Fields("tg07")) <> txt1(1) Then
                '    IsCanUpdate = True
                'End If
                'If CheckStr(AdoRecordSet3.Fields("tg08")) <> txt1(2) Then
                '    IsCanUpdate = True
                'End If
                If CheckStr(AdoRecordSet3.Fields("tg15")) <> txt1(1) Then
                    IsCanUpdate = True
                End If
                If CheckStr(AdoRecordSet3.Fields("tg07")) <> txt1(2) Then
                    IsCanUpdate = True
                End If
                If CheckStr(AdoRecordSet3.Fields("tg16")) <> txt1(3) Then
                    IsCanUpdate = True
                End If
                If CheckStr(AdoRecordSet3.Fields("tg08")) <> txt1(4) Then
                    IsCanUpdate = True
                End If
                If CheckStr(AdoRecordSet3.Fields("tg17")) <> txt1(5) Then
                    IsCanUpdate = True
                End If
                
                If IsCanUpdate = True Then
                    'edit by nickc 2008/03/28 加欄位
                    'oStrSQL = "update Tmgoods set tg06='" & ChgSQL(txt1(0)) & "',tg07='" & ChgSQL(txt1(1)) & "',tg08='" & ChgSQL(txt1(2)) & "',tg12='" & strUserNum & "',tg13=to_number(to_char(sysdate,'YYYYMMDD')),tg14=to_number(to_char(sysdate,'HH24MI')) where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
                    oStrSQL = "update Tmgoods set tg06='" & ChgSQL(txt1(0)) & "',tg15='" & ChgSQL(txt1(1)) & "',tg07='" & ChgSQL(txt1(2)) & "',tg16='" & ChgSQL(txt1(3)) & "',tg08='" & ChgSQL(txt1(4)) & "',tg17='" & ChgSQL(txt1(5)) & "',tg12='" & strUserNum & "',tg13=to_number(to_char(sysdate,'YYYYMMDD')),tg14=to_number(to_char(sysdate,'HH24MI')) where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
                    'add by nickc 2006/06/13
                    cnnConnection.Execute oStrSQL
                End If
            Else
                'edit by nickc 2008/03/28 加欄位
                'oStrSQL = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg07,tg08,tg09,tg10,tg11) values ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & Trim(.Text) & "','" & ChgSQL(txt1(0)) & "','" & ChgSQL(txt1(1)) & "','" & ChgSQL(txt1(2)) & "','" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                oStrSQL = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg15,tg07,tg16,tg08,tg17,tg09,tg10,tg11) values ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & Trim(.Text) & "','" & ChgSQL(txt1(0)) & "','" & ChgSQL(txt1(1)) & "','" & ChgSQL(txt1(2)) & "','" & ChgSQL(txt1(3)) & "','" & ChgSQL(txt1(4)) & "','" & ChgSQL(txt1(5)) & "','" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                'add by nickc 2006/06/13
                cnnConnection.Execute oStrSQL
            End If
            'edit by nickc 2006/06/13
            'cnnConnection.Execute oStrSQL
        Next i930922
        'add by nickc 2006/06/28 刪除其他不是在畫面上的類別資料 秀玲說的
        'Modify By Sindy 2014/2/19 加條件 and (TG18 is null or TG18='') 的才可以刪除
        oStrSQL = "delete from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05 not in(" & GetAddStr(AllClass) & ") and (TG18 is null or TG18='')"
        cnnConnection.Execute oStrSQL
ShowErr:
    If Err.Number = 0 Then
        cnnConnection.CommitTrans
        MsgBox "存檔成功！", vbInformation
        UpForm.ChkTG = True
    Else
        cnnConnection.RollbackTrans
        MsgBox "存檔失敗！", vbExclamation
        UpForm.ChkTG = False
        grd1.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
        grd1.Visible = True
        Me.Enabled = True
        Exit Sub
    End If
    grd1.MousePointer = flexDefault
    Screen.MousePointer = vbDefault
    grd1.Visible = True
    Me.Enabled = True
    'edit by nick 2004/10/05
    If UpForm.Tag = "frm030001_1" Then
        UpForm.cmdok(3).BackColor = &H8000000F
    End If
    End With
    Me.Hide
    UpForm.Show
    Unload Me
Case 1
    If IsSave = False Then
        If MsgBox("尚未存檔，確定離開？", vbOKCancel, "警告！") = vbOK Then
            Me.Hide
            UpForm.Show
            Unload Me
            Exit Sub
        End If
    Else
        'Modify By Sindy 2009/09/17
        If UpForm.Name <> "frm02010404_3" And _
            UpForm.Name <> "frm02010301_2" And _
            UpForm.Name <> "frm03020404_03" And _
            UpForm.Name <> "frm030203_02" Then
            Me.Hide
            UpForm.Show
        End If
        '2009/09/17 End
        Unload Me
        Exit Sub
    End If
Case 2    ' 列印定稿
        EndLetter "05", m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&000", "26", strUserNum
        NowPrint m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&000", "05", "26", False, strUserNum, 0
'add by nickc 2008/03/07
Case 3
        m_copytext = ""
        For i930922 = 1 To grd1.Rows - 1
            Select Case SSTab1.Tab
            'edit by nickc 2008/03/28 加欄位
            'Case 0
            Case 0, 1
                    'edit by nickc 2008/03/28 加欄位
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 2) & vbCrLf
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 2) & grd1.TextMatrix(i930922, 3) & vbCrLf
                    m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 3) & grd1.TextMatrix(i930922, 4) & vbCrLf
            'edit by nickc 2008/03/28 加欄位
            'Case 1
            Case 2, 3
                    'edit by nickc 2008/03/28 加欄位
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 3) & vbCrLf
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 4) & grd1.TextMatrix(i930922, 5) & vbCrLf
                    m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 5) & grd1.TextMatrix(i930922, 6) & vbCrLf
            'edit by nickc 2008/03/28 加欄位
            'Case 2
            Case 4, 5
                    'edit by nickc 2008/03/28 加欄位
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 4) & vbCrLf
                    'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 6) & grd1.TextMatrix(i930922, 7) & vbCrLf
                    m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 7) & grd1.TextMatrix(i930922, 8) & vbCrLf
            End Select
        Next i930922
        Clipboard.Clear
        Clipboard.SetText m_copytext
        MsgBox "複製完成！", vbInformation, "完成！"
Case 4
        m_copytext = ""
        For i930922 = 1 To grd1.Rows - 1
            If grd1.TextMatrix(i930922, 0) = "☆" Then
                Select Case SSTab1.Tab
                'edit by nickc 2008/03/28 加欄位
                'Case 0
                Case 0, 1
                        'edit by nickc 2008/03/28 加欄位
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 2) & vbCrLf
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 2) & grd1.TextMatrix(i930922, 3) & vbCrLf
                        m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 3) & grd1.TextMatrix(i930922, 4) & vbCrLf
                'edit by nickc 2008/03/28 加欄位
                'Case 1
                Case 2, 3
                        'edit by nickc 2008/03/28 加欄位
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 3) & vbCrLf
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 4) & grd1.TextMatrix(i930922, 5) & vbCrLf
                        m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 5) & grd1.TextMatrix(i930922, 6) & vbCrLf
                'edit by nickc 2008/03/28 加欄位
                'Case 2
                Case 4, 5
                        'edit by nickc 2008/03/28 加欄位
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 4) & vbCrLf
                        'm_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 6) & grd1.TextMatrix(i930922, 7) & vbCrLf
                        m_copytext = m_copytext & grd1.TextMatrix(i930922, 1) & "：" & grd1.TextMatrix(i930922, 7) & grd1.TextMatrix(i930922, 8) & vbCrLf
                End Select
            End If
        Next i930922
        Clipboard.Clear
        Clipboard.SetText m_copytext
        MsgBox "複製完成！", vbInformation, "完成！"
Case Else
End Select
End Sub

Private Sub SetDataListWidth()
'edit by nickc 2008/03/28 加欄位
'grd1.Cols = 11 '5 edit by nickc 2006/06/13 加欄位
grd1.Cols = 15 '14
grd1.row = 0
grd1.col = 0: grd1.Text = ""
grd1.ColWidth(0) = 250
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 1: grd1.Text = "商品類別"
grd1.ColWidth(1) = 1000
grd1.CellAlignment = flexAlignCenterCenter
'Add By Sindy 2014/2/19
grd1.col = 2: grd1.Text = "延展"
grd1.ColWidth(2) = 500
grd1.CellAlignment = flexAlignCenterCenter
'2014/2/19 END
'edit by nickc 2008/03/28 加欄位
'grd1.col = 2: grd1.Text = "中文"
grd1.col = 3: grd1.Text = "中文-1"
grd1.ColWidth(3) = 2000
grd1.CellAlignment = flexAlignCenterCenter
'edit by nickc 2008/03/28 加欄位
'grd1.col = 3: grd1.Text = "英文"
grd1.col = 4: grd1.Text = "中文-2"
grd1.ColWidth(4) = 2000
grd1.CellAlignment = flexAlignCenterCenter
'edit by nickc 2008/03/28 加欄位
'grd1.col = 4: grd1.Text = "日文"
grd1.col = 5: grd1.Text = "英文-1"
grd1.ColWidth(5) = 2000
grd1.CellAlignment = flexAlignCenterCenter
'add by nickc 2008/03/28
grd1.col = 6: grd1.Text = "英文-2"
grd1.ColWidth(6) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 7: grd1.Text = "日文-1"
grd1.ColWidth(7) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 8: grd1.Text = "日文-2"
grd1.ColWidth(8) = 2000
grd1.CellAlignment = flexAlignCenterCenter

'add by nickc 2006/06/13
'edit by nickc 2008/03/28 加欄位
'grd1.col = 5: grd1.Text = "建立人員"
'grd1.ColWidth(5) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 6: grd1.Text = "建立日期"
'grd1.ColWidth(6) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 7: grd1.Text = "建立時間"
'grd1.ColWidth(7) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 8: grd1.Text = "修改人員"
'grd1.ColWidth(8) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 9: grd1.Text = "修改日期"
'grd1.ColWidth(9) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 10: grd1.Text = "修改時間"
'grd1.ColWidth(10) = 1000
'grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 9: grd1.Text = "建立人員"
grd1.ColWidth(9) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 10: grd1.Text = "建立日期"
grd1.ColWidth(10) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 11: grd1.Text = "建立時間"
grd1.ColWidth(11) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 12: grd1.Text = "修改人員"
grd1.ColWidth(12) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 13: grd1.Text = "修改日期"
grd1.ColWidth(13) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 14: grd1.Text = "修改時間"
grd1.ColWidth(14) = 1000
grd1.CellAlignment = flexAlignCenterCenter
End Sub

Public Sub QueryData()
Dim oStrSQL As String
Dim tmpClass As Variant
Dim i930922 As Integer
Dim i930922_1 As Integer
Dim IsFind As Boolean
Dim ii As Integer, strTM09 As String, strGoods As String

Screen.MousePointer = vbHourglass
grd1.MousePointer = flexHourglass
grd1.Visible = False
grd1.Clear
grd1.Rows = 2
SetDataListWidth

m_TM01 = SystemNumber(TGKey, 1)
m_TM02 = SystemNumber(TGKey, 2)
m_TM03 = SystemNumber(TGKey, 3)
m_TM04 = SystemNumber(TGKey, 4)
Me.lbl1.Caption = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04

If AllClass = "" Then
   MsgBox "尚未建類別！", , "錯誤！"
   Me.Hide
   UpForm.ChkTG = False
   UpForm.Show
   Unload Me
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'edit by nickc 2006/06/13 加欄位
'oStrSQL = "select '',tg05,tg06,tg07,tg08 from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' order by tg05 "
'edit by nickc 2006/06/28 加入目前有的類別才出來
'oStrSQL = "select '',tg05,tg06,tg07,tg08,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
'edit by nickc 2008/03/28
'oStrSQL = "select '',tg05,tg06,tg07,tg08,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05 in (" & GetAddStr(AllClass) & ") and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
'Modify By Sindy 2010/6/23 讀取TG全部商品類別
'oStrSQL = "select '',tg05,tg06,tg15,tg07,tg16,tg08,tg17,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05 in (" & GetAddStr(AllClass) & ") and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
'Modify By Sindy 2014/2/19 +,tg18
oStrSQL = "select '',tg05,tg18,tg06,tg15,tg07,tg16,tg08,tg17,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
CheckOC3
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        'edit by nick 2005/01/19 因為長度大於 999 用 set 到 grid 時，會後面被截掉
        'Set grd1.Recordset = AdoRecordSet3
        .MoveFirst
        Do While Not .EOF
            If grd1.Rows = 2 Then
                grd1.row = grd1.Rows - 1
                grd1.col = 1
                If Trim(grd1.Text) <> "" Then
                    grd1.Rows = grd1.Rows + 1
                    grd1.row = grd1.Rows - 1
                End If
            Else
                grd1.Rows = grd1.Rows + 1
                grd1.row = grd1.Rows - 1
            End If
            grd1.col = 1
            grd1.Text = CheckStr(.Fields(1)) '商品類別
            'Add By Sindy 2014/2/19
            grd1.col = 2
            grd1.Text = CheckStr(.Fields(2)) 'N=未延展
            '2014/2/19 END
            grd1.col = 3
            grd1.Text = CheckStr(.Fields(3)) '中文
            grd1.col = 4
            grd1.Text = CheckStr(.Fields(4))
            grd1.col = 5
            grd1.Text = CheckStr(.Fields(5)) '英文
            'add by nickc 2006/06/13
            grd1.col = 6
            grd1.Text = CheckStr(.Fields(6))
            grd1.col = 7
            grd1.Text = CheckStr(.Fields(7)) '日文
            grd1.col = 8
            grd1.Text = CheckStr(.Fields(8))
            grd1.col = 9
            grd1.Text = CheckStr(.Fields(9))
            grd1.col = 10
            grd1.Text = CheckStr(.Fields(10))
            grd1.col = 11
            grd1.Text = CheckStr(.Fields(11))
            'add by nickc 2008/03/28 加欄位
            grd1.col = 12
            grd1.Text = CheckStr(.Fields(12))
            grd1.col = 13
            grd1.Text = CheckStr(.Fields(13))
            grd1.col = 14
            grd1.Text = CheckStr(.Fields(14))
            .MoveNext
        Loop
        ChgData (1)
        SetDataListWidth
        With grd1
        If AllClass <> "" Then
            tmpClass = Split(AllClass, ",")
                For i930922 = 0 To UBound(tmpClass)
                    IsFind = False
                    For i930922_1 = 1 To .Rows - 1
                        .row = i930922_1
                        .col = 1
                        If Trim(.Text) = Trim(tmpClass(i930922)) Then
                            IsFind = True
                            Exit For
                        End If
                    Next i930922_1
                    If Trim(tmpClass(i930922)) <> "" And IsFind = False Then
                        .row = 1
                        .col = 1
                        If Trim(.Text) <> "" Then
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 1
                            .Text = Trim(tmpClass(i930922))
                        Else
                            .Text = Trim(tmpClass(i930922))
                        End If
                    End If
                Next i930922
        End If
        UpForm.ChkTG = True
        For i930922_1 = 1 To .Rows - 1
            .row = i930922_1
            'edit by nickc 2008/03/28 加欄位
            '.col = 2
            'If Trim(.Text) = "" Then
            'If Trim(.TextMatrix(i930922_1, 2)) = "" And Trim(.TextMatrix(i930922_1, 3)) = "" Then
            If Trim(.TextMatrix(i930922_1, 3)) = "" And Trim(.TextMatrix(i930922_1, 4)) = "" Then
                If ChkCht = True Then
                        UpForm.ChkTG = False
                        Exit For
                End If
                'edit by nickc 2008/03/28 加欄位
                '.col = 3
                'If Trim(.Text) = "" Then
                'If Trim(.TextMatrix(i930922_1, 4)) = "" And Trim(.TextMatrix(i930922_1, 5)) = "" Then
                If Trim(.TextMatrix(i930922_1, 5)) = "" And Trim(.TextMatrix(i930922_1, 6)) = "" Then
                    If ChkEng = True Then
                        UpForm.ChkTG = False
                        Exit For
                    End If
                    'edit by nickc 2008/03/28 加欄位
                    '.col = 4
                    'If Trim(.Text) = "" Then
                    'If Trim(.TextMatrix(i930922_1, 6)) = "" And Trim(.TextMatrix(i930922_1, 7)) = "" Then
                    If Trim(.TextMatrix(i930922_1, 7)) = "" And Trim(.TextMatrix(i930922_1, 8)) = "" Then
                        UpForm.ChkTG = False
                        Exit For
                    End If
                End If
            End If
        Next i930922_1
        
        'Add By Sindy 2024/12/20 檢查與公報商品資料是否一致
        m_CompareGoods = ""
        If PubMsg = "比對公報商品資料" Then
            For ii = 1 To .Rows - 1
               strTM09 = Trim(.TextMatrix(ii, 1)) '商品類別
               strGoods = Trim(.TextMatrix(ii, 3)) & Trim(.TextMatrix(ii, 4)) '商品內容
               strExc(0) = "SELECT * from TMBULLETINGOODS" & _
                           " where TBG11='1' and TBG01='" & m_TM15 & "' and TBG02='" & strTM09 & "' and TBG07='" & m_TM08 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If strGoods <> Trim("" & RsTemp.Fields("TBG04")) & Trim(RsTemp.Fields("TBG05")) & Trim(RsTemp.Fields("TBG06")) & Trim(RsTemp.Fields("TBG08")) & Trim(RsTemp.Fields("TBG09")) & Trim(RsTemp.Fields("TBG10")) Then
                     m_CompareGoods = m_CompareGoods & "、" & strTM09
                  End If
               Else
                  m_CompareGoods = m_CompareGoods & "、" & strTM09 & "(無資料)"
               End If
            Next ii
            If m_CompareGoods = "" Then
               m_CompareGoods = "與公報商品資料比對一致！"
            Else
               m_CompareGoods = "注意！" & vbCrLf & "商品類別: " & Mid(m_CompareGoods, 2) & " 與公報商品資料比對不一致！"
            End If
        End If
        '2024/12/20 END
        
        End With
    Else
        UpForm.ChkTG = False
        If AllClass = "" Then MsgBox "尚未建類別！", , "錯誤！":  Me.Hide: UpForm.Show: Unload Me: Exit Sub
        tmpClass = Split(AllClass, ",")
        With grd1
            For i930922 = 0 To UBound(tmpClass)
                If Trim(tmpClass(i930922)) <> "" Then
                    .row = 1
                    .col = 1
                    If Trim(.Text) <> "" Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 1
                        .Text = tmpClass(i930922)
                    Else
                        .Text = tmpClass(i930922)
                    End If
                End If
            Next i930922
        End With
    End If
End With
ChgData 1
OpenTable 'Added by Morgan 2023/2/15
grd1.Visible = True
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   If cmdok(0).Visible = False Then
      cmdRemAtt.Enabled = False
      cmdAddAtt.Enabled = False
      
      'Add By Sindy 2024/12/20 檢查與公報商品資料是否一致
      If PubMsg = "比對公報商品資料" And m_CompareGoods <> "" Then
         MsgBox m_CompareGoods
      End If
      '2024/12/20 END
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   IsSave = True
   ChkCht = False
   ChkEng = False
   ChkJpn = False
   PubMsg = ""
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/22
   
   'Added by Morgan 2023/2/15
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   KillTemp
   'end 2023/2/15
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm03010303_04 = Nothing
End Sub

Sub ChgData(oIndex As Integer)
Dim i930922 As Integer
If oIndex = 0 Then Exit Sub
With grd1
    .Visible = False
    NowRow = oIndex
    For i930922 = 0 To .Rows - 1
        .row = i930922
        .col = 0
        .Text = ""
    Next i930922
    .row = oIndex
    .col = 0
    .Text = "☆"
    .col = 3 '2
    txt1(0).Text = .Text
    .col = 4 '3
    txt1(1).Text = .Text
    .col = 5 '4
    txt1(2).Text = .Text
    'add by nickc 2008/03/28 加欄位
    .col = 6 '5
    txt1(3).Text = .Text
    .col = 7 '6
    txt1(4).Text = .Text
    .col = 8 '7
    txt1(5).Text = .Text
    .Visible = True
End With
End Sub

Private Sub Grd1_Click()
ChgData (grd1.MouseRow)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
InverseTextBox txt1(Index)
'edit by nickc 2007/06/06
'If Index = 0 Then txt1(Index).IMEMode = 1
'edit by nickc 2007/09/29 避免因為輸入法切換空白，而自動回前畫面
'If Index = 0 Then OpenIme Else CloseIme
If Index = 0 And txt1(Index).Enabled = True Then OpenIme Else CloseIme
End Sub

'Added by Lydia 2021/08/17 Form 2.0的TextBox增加右鍵選單功能
Private Sub txt1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txt1(Index)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
 
   Cancel = False
   If IsEmptyText(txt1(Index)) = False Then
      'edit by nickc 2007/03/27
      'If CheckLengthIsOK(txt1(Index), 2000) = False Then
      If CheckLengthIsOK(txt1(Index), 4000) = False Then
         Cancel = True
         txt1_GotFocus Index
      End If
    End If
End Sub

Private Sub txt2_GotFocus(Index As Integer)
   InverseTextBox txt2(Index)
   CloseIme
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
    KeyAscii = UpperCase(KeyAscii)
Case Else
    'Add By Sindy 2010/3/9 開放可以打T
    If Index = 2 Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case KeyAscii
      Case 48 To 57, 8, 65 To 90
      Case Else
              KeyAscii = 0
      End Select
    Else
    '2010/3/9 End
      Select Case KeyAscii
      'edit by nickc 2008/05/30 加入可以按倒退鍵
      'Case 48 To 57
      Case 48 To 57, 8
      Case Else
              KeyAscii = 0
      End Select
    End If
End Select
End Sub

'Added by Morgan 2023/2/15

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub OpenTable()
   strExc(0) = "select '' as V, TGA05||' ('||Round(TGA06 / 1024, 2)||' KB)' as 檔案名稱 " & _
      ", sqldatet(TGA11)||' '||sqltime(TGA12)||'('||st02||')' as 上傳時間,TGA05,TGA09" & _
      " from TMGoodsAtt,staff where TGA01='" & m_TM01 & "' and TGA02='" & m_TM02 & "'" & _
      " and TGA03='" & m_TM03 & "' and TGA04='" & m_TM04 & "' and st01(+)=TGA10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set MSHFlexGrid1.Recordset = RsTemp
   SetGrid
   If intI = 1 Then
      cmdAddAtt.Enabled = False
      cmdOpenAtt.Enabled = True
      cmdSaveAtt.Enabled = True
      cmdRemAtt.Enabled = True
   Else
      cmdAddAtt.Enabled = True
      cmdRemAtt.Enabled = False
      cmdOpenAtt.Enabled = False
      cmdSaveAtt.Enabled = False
   End If
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer
   
   arrGridHeadWidth = Array(240, 5000, 2300)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
      .FormatString = "V|檔案名稱|最後修改時間|上傳時間"
   End If
   '.FixedCols = 2
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

Private Sub MSHFlexGrid1_Click()
   Dim intRow As Integer
   If MSHFlexGrid1.row > 0 Then
      intRow = MSHFlexGrid1.row
      GridClick MSHFlexGrid1, intRow, 0
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
   With MSHFlexGrid1
   If .row > 0 Then
      If cmdOpenAtt.Enabled Then OpenAtt .TextMatrix(.row, 3), .TextMatrix(.row, 4)
   End If
   End With
End Sub


Private Sub cmdAddAtt_Click()
   Dim stFileName As String, stInitDir As String
   Dim sFile() As String
   Dim ii As Integer
   Dim fs, s
   Dim f
   Dim bolAdd As Boolean
   
On Error GoTo ErrHnd
   
   With CommonDialog1
   .CancelError = True
   .FileName = ""
   .Filter = "WORD 檔案(*.doc,*.docx)|*.doc;*.docx"
   
   stInitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   If stInitDir = "" Or Dir(stInitDir, vbDirectory) = "" Then
      .InitDir = PUB_Getdesktop
   Else
      .InitDir = stInitDir
   End If
   .MaxFileSize = 3000
   '.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
   .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
   .ShowOpen
   If .FileName <> "" Then
      If InStr(.FileName, ChrW$(0)) > 0 Then
         sFile = Split(.FileName, ChrW$(0))
      Else
         ReDim sFile(1) As String
         sFile(0) = Left(.FileName, InStrRev(.FileName, "\") - 1)
         sFile(1) = Mid(.FileName, InStrRev(.FileName, "\") + 1)
      End If
      
      '記錄路徑
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
      For ii = 1 To UBound(sFile)
         If InStr(sFile(ii), "\") > 0 Then
            stFileName = sFile(ii)
         Else
            stFileName = sFile(0) & "\" & sFile(ii)
         End If
         
         If Right(Trim(UCase(stFileName)), 4) <> ".DOC" And Right(Trim(UCase(stFileName)), 5) <> ".DOCX" Then
            MsgBox "格式不符,只可存放 Word 檔!!"
            GoTo EXITSUB
         End If
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set f = fs.GetFile(stFileName)
         If f.Size = 0 Then
            ShowMsg sFile(ii) & MsgText(9221)
         ElseIf f.Size > 5242880 Then
            If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
               GoTo EXITSUB
            End If
         End If
         If IsRecordExist(sFile(ii)) = False Then
            '存檔並刪除原檔
            If AddRecord(sFile(ii), stFileName, f) = True Then
               bolAdd = True
            Else
               GoTo EXITSUB
            End If
         End If
      Next ii
   End If
   End With
   
   
EXITSUB:
   If bolAdd Then
      OpenTable
   End If
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
               
End Sub

Private Function AddRecord(pFileName As String, pFullFromPath As String, pFile As Variant) As Boolean
   Dim stSQL As String, iRecords As Integer
   Dim bolInTrans As Boolean
   Dim stFtpPath As String

On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   stSQL = "update TMGoodsAtt set TGA05=TGA05 where TGA01='" & m_TM01 & "' and TGA02='" & m_TM02 & "'" & _
      " and TGA03='" & m_TM03 & "' and TGA04='" & m_TM04 & "' and upper(TGA05)='" & ChgSQL(UCase(pFileName)) & "'"
   cnnConnection.Execute stSQL, iRecords
   If iRecords > 0 Then
      Err.Raise 999, , "檔名 " & pFileName & " 重複!!"
   End If
   
   If PUB_PutFtpFile(pFullFromPath, m_TM01 & m_TM02 & m_TM03 & m_TM04, pFileName, stFtpPath, "TMGoodsAtt") Then
      stSQL = "insert into TMGoodsAtt(TGA01,TGA02,TGA03,TGA04,TGA05,TGA06,TGA07,TGA08,TGA09,TGA10,TGA11,TGA12)" & _
         " values('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & ChgSQL(pFileName) & "'" & _
         "," & pFile.Size & "," & Format(pFile.DateLastModified, "YYYYMMDD") & "," & Format(pFile.DateLastModified, "HHMMSS") & _
         ",'" & ChgSQL(stFtpPath) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
      cnnConnection.Execute stSQL, iRecords
   Else
      Err.Raise 999, , " 檔案 " & pFileName & " 上傳失敗!!"
   End If

   cnnConnection.CommitTrans
   'pFile.Delete 'Removed by Morgan 2023/2/16 不必刪除--79020
   AddRecord = True
   
ErrHand:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal stFileName As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim adoRst As ADODB.Recordset
   
   IsRecordExist = False
   
   stSQL = "SELECT * FROM TMGoodsAtt WHERE TGA01='" & m_TM01 & "' and TGA02='" & m_TM02 & "'" & _
      " and TGA03='" & m_TM03 & "' and TGA04='" & m_TM04 & "' and upper(TGA05)=upper('" & ChgSQL(stFileName) & "')"
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      IsRecordExist = True
      MsgBox "檔案 " & stFileName & " 已存在！", vbCritical
   End If
   
   Set adoRst = Nothing
End Function

'Removed by Morgan 2023/2/16
'Private Sub cmdExit_Click()
'   Unload Me
'End Sub
'end 2023/2/16

Private Sub cmdOpenAtt_Click()
   Dim ii As Integer
   Dim bolCheck As Boolean
   
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If UCase(UCase(.TextMatrix(ii, 0))) = "V" Then
         bolCheck = True
         OpenAtt .TextMatrix(ii, 3), .TextMatrix(ii, 4)
      End If
   Next
   End With
   If Not bolCheck Then MsgBox "請點選要開啟的檔案", vbInformation
End Sub

Private Sub OpenAtt(pFileName As String, pFtpPath As String)
   Dim stSaveFileName As String
   Dim hLocalFile As Long
   
   stSaveFileName = m_AttachPath & "\" & pFileName
   If PUB_GetFtpFile(pFtpPath, stSaveFileName, "TMGoodsAtt") Then
      ShellExecute hLocalFile, "open", stSaveFileName, vbNullString, vbNullString, 1
   End If
End Sub

Private Sub cmdRemAtt_Click()
   Dim iRecord As Integer
   Dim ii As Integer, bolCheck As Boolean
   Dim stTableDir As String
   
On Error GoTo ErrHnd

   With MSHFlexGrid1
   bolCheck = False
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" Then
         bolCheck = True
         Exit For
      End If
   Next
   If Not bolCheck Then
      MsgBox "請先勾選要刪除的記錄！", vbExclamation: Exit Sub
   Else
      If MsgBox("是否確定要刪除？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then Exit Sub
   End If
   
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" Then
          '此處不包transaction 若FTP刪除成功但DB刪除失敗才會有log
          strSql = "delete TMGoodsAtt where TGA01='" & m_TM01 & "' and TGA02='" & m_TM02 & "'" & _
            " and TGA03='" & m_TM03 & "' and TGA04='" & m_TM04 & "' and TGA05='" & ChgSQL(.TextMatrix(ii, 3)) & "'"
          Pub_SeekTbLog strSql
          
         stTableDir = PUB_GetFtpTableDir("TMGoodsAtt")
         If PUB_FtpDelFile2(stTableDir & "/" & .TextMatrix(ii, 4)) Then
            cnnConnection.Execute strSql, iRecord
         Else
            Err.Raise 999, , " 刪除Ftp檔案失敗!!"
         End If
         .TextMatrix(ii, 0) = "X"
         .RowHeight(ii) = 0
      End If
   Next
   End With
   OpenTable
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub cmdSaveAtt_Click()
   Dim stFileName As String, stFolderPath As String, stFullName As String, stFtpPath As String
   Dim bMultiFile As Boolean
   Dim ii As Integer
   
   Screen.MousePointer = vbHourglass
   
   stFileName = ""
   bMultiFile = False
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" And Trim(MSHFlexGrid1.TextMatrix(ii, 3)) <> "" Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = Trim(.TextMatrix(ii, 3))
            stFtpPath = Trim(.TextMatrix(ii, 4))
         End If
      End If
   Next ii
   End With
   
   If stFileName = "" Then
      MsgBox "請勾選要下載的檔案！"
   Else
      '多選
      If bMultiFile Then
         If m_SaveFolder = "" Then m_SaveFolder = PUB_Getdesktop
         stFolderPath = PUB_GetFolder(Me.hWnd, m_SaveFolder, "請選擇欲儲存的位置:")
         If stFolderPath <> "" Then
            With MSHFlexGrid1
            For ii = 1 To .Rows - 1
               If UCase(.TextMatrix(ii, 0)) = "V" And Trim(MSHFlexGrid1.TextMatrix(ii, 1)) <> "" Then
                  stFileName = Trim(.TextMatrix(ii, 3))
                  stFullName = stFolderPath & "\" & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & Trim(.TextMatrix(ii, 1)) & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If PUB_GetFtpFile(.TextMatrix(ii, 4), stFullName, "TMGoodsAtt") = False Then
                           MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                           GoTo RunExit
                        End If
                     End If
                  End If
               End If
            Next ii
            End With
         End If
      Else
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If PUB_GetFtpFile(MSHFlexGrid1.TextMatrix(1, 4), stFullName, "TMGoodsAtt") = False Then
                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                  GoTo RunExit
               End If
            End If
         End If
      End If
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
   Screen.MousePointer = vbDefault
End Sub

Private Function GetSaveName(ByVal pFileName As String) As String

On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With

   Exit Function

ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

