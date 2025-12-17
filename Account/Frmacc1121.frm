VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1121 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據抬頭輸入"
   ClientHeight    =   3300
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   7630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7630
   Begin VB.CheckBox Check1 
      Caption         =   "收據暫不列印"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6210
      TabIndex        =   12
      Top             =   2820
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton Command4 
      Caption         =   "檢視接洽單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3090
      TabIndex        =   31
      Top             =   1620
      Width           =   1410
   End
   Begin VB.TextBox txtPrintNo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3950
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1980
      Width           =   345
   End
   Begin VB.CommandButton Command3 
      Caption         =   "開立發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5550
      TabIndex        =   15
      Top             =   2010
      Width           =   1395
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2790
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   -210
      TabIndex        =   26
      Top             =   2400
      Width           =   7800
      Begin VB.CheckBox Check2 
         Caption         =   "3.代理人請款之匯款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   9
         Top             =   30
         Width           =   2650
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2.代理人請款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3540
         TabIndex        =   8
         Top             =   30
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "1.送件日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2430
         TabIndex        =   7
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收據自動列印時間點"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   27
         Top             =   60
         Width           =   1890
      End
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1980
      Width           =   1125
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   25
      Top             =   1590
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6360
      TabIndex        =   14
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5310
      TabIndex        =   13
      Top             =   1620
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5220
      MaxLength       =   1
      TabIndex        =   3
      Top             =   870
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   1
      Top             =   510
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1980
      TabIndex        =   17
      Top             =   150
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   150
      Width           =   855
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   3990
      TabIndex        =   11
      Top             =   2790
      Width           =   1785
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3149;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1140
      TabIndex        =   2
      Top             =   870
      Width           =   4005
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7064;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Top             =   1230
      Width           =   5175
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "9128;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   315
      Left            =   2700
      TabIndex        =   19
      Top             =   510
      Width           =   3615
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "6376;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印統編       (Y:印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3210
      TabIndex        =   30
      Top             =   2010
      Width           =   1900
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "介紹案源同仁"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   2508
      TabIndex        =   29
      Top             =   2856
      Width           =   1344
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   2850
      Width           =   840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "預定收款日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   2010
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   1620
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   120
      Top             =   1500
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   22
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.不可扣繳 2.可扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5700
      TabIndex        =   21
      Top             =   930
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0j0 As New ADODB.Recordset
'Public adoacc0k0 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim strA0K01, stra0k03, strA0K04, stra0k05, stra0k08, strA0K11, stra0k20 As String
Dim lnga0k02, lnga0k06, lnga0k07 As Long
Dim m_bolSplitMail As Boolean '是否拆收據發Mail
Dim m_strMailDesc As String, m_strMailSubject As String
Dim stra0k32 As String 'Add By Sindy 2010/4/19
Dim m_strChkCompany As String, m_strCaseNo As String '檢查是否為專利商標公司 Added by Morgan 2012/9/12
Dim m_A0j04 As String 'Add By Sindy 2012/11/13
Dim m_CP09 As String 'Add By Sindy 2012/12/6
Dim m_CP31 As String, m_CP12 As String 'Add By Sindy 2013/12/17
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String 'Add By Sindy 2010/11/18
Dim m_CP10List As String 'Added by Lydia 2020/09/28 所有收文號之案件性質
Public m_CRL119 As String, m_CRL49 As String, m_CP140 As String, m_CallForm As String 'Add By Sindy 2014/2/11
Public m_CRL02 As String 'Add By Sindy 2020/3/31
Public strB_CP09 As String 'Add by Amy 2016/08/18 前畫面選取第一筆的總收文號
Dim tmpfrm As Form 'Add By Sindy 2023/1/4
Dim m_CRL153 As String 'Added by Lydia 2023/11/13 國內接洽單：DEBIT NOTE請款選項
Dim strShowCRL153 As String 'Added by Lydia 2024/08/05

'Add By Sindy 2013/12/25
Private Sub Check2_GotFocus(Index As Integer)
   Select Case Index
      Case 0
         Check2(1).Value = 0
         Check2(2).Value = 0
      Case 1
         Check2(0).Value = 0
         Check2(2).Value = 0
      Case 2
         Check2(0).Value = 0
         Check2(1).Value = 0
   End Select
End Sub

Private Sub Combo1_GotFocus()
    StatusView MsgText(65) & "100"
End Sub

Private Sub Combo1_LostFocus()
Dim m_CU173 As String
   
   'Modify By Sindy 2017/3/24
   If Combo1.Tag <> Combo1.Text Then
      m_CU173 = ""
      'Modify By Sindy 2019/5/22 + Text3
      Call GetTitleCustData(Combo1.Text, Text3, "", , , , , , , , , , , , , , , , , , , , , , , , , , , , m_CU173)
      txtPrintNo.Text = m_CU173
      Combo1.Tag = Combo1.Text
   End If
   '2017/3/24 END
   StatusView MsgText(601)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If CheckLen(Label3, Me.Combo1.Text, 100) = MsgText(603) Then
        Cancel = True
        Exit Sub
    End If
   'Modify by Amy 2014/09/24 若為境外公司 只能為1.個人且不可改
   If PUB_GetTaxNo(Combo1, 1) = "Y" Then
        Text6 = "1"
        Text6.Locked = True
   Else
        Text6.Locked = False
   End If
   'end 2014/09/24
End Sub

Private Sub Command1_Click()
Dim bCancel As Boolean
Dim strSpecCompany As String
   
   m_bolSplitMail = False
'   strDelConfirm = MsgBox(MsgText(93), vbOKCancel + vbDefaultButton1, MsgText(5))
'   If strDelConfirm = vbCancel Then
'      Exit Sub
'   End If
   If Text8 <> MsgText(601) Then
      Exit Sub
   End If
   If Text6 = MsgText(601) Then
      MsgBox MsgText(52), , MsgText(5)
      Text6.SetFocus
   End If
   If ExistCheck("acc080", "a0801", Text1, "", False) = False Then
      MsgBox MsgText(45) & Label1, , MsgText(5)
      If Text1.Enabled = True Then Text1.SetFocus 'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
      Exit Sub
   End If
   
   'Modify By Sindy 2020/3/24
   'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
   If strSrvDate(1) >= 事務所合併日 Then
      If Text1 <> "2" And Text1 <> "9" And Text1 <> "J" And Text1 <> "L" Then
         MsgBox "收據公司別只可輸入２或９或Ｊ或Ｌ", , MsgText(5)
         If Text1.Enabled = True Then Text1.SetFocus
         Exit Sub
      End If
   ElseIf strSrvDate(1) >= 智慧所更名日 Then
      If Text1 <> "1" And Text1 <> "2" And Text1 <> "9" And Text1 <> "J" And Text1 <> "L" Then
         MsgBox "收據公司別只可輸入１或２或９或Ｊ或Ｌ", , MsgText(5)
         If Text1.Enabled = True Then Text1.SetFocus
         Exit Sub
      End If
   'end 2020/04/13
   Else
   '2020/3/24 END
      'add by sonia 2013/6/5 add by sonia 瑞婷說只能開1,2,9公司
      'Add By Sindy 2013/12/17
      If strSrvDate(1) >= InvoiceStartDate Then
         If Text1 <> "1" And Text1 <> "2" And Text1 <> "9" And Text1 <> "J" Then
            MsgBox "收據公司別只可輸入１或２或９或J", , MsgText(5)
            Text1.SetFocus
            Exit Sub
         End If
      Else
      '2013/12/17 END
         If Text1 <> "1" And Text1 <> "2" And Text1 <> "9" Then
            MsgBox "收據公司別只可輸入１或２或９", , MsgText(5)
            Text1.SetFocus
            Exit Sub
         End If
      End If
      '2013/6/5 end
   End If
   
   'add by sonia 2020/4/14
   If InStr(m_CP01, "L") = 0 And Text1 = "L" Then
      MsgBox "非法務顧問案件，收據公司別不可開Ｌ公司", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   'end 2020/4/14
   
   'Add by Morgan 2008/5/5
   Text5_Validate bCancel
   If bCancel = True Then
      Text5.SetFocus
      Text5_GotFocus
      Exit Sub
   End If
   
   'Added by Morgan 2013/4/29
   If txtDate = "" Then
      MsgBox "請輸入收據日期!"
      txtDate.SetFocus
      Exit Sub
   End If
   'end 2013/4/29
   
   'add by sonia 2014/3/18
   If Frmacc1120.m_cp10N = True And Text7 = "" Then
      If MsgBox("點選案件性質有含 代辦退費, 但未輸入備註欄, 是否要輸入？", vbYesNo + vbQuestion, "代辦退費檢查") = vbYes Then
         Exit Sub
      End If
   End If
   '2014/3/18 end
   
   'Add By Sindy 2013/12/31
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select cp01, cp14, a0j04, cp10, cp02,cp03,cp04,cp140,cp151,cp09,cp31,a0j11,cp12 from caseprogress, acc0j0 where cp09 = a0j01 and a0j06 = '" & MsgText(602) & "' and a0j13=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      With adocheck
         .MoveFirst
         m_strChkCompany = "": m_strCaseNo = ""
         Do While Not .EOF
            If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
               strSpecCompany = ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
               If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                  m_strChkCompany = strSpecCompany
                  If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                  m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   adocheck.Close
   '2013/12/31 END
   
   'add by sonia 2020/6/11 檢查不可開智權公司的條件1.台灣專利商標2.ACS特定案件屬性
   If Text1 = "J" Then
      '1.台灣專利商標
      If m_A0j04 = "000" And (m_CP01 = "P" Or m_CP01 = "T" Or m_CP01 = "FCP" Or m_CP01 = "FCT") Then
         MsgBox "台灣專利商標案件，收據公司別不可開Ｊ公司！", , MsgText(5)
         Text1.SetFocus
         Exit Sub
      End If
      '2.ACS特定案件屬性
      If m_CP01 = "ACS" Then
         'Modified by Lydia 2020/10/14 (10/13) 判斷該案收文進度若有101專利布局分析，則不可開智權公司。
'         'Added by Lydia 2020/09/28 新案則改判斷案件性質，若有101專利布局分析、105企業IP經營規劃、106品牌台灣發展計畫、205驗證申請者則公司別不可輸J公司。
'         If m_CP31 = "Y" And (InStr(m_CP10List, "101,") > 0 Or InStr(m_CP10List, "105,") > 0 Or InStr(m_CP10List, "106,") > 0 Or InStr(m_CP10List, "205,") > 0) Then
'            MsgBox "ACS特定案件屬性案件，收據公司別不可開Ｊ公司！", , MsgText(5)
'            Text1.SetFocus
'            Exit Sub
'         Else
'         'end 2020/09/28
'            strSql = "SELECT LC47 FROM LAWCASE WHERE LC01='" & m_CP01 & "' AND LC02='" & m_CP02 & "' AND LC03='" & m_CP03 & "' AND LC04='" & m_CP04 & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               If ("" & RsTemp.Fields("LC47") = "專利檢索布局分析" Or "" & RsTemp.Fields("LC47") = "企業IP經營管理(TIPS)" Or "" & RsTemp.Fields("LC47") = "品牌智財運用輔導") Then
'                  MsgBox "ACS特定案件屬性案件，收據公司別不可開Ｊ公司！", , MsgText(5)
'                  Text1.SetFocus
'                  Exit Sub
'               End If
'            End If
'         End If 'Added by Lydia 2020/09/28
         strSql = "select cp09,cp10 from caseprogress where cp01='" & m_CP01 & "' AND cp02='" & m_CP02 & "' AND cp03='" & m_CP03 & "' AND cp04='" & m_CP04 & "' "
         'Modified by Lydia 2020/11/24 101專利布局分析已改為113
         strSql = strSql & "and cp10='113' and cp159=0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Modified by Lydia 2020/11/24 101專利布局分析已改為113
            MsgBox "ACS案有收文113專利布局分析，收據公司別不可開Ｊ公司！", , MsgText(5)
            Text1.SetFocus
            Exit Sub
         End If
         'end 2020/10/14
      End If
   End If
   'end 2020/6/11
   
   'Added by Morgan 2012/9/12
   'Add By Sindy 2013/12/17
   If strSrvDate(1) >= InvoiceStartDate Then
      If m_strChkCompany = "T" And Text1 <> "1" And m_CP31 = "Y" Then
         MsgBox "專利案" & m_strCaseNo & "有設定以專利商標出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         If Text1.Enabled = True Then Text1.SetFocus 'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
         Exit Sub
      ElseIf m_strChkCompany = "J" And Text1 <> "J" And m_CP31 = "Y" Then
         MsgBox m_strCaseNo & "有設定以智權公司出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         If Text1.Enabled = True Then Text1.SetFocus 'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
         Exit Sub
      End If
   Else
   '2013/12/17 END
      If m_strChkCompany <> "" And Text1 <> "1" And m_CP31 = "Y" Then
         MsgBox "專利案" & m_strCaseNo & "有設定以專利商標出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         Text1.SetFocus
         Exit Sub
      End If
   End If
   
   'Add By Sindy 2020/3/24 非法務且非顧問案按確定時，檢查該案號若曾開立收據，
   '                       新收據的公司與案件特殊出名公司不同時，不可存檔並顯示以下訊息
   If strSrvDate(1) >= 智慧所更名日 And Text1 <> "L" Then
      '查該案號曾開立收據的公司別
      'Modify By Sindy 2020/4/1 A0K11 => decode(a0k11,'1','2',a0k11) A0K11
      strSql = "SELECT a0k02,decode(a0k11,'1','2',a0k11) A0K11 FROM acc0j0,acc0k0" & _
               " where a0j02 in(" & _
               " SELECT a0j02 FROM caseprogress,acc0j0" & _
               " WHERE cp09 = a0j01 AND a0j06 = 'Y' AND a0j13=a0j01)" & _
               " AND a0j06 IS NULL AND a0j13=a0k01" & _
               " order by a0k02 desc,a0k11"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      strA0K11 = ""
      If intI = 1 Then
         strA0K11 = "" & RsTemp.Fields("a0k11")
      End If
      '有開過收據才要判斷新收據的公司與案件特殊出名公司不同時，不可存檔
      If strA0K11 <> "" Then
         If Text1 = "J" And m_strChkCompany <> "J" Then
            MsgBox "1.請先印出結餘單結算該案結餘程序" & vbCrLf & vbCrLf & _
                   "2.通知電腦中心更改公司別設定" & vbCrLf & vbCrLf & _
                   "3.開立J公司請款單", vbCritical, "收據公司別提醒"
            If Text1.Enabled = True Then Text1.SetFocus 'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
            Exit Sub
         ElseIf Text1 <> "J" And m_strChkCompany = "J" Then
            MsgBox "1.請先印出結餘單結算該案結餘程序" & vbCrLf & vbCrLf & _
                   "2.通知電腦中心更改公司別設定" & vbCrLf & vbCrLf & _
                   "3.開立2公司收據", vbCritical, "收據公司別提醒"
            If Text1.Enabled = True Then Text1.SetFocus 'Modify by Amy 2020/04/13 +if 鎖住又SetFocus 會錯
            Exit Sub
         End If
      End If
   End If
   '2020/3/24 END
   
   'Add By Sindy 2012/11/12
   If m_A0j04 = "000" And (Me.Check2(1).Value = 1 Or Me.Check2(2).Value = 1) Then
      MsgBox "非台灣案時, 收據自動列印時間點才可選擇2 或 3 !!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
'cancel by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選
'   If Me.Check1.Value = 1 And _
'      (Me.Check2(0).Value = 0 And Me.Check2(1).Value = 0 And Me.Check2(2).Value = 0) Then
'      MsgBox "勾選收據暫不列印時, 收據自動列印時間點不可空白!!!", vbExclamation + vbOKOnly
'      Exit Sub
'   End If
'   '2012/11/12 End
'   'add by sonia 2013/12/16
'   If Me.Check1.Value = 0 And _
'      (Me.Check2(0).Value = 1 Or Me.Check2(1).Value = 1 Or Me.Check2(2).Value = 1) Then
'      MsgBox "點選收據自動列印時間點, 收據暫不列印一定要勾選!!!", vbExclamation + vbOKOnly
'      Exit Sub
'   End If
'   '2013/12/16 end
'end 2015/11/26
   
   'Add By Sindy 2012/12/6
   '檢查是否可上收據自動列印時間點
   If PUB_ChkAccIsUpdCP151(m_CP09, IIf(Me.Check2(0).Value = 1, "1", IIf(Me.Check2(1).Value = 1, "2", IIf(Me.Check2(2).Value = 1, "3", "")))) = False Then
      Me.Check2(0).Value = 0
      Me.Check2(1).Value = 0
      Me.Check2(2).Value = 0
      Exit Sub
   End If
   '2012/12/6 End
   
'   'Add By Sindy 2016/12/30
'   strSql = "select cu11" & _
'            " From customer" & _
'            " where (upper(cu04)=upper('" & ChgSQL(Me.Combo1.Text) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(Me.Combo1.Text) & "') or upper(cu06)=upper('" & ChgSQL(Me.Combo1.Text) & "'))" & _
'            " and cu15<>'0'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 0 Then
'      '若A4202='04150022'者視為空值
'      'Modify By Sindy 2017/4/18 and A4202<>'04150022'==>and (A4202<>'04150022' or A4202 is null) 改語法不然抓不到資料
'      strSql = "select a4202" & _
'               " From acc420" & _
'               " where upper(a4201)=upper('" & ChgSQL(Me.Combo1.Text) & "') and (A4202<>'04150022' or A4202 is null)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 0 Then
'         strSql = "select cu11" & _
'                  " From customer" & _
'                  " where (upper(cu04)=upper('" & ChgSQL(Me.Combo1.Text) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(Me.Combo1.Text) & "') or upper(cu06)=upper('" & ChgSQL(Me.Combo1.Text) & "'))"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 0 Then
'            MsgBox "此為新的收據抬頭，請聯絡智權同仁提供基本資料以利建檔!!", vbInformation
'         End If
'      End If
'   End If
'   '2016/12/30 END
   'Add By Sindy 2017/6/19 改呼叫函數 : 檢查收據抬頭是否存在
   'Modified by Sindy 2018/9/18 拿掉chgsql
   If PUB_ChkTitleNmExist(Me.Combo1.Text) = "" Then 'Add By Sindy 2023/7/21 + if
      'MsgBox "收據抬頭不可空白！", vbExclamation + vbOKOnly
      Exit Sub
      '2023/7/21 END
   End If
   '2017/6/19 END
   
   If adoacc0j0.State = adStateOpen Then
      adoacc0j0.Close
   End If
   adoacc0j0.CursorLocation = adUseClient
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoacc0j0.Open "select distinct a0j05 from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/26 取消 a0j05 改抓 cp13
   'adoacc0j0.Open "select distinct a0j05 from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0j0.Open "select distinct cp13 from acc0j0,caseprogress where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01 and cp09(+)=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0j0.RecordCount > 1 Then
      MsgBox MsgText(92), , MsgText(5)
      adoacc0j0.Close
      Exit Sub
   End If
   adoacc0j0.Close
   
   'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   If strItemNo = "" Then
      strItemNo = AutoNo(MsgText(802), 5)
      Text8 = strItemNo
      ProduceData
   Else
      Text8 = strItemNo
      ProduceData
   End If
   Call PUB_ChkJCompanyRecv_Mail(Text8, Text1) 'Add By Sindy 2014/1/29 若收據開J公司,但案件的特殊出名公司未輸入時,同時發E-MAIL
   'Add by Morgan 2008/1/30
   If m_bolSplitMail = True Then
      PUB_SendMail strUserNum, "83002", "", m_strMailSubject, m_strMailDesc
   End If
   
   'Add by Sindy 2023/1/4 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q", tmpfrm) = True Then
      Unload tmpfrm 'frm090801_Q
   End If
   '2023/1/4 END
End Sub

Private Sub Command2_Click()
   KeyEnter vbKeyEscape
End Sub

'Add By Sindy 2013/12/24
'開立發票
Private Sub Command3_Click()
   strItemNo = Text8
   strCustNo = Text3
   strTitle = Me.Name
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   'Frmacc1127.Text1.Text = strItemNo
   Frmacc1127.Text1.Enabled = False
   Frmacc1127.Show
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

'Add By Sindy 2022/11/24 檢視檢洽單
Private Sub Command4_Click()
   Call PUB_Queryfrm090801(m_CP140, "", Me)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim adoRst As ADODB.Recordset
Dim strCU125 As String 'Add By Sindy 2016/12/9
  
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Removed by Morgan 2013/4/29 改單線固定不用再指定
   'Me.Width = 7320 '6700
   'Me.Height = 3315 '2850
   'end 2013/4/29
   'Modify by Morgan 2006/4/24 下移一點(約三筆收文)以便看到前畫面資料--辜
   'Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2 + 900
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   If strCustNo <> MsgText(601) Then
      Text3 = strCustNo
   Else
      Text3 = MsgText(601)
   End If
   Text4 = CustomerQuery(Text3, 1)
    'Modify By Cheng 2003/10/07
'   Text5 = CustomerQuery(Text3, 1)
    Me.Combo1.Text = GetReceiptTitle(Me.Text3.Text)
      
   'Add By Sindy 2021/12/14
   Combo2.Clear
   Combo2.AddItem "001-1  業務助理"
   Combo2.AddItem "73017  莊敏惠"
   'Combo2.AddItem "75033  夏慧珠" Modify By Sindy 2023/7/25 Mark
   'Combo2.AddItem "89047  謝秀珠" 'cancel by sonia 2024/4/22
   Combo2.AddItem "A8029  呂麗君"
   Combo2.AddItem "20001  台中所"
   Combo2.AddItem "F5639  北京寰華"
   '2021/12/14 END
   
   OpenTable
   'Modify by Amy 2014/09/24 若為境外公司 只能為1.個人且不可改
   If PUB_GetTaxNo(Combo1, 1) = "Y" Then
        Text6 = "1"
        Text6.Locked = True
   Else
        Text6 = "2"
        Text6.Locked = False
   End If
   'end 2014/09/24
   
   'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'   Text5 = GetRecDay
'   Text5.Tag = Text5
'   '2014/3/10 add by sonia
'   If Text5 <> "" Then
'      Text5.Locked = True
'      Text5.Enabled = False
'   Else
'      Text5.Locked = False
'      Text5.Enabled = True
'   End If
'   '2014/3/10 end
   'end 2018/08/22
   
   'Remove by Morgan 2011/5/9 移到 OpenTable,+預設收據暫不列印設定
   'SetAutoTitle 'Add by Morgan 2011/3/16
   
   'Added by Morgan 2012/9/12
   'Add By Sindy 2013/12/17
   If strSrvDate(1) >= InvoiceStartDate Then
      If m_strChkCompany <> "" Then
         If m_strChkCompany = "T" Then Text1 = "1": Text1.Enabled = False
         If m_strChkCompany = "J" Then Text1 = "J": Text1.Enabled = False
      End If
      If m_CP31 = "Y" Then
         '新案
         If m_strChkCompany <> "" Then
            MsgBox "請注意," & m_strCaseNo & "有設定特殊出名公司,請檢查與接洽記錄單是否相同!!", vbInformation, "收據公司別提醒"
         'CANCEL BY SONIA 2016/7/28 因接洽單已改預設方式,故此處不再提醒
         'Else
         '   If (m_CP01 = "P" Or m_CP01 = "T") And Left(m_CP12, 1) <> "F" And m_A0j04 = "020" Then
         '      Text1 = "J": Text1.Enabled = True
         '      MsgBox "請注意,大陸新案,請注意接洽記錄單的收據公司別!!", vbInformation, "收據公司別提醒"
         '   End If
         End If
      End If
   Else
   '2013/12/17 END
      If m_strChkCompany <> "" And Text1 <> "1" And m_CP31 = "Y" Then MsgBox "請注意,專利案" & m_strCaseNo & "有設定以專利商標出名!!", vbInformation, "收據公司別提醒"
   End If
   
   txtDate = strSrvDate(2) 'Added by Morgan 2013/4/29
   
   'Add By Sindy 2013/12/17
   Command3.Visible = False 'Add By Sindy 2013/12/24
   If strSrvDate(1) >= InvoiceStartDate Then
      Label24.Visible = True
      'Modify By Sindy 2014/12/29
      Combo2.Visible = True
      'lblSales.Visible = True
      '2014/12/29 END
   Else
      Label24.Visible = False
      'Modify By Sindy 2014/12/29
      Combo2.Visible = False
      'lblSales.Visible = False
      '2014/12/29 END
   End If
     
   'Add By Sindy 2014/2/11
   m_CallForm = ""
   If m_CP140 <> "" Then '直接收文
      strCU125 = PUB_GetApplCU125(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) 'Add By Sindy 2016/12/9 組合業務備註
      If strCU125 <> "" Then strCU125 = "業務備註：" & vbCrLf & strCU125
      'Add By Sindy 2015/8/10
      'Modified by Lydia 2023/11/13 +CRL153
      strSql = "select CRL47,CRL153 from consultrecordlist where CRL01='" & m_CP140 & "'"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & adoRst.Fields("CRL47") = "Y" Then
      'Modify By Sindy 2016/12/9
            'MsgBox "請注意, 接洽記錄單有加註「貼印花」!!", vbInformation, "提醒"
            strCU125 = "請注意, 接洽記錄單有加註「貼印花」!!" & IIf(strCU125 <> "", vbCrLf & vbCrLf & strCU125, "")
         End If
         m_CRL153 = "" & adoRst.Fields("CRL153") 'Added by Lydia 2023/11/13 國內接洽單：DEBIT NOTE請款選項
      End If
      If strCU125 <> "" Then
         MsgBox strCU125, vbInformation, "提醒"
      End If
      '2016/12/9 END
      Set adoRst = Nothing
      '2015/8/10 END
      '有特殊收據資料或收據公司與預設公司不同時,顯示接洽單的特殊收據內容給使用者看再開立收據
      'Modify By Sindy 2014/10/23 收據公司與預設公司不同時,增加檢查必須為新案(m_CP31 = "Y")
      'If m_CRL119 = "Y" Or (m_CP31 = "Y" And IIf(m_CRL49 = "3", "智權公司", IIf(m_CRL49 = "2", "專利商標", "專利法律")) <> Text2) Then
      'Modify By Sindy 2020/3/31
      If m_CRL119 = "Y" Or _
         (m_CRL02 < 事務所合併日 And m_CP31 = "Y" And IIf(m_CRL49 = "3", "智權公司", IIf(m_CRL49 = "2", "專利商標", "專利法律")) <> Text2) Or _
         (m_CRL02 >= 事務所合併日 And m_CP31 = "Y" And IIf(m_CRL49 = "J", "智權", IIf(m_CRL49 = "L", "法律所", "智慧所")) <> Text2) Then
      '2020/3/31 END
         frm090801_7.SetParent Me
         frm090801_7.m_stCRL01 = m_CP140
         m_CallForm = "frm090801_7"
         frm090801_7.Show 'vbModal Modify By Sindy 2024/3/26 開特殊收據畫面不要用強制表單方式開啟
      End If
      
   'Added by Morgan 2015/12/3
   Else
      strExc(1) = PUB_ReadCP64Tag(m_CP09, "開收據提醒")
      If strExc(1) <> "" Then
         MsgBox strExc(1), vbInformation, "提醒"
      End If
   'end 2015/12/3
   End If
   '2014/2/11 END
   
   'Added by Morgan 2023/1/18
   If SetTitle(m_CP01, m_CP02, m_CP03, m_CP04, True) = True Then
      MsgBox "此案為多人共同申請，請注意收據開立方式及內容！", vbInformation, "提醒"
   End If
   'end 2023/1/18
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label7.Visible = False
   Text5.Visible = False
   
   'Add By Sindy 2022/11/30
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Command4.Visible = True
   Else
      Command4.Visible = False
   End If
   '2022/11/30 END
End Sub

'Add by Morgan 2011/3/16
'設定自動收文的收據抬頭
Private Sub SetAutoTitle(pCRL01 As String, strCP27 As String)
   Dim stSQL As String, iR As Integer, adoRst As ADODB.Recordset
   
   'Modify by Morgan 2011/5/9 +crl50
   'Modify By Sindy 2012/11/19 +CRL92
   'Modified by Lydia 2023/11/13 +CRL153
   stSQL = "select crl41,crl42,crl50,CRL92,CRL153 from consultrecordlist" & _
      " where crl01='" & pCRL01 & "'"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      
'      If adoRst("crl50") = "Y" Then Check1.Value = 1 'Add by Morgan 2011/5/9 'cancel by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選

      'Add By Sindy 2012/11/19 接洽單自動收文時要帶出智權人員當時輸入的收據自動列印時間點
      Me.Check2(0).Value = 0
      Me.Check2(1).Value = 0
      Me.Check2(2).Value = 0
      If "" & adoRst.Fields("CRL92") <> "" Then
         If "" & adoRst.Fields("CRL92") = "1" Then Me.Check2(0).Value = 1
'cancel by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選
'         'add by sonia 2015/10/28 開收據時若為送件日列印且已送件且不必再勾暫不列印
'         If "" & adoRst.Fields("CRL92") = "1" And Val(strCP27) > 0 Then
'            Check1.Value = 0: Me.Check2(0).Value = 0
'         End If
'         'end 2015/10/28
'end 2015/11/26
         If "" & adoRst.Fields("CRL92") = "2" Then Me.Check2(1).Value = 1
         If "" & adoRst.Fields("CRL92") = "3" Then Me.Check2(2).Value = 1
      'Add By Sindy 2023/9/5
      ElseIf "" & adoRst.Fields("CRL50") <> "" Then
         Check1.Value = 1
         Check1.Visible = True
         Check1.Enabled = False
         Me.Frame1.Enabled = False
         '2023/9/5 END
      End If
      '2012/11/19 End
      m_CRL153 = "" & adoRst.Fields("CRL153") 'Added by Lydia 2023/11/13 國內接洽單：DEBIT NOTE請款選項
      If adoRst("crl41") = "1" Then
         Combo1.Text = Text4
      ElseIf adoRst("crl41") = "2" Then
         'Modified by Lydia 2023/11/13
         'MsgBox "本接洽單設定為以 DEBIT NOTE 請款！", vbExclamation, "注意"
         strExc(0) = "本接洽單設定為以 DEBIT NOTE 請款！" & vbCrLf & "DEBIT NOTE請款選項："
         If m_CRL153 = "1" Then
            strExc(0) = strExc(0) & "立即開立DEBIT NOTE"
         ElseIf m_CRL153 <> "" Then
            strExc(0) = strExc(0) & "待通知後開立，" & IIf(m_CRL153 = "2", "要", "不需要") & "加印國內收據"
         End If
         'end 2023/11/13
         'Added by Lydia 2024/08/05
         If strShowCRL153 = "" Or (m_CRL153 <> strShowCRL153 And (m_CRL153 = "1" Or m_CRL153 = "3")) Then
            MsgBox strExc(0), vbExclamation, "注意"
            strShowCRL153 = m_CRL153
         End If
         'end 2024/08/05
      ElseIf adoRst("crl41") = "3" Then
         Combo1.Text = "" & adoRst("crl42")
      Else
         MsgBox "自動收文抬頭資料設定錯誤！"
      End If
      
      
   End If
   Set adoRst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q", tmpfrm) = True Then
      Unload tmpfrm 'frm090801_Q
   End If
   '2022/12/17 END
   
   If UCase(m_CallForm) = UCase("frm090801_7") Then
      m_CallForm = ""
      Unload frm090801_7
   End If
   Reference
   tool3_enabled
   Frmacc1120.strMsgShow = MsgText(602)
   Frmacc1120.Enabled = True
   'Add By Sindy 2015/2/11 瑞婷反應開收據時,若出現有人員休假訊息後,回到該畫面不會重Load資料
   Frmacc1120.CallFormActivate
   '2015/2/11 END
   Frmacc1120.Show
   Set Frmacc1121 = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   
'modify by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選,故改控制收據自動列印時間點不可輸入
'   'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
'   If Text1 = "J" Then
'      Check1.Value = 0
'      Check1.Enabled = False
'   Else
'      Check1.Enabled = True
'   End If
'   'END 2014/5/28
   'cancel by sonia 2017/12/6婷要求開放
   'If Text1 = "J" Then
   '   Check2(0).Value = 0: Check2(1).Value = 0: Check2(2).Value = 0
   '   Check2(0).Enabled = False: Check2(1).Enabled = False: Check2(2).Enabled = False
   'Else
   'end 2017/12/6
      Check2(0).Enabled = True: Check2(1).Enabled = True: Check2(2).Enabled = True
   'End If  'cancel by sonia 2017/12/6婷要求開放
'end 2015/11/26

   Select Case Text1
      Case "1"
         Text2 = MsgText(901)
      Case "2"
         'Modify By Sindy 2020/3/24
         If strSrvDate(1) >= 事務所合併日 Then
            Text2 = A0802Query(Text1, True)
         Else
         '2020/3/24 END
            Text2 = MsgText(902)
         End If
      Case "3"
         Text2 = MsgText(903)
      Case "5"
         Text2 = MsgText(904)
      Case "7"
         Text2 = MsgText(905)
      Case "8"
         Text2 = MsgText(906)
      'Add By Sindy 2013/12/19
      Case "9"
         Text2 = MsgText(908)
      'Add By Sindy 2013/12/18
      Case "J"
         Text2 = MsgText(907)
      'Add By Sindy 2020/3/24
      Case "L"
         Text2 = A0802Query(Text1, True)
   End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme  'add by sonia 2017/2/22
End Sub

'*************************************************
'  將國內收據所需之資料放置系統變數中
'
'*************************************************
Private Sub Reference()
   If Text1 <> MsgText(601) Then
      strCompanyNo = Text1
   Else
      strCompanyNo = MsgText(601)
   End If
   If Text3 <> MsgText(601) Then
      strCustNo = Text3
   Else
      strCustNo = MsgText(601)
   End If
'   If Text5 <> MsgText(601) Then
   If Me.Combo1.Text <> MsgText(601) Then
'      strTitle = Text5
      strTitle = Me.Combo1.Text
   Else
      strTitle = MsgText(601)
   End If
   If Text6 <> MsgText(601) Then
      strComPer = Text6
   Else
      strComPer = MsgText(601)
   End If
   If Text7 <> MsgText(601) Then
      strRemark = Text7
   Else
      strRemark = MsgText(601)
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      MsgBox MsgText(188) & Label1, , MsgText(5)
      Cancel = True
      Text1.SetFocus
      Exit Sub
   Else
      If adocheck.State = adStateOpen Then
         adocheck.Close
      End If
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select * from acc080 where a0801 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MsgBox MsgText(188) & Label1, , MsgText(5)
         adocheck.Close
         Cancel = True
         Text1.SetFocus
         Exit Sub
      End If
      adocheck.Close
      
'modify by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選,故改控制收據自動列印時間點不可輸入
'      'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
'      If Text1 = "J" Then
'         Check1.Value = 0
'         Check1.Enabled = False
'      Else
'         Check1.Enabled = True
'      End If
'      'END 2014/5/28
   'cancel by sonia 2017/12/6婷要求開放
   'If Text1 = "J" Then
   '   Check2(0).Value = 0: Check2(1).Value = 0: Check2(2).Value = 0
   '   Check2(0).Enabled = False: Check2(1).Enabled = False: Check2(2).Enabled = False
   'Else
   'end 2017/12/6
      Check2(0).Enabled = True: Check2(1).Enabled = True: Check2(2).Enabled = True
   'End If  'cancel by sonia 2017/12/6婷要求開放
'end 2015/11/26
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Len(Text3) = 6 Then
      Text3 = AfterZero(Text3)
   Else
      If Len(Text3) = 8 Then
         Text3 = Text3 & "0"
      End If
   End If
   Text4 = CustomerQuery(Text3, 1)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

'Add by Morgan 2008/5/5
Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> Text5.Tag Then
      '原來有日期但被清除
      If Text5 = "" Then
         Cancel = True
         MsgBox "請輸入預訂收款日！"
      Else
         '檢查格式
         If ChkDate(Text5) = False Then
            Cancel = True
         ElseIf Val(Text5) < Val(strSrvDate(2)) Then
            Cancel = True
            MsgBox "預訂收款日不可小於系統日！"
         End If
      End If
      If Cancel = True Then
         Text5_GotFocus
      End If
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = MsgText(601) Then
      MsgBox MsgText(52), , MsgText(5)
      Cancel = True
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strAPP1 As String 'Add By Sindy 2010/11/18
Dim bolChkA0K11 As Boolean 'Add By Sindy 2010/11/18
Dim str000 As String    '2010/11/22 add by sonia P,T案要判斷台灣或非台灣
Dim strSpecCompany As String
Dim strMaxA0k02 As String 'Add By Sindy 2013/12/18
'Add by Amy 2016/08/18 前畫面選取第一筆系統別/第一筆是否為新案/第一筆申請國家/客戶檔收據公司別
Dim strCP01_F As String, strCP31_F As String, strNation_F As String, strCusReceipt As String
Dim m_CU173 As String
Dim adoquery As New ADODB.Recordset 'Add by Sindy 2020/4/28
   
On Error GoTo Checking
   
   adocheck.CursorLocation = adUseClient
   'Modify By Sindy 2010/4/19
   'adocheck.Open "select a0k11 from acc0k0 where a0k01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2013/12/17 +a0k03,a0k02
   adocheck.Open "select a0k11,a0k32,a0k03,a0k02 from acc0k0 where a0k01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
'cancel by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選
'      'Add By Sindy 2010/4/19
'      If Trim(adocheck.Fields("a0k32").Value) = "N" Then
'         Check1.Value = 1
'      Else
'         Check1.Value = 0
'      End If
'      '2010/4/19 End
'end 2015/11/26
'      'Add By Sindy 2017/3/17
'      If IsNull(adocheck.Fields("a0k40").Value) = False Then
'         txtPrintNo = adocheck.Fields("a0k40").Value
'      Else
'         txtPrintNo = ""
'      End If
'      '2017/3/17 END
      If IsNull(adocheck.Fields("a0k11").Value) = False Then
         Text1 = adocheck.Fields("a0k11").Value
         adocheck.Close
         Exit Sub
      End If
   End If
   adocheck.Close
   
   'Modify By Sindy 2017/3/24
   m_CU173 = ""
   'Modify By Sindy 2019/5/22 + Text3
   Call GetTitleCustData(Combo1.Text, Text3, "", , , , , , , , , , , , , , , , , , , , , , , , , , , , m_CU173)
   txtPrintNo.Text = m_CU173
   Combo1.Tag = Combo1.Text
   '2017/3/24 END
   m_CP10List = "" 'Added by Lydia 2020/09/28
   
'   adoacc0k0.CursorLocation = adUseClient
'   adoacc0k0.Open "select distinct a0k04 from acc0k0", adoTaie, adOpenStatic, adLockReadOnly
   adocheck.CursorLocation = adUseClient
   'Modify by Morgan 2011/5/9 +cp140 預設收據暫不列印用
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adocheck.Open "select cp01, cp14, a0j04, a0j20, cp02,cp03,cp04,cp140 from caseprogress, acc0j0 where cp09 = a0j01 and a0j06 = '" & MsgText(602) & "' and (a0j13 is null or a0j13 = '')", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/27 取消 a0j20 改抓 cp10
   'adocheck.Open "select cp01, cp14, a0j04, a0j20, cp02,cp03,cp04,cp140 from caseprogress, acc0j0 where cp09 = a0j01 and a0j06 = '" & MsgText(602) & "' and a0j13=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2012/11/13 +cp151
   'Modify By Sindy 2012/12/6 +cp09
   'Modify By Sindy 2013/12/17 +cp31,a0j11,cp12
   'modify by sonia 2015/10/28 +cp27
   'modify by sonia 2023/4/17 +order by CP05,CP09 P-127920之A,C類合併開收據才會抓到A類預設接洽單資料
   adocheck.Open "select cp01, cp14, a0j04, cp10, cp02,cp03,cp04,cp140,cp151,cp09,cp31,a0j11,cp12,cp27 from caseprogress, acc0j0 where cp09 = a0j01 and a0j06 = '" & MsgText(602) & "' and a0j13=a0j01 order by cp05,cp09", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         Select Case adocheck.Fields(0).Value
            Case "T", "CFT", "TC", "CFC", "S", "TD", "TF", "TM", "TR", "TS", "TT"
               Text1 = "1"
               '93.12.21 ADD BY SONIA   TT之文件簽證711設定為2公司
               'Modified by Morgan 2011/12/27 取消 a0j20 改判斷案件性質
               'If adocheck.Fields(0).Value = "TT" And adocheck.Fields(3).Value = "文件簽證" Then
               If adocheck.Fields(0).Value = "TT" And adocheck.Fields(3).Value = "711" Then
                  Text1 = "2"
               End If
               '93.12.21 END
               'Add By Sindy 2014/4/21 商標行政訴訟設定開專利收據
               'T台灣案403.行政訴訟 205.言詞辯論 625.參加異議 207.聲明訴訟
               '       204.準備程序 407.參加訴訟 413.訴訟     410.行政上訴答辯,  412.參加上訴  預設2公司
               'modify by sonia 2016/7/28 +408行政訴訟上訴
               If adocheck.Fields(0).Value = "T" And adocheck.Fields("a0j04").Value = "000" And _
                  (adocheck.Fields(3).Value = "403" Or adocheck.Fields(3).Value = "205" Or _
                   adocheck.Fields(3).Value = "625" Or adocheck.Fields(3).Value = "207" Or _
                   adocheck.Fields(3).Value = "204" Or adocheck.Fields(3).Value = "412" Or adocheck.Fields(3).Value = "408" Or _
                   adocheck.Fields(3).Value = "413" Or adocheck.Fields(3).Value = "410" Or adocheck.Fields(3).Value = "407") Then
                  Text1 = "2"
               End If
               '2014/4/21 END
            Case "P"
               'Modify by Morgan 2006/10/16
               'If adocheck.Fields(2).Value = "020" Then
               If adocheck.Fields(2).Value <> "000" Then
                  Text1 = "1"
                  '2012/4/24 ADD BY SONIA 此日期起非台灣P案改用專利法律2公司
                  If Val(adocheck.Fields("cp02").Value) >= 101672 Then
                     Text1 = "2"
                  End If
                  '2012/4/24 END
               Else
                  Text1 = "2"
               End If
            Case "L"
               'Modify by Morgan 2011/9/7
               '原程式判斷有誤,但目前規則也已修改為除蔣律師為5公司外其餘都用 2公司--辜
'               If IsNull(adocheck.Fields(1).Value) Then
'                  Select Case adocheck.Fields(1).Value
'                     Case "76012"
'                        Text1 = "3"
'                     Case "79037"
'                        Text1 = "5"
'                     Case "82033"
'                        Text1 = "7"
'                     Case "89007"
'                        Text1 = "8"
'                     Case Else
'                        Text1 = "2"
'                  End Select
'MODIFY BY SONIA 2013/6/5 瑞婷說不再開5公司
'               If adocheck.Fields(1) = "79037" Then
'                  Text1 = "5"
'               'end 2011/9/7
'
'               Else
'                  Text1 = "2"
'               End If
               Text1 = "2"
'2013/6/5 END
            Case "LA"
               '93.6.4 MODIFY BY SONIA
               'Text1 = "1"
               ' 'Add By Cheng 2004/05/19
               ' If Val(adocheck.Fields("cp02").Value) > 2867 Then
               '     Text1 = "2"
               ' End If
               ' 'End
               Text1 = "2"
               '93.6.4 END
            Case "TB"
               'Modify by Morgan 2005/11/16
               'Text1 = "3"
               Text1 = "1"
            'Ken 92/01/08 加入CFP之判斷
            Case "CFP"
               If Val(adocheck.Fields("cp02").Value) >= 11051 Then
                  If IsNull(adocheck.Fields(2).Value) = False Then
                     Select Case adocheck.Fields(2).Value
                        Case "221", "011", "239"
                           Text1 = "1"
                           'Add By Cheng 2004/05/19
                           If "" & adocheck.Fields(2).Value = "239" And Val(adocheck.Fields("cp02").Value) > 16183 Then
                               Text1 = "2"
                           End If
                           'End
                           '93.6.4 ADD BY SONIA
                           If "" & adocheck.Fields(2).Value = "221" And Val(adocheck.Fields("cp02").Value) > 16183 Then
                               Text1 = "2"
                           End If
                           '93.6.4 End
                           '2011/2/22 ADD BY SONIA
                           If "" & adocheck.Fields(2).Value = "011" And Val(adocheck.Fields("cp02").Value) > 23914 Then
                               Text1 = "2"
                           End If
                           '2011/2/22 End
                        Case Else
                           Text1 = "2"
                     End Select
                  Else
                     Text1 = "2"
                  End If
               Else
                  Text1 = "2"
               End If
            Case Else
               Text1 = "2"
         End Select
         'Add by Morgan 2005/6/16
         If "" & adocheck.Fields("CP31") = "Y" Then   'add by sonia 2019/1/25 新案件才考慮多申請人P-121493超項費
            SetTitle adocheck.Fields("CP01"), adocheck.Fields("CP02"), adocheck.Fields("CP03"), adocheck.Fields("CP04")
         End If  'end 2019/1/25
      Else
         Text1 = "2"
      End If
      'Add By Sindy 2010/11/18
      m_CP01 = adocheck.Fields("CP01")
      m_CP02 = adocheck.Fields("CP02")
      m_CP03 = adocheck.Fields("CP03")
      m_CP04 = adocheck.Fields("CP04")
      m_CP09 = adocheck.Fields("CP09") 'Add By Sindy 2012/12/6
      m_CP31 = "" & adocheck.Fields("CP31") 'Add By Sindy 2013/12/17
      m_CP12 = "" & adocheck.Fields("CP12") 'Add By Sindy 2013/12/18
      If InStr(m_CP10List & ",", adocheck.Fields("CP10") & ",") = 0 Then m_CP10List = m_CP10List & adocheck.Fields("CP10") & ","    'Added by Lydia 2020/09/28
      'end 2020/09/28
JumpToReset: 'Added by Lydia 2020/09/28 逐筆判斷收文號是否為新案

      'Add By Sindy 2020/4/28 (E109r10184) L案號，以收文號抓法律所案源資料的LOS06，若其案源案件類型LOS02為C類時，
      '收據客戶編號自動設定為智慧所X03072010，抬頭為此客戶編號之名稱，同時將收據抬頭鎖住。
      '例L-006203(收文號AA9014217)
      strExc(0) = "select * from lawofficesource where los06='" & m_CP09 & "' and los02='C'"
      intI = 1
      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 And m_CP01 = "L" Then
         Text3 = "X03072010" '台一國際智慧財產事務所
         m_CU173 = ""
         Call GetTitleCustData("", Text3, "", , , , , , , , , , , strExc(10), , , , , , , , , , , , , , , , , m_CU173)
         txtPrintNo.Text = m_CU173
         Text4 = strExc(10)
         Combo1.Text = strExc(10)
         Combo1.Tag = Combo1.Text
         Combo1.Enabled = False
      Else
      '2020/4/28 END
         'Add By Sindy 2014/2/11 C類該案號最新收據抬頭
         strTitle = GetReceiptTitle_C(m_CP09, m_CP01 & m_CP02 & m_CP03 & m_CP04)
         If strTitle <> "" Then
            Combo1.Text = strTitle
         End If
         '2014/2/11 END
      End If
      '2010/11/18 End
'      If adocheck.Fields("a0j20").Value = "條碼" Then
'         Text1 = "3"
'      End If
      
      'Add By Sindy 2012/11/12
      m_A0j04 = "" & adocheck.Fields("a0j04").Value '申請國家
      Me.Check2(0).Value = 0
      Me.Check2(1).Value = 0
      Me.Check2(2).Value = 0
      If "" & adocheck.Fields("cp151") <> "" Then
         If "" & adocheck.Fields("cp151") = "1" Then Me.Check2(0).Value = 1
         If "" & adocheck.Fields("cp151") = "2" Then Me.Check2(1).Value = 1
         If "" & adocheck.Fields("cp151") = "3" Then Me.Check2(2).Value = 1
      End If
      '2012/11/12 End
      
      'Add by Morgan 2011/5/9
      'modify by sonia 2020/8/15 TT-999999不抓接洽單之收據抬頭,因為固定是X82357台一國際法律事務所
      If Not IsNull(adocheck("cp140")) And m_CP01 & m_CP02 <> "TT999999" Then
         SetAutoTitle adocheck("cp140"), "" & adocheck("cp27")  'modify by sonia 2015/10/28 加傳入cp27
      End If
      '2011/5/9 End
      
      'Added by Morgan 2012/9/12
      With adocheck
      .MoveFirst
      m_strChkCompany = "": m_strCaseNo = ""
      Do While Not .EOF
         'Modify By Sindy 2013/12/17
         If (.Fields("cp01") = "P" Or .Fields("cp01") = "CFP") Or _
            strSrvDate(1) >= InvoiceStartDate Then
            If strB_CP09 = "" & .Fields("cp09") Then
                strCP01_F = "" & .Fields("cp01")
                strCP31_F = "" & .Fields("cp31")
                strNation_F = "" & .Fields("a0j04")
            End If
            If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
               strSpecCompany = ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
               If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                  m_strChkCompany = strSpecCompany
                  If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                  m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
               End If
            End If
         End If
         '2013/12/17 END
         'Add By Sindy 2023/9/6 ACS代收代付,不能鎖定收據暫不列印
         If .Fields("cp01").Value = "ACS" And .Fields("cp10").Value = "706" And Me.Frame1.Enabled = False Then
            Check1.Value = 0
            Check1.Visible = False
            Check1.Enabled = True
            Me.Frame1.Enabled = True
         End If
         '2023/9/6 END
         
         'Added by Lydia 2020/09/28 逐筆判斷收文號是否為新案; 其他如抬頭可使用第一筆記錄
         If InStr(m_CP10List & ",", adocheck.Fields("CP10") & ",") = 0 Then m_CP10List = m_CP10List & .Fields("CP10") & "," '記錄收文號之案件性質
         If m_CP31 <> "Y" And "" & .Fields("CP31") = "Y" And m_CP01 & m_CP02 & m_CP03 & m_CP04 <> "" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04") Then
             m_CP01 = .Fields("CP01")
             m_CP02 = .Fields("CP02")
             m_CP03 = .Fields("CP03")
             m_CP04 = .Fields("CP04")
             m_CP09 = .Fields("CP09")
             m_CP31 = "" & .Fields("CP31")
             m_CP12 = "" & .Fields("CP12")
             GoTo JumpToReset
         End If
         'end 2020/09/28
         .MoveNext
      Loop
      .MoveFirst
      End With
      'end 2012/9/12
      'Add by Amy 2016/08/18 +新案且案件基本檔未設特殊公司別抓客戶檔收據公司別
      If m_strChkCompany = MsgText(601) And strCP31_F = "Y" Then
         strCusReceipt = GetReceiptCmp(Left(Text3, 8), Mid(Text3, 9, 1), strCP01_F, strNation_F, False)
         If strCusReceipt <> MsgText(601) Then
            Text1 = strCusReceipt
         End If
      End If
      'end 2016/08/18
   Else
      Text1 = "2"
   End If
   
   'Add By Sindy 2020/3/24
   'Modified by Lydia 2020/09/28 改成變數
   'If InStr(adocheck.Fields("CP01"), "L") > 0 And strSrvDate(1) >= 智慧所更名日 Then
   If InStr(m_CP01, "L") > 0 And strSrvDate(1) >= 智慧所更名日 Then
      Text1 = "L"
      Text1.Enabled = False
   End If
   If Text1 = "1" And strSrvDate(1) >= 事務所合併日 Then
      Text1 = "2"
   End If
   '2020/3/24 END
      
   'Add By Sindy 2020/4/1
   If Text1.Enabled = True Then
   '2020/4/1 END
      'Add By Sindy 2010/11/18
      '舊案
      'Modify By Sindy 2020/4/1 A0K11 => decode(a0k11,'1','2',a0k11) A0K11
      strSql = "select decode(a0k11,'1','2',a0k11) A0K11 from ACC0J0,ACC0K0 where A0J02='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' and A0J13=A0K01 Order By A0K02 Desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      bolChkA0K11 = False
      If intI = 1 Then
         If Trim("" & RsTemp("A0K11")) <> "" Then
            bolChkA0K11 = True
            If Trim("" & RsTemp("A0K11")) <> Trim(Text1) Then
               MsgBox "此案號最新收據的公司別 " & Trim("" & RsTemp("A0K11")) & " 與系統預設 " & Text1 & " 不同, 請注意！"
            End If
         End If
      End If
      If bolChkA0K11 = False Then
         If m_CP01 = "P" Or m_CP01 = "T" Then
            'Modified by Lydia 2020/09/28 改成變數
            'If adocheck.Fields(2).Value = "000" Then
            '   str000 = "and a0j04='" & adocheck.Fields(2).Value & "' "
            If m_A0j04 = "000" Then
               str000 = "and a0j04='" & m_A0j04 & "' "
            'end 2020/09/28
            Else
               str000 = "and a0j04<>'000' "
            End If
         ElseIf m_CP01 = "CFP" Then
            'Modified by Lydia 2020/09/28 改成變數
            'If adocheck.Fields(2).Value = "011" Then
            '   str000 = "and a0j04='" & adocheck.Fields(2).Value & "' "
            If m_A0j04 = "011" Then
               str000 = "and a0j04='" & m_A0j04 & "' "
            'end 2020/09/28
            Else
               str000 = "and a0j04<>'011' "
            End If
         Else
            str000 = ""
         End If
         'Modify By Sindy 2020/4/1 A0K11 => decode(a0k11,'1','2',a0k11) A0K11
         strSql = "select decode(a0k11,'1','2',a0k11) A0K11 from ACC0J0,ACC0K0 where A0J11='" & Text3 & "' AND SUBSTR(A0J02, 1, Length(a0j02) - 9)='" & m_CP01 & "' " & str000 & "and A0J13=A0K01 Order By A0K02 Desc "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Trim("" & RsTemp("A0K11")) <> "" Then
               If Trim("" & RsTemp("A0K11")) <> Trim(Text1) Then
                  MsgBox "此客戶最新收據的公司別 " & Trim("" & RsTemp("A0K11")) & " 與系統預設 " & Text1 & " 不同, 請注意！"
               End If
            End If
         End If
      End If
      '2010/11/18 End
   End If
   adocheck.Close
   
   'Add By Sindy 2013/12/17 若客戶在上次收據日期後有更名,要提醒操作者
   If Text3 <> "" Then
      strExc(0) = "SELECT nvl(max(a0k02),0) FROM acc0k0 WHERE a0k03='" & Text3 & "' and nvl(a0k09,0)=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            strMaxA0k02 = RsTemp.Fields(0)
            strExc(0) = "SELECT cu04 FROM customer WHERE cu01='" & Left(Text3, 8) & "' and cu02='0' and cu82>" & DBDATE(strMaxA0k02)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "請注意,客戶已更名為" & "" & RsTemp.Fields("CU04") & ",請注意欲開立之收據抬頭!!", vbInformation, "收據公司別提醒"
            End If
         End If
      End If
   End If
   '2013/12/17 END
   
Checking:
   Set adoquery = Nothing 'Add By Sindy 2020/4/28
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生並儲存國內收據資料
'
'*************************************************
'Modify by Morgan 2011/8/12 清除a0k12相關程式(目前沒作用,保留供再使用)
Private Sub ProduceData()
'add by nickc 2007/02/08
Dim strYes

   adoTaie.BeginTrans 'Added by Morgan 2013/6/14
   
On Error GoTo Checking
   lnga0k06 = 0
   lnga0k07 = 0
   If adoacc0j0.State = adStateOpen Then
      adoacc0j0.Close
   End If
   adoacc0j0.CursorLocation = adUseClient
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoacc0j0.Open "select * from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/26 +CP
   adoacc0j0.Open "select * from acc0j0,caseprogress where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01 and cp09(+)=a0j01", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0j0.RecordCount > 2 Then
'      MsgBox MsgText(91), , MsgText(5)
'      adoacc0j0.Close
'      Exit Sub
'   End If
   If adoacc0j0.RecordCount <> 0 Then
      Do While adoacc0j0.EOF = False
         'Add by Morgan 2008/1/31
         If m_bolSplitMail = False Then
            'Modified by Morgan 2011/12/26 取消 a0j03 改抓 cp10
            m_bolSplitMail = IsSplitReceipt("" & adoacc0j0.Fields("cp10"), "" & adoacc0j0.Fields("a0j02"), "" & adoacc0j0.Fields("a0j01"))
         End If
         
         If IsNull(adoacc0j0.Fields("a0j09").Value) = False Then
            lnga0k06 = lnga0k06 + Val(adoacc0j0.Fields("a0j09").Value)
         End If
         If IsNull(adoacc0j0.Fields("a0j10").Value) = False Then
            lnga0k07 = lnga0k07 + Val(adoacc0j0.Fields("a0j10").Value)
         End If
         
         'Modified by Morgan 2011/12/26 取消 a0j05 改抓 cp13
         If IsNull(adoacc0j0.Fields("cp13").Value) Then
            stra0k20 = MsgText(601)
         Else
            stra0k20 = adoacc0j0.Fields("cp13").Value
         End If
         
         'add by sonia 2017/6/22 若屬於業績列入P1001之專利處人員則智權人員改為P1001
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select * FROM SetSpecMan where ocode='P1001' and instr(oman,'" & stra0k20 & "')>0 ", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount <> 0 Then
            stra0k20 = "P1001"
         End If
         adocheck.Close
         'end 2017/6/22
         
         'Modify by Morgan 2011/9/19 a0j13改先放收文號
         'adoTaie.Execute "update acc0j0 set a0j13 = '" & strItemNo & "' where a0j06 = '" & MsgText(602) & "' and a0j13 is null"
         adoTaie.Execute "update acc0j0 set a0j13 = '" & strItemNo & "' where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01"
         'Modify By Sindy 2012/11/13 +cp151
         adoTaie.Execute "update caseprogress set cp151=" & CNULL(IIf(Me.Check2(0).Value = 1, "1", IIf(Me.Check2(1).Value = 1, "2", IIf(Me.Check2(2).Value = 1, "3", "")))) & ", cp60 = '" & strItemNo & "', cp73 = 0, cp74 = 0, cp75 = 0, cp76 = 0, cp77 = 0, cp78 = 0, cp79 = cp16 where cp09 = '" & adoacc0j0.Fields("a0j01").Value & "'"
         adoTaie.Execute "insert into acc1m0 values ('" & strItemNo & "', '" & adoacc0j0.Fields("a0j01").Value & "')"
         'Add by Morgan 2008/5/5
         '更新預定收款日期
         'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'         If Text5 <> "" Then
'            strExc(1) = DBDATE(Text5)
'            '是否新增異動紀錄
'            strExc(2) = ""
'            '檢查是否最後異動的預定收款日期與預定收款日期不同
'            strExc(0) = "select rd02*1000+rd03||rd05 from receivablesday where rd01='" & adoacc0j0("a0j01") & "' order by 1 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               '預定收款日期不同
'               If Mid(RsTemp(0), 12) <> strExc(1) Then
'                  strExc(2) = "Y"
'               End If
'            '沒有紀錄
'            Else
'               strExc(2) = "Y"
'            End If
'            If strExc(2) = "Y" Then
'               strSql = "insert into receivablesday (rd01,rd02,rd03,rd04,rd05)" & _
'                  " select '" & adoacc0j0("a0j01") & "'," & strSrvDate(1) & ",nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(Text5) & _
'                  " from receivablesday where rd01='" & adoacc0j0("a0j01") & "' and rd02=" & strSrvDate(1)
'               adoTaie.Execute strSql, intI
'            End If
'         End If
'         'end 2008/5/5
         'end 2018/08/22
         adoacc0j0.MoveNext
      Loop
      strA0K01 = strItemNo
      adoacc0j0.MoveLast
      Acc0k0Save

      adocheck.CursorLocation = adUseClient
      adocheck.Open "select * from acc0k0 where a0k01 = '" & Text8 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         'Modify By Sindy 2010/4/19
'         adoTaie.Execute "insert into acc0k0 (a0k01, a0k02, a0k03, a0k04, a0k05, a0k06, a0k07, a0k08, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k24, a0k25, a0k26, a0k30) values " & _
'                         "('" & stra0k01 & "', " & lnga0k02 & ", '" & stra0k03 & "', '" & ChgSQL(strA0K04) & "', '" & stra0k05 & "', " & lnga0k06 & ", " & lnga0k07 & ", '" & ChgSQL(stra0k08) & "', 0, '" & stra0k11 & "', 0, 0, 0, '" & stra0k20 & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & IIf(IsNull(adoacc0j0.Fields("a0j07").Value), "", adoacc0j0.Fields("a0j07").Value) & "')"
         'Modify By Sindy 2013/12/17 +A0k34
         'Modify By Sindy 2017/3/17 +A0k40
         'modify by sonia 2018/5/22 取消A0K30,統一改用A0J07
         'adoTaie.Execute "insert into acc0k0 (a0k01, a0k02, a0k03, a0k04, a0k05, a0k06, a0k07, a0k08, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k24, a0k25, a0k26, a0k30, a0k32, a0k34) values " & _
                         "('" & strA0K01 & "', " & lnga0k02 & ", '" & stra0k03 & "', '" & ChgSQL(strA0K04) & "', '" & stra0k05 & "', " & lnga0k06 & ", " & lnga0k07 & ", '" & ChgSQL(stra0k08) & "', 0, '" & stra0k11 & "', 0, 0, 0, '" & stra0k20 & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "', '" & IIf(IsNull(adoacc0j0.Fields("a0j07").Value), "", adoacc0j0.Fields("a0j07").Value) & "'," & CNULL(stra0k32) & "," & CNULL(Left(Trim(Combo2.Text), 5)) & ")"
         adoTaie.Execute "insert into acc0k0 (a0k01, a0k02, a0k03, a0k04, a0k05, a0k06, a0k07, a0k08, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k24, a0k25, a0k26, a0k32, a0k34) values " & _
                         "('" & strA0K01 & "', " & lnga0k02 & ", '" & stra0k03 & "', '" & ChgSQL(strA0K04) & "', '" & stra0k05 & "', " & lnga0k06 & ", " & lnga0k07 & ", '" & ChgSQL(stra0k08) & "', 0, '" & strA0K11 & "', 0, 0, 0, '" & stra0k20 & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "', " & CNULL(stra0k32) & "," & CNULL(Left(Trim(Combo2.Text), 5)) & ")"
         strYes = SaveAutoNo(MsgText(802), Mid(strItemNo, 5, 5))
      Else
         'Modify By Sindy 2010/4/19
         'adoTaie.Execute "update acc0k0 set a0k01 = '" & stra0k01 & "', a0k02 = " & lnga0k02 & ", a0k03 = '" & stra0k03 & "', a0k04 = '" & ChgSQL(strA0K04) & "', a0k06 = " & lnga0k06 & ", a0k07 = " & lnga0k07 & ", a0k08 = '" & ChgSQL(stra0k08) & "', a0k11 = '" & stra0k11 & "', a0k20 = '" & stra0k20 & "', a0k24 = " & Val(ACDate(ServerDate)) & ", a0k25 = " & ServerTime & ", a0k26 = '" & strUserNum & "', a0k30 = '" & IIf(IsNull(adoacc0j0.Fields("a0j07").Value), "", adoacc0j0.Fields("a0j07").Value) & "' where a0k01 = '" & Text8 & "'"
         'Modify By Sindy 2013/12/17 +A0k34
         'Modify By Sindy 2017/3/17 +A0k40
         'modify by sonia 2018/5/22 取消A0K30,統一改用A0J07
         'adoTaie.Execute "update acc0k0 set a0k01 = '" & strA0K01 & "', a0k02 = " & lnga0k02 & ", a0k03 = '" & stra0k03 & "', a0k04 = '" & ChgSQL(strA0K04) & "', a0k06 = " & lnga0k06 & ", a0k07 = " & lnga0k07 & ", a0k08 = '" & ChgSQL(stra0k08) & "', a0k11 = '" & stra0k11 & "', a0k20 = '" & stra0k20 & "', a0k24 = " & Val(strSrvDate(2)) & ", a0k25 = " & ServerTime & ", a0k26 = '" & strUserNum & "', a0k30 = '" & IIf(IsNull(adoacc0j0.Fields("a0j07").Value), "", adoacc0j0.Fields("a0j07").Value) & "', a0k32 =" & CNULL(stra0k32) & ", a0k34 =" & CNULL(Left(Trim(Combo2.Text), 5)) & " where a0k01 = '" & Text8 & "'"
         adoTaie.Execute "update acc0k0 set a0k01 = '" & strA0K01 & "', a0k02 = " & lnga0k02 & ", a0k03 = '" & stra0k03 & "', a0k04 = '" & ChgSQL(strA0K04) & "', a0k06 = " & lnga0k06 & ", a0k07 = " & lnga0k07 & ", a0k08 = '" & ChgSQL(stra0k08) & "', a0k11 = '" & strA0K11 & "', a0k20 = '" & stra0k20 & "', a0k24 = " & Val(strSrvDate(2)) & ", a0k25 = " & ServerTime & ", a0k26 = '" & strUserNum & "', a0k32 =" & CNULL(stra0k32) & ", a0k34 =" & CNULL(Left(Trim(Combo2.Text), 5)) & " where a0k01 = '" & Text8 & "'"
      End If
      adocheck.Close
      Text8 = strItemNo
   End If
CON:
   adoacc0j0.Close
   adoacc0j0.CursorLocation = adUseClient
   adoacc0j0.Open "select * from acc0j0 where a0j13 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0j0.RecordCount <> 0 Then
      adoacc0j0.MoveFirst
      'modify by sonia 2017/6/22 若屬於業績列入P1001之專利處人員則智權人員改為P1001所以a0k22要重
      'adoTaie.Execute "update acc0k0 set a0k22 = (select MAX(cp12) from caseprogress where cp60=a0k01), a0k23 = '" & IIf(IsNull(adoacc0j0.Fields("a0j04").Value), "", adoacc0j0.Fields("a0j04").Value) & "' where a0k01 = '" & strItemNo & "'"
      adoTaie.Execute "update acc0k0 set a0k22 = (select st15 from staff where a0k20=st01), a0k23 = '" & IIf(IsNull(adoacc0j0.Fields("a0j04").Value), "", adoacc0j0.Fields("a0j04").Value) & "' where a0k01 = '" & strItemNo & "'"
   End If
   adoacc0j0.Close
   adoacc0j0.CursorLocation = adUseClient
   adoacc0j0.Open "select nvl(sum(nvl(a0j09, 0)), 0), nvl(sum(nvl(a0j10, 0)), 0) from acc0j0 where a0j13 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0j0.RecordCount <> 0 Then
      adoacc0j0.MoveFirst
      adoTaie.Execute "update acc0k0 set a0k06 = " & Val(adoacc0j0.Fields(0).Value) & ", a0k07 = " & Val(adoacc0j0.Fields(1).Value) & " where a0k01 = '" & strItemNo & "'"
   End If
   adoacc0j0.Close
   
   adoTaie.CommitTrans 'Added by Morgan 2013/6/14
   'Add By Sindy 2013/12/24
   If strSrvDate(1) >= InvoiceStartDate And Text1 = "J" Then
      Command3.Visible = True
   End If
   strItemNo = ""
   
   'Add By Sindy 2023/10/19 ACS不管制
   If m_CP01 <> "ACS" Then
   '2023/10/19 END
      'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
      Call PUB_ChkCU144isN(Mid(stra0k03, 1, 8), Mid(stra0k03, 9, 1), "", Text1, , , "A") 'Add By Sindy 2023/9/4
   End If
   
Checking:
   If Err.Number = 0 Then
      strItemNo = ""
      Exit Sub
   'Added by Morgan 2013/6/14
   Else
      adoTaie.RollbackTrans
   'end 2013/6/14
   End If
   Text8 = ""
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  將國內收據所需之資料放置系統變數中
'
'*************************************************
Private Sub Acc0k0Save()
   'Ken 91/12/27 將收據日期由收文日改為系統日
   'Modify by Morgan 2011/9/8 改抓變數就好
   'lnga0k02 = Val(ACDate(ServerDate))
   'Modified by Morgan 2013/4/29
   'lnga0k02 = strSrvDate(2)
   lnga0k02 = txtDate
   'end 2013/4/29
   
   'If IsNull(adoacc0j0.Fields("a0j12").Value) Then
   '   lnga0k02 = Val(ACDate(ServerDate))
   'Else
   '   lnga0k02 = Val(adoacc0j0.Fields("a0j12").Value)
   'End If
   stra0k03 = Text3
'   strA0K04 = Text5

    'Modify by Morgan 2004/1/29
    '清除收據抬頭右方空白
   'strA0K04 = Me.Combo1.Text
   'Modified by Morgan 2013/12/30 改保留全形空白(造字會補)
   'strA0K04 = RTrim(Me.Combo1.Text)
   'Modify By Sindy 2014/11/26 因為replace會把英文名稱中的空白都去掉,這樣還是有問題
   '改判斷第一個字若為英文則Rtrim
   'strA0K04 = Replace(Me.Combo1.Text, " ", "")
   If (Asc(Left(Me.Combo1.Text, 1)) >= 65 And Asc(Left(Me.Combo1.Text, 1)) <= 90) Or _
      (Asc(Left(Me.Combo1.Text, 1)) >= 97 And Asc(Left(Me.Combo1.Text, 1)) <= 122) Or _
      (Asc(Left(Me.Combo1.Text, 1)) >= 48 And Asc(Left(Me.Combo1.Text, 1)) <= 57) Then '英數字
      strA0K04 = RTrim(Me.Combo1.Text)
   Else '非英數字
      strA0K04 = Replace(Me.Combo1.Text, " ", "")
   End If
   '2014/11/26 END
   stra0k05 = Text6
   stra0k08 = Text7
   strA0K11 = Text1
   'Add By Sindy 2010/4/19
   'modify by sonia 2015/11/26 因取消收據暫不列印故改判斷收據自動列印時間點
   'If Check1.Value = 1 Then
   'Modify By Sindy 2023/9/5 + Or Check1.Value = 1
   If (Check2(0).Value = 1 Or Check2(1).Value = 1 Or Check2(2).Value = 1) Or Check1.Value = 1 Then
      stra0k32 = "N"
   Else
      stra0k32 = ""
   End If
   'Added by Lydia 2023/11/13 國內接洽單：DEBIT NOTE請款選項
   'Modified by Lydia 2024/08/05
   'If m_CRL153 = "1" Or m_CRL153 = "3" Then  '1=立即開立DEBIT NOTE, 3=待通知後開立，不需要加印國內收據
   If strShowCRL153 = "1" Or strShowCRL153 = "3" Then
      stra0k32 = "Z"  'Z.確定不印(為了和暫不列印做取代)
   End If
   'end 2023/11/13
   
End Sub

'Add By Cheng 2003/10/07
'取得此客戶所開的收據抬頭, 並預設最近一次開的收據抬頭, 若無則預設申請人
Private Function GetReceiptTitle(strCustNo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

Me.Combo1.Clear
Me.Combo1.AddItem CustomerQuery(Text3, 1)
'Modify by Morgan 2011/3/24 +排除已作廢
StrSQLa = "Select Distinct A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and (a0k09 is null or a0k09=0) Order By 1 "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
While Not rsA.EOF
    If "" & rsA.Fields(0).Value <> Me.Combo1.List(0) Then Me.Combo1.AddItem "" & rsA.Fields(0).Value
    rsA.MoveNext
Wend
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
GetReceiptTitle = ""
'Modify by Morgan 2011/3/24 +排除已作廢
StrSQLa = "Select A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and (a0k09 is null or a0k09=0) Order By A0K02 Desc "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetReceiptTitle = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
If GetReceiptTitle = "" Then GetReceiptTitle = CustomerQuery(Text3, 1)
End Function

'Add by Morgan 2005/6/16
'多申請人收據抬頭
'Modified by Morgan 2023/1/18 +p_ChkMultiApp可判斷是否為多人申請
Private Function SetTitle(ByVal p_Cp01 As String, ByVal p_Cp02 As String, ByVal p_Cp03 As String, ByVal p_Cp04 As String, Optional p_ChkMultiApp As Boolean = False) As Boolean
   Dim ii As Integer, st_Title As String
   '專利及服務業務才有
   Select Case CheckSys(p_Cp01)
      '專利
      Case "1"
         strSql = "SELECT PA27,PA28,PA29,PA30 FROM PATENT WHERE PA01='" & p_Cp01 & "' AND PA02='" & p_Cp02 & "' AND PA03='" & p_Cp03 & "' AND PA04='" & p_Cp04 & "' AND PA27 IS NOT NULL"
      'Add By Sindy 2011/2/21
      '商標
      Case "2"
         strSql = "SELECT TM78,TM79,TM80,TM81 FROM TRADEMARK WHERE TM01='" & p_Cp01 & "' AND TM02='" & p_Cp02 & "' AND TM03='" & p_Cp03 & "' AND TM04='" & p_Cp04 & "' AND TM78 IS NOT NULL"
      '法務
      Case "3"
         strSql = "SELECT LC43,LC44,LC45,LC46 FROM LAWCASE WHERE LC01='" & p_Cp01 & "' AND LC02='" & p_Cp02 & "' AND LC03='" & p_Cp03 & "' AND LC04='" & p_Cp04 & "' AND LC43 IS NOT NULL"
      '顧問
      Case "4"
         strSql = "SELECT HC24,HC25,HC26,HC27 FROM HIRECASE WHERE HC01='" & p_Cp01 & "' AND HC02='" & p_Cp02 & "' AND HC03='" & p_Cp03 & "' AND HC04='" & p_Cp04 & "' AND HC24 IS NOT NULL"
      '2011/2/21 End
      '服務
      Case "5", "6"
         strSql = "SELECT SP58,SP59,SP65,SP66 FROM SERVICEPRACTICE WHERE SP01='" & p_Cp01 & "' AND SP02='" & p_Cp02 & "' AND SP03='" & p_Cp03 & "' AND SP04='" & p_Cp04 & "' AND SP58 IS NOT NULL"
      '其他
      Case Else
         Exit Function
   End Select
   
On Error GoTo ErrHnd
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         'Added by Morgan 2023/1/18
         SetTitle = True
         If p_ChkMultiApp = False Then
         'end 2023/1/18
            st_Title = Text4
            For ii = 0 To 3
               If Not IsNull(.Fields(ii)) Then
                  st_Title = st_Title & "  " & CustomerQuery(.Fields(ii), 1)
               End If
            Next
            Combo1 = st_Title
         End If 'Added by Morgan 2023/1/18
      End If
   End With
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub Text7_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub
'Add by Morgan 2008/1/31
'檢查是否為拆收據補收文[其他]
Private Function IsSplitReceipt(p_CP10 As String, p_CP1234 As String, p_CP09 As String) As Boolean
   Dim stCP01 As String, stCP34 As String, stSQL As String, stCP10 As String, intR As Integer
   stCP01 = Left(p_CP1234, Len(p_CP1234) - 9)
   Select Case stCP01
      Case "P", "PS", "FCP", "FG", "CFP", "CPS"
         stCP10 = "910"
      Case "L", "LA", "FCL", "CFL"
         stCP10 = "7"
      Case Else
         stCP10 = "706"
   End Select
   If p_CP10 = stCP10 Then
      stSQL = "select cp05 from caseprogress a where cp09='" & p_CP09 & "' and cp14 is null" & _
         " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05 and b.cp14 is not null and b.cp16>0)"
      intR = 1
      Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         IsSplitReceipt = True
         stCP34 = Right(p_CP1234, 3)
         If stCP34 = "000" Then
            stCP34 = ""
         Else
            stCP34 = "-" & Left(stCP34, 1) & "-" & Right(stCP34, 2)
         End If
         m_strMailSubject = "拆收據補收文【" & stCP01 & "-" & Mid(p_CP1234, Len(stCP01) + 1, 6) & stCP34 & "】"
         m_strMailDesc = "收文日：" & ChangeWStringToTDateString(RsTemp.Fields(0)) & _
                  vbCrLf & "收文號：" & p_CP09
      End If
   End If
End Function
'Add by Morgan 2008/5/5
'抓預定收款日期
Private Function GetRecDay() As String
   '同收文號抓最後異動的日期，多個收文號時抓最小的日期
   strExc(0) = "select rd05 from receivablesday" & _
      " where (rd01,rd02*1000+rd03) in (" & _
      " select rd01,max(rd02*1000+rd03) from acc0j0,receivablesday" & _
      " where a0j06 = '" & MsgText(602) & "' and rd01(+)=a0j01 group by rd01 ) order by 1 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetRecDay = TransDate(RsTemp(0), 1)
   End If
End Function

'Added by Morgan 2012/9/12
'檢查專利案是否已專利商標出名
Private Function ChkPatentNameCompany(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As String
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   ChkPatentNameCompany = ""
   'Add By Sindy 2013/12/17
   If strSrvDate(1) >= InvoiceStartDate Then
      stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161 is not null" & _
              " union select tm130 from trademark where tm01='" & pPA01 & "' and tm02='" & pPA02 & "' and tm03='" & pPA03 & "' and tm04='" & pPA04 & "' and tm130 is not null" & _
              " union select sp85 from servicepractice where sp01='" & pPA01 & "' and sp02='" & pPA02 & "' and sp03='" & pPA03 & "' and sp04='" & pPA04 & "' and sp85 is not null" & _
              " union select lc48 from lawcase where lc01='" & pPA01 & "' and lc02='" & pPA02 & "' and lc03='" & pPA03 & "' and lc04='" & pPA04 & "' and lc48 is not null"
   Else
   '2013/12/17 END
      stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161='Y'"
   End If
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      ChkPatentNameCompany = Trim("" & adoRst.Fields(0).Value)
   End If
End Function

Private Sub txtDate_GotFocus()
   TextInverse txtDate
   CloseIme
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(KeyAscii) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate <> "" Then
      If ChkDate(txtDate) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   'add by sonia 2016/12/1
   If ChkWorkDay(txtDate + 19110000) = False Then
      MsgBox Label9 & "請輸入工作日！", vbExclamation, "日期錯誤！"
      Cancel = True
      Exit Sub
   End If
   'end 2016/12/1
   'Add by Amy 2023/08/16 +不可小於830101
   If Val(txtDate + 19110000) < 19940101 Then
      MsgBox Label9 & "不可小於83/01/01！", vbExclamation, "日期錯誤！"
      Cancel = True
      txtDate.SetFocus
      Exit Sub
   End If
End Sub

''Add By Sindy 2013/12/15
'Private Sub txtSales_GotFocus()
'   txtSales.SelStart = 0
'   txtSales.SelLength = Len(txtSales.Text)
'   '儲存未修改前之值至Tag中,供再確認時使用
'   txtSales.Tag = txtSales
'   '切換輸入法
'   CloseIme
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_Validate(Cancel As Boolean)
'Dim strTemp As String, strTemp1 As String
'
'   lblSales.Caption = ""
'   If txtSales.Text <> "" Then
'      If Not ClsPDGetStaff(txtSales.Text, strTemp, strTemp1) Then
'         Cancel = True
'         Exit Sub
'      End If
'      lblSales.Caption = strTemp
'   End If
'End Sub
'Add By Sindy 2014/12/29
Private Sub Combo2_GotFocus()
   InverseTextBox Combo2
End Sub
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo2_LostFocus()
   If Combo2.Text > "" And Len(Trim(Combo2.Text)) = 5 Then
      '抓取員工姓名
      Combo2.Text = SetCboStaffName(Combo2.Text)
   End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2, 5)) = True Then
         Call Combo2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
'2014/12/29 END
