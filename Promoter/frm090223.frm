VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090223 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件催審作業"
   ClientHeight    =   4968
   ClientLeft      =   900
   ClientTop       =   1056
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4968
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdok 
      Caption         =   "延緩催審"
      CausesValidation=   0   'False
      Height          =   380
      Index           =   2
      Left            =   5040
      TabIndex        =   16
      Top             =   120
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   7455
      _ExtentX        =   13166
      _ExtentY        =   7451
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "個案作業"
      TabPicture(0)   =   "frm090223.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblC(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblC(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblC(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblC(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblC(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "GRD1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDate(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdUpdate"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdok(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCP(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "大陸案產生清單"
      TabPicture(1)   =   "frm090223.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(5)"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "Label3(6)"
      Tab(1).Control(3)=   "Label3(7)"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "txtDate(1)"
      Tab(1).Control(6)=   "txtDate(2)"
      Tab(1).Control(7)=   "cmdProc"
      Tab(1).Control(8)=   "txtDate(3)"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   3
         Left            =   -73800
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdProc 
         Caption         =   "產生催審清單"
         Height          =   285
         Left            =   -72720
         TabIndex        =   12
         Top             =   1080
         Width           =   1440
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   2
         Left            =   -72720
         MaxLength       =   7
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   1
         Left            =   -73800
         MaxLength       =   7
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox textCP 
         Height          =   285
         Index           =   1
         Left            =   1095
         MaxLength       =   3
         TabIndex        =   0
         Top             =   480
         Width           =   480
      End
      Begin VB.TextBox textCP 
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox textCP 
         Height          =   285
         Index           =   3
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   2
         Top             =   480
         Width           =   240
      End
      Begin VB.TextBox textCP 
         Height          =   285
         Index           =   4
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "尋找(&S)"
         Default         =   -1  'True
         Height          =   285
         Index           =   0
         Left            =   3020
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "新增催審內部收文"
         Height          =   285
         Left            =   5565
         TabIndex        =   6
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   0
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090223.frx":0038
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   2190
         Width           =   7215
         _ExtentX        =   12721
         _ExtentY        =   3196
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label4 
         Caption         =   "2. 產生清單後系統自動更新催審期限為三個月後。"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74520
         TabIndex        =   31
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "催審發文日："
         Height          =   210
         Index           =   7
         Left            =   -74880
         TabIndex        =   30
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "PS：1. 催審清單依代理人產生在C:\催審函的資料夾下，請自行E-MAIL。"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   6
         Left            =   -74880
         TabIndex        =   29
         Top             =   1560
         Width           =   5670
      End
      Begin VB.Line Line1 
         X1              =   -73200
         X2              =   -72600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Caption         =   "催審期限："
         Height          =   210
         Index           =   5
         Left            =   -74880
         TabIndex        =   28
         Top             =   637
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PS：催審後系統自動更新催審期限為三個月後。"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   3240
         TabIndex        =   27
         Top             =   1920
         Width           =   3780
      End
      Begin MSForms.Label lblC 
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblC 
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "案件名稱："
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   870
         Width           =   920
      End
      Begin VB.Label Label3 
         Caption         =   "申請國家："
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1230
         Width           =   920
      End
      Begin VB.Label Label3 
         Caption         =   "申請人："
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1590
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "催審期限："
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "催審發文日："
         Height          =   210
         Index           =   3
         Left            =   3480
         TabIndex        =   19
         Top             =   1230
         Width           =   1095
      End
      Begin MSForms.Label lblC 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   870
         Width           =   5940
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10477;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblC 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   14
         Top             =   1230
         Width           =   1740
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3069;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblC 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   13
         Top             =   1590
         Width           =   5700
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10054;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件進度(&C)"
      Height          =   380
      Index           =   1
      Left            =   3840
      TabIndex        =   15
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   380
      Index           =   3
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm090223"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; lblC(index)、GRD1改字型=新細明體-ExtB
'Create by Lydia 2016/01/07 案件催審作業
Option Explicit

'本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
'-------------
Dim colNP01 As Integer
Dim colNP22 As Integer
Dim colNP08 As Integer
Dim colNo As Integer
Dim strInsList As String '新增B類收文號
Dim strWordList As String '相關總收文號'新增定稿收文號
Dim strSavePath As String
'Const ET01 As String = "11" '定稿別
'Const ET03 As String = "03"
Dim ET01 As String
Dim ET03 As String
'Added by Lydia 2016/04/01 提供給案件資料及案件進度查詢(frm100101_2)的下一筆呼叫
Public Sub PubShowNextData()
End Sub
Private Sub cmdok_Click(Index As Integer)
   'Remove by Lydia 2016/04/14 不控制頁籤
   'If SSTab1.Tab > 0 And Index > 2 Then Exit Sub
   
   Select Case Index
      Case 0 '尋找
         QueryData
      Case 1 '案件進度
         If textCP(1) & textCP(2) <> "" Then
            If textCP(3) = "" Then textCP(3) = "0"
            If textCP(4) = "" Then textCP(4) = "00"
            Me.Enabled = False
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(textCP(1) & "-" & IIf(textCP(2) = "", "000000", textCP(2)) & "-" & textCP(3) & "-" & textCP(4))
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
         End If
      Case 2 '延緩催審
         frm090222.SetParent Me
         If Me.textCP(1) <> "" And Me.textCP(2) <> "" Then
            frm090222.textCP(1) = Me.textCP(1)
            frm090222.textCP(2) = Me.textCP(2)
            frm090222.textCP(3) = Me.textCP(3)
            frm090222.textCP(4) = Me.textCP(4)
            If Me.textCP(1) <> "" And Me.textCP(2) <> "" Then
               frm090222.QueryData
            End If
         End If
         frm090222.Show
         Me.Hide
      Case 3 '結束
         Unload Me
   End Select
End Sub

Private Sub cmdProc_Click()
Dim strFN As String
Dim strFileName As String
Dim idx As Integer
Dim Cancel As Boolean
Dim RsQ As New ADODB.Recordset
Dim strQ As String
Dim StrStr As String
Dim strTitle As String
Dim tmpArr As Variant
'Added by Lydia 2016/04/01
Dim strTLine As String
Dim tmpArr2 As Variant

    strSavePath = "C:\催審函"
    If Dir(strSavePath, vbDirectory) = "" Then
       MkDir strSavePath
    End If
    
    strFN = strSrvDate(2) & Left(Format(ServerTime, "000000"), 4) & "_"
    strTitle = "本所發文日,申請/註冊號,本所案號,代號,案件名稱,案件性質,申請人,進度說明"
    'Added by Lydia 2016/04/01
    strTLine = "==========,============,===============,====,============,==========,===============,============================="
    tmpArr2 = Empty
    tmpArr2 = Split(strTLine, ",")
    'end 2016/04/01
    
    tmpArr = Empty
    tmpArr = Split(strTitle, ",")
    For idx = 1 To 3
       If txtDate(idx) = "" Then
          MsgBox "日期不可空白!", vbCritical
          txtDate(idx).SetFocus
          txtDate_GotFocus idx
          Exit Sub
       Else
          Cancel = False
          txtDate_Validate idx, Cancel
          If Cancel Then
            Exit Sub
          End If
       End If
    Next idx
    If txtDate(1) > txtDate(2) Then
       MsgBox "起始日期不可大於終止日期!", vbCritical
       txtDate(1).SetFocus
       txtDate_GotFocus 1
       Exit Sub
    End If
    'Memo by Lydia 2016/04/01 因為是台灣和大陸案一起產生,所以和AutoBatchDay.strMenu70只有大陸案的不同
    strQ = "Select np01,NP08,NP10,NP22,tm01 TM01,tm02 TM02,tm03 TM03,tm04 TM04,tm10 TM10,Nvl(tm15,tm12) TM15,tm23 TM23,Nvl(tm05,Nvl(tm06,tm07)) CaseN " & _
            "From NextProgress,TradeMark Where NP02 in ('TF','T','FCT',' ') And NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) " & _
            "and tm10 = '020' And tm29 IS NULL And NP08>=" & DBDATE(txtDate(1)) & " And NP08<=" & DBDATE(txtDate(2)) & " And NP07=305 And NP06 Is Null And Not Exists" & _
            "(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='1202' or cp10='1201') And CP09>'C') "
    strQ = strQ & " Union All " & _
            "Select np01,NP08,NP10,NP22,sp01 TM01,sp02 TM02,sp03 TM03,sp04 TM04,sp09 TM10,sp11 TM15,sp08 TM23,Nvl(sp05,Nvl(sp06,sp07)) CaseN " & _
            "From NextProgress,ServicePractice Where NP02 in ('TT','TS','TR','TM','TD','TC','TB',' ') And NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) " & _
            "and sp09 = '020' And SP29 IS NULL And NP08>=" & DBDATE(txtDate(1)) & " And NP08<=" & DBDATE(txtDate(2)) & " And NP07=305 And NP06 Is Null And Not Exists" & _
            "(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='1202' or cp10='1201') And CP09>'C') "
    strQ = strQ & " Union All " & _
            "Select np01,NP08,NP10,NP22,tm01 TM01,tm02 TM02,tm03 TM03,tm04 TM04,tm10 TM10,Nvl(tm15,tm12) TM15,tm23 TM23,Nvl(tm05,Nvl(tm06,tm07)) CaseN " & _
            "From NextProgress,TradeMark Where NP02 in ('TF','T','FCT',' ') And NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) " & _
            "and tm10 = '020' And tm29 IS NULL And NP08>=" & DBDATE(txtDate(1)) & " And NP08<=" & DBDATE(txtDate(2)) & " And NP07=305 And NP06 Is Null " & _
            "And Exists(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='1202' or cp10='1201') And CP09>'C') " & _
            "And Exists(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='203' or cp10='201' or cp10='301' or cp10='302') And CP09<'C' And (CP27='' OR CP27 Is Null) )"
    strQ = strQ & " Union All " & _
            "Select np01,NP08,NP10,NP22,sp01 TM01,sp02 TM02,sp03 TM03,sp04 TM04,sp09 TM10,sp11 TM15,sp08 TM23,Nvl(sp05,Nvl(sp06,sp07)) CaseN " & _
            "From NextProgress,ServicePractice Where NP02 in ('TT','TS','TR','TM','TD','TC','TB',' ') And NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) " & _
            "and sp09 = '020' And SP29 IS NULL And NP08>=" & DBDATE(txtDate(1)) & " And NP08<=" & DBDATE(txtDate(2)) & " And NP07=305 And NP06 Is Null " & _
            "And Exists(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='1202' or cp10='1201') And CP09>'C') " & _
            "And Exists(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='203' or cp10='201' or cp10='301' or cp10='302') And CP09<'C' And (CP27='' OR CP27 Is Null) )"
         ' 本所發文日 申請/註冊號  本所案號    代號  案件名稱            案件性質   申請人            進度說明
    'TM23, CP27       TM15         CaseNo      ST17  CaseN               cpm        Xname
    strQ = "Select CP44,decode(fa04,null,decode(fa05,null,fa05||' '||fa63||' '||fa64||' '||fa65,fa06),fa04) Yname," & _
            "sqldatew(cp27) CP27,TM15, TM01||'-'||TM02||'-'||TM03||'-'||TM04 CaseNo,S1.ST17,CaseN,Nvl(Decode(TM10,'000',cpm03,cpm04),cp10) cpm,Nvl(cu04,Nvl(cu05||cu88||cu89||cu90,cu06)) Xname " & _
            ",NP01,NP08,NP10,NP22,TM01,TM02,TM03,TM04 From (" & strQ & "),CaseProgress,Staff S1,CaseProPertyMap,Customer,Fagent " & _
            "Where NP01=CP09(+) And CP14=S1.ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) " & _
            "And SubStr(TM23,1,8)=CU01(+) And Decode(SubStr(TM23,9,1),'','0',SubStr(TM23,9,1))=CU02(+) " & _
            "And SubStr(CP44,1,8)=FA01(+) And Decode(SubStr(CP44,9,1),'','0',SubStr(CP44,9,1))=FA02(+) Order by 1,3 "
            'X72897000
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
       RsQ.MoveFirst
       Do While Not RsQ.EOF
           If StrStr <> RsQ.Fields("CP44") Then
              If strFileName <> "" Then
                 g_WordAp.ActiveDocument.SaveAs strSavePath & "\" & strFileName & ".doc"
              End If
              strFileName = strFN & RsQ.Fields("CP44")
              '開啟Word
              If TypeName(g_WordAp) <> "Application" Then Set g_WordAp = New Word.Application
        
              g_WordAp.Documents.add
              g_WordAp.Visible = False
              With g_WordAp
                 .WindowState = wdWindowStateMinimize
                 '版面
                  .Selection.PageSetup.Orientation = wdOrientLandscape '橫印
                  .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
                  .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
                  .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
                  .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
                  .Selection.PageSetup.FooterDistance = .CentimetersToPoints(2)
                  .Selection.Orientation = wdTextOrientationHorizontal
                  .Selection.Font.Size = 12
                  .Selection.Font.Name = "標楷體"
                  '檔案開頭說明
                  .Selection.TypeParagraph
                 
                  '靠左
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Selection.TypeText "致: " & Trim("" & RsQ.Fields("Yname"))
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　如下案件前委託　貴公司辦理，至今已多日，尚未收到相關通知，煩請查詢目前進度，並請填入「進度說明」欄位中，以利本所回覆客戶："
                  .Selection.TypeParagraph
                  .Selection.TypeParagraph
                  .Selection.TypeParagraph
                  '新增表格
                  idx = UBound(tmpArr) + 1
                  .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=idx
                  With .Selection.Tables(1)   '清除格線
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                    .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                    .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
                    .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
                    .Borders.Shadow = False
                  End With
                 .Selection.SelectRow
                 .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.2), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                 .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
                 .Selection.Collapse Direction:=wdCollapseStart
                 '輸入表頭
                 For idx = 0 To UBound(tmpArr)
                    'Modified by Lydia 2016/04/01 抬頭字下方加虛線
                    '.Selection.TypeText Text:=tmpArr(idX)
                    .Selection.TypeText Text:=tmpArr(idx) & vbCrLf & tmpArr2(idx)
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                 Next idx
               End With 'g_WordAp
           End If
           '本所發文日 申請/註冊號  本所案號    代號  案件名稱            案件性質   申請人            進度說明
           strExc(1) = PUB_StrToStr("" & RsQ.Fields("CP27"), 10)
           strExc(2) = PUB_StrToStr("" & RsQ.Fields("TM15"), 30)
           strExc(3) = PUB_StrToStr("" & RsQ.Fields("CaseNo"), 15)
           strExc(4) = PUB_StrToStr("" & RsQ.Fields("ST17"), 4)
           strExc(5) = PUB_StrToStr("" & RsQ.Fields("CaseN"), 20)
           strExc(6) = PUB_StrToStr("" & RsQ.Fields("CPM"), 10)
           strExc(7) = PUB_StrToStr("" & RsQ.Fields("XName"), 20)
           For idx = 1 To 8
               If idx < 8 Then
                  g_WordAp.Selection.TypeText Text:=strExc(idx)
                  If idx = 4 Then '代號置中
                     g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  End If
                  g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
               Else
                  g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
               End If
           Next idx
           StrStr = RsQ.Fields("CP44")
           RsQ.MoveNext
       Loop
       g_WordAp.ActiveDocument.SaveAs strSavePath & "\" & strFileName & ".doc"
       
       If OnSaveData2(RsQ) Then
          MsgBox "催審函清單產生完成!", vbInformation + vbOKOnly
       End If
       
       g_WordAp.Quit
    Else
       MsgBox "查無資料可產生清單!", vbCritical
    End If
    
    Set RsQ = Nothing
    Set g_WordAp = Nothing
End Sub
'大陸案催審延緩
Private Function OnSaveData2(ByRef rsR As ADODB.Recordset) As Boolean
Dim strTmp As String
Dim m_NP01 As String, m_NP22 As String, m_NP08 As String
Dim m_CP12 As String, m_CP13 As String
Dim strCP09 As String
Dim New_NP08 As String  'add by sonia 2016/3/31

   OnSaveData2 = False
    
On Error GoTo ErrorHand
   'Added by Lydia 2016/04/21 +提示
    If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") > 0 Then
       If MsgBox("將會更新下一程序催審期限+3個月,並新增B類收文之催審進度,確認繼續作業? ", vbYesNo + vbDefaultButton2) = vbNo Then
          OnSaveData2 = True
          Exit Function
       End If
    End If
    'end 2016/04/21
    
    If rsR.RecordCount > 0 Then
       rsR.MoveFirst
       cnnConnection.BeginTrans
       Do While Not rsR.EOF
          m_NP01 = rsR.Fields("NP01")
          m_NP22 = rsR.Fields("NP22")
          m_NP08 = DBDATE(rsR.Fields("NP08"))
          strCP09 = ""
          strCP09 = AutoNo("B", 6)
          '催審後系統自動更新期限為三個月後
          New_NP08 = CompDate(1, 3, m_NP08)
          strTmp = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(New_NP08, True) & ", NP09=" & CNULL(New_NP08, True) & _
                      " ,NP15=NP15||'" & ChangeTStringToTDateString(strSrvDate(2)) & " " & strUserName & " 催審過延後;" & "' WHERE np01='" & m_NP01 & "' and np22=" & CNULL(m_NP22, True)
          cnnConnection.Execute strTmp, intI
          '逐筆新增B類收文之催審進度,由催審人員將催審函交程序做正式發文(上CP27)
          m_CP13 = PUB_GetAKindSalesNo(rsR.Fields("TM01"), rsR.Fields("TM02"), rsR.Fields("TM03"), rsR.Fields("TM04"))
          m_CP12 = PUB_GetStaffST15(m_CP13, 1)
          strTmp = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP27,CP20,CP32,CP112,CP43,CP85) " & _
                        "VALUES ('" & rsR.Fields("TM01") & "','" & rsR.Fields("TM02") & "','" & rsR.Fields("TM03") & "','" & rsR.Fields("TM04") & "'," & strSrvDate(1) & "," & m_NP08 & "," & m_NP08 & "," & _
                        "'" & strCP09 & "','305','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                         CNULL(DBDATE(txtDate(3)), True) & ",'N','N','N','" & m_NP01 & "','" & CNULL(DBDATE(txtDate(3)), True) & "') "
          cnnConnection.Execute strTmp, intI
          rsR.MoveNext
       Loop
       cnnConnection.CommitTrans
    End If
    OnSaveData2 = True
    Exit Function
    
ErrorHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

'2025/3/17 sonia改按鈕名稱為新增催審內部收文，原為新增催審記錄，因T-249303按錯故修改
Private Sub cmdUpdate_Click()
Dim oText As TextBox
Dim Cancel As Boolean
Dim bCheck As Boolean
Dim i As Integer
Dim tmpArr As Variant
Dim idx As Integer
Dim st07_1 As String, st07_2 As String
Dim ArrCP09 As Variant 'Added by Lydia 2016/06/06

   Cancel = False
   For Each oText In textCP
       If oText = "" Then
          MsgBox "本所案號不可空白!", vbCritical
          Exit Sub
       End If
       textCP_Validate oText.Index, Cancel
       If Cancel Then
          Exit Sub
       End If
   Next
   
   '催審發文日
   If txtDate(0) = "" Then
       MsgBox "催審發文日不可空白!", vbCritical
       txtDate(0).SetFocus
       Exit Sub
   End If
   txtDate_Validate 0, Cancel
   If Cancel Then Exit Sub
   
   bCheck = False
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
        If GRD1.TextMatrix(i, colNo) <> textCP(1) & textCP(2) & textCP(3) & textCP(4) Then
           MsgBox "本所案號與催審期限記錄不一致!", vbCritical
           Exit Sub
        Else
           bCheck = True
        End If
      End If
   Next i
   
   If bCheck = False Then
      MsgBox "未選取催審期限記錄!", vbCritical
      Exit Sub
   Else
      intI = 1
      '操作者專業代號,分機號碼
      strExc(0) = "select st07,ed01 from staff,ExtensionData where st01='" & strUserNum & "' and st01=ed02(+) "
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         st07_1 = "" & RsTemp.Fields("st07")
         st07_2 = "" & RsTemp.Fields("ed01")
      End If
      
      If textCP(1) = "T" Then
         ET01 = "11" '定稿別
         ET03 = "03"
      ElseIf textCP(1) = "FCT" Then
         ET01 = "11"
         ET03 = "00"
      End If
      If OnSaveData Then '新增B類收文和更新下一程序
         If UpdCP123CP130 Then '顯示發文項目
            '逐筆產生定稿
            tmpArr = Empty
            tmpArr = Split(strWordList, ",")
            'Added by Lydia 2016/06/06
            ArrCP09 = Empty
            ArrCP09 = Split(strInsList, ",")
            'end 2016/06/06
            For idx = 0 To UBound(tmpArr)
               If tmpArr(idx) <> "" Then
                  'Modified by Lydia 2016/06/06 以B類收文做定稿的收文號
                  'If InsExpField(Trim(TmpArr(idx)), st07_1, st07_2) = False Then
                  If InsExpField(Trim(ArrCP09(idx)), st07_1, st07_2, Trim(tmpArr(idx))) = False Then
                     Exit For
                  End If
                  'Modified by Lydia 2016/06/06 以B類收文做定稿的收文號
                  'NowPrint Trim(TmpArr(idx)), ET01, ET03, False, strUserNum, 0
                  NowPrint Trim(ArrCP09(idx)), ET01, ET03, False, strUserNum, 0
               End If
            Next idx
         End If
         QueryData
      End If

   End If
End Sub

Private Sub Form_Load()
Dim tmpDate As String

   MoveFormToCenter Me
   FormClear
   SetGrd
   '依系統日期預設當月之1~15日或16~月底
   tmpDate = strSrvDate(1)
   'Modified by Lydia 2016/04/14 改預設為當期
'    If Val(Right(tmpDate, 2)) > 15 Then
'        txtDate(1) = ChangeWStringToTString(Left(tmpDate, 6) & "01")
'        txtDate(2) = ChangeWStringToTString(Left(tmpDate, 6) & "15")
'    Else
'        tmpDate = CompDate(1, -1, tmpDate)
'        txtDate(1) = ChangeWStringToTString(Left(tmpDate, 6) & "16")
'        txtDate(2) = ChangeWStringToTString(GetLastDay(tmpDate))
'    End If
   'Modified by Lydia 2016/04/21 Debug
   If Val(Right(tmpDate, 2)) > 15 Then
       txtDate(1) = ChangeWStringToTString(Left(tmpDate, 6) & "16")
       txtDate(2) = ChangeWStringToTString(GetLastDay(tmpDate))
   Else
       txtDate(1) = ChangeWStringToTString(Left(tmpDate, 6) & "01")
       txtDate(2) = ChangeWStringToTString(Left(tmpDate, 6) & "15")
   End If
   txtDate(3) = strSrvDate(2)
   SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090223 = Nothing
End Sub

Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String
   
   If textCP(1) = "" Or textCP(2) = "" Then
      MsgBox "請輸入案號!!!", vbExclamation + vbOKOnly
      If textCP(1) = "" Then
         Me.textCP(1).SetFocus
      ElseIf textCP(2) = "" Then
         Me.textCP(2).SetFocus
      End If
      Exit Sub
   End If

   If textCP(3) = "" Then textCP(3) = "0"
   If textCP(4) = "" Then textCP(4) = "00"
   
   m_TM01 = Trim(textCP(1))
   m_TM02 = Trim(textCP(2))
   m_TM03 = Trim(textCP(3))
   m_TM04 = Trim(textCP(4))
   If ClsPDCheckCaseCodeIsExist(m_TM01, m_TM02, m_TM03, m_TM04, strExc(1), strExc(2), strExc(3), strExc(4), strExc(5)) Then
      lblC(0).Caption = IIf(strExc(1) <> "", strExc(1), IIf(strExc(2) <> "", strExc(2), strExc(3)))
      lblC(4).Caption = strExc(4)
      lblC(1).Caption = strExc(5)
      If ClsPDGetNation(strExc(5), strExc(6)) Then
         lblC(2).Caption = strExc(6)
      End If
   Else
      FormClear
      Exit Sub
   End If
   'Added by Lydia 2016/04/14 +NP15
   strSql = "SELECT '' V,NP01,CP10,NVL(C2.CPM03,CP10) AS CP10M,NVL(S1.ST02,NP10) AS NP10,NVL(CP27 - 19110000, NULL) AS CP27," & _
            "NVL(NP08 - 19110000, NULL) AS NP08,NP15,NP22,NVL(S2.ST02,CP14) AS CP14,CP01||CP02||CP03||CP04 CASENO " & _
            "FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
            "WHERE NP02='" & m_TM01 & "' AND NP03='" & m_TM02 & "' AND NP04='" & m_TM03 & "' AND NP05='" & m_TM04 & "' " & _
            "AND NP06 IS NULL AND NP07='305' AND NP01 = C1.CP09(+) AND NP10 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP01 = C2.CPM01(+) AND CP10 = C2.CPM02(+) " & _
            "ORDER BY NP08,NP01,CP27 "
   '非台灣
   If lblC(1) > "010" Then
      Label3(3).Visible = False
      strSql = Replace(strSql, "C2.CPM03", "C2.CPM04")
      txtDate(0).Visible = False
      cmdUpdate.Visible = False
   Else
      Label3(3).Visible = True
      txtDate(0).Visible = True
      cmdUpdate.Visible = True
      txtDate(0).Text = strSrvDate(2)
   End If
   
   Call SetGrd 'Added by Lydia 2018/02/12
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      'Modified by Lydia 2018/02/12
      'SetGrd (rsTmp.RecordCount + 1)
      Call SetGrd(False)
      
      'Added by Lydia 2016/04/14 只有一筆,預設勾選
      If rsTmp.RecordCount = 1 Then
         GRD1.TextMatrix(1, 0) = "V"
      End If
      If txtDate(0).Visible = True Then txtDate(0).SetFocus
   Else
      MsgBox "無下一程序催審的記錄!!!", vbExclamation + vbOKOnly
      Me.textCP(2).SetFocus
      rsTmp.Close
      Set rsTmp = Nothing
      'SetGrd 'Remove by Lydia 2018/02/12
      Exit Sub
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Modified by Lydia 2018/02/12
'Private Sub SetGrd(Optional ByVal iR As Integer = 2)
Private Sub SetGrd(Optional ByVal pReset As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   'v, NP01,CP10,CP10M,NP10,CP27,NP08,NP15,NP22,CP14,CASENO
   arrGridHeadText = Array("V", "總收文號", "CP10", "案件性質", "NP10", "發文日", "催審期限", "備註", "NP22", "CP14", "CASENO")
   'Modified by Lydia 2016/05/20 備註加寬2200->4500
   arrGridHeadWidth = Array(200, 1100, 0, 1200, 0, 900, 900, 4500, 0, 0, 0)
   
   GRD1.Visible = False

   With GRD1
        .Cols = UBound(arrGridHeadText) + 1
        'Modified by Lydia 2018/02/12 預設先清空
        '.Rows = iR
        If pReset = True Then
             GRD1.Clear
             GRD1.Rows = 2
        End If
        'end 2018/02/12
        For iRow = 0 To .Cols - 1
           .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           .ColWidth(iRow) = arrGridHeadWidth(iRow)
        Next
   End With
   
   If colNo = 0 Then colNo = PUB_MGridGetId("CASENO", GRD1)
   If colNP01 = 0 Then colNP01 = PUB_MGridGetId("總收文號", GRD1)
   If colNP08 = 0 Then colNP08 = PUB_MGridGetId("催審期限", GRD1)
   If colNP22 = 0 Then colNP22 = PUB_MGridGetId("NP22", GRD1)
   
   GRD1.Visible = True

End Sub
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
If nCol >= 0 Then GRD1.col = nCol
If nRow >= 0 Then GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim TmpRow As Integer
Dim jj As Integer

TmpRow = GRD1.MouseRow

GRD1.Visible = False
If TmpRow > 0 Then
   '目前資料列反白
   With GRD1
      .col = 0
      .row = TmpRow
      '目前勾選資料列反白
      If .Text = "" And .TextMatrix(.row, 1) <> "" Then
          .Text = "V"
           For jj = 0 To .Cols - 1
              .col = jj
              .CellBackColor = &HFFC0C0
           Next jj
      '不勾選資料列清除反白
      ElseIf .TextMatrix(.row, 1) <> "" Then
          .Text = ""
           For jj = 0 To .Cols - 1
              .col = jj
              .CellBackColor = QBColor(15)
           Next jj
      End If
   End With
End If
GRD1.Visible = True
End Sub
Private Sub FormClear()
Dim oLabel As Object
Dim oText As TextBox

For Each oText In textCP
   oText.Text = ""
Next

For Each oText In txtDate
   oText.Text = ""
Next

Label3(3).Visible = False
txtDate(0).Visible = False
cmdUpdate.Visible = False
For Each oLabel In lblC
   oLabel.Caption = ""
Next

End Sub
'個案催審延緩--新增B類收文和更新下一程序
Private Function OnSaveData() As Boolean
Dim strTmp As String
Dim m_NP01 As String, m_NP22 As String, m_NP08 As String
Dim m_CP12 As String, m_CP13 As String
Dim ii As Integer
Dim strCP09 As String
Dim New_NP08 As String  'add by sonia 2016/3/31
   
   OnSaveData = False
   strInsList = ""
   strWordList = ""

   cnnConnection.BeginTrans
   With GRD1
       For ii = 1 To .Rows - 1
          If .TextMatrix(ii, 0) = "V" Then
             m_NP01 = .TextMatrix(ii, colNP01)
             m_NP22 = .TextMatrix(ii, colNP22)
             m_NP08 = DBDATE(.TextMatrix(ii, colNP08))
             strExc(1) = ""
             strCP09 = ""
             strCP09 = AutoNo("B", 6)
             '催審後系統自動更新期限為三個月後
             'Modified by Lydia 2025/11/12 改抓最近工作天 +UB_GetWorkDay1
             New_NP08 = PUB_GetWorkDay1(CompDate(1, 3, m_NP08), True)
             strTmp = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(New_NP08, True) & ", NP09=" & CNULL(New_NP08, True) & _
                           " ,NP15=NP15||'" & ChangeTStringToTDateString(strSrvDate(2)) & " " & strUserName & " 催審過延後;" & "' WHERE np01='" & m_NP01 & "' and np22=" & CNULL(m_NP22, True)
             cnnConnection.Execute strTmp, intI
             '逐筆新增B類收文之催審進度,由催審人員將催審函交程序做正式發文(上CP27)
             m_CP13 = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
             m_CP12 = PUB_GetStaffST15(m_CP13, 1)
             strTmp = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP32,CP112,CP43,CP84,CP85) " & _
                             "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & m_NP08 & "," & m_NP08 & "," & _
                             "'" & strCP09 & "','305','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                             "'N','N','N','" & m_NP01 & "',0,'" & CNULL(DBDATE(txtDate(0)), True) & "') "
             cnnConnection.Execute strTmp, intI
             strInsList = strInsList & strCP09 & ","
             strWordList = strWordList & m_NP01 & ","
          End If
       Next
   End With
   cnnConnection.CommitTrans
    
   OnSaveData = True
   Exit Function
   
ErrHand2:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function
'個案催審延緩--顯示發文項目
Private Function UpdCP123CP130() As Boolean
Dim tmpArr As Variant
Dim ii As Integer
Dim strTxt() As String
Dim iStep As Integer
Dim strCP123 As String, strCP130 As String

UpdCP123CP130 = False

   If strInsList <> "" Then
      tmpArr = Empty
      tmpArr = Split(strInsList, ",")
      iStep = 1
      Erase strTxt
      ReDim Preserve strTxt(1 To UBound(tmpArr) + 1)
      
      For ii = 0 To UBound(tmpArr)
         If tmpArr(ii) <> "" Then
            strCP123 = "": strCP130 = ""
            strExc(1) = tmpArr(ii): strExc(2) = ""
            If ModifyDispatchCp130(strExc(1), strExc(2), strCP123, strCP130) = True Then
               strTxt(iStep) = "UPDATE caseprogress set CP123='" & strCP123 & "',CP130='" & strCP130 & "' where cp09='" & tmpArr(ii) & "' "
               iStep = iStep + 1
            End If
         End If
      Next ii
   End If
  
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存發文項目失敗，請洽系統管理員 !", vbCritical
      Exit Function
   End If
   
   UpdCP123CP130 = True
   Exit Function

End Function
'Modified by Lydia 2016/06/06 +相關總收文號sCP43
Private Function InsExpField(sCP09 As String, sST07 As String, sST07_2 As String, sCP43 As String) As Boolean
Dim ArrXno As Variant
Dim jj As Integer
Dim strTxt(1 To 99) As String
Dim iStep As Integer
Dim strTmp As String
Dim rsR1 As New ADODB.Recordset

InsExpField = False
    iStep = 1
    intI = 1
    EndLetter ET01, sCP09, ET03, strUserNum
    'Modified by Lydia 2016/06/06
    'strTmp = " select '1' ord1,cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp14,cp27,cpm03," & _
             "decode(tm01,null,sp08||','||sp58||','||sp59||','||sp65||','||sp66,tm23||','||tm78||','||tm79||','||tm80||','||tm81) Xno " & _
             "From caseprogress,trademark, servicepractice,casepropertymap " & _
             "where cp09='" & sCP09 & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) "
    '申請意見書會先有核駁前先行通知(1202)
   ' strTmp = strTmp & " union select '2' ord1,c2.cp01,c2.cp02,c2.cp03,c2.cp04,c2.cp05,c2.cp09,c2.cp10,c2.cp14,c2.cp27,cpm03," & _
            "decode(tm01,null,sp08||','||sp58||','||sp59||','||sp65||','||sp66,tm23||','||tm78||','||tm79||','||tm80||','||tm81) Xno " & _
            "from caseprogress c1,caseprogress c2,trademark, servicepractice,casepropertymap " & _
            "where c1.cp43='" & sCP09 & "' and c1.cp10='1202' and c1.cp57 is null and c1.cp09=c2.cp43(+) and c2.cp10='202' and c2.cp57 is null " & _
            "and c2.cp01=tm01(+) and c2.cp02=tm02(+) and c2.cp03=tm03(+) and c2.cp04=tm04(+) and c2.cp01=sp01(+) and c2.cp02=sp02(+) and c2.cp03=sp03(+) and c2.cp04=sp04(+) " & _
            "and c2.cp01=cpm01(+) and c2.cp10=cpm02(+) "
    strTmp = " select '1' ord1,cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp14,cp27,cpm03," & _
             "decode(tm01,null,sp08||','||sp58||','||sp59||','||sp65||','||sp66,tm23||','||tm78||','||tm79||','||tm80||','||tm81) Xno " & _
             "From caseprogress,trademark, servicepractice,casepropertymap " & _
             "where cp09='" & sCP43 & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) "
    '申請意見書會先有核駁前先行通知(1202)
    strTmp = strTmp & " union select '2' ord1,c2.cp01,c2.cp02,c2.cp03,c2.cp04,c2.cp05,c2.cp09,c2.cp10,c2.cp14,c2.cp27,cpm03," & _
            "decode(tm01,null,sp08||','||sp58||','||sp59||','||sp65||','||sp66,tm23||','||tm78||','||tm79||','||tm80||','||tm81) Xno " & _
            "from caseprogress c1,caseprogress c2,trademark, servicepractice,casepropertymap " & _
            "where c1.cp43='" & sCP43 & "' and c1.cp10='1202' and c1.cp57 is null and c1.cp09=c2.cp43(+) and c2.cp10='202' and c2.cp57 is null " & _
            "and c2.cp01=tm01(+) and c2.cp02=tm02(+) and c2.cp03=tm03(+) and c2.cp04=tm04(+) and c2.cp01=sp01(+) and c2.cp02=sp02(+) and c2.cp03=sp03(+) and c2.cp04=sp04(+) " & _
            "and c2.cp01=cpm01(+) and c2.cp10=cpm02(+) "
    strTmp = strTmp & " order by ord1 asc,cp27 desc "
    Set rsR1 = ClsLawReadRstMsg(intI, strTmp)
    If intI = 1 Then
        rsR1.MoveFirst
        
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','操作者專業代號','" & sST07 & "')"
        iStep = iStep + 1
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','操作者分機號碼','" & sST07_2 & "')"
        iStep = iStep + 1
        
        '若有申請意見書(202),說明內容改成申請意見書
        If rsR1.Fields("cp10") = "101" And rsR1.RecordCount > 1 Then rsR1.MoveNext
        '說明
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','指定發文日','" & "" & rsR1.Fields("cp27") & "')"
        iStep = iStep + 1
        'Modified by Lydia 2016/06/04 101申請函的內文使用"註冊"而非申請
        'strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','指定案件性質','" & "" & rsR1.Fields("cpm03") & "')"
        'iStep = iStep + 1
        
        'add by sonia 2017/4/19 定稿之事件改用 '<商標卷宗性質/事件>
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','商標卷宗性質/事件','" & GetCaseType3(textCP(1), textCP(2), textCP(3), textCP(4), sCP43) & "')"
        iStep = iStep + 1
        'end 2017/4/19
        
        If rsR1.Fields("cp10") = "101" Then '申請=註冊
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
                "','指定案件性質','註冊申請')"
            iStep = iStep + 1
        ElseIf rsR1.Fields("cp10") = "202" Then '申請意見書
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
                "','指定案件性質','" & "" & rsR1.Fields("cpm03") & "')"
            iStep = iStep + 1
        Else                                    '其他
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
                "','指定案件性質','" & "" & rsR1.Fields("cpm03") & "')"
            iStep = iStep + 1
        End If
        'end 2016/06/04

        strExc(1) = ""
        If "" & rsR1.Fields("cp27") <> "" Then
           intI = DateDiff("m", ChangeTStringToTDateString(strSrvDate(1)), ChangeTStringToTDateString(rsR1.Fields("cp27")))
           If intI < 1 Then intI = intI * -1
        End If
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','逾期月數','" & Format(intI, "##0") & "')"
        iStep = iStep + 1
        'Added by Lydia 2016/06/07 因為申請函的代理人注重順序(發文有可能影響),所以指定代理人(林+閻)
        strExc(0) = "94007;81040"
        strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
            "','催審申請函代理人','" & Replace(PUB_ReadUserData(strExc(0)), ",", "、") & "')"
        iStep = iStep + 1
        
        '---------------------
        '申請人的代表人s
        ArrXno = Empty
        ArrXno = Split("" & rsR1.Fields("Xno"), ",")
        For jj = 0 To UBound(ArrXno)
           If Len(ArrXno(jj)) > 6 Then
              strTmp = "select cu39,cu42,cu45,cu48,cu51,cu54 from customer where cu01='" & Mid(ArrXno(jj), 1, 8) & "' and cu02='" & Mid(ArrXno(jj), 9, 1) & "' "
              intI = 1: strExc(1) = "": strExc(2) = "　　　　　　　　"
              Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
              If intI = 1 Then
                 If "" & RsTemp(0) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(0)
                 End If
                 If "" & RsTemp(1) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(1)
                 End If
                 If "" & RsTemp(2) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(2)
                 End If
                 If "" & RsTemp(3) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(3)
                 End If
                 If "" & RsTemp(4) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(4)
                 End If
                 If "" & RsTemp(5) <> "" Then
                    strExc(1) = strExc(1) & IIf(Len(strExc(1)) > 0, vbCrLf & strExc(2), "") & RsTemp(5)
                 End If
                 If strExc(1) <> "" Then
                     strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & ET01 & "','" & sCP09 & "','" & ET03 & "','" & strUserNum & _
                         "','" & "申請人" & jj + 1 & "代表人s" & "','" & strExc(1) & "')"
                     iStep = iStep + 1
                 End If
              End If
           End If
        Next jj
        '--------------------------
    End If
         
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存發文項目失敗，請洽系統管理員 !", vbCritical
      Exit Function
   End If
   
   InsExpField = True
   Exit Function

End Function
Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
       txtDate(1).SetFocus
       txtDate_GotFocus 1
    End If
End Sub

Private Sub textCP_GotFocus(Index As Integer)
    TextInverse textCP(Index)
    CloseIme
End Sub

Private Sub textCP_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 1
           If textCP(Index).Text <> "" And InStr(textCP(Index), "T") = 0 Then
               MsgBox "限定使用商標案!!", vbCritical
               textCP(Index).SetFocus
               Cancel = True
           End If
   End Select
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
    TextInverse txtDate(Index)
    CloseIme
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String
    Select Case Index
       Case 0, 3 '催審發文日=>CP85
            If txtDate(Index) <> "" Then
               strTmp = txtDate(Index)
               If CheckIsTaiwanDate(strTmp) = False Then
                  GoTo JumpCancel
               Else
                  strTmp = ChangeWDateStringToWString(strTmp)
                  If ChkWorkDay(strTmp) = False Then
                     MsgBox "催審發文日必須為工作天!", vbCritical
                     GoTo JumpCancel
                  End If
               End If
            End If
       Case Else
            If txtDate(Index) <> "" Then
               strTmp = txtDate(Index)
               If CheckIsTaiwanDate(strTmp) = False Then
                  GoTo JumpCancel
               End If
            End If
    End Select
    
    Exit Sub
    
JumpCancel:
    txtDate(Index).SetFocus
    txtDate_GotFocus Index
    Cancel = True
End Sub
