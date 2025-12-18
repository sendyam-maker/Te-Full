VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090608_1 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "┯快笷Θ薄琩高"
   ClientHeight    =   5724
   ClientLeft      =   -2496
   ClientTop       =   1836
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9324
   Begin VB.CheckBox Check1 
      Caption         =   "陪ボだ膀计だ瞒逆"
      Height          =   255
      Left            =   225
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "玡礶(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7944
      TabIndex        =   3
      Top             =   204
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4620
      Left            =   30
      TabIndex        =   2
      Top             =   765
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   8149
      _Version        =   393216
      Rows            =   3
      Cols            =   15
      FixedRows       =   2
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "穝灿砰-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   396
      Width           =   1470
   End
   Begin VB.Label lblMemo2 
      Caption         =   "“猔種: 108/4/1癬 1.穨叭Μゅ翴计ч衡Θ璸ン 2.糤砞龟罿翴计(=璸ンン翴计)"
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   5460
      Width           =   9195
   End
   Begin VB.Label lblMemo 
      Caption         =   "“祇ゅ笷Θ瞯キА=(祇ゅ龟罿翴计笷Θ瞯+祇ゅ膀计笷Θ瞯)/2"
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   2835
      TabIndex        =   4
      Top             =   330
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "╰参摸"
      Height          =   180
      Left            =   168
      TabIndex        =   1
      Top             =   408
      Width           =   1008
   End
End
Attribute VB_Name = "frm090608_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 эΘForm2.0 ; grd1э=穝灿砰-ExtB
'Memo By Morgan 2012/12/10 醇舦逆э
'2010/12/1 memo by sonia 絪腹逆э
'Memo by Morgan2010/8/16 ら戳逆э
Option Explicit
Dim k As Integer
Dim m_INSERTLOG As Boolean    '2012/1/20 ADD BY SONIA  FORM LOAD琩高挡狦璶LOG,COMBO1ち传琩高ぃゲ
Public m_bol108Rule As Boolean 'Added by Morgan 2019/3/22 盡矪108σ

Private Sub Check1_Click()
   If Check1.Visible = True Then Process
End Sub

Private Sub cmdok_Click()
   'edit by nickc 2005/03/04 э到 cpu  100 % 薄猵
   'Me.Hide
   frm090608.Show
   Unload Me
End Sub

Private Sub Combo1_Click()
   Process
End Sub

'Added by Morgan 2025/7/18
Private Sub Form_Activate()
   Static bolDone As Boolean
   
   If Not bolDone Then
      bolDone = True
      If grd1.Width < 11400 Then
         If mdiMain.Width > Me.Width Then
            If mdiMain.Width > 12000 Then
               Me.Width = 11800
               grd1.Width = Me.Width - 200
            Else
               Me.Width = mdiMain.Width
               grd1.Width = Me.Width - 200
            End If
            MoveFormToCenter Me
         End If
      End If
   End If
   
End Sub

Private Sub Form_Load()
   m_INSERTLOG = True  '2012/1/20 ADD BY SONIA
   
   MoveFormToCenter Me
   
   'Added by Morgan 2018/11/15
   If bolNewPromoterRule And frm090608.txt1(12) = "1" And frm090608.Check1.Value = vbUnchecked Then
      Check1.Visible = True
   End If
   'end 2018/11/15
   
   'Added by Morgan 2025/7/18
   If frm090608.txt1(12) = "3" Then
      lblMemo2.Visible = False
   End If
   'end 2025/7/18
   
   SetGrd1
   StrMenu
   Process
   
   m_INSERTLOG = False  '2012/1/20 ADD BY SONIA
End Sub

Sub StrMenu()
   strSql = "SELECT DISTINCT R102001 FROM R090608 WHERE ID='" & strUserNum & "' "
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           k = 0
           Do While .EOF = False
               'Modified by Morgan 2021/7/14 ACS北ALL㏕﹚逼材1
               'Combo1.AddItem CheckStr(.Fields(0)), k
               If .Fields(0) = "ALL" Then
                  Combo1.AddItem CheckStr(.Fields(0)), 0
               Else
                  Combo1.AddItem CheckStr(.Fields(0)), k
               End If
               'end 2021/7/14
               k = k + 1
               .MoveNext
           Loop
       End If
   End With
   CheckOC
   Combo1.Text = Combo1.List(0)
   
End Sub

Sub Process()
   
Dim ii As Integer, jj As Integer
   'Added by Morgan 2014/9/24
   If frm090608.m_bolShowMemo And Combo1.Text = "ALL" Then
      'Added by Morgan 2019/3/25 108σ
      'Removed by Morgan 2025/7/17 Ν絬钡эンゅ
      'If m_bol108Rule Then
      '   lblMemo = "“祇ゅ笷Θ瞯キА=(祇ゅ龟罿翴计笷Θ瞯+祇ゅ膀计笷Θ瞯)/2"
      'End If
      'end 2025/7/17
      'end 2019/3/25
      lblMemo.Visible = True
   Else
      lblMemo.Visible = False
   End If
   'end 2014/9/24
      
   'Modified by Morgan 2018/11/6 +R102015:ㄓ方 >> 1=ン, 2=や穿, 3=疭, 4=Μゅ
   'Modified by Morgan 2018/11/7 +R102016:だ皌翴计,R102017:ぃ膀计,R102018:ン计
   'Modified by Morgan 2018/11/15
   'Modified by Morgan 2019/3/22 +祇ゅ龟罿翴计(R102022),祇ゅ龟罿翴计笷Θ瞯(R102023),粂猭虏て
   If Check1.Value = vbChecked Then
      strSql = "select ST03,round(sum(NVL(r102003,0)),2),round(sum(NVL(r102004,0)),2),round(sum(NVL(r102005,0)),2),round(sum(DECODE(r102015,'1',nvl(r102016,0),0)),2),round(sum(NVL(r102022,0)),2),round(sum(NVL(r102006,0)),2),round(sum(DECODE(r102015,'1',nvl(r102017,0),0)),2),round(sum(DECODE(r102015,'1',nvl(r102006,0),0)),2),round(sum(DECODE(r102015,'2',nvl(r102006,0),0)),2)"
      strSql = strSql & ",round(sum(DECODE(r102015,'4',nvl(r102006,0),0)),2),round(sum(DECODE(r102015,'1',nvl(r102018,0),0)),2),round(sum(NVL(r102007,0)),2),round(sum(NVL(r102023,0)),2),round(sum(NVL(r102008,0)),2),round(sum(NVL(r102009,0)),2),round(sum(NVL(r102010,0)),2),round(sum(DECODE(r102015,'1',nvl(r102019,0),0)),2),round(sum(NVL(r102011,0)),2)"
      strSql = strSql & ",round(sum(DECODE(r102015,'1',nvl(r102020,0),0)),2),round(sum(DECODE(r102015,'1',nvl(r102011,0),0)),2),round(sum(DECODE(r102015,'2',nvl(r102011,0),0)),2),round(sum(DECODE(r102015,'4',nvl(r102011,0),0)),2),round(sum(DECODE(r102015,'1',nvl(r102021,0),0)),2),round(sum(NVL(r102012,0)),2),round(sum(NVL(r102013,0)),2),round(sum(NVL(r102014,0)),2)"
      strSql = strSql & ",r102002, ST06,nvl(st02,r102002) from r090608,staff WHERE ID='" & strUserNum & "' and r102002=st01(+) AND R102001='" & Combo1.Text & "' group by r102001,r102002,nvl(st02,r102002),ST03, ST06 "
   'Added by Morgan 2025/7/18
   ElseIf frm090608.txt1(12) = "3" Then
      strSql = "select ST03,round(sum(NVL(r102003,0)),2),round(sum(NVL(r102004,0)),2)" & _
         ",round(sum(NVL(r102010,0)),2),round(sum(NVL(r102012,0)),2),round(sum(NVL(r102011,0)),2),round(sum(NVL(r102013,0)),2)" & _
         ",round(sum(NVL(r102024,0)),2),round(sum(NVL(r102029,0)),2),round(sum(NVL(r102025,0)),2),round(sum(NVL(r102030,0)),2)" & _
         ",round(sum(NVL(r102005,0)),2),round(sum(NVL(r102007,0)),2),round(sum(NVL(r102006,0)),2),round(sum(NVL(r102008,0)),2)" & _
         ",r102002, ST06,nvl(st02,r102002) from r090608,staff" & _
         " WHERE ID='" & strUserNum & "' and r102002=st01(+) AND R102001='" & Combo1.Text & "' group by r102001,r102002,nvl(st02,r102002),ST03, ST06 "
   'end 2025/7/18
   Else
      strSql = "select ST03,round(sum(NVL(r102003,0)),2),round(sum(NVL(r102004,0)),2),round(sum(NVL(r102005,0)),2),round(sum(NVL(r102022,0)),2),round(sum(NVL(r102006,0)),2),round(sum(NVL(r102007,0)),2),round(sum(NVL(r102023,0)),2),round(sum(NVL(r102008,0)),2),round(sum(NVL(r102009,0)),2),round(sum(NVL(r102010,0)),2),round(sum(NVL(r102011,0)),2),round(sum(NVL(r102012,0)),2),round(sum(NVL(r102013,0)),2),round(sum(NVL(r102014,0)),2),r102002, ST06,nvl(st02,r102002) from r090608,staff WHERE ID='" & strUserNum & "' and r102002=st01(+) AND R102001='" & Combo1.Text & "' group by r102001,r102002,nvl(st02,r102002),ST03, ST06 "
   End If
   'end 2018/11/15
   
   '逼埃0戈
   strSql = strSql & " Having (round(sum(DECODE(r102003,0,0,NULL,0,r102003)),2)+round(sum(DECODE(r102004,0,0,NULL,0,r102004)),2)+round(sum(DECODE(r102005,0,0,NULL,0,r102005)),2)+round(sum(DECODE(r102006,0,0,NULL,0,r102006)),2)+round(sum(DECODE(r102007,0,0,NULL,0,r102007)),2)+round(sum(DECODE(r102008,0,0,NULL,0,r102008)),2)+round(sum(DECODE(r102009,0,0,NULL,0,r102009)),2)+round(sum(DECODE(r102010,0,0,NULL,0,r102010)),2)+round(sum(DECODE(r102011,0,0,NULL,0,r102011)),2)+round(sum(DECODE(r102012,0,0,NULL,0,r102012)),2)+round(sum(DECODE(r102013,0,0,NULL,0,r102013)),2)+round(sum(DECODE(r102014,0,0,NULL,0,r102014)),2)) > 0 "
   Select Case Val(frm090608.txt1(9))
   Case 1
        pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "1.祇ゅ翴计 %" 'Add By Sindy 2010/12/14
        strSql = strSql + " ORDER BY R102001,sum(R102007) DESC "
   Case 2
        pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "2.祇ゅン计 %" 'Add By Sindy 2010/12/14
        strSql = strSql + " ORDER BY R102001,sum(R102008) DESC "
   Case 3
        pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "3.祇ゅキА %" 'Add By Sindy 2010/12/14
        strSql = strSql + " ORDER BY R102001,sum(R102009) DESC "
   Case 4
        pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "4.┯快" 'Add By Sindy 2010/12/14
        strSql = strSql + " ORDER BY R102001, ST06, ST03, R102002 "
   'Added by Lydia 2016/12/19 +Ч絑膀计%
   Case 5
        pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "5.Ч絑膀计 %"
        strSql = strSql + " ORDER BY R102001,sum(R102013) DESC "
   Case Else
   End Select
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         '2012/1/20 MDOFI BY SONIA COMBO1ち传琩高ぃゲ
         'InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/14
         If m_INSERTLOG Then InsertQueryLog (.RecordCount)
         '2012/1/20 END
         Set grd1.Recordset = adoRecordset
         For ii = 2 To grd1.Rows - 1
            'Modified by Morgan 2018/11/15
            'grd1.TextMatrix(ii, 0) = grd1.TextMatrix(ii, 16 )
            'Modified by Morgan 2019/3/22
            'grd1.TextMatrix(ii, 0) = grd1.TextMatrix(ii, 16 + IIf(Check1.Value = vbChecked, 12, 0))
            'Modified by Morgan 2025/7/18
            'grd1.TextMatrix(ii, 0) = grd1.TextMatrix(ii, 18 + IIf(Check1.Value = vbChecked, 12, 0))
            grd1.TextMatrix(ii, 0) = grd1.TextMatrix(ii, grd1.Cols - 1)
            'end 2018/11/15
         Next ii
      Else
         '2012/1/20 MDOFI BY SONIA COMBO1ち传琩高ぃゲ,礚戈璶睲GRID
         'InsertQueryLog (0) 'Add By Sindy 2010/12/14
         If m_INSERTLOG Then InsertQueryLog (0)
         grd1.Rows = 3
         grd1.Clear
         '2012/1/20 END
      End If
   End With
   SetGrd1
   CheckOC
   
   'Added by Morgan 2019/4/11
   '翴计陪ボ计1膀计2
   For ii = grd1.FixedRows To grd1.Rows - 1
      For jj = 2 To grd1.Cols - 1
         If Right(grd1.TextMatrix(1, jj), 2) <> "ン计" Then
            If Right(grd1.TextMatrix(1, jj), 2) = "翴计" And Right(grd1.TextMatrix(0, jj), 1) <> "%" Then
               grd1.TextMatrix(ii, jj) = Format(grd1.TextMatrix(ii, jj), "#0.0")
            Else
               grd1.TextMatrix(ii, jj) = Format(grd1.TextMatrix(ii, jj), "#0.00")
            End If
         End If
      Next
   Next ii
   'end 2019/4/11
End Sub
'Added by Morgan 2018/11/15
Private Sub SetGrd2()
   Dim ii As Integer
   
   'Modified by Morgan 2019/3/22 +祇ゅ龟罿翴计,祇ゅ龟罿翴计笷Θ瞯
   With grd1
       .Visible = False
       .Cols = 29
       .row = 0: .col = 0: .Text = "┯快"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 0: .Text = "┯快"
       .ColWidth(0) = 800
       .CellAlignment = flexAlignCenterCenter
      
       .row = 0: .col = 1: .Text = "场腹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 1: .Text = "场腹"
       .ColWidth(1) = 800
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 2: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 2: .Text = "翴计"
       .ColWidth(2) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 3: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 3: .Text = "膀计"
       .ColWidth(3) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 4: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 4: .Text = "翴计"
       .ColWidth(4) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 5: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 5: .Text = "だ皌翴计"
       .ColWidth(5) = 800
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 6: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 6: .Text = "龟罿翴计"
       If m_bol108Rule Then
         .ColWidth(6) = 800
       Else
         .ColWidth(6) = 0
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 7: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 7: .Text = "膀计"
       .ColWidth(7) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 8: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 8: .Text = "﹍"
       .ColWidth(8) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 9: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 9: .Text = ""
       .ColWidth(9) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 10: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 10: .Text = "や穿"
       .ColWidth(10) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 11: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 11: .Text = "Μゅ"
       If m_bol108Rule Then
         .ColWidth(11) = 0
       Else
         .ColWidth(11) = 700
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 12: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 12: .Text = "ン计"
       .ColWidth(12) = 500
       .CellAlignment = flexAlignCenterCenter
              
       .row = 0: .col = 13: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 13: .Text = "翴计"
       .ColWidth(13) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 14: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 14: .Text = "龟罿翴计"
       If m_bol108Rule Then
         .ColWidth(14) = 900
       Else
         .ColWidth(14) = 0
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 15: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 15: .Text = "膀计"
       .ColWidth(15) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 16: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 16: .Text = "キА"
       .ColWidth(16) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 17: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 17: .Text = "翴计"
       .ColWidth(17) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 18: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 18: .Text = "だ皌"
       .ColWidth(18) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 19: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 19: .Text = "膀计"
       .ColWidth(19) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 20: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 20: .Text = "﹍"
       .ColWidth(20) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 21: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 21: .Text = ""
       .ColWidth(21) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 22: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 22: .Text = "や穿"
       .ColWidth(22) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 23: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 23: .Text = "Μゅ"
       If m_bol108Rule Then
         .ColWidth(23) = 0
       Else
         .ColWidth(23) = 700
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 24: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 24: .Text = "ン计"
       .ColWidth(24) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 25: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 25: .Text = "翴计"
       .ColWidth(25) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 26: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 26: .Text = "膀计"
       .ColWidth(26) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 27: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 27: .Text = "キА"
       .ColWidth(27) = 700
       .CellAlignment = flexAlignCenterCenter
              
       .row = 0: .col = 28: .Text = ""
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 28: .Text = ""
       .ColWidth(28) = 0
       .CellAlignment = flexAlignCenterCenter
       For ii = 2 To Me.grd1.Cols - 1
         .ColAlignment(ii) = flexAlignRightCenter
       Next ii
          
       '夹肈ㄖ陪ボ
       .MergeCells = flexMergeRestrictRows
       .MergeRow(0) = True
       .MergeRow(1) = True
       .MergeRow(2) = False
       .MergeCol(0) = True
       .MergeCol(1) = True
       .MergeCol(2) = False
       .Visible = True
   End With
End Sub

Private Sub SetGrd1()
Dim ii As Integer
Dim strColName As String
   
   If frm090608.txt1(12) = "3" Then SetGrd3: Exit Sub 'Added by Morgan 2025/7/18
   
   If Check1.Value = vbChecked Then SetGrd2: Exit Sub 'Added by Morgan 2018/11/15
   
   'Modify by Morgan 2010/10/19
   'Modified by Morgan 2013/5/22 穝ゼ匡龟悔ン计ノ膀计
   If bolNewPromoterRule And frm090608.txt1(12) <> "2" And frm090608.Check1.Value = vbUnchecked Then
      'Modify by Morgan 2010/12/30 璸翴->膀计
      strColName = "膀计"
   Else
      strColName = "ン计"
   End If
   
   'Modified by Morgan 2019/3/22 +祇ゅ龟罿翴计,祇ゅ龟罿翴计笷Θ瞯
   '┯快场腹
   With grd1
       .Visible = False
       .Cols = 17
       .row = 0: .col = 0: .Text = "┯快"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 0: .Text = "┯快"
       .ColWidth(0) = 800
       .CellAlignment = flexAlignCenterCenter
      
       .row = 0: .col = 1: .Text = "场腹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 1: .Text = "场腹"
       .ColWidth(1) = 800
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 2: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 2: .Text = "翴计"
       .ColWidth(2) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 3: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 3: .Text = "膀计"
       .ColWidth(3) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 4: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 4: .Text = "翴计"
       .ColWidth(4) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 5: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 5: .Text = "龟膀翴计"
       If m_bol108Rule Then
         .ColWidth(5) = 900
       Else
         .ColWidth(5) = 0
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 6: .Text = "ヘ夹笷Θ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 6: .Text = strColName
       .ColWidth(6) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 7: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 7: .Text = "翴计"
       If m_bol108Rule Then
         .ColWidth(7) = 700
       Else
         .ColWidth(7) = 0
       End If
       .CellAlignment = flexAlignCenterCenter
              
       .row = 0: .col = 8: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 8: .Text = "龟罿翴计"
       If m_bol108Rule Then
         .ColWidth(8) = 900
       Else
         .ColWidth(8) = 0
       End If
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 9: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 9: .Text = strColName
       .ColWidth(9) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 10: .Text = "祇ゅ笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 10: .Text = "キА"
       .ColWidth(10) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 11: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 11: .Text = "翴计"
       .ColWidth(11) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 12: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 12: .Text = strColName
       .ColWidth(12) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 13: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 13: .Text = "翴计"
       .ColWidth(13) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 14: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 14: .Text = strColName
       .ColWidth(14) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 15: .Text = "Ч絑笷Θ瞯%"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 15: .Text = "キА"
       .ColWidth(15) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 16: .Text = ""
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 16: .Text = ""
       .ColWidth(16) = 0
       .CellAlignment = flexAlignCenterCenter
       For ii = 2 To Me.grd1.Cols - 1
         .ColAlignment(ii) = flexAlignRightCenter
       Next ii
          
       '夹肈ㄖ陪ボ
       .MergeCells = flexMergeRestrictRows
       .MergeRow(0) = True
       .MergeRow(1) = True
       .MergeRow(2) = False
       .MergeCol(0) = True
       .MergeCol(1) = True
       .MergeCol(2) = False
       .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090608_1 = Nothing
End Sub

'Added by Morgan 2025/7/18
Private Sub SetGrd3()
   Dim ii As Integer
   
   With grd1
       .Visible = False
       .WordWrap = True
       .RowHeight(1) = .RowHeight(0) * 2
       .Cols = 17
       .row = 0: .col = 0: .Text = "┯快"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 0: .Text = "┯快"
       .ColWidth(0) = 800
       .CellAlignment = flexAlignCenterCenter
       .row = 0: .col = 1: .Text = "场腹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 1: .Text = "场腹"
       .ColWidth(1) = 500
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 2: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 2: .Text = "翴计"
       .ColWidth(2) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 3: .Text = "ヘ夹"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 3: .Text = "膀计"
       .ColWidth(3) = 700
       .CellAlignment = flexAlignCenterCenter
       'Ч絑
       .row = 0: .col = 4: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 4: .Text = "翴计"
       .ColWidth(4) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 5: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 5: .Text = "翴计笷Θ瞯%"
      .ColWidth(5) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 6: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 6: .Text = "膀计"
       .ColWidth(6) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 7: .Text = "Ч絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 7: .Text = "膀计笷Θ瞯%"
       .ColWidth(7) = 700
       .CellAlignment = flexAlignCenterCenter
              
       '穦絑
       .row = 0: .col = 8: .Text = "穦絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 8: .Text = "翴计"
       .ColWidth(8) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 9: .Text = "穦絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 9: .Text = "翴计笷Θ瞯%"
      .ColWidth(9) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 10: .Text = "穦絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 10: .Text = "膀计"
       .ColWidth(10) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 11: .Text = "穦絑"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 11: .Text = "膀计笷Θ瞯%"
       .ColWidth(11) = 700
       .CellAlignment = flexAlignCenterCenter
       
       '祇ゅ
       .row = 0: .col = 12: .Text = "祇ゅ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 12: .Text = "翴计"
       .ColWidth(12) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 13: .Text = "祇ゅ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 13: .Text = "翴计笷Θ瞯%"
      .ColWidth(13) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 14: .Text = "祇ゅ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 14: .Text = "膀计"
       .ColWidth(14) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 15: .Text = "祇ゅ"
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 15: .Text = "膀计笷Θ瞯%"
       .ColWidth(15) = 700
       .CellAlignment = flexAlignCenterCenter
       
       .row = 0: .col = 16: .Text = ""
       .CellAlignment = flexAlignCenterCenter
       .row = 1: .col = 16: .Text = ""
       .ColWidth(16) = 0
       .CellAlignment = flexAlignCenterCenter
       For ii = 2 To Me.grd1.Cols - 1
         .ColAlignment(ii) = flexAlignRightCenter
       Next ii
          
       '夹肈ㄖ陪ボ
       .MergeCells = flexMergeRestrictRows
       .MergeRow(0) = True
       .MergeRow(1) = True
       .MergeRow(2) = False
       .MergeCol(0) = True
       .MergeCol(1) = True
       .MergeCol(2) = False
       .Visible = True
   End With
End Sub
