VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090609_1 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "┯快秖琩高"
   ClientHeight    =   6228
   ClientLeft      =   216
   ClientTop       =   900
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6228
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdok 
      Caption         =   "穝(&R)"
      Height          =   400
      Index           =   2
      Left            =   5304
      TabIndex        =   8
      Top             =   72
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "玡礶(&U)"
      Height          =   400
      Index           =   1
      Left            =   8040
      TabIndex        =   3
      Top             =   72
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "筄セ┮戳(&T)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6720
      TabIndex        =   2
      Top             =   72
      Width           =   1300
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4596
      Left            =   0
      TabIndex        =   0
      Top             =   636
      Width           =   9276
      _ExtentX        =   16362
      _ExtentY        =   8107
      _Version        =   393216
      Rows            =   3
      Cols            =   1
      FixedRows       =   2
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "3.禬筁┮度参璸禬筁┮6るン计"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Top             =   6000
      Width           =   3456
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ネて: 0.67)  ㄤ玥璸ンン计"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4968
      TabIndex        =   7
      Top             =   5520
      Width           =   3072
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2.や穿だ计ぃσ納瓣產の╰参兵ン常璸獶砞璸┯快秖"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3840
      TabIndex        =   6
      Top             =   5736
      Width           =   4272
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "称爹: 1.珹┓ず计: 盡獶砞璸>=疭﹚σン计(诀篶: 1 筿: 0.83"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3384
      TabIndex        =   5
      Top             =   5304
      Width           =   5880
   End
   Begin VB.Label lbl2 
      Height          =   180
      Left            =   72
      TabIndex        =   4
      Top             =   5292
      Width           =   3048
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   396
      Width           =   1776
   End
End
Attribute VB_Name = "frm090609_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 эΘForm2.0 ; grd1э=穝灿砰-ExtB
'Memo By Morgan 2012/12/10 醇舦逆э
'2010/12/1 memo by sonia 絪腹逆э
'Memo by Morgan2010/8/16 ら戳逆э
Option Explicit

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Me.Hide
     frm090609_2.Show
Case 1
     Me.Hide
     If frm090609.ObjForm = 1 Then
        frm090609.Show
        Unload Me
        Exit Sub
     Else
        frm090609_2.Show
        Unload Me
        Exit Sub
     End If
     
'Added by Morgan 2024/4/19
Case 2
   If frm090609.ObjForm = 2 Then
      Unload frm090609_2
   End If
   frm090609.Show
   Me.Hide
   frm090609.m_bolRedo = True
   frm090609.cmdOK(0).Value = True
   frm090609.m_bolRedo = False
   ReadData
   
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Modified by Morgan 2024/4/19
'SetGrd1
''Modify by Morgan 2011/3/31
''lbl1.Caption = Mid(GetTaiwanTodayDate, 1, 2) & "  " & Mid(GetTaiwanTodayDate, 3, 2) & " る "
'lbl1.Caption = (Val(strSrvDate(2)) \ 10000) & "  " & Right(strSrvDate(2) \ 100, 2) & " る "
'StrMenu
''lbl2.Caption = "璸" & str(GRD1.Rows - 1)
''Modified by Lydia 2015/01/22 癘魁掸计
'lbl2.Caption = "璸" & str(grd1.Rows - 2)
''Process
ReadData
'end 2024/4/19
End Sub

Sub StrMenu()
Dim i As Integer
'Added by Lydia 2017/12/28 盢ゲ璶兵ン既郎
Dim rsAD As New ADODB.Recordset
Dim mSeq As String '既TB絪腹
'end 2017/12/28
    
'Added by Lydia 2017/12/28 盢ゲ璶兵ン既郎
strSql = "select eep01,max(eep02) eep02 from caseprogress,EmpElectronProcess where cp158=0 and cp159=0 and cp09=eep01(+) and eep04 in('" & EMP_癳穦 & "','" & EMP_穦 & "') group by eep01 "
i = 1
Set RsTemp = ClsLawReadRstMsg(i, strSql)
If i = 1 Then
    Set rsAD = PUB_CreateRecordset(RsTemp, , , , Me.Name, mSeq)
End If
'end 2017/12/28

'Modify by Morgan 2010/6/24 + 陪ボ穝
'strSql = "select NVL(ST02,r103001),sum(r103002),sum(r103003),sum(r103004),sum(r103005),sum(r103006),sum(r103007),sum(r103008),sum(r103009),sum(r103010),sum(r103011),sum(r103012) from r090609_1,STAFF where id='" & strUserNum & "' AND r103001=ST01(+) group by ST06,ST03,r103001,NVL(ST02,r103001) order by ST06,ST03,r103001,NVL(ST02,r103001) "
'Modify by Sindy 2013/10/22
'strSql = "select NVL(ST02,r103001),sum(r103002),sum(r103013)||'('||sum(r103003)||')',sum(r103014)||'('||sum(r103004)||')',sum(r103015)||'('||sum(r103005)||')',sum(r103016)||'('||sum(r103006)||')',sum(r103017)||'('||sum(r103007)||')',sum(r103018)||'('||sum(r103008)||')',sum(r103019),sum(r103020),sum(r103011),sum(r103012) from r090609_1,STAFF where id='" & strUserNum & "' AND r103001=ST01(+) group by ST06,ST03,r103001,NVL(ST02,r103001) order by ST06,ST03,r103001,NVL(ST02,r103001) "
'Modified by Lydia 2024/08/02 玂痙Τ/礚;秸俱逆抖---琭揩
'strSql = "select NVL(ST02,r103001),ST01,sum(r103002),sum(r103013)||'('||sum(r103003)||')',sum(r103014)||'('||sum(r103004)||')',sum(r103015)||'('||sum(r103005)||')',sum(r103016)||'('||sum(r103006)||')',' ',sum(r103017)||'('||sum(r103007)||')',sum(r103018)||'('||sum(r103008)||')',sum(r103019),sum(r103020),sum(r103011),sum(r103012) from r090609_1,STAFF where id='" & strUserNum & "' AND r103001=ST01(+) group by ST06,ST03,ST01,r103001,NVL(ST02,r103001) order by ST06,ST03,ST01,r103001,NVL(ST02,r103001) "
'strSql = "select nvl(st02,r103001) as ┯快,st01 as ID,sum(r103002) as 禬筁猭﹚,sum(r103013)||'('||sum(r103003)||')' as 砞璸┯快秖,sum(r103014)||'('||sum(r103004)||')' as 獶砞璸┯快秖,sum(r103015)||'('||sum(r103005)||')' as 砞璸快秖,sum(r103016)||'('||sum(r103006)||')' as 獶砞璸快秖,' ' as 穦快秖,sum(r103017)||'('||sum(r103007)||')' as 砞璸だ秖,sum(r103018)||'('||sum(r103008)||')' as 獶砞璸だ秖,sum(r103019) as 砞璸祇ゅン计 ,sum(r103020) as 獶砞璸祇ゅン计,sum(r103011) as 砞璸祇ゅ翴计,sum(r103012) as 獶砞璸祇ゅ翴计 from r090609_1,staff where id='" & strUserNum & "' AND r103001=ST01(+) group by ST06,ST03,ST01,r103001,NVL(ST02,r103001) order by ST06,ST03,ST01,r103001,NVL(ST02,r103001) "
'1.快秖簿程オ娩,
'2.獶砞璸эΘ砞璸オ娩 (┮Τ逆常э硂妓ぃ穦睹)
'ヘ琌琵快秖獶砞璸候綟祘畍﹎
'Modified by Morgan 2025/2/18 逼埃戮祘畍絪腹--琭揩
'strSql = "select nvl(st02,r103001) as ┯快,st01 as id,sum(r103016)||'('||sum(r103006)||')' as 獶砞璸快秖,sum(r103015)||'('||sum(r103005)||')' as 砞璸快秖,' ' as 穦快秖,sum(r103002) as 禬筁猭﹚,sum(r103014)||'('||sum(r103004)||')' as 獶砞璸┯快秖,sum(r103013)||'('||sum(r103003)||')' as 砞璸┯快秖,sum(r103018)||'('||sum(r103008)||')' as 獶砞璸だ秖,sum(r103017)||'('||sum(r103007)||')' as 砞璸だ秖 ,sum(r103020) as 獶砞璸祇ゅン计,sum(r103019) as 砞璸祇ゅン计,sum(r103012) as 獶砞璸祇ゅ翴计,sum(r103011) as 砞璸祇ゅ翴计 " & _
         "from r090609_1,staff where id='" & strUserNum & "' AND r103001=ST01(+) group by ST06,ST03,ST01,r103001,NVL(ST02,r103001) order by ST06,ST03,ST01,r103001,NVL(ST02,r103001) "
strSql = "select nvl(st02,r103001) as ┯快,st01 as id,sum(r103016)||'('||sum(r103006)||')' as 獶砞璸快秖,sum(r103015)||'('||sum(r103005)||')' as 砞璸快秖,' ' as 穦快秖,sum(r103002) as 禬筁猭﹚,sum(r103014)||'('||sum(r103004)||')' as 獶砞璸┯快秖,sum(r103013)||'('||sum(r103003)||')' as 砞璸┯快秖,sum(r103018)||'('||sum(r103008)||')' as 獶砞璸だ秖,sum(r103017)||'('||sum(r103007)||')' as 砞璸だ秖 ,sum(r103020) as 獶砞璸祇ゅン计,sum(r103019) as 砞璸祇ゅン计,sum(r103012) as 獶砞璸祇ゅ翴计,sum(r103011) as 砞璸祇ゅ翴计 " & _
         "from r090609_1,staff where id='" & strUserNum & "' and r103001=ST01(+) and substr(st01,4,1)<'9'" & _
         " group by ST06,ST03,ST01,r103001,NVL(ST02,r103001) order by ST06,ST03,ST01,r103001 "
'end 2024/08/02

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set GRD1.Recordset = adoRecordset
        'Add By Sindy 2013/10/23
        For i = 2 To GRD1.Rows - 1
            'Modify By Sindy 2017/2/8 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
            'Modified by Lydia 2017/12/28 盢ゲ璶兵ン既郎
'            strSql = "select e2.eep01,e2.eep02,e2.eep04,e2.eep05 from" & _
'                     " (select eep01,max(eep02) eep02" & _
'                     " from EmpElectronProcess,caseprogress where eep01=cp09(+) and CP158=0 and CP159=0 and eep04 in('" & EMP_癳穦 & "','" & EMP_穦 & "')" & _
'                     " group by eep01) e1,EmpElectronProcess e2" & _
'                     " where e1.eep01=e2.eep01(+) and e1.eep02=e2.eep02(+)" & _
'                     " and e2.eep04='" & EMP_穦 & "'" & _
'                     " and e2.eep05='" & GRD1.TextMatrix(i, 1) & "'"
            strSql = "select e2.eep01,e2.eep02,e2.eep04,e2.eep05 from rdatafactory e1,EmpElectronProcess e2" & _
                     " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mSeq & "' " & _
                     " and e1.r001=e2.eep01(+) and e1.r002=e2.eep02(+)" & _
                     " and e2.eep04='" & EMP_穦 & "'" & _
                     " and e2.eep05='" & GRD1.TextMatrix(i, 1) & "'"
            'end 2017/12/28
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               'Modified by Lydia 2024/08/02 э竚7=>4
               GRD1.TextMatrix(i, 4) = RsTemp.RecordCount
            Else
               'Modified by Lydia 2024/08/02 э竚7=>4
               GRD1.TextMatrix(i, 4) = 0
            End If
        Next i
        '2013/10/23 END
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
SetGrd1
Me.GRD1.row = 2: Me.GRD1.col = 0
Set rsAD = Nothing 'Added by Lydia 2017/12/28
End Sub

'Moddified by Lydia 2024/08/02 SetGrd1эSetGrd1_old
Private Sub SetGrd1_old()
With GRD1
    .Cols = 14
    'Add By Cheng 2003/06/10
    .row = 0
    .col = 0:   .Text = "┯快"
    .ColWidth(0) = 670
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "ID"
    .ColWidth(1) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "禬筁猭﹚"
    .ColWidth(2) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "┯快秖"
    .ColWidth(3) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "┯快秖"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "快秖"
    .ColWidth(5) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "快秖"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "快秖"
    .ColWidth(7) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "だ秖"
    .ColWidth(8) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "だ秖"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "祇ゅン计"
    .ColWidth(10) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "祇ゅン计"
    .ColWidth(11) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "祇ゅ翴计"
    .ColWidth(12) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "祇ゅ翴计"
    .ColWidth(13) = 800
    .CellAlignment = flexAlignCenterCenter
    
    .row = 1
    .col = 0:   .Text = "┯快"
    .ColWidth(0) = 670
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "ID"
    .ColWidth(1) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "禬筁猭﹚"
    .ColWidth(2) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "砞璸"
    .ColWidth(3) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "獶砞璸"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "砞璸"
    .ColWidth(5) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "獶砞璸"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "穦"
    .ColWidth(7) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "砞璸"
    .ColWidth(8) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "獶砞璸"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "砞璸"
    .ColWidth(10) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "獶砞璸"
    .ColWidth(11) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "砞璸"
    .ColWidth(12) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "獶砞璸"
    .ColWidth(13) = 800
    .CellAlignment = flexAlignCenterCenter
    .MergeCells = flexMergeRestrictRows
    .MergeRow(0) = True: .MergeRow(1) = True: .MergeRow(2) = False
    .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True
End With
End Sub

'Added by Lydia 2024/08/02 秸俱逆抖
Private Sub SetGrd1()
'1.快秖簿程オ娩,
'2.獶砞璸эΘ砞璸オ娩 (┮Τ逆常э硂妓ぃ穦睹)
'ヘ琌琵快秖獶砞璸候綟祘畍﹎

With GRD1
    .Cols = 14
    .row = 0
    .col = 0:   .Text = "┯快"
    .ColWidth(0) = 670
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "ID"
    .ColWidth(1) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "快秖"
    .ColWidth(2) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "快秖"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "快秖"
    .ColWidth(4) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "禬筁" '禬筁┮6る
    .ColWidth(5) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "┯快秖"
    .ColWidth(6) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "┯快秖"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "だ秖"
    .ColWidth(8) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "だ秖"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "祇ゅン计"
    .ColWidth(10) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "祇ゅン计"
    .ColWidth(11) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "祇ゅ翴计"
    .ColWidth(12) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "祇ゅ翴计"
    .ColWidth(13) = 800
    .CellAlignment = flexAlignCenterCenter

    
    .row = 1
    .col = 0:   .Text = "┯快"
    .ColWidth(0) = 670
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "ID"
    .ColWidth(1) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "獶砞璸"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3:  .Text = "砞璸"
    .ColWidth(3) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 4:  .Text = "穦"
    .ColWidth(4) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "┮" '禬筁┮6る
    .ColWidth(5) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "獶砞璸"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "砞璸"
    .ColWidth(7) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "獶砞璸"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "砞璸"
    .ColWidth(9) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "獶砞璸"
    .ColWidth(10) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 11:   .Text = "砞璸"
    .ColWidth(11) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 12:   .Text = "獶砞璸"
    .ColWidth(12) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 13:   .Text = "砞璸"
    .ColWidth(13) = 700
    .CellAlignment = flexAlignCenterCenter

    .MergeCells = flexMergeRestrictRows
    .MergeRow(0) = True: .MergeRow(1) = True: .MergeRow(2) = False
    .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090609_1 = Nothing
End Sub

'Added by Morgan 2024/4/19
Private Sub ReadData()
   Me.Visible = False
   SetGrd1
   LBL1.Caption = (Val(strSrvDate(2)) \ 10000) & "  " & Right(strSrvDate(2) \ 100, 2) & " る "
   StrMenu
   lbl2.Caption = "璸" & str(GRD1.Rows - 2)
   Me.Visible = True
End Sub
