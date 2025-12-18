VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090611_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦天數統計查詢(統計)"
   ClientHeight    =   5715
   ClientLeft      =   -3030
   ClientTop       =   2475
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   8040
      TabIndex        =   1
      Top             =   24
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5232
      Left            =   36
      TabIndex        =   0
      Top             =   468
      Width           =   9252
      _ExtentX        =   16325
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
   Begin VB.Label Label3 
      Height          =   180
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   4455
   End
End
Attribute VB_Name = "frm090611_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Private Sub cmdOK_Click()
Me.Hide
frm090611.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Process
SetGrd1
'Add By Cheng 2003/06/10
If frm090611.Option1(0).Value Then
    Me.Label3.Caption = "會稿日期：" & frm090611.Txt1(3).Text & "－" & frm090611.Txt1(4).Text
Else
    Me.Label3.Caption = "完稿日期：" & frm090611.Txt1(24).Text & "－" & frm090611.Txt1(25).Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090611_2 = Nothing
End Sub

Sub Process()
Dim ii As Integer
Dim jj As Integer
Dim dblTotal As Double

'strSQL = "select r108001,sum(r108003)+sum(r108004)+sum(r108005)+sum(r108006)+sum(r108007),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' group by r108001 order by r108001 "
'Modify By Cheng 2003/06/10
'列印對象為承辦人
If frm090611.Txt1(7).Text = "1" Then
    pub_QL05 = pub_QL05 & ";" & frm090611.Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/17
    strSql = "select ST02,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007), ST06, ST03, R108001 from r090611_2, Staff where R108001=ST01(+) And id='" & strUserNum & "' group by r108001, ST06, ST03, ST02 order by ST06, ST03, r108001 "
'列印對象為智權人員
Else
    pub_QL05 = pub_QL05 & ";" & frm090611.Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/17
    strSql = "select ST02,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007), ST06, ST15, R108001 from r090611_2, Staff where R108001=ST01(+) And id='" & strUserNum & "' group by r108001, ST06, ST15, ST02 order by ST06, ST15, r108001 "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set grd1.Recordset = adoRecordset
        Me.grd1.AddItem "合　計"
        For ii = 1 To 6
            dblTotal = 0
            For jj = 1 To Me.grd1.Rows - 2
                dblTotal = dblTotal + Val(Me.grd1.TextMatrix(jj, ii))
            Next jj
            Me.grd1.TextMatrix(Me.grd1.Rows - 1, ii) = dblTotal
        Next ii
        Me.grd1.AddItem "百分比"
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 1) = " "
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 2) = Format(Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 2)) / Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 1)) * 100, "##0.00") & "%"
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 3) = Format(Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 3)) / Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 1)) * 100, "##0.00") & "%"
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 4) = Format(Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 4)) / Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 1)) * 100, "##0.00") & "%"
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 5) = Format(Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 5)) / Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 1)) * 100, "##0.00") & "%"
        Me.grd1.TextMatrix(Me.grd1.Rows - 1, 6) = Format(Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 6)) / Val(Me.grd1.TextMatrix(Me.grd1.Rows - 2, 1)) * 100, "##0.00") & "%"
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 7
    .row = 0
    .col = 0: .Text = "承辦人"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    'Modify By Cheng 2003/06/10
    If frm090611.Option1(0).Value Then
        .col = 1: .Text = "當月會稿"
    Else
        .col = 1: .Text = "當月完稿"
    End If
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = frm090611.Txt1(12) & "-" & frm090611.Txt1(13)
    .ColWidth(2) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = frm090611.Txt1(14) & "-" & frm090611.Txt1(15)
    .ColWidth(3) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = frm090611.Txt1(16) & "-" & frm090611.Txt1(17)
    .ColWidth(4) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = frm090611.Txt1(18) & "-" & frm090611.Txt1(19)
    .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = frm090611.Txt1(20) & "-" & frm090611.Txt1(21)
    .ColWidth(6) = 1000
    .CellAlignment = flexAlignCenterCenter
End With

End Sub

