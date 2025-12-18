VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090704_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員作業天數統計查詢(統計)"
   ClientHeight    =   8880
   ClientLeft      =   -2235
   ClientTop       =   2400
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   13290
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   8300
      Left            =   0
      TabIndex        =   1
      Top             =   468
      Width           =   14900
      _ExtentX        =   26273
      _ExtentY        =   14631
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
End
Attribute VB_Name = "frm090704_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim i As Integer
Private Sub cmdOK_Click()
Me.Hide
frm090704.Show
Unload Me
End Sub

Private Sub Form_Activate()
If Process = True Then
   SetGrd1
Else
   frm090704.Show
   Unload Me
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090704_2 = Nothing
End Sub

Function Process() As Boolean
Process = True
'Modify By Cheng 2003/07/17
'strSQL = "select r108001,R108002," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/(" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & ")," & SQLSum("R108012") & " from r090704_2 where id='" & strUserNum & "' group by r108001,r108002 order by r108001,r108002 "
strSql = "select ST02, R108002," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/(" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & ")," & SQLSum("R108012") & " ,r108001, ST06 from r090704_2, Staff where r108001=ST01(+) And  id='" & strUserNum & "' group by ST06, ST02, r108001,r108002 order by ST06, r108001, r108002 "
'strSQL = "select r108001," & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & "," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & " from r090611_2 where id='" & strUserNum & "' group by r108001 order by r108001 "
'strSQL = "select r108001,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' group by r108001 order by r108001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
        For i = 0 To grd1.Rows - 1
            grd1.row = i
            grd1.col = 1
            If Val(grd1.Text) = 1 Then
                grd1.Text = "草圖"
            Else
                grd1.Text = "墨圖"
            End If
        Next i
    Else
        Process = False
        ShowNoData
        Exit Function
    End If
End With
CheckOC
End Function

Private Sub SetGrd1()
With grd1
    .Cols = 12
    .row = 0
    .col = 0: .Text = "承辦人"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "當月承辦"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = frm090704.Txt1(8) & "-" & frm090704.Txt1(9)
    .ColWidth(2) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = frm090704.Txt1(10) & "-" & frm090704.Txt1(11)
    .ColWidth(3) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = frm090704.Txt1(12) & "-" & frm090704.Txt1(13)
    .ColWidth(4) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = frm090704.Txt1(14) & "-" & frm090704.Txt1(15)
    .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = frm090704.Txt1(16) & "-" & frm090704.Txt1(17)
    .ColWidth(6) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = frm090704.Txt1(18) & "-" & frm090704.Txt1(19)
    .ColWidth(7) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = frm090704.Txt1(20) & "-" & frm090704.Txt1(21)
    .ColWidth(8) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 9: .Text = frm090704.Txt1(22) & "-" & frm090704.Txt1(23)
    .ColWidth(9) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 10: .Text = "平均天數"
    .ColWidth(10) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 11: .Text = "超過件數"
    .ColWidth(11) = 1000
    .CellAlignment = flexAlignCenterCenter
End With

End Sub


