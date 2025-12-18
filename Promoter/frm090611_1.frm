VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090611_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦天數統計查詢(明細)"
   ClientHeight    =   5925
   ClientLeft      =   -2910
   ClientTop       =   1770
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9315
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   2310
      Left            =   75
      TabIndex        =   3
      Top             =   3570
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4075
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7848
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2190
      Left            =   75
      TabIndex        =   1
      Top             =   1155
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   3863
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
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   870
      TabIndex        =   7
      Top             =   420
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2963;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   135
      TabIndex        =   6
      Top             =   900
      Width           =   9150
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "16140;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   180
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "承辦天數："
      Height          =   180
      Left            =   105
      TabIndex        =   4
      Top             =   3390
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   840
   End
End
Attribute VB_Name = "frm090611_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; grd1改字型=新細明體-ExtB、lbl1、Combo1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim Int1 As Integer, j As Integer

Private Sub cmdOK_Click()
Me.Hide
frm090611.Show
Unload Me
End Sub

Private Sub Combo1_Click()
Process
SetGrd1
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
StrMenu
Process
SetGrd1
'Add By Cheng 2003/06/10
If frm090611.Option1(0).Value Then
    Me.Label3.Caption = "會稿日期：" & frm090611.txt1(3).Text & "－" & frm090611.txt1(4).Text
Else
    Me.Label3.Caption = "完稿日期：" & frm090611.txt1(24).Text & "－" & frm090611.txt1(25).Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090611_1 = Nothing
End Sub

Sub StrMenu()
'Modify By Cheng 2003/06/10
'strSQL = "SELECT DISTINCT R106001 FROM R090611_1 WHERE ID='" & strUserNum & "' "
'若列印對象為承辦人
If frm090611.txt1(7).Text = "1" Then
    pub_QL05 = pub_QL05 & ";" & frm090611.Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/17
    strSql = "SELECT DISTINCT R106001, ST02, ST06, ST03 FROM R090611_1, Staff WHERE R106001=ST01(+) And ID='" & strUserNum & "' Order By ST06, ST03, R106001 "
'若列印對象為智權人員
Else
    pub_QL05 = pub_QL05 & ";" & frm090611.Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/17
    strSql = "SELECT DISTINCT R106001, ST02, ST06, ST15 FROM R090611_1, Staff WHERE R106001=ST01(+) And ID='" & strUserNum & "' Order By ST06, ST15, R106001 "
End If
CheckOC
j = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            'Modify By Cheng 2003/06/10
'            Combo1.AddItem CheckStr(.Fields(0)), j
            Combo1.AddItem CheckStr(.Fields(0)) & " " & .Fields(1).Value, j
            j = j + 1
            .MoveNext
        Loop
    End If
End With
CheckOC
Combo1.Text = Combo1.List(0)
End Sub

Sub Process()
'Modify By Cheng 2003/06/10
'strSQL = " SELECT r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090611_1 WHERE ID='" & strUserNum & "' AND R106001='" & Combo1.Text & "' "
strSql = " SELECT r106002,r106003,r106004,r106005,r106006,ST02,r106008,r106009,r106010,r106011 FROM R090611_1,Staff WHERE R106007=ST01(+) And ID='" & strUserNum & "' AND R106001='" & Trim(Left(Combo1.Text, 6)) & "' "
'Add By Cheng 2003/06/05
strSql = strSql & " Order By R106001, R106002, R106003 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    pub_QL05 = pub_QL05 & ";明細" 'Add By Sindy 2010/12/17
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set grd1.Recordset = adoRecordset
        Int1 = .RecordCount
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
'Modify By Cheng 2003/06/10
'strSQL = "select sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' and r108001='" & Combo1.Text & "' "
strSql = "select sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' and r108001='" & Trim(Left(Combo1.Text, 6)) & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    pub_QL05 = pub_QL05 & ";統計" 'Add By Sindy 2010/12/17
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set grd2.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 10
    .row = 0
    .col = 0:   .Text = "天數"
    .ColWidth(0) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "本所案號"
    .ColWidth(1) = 1550
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "案件名稱"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "種類"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "案件性質"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5
    If Val(frm090611.txt1(7)) = 1 Then
        .Text = "智權人員"
        Label1.Caption = "承辦人："
        'Modify By Cheng 2003/06/10
'        lbl1.Caption = "承辦人：" & Combo1.Text & ", 收文共 " & str(Int1) & " 件 "
        lbl1.Caption = "承辦人：" & Trim(Mid(Combo1.Text, 7, Len(Me.Combo1.Text) - 6)) & ", 收文共 " & str(Int1) & " 件 "
    Else
        .Text = "承辦人"
        Label1.Caption = "智權人員："
        'Modify By Cheng 2003/06/10
'        lbl1.Caption = "智權人員：" & Combo1.Text & ", 收文共 " & str(Int1) & " 件 "
        lbl1.Caption = "智權人員：" & Trim(Mid(Combo1.Text, 7, Len(Me.Combo1.Text) - 6)) & ", 收文共 " & str(Int1) & " 件 "
    End If
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "齊備日"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "完稿日"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = "會稿日"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9: .Text = "發文日"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
End With
With grd2
    .Cols = 5
    .row = 0
    .col = 0:   .Text = frm090611.txt1(12) & "-" & frm090611.txt1(13)
    .ColWidth(0) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = frm090611.txt1(14) & "-" & frm090611.txt1(15)
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = frm090611.txt1(16) & "-" & frm090611.txt1(17)
    .ColWidth(2) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = frm090611.txt1(18) & "-" & frm090611.txt1(19)
    .ColWidth(3) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = frm090611.txt1(20) & "-" & frm090611.txt1(21)
    .ColWidth(4) = 1500
    .CellAlignment = flexAlignCenterCenter
End With

End Sub


