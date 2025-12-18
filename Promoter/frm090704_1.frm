VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090704_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員作業天數統計查詢(明細)"
   ClientHeight    =   8880
   ClientLeft      =   -3735
   ClientTop       =   2910
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
      Height          =   390
      Left            =   13260
      TabIndex        =   1
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   3700
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6535
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3700
      Left            =   0
      TabIndex        =   2
      Top             =   972
      Width           =   14900
      _ExtentX        =   26273
      _ExtentY        =   6535
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
      Left            =   1020
      TabIndex        =   3
      Top             =   390
      Width           =   2430
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4286;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Left            =   45
      TabIndex        =   6
      Top             =   457
      Width           =   990
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   48
      TabIndex        =   5
      Top             =   756
      Width           =   7200
   End
   Begin VB.Label Label2 
      Caption         =   "承辦天數："
      Height          =   180
      Left            =   45
      TabIndex        =   4
      Top             =   4800
      Width           =   7200
   End
End
Attribute VB_Name = "frm090704_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; Combo1、grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim Int1 As Integer, Int2 As Integer, Int3 As Integer, j As Integer

Private Sub cmdOK_Click()
Me.Hide
frm090704.Show
Unload Me
End Sub

'Modified by Lydia 2022/02/07 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Private Sub Combo1_Click()
Private Sub Combo1_DropButtonClick()
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2022/02/07

If Process = True Then
   SetGrd1
End If
End Sub

Private Sub Form_Activate()
StrMenu
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
Set frm090704_1 = Nothing
End Sub

Sub StrMenu()
'combo1 的資料
'Modify By Cheng 2003/07/17
'strSQL = "SELECT DISTINCT R106001 FROM R090704_1 WHERE ID='" & strUserNum & "' "
strSql = "SELECT DISTINCT ST06, R106001, ST02  FROM R090704_1, Staff WHERE R106001=ST01(+) And ID='" & strUserNum & "' Group By ST02, R106001, ST06 Order By ST06, R106001 "
CheckOC
j = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Combo1.AddItem CheckStr(.Fields(1)) & " " & .Fields(2).Value, j
            j = j + 1
            .MoveNext
        Loop
    End If
End With
CheckOC
Combo1.Text = Combo1.List(0)
End Sub

Function Process() As Boolean
Process = True
'讀資料並塞入grid
'Modify By Cheng 2003/07/17
'strSQL = " SELECT r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090704_1 WHERE ID='" & strUserNum & "' AND R106001='" & Combo1.Text & "' "
strSql = " SELECT r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090704_1 WHERE ID='" & strUserNum & "' AND R106001='" & Trim(Left(Combo1.Text, 6)) & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
        Int1 = .RecordCount
    Else
        ShowNoData
        Process = False
        Exit Function
    End If
End With
CheckOC
'Modify By Cheng 2003/07/17
'strSQL = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & Combo1.Text & "' and r106009 is not null "
strSql = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & Trim(Left(Combo1.Text, 6)) & "' and r106009 is not null "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Int2 = Val(CheckStr(.Fields(0)))
    End If
    'Modify By Cheng 2003/07/17
'    strSQL = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & Combo1.Text & "' and r106009 is not null "
    strSql = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & Trim(Left(Combo1.Text, 6)) & "' and r106009 is not null "
    CheckOC
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenDynamic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Int3 = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC
'Modify By Cheng 2003/07/17
'strSQL = "select DECODE(R108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & " from r090704_2 where id='" & strUserNum & "' and r108001='" & Combo1.Text & "' GROUP BY DECODE(R108002,1,'草圖','墨圖')"
strSql = "select DECODE(R108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & " from r090704_2 where id='" & strUserNum & "' and r108001='" & Trim(Left(Combo1.Text, 6)) & "' GROUP BY DECODE(R108002,1,'草圖','墨圖')"
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd2.Recordset = adoRecordset
    End If
End With
CheckOC
End Function

Private Sub SetGrd1()
'設定grid
With grd1
    .Cols = 10
    .row = 0
    .col = 0:   .Text = "作業天數"
    .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "本所案號"
    .ColWidth(1) = 1600
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "案件名稱"
    .ColWidth(2) = 2600
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "種類"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "案件性質"
    .ColWidth(4) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "承辦人"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "草圖齊備日"
    .ColWidth(6) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "草圖完稿日"
    .ColWidth(7) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = "墨圖完稿日"
    .ColWidth(8) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 9: .Text = "發文日"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
End With
With grd2
    .Cols = 9
    .row = 0
    .col = 0:   .Text = ""
    .ColWidth(0) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = frm090704.Txt1(8) & "-" & frm090704.Txt1(9)
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = frm090704.Txt1(10) & "-" & frm090704.Txt1(11)
    .ColWidth(2) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = frm090704.Txt1(12) & "-" & frm090704.Txt1(13)
    .ColWidth(3) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = frm090704.Txt1(14) & "-" & frm090704.Txt1(15)
    .ColWidth(4) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = frm090704.Txt1(16) & "-" & frm090704.Txt1(17)
    .ColWidth(5) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = frm090704.Txt1(18) & "-" & frm090704.Txt1(19)
    .ColWidth(6) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = frm090704.Txt1(20) & "-" & frm090704.Txt1(21)
    .ColWidth(7) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = frm090704.Txt1(22) & "-" & frm090704.Txt1(23)
    .ColWidth(8) = 1500
    .CellAlignment = flexAlignCenterCenter
End With

End Sub


