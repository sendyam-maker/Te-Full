VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090615 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人、繪圖人員目標資料維護"
   ClientHeight    =   5736
   ClientLeft      =   540
   ClientTop       =   3840
   ClientWidth     =   9300
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9300
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1056
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3380
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1056
      Width           =   1100
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1085
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1056
      Width           =   1100
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "複製資料(&C)"
      Height          =   400
      Index           =   1
      Left            =   8040
      TabIndex        =   13
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   400
      Index           =   0
      Left            =   7224
      TabIndex        =   12
      Top             =   600
      Width           =   756
   End
   Begin VB.Frame Frame1 
      Height          =   4356
      Left            =   0
      TabIndex        =   18
      Top             =   1368
      Width           =   9252
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   2628
         Left            =   72
         TabIndex        =   29
         Top             =   1692
         Width           =   9132
         _ExtentX        =   16108
         _ExtentY        =   4636
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   9
         Left            =   5595
         TabIndex        =   9
         Top             =   1368
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   8
         Left            =   3330
         TabIndex        =   8
         Top             =   1368
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   7
         Left            =   1035
         TabIndex        =   7
         Top             =   1368
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   6
         Left            =   3330
         TabIndex        =   6
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   5
         Left            =   1035
         TabIndex        =   5
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   4
         Left            =   3330
         TabIndex        =   4
         Top             =   804
         Width           =   1100
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   3
         Left            =   1035
         TabIndex        =   3
         Top             =   804
         Width           =   1100
      End
      Begin VB.CommandButton cmd 
         Caption         =   "刪除(&D)"
         Height          =   400
         Index           =   1
         Left            =   8352
         TabIndex        =   11
         Top             =   144
         Width           =   810
      End
      Begin VB.CommandButton cmd 
         Caption         =   "加入(A)"
         Height          =   400
         Index           =   0
         Left            =   7500
         TabIndex        =   10
         Top             =   144
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "張"
         Height          =   180
         Index           =   17
         Left            =   8760
         TabIndex        =   37
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Height          =   180
         Index           =   16
         Left            =   8280
         TabIndex        =   36
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "點"
         Height          =   180
         Index           =   15
         Left            =   7920
         TabIndex        =   35
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Height          =   180
         Index           =   14
         Left            =   7440
         TabIndex        =   34
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   " 1 件"
         Height          =   180
         Index           =   13
         Left            =   6960
         TabIndex        =   33
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "點"
         Height          =   180
         Index           =   12
         Left            =   7920
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Height          =   180
         Index           =   11
         Left            =   7440
         TabIndex        =   31
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   " 1 件"
         Height          =   180
         Index           =   10
         Left            =   6960
         TabIndex        =   30
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖張數："
         Height          =   180
         Index           =   9
         Left            =   4680
         TabIndex        =   28
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖件數："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1368
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖點數："
         Height          =   180
         Index           =   3
         Left            =   2400
         TabIndex        =   25
         Top             =   1368
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "其他件數："
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "其他點數："
         Height          =   180
         Index           =   7
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "專業件數："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   804
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "專業點數："
         Height          =   180
         Index           =   5
         Left            =   2400
         TabIndex        =   21
         Top             =   804
         Width           =   915
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   1
         Left            =   2430
         TabIndex        =   20
         Top             =   540
         Width           =   2385
         VariousPropertyBits=   27
         Caption         =   "111"
         Size            =   "4207;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   540
         Width           =   1995
         VariousPropertyBits=   27
         Caption         =   "111"
         Size            =   "3519;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   108
         X2              =   7293
         Y1              =   1656
         Y2              =   1656
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8136
      Top             =   840
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090615.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl1 
      Caption         =   "111"
      Height          =   180
      Left            =   6372
      TabIndex        =   17
      Top             =   1092
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "目標年月："
      Height          =   180
      Index           =   2
      Left            =   2450
      TabIndex        =   16
      Top             =   1092
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "部門別："
      Height          =   180
      Index           =   1
      Left            =   4850
      TabIndex        =   15
      Top             =   1092
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   170
      TabIndex        =   14
      Top             =   1092
      Width           =   948
   End
End
Attribute VB_Name = "frm090615"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (grd1,Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim i As Integer, j As Integer, k As Integer, s As Integer, TextOk As Boolean, SeekAction As Integer, SeekRec As Variant
Dim StrSQL6 As String, strTemp1 As Variant, SeekTemp As String, DELMenu() As String, DELTemp() As String, SeekBmk1 As Variant, SeekBmk2 As Variant, SeekBmk3 As Variant
Dim strTemp(0 To 9) As String, PLeft(0 To 9) As Integer, Page As Integer, iPrint As Integer, seekbmk, BolDbOk As Boolean
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Add By Cheng 2002/05/24
Dim m_blnCancel As Boolean

Private Sub cmd_Click(Index As Integer)
If SeekAction <> 1 And SeekAction <> 0 Then
    Exit Sub
End If
With grd1
Select Case Index
Case 0
     For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            For j = 3 To 9
                .col = j
                .Text = Txt1(j)
            Next j
            Exit For
        End If
     Next i
Case 1
     For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
           'Modify by Morgan 2004/4/22
           '刪除時移除
'            .Col = 1
'            For j = 3 To 9
'                .Col = j
'                .Text = "0"
'            Next j
'            grd1_RowColChange
            .RemoveItem .row
            Exit For
        End If
     Next i
Case Else
End Select
End With
End Sub

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     PrintData
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Me.Hide
     frm090615_1.Show
Case Else
End Select
End Sub

Sub REFormLoad()
    SeekAction = 4
    ProcessUp
    ProcessDown
    TxtLock 3
    TxtSitu True
    ReDim DELMenu(0) As String
    ReDim DELTemp(0) As String
End Sub

Sub PrintData()
Page = 1
With adoRecordset
    'seekbmk = .Bookmark
    'If .RecordCount <> 0 And .RecordCount > 0 Then
    '    .MoveFirst
    '    Do While .EOF = False
            'strSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),pe01,st02,pe06,pe05,pe08,pe07,PE11,PE09,PE10,'',st06 from performance,staff where PE01=ST01(+) AND pe02='" & CheckStr(.Fields(0)) & "' and pe03=" & Val(CheckStr(.Fields(1))) + 191100 & " and ST04='1' AND st03='" & CheckStr(.Fields(2)) & "' order by st06,2,3 "
            strSql = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),pe01,st02,pe05,pe06,pe07,pe08,PE09,PE11,PE10,st06 from performance,staff where PE01=ST01(+) AND pe02='" & Trim(Txt1(0)) & "' and pe03=" & Val(Txt1(1)) + 191100 & " and ST04='1' AND st03='" & Trim(Txt1(2)) & "' order by st06,2,3 "
            CheckOC2
            With adoRecordset1
                .CursorLocation = adUseClient
                .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If .RecordCount <> 0 And .RecordCount > 0 Then
                    .MoveFirst
'edit by nickc 2007/04/02 以前有拿掉，卻沒將此段移除，會導致每次都會先出一張空白
'                    If adoRecordset.Bookmark <> 1 Then
'                        Page = Page + 1
'                        Printer.NewPage
'                    End If
                    PrintTitle
                    Do While .EOF = False
                        For i = 0 To 9
                            strTemp(i) = CheckStr(.Fields(i))
                            If i >= 3 And Len(strTemp(i)) = 0 Then
                                strTemp(i) = "0"
                            End If
                        Next i
                        PrintDatil
                        If iPrint >= 15000 Then
                            Page = Page + 1
                            Printer.NewPage
                            PrintTitle
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            CheckOC2
            '.MoveNext
    '    Loop
    'Else
    'End If
    '.Bookmark = seekbmk
End With
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintDatil() '列印資料

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 3 To 7 Step 2
    Printer.CurrentX = PLeft(i) + 1000 - Printer.TextWidth(Format(strTemp(i), "####0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0.00")
    Printer.CurrentX = PLeft(i + 1) + 1000 - Printer.TextWidth(Format(strTemp(i + 1), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "####0")
Next i
Printer.CurrentX = PLeft(9) + 1000 - Printer.TextWidth(Format(strTemp(9), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(9), "####0")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
PLeft(3) = 4500
PLeft(4) = PLeft(3) + 1750
PLeft(5) = PLeft(4) + 1750
PLeft(6) = PLeft(5) + 1750
PLeft(7) = PLeft(6) + 1750
PLeft(8) = PLeft(7) + 1750
PLeft(9) = PLeft(8) + 1750
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

Sub PrintTitle() '列印抬頭

GetPleft
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "個人、繪圖人員目標資料明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & CheckStr(adoRecordset.Fields(0))
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "部門別：" & CheckStr(adoRecordset.Fields(2)) & "  " & lbl1.Caption
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "目標年月：" & CheckStr(adoRecordset.Fields(1))
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(16000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "所別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "員工編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "姓名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "專業件數"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "專業點數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "其他件數"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "其他點數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "繪圖件數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "繪圖點數"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "繪圖張數"
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(16000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF2
     If SeekAction >= 4 Then
        YNEdit 0
     End If
Case vbKeyF3
     If SeekAction >= 4 Then
        YNEdit 1
     End If
Case vbKeyF5
     If SeekAction >= 4 Then
        YNEdit 2
     End If
Case vbKeyF4
     If SeekAction >= 4 Then
        YNEdit 3
     End If
Case vbKeyHome
     If SeekAction >= 4 Then
        MoveRec 0
     End If
Case vbKeyPageUp
     If SeekAction >= 4 Then
        MoveRec 1
     End If
Case vbKeyPageDown
     If SeekAction >= 4 Then
        MoveRec 2
     End If
Case vbKeyEnd
     If SeekAction >= 4 Then
        MoveRec 3
     End If
Case vbKeyF9, vbKeyReturn
     If SeekAction >= 0 And SeekAction <= 3 Then
         YNEdit 4
     End If
Case vbKeyF10
     If SeekAction >= 0 And SeekAction <= 3 Then
        YNEdit 5
     End If
Case vbKeyEscape
     If SeekAction >= 4 Then
        Unload Me
     End If
Case Else
End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
      If SeekAction > 3 Then
         If m_bInsert Then
             TBar1.Buttons(1).Enabled = True
         Else
             TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
             TBar1.Buttons(2).Enabled = True
         Else
             TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
             TBar1.Buttons(3).Enabled = True
         Else
             TBar1.Buttons(3).Enabled = False
         End If
      End If
   End If

End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm090615", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090615", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090615", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090615", strFind, False)
    MoveFormToCenter Me
    SeekAction = 4
    grd1.Cols = 10
    ProcessUp
    ProcessDown
    TxtLock 3
    TxtSitu True
    ReDim DELMenu(0) As String
    ReDim DELTemp(0) As String
       If m_bInsert Then
       TBar1.Buttons(1).Enabled = True
   Else
       TBar1.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       TBar1.Buttons(2).Enabled = True
   Else
       TBar1.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       TBar1.Buttons(3).Enabled = True
   Else
       TBar1.Buttons(3).Enabled = False
   End If

   '92.5.30 ADD BY SONIA
   Select Case Txt1(0)
      Case "P"
         Label1(11) = "13": Label1(14) = "14.5": Label1(16) = "5"
         Label1(10).Visible = True: Label1(12).Visible = True: Label1(13).Visible = True
         Label1(15).Visible = True: Label1(17).Visible = True
         
      Case "CFP"
         Label1(11) = "30": Label1(14) = "14.5": Label1(16) = "5"
         Label1(10).Visible = True: Label1(12).Visible = True: Label1(13).Visible = True
         Label1(15).Visible = True: Label1(17).Visible = True
         
      Case Else
         Label1(11) = "": Label1(14) = "": Label1(16) = ""
         Label1(10).Visible = False: Label1(12).Visible = False: Label1(13).Visible = False
         Label1(15).Visible = False: Label1(17).Visible = False
   End Select
   '92.5.30 END
End Sub

Sub ProcessUp()
StrSQL6 = " "
If Len(Systemkind_g) <> 0 Then
    'strTemp1 = Split(Systemkind_g, ",")
    'For i = 0 To UBound(strTemp1)
    StrSQL6 = StrSQL6 & " and PE02 in (" & GetAddStr(Systemkind_g) & ") "
    '    If i <> UBound(strTemp1) Then
    '        StrSQL6 = StrSQL6 + " OR "
    '    End If
    'Next i
    'StrSQL6 = StrSQL6 + " ) "
End If
'取的上半部資料
strSql = "select DISTINCT pe02 AS A,pe03-191100 AS B,st03 AS C,nvl(a0902,a0903),PE02||TO_CHAR(PE03-191100)||ST03 AS D from performance,staff,acc090 where pe01=st01(+) and ST04='1' AND st03=a0901(+) and substr(st03,1,1)='P' " & StrSQL6 & " order by 1,2,3 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        BolDbOk = True
        GetDataUp
    Else
        BolDbOk = False
        For i = 0 To 2
            Txt1(i) = ""
        Next i
    End If
End With
End Sub
        
Private Sub GetDataUp()         '取得上半部資料
If adoRecordset.RecordCount = 0 Then
    For i = 0 To 2
        Txt1(i) = ""
    Next i
    lbl1.Caption = ""
Else
    For i = 0 To 2
        Txt1(i) = CheckStr(adoRecordset.Fields(i))
    Next i

    lbl1.Caption = CheckStr(adoRecordset.Fields(3))

End If
End Sub

Sub ProcessDown()
'grd1.Clear
'grd1.Rows = 2
'SetGrd1
strSql = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),ST01,st02,pe05,pe06,pe07,pe08,PE09,PE11,PE10,st06 from staff,performance where ST01=PE01(+) AND '" & Trim(Txt1(0)) & "'=pe02(+) and " & Val(Txt1(1)) + 191100 & "=pe03(+) and ST04='1' AND st03='" & Trim(Txt1(2)) & "' order by st06,2,3 "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset1
        grd1.row = 1
        TextOk = True
        'grd1_Click
        GetDataDown
    Else
        grd1.Clear
        grd1.Rows = 2
        GetDataDown
    End If
    SetPieceRate 'Add by Morgan 2011/3/31
End With
CheckOC2
SetGrd1
End Sub
'Add by Morgan 2011/3/31
'設定張數/件
Private Sub SetPieceRate()
   If Trim(Txt1(0)) = "P" Or Trim(Txt1(0)) = "CFP" Then
      If Val(Txt1(1)) < 10004 Then
         Label1(16) = "5"
      Else
         Label1(16) = "7"
      End If
   Else
      Label1(16) = ""
   End If
End Sub

Private Sub GetDataDown()         '取得下半部資料
grd1_RowColChange
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As TextBox, i As Integer
   Select Case Lt
      Case 0
         TxtLock 1
         For Each txt In frm090615.Txt1
            txt.Locked = True
         Next
      Case 1
         For Each txt In frm090615.Txt1
            txt.Locked = False
            txt.Enabled = True
         Next
      Case 2
         For Each txt In frm090615.Txt1
            txt.Text = ""
         Next
         lbl1.Caption = ""
         lbl2(0).Caption = ""
         lbl2(1).Caption = ""
         grd1.Clear
         grd1.Rows = 2
         SetGrd1
      Case 3
         For i = 0 To 2
            If SeekAction = 0 Or SeekAction = 1 Then
                Txt1(i).Enabled = False
            Else
                Txt1(i).Locked = True
            End If
         Next i
      Case 4
         For i = 3 To 9
            Txt1(i).Enabled = False
         Next i
   End Select
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090615 = Nothing
End Sub

Private Sub YNEdit(ByVal Strindex As Integer)

'911107 nickchen
On Error GoTo CheckingErr

Select Case Strindex
Case 0  'ADD
     TxtSitu False
     SeekAction = 0
     TxtLock 2
     TxtLock 4
     cmdok(0).Enabled = False
     cmdok(1).Enabled = False
     Txt1(0).SetFocus
Case 1  'EDIT
     TxtSitu False
     SeekAction = 1
     TxtLock 3
Case 2  'DEL
     TxtSitu False
     TxtLock 4
     SeekAction = 2
     SeekRec = adoRecordset.Bookmark
     If MsgBox("是否要刪除此筆資料??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
        YNEdit 4
     End If
     TxtSitu True
     TxtLock 0
     Txt1(0).SetFocus
     txt1_GotFocus (0)
     SeekAction = 4
Case 3  'FIND
     TxtSitu False
     TxtLock 2
     TxtLock 4
     SeekAction = 3
     SeekRec = adoRecordset.Bookmark
     Txt1(0).SetFocus
     Exit Sub
Case 4  'ENTER
     Select Case SeekAction
     Case 0
          grd1.row = 1
          grd1.col = 1
          If Len(Txt1(0)) <> 0 And Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 And Len(grd1.Text) <> 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
                
                '911107 nickchen
                cnnConnection.BeginTrans
                
                For i = 1 To grd1.Rows - 1
                    grd1.row = i
                    strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06,PE07,PE08,PE09,PE10,PE11) VALUES ('"
                    grd1.col = 1
                    strSql = strSql & Trim(grd1.Text) & "','" & Txt1(0) & "'," & Val(Txt1(1)) + 191100 & ","
                    grd1.col = 3
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 4
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 5
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 6
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 7
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 9
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 8
                    strSql = strSql & Val(grd1.Text) & ")"
                    cnnConnection.Execute strSql
                Next i
                
                '911107 nickchen
                cnnConnection.CommitTrans
          Else
              s = MsgBox("沒有資料可存入資料庫!!", , "USER 輸入錯誤")
          End If
     Case 1
          SeekRec = adoRecordset.Bookmark
          grd1.row = 1
          grd1.col = 1
          If Len(Txt1(0)) <> 0 And Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
                
                '911107 nickchen
                cnnConnection.BeginTrans
                'Modify by Morgan 2004/4/22
                '考慮移除單筆明細，故先全部刪除
                cnnConnection.Execute ("delete performance where pe02='" & Txt1(0) & "' and pe03=" & Val(Txt1(1)) + 191100 & " AND pe01 in (select st01 from staff where st03='" & Txt1(2) & "' )")
                For i = 1 To grd1.Rows - 1
                   grd1.row = i
                   '92.5.13 MODIFY BY SONIA
                   strExc(0) = "SELECT * FROM PERFORMANCE WHERE PE01='" & Trim(grd1.Text) & "' AND PE02='" & Txt1(0) & "' AND PE03=" & Val(Txt1(1)) + 191100
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                       strSql = "UPDATE PERFORMANCE SET PE05="
                       grd1.col = 3
                       strSql = strSql & Val(grd1.Text) & ",PE06="
                       grd1.col = 4
                       strSql = strSql & Val(grd1.Text) & ",PE07="
                       grd1.col = 5
                       strSql = strSql & Val(grd1.Text) & ",PE08="
                       grd1.col = 6
                       strSql = strSql & Val(grd1.Text) & ",PE09="
                       grd1.col = 7
                       strSql = strSql & Val(grd1.Text) & ",PE10="
                       grd1.col = 9
                       strSql = strSql & Val(grd1.Text) & ",PE11="
                       grd1.col = 8
                       strSql = strSql & Val(grd1.Text) & " "
                       grd1.col = 1
                       strSql = strSql & " WHERE PE01='" & Trim(grd1.Text) & "' AND PE02='" & Txt1(0) & "' AND PE03=" & Val(Txt1(1)) + 191100
                   Else
                       strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06,PE07,PE08,PE09,PE10,PE11) VALUES ('"
                       grd1.col = 1
                       strSql = strSql & Trim(grd1.Text) & "','" & Txt1(0) & "'," & Val(Txt1(1)) + 191100 & ","
                       grd1.col = 3
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 4
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 5
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 6
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 7
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 9
                       strSql = strSql & Val(grd1.Text) & ","
                       grd1.col = 8
                       strSql = strSql & Val(grd1.Text) & ")"
                       grd1.col = 1
                  End If
                  cnnConnection.Execute strSql
                Next i
                
                '911107 nickchen
                cnnConnection.CommitTrans
                
          End If
          TxtSitu True
          ProcessUp
          adoRecordset.Bookmark = SeekRec
          GetDataUp
         ProcessDown
         cmdok(0).Enabled = True
         cmdok(1).Enabled = True
         SeekAction = 4
         ReDim DELMenu(0) As String
         ReDim DELTemp(0) As String
         TxtLock 1
         TxtLock 0
         Txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
     Case 2
          SeekRec = adoRecordset.Bookmark
          strSql = "SELECT ST01 FROM STAFF WHERE ST03='" & Trim(Txt1(2)) & "' AND ST04='1' "
          CheckOC2
          adoRecordset1.CursorLocation = adUseClient
          adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
              adoRecordset1.MoveFirst
              
              '911107 nickchen
              cnnConnection.BeginTrans
              
              Do While adoRecordset1.EOF = False
              If Len(CheckStr(adoRecordset1.Fields(0))) <> 0 Then
                  strSql = "DELETE FROM PERFORMANCE WHERE PE01='" & CheckStr(adoRecordset1.Fields(0)) & "' AND PE02='" & Trim(Txt1(0)) & "' AND PE03=" & Val(Txt1(1)) + 191100
                  cnnConnection.Execute strSql
              End If
                adoRecordset1.MoveNext
              Loop
              
              '911107 nickchen
              cnnConnection.CommitTrans
              
          End If
          TxtSitu True
          ProcessUp
          If SeekRec > adoRecordset.RecordCount Then
            adoRecordset.MoveFirst
          Else
            adoRecordset.Bookmark = SeekRec
          End If
          GetDataUp
          ProcessDown
          cmdok(0).Enabled = True
          cmdok(1).Enabled = True
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          Txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 3
          SeekRec = adoRecordset.Bookmark
          adoRecordset.Find "D='" & Txt1(0) & Txt1(1) & Txt1(2) & "'", 0, adSearchForward, 1
          If adoRecordset.EOF Then
              s = MsgBox("沒有符合資料!!", , "錯誤")
              adoRecordset.Bookmark = SeekRec
          End If
          TxtSitu True
          GetDataUp
          ProcessDown
          cmdok(0).Enabled = True
          cmdok(1).Enabled = True
          SeekAction = 4
          TxtLock 0
          Txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case Else
     End Select
     TxtSitu True
     ProcessUp
     ProcessDown
     cmdok(0).Enabled = True
     cmdok(1).Enabled = True
     SeekAction = 4
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 1
     TxtLock 0
     Txt1(0).SetFocus
     txt1_GotFocus (0)
Case 5  'CHANCL
     Select Case SeekAction
     Case 0
          If Len(Txt1(0)) <> 0 And Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
     Case 1
          If Len(Txt1(0)) <> 0 And Len(Txt1(1)) <> 0 And Len(Txt1(2)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
          TxtLock 1
     Case 2
          adoRecordset.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          cmdok(0).Enabled = True
          cmdok(1).Enabled = True
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          Txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 3
          adoRecordset.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          cmdok(0).Enabled = True
          cmdok(1).Enabled = True
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          Txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case Else
     End Select
     TxtSitu True
     ProcessUp
     ProcessDown
     cmdok(0).Enabled = True
     cmdok(1).Enabled = True
     SeekAction = 4
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 1
     TxtLock 0
     Txt1(0).SetFocus
     txt1_GotFocus (0)
Case Else
End Select

 '911107 nick transation
   Exit Sub
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Sub

Private Sub MoveRec(ByVal Strindex As Integer)
With adoRecordset
    If .EOF = True And .EOF = True Then Exit Sub
    Select Case Strindex
    Case 0
         .MoveFirst
    Case 1
         .MovePrevious
         If .BOF Then
            DataErrorMessage (6)
            .MoveFirst
         End If
    Case 2
         .MoveNext
         If .EOF Then
            DataErrorMessage (7)
            .MoveLast
         End If
    Case 3
         .MoveLast
    Case Else
    End Select
    GetDataUp
    ProcessDown
End With
End Sub

Private Sub grd1_RowColChange()
With grd1
   s = .MouseRow
    .Visible = False
    .Cols = 10
    For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            For k = 0 To .Cols - 1
                .col = k
                .CellBackColor = QBColor(15)
            Next k
            Exit For
        End If
    Next i
    .col = 0
    If TextOk = True Then
        .row = 0
        TextOk = False
    Else
        .row = s
        '.mousec
    End If
    If .row = 0 Then
        .row = 1
    End If
    .col = 1
    lbl2(0).Caption = "員工編號：" & .Text
    .col = 2
    lbl2(1).Caption = "姓名：" & .Text
    .col = 3
    Txt1(3).Text = .Text
    .col = 4
    Txt1(4).Text = .Text
    .col = 5
    Txt1(5).Text = .Text
    .col = 6
    Txt1(6).Text = .Text
    .col = 7
    Txt1(7).Text = .Text
    .col = 8
    Txt1(8).Text = .Text
    .col = 9
    Txt1(9).Text = .Text
    For j = 0 To .Cols - 1
        .col = j
        If Len(.Text) <> 0 Then
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = &HFFC0C0
            Next i
            Exit For
        End If
    Next j
    'SetGrd1
    .Visible = True
End With

End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         YNEdit 0
      Case 2
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 1
         End If
      Case 3
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 2
         End If
      Case 4
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 3
         End If
      Case 6
         If BolDbOk = False Then Exit Sub
         MoveRec 0
      Case 7
         If BolDbOk = False Then Exit Sub
         MoveRec 1
      Case 8
         If BolDbOk = False Then Exit Sub
         MoveRec 2
      Case 9
         If BolDbOk = False Then Exit Sub
         MoveRec 3
      Case 11
         YNEdit 4
      Case 12
         YNEdit 5
      Case 14
         Unload Me
      End Select
         If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      Else
          TBar1.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
      Else
          TBar1.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          TBar1.Buttons(3).Enabled = True
      Else
          TBar1.Buttons(3).Enabled = False
      End If
   End If

End Sub

Private Sub txt1_GotFocus(Index As Integer)
'Txt1(Index).SetFocus
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 11
    .row = 0
    .col = 0:   .Text = "所別"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "員工編號"
    .ColWidth(1) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "姓名"
    .ColWidth(2) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "專業件數"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "專業點數"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "其他件數"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "其他點數"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "繪圖件數"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "繪圖點數"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "繪圖張數"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = ""
    .ColWidth(10) = 0
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
m_blnCancel = False
If Txt1(Index).Locked = True Then
    Exit Sub
End If
Select Case Index
Case 0
     If Len(Trim(Txt1(0))) <> 0 Then
         strSql = "SELECT COUNT(*) FROM SYSTEMKIND WHERE SK01='" & Trim(Txt1(0)) & "' "
         CheckOC2
         adoRecordset1.CursorLocation = adUseClient
         adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
             If Val(CheckStr(adoRecordset1.Fields(0))) = 0 Then
                CheckOC2
                s = MsgBox("無此系統別 " & Txt1(0) & " !!", , "USER 輸入錯誤")
                Txt1(0).SetFocus
                Txt1(0).SelStart = 0
                Txt1(0).SelLength = Len(Txt1(0))
                'Add By Cheng 2002/05/24
                m_blnCancel = True
                
                Exit Sub
             End If
         '92.5.30 ADD BY SONIA
         Select Case Txt1(0)
            Case "P"
               Label1(11) = "13": Label1(14) = "14.5": Label1(16) = "5"
               Label1(10).Visible = True: Label1(12).Visible = True: Label1(13).Visible = True
               Label1(15).Visible = True: Label1(17).Visible = True
            Case "CFP"
               Label1(11) = "30": Label1(14) = "14.5": Label1(16) = "5"
               Label1(10).Visible = True: Label1(12).Visible = True: Label1(13).Visible = True
               Label1(15).Visible = True: Label1(17).Visible = True
            Case Else
               Label1(11) = "": Label1(14) = "": Label1(16) = ""
               Label1(10).Visible = False: Label1(12).Visible = False: Label1(13).Visible = False
               Label1(15).Visible = False: Label1(17).Visible = False
         End Select
         '92.5.30 END
         End If
     End If
     SetPieceRate 'Add by Morgan 2011/3/31
     
Case 1
     If Len(Trim(Txt1(1))) <> 0 Then
         If IsDate(ChangeTStringToTDateString(Txt1(1) & "01")) = False Then
            s = MsgBox("年月輸入錯誤", , "USER 輸入錯誤")
            Txt1(1).SetFocus
            Txt1(1).SelStart = 0
            Txt1(1).SelLength = Len(Txt1(1))
            'Add By Cheng 2002/05/24
            m_blnCancel = True
            Exit Sub
         End If
     End If
      SetPieceRate 'Add by Morgan 2011/3/31
      
Case 2
     If Len(Trim(Txt1(2))) <> 0 Then
        strSql = "SELECT NVL(A0902,A0903) FROM ACC090 WHERE A0901='" & Txt1(2) & "' "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1.Caption = CheckStr(adoRecordset1.Fields(0))
        Else
            s = MsgBox("部門別輸入錯誤找不到!!", , "USER 輸入錯誤")
            lbl1.Caption = ""
            Txt1(2).SetFocus
            Txt1(2).SelStart = 0
            Txt1(2).SelLength = Len(Txt1(2))
            CheckOC2
            'Add By Cheng 2002/05/24
            m_blnCancel = True
            Exit Sub
        End If
     End If
     If SeekAction = 0 Then
        For i = 0 To 2
            If Len(Txt1(i)) = 0 Then
                s = MsgBox("系統類別與目標年月與部門別不可空白!!", , "USER 輸入錯誤")
                If Len(Txt1(2)) = 0 Then Txt1(2).SetFocus
                If Len(Txt1(1)) = 0 Then Txt1(1).SetFocus
                If Len(Txt1(0)) = 0 Then Txt1(0).SetFocus
                'Add By Cheng 2002/05/24
                m_blnCancel = True
                
                Exit Sub
            End If
        Next i
        strSql = "SELECT ST01 FROM STAFF WHERE ST04='1' AND ST03='" & Txt1(2) & "' AND SUBSTR(ST03,1,1)='P' "
        CheckOC2
        With adoRecordset1
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                SeekTemp = CheckStr(.Fields(0))
                strSql = "SELECT * FROM PERFORMANCE WHERE PE01='" & SeekTemp & "' AND PE02='" & Txt1(0) & "'  AND PE03=" & Val(Txt1(1)) + 191100
                CheckOC2
                .CursorLocation = adUseClient
                .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If .RecordCount <> 0 And .RecordCount > 0 Then
                    s = MsgBox("此系統類別與年月與部門已存在!!", , "USER 輸入錯誤")
                    CheckOC
                    Txt1(0).SetFocus
                    Txt1(0).SelStart = 0
                    Txt1(0).SelLength = Len(Txt1(0))
                    'Add By Cheng 2002/05/24
                    m_blnCancel = True
                    
                    Exit Sub
                End If
            End If
        End With
        CheckOC2
        strSql = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),ST01,st02,0,0,0,0,0,0,0,ST06 from staff where st03='" & Trim(Txt1(2)) & "' AND ST04='1' AND SUBSTR(ST03,1,1)='P' order by ST06,2,3 "
        CheckOC2
        With adoRecordset1
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                Set grd1.Recordset = adoRecordset1
                grd1.row = 1
                TextOk = True
                GetDataDown
            End If
        End With
        CheckOC2
        TxtLock 1
        TxtLock 3
     End If
     SetGrd1
'92.5.30 ADD BY SONIA
Case 3
     If IsNumeric(Txt1(Index)) = False And Len(Txt1(Index)) <> 0 Then
        s = MsgBox("輸入錯誤, 請輸入數字", , "USER 輸入錯誤")
        Txt1(Index).SetFocus
        Txt1(Index).SelStart = 0
        Txt1(Index).SelLength = Len(Txt1(Index))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
     If Txt1(4) = "" Then Txt1(4) = 0
     If Txt1(Index) <> "" And Txt1(4) = 0 And Label1(11) <> "" Then
        Txt1(4) = Format(Txt1(Index) * Label1(11), "####0.0")
     End If
Case 7
     If IsNumeric(Txt1(Index)) = False And Len(Txt1(Index)) <> 0 Then
        s = MsgBox("輸入錯誤, 請輸入數字", , "USER 輸入錯誤")
        Txt1(Index).SetFocus
        Txt1(Index).SelStart = 0
        Txt1(Index).SelLength = Len(Txt1(Index))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
     If Txt1(8) = "" Then Txt1(8) = 0
     If Txt1(Index) <> "" And Txt1(8) = 0 And Label1(14) <> "" Then
        Txt1(8) = Format(Txt1(Index) * Label1(14), "####0.0")
     End If
     If Txt1(9) = "" Then Txt1(9) = 0
     If Txt1(Index) <> "" And Txt1(9) = 0 And Label1(16) <> "" Then
        Txt1(9) = Format(Txt1(Index) * Label1(16), "####0.0")
     End If
Case 3, 4, 5, 6, 7, 8, 9
'92.5.30 END
Case 4, 5, 6, 8, 9
     If IsNumeric(Txt1(Index)) = False And Len(Txt1(Index)) <> 0 Then
        s = MsgBox("輸入錯誤, 請輸入數字", , "USER 輸入錯誤")
        Txt1(Index).SetFocus
        Txt1(Index).SelStart = 0
        Txt1(Index).SelLength = Len(Txt1(Index))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
Case Else
End Select
End Sub

Function CheckRec() As Boolean
   '92.5.30 modify by sonia
   'If adoRecordset.RecordCount <> 0 Then CheckRec = True Else CheckRec = False
   If grd1.Rows - 1 <> 0 Then
      CheckRec = True
   Else
      CheckRec = False
   End If
   '92.5.30 end
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.Txt1
   If objTxt.Enabled = True Then
      Cancel = False
      txt1_GotFocus objTxt.Index
      If m_blnCancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

