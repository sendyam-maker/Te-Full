VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010023 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "櫃檯每日信件查詢"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox textLI08 
      Height          =   300
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   19
      Top             =   390
      Width           =   1785
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   6750
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   7770
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox textLI02 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   390
      Width           =   765
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   30
      TabIndex        =   6
      Top             =   1770
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6959
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "大陸-一般信件"
      TabPicture(0)   =   "frm010023.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "非大陸-一般信件"
      TabPicture(1)   =   "frm010023.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "國外信件"
      TabPicture(2)   =   "frm010023.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grd1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "客戶信件"
      TabPicture(3)   =   "frm010023.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grd1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "退件"
      TabPicture(4)   =   "frm010023.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grd1(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3525
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6218
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
         Height          =   3525
         Index           =   1
         Left            =   -74970
         TabIndex        =   8
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6218
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
         Height          =   3525
         Index           =   2
         Left            =   -74970
         TabIndex        =   9
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6218
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
         Height          =   3525
         Index           =   3
         Left            =   -74970
         TabIndex        =   10
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6218
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
         Height          =   3525
         Index           =   4
         Left            =   -74970
         TabIndex        =   21
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6218
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
   Begin VB.TextBox textLI04 
      Height          =   300
      Left            =   5790
      MaxLength       =   12
      TabIndex        =   5
      Top             =   720
      Width           =   2265
   End
   Begin VB.TextBox textLI01 
      Alignment       =   2  '置中對齊
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   0
      Top             =   60
      Width           =   1035
   End
   Begin MSForms.TextBox textLI15 
      Height          =   300
      Left            =   840
      TabIndex        =   23
      Top             =   1380
      Width           =   4065
      VariousPropertyBits=   679493659
      MaxLength       =   100
      Size            =   "7170;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textLI05 
      Height          =   300
      Left            =   840
      TabIndex        =   22
      Top             =   1050
      Width           =   4065
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "7170;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textLI06 
      Height          =   300
      Left            =   5790
      TabIndex        =   17
      Top             =   1050
      Width           =   2925
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "5159;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textLI03 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   3705
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "6535;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Left            =   270
      TabIndex        =   20
      Top             =   1410
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "信件種類："
      Height          =   180
      Left            =   270
      TabIndex        =   18
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Left            =   270
      TabIndex        =   16
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "序號："
      Height          =   180
      Left            =   3090
      TabIndex        =   15
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收件人："
      Height          =   180
      Left            =   5040
      TabIndex        =   14
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "文號："
      Height          =   180
      Left            =   5220
      TabIndex        =   13
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公司名稱："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   12
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "信件日期："
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   270
      TabIndex        =   11
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm010023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/30 Form2.0已修改 textLI03/textLI05/textLI06/textLI15/grd1(全)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 第一筆資料的key
Dim m_FirstKEY(3) As String
' 最後一筆資料的key
Dim m_LastKEY(3) As String
' 目前正在顯示的key
Dim m_CurrKEY(3) As String
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iLine As Integer
Dim MaxLine As Integer
Dim PLeft(6) As Integer
Dim strTemp(6) As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        m_CurrKEY(0) = ""
        GetAllData
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer

MoveFormToCenter Me

textLI01.Text = strSrvDate(2)
textLI01.Locked = False
RefreshRange
GetAllData
ShowLastRecord
SetCtrlReadOnly True
SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010023 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   GRD1(0).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1(0).Cols - 1
      GRD1(0).row = 0
      GRD1(0).col = iRow
      GRD1(0).Text = arrGridHeadText(iRow)
      GRD1(0).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(0).CellAlignment = flexAlignCenterCenter
   Next
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   GRD1(1).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1(1).Cols - 1
      GRD1(1).row = 0
      GRD1(1).col = iRow
      GRD1(1).Text = arrGridHeadText(iRow)
      GRD1(1).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(1).CellAlignment = flexAlignCenterCenter
   Next
   'Modify By Sindy 2013/2/4 +收件人,但資料在畫面上不顯示出來
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 0, 2500)
   GRD1(2).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1(2).Cols - 1
      GRD1(2).row = 0
      GRD1(2).col = iRow
      GRD1(2).Text = arrGridHeadText(iRow)
      GRD1(2).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(2).CellAlignment = flexAlignCenterCenter
   Next
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   GRD1(3).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1(3).Cols - 1
      GRD1(3).row = 0
      GRD1(3).col = iRow
      GRD1(3).Text = arrGridHeadText(iRow)
      GRD1(3).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(3).CellAlignment = flexAlignCenterCenter
   Next
   'Add By Sindy 98/03/20
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   GRD1(4).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1(4).Cols - 1
      GRD1(4).row = 0
      GRD1(4).col = iRow
      GRD1(4).Text = arrGridHeadText(iRow)
      GRD1(4).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(4).CellAlignment = flexAlignCenterCenter
   Next
   '98/03/20 End
End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1(Index), x, y, nCol, nRow
GRD1(Index).col = nCol
GRD1(Index).row = nRow
End Sub

Private Sub grd1_SelChange(Index As Integer)
Dim tmpMouseRow
Dim i, j

GRD1(Index).Visible = False
tmpMouseRow = GRD1(Index).row
GRD1(Index).Visible = True
If tmpMouseRow <> 0 Then
    GRD1(Index).row = tmpMouseRow
    GRD1(Index).col = 0
    If GRD1(Index).CellBackColor <> &HFFC0C0 Then
         GRD1(Index).Visible = False
         For j = 1 To GRD1(Index).Rows - 1
             GRD1(Index).row = j
             For i = 0 To GRD1(Index).Cols - 1
                  GRD1(Index).col = i
                  GRD1(Index).CellBackColor = QBColor(15)
             Next i
        Next j
        GRD1(Index).row = tmpMouseRow
         For i = 0 To GRD1(Index).Cols - 1
             GRD1(Index).col = i
             GRD1(Index).CellBackColor = &HFFC0C0
         Next i
         m_CurrKEY(0) = Val(ChangeTStringToWString(textLI01))
         m_CurrKEY(1) = GRD1(Index).TextMatrix(tmpMouseRow, 0)
         m_CurrKEY(2) = Left(textLI08.Text, 1)
         UpdateCtrlData
         GRD1(Index).Visible = True
    End If
End If
End Sub

Private Sub ChgGrdData(Index As Integer, iRow As Integer)
GRD1(Index).Visible = False
Dim i, j, k
For i = 0 To 4 '3
    For j = 1 To GRD1(i).Rows - 1
        GRD1(i).row = j
        For k = 0 To GRD1(i).Cols - 1
            GRD1(i).col = k
            GRD1(i).CellBackColor = QBColor(15)
        Next k
    Next j
Next i
SSTab1.Tab = Index
If SSTab1.Tab = 2 Then
    Label4.Visible = False
    textLI06.Visible = False
Else
    Label4.Visible = True
    textLI06.Visible = True
End If
GRD1(Index).row = iRow
For j = 0 To GRD1(Index).Cols - 1
    GRD1(Index).col = j
    GRD1(Index).CellBackColor = &HFFC0C0
Next j
GRD1(Index).TopRow = iRow
GRD1(Index).Visible = True
End Sub

Private Sub ChgToNowData()
Dim i, j As Integer
 j = 0
For i = 1 To GRD1(Left(textLI08.Text, 1) - 1).Rows - 1
    If GRD1(Left(textLI08.Text, 1) - 1).TextMatrix(i, 0) = textLI02 Then
        j = i
        Exit For
    End If
Next i
If j <> 0 Then ChgGrdData Left(textLI08.Text, 1) - 1, j
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim IsRun As Boolean
If SSTab1.Tab = 2 Then
   Label4.Visible = False
   textLI06.Visible = False
Else
   Label4.Visible = True
   textLI06.Visible = True
End If
If GRD1(SSTab1.Tab).Rows > 1 Then
   m_CurrKEY(0) = Val(ChangeTStringToWString(textLI01))
   m_CurrKEY(1) = GRD1(SSTab1.Tab).TextMatrix(GRD1(SSTab1.Tab).Rows - 1, 0)
   m_CurrKEY(2) = Trim(SSTab1.Tab + 1)
   UpdateCtrlData
Else
   Call SettextLI08(Trim(SSTab1.Tab + 1))
   ClearField
End If
End Sub

Private Sub textLI01_GotFocus()
InverseTextBox textLI01
End Sub

Private Sub textLI01_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textLI01_LostFocus()
If Trim(textLI01) <> "" And textLI01.Locked = False Then
    m_CurrKEY(0) = ""
    'GetAllData
End If
End Sub

Private Sub textLI01_Validate(Cancel As Boolean)
If Trim(textLI01) <> "" And m_EditMode = 1 Then
    If CheckIsTaiwanDate(textLI01, False) = False Then
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
    ElseIf ChkWorkDay(ChangeTStringToWString(textLI01)) = False Then
        Cancel = True
        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
    End If
End If
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT li01,min(li02) li02,li08 FROM letterinput " & _
            "WHERE li01 = (select min(li01) from letterinput)  AND " & _
                  "li08 = (SELECT MIN(li08) FROM letterinput  " & _
                           "where li01 = (select min(li01) from letterinput)) group by li01,li08"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_FirstKEY(2) = rsTmp.Fields("li08")
   End If
   rsTmp.Close

   strSql = "SELECT li01,max(li02) li02,li08 FROM letterinput " & _
            "WHERE li01 = (select max(li01) from letterinput)  AND " & _
                  "li08 = (SELECT max(li08) FROM letterinput  " & _
                           "where li01 = (select max(li01) from letterinput)) group by li01,li08"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_LastKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_LastKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_LastKEY(2) = rsTmp.Fields("li08")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textLI08.Locked = bEnable
   textLI03.Locked = bEnable
   textLI04.Locked = bEnable
   textLI05.Locked = bEnable
   textLI06.Locked = bEnable
   SSTab1.Enabled = bEnable
   'Add By Sindy 98/03/20
   textLI15.Locked = bEnable
End Sub

Private Sub ClearField()
   textLI02 = Empty
   textLI03 = Empty
   textLI04 = Empty
   textLI05 = Empty
   textLI06 = Empty
   'Add By Sindy 98/03/20
   textLI15 = Empty
End Sub

Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM letterinput " & _
            "WHERE li01 = " & Val(m_CurrKEY(0)) & " AND " & _
                  "li02 = " & Val(m_CurrKEY(1)) & " and li08='" & m_CurrKEY(2) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      '信件種類
      If IsNull(rsTmp.Fields("li08")) = False Then
         Call SettextLI08(rsTmp.Fields("li08"))
         m_CurrKEY(2) = rsTmp.Fields("li08")
      End If
      SSTab1.Tab = rsTmp.Fields("li08") - 1
      If SSTab1.Tab = 2 Then
          Label4.Visible = False
          textLI06.Visible = False
      Else
          Label4.Visible = True
          textLI06.Visible = True
      End If
      If Val(m_CurrKEY(0)) <> Val(ChangeTStringToWString(textLI01)) Then
           GetAllData
      End If
      If IsNull(rsTmp.Fields("li01")) = False Then: textLI01 = ChangeWStringToTString(rsTmp.Fields("li01"))
      If IsNull(rsTmp.Fields("li02")) = False Then: textLI02 = rsTmp.Fields("li02"): m_CurrKEY(1) = textLI02
      If IsNull(rsTmp.Fields("li03")) = False Then: textLI03 = rsTmp.Fields("li03")
      If IsNull(rsTmp.Fields("li04")) = False Then: textLI04 = rsTmp.Fields("li04")
      If IsNull(rsTmp.Fields("li05")) = False Then: textLI05 = rsTmp.Fields("li05")
      If IsNull(rsTmp.Fields("li06")) = False Then: textLI06 = rsTmp.Fields("li06")
      'Add By Sindy 98/03/20
      If IsNull(rsTmp.Fields("li15")) = False Then: textLI15 = rsTmp.Fields("li15")
      ChgToNowData
   End If
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

'抓當日所有資料
Private Sub GetAllData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim Rani As Integer
   ClearField
   For Rani = 0 To 4 '3
        'Modify By Sindy 98/03/20
        'strSQL = "SELECT li02,li03,li04,li05,li06,li08 FROM letterinput "
        strSql = "SELECT li02,li03,li04,li05,li06,li15,li08 FROM letterinput " & _
                 "WHERE li01 = " & IIf(m_CurrKEY(0) = "", Val(ChangeTStringToWString(textLI01)), Val(m_CurrKEY(0))) & " AND " & _
                       " li08='" & Trim(Rani + 1) & "' order by li02"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        Set GRD1(Rani).Recordset = rsTmp
        rsTmp.Close
    Next Rani
    SetGrd
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SettextLI08(Index As Integer)
   If Index = 1 Then
      textLI08.Text = "1.大陸 -一般信件"
   ElseIf Index = 2 Then
      textLI08.Text = "2.非大陸一般信件"
   ElseIf Index = 3 Then
      textLI08.Text = "3.國外信件"
   ElseIf Index = 4 Then
      textLI08.Text = "4.客戶信件"
   'Add By Sindy 98/03/20
   ElseIf Index = 5 Then
      textLI08.Text = "5.退件"
   End If
End Sub
