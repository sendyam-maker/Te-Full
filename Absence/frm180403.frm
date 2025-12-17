VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180403 
   BorderStyle     =   1  '單線固定
   Caption         =   "職代/簽核主管關聯查詢"
   ClientHeight    =   6550
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6550
   ScaleWidth      =   8950
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   30
      TabIndex        =   18
      Top             =   1110
      Width           =   8925
      _ExtentX        =   15752
      _ExtentY        =   9119
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "職代/簽核主管"
      TabPicture(0)   =   "frm180403.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "grd1"
      Tab(0).Control(4)=   "Check1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "簽核主管特殊對象的簽核職代"
      TabPicture(1)   =   "frm180403.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grd2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CheckBox Check1 
         Caption         =   "設定99天一併顯示"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -71520
         TabIndex        =   21
         Top             =   360
         Width           =   2205
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Bindings        =   "frm180403.frx":0038
         Height          =   4095
         Left            =   -74940
         TabIndex        =   19
         Top             =   570
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   7214
         _Version        =   393216
         Cols            =   48
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   48
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Bindings        =   "frm180403.frx":004D
         Height          =   4725
         Left            =   60
         TabIndex        =   24
         Top             =   360
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   8326
         _Version        =   393216
         Cols            =   27
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   27
      End
      Begin VB.Label Label3 
         Caption         =   "案件代理類型：空白->所有案件"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -70290
         TabIndex        =   23
         Top             =   4710
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "                            1->台灣案，2->非台灣案"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -70290
         TabIndex        =   22
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "職代有分”人事”職代和”案件”職代"
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   -74850
         TabIndex        =   20
         Top             =   360
         Width           =   6645
      End
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   2
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   4
      Top             =   690
      Width           =   405
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   3
      Left            =   1770
      MaxLength       =   1
      TabIndex        =   5
      Top             =   690
      Width           =   405
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   1
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   2
      Top             =   30
      Width           =   495
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   1230
      MaxLength       =   3
      TabIndex        =   1
      Top             =   30
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   4020
      TabIndex        =   13
      Top             =   480
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   360
      Left            =   6210
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtB0101 
      Height          =   300
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   7050
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7890
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所  別："
      Height          =   180
      Index           =   2
      Left            =   300
      TabIndex        =   17
      Top             =   750
      Width           =   630
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   2040
      Y1              =   840
      Y2              =   840
   End
   Begin MSForms.TextBox txtB0101_2 
      Height          =   255
      Left            =   2010
      TabIndex        =   16
      Top             =   390
      Width           =   825
      VariousPropertyBits=   679495711
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "1455;450"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "（部門別查詢時，僅顯示部門職代表）"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   2430
      TabIndex        =   15
      Top             =   90
      Width           =   3060
   End
   Begin VB.Line Line2 
      X1              =   1650
      X2              =   2130
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註：離職人員為紅色底，輸入員工代號查詢者為藍色底"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   6330
      Width           =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Index           =   15
      Left            =   300
      TabIndex        =   11
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(已離職)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   2880
      TabIndex        =   10
      Top             =   420
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   9
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm180403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/9/9
Option Explicit

'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim i As Integer, j As Integer
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 25) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblPrevRow As Double, dblPrevRow2 As Double
Dim strPrinter As String 'Added by Sindy 2021/1/26


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   strSystemKind = GetSystemKindByNick
   
   'Modify By Sindy 2021/1/26
'   strSql = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'      Set Printer = Printers(i)
'      Combo1.AddItem Printer.DeviceName, j
'      j = j + 1
'      If Printer.DeviceName = strSql Then
'         SeekPrint = i
'      End If
'   Next i
'   Set Printer = Printers(SeekPrint)
'   Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   '2021/1/26 END
   
   MoveFormToCenter Me
   Label2(0).Visible = False
   SetDataListWidth
   SetDataListWidth2 'Add By Sindy 2022/5/30
   
   Check1.Visible = False 'Add By Sindy 2022/5/3
   If ChkIsAbsBoss(strUserNum) = True Or _
      GetStaffDepartment(strUserNum) = "M51" Or _
      GetStaffDepartment(strUserNum) = "M21" Then
      cmdPrint.Visible = True
      Frame1.Visible = True
      'grd1.Height = 4395
      Check1.Visible = True 'Add By Sindy 2022/5/3
   Else
      cmdPrint.Visible = False
      Frame1.Visible = False
      'grd1.Height = 5025
   End If
   SSTab1.Tab = 0 'Add By Sindy 2022/5/30
End Sub

'Add By Sindy 2021/7/12
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DestroyToolTip '清除物件 Add By Sindy 2021/7/12
   Set frm180403 = Nothing
End Sub

Private Function CheckData() As Boolean
Dim Cancel As Boolean
   
   CheckData = False
   
   If txtData(0).Text = "" And txtB0101.Text = "" And txtData(2).Text = "" Then
      MsgBox "至少輸入一項查詢條件！", vbExclamation
      txtData(0).SetFocus
      Exit Function
   End If
   
   Cancel = False
   If txtData(0) <> "" Then
      Call Txtdata_Validate(0, Cancel)
      If Cancel = True Then
         Exit Function
      End If
   End If
   If txtData(1) <> "" Then
      Call Txtdata_Validate(1, Cancel)
      If Cancel = True Then
         Exit Function
      End If
   End If
   If txtB0101 <> "" Then
      Call txtB0101_Validate(Cancel)
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   CheckData = True
End Function

'Add By Sindy 2022/5/30
Private Sub SetDataListWidth2()
grd2.Visible = False
grd2.FixedCols = 0
grd2.row = 0
grd1.col = 0: grd2.Text = "部門"
grd1.ColWidth(0) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd2.col = 1: grd2.Text = "被簽核對象" '員工姓名
grd2.ColWidth(1) = 1000
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 2: grd2.Text = "簽核主管"
grd2.ColWidth(2) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 3: grd2.Text = "職代一(1)"
grd2.ColWidth(3) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 4: grd2.Text = "職代一(2)"
grd2.ColWidth(4) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 5: grd2.Text = "職代二(1)"
grd2.ColWidth(5) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 6: grd2.Text = "職代二(2)"
grd2.ColWidth(6) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 7: grd2.Text = "職代三(1)"
grd2.ColWidth(7) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 8: grd2.Text = "職代三(2)"
grd2.ColWidth(8) = 850
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 9: grd2.Text = "s0.ST04"
grd2.ColWidth(9) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 10: grd2.Text = "s1.ST04"
grd2.ColWidth(10) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 11: grd2.Text = "s2.ST04"
grd2.ColWidth(11) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 12: grd2.Text = "s3.ST04"
grd2.ColWidth(12) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 13: grd2.Text = "s4.ST04"
grd2.ColWidth(13) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 14: grd2.Text = "s5.ST04"
grd2.ColWidth(14) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 15: grd2.Text = "s6.ST04"
grd2.ColWidth(15) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 16: grd2.Text = "s7.ST04"
grd2.ColWidth(16) = 0
grd2.CellAlignment = flexAlignLeftCenter

grd2.col = 17: grd2.Text = "B0201"
grd2.ColWidth(17) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 18: grd2.Text = "B0202"
grd2.ColWidth(18) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 19: grd2.Text = "B0203"
grd2.ColWidth(19) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 20: grd2.Text = "B0204"
grd2.ColWidth(20) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 21: grd2.Text = "B0205"
grd2.ColWidth(21) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 22: grd2.Text = "B0206"
grd2.ColWidth(22) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 23: grd2.Text = "B0207"
grd2.ColWidth(23) = 0
grd2.CellAlignment = flexAlignLeftCenter
grd2.col = 24: grd2.Text = "B0208"
grd2.ColWidth(24) = 0
grd2.CellAlignment = flexAlignLeftCenter
'Add By Sindy 2023/5/4
grd2.col = 25: grd2.Text = "簽核種類"
grd2.ColWidth(25) = 850
grd2.CellAlignment = flexAlignLeftCenter
'2023/5/4 END
'Add By Sindy 2025/3/6
grd2.col = 26: grd2.Text = "A0921"
grd2.ColWidth(26) = 0
grd2.CellAlignment = flexAlignLeftCenter
'2025/3/6 END

grd2.Visible = True
End Sub

'Add By Sindy 2022/5/30 簽核主管特殊對象的簽核職代
Private Sub Query_Spec()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
      
   grd2.Rows = 2
   grd2.Clear
   SetDataListWidth2
   
   If CheckData = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   strCon = ""
   If txtData(0) <> "" Then
'      'Modify By Sindy 2023/12/20
'      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & "and s1.ST93>='" & txtData(0) & "' "
'      Else
'      '2023/12/20 END
'         strCon = strCon & "and s0.ST03>='" & txtData(0) & "' "
'      End If
   End If
   If txtData(1) <> "" Then
'      'Modify By Sindy 2023/12/20
'      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & "and s1.ST93<='" & txtData(1) & "' "
'      Else
'      '2023/12/20 END
'         strCon = strCon & "and s0.ST03<='" & txtData(1) & "' "
'      End If
   End If
   If txtData(2) <> "" Then
      strCon = strCon & "and s1.ST06>='" & txtData(2) & "' "
   End If
   If txtData(3) <> "" Then
      strCon = strCon & "and s1.ST06<='" & txtData(3) & "' "
   End If
   If txtB0101 <> "" Then
      strCon = strCon & "and (B0201='" & txtB0101 & "' or B0202='" & txtB0101 & "' or B0203='" & txtB0101 & "'" & _
                        " or B0204='" & txtB0101 & "' or B0205='" & txtB0101 & "' or B0206='" & txtB0101 & "'" & _
                        " or B0207='" & txtB0101 & "' or B0208='" & txtB0101 & "')"
   End If
   
   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2023/12/20
   'Modify By Sindy 2025/3/6 再調整SQL
   strSql = "select decode(B0201,B0208,'',A0922) A0922,decode(B0209,'1',s0.ST02,decode(B0208,B0201,'',s0.ST02))" & _
                  ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02" & _
                  ",s0.ST04,s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04" & _
                  ",B0201,B0202,B0203,B0204,B0205,B0206,B0207,B0208,decode(B0209,'1','人事','2','案件',B0209),A0921" & _
            " from ABS002,ACC090NEW,STAFF s0,STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7" & _
            " where s0.ST01=B0208 and s0.ST01 is not null and length(B0208)=5" & _
            " and s0.ST93=A0921 and B0201=s1.ST01(+)" & _
            " and B0202=s2.ST01(+) and B0203=s3.ST01(+)" & _
            " and B0204=s4.ST01(+) and B0205=s5.ST01(+)" & _
            " and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & strCon & _
            " and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0208 and sc02=(select max(sc02) from Staff_Change where sc01=B0208)))"
   strSql = strSql & " union " & _
            "select A0922,A0922" & _
                  ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02" & _
                  ",'',s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04" & _
                  ",B0201,B0202,B0203,B0204,B0205,B0206,B0207,B0208,decode(B0209,'1','人事','2','案件',B0209),A0921" & _
            " from ABS002,ACC090NEW,STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7" & _
            " where A0921=B0208 and length(B0208)<>5" & _
            " and B0201=s1.ST01(+)" & _
            " and B0202=s2.ST01(+) and B0203=s3.ST01(+)" & _
            " and B0204=s4.ST01(+) and B0205=s5.ST01(+)" & _
            " and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & strCon
   strSql = strSql & " order by A0921,B0208"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
      SetDataListWidth2
      grd2.FixedCols = 3
      GetSelChage2
   Else
      SSTab1.Tab = 0
   End If
   rsTmp.Close
   'SetDataListWidth2
   dblPrevRow2 = 0
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2022/5/30
Private Sub GetSelChage2(Optional intCurRow As Integer = 0)
Dim k As Integer
Dim strText As String, varArr As Variant
Dim intTotRows As Integer
Dim intChangRow As Integer '要異動第幾筆
   
   grd2.Visible = False
   If grd2.Rows - 1 > 0 Then
      intTotRows = grd2.Rows - 1
      If intCurRow > 0 Then
         If intCurRow > intTotRows Then
            intCurRow = intTotRows
         End If
         intTotRows = intCurRow
         intChangRow = intCurRow
      Else
         intChangRow = 1
      End If
      For j = intChangRow To intTotRows
         grd2.row = j
         '第一個欄位
         grd2.col = 0
         grd2.CellBackColor = &H8000000F
         '第二,三個欄位
         For i = 1 To 2
            grd2.col = i
            If i = 1 Then k = 9
            If i = 2 Then k = 10
            If grd2.TextMatrix(j, k) <> "1" And grd2.TextMatrix(j, k) <> "" Then '已離職
               grd2.CellBackColor = &HFF& '紅色
            ElseIf Trim(txtB0101_2.Text) <> "" And Trim(grd2.TextMatrix(j, i)) = Trim(txtB0101_2.Text) Then '為輸入員工代號者
               grd2.CellBackColor = &HFFFF80 '水藍色
            Else
               grd2.CellBackColor = &H8000000F '灰色
            End If
         Next i
         '其他欄位
         For i = 3 To 8
            grd2.col = i
            If i = 3 Then k = 11
            If i = 4 Then k = 12
            If i = 5 Then k = 13
            If i = 6 Then k = 14
            If i = 7 Then k = 15
            If i = 8 Then k = 16
            If grd2.TextMatrix(j, k) <> "1" And grd2.TextMatrix(j, k) <> "" Then '已離職
               grd2.CellBackColor = &HFF& '紅色
            ElseIf Trim(txtB0101_2.Text) <> "" And Trim(grd2.TextMatrix(j, i)) = Trim(txtB0101_2.Text) Then '為輸入員工代號者
               grd2.CellBackColor = &HFFFF80 '水藍色
            Else
               If intCurRow = 0 Then
                  grd2.CellBackColor = QBColor(15) '清除反白
               Else
                  If grd2.CellBackColor = &HFFC0C0 Then '已反白
                     grd2.CellBackColor = QBColor(15) '清除反白
                  Else
                     grd2.CellBackColor = &HFFC0C0 '資料列反白
                  End If
               End If
            End If
         Next i
      Next j
   End If
   grd2.Visible = True
End Sub

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim str99SQL As String
   
   If CheckData = False Then Exit Sub
   
   Call Query_Spec 'Add By Sindy 2022/5/30
   
   grd1.Rows = 2
   grd1.Clear
   SetDataListWidth
    
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   strCon = ""
   If txtData(0) <> "" Then
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and s0.ST93>='" & txtData(0) & "' "
      Else
      '2023/12/20 END
         strCon = strCon & " and s0.ST03>='" & txtData(0) & "' "
      End If
   End If
   If txtData(1) <> "" Then
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and s0.ST93<='" & txtData(1) & "' "
      Else
      '2023/12/20 END
         strCon = strCon & " and s0.ST03<='" & txtData(1) & "' "
      End If
   End If
   'Add By Sindy 2022/5/3
   If txtData(2) <> "" Then
      strCon = strCon & " and s0.ST06>='" & txtData(2) & "' "
   End If
   If txtData(3) <> "" Then
      strCon = strCon & " and s0.ST06<='" & txtData(3) & "' "
   End If
   '2022/5/3 END
   If txtB0101 <> "" Then
      'Modify By Sindy 2021/7/12 + ,B0128,B0129,B0124
      strCon = strCon & " and (B0101='" & txtB0101 & "' or B0102='" & txtB0101 & "' or B0103='" & txtB0101 & "'" & _
                         " or B0104='" & txtB0101 & "' or B0105='" & txtB0101 & "' or B0106='" & txtB0101 & "'" & _
                         " or B0107='" & txtB0101 & "' or B0108='" & txtB0101 & "' or B0109='" & txtB0101 & "'" & _
                         " or B0110='" & txtB0101 & "' or B0111='" & txtB0101 & "' or B0117='" & txtB0101 & "'" & _
                         " or B0119='" & txtB0101 & "' or B0121='" & txtB0101 & "' or B0123='" & txtB0101 & "'" & _
                         " or B0128='" & txtB0101 & "' or B0129='" & txtB0101 & "' or instr(B0124,'" & txtB0101 & "')>0" & _
                        ")"
   End If
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/5/27 簽核天數99時,該審核主管及天數都不要顯示
   'Modified by Lydia 2017/03/28 ST14改成多個編號 st14<>'99997'=> instr(st14,'99997')=0
   'Modify By Sindy 2017/8/21 + B0126 : 全發
   'Modify By Sindy 2021/7/12 + ,B0128,B0129,B0124
   'Modify By Sindy 2022/5/3 可選擇99天要不要顯示
   If Check1.Visible = True And Check1.Value = 1 Then
      str99SQL = "s7.ST02,B0112,s8.ST02,B0113,s9.ST02,B0114,s10.ST02,B0115"
   Else
      str99SQL = "decode(B0112,99,'',s7.ST02),decode(B0112,99,'',B0112),decode(B0113,99,'',s8.ST02),decode(B0113,99,'',B0113),decode(B0114,99,'',s9.ST02),decode(B0114,99,'',B0114),decode(B0115,99,'',s10.ST02),decode(B0115,99,'',B0115)"
   End If
   '2022/5/3 END
   'Modify By Sindy 2023/12/20
   strSql = "select A0922,s0.ST02 " & _
            ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02," & str99SQL & ",B0101,A0921,s0.ST04 " & _
            ",B0116,s11.ST02,B0118,s12.ST02,B0120,s13.ST02,B0122,s14.ST02,s15.ST02,s16.ST02 " & _
            ",s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04,s8.ST04,s9.ST04,s10.ST04,s11.ST04,s12.ST04,s13.ST04,s14.ST04,s15.ST04,s16.ST04,B0126,GETSTAFFNAMELIST(replace(B0124,';',',')) as B0124Name,B0124 " & _
            "from ABS001,ACC090NEW,STAFF s0 " & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7,STAFF s8,STAFF s9,STAFF s10,STAFF s11,STAFF s12,STAFF s13,STAFF s14,STAFF s15,STAFF s16 " & _
            "where s0.ST01=B0101(+) " & _
            "and (instr(s0.st14,'99997')=0 or s0.ST14 is null) and substr(s0.ST01,1,1) in(" & ST01CodeNum1 & ") and substr(s0.ST01,4,1)<>'9' and s0.st01 not in('60000','96029','96030','86026','67004','68007','63001') " & _
            "and s0.ST93=A0921(+) and B0102=s1.ST01(+) " & _
            "and B0103=s2.ST01(+) and B0104=s3.ST01(+) " & _
            "and B0105=s4.ST01(+) and B0106=s5.ST01(+) " & _
            "and B0107=s6.ST01(+) and B0108=s7.ST01(+) " & _
            "and B0109=s8.ST01(+) and B0110=s9.ST01(+) " & _
            "and B0111=s10.ST01(+) and B0117=s11.ST01(+) " & _
            "and B0119=s12.ST01(+) and B0121=s13.ST01(+) " & _
            "and B0123=s14.ST01(+) and B0128=s15.ST01(+) " & _
            "and B0129=s16.ST01(+)" & strCon & _
            "and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) " & _
            "order by A0921,B0101"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      SetDataListWidth
      grd1.FixedCols = 2
   End If
   rsTmp.Close
   'SetDataListWidth
   GetSelChage
   dblPrevRow = 0
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub GetSelChage(Optional intCurRow As Integer = 0)
Dim k As Integer
Dim strText As String, varArr As Variant
Dim intTotRows As Integer
Dim intChangRow As Integer '要異動第幾筆
   
   grd1.Visible = False
   If grd1.Rows - 1 > 0 Then
      intTotRows = grd1.Rows - 1
      If intCurRow > 0 Then
         If intCurRow > intTotRows Then
            intCurRow = intTotRows
         End If
         intTotRows = intCurRow
         intChangRow = intCurRow
      Else
         intChangRow = 1
      End If
      For j = intChangRow To intTotRows
         grd1.row = j
         '第一個欄位
         grd1.col = 0
         grd1.CellBackColor = &H8000000F
         '第二個欄位
         grd1.col = 1
         If grd1.TextMatrix(j, 18) <> "1" And grd1.TextMatrix(j, 18) <> "" Then '已離職
            grd1.CellBackColor = &HFF& '紅色
         ElseIf Trim(txtB0101_2.Text) <> "" And Trim(grd1.TextMatrix(j, 1)) = Trim(txtB0101_2.Text) Then '為輸入員工代號者
            grd1.CellBackColor = &HFFFF80    '水藍色
         Else
            grd1.CellBackColor = &H8000000F '灰色
         End If
         '其他欄位
         For i = 2 To grd1.Cols - 1
            grd1.col = i
            If i = 2 Then k = 29
            If i = 3 Then k = 30
            If i = 4 Then k = 31
            If i = 5 Then k = 32
            If i = 6 Then k = 33
            If i = 7 Then k = 34
            If i = 8 Then k = 35
            If i = 10 Then k = 36
            If i = 12 Then k = 37
            If i = 14 Then k = 38
            If i = 20 Then k = 39
            If i = 22 Then k = 40
            If i = 24 Then k = 41
            If i = 26 Then k = 42
            If i = 27 Then k = 43
            If i = 28 Then k = 44
            If grd1.TextMatrix(j, k) <> "1" And grd1.TextMatrix(j, k) <> "" Then '已離職
               grd1.CellBackColor = &HFF& '紅色
            ElseIf Trim(txtB0101_2.Text) <> "" And Trim(grd1.TextMatrix(j, i)) = Trim(txtB0101_2.Text) Then '為輸入員工代號者
               grd1.CellBackColor = &HFFFF80 '水藍色
            Else
               If intCurRow = 0 Then
                  grd1.CellBackColor = QBColor(15) '清除反白
               Else
                  If grd1.CellBackColor = &HFFC0C0 Then '已反白
                     grd1.CellBackColor = QBColor(15) '清除反白
                  Else
                     grd1.CellBackColor = &HFFC0C0 '資料列反白
                  End If
               End If
            End If
         Next i
         'Add By Sindy 2021/7/12
         strText = Trim(grd1.TextMatrix(j, 47)) '請假、出差核准後通知人員
         If strText <> "" Then
            varArr = Split(strText, ";")
            If UBound(varArr) >= 0 Then
               grd1.col = 46
               For i = 0 To UBound(varArr)
                  '檢查人員是否存在或離職
                  'Modify By Sindy 2023/7/26 +, False: 不彈訊息
                  If ChkStaffST04(Trim(varArr(i)), False) = True Then
                     grd1.CellBackColor = &HFF& '紅色
                     Exit For
                  ElseIf Trim(txtB0101_2.Text) <> "" And InStr(Trim(grd1.TextMatrix(j, 46)), Trim(txtB0101_2.Text)) > 0 Then '為輸入員工代號者
                     grd1.CellBackColor = &HFFFF80 '水藍色
                     Exit For
                  End If
               Next i
            End If
         End If
         '2021/7/12 END
      Next j
   End If
   grd1.Visible = True
End Sub

Private Sub SetDataListWidth()
grd1.Visible = False
grd1.FixedCols = 0
grd1.row = 0
grd1.col = 0: grd1.Text = "部門"
grd1.ColWidth(0) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 1: grd1.Text = "員工姓名"
grd1.ColWidth(1) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 2: grd1.Text = "職代一(1)"
grd1.ColWidth(2) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 3: grd1.Text = "職代一(2)"
grd1.ColWidth(3) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 4: grd1.Text = "職代二(1)"
grd1.ColWidth(4) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 5: grd1.Text = "職代二(2)"
grd1.ColWidth(5) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 6: grd1.Text = "職代三(1)"
grd1.ColWidth(6) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 7: grd1.Text = "職代三(2)"
grd1.ColWidth(7) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 8: grd1.Text = "審核主管1"
grd1.ColWidth(8) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 9: grd1.Text = "天數"
grd1.ColWidth(9) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 10: grd1.Text = "審核主管2"
grd1.ColWidth(10) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 11: grd1.Text = "天數"
grd1.ColWidth(11) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 12: grd1.Text = "審核主管3"
grd1.ColWidth(12) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 13: grd1.Text = "天數"
grd1.ColWidth(13) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 14: grd1.Text = "審核主管4"
grd1.ColWidth(14) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 15: grd1.Text = "天數"
grd1.ColWidth(15) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 16: grd1.Text = "B0101"
grd1.ColWidth(16) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 17: grd1.Text = "A0901"
grd1.ColWidth(17) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 18: grd1.Text = "ST04"
grd1.ColWidth(18) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 19: grd1.Text = "(1-1)類型"
grd1.ColWidth(19) = 600
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 20: grd1.Text = "案件職代一(1)"
grd1.ColWidth(20) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 21: grd1.Text = "(1-2)類型"
grd1.ColWidth(21) = 600
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 22: grd1.Text = "案件職代一(2)"
grd1.ColWidth(22) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 23: grd1.Text = "(2-1)類型"
grd1.ColWidth(23) = 600
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 24: grd1.Text = "案件職代二(1)"
grd1.ColWidth(24) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 25: grd1.Text = "(2-2)類型"
grd1.ColWidth(25) = 600
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 26: grd1.Text = "案件職代二(2)"
grd1.ColWidth(26) = 1000
grd1.CellAlignment = flexAlignLeftCenter
'Modify By Sindy 2021/7/12
grd1.col = 27: grd1.Text = "居家職代(1)"
grd1.ColWidth(27) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 28: grd1.Text = "居家職代(2)"
grd1.ColWidth(28) = 1000
grd1.CellAlignment = flexAlignLeftCenter
'2021/7/12 END
grd1.col = 29: grd1.Text = "s1.ST04"
grd1.ColWidth(29) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 30: grd1.Text = "s2.ST04"
grd1.ColWidth(30) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 31: grd1.Text = "s3.ST04"
grd1.ColWidth(31) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 32: grd1.Text = "s4.ST04"
grd1.ColWidth(32) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 33: grd1.Text = "s5.ST04"
grd1.ColWidth(33) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 34: grd1.Text = "s6.ST04"
grd1.ColWidth(34) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 35: grd1.Text = "s7.ST04"
grd1.ColWidth(35) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 36: grd1.Text = "s8.ST04"
grd1.ColWidth(36) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 37: grd1.Text = "s9.ST04"
grd1.ColWidth(37) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 38: grd1.Text = "s10.ST04"
grd1.ColWidth(38) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 39: grd1.Text = "s11.ST04"
grd1.ColWidth(39) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 40: grd1.Text = "s12.ST04"
grd1.ColWidth(40) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 41: grd1.Text = "s13.ST04"
grd1.ColWidth(41) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 42: grd1.Text = "s14.ST04"
grd1.ColWidth(42) = 0
grd1.CellAlignment = flexAlignLeftCenter
'Modify By Sindy 2021/7/12
grd1.col = 43: grd1.Text = "s15.ST04"
grd1.ColWidth(43) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 44: grd1.Text = "s16.ST04"
grd1.ColWidth(44) = 0
grd1.CellAlignment = flexAlignLeftCenter
'2021/7/12
'Add By Sindy 2017/8/21
'Modify By Sindy 2025/8/12
'grd1.col = 45: grd1.Text = "全發(案職)"
grd1.col = 45: grd1.Text = "全發"
'2025/8/12 END
grd1.ColWidth(45) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 46: grd1.Text = "請假、出差核准後通知人員"
grd1.ColWidth(46) = 2000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 47: grd1.Text = "B0124"
grd1.ColWidth(47) = 0
grd1.CellAlignment = flexAlignLeftCenter
'2017/8/21 END
grd1.Visible = True
End Sub

'Add By Sindy 2022/5/30
Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd2, x, y, nCol, nRow
grd2.col = nCol
grd2.row = nRow
End Sub
Private Sub GRD2_SelChange()
'GRD2.Visible = False
If grd2.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow2 > 0 Then
      Call GetSelChage2(CInt(dblPrevRow2))
   End If
   '目前資料列反白
   grd2.col = grd2.MouseCol
   grd2.row = grd2.MouseRow
   dblPrevRow2 = grd2.row
   Call GetSelChage2(grd2.row)
End If
'GRD2.Visible = True
End Sub
Private Sub grd2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If grd2.MouseRow <> 0 Then
      If iRow <> grd2.MouseRow Or iCol <> grd2.MouseCol Then
         If grd2.MouseCol = 46 Then '請假、出差核准後通知人員
            CreateToolTip GetHWndForToolTip(grd2), grd2.TextMatrix(grd2.MouseRow, grd2.MouseCol)
         End If
         iRow = grd2.MouseRow
         iCol = grd2.MouseCol
      End If
   End If
End Sub
'2022/5/30 END

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd1, x, y, nCol, nRow
grd1.col = nCol
grd1.row = nRow
End Sub

Private Sub grd1_SelChange()
'grd1.Visible = False
If grd1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      Call GetSelChage(CInt(dblPrevRow))
   End If
   '目前資料列反白
   grd1.col = grd1.MouseCol
   grd1.row = grd1.MouseRow
   dblPrevRow = grd1.row
   Call GetSelChage(grd1.row)
End If
'grd1.Visible = True
End Sub

'Add By Sindy 2021/7/12
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If grd1.MouseRow <> 0 Then
      If iRow <> grd1.MouseRow Or iCol <> grd1.MouseCol Then
         If grd1.MouseCol = 46 Then '請假、出差核准後通知人員
            CreateToolTip GetHWndForToolTip(grd1), grd1.TextMatrix(grd1.MouseRow, grd1.MouseCol)
         End If
         iRow = grd1.MouseRow
         iCol = grd1.MouseCol
      End If
   End If
End Sub

Private Sub txtB0101_GotFocus()
   InverseTextBox txtB0101
End Sub

Private Sub txtB0101_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtB0101_LostFocus()
   If txtB0101 <> "" Then
      txtB0101_2 = GetStaffName(txtB0101, True)
   End If
End Sub

Private Sub txtB0101_Validate(Cancel As Boolean)
   Label2(0).Visible = False
   If txtB0101.Text = "" Then txtB0101_2 = ""
   
   If txtB0101 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffID(txtB0101) Then
         Call txtB0101_GotFocus
         Cancel = True
         Exit Sub
      End If
      txtB0101_2 = GetStaffName(txtB0101, True)
      If txtB0101_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call txtB0101_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否離職
      If ChkStaffST04(txtB0101, False) = True Then
         Label2(0).Visible = True
      End If
   End If
End Sub

Private Sub cmdPrint_Click()
Dim strCon As String
Dim str99SQL As String 'Add By Sindy  2022/5/3

strCon = ""
If txtData(0) <> "" Then
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= 新部門啟用日 Then
      strCon = strCon & "and s1.ST93>='" & txtData(0) & "' "
   Else
   '2023/12/20 END
      strCon = strCon & "and s1.ST03>='" & txtData(0) & "' "
   End If
End If
If txtData(1) <> "" Then
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= 新部門啟用日 Then
      strCon = strCon & "and s1.ST93<='" & txtData(1) & "' "
   Else
   '2023/12/20 END
      strCon = strCon & "and s1.ST03<='" & txtData(1) & "' "
   End If
End If
'Add By Sindy 2022/5/3
If txtData(2) <> "" Then
   strCon = strCon & "and s1.ST06>='" & txtData(2) & "' "
End If
If txtData(3) <> "" Then
   strCon = strCon & "and s1.ST06<='" & txtData(3) & "' "
End If
'2022/5/3 END
If txtB0101 <> "" Then
   'Modify By Sindy 2024/2/22 + ,B0128,B0129,B0124
   strCon = strCon & "and (B0101='" & txtB0101 & "' or B0102='" & txtB0101 & "' or B0103='" & txtB0101 & "' or B0104='" & txtB0101 & "' or B0105='" & txtB0101 & "' or B0106='" & txtB0101 & "' or B0107='" & txtB0101 & "' or B0108='" & txtB0101 & "' or B0109='" & txtB0101 & "' or B0110='" & txtB0101 & "' or B0111='" & txtB0101 & "' or B0117='" & txtB0101 & "' or B0119='" & txtB0101 & "' or B0121='" & txtB0101 & "' or B0123='" & txtB0101 & "' or B0128='" & txtB0101 & "' or B0129='" & txtB0101 & "' or instr(B0124,'" & txtB0101 & "')>0) "
End If

Screen.MousePointer = vbHourglass

'Modify By Sindy 2021/1/26
'Set Printer = Printers(Combo1.ListIndex)
PUB_RestorePrinter Combo1
'2021/1/26 END
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
'Modify By Sindy 2016/5/27 簽核天數99時,該審核主管及天數都不要顯示
'Modified by Lydia 2017/03/28 ST14改成多個編號 s1.st14<>'99997'=> instr(s1.st14,'99997')=0
'Modify By Sindy 2022/5/3 可選擇99天要不要顯示
If Check1.Visible = True And Check1.Value = 1 Then
   str99SQL = ",s8.ST02||'('||B0112||')',s9.ST02||'('||B0113||')' " & _
              ",s10.ST02||'('||B0114||')',s11.ST02||'('||B0115||')' "
Else
   str99SQL = ",decode(B0112,'99','',s8.ST02||decode(nvl(B0112,''),'','','('||B0112||')')),decode(B0113,'99','',s9.ST02||decode(nvl(B0113,''),'','','('||B0113||')')) " & _
              ",decode(B0114,'99','',s10.ST02||decode(nvl(B0114,''),'','','('||B0114||')')),decode(B0115,'99','',s11.ST02||decode(nvl(B0115,''),'','','('||B0115||')')) "
End If
'2022/5/3 END
'Modify By Sindy 2023/12/20
m_str = "select A0922,s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02 " & _
        ",s17.ST02||decode(nvl(B0116,''),'','','('||B0116||')'),s19.ST02||decode(nvl(B0118,''),'','','('||B0118||')') " & _
        ",s21.ST02||decode(nvl(B0120,''),'','','('||B0120||')'),s23.ST02||decode(nvl(B0122,''),'','','('||B0122||')') " & _
        str99SQL & _
        ",GETSTAFFNAMELIST(replace(B0124,';',',')) as B0124Name,GETSTAFFNAMELIST(B0128) as B0128Name,GETSTAFFNAMELIST(B0129) as B0129Name " & _
        "from ABS001,ACC090NEW,Staff s1,Staff s2,Staff s3,Staff s4,Staff s5,Staff s6,Staff s7 " & _
        ",Staff s8,Staff s9,Staff s10,Staff s11 " & _
        ",Staff s17,Staff s19,Staff s21,Staff s23 " & _
        "where s1.ST01=B0101(+) and s1.ST93=A0921(+) " & _
        "and (instr(s1.st14,'99997')=0 or s1.ST14 is null) and substr(s1.ST01,1,1) in(" & ST01CodeNum1 & ") and substr(s1.ST01,4,1)<>'9' and s1.st01 not in('60000','96029','96030','86026','67004','68007','63001') " & _
        "and B0102=s2.ST01(+) and B0103=s3.ST01(+) and B0104=s4.ST01(+) and B0105=s5.ST01(+) and B0106=s6.ST01(+) and B0107=s7.ST01(+) " & _
        "and B0108=s8.ST01(+) and B0109=s9.ST01(+) and B0110=s10.ST01(+) and B0111=s11.ST01(+) " & _
        "and B0117=s17.ST01(+) and B0119=s19.ST01(+) and B0121=s21.ST01(+) and B0123=s23.ST01(+) " & strCon & _
        "and (s1.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) " & _
        "order by s1.ST93,B0101 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   With m_rs
      m_rs.MoveFirst
      
      '預設值
      iLine = 1
      strType = "" '切頁條件
      Do While Not m_rs.EOF
         For m_i = 1 To 19 '16
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = Left(CheckStr(m_rs.Fields(0)), 5)
         strTemp(2) = CheckStr(m_rs.Fields(1))
         strTemp(3) = CheckStr(m_rs.Fields(2))
         strTemp(4) = CheckStr(m_rs.Fields(3))
         strTemp(5) = CheckStr(m_rs.Fields(4))
         strTemp(6) = CheckStr(m_rs.Fields(5))
         strTemp(7) = CheckStr(m_rs.Fields(6))
         strTemp(8) = CheckStr(m_rs.Fields(7))
         strTemp(9) = CheckStr(m_rs.Fields(8))
         strTemp(10) = CheckStr(m_rs.Fields(9))
         strTemp(11) = CheckStr(m_rs.Fields(10))
         strTemp(12) = CheckStr(m_rs.Fields(11))
         strTemp(13) = CheckStr(m_rs.Fields(12))
         strTemp(14) = CheckStr(m_rs.Fields(13))
         strTemp(15) = CheckStr(m_rs.Fields(14))
         strTemp(16) = CheckStr(m_rs.Fields(15))
         'Add By Sindy 2024/2/22
         strTemp(17) = CheckStr(m_rs.Fields("B0124Name"))
         strTemp(18) = CheckStr(m_rs.Fields("B0128Name"))
         strTemp(19) = CheckStr(m_rs.Fields("B0129Name"))
         '2024/2/22 END
         If iLine > 36 Or iLine = 1 Then
            If strType <> "" Then Printer.NewPage
            iLine = 1
            PrintTitle '列印表頭
         End If
         
         PrintDetail '列印明細
         
         strType = CheckStr(m_rs.Fields(0))
         m_rs.MoveNext
      Loop
   End With
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "註：職代(1-1)(1-2)：第一組職代　職代(2-1)(2-2)：第二組職代　職代(3-1)(3-2)：第三組職代"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　主管括號內的數字是指幾天(不含)以上須經過該主管審核"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　案職代(1-1)(1-2)：案件第一組職代　案職代(2-1)(2-2)：案件第二組職代　＜註：(1)台灣案(2)非台灣案＞"
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Printer.EndDoc
Print_Spec 'Add By Sindy 2022/5/31

PUB_RestorePrinter strPrinter 'Modify By Sindy 2021/1/26
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 14
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("職代/審核主管關聯資料表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "職代/審核主管關聯資料表"

Printer.Font.Size = 10
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-1)"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-2)"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-1)"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-2)"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-1)"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-2)"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(1-1)"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(1-2)"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(2-1)"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(2-2)"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iLine * 300
Printer.Print "主管1"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iLine * 300
Printer.Print "主管2"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iLine * 300
Printer.Print "主管3"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iLine * 300
Printer.Print "主管4"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(255, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 300
PLeft(2) = 1500
PLeft(3) = 2500
PLeft(4) = 3500
PLeft(5) = 4500
PLeft(6) = 5500
PLeft(7) = 6500
PLeft(8) = 7500
PLeft(9) = 8500
PLeft(10) = 9500
PLeft(11) = 10500
PLeft(12) = 11500
PLeft(13) = 12500
PLeft(14) = 13500
PLeft(15) = 14500
PLeft(16) = 15500
End Sub

Sub PrintDetail()
Dim i As Integer
   
   For i = 1 To 16
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
   iLine = iLine + 1
   'Add Sindy 2024/2/22
   For i = 17 To 19
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      If i = 17 Then
         strExc(10) = "　請假、出差核准後通知人員編號："
      ElseIf i = 18 Then
         strExc(10) = "　居家職代(1)："
      Else
         strExc(10) = "　居家職代(2)："
      End If
      If strTemp(i) <> "" Then
         Printer.Print strExc(10) & strTemp(i)
         iLine = iLine + 1
         If iLine > 36 Or iLine = 1 Then
            Printer.NewPage
            iLine = 1
            PrintTitle '列印表頭
         End If
      End If
   Next i
   '2024/2/22 END
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   InverseTextBox txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If txtData(Index) <> "" And txtData(Index + 1) = "" Then
         txtData(Index + 1) = txtData(Index)
      End If
   ElseIf Index = 1 Then
      If txtData(Index) <> "" And txtData(Index - 1) = "" Then
         txtData(Index - 1) = txtData(Index)
      End If
      If RunNick(txtData(Index - 1), txtData(Index)) Then
         Call Txtdata_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2022/5/31
Private Sub Print_Spec()
Dim strCon As String

   strCon = ""
   If txtData(0) <> "" Then
      'Modify By Sindy 2023/12/20
'      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & "and s1.ST93>='" & txtData(0) & "' "
'      Else
'      '2023/12/20 END
'         strCon = strCon & "and s0.ST03>='" & txtData(0) & "' "
'      End If
   End If
   If txtData(1) <> "" Then
      'Modify By Sindy 2023/12/20
'      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & "and s1.ST93<='" & txtData(1) & "' "
'      Else
'      '2023/12/20 END
'         strCon = strCon & "and s0.ST03<='" & txtData(1) & "' "
'      End If
   End If
   If txtData(2) <> "" Then
      strCon = strCon & "and s1.ST06>='" & txtData(2) & "' "
   End If
   If txtData(3) <> "" Then
      strCon = strCon & "and s1.ST06<='" & txtData(3) & "' "
   End If
   If txtB0101 <> "" Then
      strCon = strCon & "and (B0201='" & txtB0101 & "' or B0202='" & txtB0101 & "' or B0203='" & txtB0101 & "'" & _
                         " or B0204='" & txtB0101 & "' or B0205='" & txtB0101 & "' or B0206='" & txtB0101 & "'" & _
                         " or B0207='" & txtB0101 & "' or B0208='" & txtB0101 & "')"
   End If
   
   Screen.MousePointer = vbHourglass
   
   Printer.EndDoc
   Printer.Orientation = 2 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2023/12/20
   'OLD: decode(length(B0208),3,A0922,s0.ST02)
'   m_str = "select decode(B0201,B0208,'',A0922) A0922,decode(length(B0208),3,A0922,decode(B0209,'1',s0.ST02,decode(B0208,B0201,'',s0.ST02))),s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02" & _
'                  ",s0.ST04,s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04" & _
'                  ",B0201,B0202,B0203,B0204,B0205,B0206,B0207,B0208,decode(B0209,'1','人事','2','案件',B0209)" & _
'            " from ABS002,ACC090NEW,STAFF s0,STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7" & _
'            " where s0.ST01=B0208(+) and B0201 is not null" & _
'            " and s0.ST93=A0921(+) and B0201=s1.ST01(+)" & _
'            " and B0202=s2.ST01(+) and B0203=s3.ST01(+)" & _
'            " and B0204=s4.ST01(+) and B0205=s5.ST01(+)" & _
'            " and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & strCon & _
'            " and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0208 and sc02=(select max(sc02) from Staff_Change where sc01=B0208)))" & _
'            " order by A0921,B0208"
   'Modify By Sindy 2025/3/6 再調整SQL
   m_str = "select decode(B0201,B0208,'',A0922) A0922,decode(B0209,'1',s0.ST02,decode(B0208,B0201,'',s0.ST02))" & _
                  ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02" & _
                  ",s0.ST04,s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04" & _
                  ",B0201,B0202,B0203,B0204,B0205,B0206,B0207,B0208,decode(B0209,'1','人事','2','案件',B0209),A0921" & _
            " from ABS002,ACC090NEW,STAFF s0,STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7" & _
            " where s0.ST01=B0208 and s0.ST01 is not null and length(B0208)=5" & _
            " and s0.ST93=A0921 and B0201=s1.ST01(+)" & _
            " and B0202=s2.ST01(+) and B0203=s3.ST01(+)" & _
            " and B0204=s4.ST01(+) and B0205=s5.ST01(+)" & _
            " and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & strCon & _
            " and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0208 and sc02=(select max(sc02) from Staff_Change where sc01=B0208)))"
   m_str = m_str & " union " & _
            "select A0922,A0922" & _
                  ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02" & _
                  ",'',s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04,s7.ST04" & _
                  ",B0201,B0202,B0203,B0204,B0205,B0206,B0207,B0208,decode(B0209,'1','人事','2','案件',B0209),A0921" & _
            " from ABS002,ACC090NEW,STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7" & _
            " where A0921=B0208 and length(B0208)<>5" & _
            " and B0201=s1.ST01(+)" & _
            " and B0202=s2.ST01(+) and B0203=s3.ST01(+)" & _
            " and B0204=s4.ST01(+) and B0205=s5.ST01(+)" & _
            " and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & strCon
   m_str = m_str & " order by A0921,B0208"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "" '切頁條件
         Do While Not m_rs.EOF
            For m_i = 1 To 17 '16
                strTemp(m_i) = ""
            Next m_i
            strTemp(1) = Left(CheckStr(m_rs.Fields(0)), 5)
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields(2))
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            strTemp(7) = CheckStr(m_rs.Fields(6))
            strTemp(8) = CheckStr(m_rs.Fields(7))
            strTemp(9) = CheckStr(m_rs.Fields(8))
            strTemp(10) = CheckStr(m_rs.Fields(9))
            strTemp(11) = CheckStr(m_rs.Fields(10))
            strTemp(12) = CheckStr(m_rs.Fields(11))
            strTemp(13) = CheckStr(m_rs.Fields(12))
            strTemp(14) = CheckStr(m_rs.Fields(13))
            strTemp(15) = CheckStr(m_rs.Fields(14))
            strTemp(16) = CheckStr(m_rs.Fields(15))
            strTemp(10) = CheckStr(m_rs.Fields(25)) 'Add By Sindy 2023/5/4
            
            If iLine > 36 Or iLine = 1 Then
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle2 '列印表頭
            End If
            
            PrintDetail2 '列印明細
            
            strType = CheckStr(m_rs.Fields(0))
            m_rs.MoveNext
         Loop
      End With
      iLine = iLine + 1
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "註：職代(1-1)(1-2)：第一組職代　職代(2-1)(2-2)：第二組職代　職代(3-1)(3-2)：第三組職代"
'   Else
'      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
'      Screen.MousePointer = vbDefault
'      Exit Sub
   End If
   Printer.EndDoc
End Sub

'Add By Sindy 2022/5/31
Sub PrintTitle2()
GetPleft2

Printer.Font.Size = 14
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("簽核主管特殊對象的簽核職代表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "簽核主管特殊對象的簽核職代表"

Printer.Font.Size = 10
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "簽核主管"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-1)"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-2)"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-1)"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-2)"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-1)"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-2)"
'Add By Sindy 2023/5/4
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "簽核種類"
'2023/5/4 END

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(255, "-")

iLine = iLine + 1
End Sub

'Add By Sindy 2022/5/31
Sub GetPleft2()
PLeft(1) = 300
PLeft(2) = 1500
PLeft(3) = 2500
PLeft(4) = 3500
PLeft(5) = 4500
PLeft(6) = 5500
PLeft(7) = 6500
PLeft(8) = 7500
PLeft(9) = 8500
PLeft(10) = 9500 'Add By Sindy 2023/5/4
End Sub

'Add By Sindy 2022/5/31
Sub PrintDetail2()
Dim i As Integer
   
   For i = 1 To 10 '9
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
   iLine = iLine + 1
End Sub
