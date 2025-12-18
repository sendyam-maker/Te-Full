VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100101_24 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利法律案件關聯查詢"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9435
   Begin VB.TextBox txtCodeP 
      Height          =   300
      Index           =   4
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   3
      Top             =   90
      Width           =   345
   End
   Begin VB.TextBox txtCodeP 
      Height          =   300
      Index           =   3
      Left            =   2685
      MaxLength       =   1
      TabIndex        =   2
      Top             =   90
      Width           =   225
   End
   Begin VB.TextBox txtCodeP 
      Height          =   300
      Index           =   2
      Left            =   1845
      MaxLength       =   6
      TabIndex        =   1
      Top             =   90
      Width           =   765
   End
   Begin VB.TextBox txtCodeP 
      Height          =   300
      Index           =   1
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   0
      Top             =   90
      Width           =   465
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   4
      Left            =   3015
      MaxLength       =   2
      TabIndex        =   9
      Top             =   840
      Width           =   345
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   3
      Left            =   2685
      MaxLength       =   1
      TabIndex        =   8
      Top             =   840
      Width           =   225
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   2
      Left            =   1845
      MaxLength       =   6
      TabIndex        =   7
      Top             =   840
      Width           =   765
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   1
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   6
      Top             =   840
      Width           =   465
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7584
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8484
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4065
      Left            =   135
      TabIndex        =   14
      Top             =   1560
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   7170
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1305
      TabIndex        =   10
      Top             =   1200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2685
      TabIndex        =   11
      Top             =   1200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1305
      TabIndex        =   4
      Top             =   465
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   2685
      TabIndex        =   5
      Top             =   465
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Line Line4 
      X1              =   1305
      X2              =   3015
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   1605
      X2              =   3315
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line1 
      X1              =   2430
      X2              =   2700
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   2430
      X2              =   2700
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法務案收文日"
      Height          =   180
      Left            =   180
      TabIndex        =   19
      Top             =   1260
      Width           =   1080
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   5430
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利案號"
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "專利案收文日"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   525
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法務案號"
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   900
      Width           =   720
   End
End
Attribute VB_Name = "frm100101_24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/07 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Morgan 2011/5/27
Option Explicit

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .Cols = 11: .FixedRows = 1: .FixedCols = 0
         .MergeCol(0) = True
         .MergeCells = flexMergeRestrictColumns
      End If
      .row = 0
      .col = 0: .ColWidth(.col) = 1215: .Text = "法務案號"
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 850: .Text = "收文日"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1845: .Text = "案件性質"
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      
      .col = 3: .ColWidth(.col) = 1095: .Text = "專利案號 "
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 850: .Text = "收文日"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1275: .Text = "案件性質"
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 6: .ColWidth(.col) = 990: .Text = "分配點數"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 7: .ColWidth(.col) = 720: .Text = "總點數"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      
      For ii = 8 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      .Refresh
      .Visible = True
   End With
End Sub

Private Function doQuery() As Boolean
   Dim stCon As String, stDate As String
   
   If txtCode(1) <> "" Then
      stCon = stCon & " and c1.cp01='" & txtCode(1) & "'"
   End If
   If txtCode(2) <> "" Then
      stCon = stCon & " and c1.cp02='" & txtCode(2) & "'"
   End If
   If txtCode(3) <> "" Then
      stCon = stCon & " and c1.cp03='" & txtCode(3) & "'"
   End If
   If txtCode(4) <> "" Then
      stCon = stCon & " and c1.cp04='" & txtCode(4) & "'"
   End If
   
   stDate = Replace(Replace(MaskEdBox1, "/", ""), "_", "")
   If stDate <> "" Then
      If ChkDate(stDate) = False Then
         MaskEdBox1.SetFocus
         Exit Function
      Else
         stCon = stCon & " and c1.cp05>=" & DBDATE(stDate)
      End If
   End If
   
   stDate = Replace(Replace(MaskEdBox2, "/", ""), "_", "")
   If stDate <> "" Then
      If ChkDate(stDate) = False Then
         MaskEdBox2.SetFocus
         Exit Function
      Else
         stCon = stCon & " and c1.cp05<=" & DBDATE(stDate)
      End If
   End If
   
   stDate = Replace(Replace(MaskEdBox3, "/", ""), "_", "")
   If stDate <> "" Then
      If ChkDate(stDate) = False Then
         MaskEdBox3.SetFocus
         Exit Function
      Else
         stCon = stCon & " and c2.cp05>=" & DBDATE(stDate)
      End If
   End If
   
   stDate = Replace(Replace(MaskEdBox4, "/", ""), "_", "")
   If stDate <> "" Then
      If ChkDate(stDate) = False Then
         MaskEdBox4.SetFocus
         Exit Function
      Else
         stCon = stCon & " and c2.cp05<=" & DBDATE(stDate)
      End If
   End If
   
   If txtCodeP(1) <> "" Then
      stCon = stCon & " and c2.cp01='" & txtCodeP(1) & "'"
   End If
   If txtCodeP(2) <> "" Then
      stCon = stCon & " and c2.cp02='" & txtCodeP(2) & "'"
   End If
   If txtCodeP(3) <> "" Then
      stCon = stCon & " and c2.cp03='" & txtCodeP(3) & "'"
   End If
   If txtCodeP(4) <> "" Then
      stCon = stCon & " and c2.cp04='" & txtCodeP(4) & "'"
   End If


On Error GoTo ErrHnd

   strExc(0) = "select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) 法務案號" & _
      ",substrb(sqldatet(c1.cp05),1,10) 收文日,substrb(m1.cpm03,1,18) 案件性質" & _
      ",c2.cp01||'-'||c2.cp02||decode(c2.cp03||c2.cp04,'000','','-'||c2.cp03||'-'||c2.cp04) 專利案號" & _
      ",substrb(sqldatet(c2.cp05),1,10) 收文日,substrb(m2.cpm03,1,12) 案件性質" & _
      ",''||(a0n03/1000) 分配點數,c1.cp18/1 總點數" & _
      " from acc0n0,caseprogress c1,casepropertymap m1,lawcase" & _
      ",caseprogress c2,casepropertymap m2,patent" & _
      " Where a0n01 <> a0n02" & _
      " and c1.cp09(+)=a0n01 and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10" & _
      " and lc01(+)=c1.cp01 and lc02(+)=c1.cp02 and lc03(+)=c1.cp03 and lc04(+)=c1.cp04" & _
      " and c2.cp09(+)=a0n02 and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10" & _
      " and pa01(+)=c2.cp01 and pa02(+)=c2.cp02 and pa03(+)=c2.cp03 and pa04(+)=c2.cp04" & stCon
      
   'Added by Morgan 2021/3/16 +案源資料
   strExc(0) = strExc(0) & " union select max(lc01||'-'||lc02||decode(lc03||lc04,'000','','-'||lc03||'-'||lc04)) 法務案號" & _
      ",max(substrb(sqldatet(c1.cp05),1,10)) 收文日,max(substrb(m1.cpm03,1,18)) 案件性質" & _
      ",max(pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)) 專利案號" & _
      ",max(substrb(sqldatet(c2.cp05),1,10)) 收文日,max(substrb(m2.cpm03,1,12)) 案件性質" & _
      ",'案源' 分配點數,sum(c3.cp18) 總點數" & _
      " from LAWOFFICESOURCE,caseprogress c1,casepropertymap m1,lawcase,caseprogress c2,casepropertymap m2,patent,caseprogress c3" & _
      " where los02 like 'B%' and c1.cp09(+)=los06 and c2.cp09(+)=los01" & _
      " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10" & _
      " and lc01(+)=c1.cp01 and lc02(+)=c1.cp02 and lc03(+)=c1.cp03 and lc04(+)=c1.cp04" & _
      " and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10 and pa01(+)=c2.cp01" & _
      " and pa02(+)=c2.cp02 and pa03(+)=c2.cp03 and pa04(+)=c2.cp04 and pa01 is not null" & _
      " and c3.cp162(+)=los15" & stCon & " group by los15"
   'end 2021/3/16
   
   strExc(0) = strExc(0) & " order by 1,2,4,5"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set grdDataList.Recordset = RsTemp.Clone
      Call SetDataListWidth(True)
   Else
      MsgBox "無符合資料！", vbInformation
   End If
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Dim iOpt As Integer, oOption As OptionButton
   Screen.MousePointer = vbHourglass
   doQuery
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set frm100101_24 = Nothing
End Sub

Private Sub MaskEdBoxInverse(pBox As MaskEdBox)
   pBox.SelStart = 0
   pBox.SelLength = Len(pBox.Text)
End Sub

Private Sub MaskEdBox1_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox1
End Sub

Private Sub MaskEdBox2_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox2
End Sub

Private Sub MaskEdBox3_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox3
End Sub

Private Sub MaskEdBox4_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox4
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
   CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCodeP_GotFocus(Index As Integer)
   TextInverse txtCodeP(Index)
   CloseIme
End Sub

Private Sub txtCodeP_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
