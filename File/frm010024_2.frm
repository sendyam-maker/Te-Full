VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010024_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "補輸發文字號"
   ClientHeight    =   3990
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6675
   Begin VB.TextBox txtCP84 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4545
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   900
      Width           =   1410
   End
   Begin VB.TextBox txtCP09 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4545
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   570
      Width           =   1410
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   570
      Width           =   1410
   End
   Begin VB.TextBox txtCP10 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   900
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4425
      TabIndex        =   0
      Top             =   45
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   45
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2370
      Left            =   60
      TabIndex        =   12
      Top             =   1560
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   4180
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   255
      Left            =   1335
      TabIndex        =   7
      Top             =   1230
      Width           =   1410
      VariousPropertyBits=   679493661
      BackColor       =   -2147483633
      Size            =   "8555;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "補發總收文號："
      Height          =   255
      Left            =   3150
      TabIndex        =   11
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "有無規費："
      Height          =   255
      Left            =   3150
      TabIndex        =   10
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "進度備註："
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lblFund 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
End
Attribute VB_Name = "frm010024_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 txtCP64/GrdDataList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Public strCP09 As String      '總收文號(傳入)
Public strCP28 As String      '發文字號(回傳)
Public strCP124 As String    '發文室發文日(回傳)
Public BolOk As Boolean     'True: 確定  False: 取消


Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "總收文號"
grdDataList.ColWidth(1) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "案件性質"
grdDataList.ColWidth(2) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "發文字號"
grdDataList.ColWidth(3) = 900
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "有無規費"
grdDataList.ColWidth(4) = 600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "發文室發文日"
grdDataList.ColWidth(5) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "發文室發文時間"
grdDataList.ColWidth(6) = 1300
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean, j As Integer
   
   '確定
   If Index = 0 Then
      Cancel = True
      For j = 1 To grdDataList.Rows - 1
         If Trim(grdDataList.TextMatrix(j, 0)) = "V" And Trim(grdDataList.TextMatrix(j, 3)) <> "" Then
            strCP28 = Trim(grdDataList.TextMatrix(j, 3))   '發文字號
            strCP124 = ChangeTStringToWString(ChangeTDateStringToTString(Trim(grdDataList.TextMatrix(j, 5)))) '發文室發文日
            Cancel = False
         End If
      Next j
      If Cancel = True Then
         MsgBox "請點選一筆已發文資料！", vbExclamation + vbOKOnly, Me.Caption
         Exit Sub
      End If
      BolOk = True
      
   '回前畫面(取消)
   Else
      strCP28 = ""
      strCP124 = ""
      BolOk = False
   End If
   Me.Hide
End Sub

Public Function CheckShowList() As Boolean
Dim strSql As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim dblCP27 As Double
Dim intIdx As Integer, i As Integer
   
   CheckShowList = False
   
   '取得補發總收文號
   strSql = "SELECT * FROM CaseProgress,casepropertymap WHERE CP09='" & Trim(strCP09) & "' and CP01=CPM01(+) and CP10=CPM02(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCP01 = "" & RsTemp("cp01")
      strCP02 = "" & RsTemp("cp02")
      strCP03 = "" & RsTemp("cp03")
      strCP04 = "" & RsTemp("cp04")
      dblCP27 = "" & RsTemp("cp27")
      txtCaseNo = RsTemp("cp01") & "-" & RsTemp("cp02") & IIf(RsTemp("cp03") & RsTemp("cp04") = "000", "", "-" & RsTemp("cp03") & "-" & RsTemp("cp04"))
      txtCP10 = RsTemp("cpm03")
      txtCP64 = "" & RsTemp("cp64")
      txtCP09 = strCP09
      If IsNull(RsTemp("cp84")) Or RsTemp("cp84") <= 0 Then
         txtCP84 = "無"
      Else
         txtCP84 = "有"
      End If
      CheckShowList = True
   End If
   
   '取得當天相同案號已發文資料
   strSql = "SELECT ' ' AS V,CP09,CPM03,CP28,decode(CP84,0,'無',null,'無','有'),sqldateT(CP124),SqlTime(CP125) FROM CaseProgress,casepropertymap " & _
                  "WHERE CP01=CPM01(+) and CP10=CPM02(+) " & _
                  "and CP27=" & dblCP27 & " and CP124>0 " & _
                  "and CP01='" & strCP01 & "' and CP02='" & strCP02 & "' and CP03='" & strCP03 & "' and CP04='" & strCP04 & "' "
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   'GrdDataList.FixedCols = 0
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
       CheckShowList = True
   Else
       'ShowNoData
       grdDataList.Clear
   End If
   SetDataListWidth
   'GrdDataList.FixedCols = 3
   CheckOC
   
'   '若只有一筆資料, 則直接設定為點選此筆資料
'   With Me.GrdDataList
'      If .Rows = 2 Then
'         .row = 1
'         .col = 1
'         If .Text <> "" Then
'           .Visible = False
'           .row = 1
'           .col = 0
'           .Text = "V"
'           For i = 0 To .Cols - 1
'               .col = i
'               .CellBackColor = &HFFC0C0
'               If i <= 2 Then
'                 GrdDataList.CellBackColor = &H8000000F
'               End If
'           Next i
'           .Visible = True
'         End If
'      End If
'   End With
   Screen.MousePointer = vbDefault
   
   BolOk = True
   Exit Function
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub GrdDataList_Click()
Dim j As Integer, i As Integer
   
   grdDataList.Visible = False
   
   For j = 1 To grdDataList.Rows - 1
      grdDataList.row = j
      grdDataList.col = 0
      grdDataList.Text = ""
      For i = 0 To grdDataList.Cols - 1
           grdDataList.col = i
           grdDataList.CellBackColor = QBColor(15)
           If i <= 2 Then
              grdDataList.CellBackColor = &H8000000F
           End If
      Next i
   Next j
   
   '勾選
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
           grdDataList.Text = ""
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = QBColor(15)
               If i <= 2 Then
                  grdDataList.CellBackColor = &H8000000F
               End If
          Next i
      Else
           grdDataList.Text = "V"
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
               If i <= 2 Then
                  grdDataList.CellBackColor = &H8000000F
               End If
           Next i
      End If
   End If
   
   grdDataList.Visible = True
End Sub
