VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100120_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以發明人查詢"
   ClientHeight    =   5920
   ClientLeft      =   -270
   ClientTop       =   980
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5920
   ScaleWidth      =   9320
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1200
      Width           =   1300
   End
   Begin VB.OptionButton Option2 
      Caption         =   "發明人ID："
      Height          =   180
      Index           =   2
      Left            =   6480
      TabIndex        =   26
      Top             =   1290
      Width           =   1200
   End
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   1740
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   1665
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2535
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   350
      Left            =   1584
      TabIndex        =   23
      Top             =   840
      Width           =   3120
      Begin VB.OptionButton Option1 
         Caption         =   "日文"
         Height          =   180
         Index           =   2
         Left            =   2250
         TabIndex        =   4
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "英文"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "中文"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Value           =   -1  'True
         Width           =   732
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "發明人名稱："
      Height          =   180
      Index           =   1
      Left            =   132
      TabIndex        =   1
      Top             =   1260
      Width           =   1410
   End
   Begin VB.OptionButton Option2 
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   19
      Top             =   540
      Value           =   -1  'True
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1650
      MaxLength       =   9
      TabIndex        =   0
      Top             =   510
      Width           =   1932
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2210
      Width           =   885
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1885
      Width           =   885
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   2772
   End
   Begin VB.TextBox Text7 
      Height          =   264
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2228
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1885
      Width           =   852
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3870
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發明人資料(&I)"
      Height          =   400
      Index           =   2
      Left            =   5964
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   70
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件資料(&B)"
      Height          =   400
      Index           =   1
      Left            =   7284
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人資料(&A)"
      Height          =   400
      Index           =   0
      Left            =   4650
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   90
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8508
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   70
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3000
      Left            =   30
      TabIndex        =   24
      Top             =   2880
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5292
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   1590
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
      VariousPropertyBits=   671107099
      Size            =   "8488;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl_Name 
      Height          =   255
      Left            =   3630
      TabIndex        =   27
      Top             =   540
      Width           =   5505
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "9710;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2211
      X2              =   2040
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2040
      X2              =   2211
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                  (ALL：全部)"
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   1620
      Width           =   4860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Left            =   150
      TabIndex        =   21
      Top             =   1945
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   150
      TabIndex        =   20
      Top             =   2270
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "是否含來函資料：           （N：不含）"
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   2580
      Width           =   2955
   End
End
Attribute VB_Name = "frm100120_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Text2、Lbl_Name
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, s As Integer
Dim StrTag As String, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/13 只記錄於此Form


Private Sub SetDataListWidth()
   grdDataList.row = 0
   grdDataList.col = 0
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   
   grdDataList.col = 1: grdDataList.Text = "發明人編號"
   grdDataList.ColWidth(1) = 1100
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "發明人名稱"
   grdDataList.ColWidth(2) = 4600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "國籍"
   grdDataList.ColWidth(3) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                    grdDataList.col = j
                    grdDataList.CellBackColor = QBColor(15)
              Next j
              If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
              End If
              Dim Str01 As String
              grdDataList.col = 1
              Screen.MousePointer = vbHourglass
              frm100101_11.Show
              frm100101_11.Tag = Left(Trim(grdDataList.Text), 8) + "0"
              frm100101_11.StrMenu
              Screen.MousePointer = vbDefault
              Me.Enabled = True
              Exit Sub
            End If
            Next i
            Me.Enabled = True
      Case 1
            If PUB_CheckKeyInDate(Me.Text4) = -1 Then
               Me.Text4.SetFocus
               Text4_GotFocus
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text5) = -1 Then
               Me.Text5.SetFocus
               Text5_GotFocus
               Exit Sub
            End If
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
               grdDataList.Text = " "
               For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
              Next j
              If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
              End If
              grdDataList.col = 1
              Screen.MousePointer = vbHourglass
              frm100120_2.Show
              'frm100120_2.MousePointer = vbHourglass
              frm100120_2.Tag = grdDataList.Text
              frm100120_2.StrMenu
              Screen.MousePointer = vbDefault
              Me.Enabled = True
              Exit Sub
            End If
            Next i
            Me.Enabled = True
      Case 2
           Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
              grdDataList.col = 0
              grdDataList.row = i
              If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 0
                   grdDataList.Text = " "
                   For j = 0 To grdDataList.Cols - 1
                    grdDataList.col = j
                    grdDataList.CellBackColor = QBColor(15)
                  Next j
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  grdDataList.col = 1
                  Screen.MousePointer = vbHourglass
                  frm100101_12.Show
                  frm100101_12.Tag = grdDataList.Text
                  frm100101_12.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
              End If
            Next i
            Me.Enabled = True
      Case 3
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

'2011/12/6 add by sonia
Private Sub chk_Click()
   '若勾選所有系統類別
   If Me.Chk.Value = vbChecked Then
       Me.Text3.Text = "ALL"
   '若取消勾選所有系統類別
   Else
       Me.Text3.Text = Systemkind_g
   End If
End Sub
'2011/12/6 end

Private Sub cmdok_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text3.Text)) = 0 Then
       Me.Text3.Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
End Sub

Private Sub cmdSearch_Click()

'已發明人查詢之資料庫
'查詢 INVENTOR   只輸發明人資料
Dim s As Integer
   'Add By Cheng 2002/03/14
   ''Add By Cheng 2002/01/07
   'Text3_LostFocus
   
   If Option2(0).Value = True Then
       If Len(Trim(Text1)) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           Exit Sub
       End If
       Call Text1_LostFocus 'Modify by Amy 2014/04/22
   End If
   If Option2(1).Value = True Then
       If Len(Trim(Text2)) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           Exit Sub
       End If
   End If
   'Add by Amy 2014/04/22 +發明人ID查詢
    If Option2(2).Value = True Then
       If Len(Trim(Text9)) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           Exit Sub
       End If
   End If
   'end 2014/04/22
   '以編號或名稱
   Screen.MousePointer = vbHourglass
    'StrToSystem(ArrTemp) = ""
   grdDataList.Clear
   SetDataListWidth
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/27 清除查詢印表記錄檔欄位
   '以申請人編號查詢
   If Option2(0).Value = True Then
      'Modify By Cheng 2002/01/25
      '發明人名稱顯示方式, 若國籍為台灣則中英日,若否則英中日
      'strSQL = "SELECT ' ' AS V ,IN01||'-'||IN02 AS 發明人編號,NVL(IN04,NVL(IN05,IN06)) AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND IN01='" & Mid(GetNewFagent(Text1), 1, 8) & "' "
      '2008/9/2 MODIFY BY SONIA若國籍為台灣,大陸,香港,澳門則中英日,若否則英中日
      'strSQL = "SELECT ' ' AS V ,IN01||'-'||IN02 AS 發明人編號,DeCode( SubStr(NA01,1,2), '00', NVL(IN04,NVL(IN05,IN06)),NVL(IN05,NVL(IN04,IN06))) AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND IN01='" & Mid(GetNewFagent(Text1), 1, 8) & "' "
      strSql = "SELECT ' ' AS V ,IN01||'-'||IN02 AS 發明人編號,DeCode(NA01,'020', NVL(IN04,NVL(IN05,IN06)),'013', NVL(IN04,NVL(IN05,IN06)),'044', NVL(IN04,NVL(IN05,IN06)),DECODE(SubStr(NA01,1,2), '00', NVL(IN04,NVL(IN05,IN06)),NVL(IN05,NVL(IN04,IN06)))) AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND IN01='" & Mid(GetNewFagent(Text1), 1, 8) & "' "
      pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Text1 'Add By Sindy 2010/10/27
   '以發明人名稱查詢
   ElseIf Option2(1).Value = True Then
      '查中文
       If Option1(0).Value = True Then
          strSql = "SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN04 AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NAtION WHERE IN11=NA01(+) AND instr(IN04,'" & Text2 & "')>0 "
          pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/10/27
       Else
          '查英文
          If Option1(1).Value = True Then
              strSql = "SELECT ' ' AS V,IN01||'-'||IN02 AS 發明人編號,IN05 AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND INSTR(UPPER(IN05),'" & UCase(Text2) & "')>0 "
              pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/10/27
          Else
               '查日文
              If Option1(2).Value = True Then
                  strSql = "SELECT ' ' AS V,IN01||'-'||IN02 AS 發明人編號,IN06 AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND instr(IN06,'" & Text2 & "')>0 "
                  pub_QL05 = pub_QL05 & ";" & Option1(2).Caption 'Add By Sindy 2010/10/27
              End If
          End If
       End If
       pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Text2 'Add By Sindy 2010/10/27
   '發明人ID查詢 Add by Amy 2014/04/22
   Else
        strSql = "SELECT ' ' AS V ,IN01||'-'||IN02 AS 發明人編號,DeCode(NA01,'020', NVL(IN04,NVL(IN05,IN06)),'013', NVL(IN04,NVL(IN05,IN06)),'044', NVL(IN04,NVL(IN05,IN06)),DECODE(SubStr(NA01,1,2), '00', NVL(IN04,NVL(IN05,IN06)),NVL(IN05,NVL(IN04,IN06)))) AS 發明人名稱,NA03 AS 國籍 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND IN03='" & Text9 & "' "
        pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Text9
   End If
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/13 記錄此Form的查詢條件
   If adoRecordset.RecordCount <> 0 Then
       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/27
       grdDataList.Rows = adoRecordset.RecordCount + 1
       If Not cmdok(0).Enabled Then cmdok(0).Enabled = True
       If Not cmdok(1).Enabled Then cmdok(1).Enabled = True
       If Not cmdok(2).Enabled Then cmdok(2).Enabled = True
       '911029 nick move from down
       Set grdDataList.Recordset = adoRecordset
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/10/27
       grdDataList.Rows = 2 'Add by Amy 2014/04/22
       ShowNoData
       cmdok(0).Enabled = False
       cmdok(1).Enabled = False
       cmdok(2).Enabled = False
   End If
   '911029 nick move to up
   'Set GrdDataList.Recordset = adoRecordset
   CheckOC
   SetDataListWidth
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/13 還原此Form的查詢條件記錄
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   cmdok(0).Enabled = False
   cmdok(1).Enabled = False
   cmdok(2).Enabled = False
   'Text2.Enabled = False
   Option2(0).Value = True
   Option1(0).Enabled = False
   Option1(1).Enabled = False
   Option1(2).Enabled = False
   '2011/12/6 modify by sonia
   'Text3 = Systemkind_g
   Me.Chk.Value = vbChecked
   lbl_Name.Caption = "" 'Modify by Amy 201/04/22 +申請人名稱
   Text3 = "ALL"
   '2011/12/6 end
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
   Set frm100120_1 = Nothing
End Sub

Private Sub GrdDataList_Click()
   Me.grdDataList.Enabled = True
   grdDataList.Visible = False
   'Modify By Cheng 2002/03/14
   'grdDataList.Row = grdDataList.MouseRow
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
           grdDataList.Text = ""
           For i = 0 To grdDataList.Cols - 1
                grdDataList.col = i
                grdDataList.CellBackColor = QBColor(15)
          Next i
      
      Else
           grdDataList.Text = "V"
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
           Next i
      
      End If
   End If
   grdDataList.Visible = True
   Me.grdDataList.Enabled = True
   
End Sub

Private Sub grdDataList_SelChange()
   'Modify By Cheng 2002/03/14
   'grdDataList.Visible = False
   ''Modify By Cheng 2002/03/14
   ''grdDataList.Row = grdDataList.MouseRow
   'grdDataList.Row = grdDataList.MouseRow + 1
   'grdDataList.Col = 0
   'If grdDataList.Row <> 0 Then
   '   If grdDataList.Text = "V" Then
   '        grdDataList.Text = ""
   '        For i = 0 To grdDataList.Cols - 1
   '             grdDataList.Col = i
   '             grdDataList.CellBackColor = QBColor(15)
   '       Next i
   '
   '   Else
   '        grdDataList.Text = "V"
   '        For i = 0 To grdDataList.Cols - 1
   '            grdDataList.Col = i
   '            grdDataList.CellBackColor = &HFFC0C0
   '        Next i
   '
   '   End If
   'End If
   'grdDataList.Visible = True
End Sub

Sub StrMenu()
   grdDataList.Clear
   SetDataListWidth
   Dim strSql As String, lngCounter As Long, lngCounterI As Long
   lngCounterI = 0
   '已申請人查詢之資料庫
   '查詢 CUSTOMER   使用 LIKE
   Dim s As Integer
   If Option2(0).Value = True Then
       If Len(Trim(Text1)) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           Exit Sub
       End If
   End If
   If Option2(1).Value = True Then
       If Len(Trim(Text2)) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           Exit Sub
       End If
   End If
   '以編號或名稱
   Screen.MousePointer = vbHourglass
   If Option2(0).Value = True Then
   strSql = "SELECT ' ' AS V ,CU01||CU02 AS 申請人編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人名稱,NA03 AS 國籍 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(Text1, 6) & "000' AND CU01<='" & Left(Text1, 6) & "zzz' "
   Else
       If Option1(0).Value = True Then
          strSql = "SELECT ' ' as V,CU01||CU02 AS 申請人編號,CU04 AS 申請人名稱,NA03 AS 國籍 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND instr(CU04,'" & Text2 & "')>0 "
       Else
          If Option1(1).Value = True Then
              strSql = "SELECT ' ' AS V,CU01||CU02 AS 申請人編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 申請人名稱,NA03 AS 國籍 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND instr(cu05||' '||cu88||' '||cu89||' '||cu90,'" & Text2 & "')>0 "
          Else
              If Option1(2).Value = True Then
                  strSql = "SELECT ' ' AS V,CU01||CU02 AS 申請人編號,CU06 AS 申請人名稱,NA03 AS 國籍 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND instr(CU06,'" & Text2 & "')>0 "
              End If
          End If
       End If
   End If
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       grdDataList.Rows = adoRecordset.RecordCount + 1
       If Not cmdok(0).Enabled Then cmdok(0).Enabled = True
       If Not cmdok(1).Enabled Then cmdok(1).Enabled = True
       If Not cmdok(2).Enabled Then cmdok(2).Enabled = True
   Else
       cmdok(0).Enabled = False
       cmdok(1).Enabled = False
       cmdok(2).Enabled = False
   End If
   Set grdDataList.Recordset = adoRecordset
   CheckOC
   
   'Add By Cheng 2001/12/26
   '若查詢結果僅有一筆, 則直接點選
   If Me.grdDataList.Rows = 2 Then
      grdDataList.Visible = False
      grdDataList.row = 1
      grdDataList.col = 0
      grdDataList.Text = "V"
      For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = &HFFC0C0
      Next i
      grdDataList.Visible = True
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Option1_Click(Index As Integer)
   'Add By Cheng 2002/04/23
   Select Case Index
      Case 0, 2
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Me.Text2.IMEMode = 1
         OpenIme
      '   Do While Me.Text2.IMEMode <> 1
      '      Me.Text2.IMEMode = 1
      '   Loop
         
      Case 1
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Me.Text2.IMEMode = 2
         CloseIme
      '   Do While Me.Text2.IMEMode <> 2
      '      Me.Text2.IMEMode = 2
      '   Loop
   End Select
End Sub

Private Sub Option2_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option2(0).Value = True Then
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              'Text1.Enabled = True
              'Text2.Enabled = False
              Text1.SetFocus
              Text1_GotFocus
           End If
      Case 1
           If Option2(1).Value = True Then
              Option1(0).Enabled = True
              Option1(0).Value = True
              Option1(1).Enabled = True
              Option1(2).Enabled = True
              Text2.SetFocus
              Text2_GotFocus
           End If
      'Add by Amy 2014/04/22
      Case 2
           If Option2(2).Value = True Then
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              Text9.SetFocus
              Text9_GotFocus
           End If
      Case Else
   End Select
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2014/04/22 +顯示申請人名稱
Private Sub Text1_LostFocus()
    If Trim(Text1) = MsgText(601) Then Exit Sub
    
    If Len(Trim(Text1)) < 8 Then Text1 = Trim(Text1) & String(8 - Len(Trim(Text1)), "0")
    strExc(0) = "Select Nvl(cu04,Nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) as CName From Customer " & _
                     "Where cu01='" & Text1 & "' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        lbl_Name.Caption = RsTemp("CName")
        Exit Sub
    End If
    lbl_Name.Caption = ""
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(0).Value = True

End Sub

Private Sub Text2_GotFocus()
   '若為英文
   If Option1(1).Value = True Then
   'Modify By Cheng 2002/04/23
   '   Text2.IMEMode = 2: DoEvents
   'edit by nickc 2007/06/06 切換輸入法改用API
   CloseIme
   '若為中文或日文
   ElseIf Me.Option1(0).Value Or Me.Option1(2).Value Then
   'Modify By Cheng 2002/04/23
   '   Text2.IMEMode = 1: DoEvents
   'edit by nickc 2007/06/06 切換輸入法改用API
   OpenIme
   End If
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(1).Value = True

End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_LostFocus()
   'Modify By Cheng 2002/03/14
   ''Add By Cheng 2002/01/07
   'Me.Text3.Text = GetAllSysKind(Me.Text3)
End Sub

Private Sub Text4_GotFocus()
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If PUB_CheckKeyInDate(Me.Text4) = -1 Then
      Me.Text4.SetFocus
      Text4_GotFocus
   End If
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If PUB_CheckKeyInDate(Me.Text5) = -1 Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Sub
   End If
   If Not nickChgRan(Text4, Text5, "收文日期") Then
      Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6)
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
   If RunNick(Text6, Text7) Then
      Text6.SetFocus
      Text6_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8)
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_LostFocus()
   If Text8 <> "N" And Text8 <> "n" And Text8 <> "" Then
       s = MsgBox("請輸入N或空白", , "USER 輸入錯誤")
       Text8.SetFocus
       Text8_GotFocus
   End If
End Sub

'Add by Amy 2014/04/22
Private Sub Text9_GotFocus()
   Text9.SelStart = 0
   Text9.SelLength = Len(Text9)
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Option2(2).Value = True
End Sub
