VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm077004_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "介紹法務所案源查詢-明細"
   ClientHeight    =   5790
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdOK 
      Caption         =   "法律案源接洽單"
      Height          =   375
      Index           =   7
      Left            =   60
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "P/T接洽單"
      Height          =   315
      Index           =   6
      Left            =   570
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   5430
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "法律卷宗區"
      Height          =   375
      Index           =   5
      Left            =   3690
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "法律進度"
      Height          =   375
      Index           =   4
      Left            =   2775
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "法律案件資料"
      Height          =   375
      Index           =   3
      Left            =   1530
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案源案件資料"
      Height          =   375
      Index           =   0
      Left            =   4890
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案源進度"
      Height          =   375
      Index           =   1
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案源卷宗區"
      Height          =   375
      Index           =   2
      Left            =   7080
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8310
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4845
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   9255
      _ExtentX        =   16334
      _ExtentY        =   8537
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "V|介紹日|管制日期|業務區|介紹人|介紹客戶|介紹內容|案源案號|法律所案號|收文日|進度備註|法律所業務|案源單號"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7590
      TabIndex        =   9
      Top             =   5490
      Width           =   1710
   End
End
Attribute VB_Name = "frm077004_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Sindy 2020/5/4
Option Explicit

'Memo by Lydia 2022/01/07 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、P/T接洽單(Printer列印未改)
Dim m_adoRst As ADODB.Recordset
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public cmdState As Integer


Private Sub cmdExit_Click()
   frm077004.Show
   Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim i As Integer, StrTag As String
Dim Str01 As String
Dim ii As Integer
Dim strKey As String, StrKey2 As String
Dim frmTmp As Form

On Error GoTo ErrorHandler

   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   If cmdState = 6 Then 'P/T接洽單
      For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            '有法律案接洽單號,無案源總收文號時才可以列印P/T接洽單
            If Trim(grdDataList.TextMatrix(i, 13)) = "" And _
               Trim(grdDataList.TextMatrix(i, 15)) <> "" Then
               
               StrKey2 = Trim(grdDataList.TextMatrix(i, 16)) 'PT接洽單號
               If StrKey2 = "" Then
                  MsgBox "案源單號(" & Trim(grdDataList.TextMatrix(i, 12)) & ")無PT接洽單！", vbExclamation
                  GoTo ErrorHandler
               Else
                  grdDataList.col = 0
                  grdDataList.Text = ""
                  For ii = 0 To grdDataList.Cols - 1
                     grdDataList.col = ii
                     grdDataList.CellBackColor = grdDataList.BackColor 'lngColor
                  Next
                  
                  strKey = Trim(grdDataList.TextMatrix(i, 12)) '案源單號
'                  If PUB_CheckFormExist("frm090801") Then
'                     MsgBox "請先關閉接洽單畫面！"
'                     GoTo ErrorHandler
'                  End If
                  Set frmTmp = Forms(0).GetForm("frm090801")
                  'With frm090801
                  With frmTmp
                     .SetParent Me
                     .Load4Print strKey, StrKey2, False
                     .Show
                     GoTo ErrorHandler
                  End With
               End If
            Else
               MsgBox "案源單號(" & Trim(grdDataList.TextMatrix(i, 12)) & ")不可列印PT接洽單！", vbExclamation
               GoTo ErrorHandler
            End If
         End If
      Next i
      
   Else
      For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            bolRefresh = False
            grdDataList.col = 0
            grdDataList.Text = ""
            'grdDataList.CellBackColor = grdDataList.BackColor
            'grdDataList.col = 3
            'lngColor = grdDataList.CellBackColor
            'For ii = 1 To 2
            For ii = 0 To grdDataList.Cols - 1
               grdDataList.col = ii
               grdDataList.CellBackColor = grdDataList.BackColor 'lngColor
            Next
            
            'Added by Morgan 2023/2/6 法律案源接洽單
            If cmdState = 7 Then
               StrTag = grdDataList.TextMatrix(i, 15)
               strExc(1) = DBDATE(grdDataList.TextMatrix(i, 1))
               Call PUB_Queryfrm090801(StrTag, strExc(1), Me)
               'Me.Hide 'Modify By Sindy 2023/5/9 Mark
               Exit For
            Else
            'end 2023/2/6
               If cmdState = 0 Or cmdState = 1 Or cmdState = 2 Then
                  StrTag = grdDataList.TextMatrix(i, 7) '案源案號
               Else
                  StrTag = grdDataList.TextMatrix(i, 8) '法律案號
               End If
               If StrTag <> "" Then
                  If Left(Right(StrTag, 7), 1) = "-" Then
                     StrTag = StrTag & "-0-00"
                  ElseIf Left(Right(StrTag, 2), 1) = "-" Then
                     StrTag = StrTag & "-00"
                  End If
                  If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
                     StrTag = Right(StrTag, Len(StrTag) - 1)
                  End If
                  Str01 = SystemNumber(StrTag, 1)
                  If fnSaveParentForm(Me) = False Then
                     Exit For
                  End If
                  Me.Show
                  Select Case cmdState
                     Case 0, 3 '案件基本資料
                        Select Case Pub_RplStr(Str01)
                           Case "CFP", "FCP", "P"   '專利
                                 Screen.MousePointer = vbHourglass
                                 frm100101_3.Show
                                 frm100101_3.Tag = Pub_RplStr(StrTag)
                                 frm100101_3.StrMenu
                                 Screen.MousePointer = vbDefault
                           Case "CFT", "FCT", "T", "TF"   '商標
                                 Screen.MousePointer = vbHourglass
                                 frm100101_4.Show
                                 frm100101_4.Tag = Pub_RplStr(StrTag)
                                 frm100101_4.StrMenu
                                 Screen.MousePointer = vbDefault
                           'Modify By Sindy 2009/07/24 增加LIN系統類別
                           'modify by sonia 2019/7/29 +ACS系統類別
                           Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
                                 Screen.MousePointer = vbHourglass
                                 frm100101_5.Show
                                 frm100101_5.Tag = Pub_RplStr(StrTag)
                                 frm100101_5.StrMenu
                                 Screen.MousePointer = vbDefault
                           Case "LA"            '顧問
                                 Screen.MousePointer = vbHourglass
                                 frm100101_6.Show
                                 frm100101_6.Tag = Pub_RplStr(StrTag)
                                 frm100101_6.StrMenu
                                 Screen.MousePointer = vbDefault
                           Case Else                  '服務
                                Select Case Pub_RplStr(Str01)
                                    Case "TB"    '條碼
                                       Screen.MousePointer = vbHourglass
                                       frm100101_7.Show
                                       frm100101_7.Tag = Pub_RplStr(StrTag)
                                       frm100101_7.StrMenu
                                       Screen.MousePointer = vbDefault
                                    Case "TM"
                                       Screen.MousePointer = vbHourglass
                                       frm100101_8.Show
                                       frm100101_8.Tag = Pub_RplStr(StrTag)
                                       frm100101_8.StrMenu
                                       Screen.MousePointer = vbDefault
                                    Case "TD"
                                       Screen.MousePointer = vbHourglass
                                       frm100101_9.Show
                                       frm100101_9.Tag = Pub_RplStr(StrTag)
                                       frm100101_9.StrMenu
                                       Screen.MousePointer = vbDefault
                                    Case "TC", "CFC"
                                       Screen.MousePointer = vbHourglass
                                       frm100101_A.Show
                                       frm100101_A.Tag = Pub_RplStr(StrTag)
                                       frm100101_A.StrMenu
                                       Screen.MousePointer = vbDefault
                                    Case Else
                                       Screen.MousePointer = vbHourglass
                                       frm100101_B.Show
                                       frm100101_B.Tag = Pub_RplStr(StrTag)
                                       frm100101_B.StrMenu
                                       Screen.MousePointer = vbDefault
                                 End Select
                        End Select
                        Me.Hide
                     Case 1, 4 '案件進度
                        'Modify By Sindy 2020/5/29
                        '法律所人員看999999只需看該筆文號資料
                        If PUB_ChkLCompStaff(strUserNum) = True _
                           And InStr(StrTag, "TT-999999") > 0 _
                           And cmdState = 1 Then
                           frm100101_C.Show
                           frm100101_C.Tag = StrTag & "=" & Pub_RplStr(grdDataList.TextMatrix(i, 13)) '案源總收文號
                           frm100101_C.StrMenu
                        Else
                        '2020/5/29 END
                           frm100101_2.Show
                           frm100101_2.Tag = StrTag
                           frm100101_2.StrMenu
                        End If
                        Me.Hide
                        
                     Case 2, 5 '案源卷宗區
                        If cmdState = 2 Then
                           'StrTag = grdDataList.TextMatrix(i, 13) '案源總收文號
                           StrTag = grdDataList.TextMatrix(i, 7) '案源案號
                        Else
                           'StrTag = grdDataList.TextMatrix(i, 14) '法律總收文號
                           StrTag = grdDataList.TextMatrix(i, 8) '法律案號
                        End If
                        frm100101_L.m_strKey = Pub_RplStr(StrTag)
                        frm100101_L.SetParent Me
                        If frm100101_L.QueryData = True Then
                           frm100101_L.Show
                           Me.Hide
                        End If
                  End Select
                  Exit For
               End If
            End If 'Added by Morgan 2023/2/6
         End If
      Next i
   End If
'   If bolRefresh = True Then
'      cmdQuery_Click 0
'   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Caption = frm077004.Caption & "-明細"
   
'   Screen.MousePointer = vbHourglass
'   grdDataList.MousePointer = flexHourglass
'   Me.Enabled = False
'   doQuery
'   Me.Enabled = True
'   grdDataList.MousePointer = flexDefault
'   Screen.MousePointer = vbDefault

'   'Modify By Sindy 2022/11/1
'   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      cmdOK(6).Visible = False
'   End If
'   '2022/11/1 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm077004_1 = Nothing
End Sub

Public Sub doQuery(RsTemp As ADODB.Recordset)
   Set m_adoRst = RsTemp.Clone
   SetRst2Grid
End Sub

'有管制期限變黃
Private Sub SetRowColor()
   Dim jj As Integer
   With grdDataList
   If .TextMatrix(.row, 1) <> "" Then
      For jj = 0 To .Cols - 1
         .col = jj
         '黃
         .CellBackColor = &HFFFF&
      Next
   End If
   End With
End Sub

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      '                0  1        2        3      4      5          6              7        8          9        10           11         12
      .FormatString = "V |介紹日　|管制日期|業務區|介紹人|介紹客戶　|介紹內容　　　|案源案號|法律所案號|收文日　|進度備註　　|法律所業務|案源單號　"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         If intI > 12 Then
            .ColWidth(intI) = 0
         End If
      Next
      .Visible = True
   End With
End Sub

Private Sub SetRst2Grid()
Dim ii As Integer, jj As Integer
   
   'grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   'grdDataList.FixedCols = 4
   SetGrid
   With grdDataList
      If .Rows > 1 Then
         .Visible = False
         For ii = 1 To .Rows - 1
            .TextMatrix(ii, 4) = PUB_ReadUserData(Replace(.TextMatrix(ii, 4), ",", ";"))
            'Modify By Sindy 2020/6/22 TT-999999 案號不顯示
            If .TextMatrix(ii, 7) = "TT-999999" Then
               .TextMatrix(ii, 7) = ""
            End If
            '2020/6/22 END
   '            .RowHeight(ii) = 255
   '            .row = ii
   '            '固定欄位變回底色
   '            For jj = 0 To .FixedCols - 1
   '               .col = jj
   '               .CellBackColor = .BackColor
   '               .CellAlignment = flexAlignRightTop
   '               .CellFontSize = 9
   '            Next
   '            '有管制期限變黃
   '            SetRowColor
         Next
         .Visible = True
      End If
   End With

   LblCnt.Caption = "共 " & m_adoRst.RecordCount & " 筆"
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow grdDataList, X, Y, nCol, nRow
   grdDataList.col = IIf(nCol < 0, 0, nCol) 'nCol
   grdDataList.row = IIf(nRow < 0, 0, nRow) 'nRow
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iCol As Integer
   iCol = grdDataList.col
   If grdDataList.row < 1 Then
     grdDataList.Visible = False
'     ChgEmptyDate True
     Set grdDataList.Recordset = Nothing
     If m_blnColOrderAsc = True Then
        m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc,介紹日 desc" '& m_stSort
        m_blnColOrderAsc = False
     Else
        m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc,介紹日 asc" '& m_stSort
        m_blnColOrderAsc = True
     End If
     SetRst2Grid
     SetGrid
'     SetColor
     grdDataList.Visible = True
   End If
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            '.col = 0
            '.CellBackColor = .BackColor
            '.col = 3
            'lngColor = .CellBackColor
            'For ii = 1 To 2
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = .BackColor 'lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub
