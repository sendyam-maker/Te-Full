VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100129_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部業務消長分析表"
   ClientHeight    =   5360
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5360
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdok 
      Caption         =   "Word(&W)"
      Enabled         =   0   'False
      Height          =   405
      Index           =   2
      Left            =   6255
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4485
      Left            =   15
      TabIndex        =   2
      Top             =   540
      Width           =   9405
      _ExtentX        =   16581
      _ExtentY        =   7902
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "發　文|增減|比率|98.4|事務所Ａ|比率|廠商Ｂ|比率|97.4|事務所Ａ|比率|廠商Ｂ|比率|96.4|事務所Ａ|比率|廠商Ｂ|比率"
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
      _Band(0).Cols   =   18
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   7200
      TabIndex        =   0
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   8370
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblMemo 
      Caption         =   "備註：＊為互惠代理人"
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   5070
      Width           =   2085
   End
End
Attribute VB_Name = "frm100129_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Create by Morgan 2010/9/2
Option Explicit

Public cmdState As Integer
Public m_RptType As String

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100129_2 = Nothing
End Sub

Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      Case 2
         Screen.MousePointer = vbHourglass
         runWord
         Screen.MousePointer = vbDefault
      Case Else
   End Select
End Sub

Public Sub SetGrid(p_Rst As ADODB.Recordset, p_FormatString As String)
   Dim iRow As Integer, iCol As Integer, lngTot1(9) As Long, iNum As Integer
   
   With grdDataList
      .Visible = False
      Set .Recordset = p_Rst.Clone
      .FormatString = p_FormatString
      .AllowUserResizing = flexResizeBoth
      .RowHeight(0) = .RowHeight(0) * 2
      Select Case m_RptType
      Case "11", "12"
         .WordWrap = True
         .ColAlignmentFixed = flexAlignCenterCenter
         .ColWidth(1) = 430
         .ColWidth(2) = 445
         .ColWidth(3) = 480
         .ColWidth(4) = 450
         
         .ColWidth(6) = 430
         .ColWidth(7) = 445
         .ColWidth(8) = 480
         .ColWidth(9) = 450
         .ColWidth(11) = 360
         .ColWidth(12) = 580
         
         .ColWidth(13) = 430
         .ColWidth(14) = 445
         .ColWidth(15) = 480
         .ColWidth(16) = 450
         .ColWidth(18) = 360
         .ColWidth(19) = 580
         
         For iRow = 1 To .Rows - 1
            If Val(.TextMatrix(iRow, 11)) > 0 Then
               .TextMatrix(iRow, 11) = "+" & .TextMatrix(iRow, 11)
               .TextMatrix(iRow, 12) = "+" & .TextMatrix(iRow, 12)
            End If
            If Val(.TextMatrix(iRow, 18)) > 0 Then
               .TextMatrix(iRow, 18) = "+" & .TextMatrix(iRow, 18)
               .TextMatrix(iRow, 19) = "+" & .TextMatrix(iRow, 19)
            End If
            lngTot1(1) = lngTot1(1) + Val(.TextMatrix(iRow, 1))
            lngTot1(2) = lngTot1(2) + Val(.TextMatrix(iRow, 2))
            lngTot1(3) = lngTot1(3) + Val(.TextMatrix(iRow, 4))
            
            lngTot1(4) = lngTot1(4) + Val(.TextMatrix(iRow, 6))
            lngTot1(5) = lngTot1(5) + Val(.TextMatrix(iRow, 7))
            lngTot1(6) = lngTot1(6) + Val(.TextMatrix(iRow, 9))
            
            lngTot1(7) = lngTot1(7) + Val(.TextMatrix(iRow, 13))
            lngTot1(8) = lngTot1(8) + Val(.TextMatrix(iRow, 14))
            lngTot1(9) = lngTot1(9) + Val(.TextMatrix(iRow, 16))
            
         Next
         
         
         If lngTot1(1) > 0 Then
            strExc(2) = Round(lngTot1(2) / lngTot1(1) * 100) & "%"
            strExc(3) = Round(lngTot1(3) / lngTot1(1) * 100) & "%"
         Else
            strExc(2) = ""
            strExc(3) = ""
         End If
         strExc(0) = "合　計" & vbTab & lngTot1(1) & vbTab & lngTot1(2) & vbTab & strExc(2) & vbTab & lngTot1(3) & vbTab & strExc(3)
                 
         
         If lngTot1(4) > 0 Then
            strExc(2) = Round(lngTot1(5) / lngTot1(4) * 100) & "%"
            strExc(3) = Round(lngTot1(6) / lngTot1(4) * 100) & "%"
         Else
            strExc(2) = ""
            strExc(3) = ""
         End If
         strExc(0) = strExc(0) & vbTab & lngTot1(4) & vbTab & lngTot1(5) & vbTab & strExc(2) & vbTab & lngTot1(6) & vbTab & strExc(3)
         
         If lngTot1(4) - lngTot1(1) > 0 Then
            strExc(1) = "+"
         Else
            strExc(1) = ""
         End If
         If lngTot1(1) > 0 Then
            strExc(2) = strExc(1) & Round((lngTot1(4) - lngTot1(1)) / lngTot1(1) * 100) & "%"
         Else
            strExc(2) = ""
         End If
         
         strExc(0) = strExc(0) & vbTab & strExc(1) & (lngTot1(4) - lngTot1(1)) & vbTab & strExc(2)
         
         If lngTot1(7) > 0 Then
            strExc(2) = Round(lngTot1(8) / lngTot1(7) * 100) & "%"
            strExc(3) = Round(lngTot1(9) / lngTot1(7) * 100) & "%"
         Else
            strExc(2) = ""
            strExc(3) = ""
         End If
         strExc(0) = strExc(0) & vbTab & lngTot1(7) & vbTab & lngTot1(8) & vbTab & strExc(2) & vbTab & lngTot1(9) & vbTab & strExc(3)
         
         If lngTot1(7) - lngTot1(4) > 0 Then
            strExc(1) = "+"
         Else
            strExc(1) = ""
         End If
         If lngTot1(4) > 0 Then
            strExc(2) = strExc(1) & Round((lngTot1(7) - lngTot1(4)) / lngTot1(4) * 100) & "%"
         Else
            strExc(2) = ""
         End If
         
         strExc(0) = strExc(0) & vbTab & strExc(1) & (lngTot1(7) - lngTot1(4)) & vbTab & strExc(2)
         .AddItem strExc(0)
         
      Case Else
         .WordWrap = False
         If Right(m_RptType, 1) <= "2" Then
            .ColWidth(0) = 1000
            iNum = 5
         Else
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(0) = 950
            .ColWidth(1) = 1800
            If Right(m_RptType, 1) = "3" Then
               .ColWidth(2) = 400
               iNum = 7
            Else
               iNum = 6
            End If
         End If
         
         For iRow = 1 To .Rows - 1
            For intI = 1 To 3
               iCol = 5 * (intI - 1) + iNum
               If iCol < .Cols Then
                  If Val(.TextMatrix(iRow, iCol)) > 0 Then
                     .TextMatrix(iRow, iCol) = "+" & .TextMatrix(iRow, iCol)
                  End If
                  iCol = 5 * (intI - 1) + iNum - 2
                  If Val(.TextMatrix(iRow, iCol)) > 0 Then
                     .TextMatrix(iRow, iCol) = "+" & .TextMatrix(iRow, iCol)
                  End If
               End If
            Next
         Next
         
      End Select
      
      .Visible = True
   End With
End Sub

Private Sub runWord()
   
   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim iTable As Integer
   Dim iLine As Integer
   Dim iJudgeCol As Integer, bolNoCaseTag As Boolean  'Added by Morgan 2015/1/20
   
On Error GoTo ErrHnd
   
   If Right(m_RptType, 1) > "2" Then
      bolPrintCountry = True
      'Added by Morgan 2015/1/20
      If Right(m_RptType, 1) = 3 Then
         iJudgeCol = 6
      Else
         iJudgeCol = 5
      End If
      ''end 2015/1/20
   
   Else
      bolPrintCountry = False
   End If
            
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   'g_WordAp.Visible = True
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Visible = False
      .Selection.Font.Name = "標楷體"
      
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        '.DefaultBorderColor = wdColorBlack 'Word97 沒有這個屬性及常數(Word2007 有)
      End With
            
      If grdDataList.Cols > 8 Then
         'Modified by Morgan 2012/6/25 --David
         '.Selection.PageSetup.Orientation = wdOrientLandscape
         'strFontSize = 14
         .Selection.PageSetup.Orientation = wdOrientPortrait
         strFontSize = 10
      Else
         .Selection.PageSetup.Orientation = wdOrientPortrait
         strFontSize = 12
      End If
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      '邊界
      'Modified by Morgan 2012/6/25 --David
      '.Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
      '.Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      stTmp = Me.Caption
      
      .Selection.Font.Size = 18
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=stTmp
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.Font.Size = strFontSize
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      For iRow = 0 To grdDataList.Rows - 1
         If iRow = 0 Then
            If bolPrintCountry Then
               bolNoCaseTag = False 'Added by Morgan 2015/1/20
               strLastCountry = grdDataList.TextMatrix(1, 0)
               .Selection.Font.Size = 14
               .Selection.TypeText Text:=strLastCountry
               .Selection.Font.Size = strFontSize
               iFCol = 1
               iCols = grdDataList.Cols - 1
               
               'Added by Morgan 2015/1/20
               If Val(grdDataList.TextMatrix(1, iJudgeCol)) = 0 Then
                  bolNoCaseTag = True
                  .Selection.TypeParagraph
                  .Selection.Font.Size = 10
                  .Selection.TypeText Text:="當期無案件："
                  .Selection.Font.Size = strFontSize
               End If
               'end 2015/1/20
            Else
               iFCol = 0
               iCols = grdDataList.Cols
            End If
            
            AddNewTable iCols, bolPrintCountry, iFCol
            
         Else
            If bolPrintCountry Then
               If strLastCountry <> grdDataList.TextMatrix(iRow, 0) Then
                  bolNoCaseTag = False 'Added by Morgan 2015/1/20
                  strLastCountry = grdDataList.TextMatrix(iRow, 0)
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  .Selection.TypeParagraph
                  .Selection.Font.Size = 14
                  .Selection.TypeText Text:=strLastCountry
                  .Selection.Font.Size = strFontSize
                  
                  'Added by Morgan 2015/1/20
                  If Val(grdDataList.TextMatrix(iRow, iJudgeCol)) = 0 Then
                     bolNoCaseTag = True
                     .Selection.TypeParagraph
                     .Selection.Font.Size = 10
                     .Selection.TypeText Text:="當期無案件："
                     .Selection.Font.Size = strFontSize
                  End If
                  'end 2015/1/20
                  
                  AddNewTable iCols, bolPrintCountry, iFCol
               End If
            
               'Added by Morgan 2015/1/20
               If bolNoCaseTag = False Then
                  If Val(grdDataList.TextMatrix(iRow, iJudgeCol)) = 0 Then
                     bolNoCaseTag = True
                     .Selection.MoveDown Unit:=wdLine, Count:=1
                     .Selection.Font.Size = 10
                     .Selection.TypeText Text:="當期無案件："
                     .Selection.Font.Size = strFontSize
                     AddNewTable iCols, bolPrintCountry, iFCol
                  End If
               End If
               'end 2015/1/20
            End If
            
            .Selection.InsertRows 1
            .Selection.Cells.SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1
            For iCol = iFCol To grdDataList.Cols - 1
               .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            Next
         End If
         
      Next
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.TypeText Text:=lblMemo
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      'Add by Morgan 2010/9/23 插入頁碼
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
      .Selection.HeaderFooter.PageNumbers.add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
      .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      
      .Visible = True
      .Activate
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & iLine & " : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Sub

Private Sub AddNewTable(iColCount As Integer, bolCountryType As Boolean, iFromCol As Integer)
   Dim iCol As Integer
   
   With g_WordAp.Application
      '列數,欄數
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=iColCount
      
      'Added by Morgan 2011/12/13
      .Selection.SelectRow
      With .Selection.Borders(wdBorderTop)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
          '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      With .Selection.Borders(wdBorderLeft)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
          '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      With .Selection.Borders(wdBorderBottom)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
          '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      With .Selection.Borders(wdBorderRight)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
          ''Word97 沒有這個屬性及常數(Word2007 有).Color = g_WordAp.Options.DefaultBorderColor
      End With
      With .Selection.Borders(wdBorderHorizontal)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          '.LineWidth = g_WordAp.Options.DefaultBorderLineWidth'Word巨集正常但vb跑會有錯
          '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      With .Selection.Borders(wdBorderVertical)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
          '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      'end 2011/12/13
      
      '設定表格高度
      .Selection.SelectRow
      '.Selection.Font.Bold = wdToggle
      .Selection.Cells.SetHeight RowHeight:=36, HeightRule:=wdRowHeightExactly
      If bolCountryType Then
         .Selection.MoveLeft Unit:=wdCharacter, Count:=1
         .Selection.SelectColumn
         .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(6), RulerStyle:=wdAdjustSameWidth
         .Selection.SelectRow
      End If
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.MoveLeft Unit:=wdCharacter, Count:=1
      For iCol = iFromCol To grdDataList.Cols - 1
         .Selection.TypeText Text:=grdDataList.TextMatrix(0, iCol)
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Next
   End With
End Sub
