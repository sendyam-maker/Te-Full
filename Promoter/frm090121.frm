VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090121 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標查名報告"
   ClientHeight    =   5745
   ClientLeft      =   900
   ClientTop       =   1050
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6225
   Begin VB.CheckBox Check2 
      Caption         =   "其他補充說明："
      Height          =   285
      Left            =   300
      TabIndex        =   8
      Top             =   4020
      Width           =   2085
   End
   Begin VB.CheckBox Check1 
      Caption         =   "附件"
      Height          =   285
      Left            =   300
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtLetterHead 
      Height          =   270
      Left            =   4245
      MaxLength       =   1
      TabIndex        =   11
      Top             =   5130
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   4890
      TabIndex        =   13
      Top             =   105
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Word編輯(&W)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3570
      TabIndex        =   12
      Top             =   105
      Width           =   1290
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   540
      TabIndex        =   9
      Top             =   4320
      Width           =   5565
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "9816;529"
      Value           =   "如不申請商標註冊，本所將酌收查詢工本費新台幣　　元。"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   3540
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   3180
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1770
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   2115
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   2475
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   2835
      Width           =   4900
      VariousPropertyBits=   -1467989989
      MaxLength       =   30
      Size            =   "8643;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   660
      Width           =   3465
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "6112;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   1
      Top             =   990
      Width           =   3465
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "6112;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Left            =   990
      TabIndex        =   23
      Top             =   3600
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Left            =   990
      TabIndex        =   22
      Top             =   3240
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Left            =   990
      TabIndex        =   21
      Top             =   1815
      Width           =   90
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Left            =   990
      TabIndex        =   20
      Top             =   2145
      Width           =   90
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Left            =   990
      TabIndex        =   19
      Top             =   2505
      Width           =   90
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Left            =   990
      TabIndex        =   18
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label3 
      Caption         =   "委查文字：（圖形以文字代替）"
      Height          =   165
      Left            =   420
      TabIndex        =   17
      Top             =   1560
      Width           =   3195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人："
      Height          =   180
      Left            =   480
      TabIndex        =   16
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱："
      Height          =   180
      Left            =   300
      TabIndex        =   15
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "是否印信頭：        （N:不印）"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   14
      Top             =   5130
      Width           =   2370
   End
End
Attribute VB_Name = "frm090121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Lydia 2022/01/26 改成Form2.0 ; txt1(index)
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim bolRetry As Boolean '是否已發生錯誤且重試


Private Sub Command1_Click()
   '檢查資料
   If txt1(0).Text = "" And txt1(1).Text = "" Then
       MsgBox "請輸入客戶名稱或聯絡人!", vbExclamation + vbOKOnly
       txt1(0).SetFocus
       txt1_GotFocus (0)
       Exit Sub
   End If
   If txt1(2).Text = "" And txt1(3).Text = "" And txt1(4).Text = "" And _
      txt1(5).Text = "" And txt1(6).Text = "" And txt1(7).Text = "" Then
       MsgBox "請輸入委查文字!", vbExclamation + vbOKOnly
       txt1(2).SetFocus
       txt1_GotFocus (2)
       Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Call WordEdit
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub WordEdit()
Dim stFileName As String '暫存圖檔檔名
Dim iPicNo As Integer '圖檔代碼 1:外商 2:外專/外法 3.CFP 4.其他
Dim strText As String, tmpArr As Variant
Dim i As Integer
Dim oShape 'Added by Morgan 2015/8/18
Dim iPicNo2 As Integer 'Added by Morgan 2020/3/31
   
   bolRetry = False
   
'On Error GoTo ERRORSECTION1
   
   strText = ""
   For i = 2 To 7
      If Trim(txt1(i)) <> "" Then
         If strText <> "" Then strText = strText & "、"
         strText = strText & "「" & Trim(txt1(i)) & "」"
      Else
         Exit For
      End If
   Next i
   
   'If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
      'Added by Lydia 2016/1/13
      '切換為整頁模式,信頭才會正常顯示
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      
      '設定字型版面(參照定稿)
      '.Selection.Font.Name = "Times New Roman"
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      'Added by Morgan 2020/3/31 改和撰寫信函一致
      If strSrvDate(1) >= 智慧所更名日 Then
         .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
         .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
         .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.2)
         .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      Else
      'end 2020/3/31
         .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.175)
         .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.175)
         .Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
         .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      End If 'Added by Morgan 2020/3/31
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      '不要分散對齊
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      
      '信函信頭
      If txtLetterHead <> "N" Then
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= 智慧所更名日 Then
            PUB_GetLetterPicID "2", "T", iPicNo, iPicNo2, 1, False, Pub_StrUserSt03
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0)
               If iPicNo2 > 0 Then
                  If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                     .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     oShape.Top = .CentimetersToPoints(27.2)
                  End If
                  .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
               End If
               .Selection.EndKey Unit:=wdStory
            End If
         Else
         'end 2020/3/31
            
            iPicNo = "7"
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               '插入圖片檔案
               'Modified by Morgan 2015/8/18  Word2007會有錯(找不到項目)
              ' .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
              ' .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.Select
               'end 2015/8/18
               'Modified by Lydia 2015/12/02 先前Word2010會出錯,在於後面的程式應該改為對物件操作,
   '            .Selection.ShapeRange.ZOrder 4
   '            .Selection.ShapeRange.LockAnchor = True
   '            .Selection.ShapeRange.LockAspectRatio = -1
   '            .Selection.ShapeRange.Width = 546.5
   '            .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
   '            .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
   '            .Selection.ShapeRange.Left = .CentimetersToPoints(1)
   '            .Selection.ShapeRange.Top = .CentimetersToPoints(1)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = 546.5
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(1)
               oShape.Top = .CentimetersToPoints(1)
               'end 2015/12/02
               .Selection.EndKey Unit:=wdStory
            End If
         End If 'Added by Morgan 2020/3/31
      Else
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      End If
      .Selection.TypeParagraph
      'Modified by Lydia 2015/09/18 修改格式(下移3列)
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      'end 2015/09/18
      
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText Trim(txt1(0).Text)
      If Trim(txt1(1).Text) <> "" Then
         .Selection.TypeParagraph
         .Selection.TypeText Trim(txt1(1).Text)
      End If
      .Selection.TypeParagraph
      '系統日期
      '置右
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Mid(strSrvDate(1), 1, 4) & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日"
      .Selection.TypeParagraph
      '置中
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText "事由：查詢結果報告"
      .Selection.TypeParagraph
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText "敬啟者："
      .Selection.TypeParagraph
      .Selection.TypeText "　　貴公司委託本所代為查詢" & strText & "於指定使用之商品或服務，茲將查詢結果，報告如下" & IIf(Check1.Value = 1, "(近似商標資料請參附件)", "") & "："
      .Selection.TypeParagraph
      
      tmpArr = Split(strText, "、")
      For i = 0 To UBound(tmpArr)
         'Modified by Lydia 2015/09/18 修改格式
'         .Selection.TypeText "　　" & i + 1 & ". " & Mid(TmpArr(i), 2, Len(TmpArr(i)) - 2) & "："
'         .Selection.TypeParagraph
'         .ActiveDocument.Tables.Add Range:=.Selection.Range, NumRows:=3, NumColumns:=4
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.TypeText Text:="類別組群"
'         .Selection.MoveRight Unit:=wdCell
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.TypeText Text:="查名結果"
'         .Selection.MoveRight Unit:=wdCell
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.TypeText Text:="近似商標"
'         .Selection.MoveRight Unit:=wdCell
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.TypeText Text:="號數"
'         '.Selection.MoveRight Unit:=wdCharacter, Count:=12
         .Selection.TypeText "　　" & i + 1 & ". 「" & Mid(tmpArr(i), 2, Len(tmpArr(i)) - 2) & "」："
         .Selection.TypeParagraph
          '新增表格(1X5)
          .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=5
 
          '畫格線
          With .Selection.Tables(1)
            .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
            .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
            .Borders.Shadow = False
          End With

         .Selection.SelectRow
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.3), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.3), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.3), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(4.6), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="類別組群"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.TypeText Text:="主要商品"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.TypeText Text:="查名結果"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.TypeText Text:="引證資料"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.TypeText Text:="建議"
         '第2行
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.MoveRight Unit:=wdCell, Count:=3
         'Modified by Lydia 2015/10/07
         '.Selection.Cells.Split NumColumns:=2
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
         
         .Selection.TypeText Text:="近似商標"

         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.TypeText Text:="號數"
         'Modified by Lydia 2015/10/14 因為游標問題,修改
         '.Selection.MoveUp Unit:=wdLine, Count:=1
         .Selection.MoveLeft Unit:=wdCell, Count:=6
         .Selection.SelectRow
         .Selection.Collapse Direction:=wdCollapseStart
         'Modified by Lydia 2015/10/07 word2000 無法指定特定儲存格合併
'         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
'         .Selection.Cells.Merge
'         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
'         .Selection.Cells.Merge
'         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
'         .Selection.Cells.Merge
'         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
'         .Selection.MoveRight Unit:=wdCell, Count:=2
'         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
'         .Selection.Cells.Merge
         .Selection.SelectColumn
         .Selection.Cells.Merge
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.SelectColumn
         .Selection.Cells.Merge
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.SelectColumn
         .Selection.Cells.Merge
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.MoveRight Unit:=wdCell, Count:=2
         .Selection.SelectColumn
         .Selection.Cells.Merge
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.MoveRight Unit:=wdCell, Count:=1
'--------end 2015/10/07
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.MoveRight Unit:=wdCell, Count:=20 '---Insert 3列
         .Selection.EndKey Unit:=wdStory
         .Selection.TypeParagraph
         'end 2015/09/18
      Next i
         
      .Selection.TypeParagraph
      'Modified by Lydia 2015/09/18
'      .Selection.TypeText "　　本查詢資料來源係依據智慧財產局所發行之商標公報。倘有爭議發生時，仍須以智慧財產局之資料為準，此查詢結果僅供參考。"
       strExc(1) = "　　本查詢資料來源係依據智慧財產局所發行之商標公報，其內容僅供參考。" & _
                   "不得作為准駁之依據，商標申請准駁與否仍應以智慧財產局的審查結果為準。" & _
                   "若商標於申請核准即實際使用，應注意是否可能對他人商標構成侵權。"
      .Selection.TypeText strExc(1)
      .Selection.TypeParagraph
      .Selection.TypeText "　　" & IIf(Trim(txt1(8).Text) <> "" And Check2.Value = 1, Trim(txt1(8).Text), "") & "若尚有任何問題，請隨時與本所聯繫，本所竭誠為　貴公司服務！"
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "　　　　耑此     順頌"
      .Selection.TypeParagraph
      .Selection.TypeText "商祺"
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
      'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
      '.Selection.TypeText "台一國際專利商標事務所  敬上"
      .Selection.TypeText PUB_GetCompName2("1") & "  敬上"
      'end 2020/3/30
      .Selection.TypeParagraph
      
      .Selection.WholeStory
      ChgWordFormat g_WordAp, .Selection.Text
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   Set g_WordAp = Nothing
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            If bolRetry = True Then
               MsgBox Err.Description, vbCritical
            Else
               Set g_WordAp = New Word.Application
               g_WordAp.Documents.add
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox Err.Description, vbCritical
            Resume
      End Select
   End If
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090121 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   'TextInverse txt1(Index)
   InverseTextBox txt1(Index)
   'OpenIme
End Sub

'Modified by Lydia 2022/01/26 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   'CloseIme
End Sub

Private Sub txtLetterHead_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
   End If
End Sub
