VERSION 5.00
Begin VB.Form frm160309 
   BorderStyle     =   1  '單線固定
   Caption         =   "尾牙抽獎中獎名單"
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   5040
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   1704
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1296
      Width           =   324
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtYEAR 
      Height          =   270
      Left            =   1692
      MaxLength       =   3
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   5
      Top             =   2490
      Width           =   4875
      Begin VB.ComboBox cmbPrinter 
         Height          =   276
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年度：                   (ex:112)"
      Height          =   180
      Left            =   1104
      TabIndex        =   8
      Top             =   936
      Width           =   2028
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所別：           (1:北 2:中 3:南 4:高)"
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   1344
      Width           =   2580
   End
End
Attribute VB_Name = "frm160309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/11/20
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txtYEAR = "" Then
            MsgBox "請輸入年度！", vbExclamation
            txtYEAR.SetFocus
            Exit Sub
         ElseIf Val(txtYEAR) < 100 Or Val(txtYEAR) > 200 Then
            MsgBox "年度輸入錯誤！", vbCritical
            txtYEAR.SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         Call StrMenu
         PUB_SetOsDefaultPrinter pub_OsPrinter
         '若印表機變動, 則更新列印設定
         If cmbPrinter.Tag <> cmbPrinter Then
             PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
         End If
         Screen.MousePointer = vbDefault
         
      Case 1
         Unload Me
   End Select
End Sub

'明細表
Sub StrMenu()
   Dim rsQuery As ADODB.Recordset
   strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) Amt,mb04,st02,mb03" & _
      " from MiscBonus,staff where mb01=" & (Val(txtYEAR) + 1911) & IIf(txtZone <> "", " and mb10='" & txtZone & "'", "") & _
      " and mb02='01' and st01(+)=mb04" & _
      " order by mb03,mb04,mb05"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If doPrint(rsQuery) = True Then
         ShowPrintOk
      End If
   Else
      ShowNoData
   End If
   Set rsQuery = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160309 = Nothing
End Sub

Private Sub txtYEAR_GotFocus()
   TextInverse txtYEAR
End Sub

Private Sub txtYEAR_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("4")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function doPrint(pRst As ADODB.Recordset) As Boolean
   Dim iCount As Integer, stMB03 As String, stText As String, stList As String
On Error GoTo ErrHnd
   If Pub_NewWordDoc(g_WordAp) = True Then
      PUB_SetOsDefaultPrinter cmbPrinter
      PUB_SetWordActivePrinter
      With pRst
      .MoveFirst
      stMB03 = pRst.Fields("MB03")
      iCount = 1
      Do While Not pRst.EOF
         If stMB03 <> pRst.Fields("MB03") Then
            iCount = iCount + 1
            stMB03 = pRst.Fields("MB03")
         End If
         .MoveNext
      Loop
      End With
      
      With g_WordAp
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2.5)
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2.5)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      '.Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      '新增表格(2*N)
      .Selection.Tables.add Range:=.Selection.Range, NumColumns:=2, NumRows:=iCount + 2
      With .Selection.Tables(1)
         .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
         .Columns(1).Cells.VerticalAlignment = wdAlignVerticalCenter
         .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
         .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
         .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
         .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
         .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
         .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
         .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
         .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
         .Borders.Shadow = False
         .Rows(1).Select
      End With
      .Selection.Cells.Merge
      
      '設定表格高度欄寬
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Font.Size = 16
      strExc(1) = ""
      If txtZone = "1" Then
         strExc(1) = "(北所)"
      ElseIf txtZone = "2" Then
         strExc(1) = "(中所)"
      ElseIf txtZone = "3" Then
         strExc(1) = "(南所)"
      ElseIf txtZone = "4" Then
         strExc(1) = "(高所)"
      End If
      strExc(0) = txtYEAR & "年度尾牙抽獎中獎名單" & strExc(1)
      .Selection.TypeText Text:=strExc(0)
      
      .Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
      
      .Selection.Tables(1).Rows(2).Cells(1).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="獎　別　紅　包"
      .Selection.Tables(1).Rows(2).Cells(2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="姓　　　　　名"
      
      iCount = 2
      stMB03 = ""
      pRst.MoveFirst
      Do While Not pRst.EOF
         If stMB03 <> pRst.Fields("MB03") Then
            If stMB03 <> "" Then
               stList = stList & stText
               .Selection.Tables(1).Rows(iCount).Cells(2).Select
               .Selection.TypeText Text:=stList
            End If
            iCount = iCount + 1
            .Selection.Tables(1).Rows(iCount).Cells(1).Select
            .Selection.TypeText Text:=pRst.Fields("mb06") & " " & pRst.Fields("amt") & "元"
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            stMB03 = pRst.Fields("MB03")
            stText = pRst.Fields("st02")
            stList = ""
         Else
            If Len(stText & "、" & pRst.Fields("st02")) > 28 Then
               stList = stList & stText & vbCrLf
               stText = pRst.Fields("st02")
            Else
               stText = stText & "、" & pRst.Fields("st02")
            End If
         End If
         pRst.MoveNext
      Loop
      stList = stList & stText
      .Selection.Tables(1).Rows(iCount).Cells(2).Select
      .Selection.TypeText Text:=stList
               
      .PrintOut Background:=False, Copies:=1, Collate:=True
      .ActiveDocument.Close wdDoNotSaveChanges
      .Quit wdDoNotSaveChanges
      End With
      Set g_WordAp = Nothing
      doPrint = True
   End If
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function
