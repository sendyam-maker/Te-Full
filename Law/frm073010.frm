VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm073010 
   BorderStyle     =   1  '≥ÊΩu©T©w
   Caption         =   "≈U∞›¶aß}±¯"
   ClientHeight    =   5820
   ClientLeft      =   210
   ClientTop       =   720
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.Frame Frame2 
      Caption         =   "¶aß}±¯µßº∆°G¶@0µß!!!"
      Height          =   4485
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   9225
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4110
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   7250
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "∑s≤”©˙≈È-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "≥]©w"
      Height          =   600
      Left            =   180
      TabIndex        =   3
      Top             =   5100
      Width           =   5256
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '≥ÊØ¬§U©‘¶°
         TabIndex        =   4
         Top             =   168
         Width           =   4392
      End
      Begin VB.Label Label2 
         Caption         =   "¶L™Ìæ˜"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "•˛≥°øÔ®˙(&A)"
      Height          =   400
      Left            =   6144
      TabIndex        =   0
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "¶^´eµe≠±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   2
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "¶C¶L(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7272
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "µ˘°G°¥•N™Ì¨∞¡¬´ﬂÆv™A∞»´»§·°A§¥∑|¶L¶aß}±¯"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   5490
      TabIndex        =   8
      Top             =   5100
      Width           =   3600
   End
End
Attribute VB_Name = "frm073010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 ¥º≈v§H≠˚ƒÊ§w≠◊ßÔ
'Memo By Sindy 2011/2/16 SQLDate§w¿À¨d
'Memo By Sindy 2010/11/26 ≠˚§uΩs∏πƒÊ§w≠◊ßÔ
'Memo by Lydia 2022/02/11 ßÔ¶®Form2.0 ; MSHFlexGrid1ßÔ¶r´¨=∑s≤”©˙≈È-ExtB ; Printer¶C¶L•ºßÔ
'Memo By Sindy 2010/8/3 §È¥¡ƒÊ§w≠◊ßÔ
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim SeekPrint As Integer
Dim SeekPrintL As Integer

Private Sub cmdBack_Click()
   frm073009.Show
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim i  As Integer
Dim nPageNo As Integer
'Add By Cheng 2003/04/02
Dim strArrCaseNo '•ª©“Æ◊∏π∞}¶C
    
'   'Add By Cheng 2003/05/21
'   'ßR∞£¶aß}±¯¶C™Ì∏ÍÆ∆
'   PUB_DeleteAddressList strUserNum
'   '™Ï©l§∆ß«∏π
'   pub_AddressListSN = 0

   cmdPrint.Enabled = False 'Add By Sindy 2012/1/30
   Screen.MousePointer = vbHourglass 'Add By Sindy 2012/1/30
   
   With MSHFlexGrid1
      nPageNo = 1
      For i = 1 To .Rows - 1
         .row = i
         .col = 0
         '2011/8/8 ¡Ÿ≠Ï by sonia §¥≠n¶L•”Ω–§H2
         'Modify By Sindy 2011/7/25 °¥•N™Ì¨∞¡¬´ﬂÆv™A∞»´»§·°A§£¶L¶aß}±¯
         If .Text = "v" Then
         'If .Text = "v" And Left(Trim(.TextMatrix(i, 2)), 1) <> "°¥" Then  '2011/8/8 cancel by sonia
         '2011/7/25 End
            'Modify By Cheng 2003/04/02
            Load frm083014
            frm083014.Hide
            frm083014.Opt1(0).Value = True
            '2011/8/8 modify by sonia
            'frm083014.Text1(0).Text = .TextMatrix(i, 2)
            If Left(Trim(.TextMatrix(i, 2)), 1) <> "°¥" Then
               frm083014.Text1(0).Text = .TextMatrix(i, 2)
            Else
               frm083014.Text1(0).Text = Pub_RplStr(.TextMatrix(i, 2))
            End If
            '2011/8/8 end
            frm083014.Text1(3).Text = "1"  '¶C¶L•˜º∆
            frm083014.Text1(4).Text = "1"  '§§§ÂÆÊ¶°
            frm083014.Text1(5).Text = "Y"  '2011/8/8 add by sonia
            frm083014.SetPageNo nPageNo
            frm083014.SetPrinter Me.Combo1.Text
            frm083014.cmdPrint_Click
            nPageNo = nPageNo + 1
            frm083014.cmdBack_Click
'            '≠Y¶≥•ª©“Æ◊∏π∏ÍÆ∆
'            If .TextMatrix(i, 5) <> "" Then
'                strArrCaseNo = Split(.TextMatrix(i, 5), "-")
'                '≥]©w¶aß}±¯¶C™Ì¨y§Ù∏π
'                pub_AddressListSN = pub_AddressListSN + 1
'                '∑sºW¶aß}±¯¶C™Ì∏ÍÆ∆
'                PUB_AddNewAddressList strUserNum, "" & strArrCaseNo(0), "" & strArrCaseNo(1), "" & strArrCaseNo(2), "" & strArrCaseNo(3), "" & pub_AddressListSN, "0"
'            End If
         End If
      Next i
   End With
   
   If Err.Number = 0 Then
'        'Add By Cheng 2003/01/29
'        '¶C¶L¶aß}±¯
'        PUB_PrintAddressList strUserNum, Me.Combo1.Text
'        'ßR∞£¶aß}±¯¶C™Ì∏ÍÆ∆
'        PUB_DeleteAddressList strUserNum
'        '™Ï©l§∆ß«∏π
'        pub_AddressListSN = 0
        MsgBox "¶C¶Lßπ¶®!", vbInformation, "≈U∞›¶aß}±¯"
        frm073009.Show
        Unload Me
   End If
   
   Screen.MousePointer = vbDefault 'Add By Sindy 2012/1/30
   cmdPrint.Enabled = True 'Add By Sindy 2012/1/30
End Sub

Private Sub cmdSelect_Click()
 Dim i As Integer
   With MSHFlexGrid1
      .col = 0
      For i = 1 To MSHFlexGrid1.Rows - 1
         .row = i
         .Text = "v"
      Next
   End With
   cmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
Dim i As Integer
   
    MoveFormToCenter Me
    cmdPrint.Enabled = False
    Set MSHFlexGrid1.Recordset = RsTemp
    Grid
    'Add By Sindy 2011/2/11 ∑Ì®∆§H1≠Y¨∞X65299Æ…, ´h∑Ì®∆§H∏ÍÆ∆ßÔßÏ∑Ì®∆§H2°C
    For i = 1 To MSHFlexGrid1.Rows - 1
      If Left(Trim(MSHFlexGrid1.TextMatrix(i, 2)), 6) = "X65299" Then
         MSHFlexGrid1.TextMatrix(i, 2) = MSHFlexGrid1.TextMatrix(i, 7)
         MSHFlexGrid1.TextMatrix(i, 3) = MSHFlexGrid1.TextMatrix(i, 8)
      End If
    Next i
    '2011/2/11 End
'*****************
'¶L™Ì≥]©w
'*****************
    SeekPrintL = Printer.Orientation
    PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint 'Modified by Morgan 2017/11/9 ≥]©w¶L™Ìæ˜ßÔ©I•s§Ω•Œ®Áº∆,≠Ïµ{¶°≤æ∞£

End Sub

Private Sub Grid()
   With MSHFlexGrid1
      .Visible = False
      .Cols = 7 '9
      .row = 0
      .col = 0: .ColWidth(0) = 300: .Text = "v" '
      .col = 1: .ColWidth(1) = 1000: .Text = "¥º≈v§H≠˚"
      .col = 2: .ColWidth(2) = 1200: .Text = "´»§·Ωs∏π"
      .col = 3: .ColWidth(3) = 3000: .Text = "´»§·¶W∫Ÿ"
      .col = 4: .ColWidth(4) = 1000: .Text = "∏u•Ù®Ï¥¡§È"
      '  91.08.07   nick ¡Ù¬√ƒÊ¶Ï •ª©“Æ◊∏π
      .col = 5: .ColWidth(5) = 0: .Text = ""
        'Add By Cheng 2003/04/18
        '¡Ùê¯ƒÊ¶Ï--∑~∞»∞œΩs∏π
      .col = 6: .ColWidth(6) = 0: .Text = ""
      '    end
'      'Add By Sindy 2011/2/11
'      .col = 7: .ColWidth(7) = 0: .Text = "∑Ì®∆§H2"
'      .col = 8: .ColWidth(8) = 0: .Text = "∑Ì®∆§H2´»§·¶W∫Ÿ"
'      '2011/2/11 End
      .Visible = True
   End With
    'Add By Cheng 2003/04/17
    Me.Frame2.Caption = "¶aß}±¯µßº∆°G¶@ " & Me.MSHFlexGrid1.Rows - 1 & " µß!!!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2003/04/02
    '≠Y¶L™Ìæ˜≈‹∞ , ´hßÛ∑s¶C¶L≥]©w
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    'Add By Cheng 2002/07/18
    Set frm073010 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
 Dim i As Integer
   intCols = MSHFlexGrid1.Cols - 1
   ShowBar MSHFlexGrid1, intLastRow, intCols
   With MSHFlexGrid1
      .col = 0
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
      End If
      CheckCmd
   End With
End Sub

Private Sub CheckCmd()
 Dim i As Integer, n As Integer
   With MSHFlexGrid1
      n = 0
      For i = 1 To .Rows - 1
         .row = i
         .col = 0
         If .Text = "v" Then
            cmdPrint.Enabled = True
            Exit For
         Else
            If i = .Rows - 1 Then
               cmdPrint.Enabled = False
               Exit Sub
            End If
         End If
      Next
   End With
End Sub

Private Sub SetPrint()
'   Printer.Font.Size = 12
'   Printer.Height = 2200
'   Printer.Width = 10000
'
'   With MSHFlexGrid1
'        nPageNo = 1
'        For i = 1 To .Rows - 1
'            .Row = i
'            If .TextMatrix(i, 0) = "v" Then
'                 nRow = 0
'                 Printer.CurrentX = 1000
'                     ' 90.07.12 modify by louis
'                     'Printer.CurrentY = i * 250
'                     'If IsNull(.Fields(i)) = False Then Printer.Print .Fields(i)
'                     If IsNull(.Fields(i)) = False Then
'                        If IsEmptyText(.Fields(i)) = False Then
'                           ' ªy§Â¨∞≠^§ÂÆ…§£™≈¶Ê
'                           If Text1(4) = "2" Then
'                              Printer.CurrentY = nRow * 250
'                           Else
'                              Printer.CurrentY = i * 250
'                           End If
'                           nRow = nRow + 1
'                        End If
'                     End If
'
'                     If IsNull(.Fields(i)) = False Then
'                        If IsEmptyText(.Fields(i)) = False Then
'                           Printer.Print .Fields(i)
'                        End If
'                     End If
'                  Next
'                  Printer.CurrentX = 5000
'                  'Printer.CurrentY = (i - 1) * 250
'                  If Text1(4) = "2" Then
'                     Printer.CurrentY = (nRow - 1) * 250
'                  Else
'                     Printer.CurrentY = (i - 1) * 250
'                  End If
'                  ' 90.07.12 modify by louis
'                  'Printer.Print Format(iPrint, "000000")
'                  If m_PageNo > 0 Then
'                     Printer.Print Format(m_PageNo, "000000")
'                  Else
'                     Printer.Print Format(iPrint, "000000")
'                  End If
'                  iPrint = iPrint + 1
'                  Printer.NewPage
'                  .MoveNext
'               Loop
'         End Select
End Sub
