VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050719 
   BorderStyle     =   1  '單線固定
   Caption         =   "程序人員核判表維護"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7875
   Begin VB.TextBox txtLR 
      Height          =   285
      Index           =   6
      Left            =   5490
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1350
      Width           =   795
   End
   Begin VB.TextBox txtLR 
      Height          =   285
      Index           =   5
      Left            =   1035
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1350
      Width           =   795
   End
   Begin VB.TextBox txtLR 
      Height          =   285
      Index           =   2
      Left            =   4950
      MaxLength       =   30
      TabIndex        =   1
      Top             =   690
      Width           =   390
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4410
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050719.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   3015
      Left            =   90
      TabIndex        =   7
      Top             =   2100
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      HighLight       =   0
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   4
      Left            =   4950
      TabIndex        =   18
      Top             =   1020
      Width           =   2760
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4868;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   3
      Left            =   1035
      TabIndex        =   2
      Top             =   1020
      Width           =   1185
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "2090;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   7
      Left            =   1035
      TabIndex        =   5
      Top             =   1680
      Width           =   6720
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11853;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   1035
      TabIndex        =   0
      Top             =   690
      Width           =   2040
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3598;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "XXX"
      Height          =   180
      Index           =   6
      Left            =   6300
      TabIndex        =   17
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "XXX"
      Height          =   180
      Index           =   5
      Left            =   1845
      TabIndex        =   16
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "相關號案件性質："
      Height          =   180
      Index           =   6
      Left            =   4050
      TabIndex        =   15
      Top             =   1395
      Width           =   1440
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   1395
      Width           =   900
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "可維護群組："
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1080
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "核判類別：          (1: 客戶函 2: 指示信)"
      Height          =   180
      Index           =   2
      Left            =   4050
      TabIndex        =   12
      Top             =   735
      Width           =   2955
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "判發人員："
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   1725
      Width           =   900
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   4
      Left            =   4050
      TabIndex        =   10
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label lblLR 
      AutoSize        =   -1  'True
      Caption         =   "程序人員："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   735
      Width           =   900
   End
End
Attribute VB_Name = "frm050719"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (Combo1)
'Created by Morgan 2018/7/23
Option Explicit

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_EditMode As Integer

Dim m_CurrKEY As String '目前正在顯示的key
Dim m_iPreRow As Integer '前次顯示資料列

Private Sub Combo1_Click(Index As Integer)
   If Index = 3 Then
      SetCombo4
   End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
   Dim ii As Integer
   
   If Combo1(Index).Locked Or Combo1(Index).Text = "" Then Exit Sub
   If Index = 1 Or Index = 4 Or Index = 7 Then
      If Combo1(Index).ListIndex = -1 Then
         For ii = 0 To Combo1(Index).ListCount - 1
            If GetStr(Combo1(Index).List(ii)) = GetStr(Combo1(Index)) Then
               Combo1(Index).ListIndex = ii
               Exit For
            End If
         Next
         If Combo1(Index).ListIndex = -1 Then Cancel = True
      End If
      If Index = 7 And Combo1(Index).ListIndex = -1 Then
         strExc(1) = ""
         Cancel = Not GetName(Combo1(Index).Text, strExc(1))
         Combo1(Index).Text = strExc(1)
      End If
   End If
End Sub

Private Function GetName(pStr As String, pName As String) As Boolean
   Dim strText As String
   Dim strChk As String
   
   strChk = GetStr(pStr)
   If IsNumeric(Mid(Trim(pStr), 2, 4)) Then
      strText = GetPrjSalesNM(Left(pStr, 5))
      If strText <> "" Then
         pName = Left(pStr, 5) & " " & strText
         If ChkStaffST04(Left(pStr, 5)) = False Then
            GetName = True
         End If
      End If
   Else
      strText = GetPrjSalesNM_2(pStr)
      If strText <> "" Then
         pName = strText & " " & pStr
         If ChkStaffST04(strText) = False Then
            GetName = True
         End If
      End If
   End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
               
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢, 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyF4, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

SetCombo

GetData
ShowFirstRecord
SetCtrlReadOnly True
UpdateToolbarState

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050719 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   
   arrGridHeadText = Array("程序人員", "核判類別", "系統別", "申請國家", "案件性質", "相關號案件性質", "判發人")
   arrGridHeadWidth = Array(800, 800, 650, 1000, 1600, 1600, 800)
   Grd1.Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To Grd1.Cols - 1
      With Grd1
      .TextMatrix(0, iCol) = arrGridHeadText(iCol)
      .ColWidth(iCol) = arrGridHeadWidth(iCol)
      .ColAlignmentFixed(iCol) = flexAlignCenterCenter
      If iCol <> 4 And iCol <> 5 Then
         .ColAlignment(iCol) = flexAlignCenterCenter
      End If
      End With
   Next
   m_iPreRow = 0
End Sub

Private Sub grd1_SelChange()
   Dim TmpRow As Integer
   
   TmpRow = Grd1.MouseRow
   Grd1.col = 0
   If TmpRow <> 0 Then
       
       Grd1.Recordset.MoveFirst
       If TmpRow > 1 Then
         Grd1.Recordset.Move TmpRow - 1
       End If
       FormShow
   End If
End Sub

Private Sub ChgToNowData()
   With Grd1.Recordset
   Do While Not .EOF
      If .Fields("LR01") = GetStr(Combo1(1)) Then
         If .Fields("LR02") = txtLR(2) Then
            If .Fields("LR03") = Combo1(3) Then
               If .Fields("LR04") = IIf(Combo1(4) = "", "*", GetStr(Combo1(4))) Then
                  If .Fields("LR05") = IIf(txtLR(5) = "", "*", txtLR(5)) Then
                     If .Fields("LR06") = IIf(txtLR(6) = "", "*", txtLR(6)) Then
                        If "" & .Fields("LR07") = GetStr(Combo1(7)) Then
                           m_CurrKEY = Grd1.Recordset.AbsolutePosition
                           ChgGrdData Val(m_CurrKEY)
                           Exit Do
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      .MoveNext
   Loop
   End With
End Sub


Private Sub ChgGrdData(iRow As Integer)
   Dim iCol As Integer
   
   If m_iPreRow > 0 And m_iPreRow <> iRow Then
      Grd1.Visible = False
      Grd1.row = m_iPreRow
      For iCol = 0 To Grd1.Cols - 1
          Grd1.col = iCol
          Grd1.CellBackColor = QBColor(15)
      Next
   End If
      
   Grd1.row = iRow
   For iCol = 0 To Grd1.Cols - 1
       Grd1.col = iCol
       Grd1.CellBackColor = &HFFC0C0
   Next
   
   m_iPreRow = Grd1.row
   
   If Grd1.TopRow > iRow Or Grd1.TopRow + 8 < iRow Then
      Grd1.TopRow = iRow
   End If
   Grd1.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub ClearField()
   Dim oText As TextBox
   For Each oText In txtLR
      oText = ""
   Next
   Combo1(1).ListIndex = -1
   Combo1(3).ListIndex = -1
   Combo1(4).ListIndex = -1
   lbl1(5) = ""
   lbl1(6) = ""
   Combo1(7) = ""
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         Combo1(1).SetFocus
         
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         UpdateToolbarState
         Combo1(7).Locked = False
         Combo1(7).SetFocus
         
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         Combo1(1).SetFocus
         
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbNo Then
                  Exit Sub
               End If
         End Select
         
         m_EditMode = 0
         FormShow
         SetCtrlReadOnly True
         UpdateToolbarState
         CloseIme
         
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub FormShow()
   Dim oText As TextBox
   Dim ii As Integer
   
   If Grd1.Recordset.BOF Or Grd1.Recordset.EOF Then Exit Sub
   
   ClearField
   
   For Each oText In txtLR
      If Grd1.Recordset.Fields("LR" & Format(oText.Index, "00")) = "*" Then
         oText.Text = ""
      Else
         oText.Text = "" & Grd1.Recordset.Fields("LR" & Format(oText.Index, "00"))
      End If
   Next
   
   '程序人員
   For ii = 0 To Combo1(1).ListCount - 1
      If GetStr(Combo1(1).List(ii)) = Grd1.Recordset.Fields("LR01") Then
         Combo1(1).ListIndex = ii
         Exit For
      End If
   Next
   If Combo1(1).ListIndex = -1 Then
      Combo1(1) = "" & Grd1.Recordset.Fields("LR01") & " " & Grd1.Recordset.Fields("程序人員")
   End If
   
   Combo1(3) = "" & Grd1.Recordset.Fields("LR03")
   
   '申請國家
   If Grd1.Recordset.Fields("LR04") = "*" Then
      Combo1(4).ListIndex = -1
   Else
      For ii = 0 To Combo1(4).ListCount - 1
         If GetStr(Combo1(4).List(ii)) = Grd1.Recordset.Fields("LR04") Then
            Combo1(4).ListIndex = ii
            Exit For
         End If
      Next
      If Combo1(4).ListIndex = -1 Then
         Combo1(4) = "" & Grd1.Recordset.Fields("LR04") & " " & Grd1.Recordset.Fields("申請國家")
      End If
   End If
   
   lbl1(5) = "" & Grd1.Recordset.Fields("案件性質")
   lbl1(6) = "" & Grd1.Recordset.Fields("相關號案件性質")
   
   '判發人
   If IsNull(Grd1.Recordset.Fields("LR07")) Then
      Combo1(7).ListIndex = -1
   Else
      For ii = 0 To Combo1(7).ListCount - 1
         If GetStr(Combo1(7).List(ii)) = Grd1.Recordset.Fields("LR07") Then
            Combo1(7).ListIndex = ii
            Exit For
         End If
      Next
      If Combo1(7).ListIndex = -1 Then
         Combo1(7) = "" & Grd1.Recordset.Fields("LR07") & " " & Grd1.Recordset.Fields("判發人")
      End If
   End If
   
   m_CurrKEY = Grd1.Recordset.AbsolutePosition
   ChgGrdData Val(m_CurrKEY)
End Sub
' 顯示第一筆資料
Private Sub ShowFirstRecord()
   With Grd1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      FormShow
   End If
   End With
End Sub
' 顯示上一筆資料
Private Sub ShowPrevRecord()
   With Grd1.Recordset
   If .RecordCount > 0 Then
      .MovePrevious
      If Not .BOF Then
         FormShow
      Else
         .MoveFirst
         MsgBox "已經是第一筆！", vbInformation
      End If
   End If
   End With
End Sub
' 顯示下一筆資料
Private Sub ShowNextRecord()
   With Grd1.Recordset
   If .RecordCount > 0 Then
      .MoveNext
      If Not .EOF Then
         FormShow
      Else
         .MoveLast
         MsgBox "已經是最後一筆！", vbInformation
      End If
   End If
   End With
End Sub
' 顯示最後一筆資料
Private Sub ShowLastRecord()
   With Grd1.Recordset
   If .RecordCount > 0 Then
      .MoveLast
      FormShow
   End If
   End With
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            Toolbar1.Buttons(1).Enabled = True
         Else
            Toolbar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            Toolbar1.Buttons(2).Enabled = True
         Else
            Toolbar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            Toolbar1.Buttons(3).Enabled = True
         Else
            Toolbar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(4).Enabled = True
         Else
            Toolbar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(8).Enabled = True
            Toolbar1.Buttons(9).Enabled = True
         Else
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
            Toolbar1.Buttons(8).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
         End If
         Toolbar1.Buttons(11).Enabled = False
         Toolbar1.Buttons(12).Enabled = False
         Toolbar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(2).Enabled = False
         Toolbar1.Buttons(3).Enabled = False
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Toolbar1.Buttons(7).Enabled = False
         Toolbar1.Buttons(8).Enabled = False
         Toolbar1.Buttons(9).Enabled = False
         Toolbar1.Buttons(11).Enabled = True
         Toolbar1.Buttons(12).Enabled = True
         Toolbar1.Buttons(14).Enabled = False
   End Select
   
End Sub
' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   Dim oText As TextBox
   For Each oText In txtLR
      oText.Locked = bEnable
   Next
   Combo1(1).Locked = bEnable
   Combo1(3).Locked = bEnable
   Combo1(4).Locked = bEnable
   Combo1(7).Locked = bEnable
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If Not Combo1(1).Locked Then
      If Combo1(1).ListIndex = -1 Then
         MsgBox "請選擇程序人員！", vbExclamation
         Combo1(1).SetFocus
         Exit Function
      End If
   End If
   
   If Not txtLR(2).Locked Then
      If txtLR(2) = "" Then
         MsgBox "請輸入核判類別！", vbExclamation
         txtLR(2).SetFocus
         Exit Function
      End If
   End If
   
   If Not Combo1(3).Locked Then
      If Combo1(3).ListIndex = -1 Then
         MsgBox "系統別不可空白！", vbExclamation
         Combo1(3).SetFocus
         Exit Function
      End If
   End If
   
   If Not Combo1(4).Locked Then
      If Combo1(3) = "P" And Combo1(4) = "" Then
         MsgBox "請選擇申請國家！", vbExclamation
         Combo1(4).SetFocus
         Exit Function
      ElseIf Combo1(4).ListIndex = -1 Then
         Combo1_Validate 4, bCancel
         If bCancel Then
            MsgBox "申請國家選擇錯誤！", vbExclamation
            Combo1(4).SetFocus
            Exit Function
         End If
      End If
   End If
   
   If Not txtLR(5).Locked Then
      If txtLR(5) <> "" And lbl1(5) = "" Then
         MsgBox "案件性質輸入錯誤！", vbExclamation
         txtLR(5).SetFocus
         Exit Function
      End If
   End If
   
   If Not txtLR(6).Locked Then
      If txtLR(6) <> "" And lbl1(6) = "" Then
         MsgBox "相關號案件性質輸入錯誤！", vbExclamation
         txtLR(6).SetFocus
         Exit Function
      End If
   End If
   
   If Not Combo1(7).Locked Then
      If Combo1(7).ListIndex = -1 Then
         Combo1_Validate 7, bCancel
         If bCancel Then
            MsgBox "判發人員輸入錯誤！", vbExclamation
            Combo1(7).SetFocus
            Exit Function
         End If
      End If
   End If
   
   If GetStr(Combo1(1)) = GetStr(Combo1(7)) Then
      MsgBox "自行判發不必設定判發人員！", vbExclamation
      Combo1(7).SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If Not TxtValidate Then Exit Function
         If Not AddRecord Then Exit Function
         If GetData Then
            ChgToNowData
         End If
            
      Case 2: '修改
         If Not TxtValidate Then Exit Function
         If Not ModRecord Then Exit Function
         If GetData Then
            ChgToNowData
         End If
         
      Case 3: '刪除
         If Not DelRecord Then Exit Function
         If GetData Then
            ShowFirstRecord
         End If
         
      Case 4: '查詢
         If GetData(False) = True Then
            FormShow
         Else
            Exit Function
         End If
         
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

Private Function AddRecord() As Boolean
   Dim stValues As String
   
   strSql = "insert into LETTERREVIEWER(LR01,LR02,LR03"
   stValues = " values ('" & GetStr(Combo1(1)) & "','" & txtLR(2) & "','" & Combo1(3) & "'"
   If Combo1(4) <> "" Then
      strSql = strSql & ",LR04"
      stValues = stValues & ",'" & GetStr(Combo1(4)) & "'"
   End If
   If txtLR(5) <> "" Then
      strSql = strSql & ",LR05"
      stValues = stValues & ",'" & txtLR(5) & "'"
   End If
   If txtLR(6) <> "" Then
      strSql = strSql & ",LR06"
      stValues = stValues & ",'" & txtLR(6) & "'"
   End If
   stValues = stValues & ",'" & GetStr(Combo1(7)) & "')"
   
   strSql = strSql & ",LR07)" & stValues
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

Private Function ModRecord() As Boolean
   strSql = "UPDATE LETTERREVIEWER SET LR07='" & GetStr(Combo1(7)) & "'"
   strSql = strSql & " WHERE LR01='" & GetStr(Combo1(1)) & "' AND LR02='" & txtLR(2) & "' AND LR03='" & Combo1(3) & "'"
   If Combo1(4) = "" Then
      strSql = strSql & " AND LR04='*'"
   Else
      strSql = strSql & " AND LR04='" & GetStr(Combo1(4)) & "'"
   End If
   If txtLR(5) = "" Then
      strSql = strSql & " AND LR05='*'"
   Else
      strSql = strSql & " AND LR05='" & txtLR(5) & "'"
   End If
   If txtLR(6) = "" Then
      strSql = strSql & " AND LR06='*'"
   Else
      strSql = strSql & " AND LR06='" & txtLR(6) & "'"
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
   
End Function

Private Function DelRecord() As Boolean
   strSql = "DELETE LETTERREVIEWER"
   strSql = strSql & " WHERE LR01='" & GetStr(Combo1(1)) & "' AND LR02='" & txtLR(2) & "' AND LR03='" & Combo1(3) & "'"
   If Combo1(4) = "" Then
      strSql = strSql & " AND LR04='*'"
   Else
      strSql = strSql & " AND LR04='" & GetStr(Combo1(4)) & "'"
   End If
   If txtLR(5) = "" Then
      strSql = strSql & " AND LR05='*'"
   Else
      strSql = strSql & " AND LR05='" & txtLR(5) & "'"
   End If
   If txtLR(6) = "" Then
      strSql = strSql & " AND LR06='*'"
   Else
      strSql = strSql & " AND LR06='" & txtLR(6) & "'"
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   DelRecord = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox " 刪除失敗！" & vbCrLf & Err.Description
End Function
'取得空白前面的字串
Private Function GetStr(pStr As String) As String
   Dim iPos As Integer
   iPos = InStr(pStr, " ")
   If iPos > 0 Then
      GetStr = Trim(Left(pStr, iPos))
   Else
      GetStr = Trim(pStr)
   End If
      
End Function

Private Function GetData(Optional pAll As Boolean = True) As Boolean
   Dim stCon As String
   If Not pAll Then
      If Combo1(1) <> "" Then stCon = stCon & " and lr01='" & GetStr(Combo1(1)) & "'"
      If txtLR(2) <> "" Then stCon = stCon & " and lr02='" & txtLR(2) & "'"
      If Combo1(3) <> "" Then stCon = stCon & " and lr03='" & Combo1(3) & "'"
      If Combo1(4) <> "" Then stCon = stCon & " and lr04='" & GetStr(Combo1(4)) & "'"
      If txtLR(5) <> "" Then stCon = stCon & " and lr05='" & txtLR(5) & "'"
      If txtLR(6) <> "" Then stCon = stCon & " and lr06='" & txtLR(6) & "'"
      If Combo1(7) <> "" Then stCon = stCon & " and lr07='" & GetStr(Combo1(7)) & "'"
   Else
      If txtLR(2).Tag = txtLR(2) And txtLR(2) <> "" Then stCon = stCon & " and lr02='" & txtLR(2) & "'"
      If Combo1(3).Tag = Combo1(3) And Combo1(3) <> "" Then stCon = stCon & " and lr03='" & Combo1(3) & "'"
   End If
   
   txtLR(2).Tag = txtLR(2)
   Combo1(3).Tag = Combo1(3)
   
    strExc(0) = "select decode(lr01,'00000','不指定',s1.st02) 程序人員" & _
      ",decode(lr02,'1','客戶函','2','指示信') 核判類別" & _
      ",lr03 系統別" & _
      ",decode(lr04,'*',' ','999','非台灣',na03) 申請國家" & _
      ",decode(lr05,'*',' ',decode(lr04,'000',c1.cpm03,c1.cpm04)) 案件性質" & _
      ",decode(lr06,'*',' ',decode(lr04,'000',c2.cpm03,c2.cpm04)) 相關號案件性質" & _
      ",decode(lr07,'','自判',nvl(s2.st02,lr07)) 判發人,l.*" & _
      " from LETTERREVIEWER l,staff s1,nation,casepropertymap c1,casepropertymap c2,staff s2" & _
      " where s1.st01(+)=lr01 and na01(+)=lr04" & stCon & _
      " and c1.cpm01(+)=lr03 and c1.cpm02(+)=lr05" & _
      " and c2.cpm01(+)=lr03 and c2.cpm02(+)=lr06" & _
      " and s2.st01(+)=lr07 order by lr01,lr02,lr03,lr04,lr05,lr06"
      
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
      GetData = True
    ElseIf Not pAll Then
       MsgBox "無符合資料！"
       Exit Function
    End If
    Set Grd1.Recordset = RsTemp
    SetGrd
    
End Function

Private Sub txtLR_Change(Index As Integer)
   Select Case Index
      Case 4, 5, 6
         lbl1(Index) = ""
   End Select
End Sub

Private Sub SetCombo4()
   Combo1(4).Clear
   
   If Combo1(3) <> "CFP" Then
      strExc(0) = "select na01||' '||na03 from nation where na01 in ('000','020','013','044','056') order by 1"
   Else
      strExc(0) = "select na01||' '||na03 from nation where length(na01)=3 and na01>='010' and na01<'4' and na01 not in ('020','044') order by 1"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo1(4).AddItem .Fields(0)
         .MoveNext
      Loop
      End With
   End If
   
   If Combo1(3) <> "CFP" Then
      Combo1(4).AddItem "999 非台灣"
   End If
End Sub

Private Sub SetCombo()
   Combo1(1).Clear
   Combo1(3).Clear
   Combo1(7).Clear
   
   Combo1(3).AddItem "P"
   Combo1(3).AddItem "CFP"
   
   Combo1(1).AddItem "00000 不指定", 0

   strExc(0) = "select st01||' '||st02 C00,st01 from staff where st03='P12' and st04='1' and st01<'F' order by st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo1(1).AddItem .Fields(0)
         Combo1(7).AddItem .Fields(0)
         .MoveNext
      Loop
      End With
   End If
   
   strExc(0) = "SELECT OCODE||' '||OEXPLAIN||'('||OMAN||')' FROM SetSpecMan WHERE OCODE IN ('PS4','PS5','PS6','PS7') ORDER BY 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo1(7).AddItem RsTemp.Fields(0)
         .MoveNext
      Loop
      End With
   End If
   SetCombo4
End Sub


Private Sub txtLR_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtLR(Index)
End Sub

Private Sub txtLR_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 2 Then
      If Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txtLR_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 5, 6
         If Combo1(3) = "" Then
            strExc(1) = "P"
         Else
            strExc(1) = Combo1(3)
         End If
         
         If Index = 5 And txtLR(5) <> "" Then
            If ClsPDGetCaseProperty(strExc(1), txtLR(5), strExc(2), IIf(GetStr(Combo1(4)) = "000", False, True)) Then
               lbl1(5) = strExc(2)
            End If
         End If
         If Index = 6 And txtLR(6) <> "" Then
            If ClsPDGetCaseProperty(strExc(1), txtLR(6), strExc(2), IIf(GetStr(Combo1(4)) = "000", False, True)) Then
               lbl1(6) = strExc(2)
            End If
         End If
   End Select
End Sub
