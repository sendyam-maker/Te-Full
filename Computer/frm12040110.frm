VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040110 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利商標種類/著作權登記項目對照表"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8085
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2016
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
            Picture         =   "frm12040110.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040110.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
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
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSForms.TextBox Text1 
      Height          =   450
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   5505
      VariousPropertyBits=   671105055
      MaxLength       =   30
      Size            =   "9710;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   450
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   3210
      Width           =   5505
      VariousPropertyBits=   671105055
      MaxLength       =   30
      Size            =   "9710;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   450
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   5505
      VariousPropertyBits=   671105055
      MaxLength       =   30
      Size            =   "9710;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   450
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   1890
      Width           =   5505
      VariousPropertyBits=   671105055
      MaxLength       =   30
      Size            =   "9710;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1380
      Width           =   255
      VariousPropertyBits=   671105055
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   810
      Width           =   255
      VariousPropertyBits=   671105055
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "英文名稱："
      Height          =   252
      Left            =   960
      TabIndex        =   11
      Top             =   3220
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "日文名稱："
      Height          =   252
      Left            =   960
      TabIndex        =   10
      Top             =   3870
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "大陸名稱："
      Height          =   252
      Left            =   960
      TabIndex        =   9
      Top             =   2570
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "國內名稱："
      Height          =   252
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "種類代號/登記項目："
      Height          =   252
      Left            =   960
      TabIndex        =   7
      Top             =   1425
      Width           =   1692
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標或專利/著作權：             (1:專利 2:商標 3:著作權 4.創新業務)"
      Height          =   180
      Left            =   972
      TabIndex        =   6
      Top             =   861
      Width           =   5064
   End
End
Attribute VB_Name = "frm12040110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/15 改成Form2.0 ; Text1(index)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim Computer As New ADODB.Recordset, cp As New ADODB.Recordset
Dim EditSelect As Integer, i As Integer
Dim a(2) As String

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Dim bolMsgEnter As Boolean 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; 因為MsgBox的Enter鍵都會觸發Toolbar的”確定KeyF9”動作反之用滑鼠就不會

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Memo by Lydia 2021/10/21 原程式搬到Form_KeyUp

End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
'Remove by Lydia 2021/10/21 改到Form_KeyUp
'    Select Case KeyAscii
'      Case vbKeyReturn:
'         If EditSelect <> 0 Then
'            KeyAscii = 0
'            Form_KeyDown vbKeyF9, 0
'         End If
'    End Select
End Sub

'Added by Lydia 2021/10/21
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Memo by Lydia 2021/10/21 從Form_KeyDown搬來
Select Case KeyCode
       Case vbKeyF2
       EditTool (1)
       Case vbKeyF3
       EditTool (2)
       Case vbKeyF5
       EditTool (3)
       Case vbKeyF4
       EditTool (4)
       Case vbKeyF9
       EditTool (9)
       'Added by Lydia 2021/10/21 從Form_KeyPress改到這裡
       'Remove by Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
       'Case vbKeyReturn:
       '   If EditSelect <> 0 Then
       '       'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; 因為MsgBox的Enter鍵都會觸發Toolbar的”確定KeyF9”動作
       ''       If bolMsgEnter = True Then
        '          bolMsgEnter = False
        '      Else
        '      'end 2021/10/25
        '          EditTool (9)
        '      End If 'Added by Lydia 2021/10/25
        '  End If
      ''end 2021/10/21
      'end 2021/11/22
      
       Case vbKeyF10
       EditTool (10)
       Case vbKeyHome
       EditTool (5)
       Case vbKeyEnd
       EditTool (8)
       Case vbKeyPageUp
       EditTool (6)
       Case vbKeyPageDown
       EditTool (7)
       Case vbKeyEscape
       EditTool (11)
End Select
   
'   ' Ken 90.07.19 -- Start
'   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
'         If m_bInsert Then
'             TBar1.Buttons(1).Enabled = True
'         Else
'             TBar1.Buttons(1).Enabled = False
'         End If
'         If m_bUpdate Then
'             TBar1.Buttons(2).Enabled = True
'         Else
'             TBar1.Buttons(2).Enabled = False
'         End If
'         If m_bDelete Then
'             TBar1.Buttons(3).Enabled = True
'         Else
'             TBar1.Buttons(3).Enabled = False
'         End If
'   End If
'   ' Ken 90.07.19 -- End
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040110", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040110", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040110", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040110", strFind, False)
   ' Ken 90.07.16 -- End
   
MoveFormToCenter Me
If Computer.State = adStateOpen Then Computer.Close
Computer.CursorLocation = adUseClient
strExc(0) = "select ptm01,ptm02,ptm03,ptm04,ptm05,ptm06,ptm01||ptm02 as test from patenttrademarkmap ORDER BY PTM01,PTM02"
Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
For i = 0 To 5
If IsNull(Computer.Fields(i).Value) Then
Text1(i) = ""
Else
Text1(i) = Computer.Fields(i).Value
End If
Next i
For i = 1 To 4
  TBar1.Buttons(i).Enabled = True
Next i
For i = 6 To 9
  TBar1.Buttons(i).Enabled = True
Next i
TBar1.Buttons(11).Enabled = False
TBar1.Buttons(12).Enabled = False
TBar1.Buttons(14).Enabled = True

   ' Ken 90.07.16 -- start
   If m_bInsert Then
       TBar1.Buttons(1).Enabled = True
   Else
       TBar1.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       TBar1.Buttons(2).Enabled = True
   Else
       TBar1.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       TBar1.Buttons(3).Enabled = True
   Else
       TBar1.Buttons(3).Enabled = False
   End If
   ' Ken 90.07.16 -- End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040110 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
       Case 1
       EditTool (1)
       Case 2
       EditTool (2)
       Case 3
       EditTool (3)
       Case 4
       EditTool (4)
       Case 6
       EditTool (5)
       Case 7
       EditTool (6)
       Case 8
       EditTool (7)
       Case 9
       EditTool (8)
       Case 11
       EditTool (9)
       Case 12
       EditTool (10)
       Case 14
       EditTool (11)
End Select

   ' Ken 90.07.16 -- Start
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      Else
          TBar1.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
      Else
          TBar1.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          TBar1.Buttons(3).Enabled = True
      Else
          TBar1.Buttons(3).Enabled = False
      End If
   End If
   ' Ken 90.07.16 -- End
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Select Case Index
      Case 0
         Text1(0).SelStart = 0
         Text1(0).SelLength = Len(Text1(0))
      Case 1
         Text1(1).SelStart = 0
         Text1(1).SelLength = Len(Text1(1))
      Case 2
         Text1(2).SelStart = 0
         Text1(2).SelLength = Len(Text1(2))
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(2).IMEMode = 1
         OpenIme
      Case 3
         Text1(3).SelStart = 0
         Text1(3).SelLength = Len(Text1(3))
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(3).IMEMode = 1
         OpenIme
      Case 4
         Text1(4).SelStart = 0
         Text1(4).SelLength = Len(Text1(4))
      Case 5
         Text1(5).SelStart = 0
         Text1(5).SelLength = Len(Text1(5))
   End Select
End Sub

'Modified by Lydia 2021/10/15 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 0
         KeyAscii = UpperCase(KeyAscii)
      Case 1
         If KeyAscii = 13 And EditSelect = 4 Then
            Computer.Find "test='" & Text1(0) + Text1(1) & "'", 0, adSearchForward, 1
            If Computer.EOF Then
               'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
               MsgBox "找不到搜尋之資料", vbInformation
               If Computer.State = adStateOpen Then Computer.Close
               strExc(0) = "select ptm01,ptm02,ptm03,ptm04,ptm05,ptm06,ptm01||ptm02 as test from patenttrademarkmap ORDER BY PTM01,PTM02"
               Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
               Computer.MoveFirst
               For i = 0 To 5
                  If IsNull(Computer.Fields(i)) Then
                     Text1(i).Text = ""
                  Else
                     Text1(i).Text = Computer.Fields(i).Value
                  End If
               Next i
               For i = 1 To 4
                  TBar1.Buttons(i).Enabled = True
               Next i
               For i = 6 To 9
                  TBar1.Buttons(i).Enabled = True
               Next i
               TBar1.Buttons(11).Enabled = False
               TBar1.Buttons(12).Enabled = False
               TBar1.Buttons(14).Enabled = True
               EditSelect = 0
               Exit Sub
            End If
            For i = 0 To 5
               If IsNull(Computer.Fields(i)) Then
                  Text1(i).Text = ""
               Else
                  Text1(i).Text = Computer.Fields(i).Value
               End If
            Next i
            For i = 1 To 4
               TBar1.Buttons(i).Enabled = True
            Next i
            For i = 6 To 9
               TBar1.Buttons(i).Enabled = True
            Next i
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         'modify by sonia 2019/7/24 +4創新業務
         If Not (Text1(0) >= "1" And Text1(0) <= "4") Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "輸入錯誤"
            'Remove by Lydia 2021/10/25 造成重複彈訊息
            'Remark by Lydia 2021/11/22 恢復控制
            Text1(0).SetFocus
            Text1(0).SelStart = 0
            Text1(0).SelLength = Len(Text1(0))
            'end 2021/10/25
            Cancel = True
         Else
            Cancel = False
         End If
      Case 1
         If EditSelect = 1 Then
            If cp.State = adStateOpen Then cp.Close
            strExc(1) = "select count(ptm01) from patenttrademarkmap where ptm01='" & Text1(0) & "' and ptm02='" & Text1(1) & "'"
            cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
            If cp.Fields(0).Value <> "0" Then
               'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
               MsgBox "此種類代號已存在"
               'Remove by Lydia 2021/10/25 造成重複彈訊息
               'Remark by Lydia 2021/11/22 恢復控制
               Text1(0).SetFocus
               Text1(0).SelStart = 0
               Text1(0).SelLength = Len(Text1(0))
               'end 2021/10/25
               Cancel = True
            Else
               Cancel = False
            End If
         End If
      Case 2
         'modify by sonia 2019/8/6 改長度20->30
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If CheckLengthIsOK(Text1(2).Text, 30) = False Then
            'Remove by Lydia 2021/10/25 造成重複彈訊息
            'Remark by Lydia 2021/11/22 恢復控制
            Text1(2).SetFocus
            Text1(2).SelStart = 0
            Text1(2).SelLength = Len(Text1(2))
            'end 2021/10/25
            Cancel = True
         Else
            Cancel = False
         End If
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(2).IMEMode = 2
         If Cancel = False Then CloseIme
      Case 3
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If CheckLengthIsOK(Text1(3).Text, 20) = False Then
            'Remove by Lydia 2021/10/25 造成重複彈訊息
            'Remark by Lydia 2021/11/22 恢復控制
            Text1(3).SetFocus
            Text1(3).SelStart = 0
            Text1(3).SelLength = Len(Text1(3))
            'end 2021/10/25
            Cancel = True
         Else
            Cancel = False
         End If
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(3).IMEMode = 2
         If Cancel = False Then CloseIme
      Case 4
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If CheckLengthIsOK(Text1(4).Text, 40) = False Then
            'Remove by Lydia 2021/10/25 造成重複彈訊息
            'Remark by Lydia 2021/11/22 恢復控制
            Text1(4).SetFocus
            Text1(4).SelStart = 0
            Text1(4).SelLength = Len(Text1(4))
            'end 2021/10/25
            Cancel = True
         Else
            Cancel = False
         End If
      Case 5
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If CheckLengthIsOK(Text1(5).Text, 20) = False Then
            'Remove by Lydia 2021/10/25 造成重複彈訊息
            'Remark by Lydia 2021/11/22 恢復控制
            Text1(5).SetFocus
            Text1(5).SelStart = 0
            Text1(5).SelLength = Len(Text1(5))
            'end 2021/01/25
            Cancel = True
         Else
            Cancel = False
         End If
   End Select
End Sub

Private Function EditTool(Index As Integer)
Select Case Index
       Case 1 '新增
       If TBar1.Buttons(1).Enabled = True Then
       EditSelect = 1
       For i = 1 To 4
       TBar1.Buttons(i).Enabled = False
       Next i
       For i = 6 To 9
       TBar1.Buttons(i).Enabled = False
       Next i
       TBar1.Buttons(11).Enabled = True
       TBar1.Buttons(12).Enabled = True
       TBar1.Buttons(14).Enabled = False
       For i = 0 To 5
       Text1(i).Locked = False
       Next i
       For i = 0 To 5
       Text1(i).Text = ""
       Next i
       Text1(0).SetFocus
       End If
       Case 2 '修改
       If TBar1.Buttons(2).Enabled = True Then
       EditSelect = 2
       For i = 1 To 4
       TBar1.Buttons(i).Enabled = False
       Next i
       For i = 6 To 9
       TBar1.Buttons(i).Enabled = False
       Next i
       TBar1.Buttons(11).Enabled = True
       TBar1.Buttons(12).Enabled = True
       TBar1.Buttons(14).Enabled = False
       Text1(0).Locked = True
       Text1(1).Locked = True
       For i = 2 To 5
       Text1(i).Locked = False
       Next i
       Text1(2).SetFocus
       End If
       Case 3 '刪除
       If TBar1.Buttons(3).Enabled = True Then
       EditSelect = 3
       For i = 1 To 4
       TBar1.Buttons(i).Enabled = False
       Next i
       For i = 6 To 9
       TBar1.Buttons(i).Enabled = False
       Next i
       TBar1.Buttons(11).Enabled = True
       TBar1.Buttons(12).Enabled = True
       TBar1.Buttons(14).Enabled = False
       'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
       If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
       Computer.MoveNext
       If Computer.EOF Then
       Computer.MovePrevious
       Computer.MovePrevious
       End If
       a(0) = Computer.Fields(0).Value
       a(1) = Computer.Fields(1).Value
       cnnConnection.Execute "delete patenttrademarkmap where ptm01='" & Text1(0) & "' and ptm02='" & Text1(1) & "'"
       If Computer.State = adStateOpen Then Computer.Close
       strExc(0) = "select ptm01,ptm02,ptm03,ptm04,ptm05,ptm06,PTM01||PTM02 AS TEST from patenttrademarkmap ORDER BY PTM01,PTM02"
       Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
       Computer.Find "TEST='" & a(0) + a(1) & "'", 0, adSearchForward, 1
       For i = 0 To 5
       If IsNull(Computer.Fields(i)) Then
         Text1(i).Text = ""
       Else
         Text1(i).Text = Computer.Fields(i).Value
       End If
       Next i
       For i = 1 To 4
          TBar1.Buttons(i).Enabled = True
       Next i
       For i = 6 To 9
          TBar1.Buttons(i).Enabled = True
       Next i
       TBar1.Buttons(11).Enabled = False
       TBar1.Buttons(12).Enabled = False
       TBar1.Buttons(14).Enabled = True
       EditSelect = 0
       Else
       For i = 1 To 4
          TBar1.Buttons(i).Enabled = True
       Next i
       For i = 6 To 9
          TBar1.Buttons(i).Enabled = True
       Next i
       TBar1.Buttons(11).Enabled = False
       TBar1.Buttons(12).Enabled = False
       TBar1.Buttons(14).Enabled = True
       EditSelect = 0
       Exit Function
       End If
       Text1(0).SetFocus
       End If
       Case 4 '查詢
       If TBar1.Buttons(4).Enabled = True Then
       EditSelect = 4
       For i = 1 To 4
       TBar1.Buttons(i).Enabled = False
       Next i
       For i = 6 To 9
       TBar1.Buttons(i).Enabled = False
       Next i
       TBar1.Buttons(11).Enabled = True
       TBar1.Buttons(12).Enabled = True
       TBar1.Buttons(14).Enabled = False
       Text1(0).Locked = False
       Text1(1).Locked = False
       For i = 2 To 5
       Text1(i).Locked = True
       Next i
       For i = 0 To 5
       Text1(i).Text = ""
       Next i
       Text1(0).SetFocus
       End If
       Case 5 '第一筆
       If TBar1.Buttons(6).Enabled = True Then
       Computer.MoveFirst
        For i = 0 To 5
        If IsNull(Computer.Fields(i)) Then
           Text1(i).Text = ""
        Else
           Text1(i).Text = Computer.Fields(i).Value
        End If
        Next i
        End If
       Case 6 '前一筆
       If TBar1.Buttons(7).Enabled = True Then
       Computer.MovePrevious
       If Computer.BOF Then
       DataErrorMessage (6)
       Computer.MoveFirst
       End If
       For i = 0 To 5
       If IsNull(Computer.Fields(i)) Then
          Text1(i).Text = ""
       Else
          Text1(i).Text = Computer.Fields(i).Value
       End If
       Next i
       End If
       Case 7 '後一筆
       If TBar1.Buttons(8).Enabled = True Then
       Computer.MoveNext
       If Computer.EOF Then
       DataErrorMessage (7)
       Computer.MoveLast
       End If
       For i = 0 To 5
       If IsNull(Computer.Fields(i)) Then
          Text1(i).Text = ""
       Else
          Text1(i).Text = Computer.Fields(i).Value
       End If
       Next i
       End If
       Case 8 '最後一筆
        If TBar1.Buttons(9).Enabled = True Then
        Computer.MoveLast
        For i = 0 To 5
        If IsNull(Computer.Fields(i)) Then
           Text1(i).Text = ""
        Else
           Text1(i).Text = Computer.Fields(i).Value
        End If
        Next i
        End If
       Case 9 '確定
           If TBar1.Buttons(11).Enabled = True Then
           If EditSelect = 1 Or EditSelect = 4 Then
             'Modified by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; bolMsgEnter = True; 造成重複彈訊息,所以去掉SetFocus
             'Remark by Lydia 2021/11/22 恢復控制
             If Text1(0).Text = "" Then MsgBox "商標或專利/著作權不可空白", vbInformation: Text1(0).SetFocus: Exit Function
             If Text1(1).Text = "" Then MsgBox "種類代號/登記項目不可空白", vbInformation: Text1(1).SetFocus: Exit Function
             'Mark by Lydia 2021/11/22
             'If Text1(0).Text = "" Then bolMsgEnter = True:  MsgBox "商標或專利/著作權不可空白", vbInformation: Exit Function
             'If Text1(1).Text = "" Then bolMsgEnter = True:  MsgBox "種類代號/登記項目不可空白", vbInformation: Exit Function
             'end 2021/10/25
           End If
           If EditSelect = 1 Then
           'Modified by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; bolMsgEnter = True; 造成重複彈訊息,所以去掉SetFocus
           'Remark by Lydia 2021/11/22 恢復控制
           If Text1(2).Text = "" Then MsgBox "國內名稱不可空白", vbInformation: Text1(2).SetFocus: Exit Function
           'Mark by Lydia 2021/11/22
           'If Text1(2).Text = "" Then bolMsgEnter = True:  MsgBox "國內名稱不可空白", vbInformation: Exit Function
           'end 2021/11/22
           If cp.State = adStateOpen Then cp.Close
           strExc(1) = "select count(ptm01) from patenttrademarkmap where ptm01='" & Text1(0) & "' and ptm02='" & Text1(1) & "'"
           cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
           If cp.Fields(0).Value <> "0" Then
              'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
              MsgBox "此種類代號已存在"
              'Remove by Lydia 2021/10/25 造成重複彈訊息
              'Remark by Lydia 2021/11/22 恢復控制
              Text1(0).SetFocus
              Text1(0).SelStart = 0
              Text1(0).SelLength = Len(Text1(0))
              'end 2021/10/25
           Exit Function
           End If
           End If
          Select Case EditSelect
                 Case 1
                  'Add By Cheng 2002/05/23
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Function
                  
                  cnnConnection.Execute "insert into patenttrademarkmap values('" & Text1(0) & "','" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "','" & Text1(5) & "')"
                 EditSelect = 0
                 Case 2
                  'Add By Cheng 2002/05/23
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Function
                  
                  cnnConnection.Execute "update patenttrademarkmap set ptm03='" & Text1(2) & "',ptm04='" & Text1(3) & "',ptm05='" & Text1(4) & "',ptm06='" & Text1(5) & "' where ptm01='" & Text1(0) & "' and ptm02='" & Text1(1) & "'"
                 EditSelect = 0
                 Case 4
                 Computer.Find "test='" & Text1(0) + Text1(1) & "'", 0, adSearchForward, 1
                 If Computer.EOF Then
                 'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
                 MsgBox "找不到搜尋之資料", vbInformation
                 For i = 1 To 4
                 TBar1.Buttons(i).Enabled = True
                 Next i
                 For i = 6 To 9
                 TBar1.Buttons(i).Enabled = True
                 Next i
                 TBar1.Buttons(11).Enabled = False
                 TBar1.Buttons(12).Enabled = False
                 TBar1.Buttons(14).Enabled = True
                 EditSelect = 0
                 If Computer.State = adStateOpen Then Computer.Close
                 strExc(0) = "select ptm01,ptm02,ptm03,ptm04,ptm05,ptm06,PTM01||PTM02 AS TEST from patenttrademarkmap order by ptm01,ptm02"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.MoveFirst
                 For i = 0 To 5
                 If IsNull(Computer.Fields(i)) Then
                 Text1(i).Text = ""
                 Else
                 Text1(i).Text = Computer.Fields(i).Value
                 End If
                 Next i
                 Exit Function
                 Else
                 For i = 0 To 5
                 If IsNull(Computer.Fields(i)) Then
                 Text1(i).Text = ""
                 Else
                 Text1(i).Text = Computer.Fields(i).Value
                 End If
                 Next i
                 End If
                 End Select
                 If Computer.State = adStateOpen Then Computer.Close
                   strExc(0) = "select ptm01,ptm02,ptm03,ptm04,ptm05,ptm06,PTM01||PTM02 AS TEST from patenttrademarkmap order by ptm01,ptm02"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.Find "TEST='" & Text1(0).Text + Text1(1).Text & "'", 0, adSearchForward, 1
                 For i = 0 To 5
                 If IsNull(Computer.Fields(i)) Then
                 Text1(i).Text = ""
                 Else
                 Text1(i).Text = Computer.Fields(i).Value
                 End If
                 Next i
                 For i = 1 To 4
                 TBar1.Buttons(i).Enabled = True
                 Next i
                 For i = 6 To 9
                 TBar1.Buttons(i).Enabled = True
                 Next i
                 TBar1.Buttons(11).Enabled = False
                 TBar1.Buttons(12).Enabled = False
                 TBar1.Buttons(14).Enabled = True
                 EditSelect = 0
                 End If
       Case 10 '取消
            If TBar1.Buttons(12).Enabled = True Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            If MsgBox("妳並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
               If EditSelect = 1 Or EditSelect = 4 Then
                  Computer.MoveFirst
                  For i = 0 To 5
                    If IsNull(Computer.Fields(i).Value) Then
                         Text1(i).Text = ""
                    Else
                       Text1(i).Text = Computer.Fields(i).Value
                    End If
                  Next i
                  For i = 1 To 4
                       TBar1.Buttons(i).Enabled = True
                  Next i
                  For i = 6 To 9
                       TBar1.Buttons(i).Enabled = True
                  Next i
                  TBar1.Buttons(11).Enabled = False
                  TBar1.Buttons(12).Enabled = False
                   TBar1.Buttons(14).Enabled = True
               Else
                  Exit Function
               End If
               EditSelect = 0
            End If
         End If
       Case 11 '離開
       If TBar1.Buttons(14).Enabled = True Then
         Unload Me
       End If
End Select
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

