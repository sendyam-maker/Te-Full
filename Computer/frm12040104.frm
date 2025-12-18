VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040104 
   BorderStyle     =   1  '單線固定
   Caption         =   "系統種類對照表"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7650
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2880
      Width           =   5295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3720
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
            Picture         =   "frm12040104.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040104.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
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
   Begin VB.Label Label7 
      Caption         =   "0:內對內;1:內對外;2:外對內"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "系統代號："
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "服務業務："
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "國內外：　"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "中文敘述："
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "5 專利 ; 6 商標 ; 7 法務 ; 8 顧問"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "1 專利 ; 2 商標 ; 3 法務 ; 4 顧問"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frm12040104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/10/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim Computer As New ADODB.Recordset, cp As New ADODB.Recordset
Dim EditSelect As Integer, i As Integer
Dim a(4) As String

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
       Case vbKeyF10
       EditTool (10)
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

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If EditSelect <> 4 Then
            KeyAscii = 0
            EditTool (9)
         End If
    End Select
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040104", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040104", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040104", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040104", strFind, False)
   ' Ken 90.07.16 -- End
   
MoveFormToCenter Me
If Computer.State = adStateOpen Then Computer.Close
Computer.CursorLocation = adUseClient
strExc(0) = "select sk01,sk02,sk03,sk04 from systemkind order by sk01 "
Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
Text1.Text = Computer.Fields(0).Value
Text2.Text = Computer.Fields(1).Value
Text3.Text = Computer.Fields(2).Value
Text4.Text = Computer.Fields(3).Value
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
Set frm12040104 = Nothing
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

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii = 13 And EditSelect = 4 Then
  Computer.Find "sk01='" & Text1 & "'", 0, adSearchForward, 1
  If Computer.EOF Then
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
                 strExc(0) = "select sk01,sk02,sk03,sk04 from systemkind order by sk01"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.MoveFirst
                 Text1.Text = Computer.Fields(0).Value
                 Text2.Text = Computer.Fields(1).Value
                 Text3.Text = Computer.Fields(2).Value
                 Text4.Text = Computer.Fields(3).Value
  Exit Sub
  End If
  If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
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
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If EditSelect = 1 Then
If cp.State = adStateOpen Then cp.Close
cp.CursorLocation = adUseClient
strExc(1) = "SELECT COUNT(SK01) FROM SYSTEMKIND WHERE SK01='" & Text1.Text & "'"
cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
Cancel = False
If cp.Fields(0).Value <> 0 Then
MsgBox "此系統代號已存在"
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Cancel = True
End If
cp.Close
End If
End Sub

Private Sub Text2_GotFocus()
If EditSelect = 1 Then
    If cp.State = adStateOpen Then cp.Close
    strExc(1) = "SELECT COUNT(SK01) FROM SYSTEMKIND WHERE "
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
 Cancel = False
If Not (Text2.Text = "" Or Text2.Text >= "1" And Text2.Text <= "8") Then
     MsgBox "輸入錯誤", vbInformation
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Cancel = True
End If
End Sub
Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
Cancel = False
If Not (Text3.Text = "" Or Text3.Text = "0" Or Text3.Text = "1" Or Text3.Text = "2") Then
MsgBox "輸入錯誤", vbInformation
Text3.SelStart = 0
Text3.SelLength = Len(Text3)
Cancel = True
End If
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text4.IMEMode = 1
OpenIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text4.IMEMode = 2
CloseIme
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
       Text1.Locked = False
       Text2.Locked = False
       Text3.Locked = False
       Text4.Locked = False
       Text1.Text = ""
       Text2.Text = ""
       Text3.Text = ""
       Text4.Text = ""
       Text1.SetFocus
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
       Text1.Locked = True
       Text2.Locked = False
       Text3.Locked = False
       Text4.Locked = False
       Text2.SetFocus
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
       If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
       Computer.MoveNext
       If Computer.EOF Then
       Computer.MovePrevious
       Computer.MovePrevious
       End If
       a(0) = Computer.Fields(0).Value
       cnnConnection.Execute "delete systemkind where sk01='" & Text1.Text & "'"
                 If Computer.State = adStateOpen Then Computer.Close
                 strExc(0) = "select sk01,sk02,sk03,sk04 from systemkind order by sk01"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.Find "sk01='" & a(0) & "'", 0, adSearchForward, 1
                 If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
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
       Text1.SetFocus
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
       Text1.Locked = False
       Text2.Locked = True
       Text3.Locked = True
       Text4.Locked = True
       Text1.Text = ""
       Text2.Text = ""
       Text3.Text = ""
       Text4.Text = ""
       Text1.SetFocus
       End If
       Case 5 '第一筆
       If TBar1.Buttons(6).Enabled = True Then
       Computer.MoveFirst
        If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
                 End If
       Case 6 '前一筆
       If TBar1.Buttons(7).Enabled = True Then
       Computer.MovePrevious
       If Computer.BOF Then
       DataErrorMessage (6)
       Computer.MoveFirst
       End If
        If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
                 End If
       Case 7 '後一筆
       If TBar1.Buttons(8).Enabled = True Then
       Computer.MoveNext
       If Computer.EOF Then
       DataErrorMessage (7)
       Computer.MoveLast
       End If
                 If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
                 End If
       Case 8 '最後一筆
        If TBar1.Buttons(9).Enabled = True Then
        Computer.MoveLast
        If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
                 End If
       Case 9 '確定
           If TBar1.Buttons(11).Enabled = True Then
           If Text1.Text = "" Then MsgBox "系統代號不可為空值", vbInformation: Text1.SetFocus: Exit Function
           If EditSelect = 1 Then
           If Text2.Text = "" Then MsgBox "服務業務不可為空值", vbInformation: Text2.SetFocus: Exit Function
           If Text3.Text = "" Then MsgBox "內;外不可為空值", vbInformation: Text3.SetFocus: Exit Function
           If Text4.Text = "" Then MsgBox "中文敘述不可為空值", vbInformation: Text4.SetFocus: Exit Function
           End If
           If EditSelect = 1 Then
           If cp.State = adStateOpen Then cp.Close
           cp.CursorLocation = adUseClient
           strExc(1) = "SELECT COUNT(SK01) FROM SYSTEMKIND WHERE SK01='" & Text1.Text & "'"
           cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
           If cp.Fields(0).Value <> 0 Then
           MsgBox "此系統代號已存在"
           Text1.SetFocus
           Exit Function
           End If
           cp.Close
           If Not (Text2.Text = "" Or Text2.Text >= "1" And Text2.Text <= "8") Then
           MsgBox "輸入錯誤", vbInformation
           Text2.SetFocus
           Exit Function
           End If
           If Not (Text3.Text = "" Or Text3.Text = "0" Or Text3.Text = "1" Or Text3.Text = "2") Then
           MsgBox "輸入錯誤", vbInformation
           Text3.SetFocus
           Exit Function
           End If
           End If
           Select Case EditSelect
                 Case 1
                  'Add By Cheng 2002/05/23
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Function
                 
                 cnnConnection.Execute "insert into systemkind values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
                 EditSelect = 0
                 Case 2
                  'Add By Cheng 2002/05/23
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Function
                 
                 cnnConnection.Execute "update systemkind set sk02='" & Text2.Text & "',sk03='" & Text3.Text & "',sk04='" & Text4.Text & "' where sk01='" & Text1.Text & "'"
                 EditSelect = 0
                 Case 4
                 Computer.Find "sk01='" & Text1 & "'", 0, adSearchForward, 1
                 If Computer.EOF Then
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
                 strExc(0) = "select sk01,sk02,sk03,sk04 from systemkind order by sk01"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.MoveFirst
                 Text1.Text = Computer.Fields(0).Value
                 Text2.Text = Computer.Fields(1).Value
                 Text3.Text = Computer.Fields(2).Value
                 Text4.Text = Computer.Fields(3).Value
                 Exit Function
                 Else
                 If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
                 End If
                 End Select
                 If Computer.State = adStateOpen Then Computer.Close
                 strExc(0) = "select sk01,sk02,sk03,sk04 from systemkind order by sk01"
                 Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                 Computer.Find "sk01='" & Text1.Text & "'", 0, adSearchForward, 1
                 If IsNull(Computer.Fields(0)) Then
                 Text1.Text = ""
                 Else
                 Text1.Text = Computer.Fields(0).Value
                 End If
                 If IsNull(Computer.Fields(1)) Then
                 Text2.Text = ""
                 Else
                 Text2.Text = Computer.Fields(1).Value
                 End If
                 If IsNull(Computer.Fields(2).Value) Then
                 Text3.Text = ""
                 Else
                 Text3.Text = Computer.Fields(2).Value
                 End If
                 If IsNull(Computer.Fields(3).Value) Then
                 Text4.Text = ""
                 Else
                 Text4.Text = Computer.Fields(3).Value
                 End If
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
              If MsgBox("妳並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
         If EditSelect = 1 Or EditSelect = 4 Then
         Computer.MoveFirst
         Text1.Text = Computer.Fields(0).Value
         Text2.Text = Computer.Fields(1).Value
         Text3.Text = Computer.Fields(2).Value
         Text4.Text = Computer.Fields(3).Value
         End If
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
If Me.Text1.Enabled = True Then
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text2.Enabled = True Then
   Cancel = False
   Text2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text3.Enabled = True Then
   Cancel = False
   Text3_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text4.Enabled = True Then
   Cancel = False
   Text4_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

