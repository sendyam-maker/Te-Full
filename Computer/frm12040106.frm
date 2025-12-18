VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040106 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工等級檔維護"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7515
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   0
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
            Picture         =   "frm12040106.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040106.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工等級編號"
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工等級描述"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   1080
   End
End
Attribute VB_Name = "frm12040106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/10/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim RcMain As New ADODB.Recordset, cp As New ADODB.Recordset
Dim TmpField(0 To 1) As String, ActionEdit As Integer
Dim Bmk As Variant

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF2
        Text1(0).SetFocus
        RcEdit 0
        Case vbKeyF3
        Text1(1).SetFocus
        Text1(0).TabStop = False
        RcEdit 1
        Case vbKeyF5
        RcEdit 2
        Case vbKeyF4
        RcEdit 5
        Case vbKeyHome
        ActionRc 0
        Case vbKeyPageUp
        ActionRc 1
        Case vbKeyPageDown
        ActionRc 2
        Case vbKeyEnd
        ActionRc 3
        Case vbKeyF9
        If Text1(0).Text = "" Then MsgBox "員工等級編號不可為空值", vbInformation: Exit Sub
        If ActionEdit = 0 Or ActionEdit = 1 Then
        If Text1(1).Text = "" Then MsgBox "員工等級描述不可為空值", vbInformation: Exit Sub
        End If
        RcEdit 3
        RcMain.ReQuery
        RcMain.Find "sl01='" & Text1(0) & "'", 0, adSearchForward, 1
        Text1(0).TabStop = True
        Case vbKeyF10
         If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
        RcEdit 4
        Text1(0).TabStop = True
        End If
        Case vbKeyEscape
        Unload Me
        Set frm12040105 = Nothing
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
         If ActionEdit <> 3 Then
            KeyAscii = 0
            Form_KeyDown vbKeyF9, 0
         End If
    End Select
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040106", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040106", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040106", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040106", strFind, False)
   ' Ken 90.07.16 -- End
   
   MoveFormToCenter Me
   If RcMain.State = adStateOpen Then RcMain.Close
   strExc(0) = "SELECT SL01,SL02 FROM STAFF_LEVEL ORDER BY SL01"
   RcMain.CursorType = adOpenDynamic
   RcMain.CursorLocation = adUseClient
   RcMain.LockType = adLockBatchOptimistic
   RcMain.Open strExc(0), cnnConnection
   If Not RcMain.BOF Then ActionRc 0
   TxtSitu True
   ActionEdit = 3
   
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

Private Sub ActionRc(ByVal Sty As Integer)
   TxtLock 2
   If RcMain.EOF And RcMain.BOF Then MsgBox "資料庫內無資料 !", vbInformation: Exit Sub
   With RcMain
      Select Case Sty
         Case 0
            .MoveFirst
         Case 1
            .MovePrevious
            If .BOF Then
               Beep
               MsgBox "巳是第一筆了 ! ", vbInformation
               .MoveFirst
            End If
         Case 2
            .MoveNext
            If .EOF Then
               Beep
               MsgBox "巳是最後一筆了 ! ", vbInformation
               .MoveLast
            End If
         Case 3
            .MoveLast
      End Select
   End With
   SetTxtValue
End Sub

Private Sub SetTxtValue()
 Dim i As Integer, j As Integer
   For i = 0 To 1
      If IsNull(RcMain.Fields(i).Value) = False Then
         Text1(i).Text = RcMain.Fields(i).Value
      End If
   Next
End Sub

Private Sub RcEdit(Situ As Integer)
 Dim i As Integer
   Select Case Situ
      Case 0 'add
         TxtSitu False
         ActionEdit = 0
         TxtLock 2
         TextInverse Text1(0)
      Case 1 'modi
         TxtSitu False
         ActionEdit = 1
         For i = 0 To 1
            TmpField(i) = Text1(i).Text
         Next
         Text1(0).Locked = True
      Case 2 'delete
          If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            RcMain.Delete
            RcMain.UpdateBatch
            If RcMain.EOF = True Then
               ActionRc 1
            Else
               ActionRc 2
            End If
         End If
      Case 3 'update
         If ActionEdit = 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
         
            RcMain.AddNew
            If GetVal = False Then Exit Sub
            ActionRc 3
         ElseIf ActionEdit = 1 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            If GetVal = False Then Exit Sub
         ElseIf ActionEdit = 2 Then
            RcMain.Find "SL01='" & Text1(0).Text & "'", 0, adSearchForward, 1
            If RcMain.EOF Then
               MsgBox "無此記錄之資料 !", vbCritical
               RcMain.Bookmark = Bmk
            End If
            SetTxtValue
         End If
         TxtSitu True
         ActionEdit = 3
      Case 4 'cancel
         TxtSitu True
         If ActionEdit = 0 Then
            ActionRc 3
         ElseIf ActionEdit = 1 Then
            For i = 0 To 1
               Text1(i).Text = TmpField(i)
            Next
         ElseIf ActionEdit = 2 Then
            RcMain.Bookmark = Bmk
            SetTxtValue
         End If
         ActionEdit = 3
      Case 5 'query
         Bmk = RcMain.Bookmark
         TxtSitu False
         TxtLock 2
         Text1(1).Locked = True
         ActionEdit = 2
         Text1(0).SetFocus
   End Select
End Sub

Private Function GetVal() As Boolean
 Dim i As Integer
On Error GoTo ErrHand
   For i = 0 To 1
      If Text1(i).Text <> "" Then
         RcMain.Fields(i).Value = Text1(i).Text
      Else
         RcMain.Fields(i).Value = Null
      End If
   Next
   RcMain.UpdateBatch
   GetVal = True
   Exit Function
ErrHand:
   GetVal = False
   RcMain.CancelUpdate
   RcMain.ReQuery
   If Err.Number = -2147217887 Then
      MsgBox "資料錯誤，新增失敗 !", vbInformation
   Else
      MsgBox "錯誤 : " & Err.Description, vbInformation
   End If
End Function

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As TextBox, i As Integer
   Select Case Lt
      Case 0
         For Each txt In frm12040106.Text1
            txt.Locked = True
         Next
      Case 1
         For Each txt In frm12040106.Text1
            txt.Locked = False
         Next
      Case 2
         For Each txt In frm12040106.Text1
            txt.Text = ""
         Next
   End Select
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040106 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         Text1(0).SetFocus
         RcEdit 0
      Case 2
         Text1(1).SetFocus
         Text1(0).TabStop = False
         RcEdit 1
      Case 3
         RcEdit 2
      Case 4
         RcEdit 5
      Case 6
         ActionRc 0
      Case 7
         ActionRc 1
      Case 8
         ActionRc 2
      Case 9
         ActionRc 3
      Case 11
        If Text1(0).Text = "" Then MsgBox "員工等級編號不可為空值", vbInformation: Exit Sub
        If ActionEdit = 0 Or ActionEdit = 1 Then
        If Text1(1).Text = "" Then MsgBox "員工等級描述不可為空值", vbInformation: Exit Sub
        End If
         RcEdit 3
          RcMain.ReQuery
        RcMain.Find "sl01='" & Text1(0) & "'", 0, adSearchForward, 1
         Text1(0).TabStop = True
      Case 12
         If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
        RcEdit 4
        Text1(0).TabStop = True
        End If
      Case 14
         Unload Me
         Set frm12040106 = Nothing
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
   TextInverse Text1(Index)
   Select Case Index
          Case 1
          'edit by nickc 2007/07/11 切換輸入法改用API
          'Text1(1).IMEMode = 1
          OpenIme
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
          Case 0
          KeyAscii = UpperCase(KeyAscii)
   End Select
   If Index = 0 Then
      If KeyAscii = 13 And ActionEdit = 2 Then
         RcEdit 3
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
       Case 0
       If ActionEdit = 0 Then
       If cp.State = adStateOpen Then cp.Close
       strExc(1) = "select count(sl01) from staff_level where sl01='" & Text1(0) & "'"
       cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
       If cp.Fields(0).Value <> "0" Then
       MsgBox "員工等級編號已存在"
       Text1(0).SetFocus
       Text1(0).SelStart = 0
       Text1(0).SelLength = Len(Text1(0))
       Cancel = True
       Else
       Cancel = False
       End If
       End If
       Case 1
       'edit by nickc 2007/07/11 切換輸入法改用API
       'Text1(1).IMEMode = 2
       If CheckLengthIsOK(Text1(1).Text, 30) = False Then
       Text1(1).SetFocus
       Text1(1).SelStart = 0
       Text1(1).SelLength = Len(Text1(1))
       Cancel = True
       Else
       Cancel = False
       End If
       'add by nickc 2007/07/11 切換輸入法改用API
       If Cancel = False Then CloseIme
End Select
End Sub

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

