VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040146 
   BorderStyle     =   1  '單線固定
   Caption         =   "特殊專利商標資料維護"
   ClientHeight    =   3216
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7644
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3216
   ScaleWidth      =   7644
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   0
      Top             =   922
      Width           =   360
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1368
      Width           =   360
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2448
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1812
      Width           =   4884
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6750
      Top             =   780
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040146.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7644
      _ExtentX        =   13483
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
   Begin MSForms.TextBox textSPT04 
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   2232
      Width           =   4884
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "8615;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "特殊專利商標代碼："
      Height          =   252
      Left            =   420
      TabIndex        =   9
      Top             =   1390
      Width           =   1620
   End
   Begin VB.Label Label6 
      Caption         =   "大陸名稱："
      Height          =   228
      Left            =   420
      TabIndex        =   8
      Top             =   2280
      Width           =   1368
   End
   Begin VB.Label Label5 
      Caption         =   "特殊專利商標代碼名稱：　"
      Height          =   228
      Left            =   420
      TabIndex        =   7
      Top             =   1848
      Width           =   1992
   End
   Begin VB.Label Label3 
      Caption         =   "專利或商標："
      Height          =   225
      Left            =   420
      TabIndex        =   6
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "1 專利 ; 2 商標"
      Height          =   252
      Left            =   2496
      TabIndex        =   4
      Top             =   946
      Width           =   1212
   End
End
Attribute VB_Name = "frm12040146"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/11/14 增加textSPT04=>Form2.0元件
'Memo By Sonia 2021/12/10 Form2.0不用改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim Computer As New ADODB.Recordset, cp As New ADODB.Recordset
Dim EditSelect As Integer 'Memo by Lydia 2023/11/14 0:瀏覽 1:新增 2:修改 3:刪除 4:查詢


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
   ' Ken 90.07.19 -- Start
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
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
   ' Ken 90.07.19 -- End
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040146", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040146", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040146", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040146", strFind, False)
   ' Ken 90.07.16 -- End
   
MoveFormToCenter Me
If Computer.State = adStateOpen Then Computer.Close
Computer.CursorLocation = adUseClient
'Modified by Lydia 2023/11/14 +SPT04
strExc(0) = "Select SPT01, SPT02, SPT03, SPT04 From SpecialPatentTrademark Order By 1, 2 "
Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
Text2.Text = "" & Computer.Fields(0).Value
Text3.Text = "" & Computer.Fields(1).Value
Text4.Text = "" & Computer.Fields(2).Value
textSPT04.Text = "" & Computer.Fields(3).Value 'Added by Lydia 2023/11/14

For intI = 1 To 4
  TBar1.Buttons(intI).Enabled = True
Next intI
For intI = 6 To 9
  TBar1.Buttons(intI).Enabled = True
Next intI
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
Set frm12040146 = Nothing
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

Private Sub Text2_GotFocus()
If EditSelect = 1 Then
    'Modified by Lydia 2023/11/14 改模組
    TextInverse Text2
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Cancel = False
If Not (Text2.Text = "" Or Text2.Text >= "1" And Text2.Text <= "2") Then
    MsgBox "輸入錯誤", vbInformation
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Cancel = True
End If
End Sub
Private Sub Text3_GotFocus()
'Modified by Lydia 2023/11/14 改模組
TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
Cancel = False

'Modified by Lydia 2023/11/14 以日期區分
'If Not (Text3.Text = "" Or (Me.Text3.Text >= 1 And Me.Text3.Text <= 9)) Then
If strSrvDate(1) > "20231114" Then
  If Not (Text3.Text = "" Or (Me.Text3.Text >= "A" And Me.Text3.Text <= "Z")) Then
     Cancel = True
  End If
Else
  If Not (Text3.Text = "" Or (Me.Text3.Text >= 1 And Me.Text3.Text <= 9)) Then
     Cancel = True
  End If
End If
If Cancel = True Then
'end 2023/11/14
    MsgBox "輸入錯誤", vbInformation
    TextInverse Text3 'Modified by Lydia 2023/11/14 改模組
    Cancel = True
End If
End Sub

Private Sub Text4_GotFocus()
'Modified by Lydia 2023/11/14 改模組
TextInverse Text4
OpenIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If CheckLengthIsOK(Me.Text4.Text, 20) = False Then
    Cancel = True
    Me.Text4.SetFocus
    Text4_GotFocus
End If
If Cancel = False Then CloseIme
End Sub
Private Function EditTool(Index As Integer)
Select Case Index
Case 1 '新增
    If TBar1.Buttons(1).Enabled = True Then
        EditSelect = 1
        For intI = 1 To 4
            TBar1.Buttons(intI).Enabled = False
        Next intI
        For intI = 6 To 9
            TBar1.Buttons(intI).Enabled = False
        Next intI
        TBar1.Buttons(11).Enabled = True
        TBar1.Buttons(12).Enabled = True
        TBar1.Buttons(14).Enabled = False
        Text2.Locked = False
        Text3.Locked = False
        Text4.Locked = False
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        'Added by Lydia 2023/11/14
        textSPT04.Locked = False
        textSPT04.Text = ""
        'end 2023/11/14
        Text2.SetFocus
    End If
Case 2 '修改
    If TBar1.Buttons(2).Enabled = True Then
        EditSelect = 2
        For intI = 1 To 4
            TBar1.Buttons(intI).Enabled = False
        Next intI
        For intI = 6 To 9
            TBar1.Buttons(intI).Enabled = False
        Next intI
        TBar1.Buttons(11).Enabled = True
        TBar1.Buttons(12).Enabled = True
        TBar1.Buttons(14).Enabled = False
        Text2.Locked = True
        Text3.Locked = True
        Text4.Locked = False
        textSPT04.Locked = False 'Added by Lydia 2023/11/14
        Text4.SetFocus
    End If
Case 3 '刪除
    If TBar1.Buttons(3).Enabled = True Then
        EditSelect = 3
        For intI = 1 To 4
            TBar1.Buttons(intI).Enabled = False
        Next intI
        For intI = 6 To 9
            TBar1.Buttons(intI).Enabled = False
        Next intI
        TBar1.Buttons(11).Enabled = True
        TBar1.Buttons(12).Enabled = True
        TBar1.Buttons(14).Enabled = False
        If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            Computer.MoveNext
            If Computer.EOF Then
                Computer.MovePrevious
                Computer.MovePrevious
            End If
            strExc(1) = Computer.Fields(0).Value
            cnnConnection.Execute "Delete SpecialPatentTrademark Where SPT01='" & Me.Text2.Text & "' And SPT02='" & Me.Text3.Text & "' "
            If Computer.State = adStateOpen Then Computer.Close
            'Modified by Lydia 2023/11/14 +SPT04
            strExc(0) = "Select SPT01, SPT02, SPT03, SPT04 From SpecialPatentTrademark Order By 1, 2 "
            Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
            Computer.Find "SPT01='" & strExc(1) & "'", 0, adSearchForward, 1
            If IsNull(Computer.Fields(0)) Then
                Text2.Text = ""
            Else
                Text2.Text = Computer.Fields(0).Value
            End If
            If IsNull(Computer.Fields(1).Value) Then
                Text3.Text = ""
            Else
                Text3.Text = Computer.Fields(1).Value
            End If
            If IsNull(Computer.Fields(2).Value) Then
                Text4.Text = ""
            Else
                Text4.Text = Computer.Fields(2).Value
            End If
            textSPT04.Text = "" & Computer.Fields(3).Value 'Added by Lydia 2023/11/14
            
            For intI = 1 To 4
                TBar1.Buttons(intI).Enabled = True
            Next intI
            
            For intI = 6 To 9
                TBar1.Buttons(intI).Enabled = True
            Next intI
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
            EditSelect = 0
        Else
            For intI = 1 To 4
                TBar1.Buttons(intI).Enabled = True
            Next intI
            For intI = 6 To 9
                TBar1.Buttons(intI).Enabled = True
            Next intI
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
            EditSelect = 0
            Exit Function
        End If
        Text2.SetFocus
    End If
Case 4 '查詢
    If TBar1.Buttons(4).Enabled = True Then
        EditSelect = 4
        For intI = 1 To 4
            TBar1.Buttons(intI).Enabled = False
        Next intI
        For intI = 6 To 9
            TBar1.Buttons(intI).Enabled = False
        Next intI
        TBar1.Buttons(11).Enabled = True
        TBar1.Buttons(12).Enabled = True
        TBar1.Buttons(14).Enabled = False
        Text2.Locked = False
        Text3.Locked = True
        Text4.Locked = True
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        'Added by Lydia 2023/11/14
        textSPT04.Locked = True
        textSPT04.Text = ""
        'end 2023/11/14
        Text2.SetFocus
    End If
Case 5 '第一筆
    If TBar1.Buttons(6).Enabled = True Then
        Computer.MoveFirst
        If IsNull(Computer.Fields(0)) Then
            Text2.Text = ""
        Else
            Text2.Text = Computer.Fields(0).Value
        End If
        If IsNull(Computer.Fields(1).Value) Then
            Text3.Text = ""
        Else
            Text3.Text = Computer.Fields(1).Value
        End If
        If IsNull(Computer.Fields(2).Value) Then
            Text4.Text = ""
        Else
            Text4.Text = Computer.Fields(2).Value
        End If
        textSPT04.Text = "" & Computer.Fields(3).Value  'Added by Lydia 2023/11/14
    End If
Case 6 '前一筆
    If TBar1.Buttons(7).Enabled = True Then
        Computer.MovePrevious
        If Computer.BOF Then
            DataErrorMessage (6)
            Computer.MoveFirst
        End If
        If IsNull(Computer.Fields(0)) Then
            Text2.Text = ""
        Else
            Text2.Text = Computer.Fields(0).Value
        End If
        If IsNull(Computer.Fields(1).Value) Then
            Text3.Text = ""
        Else
            Text3.Text = Computer.Fields(1).Value
        End If
        If IsNull(Computer.Fields(2).Value) Then
            Text4.Text = ""
        Else
            Text4.Text = Computer.Fields(2).Value
        End If
        textSPT04.Text = "" & Computer.Fields(3).Value  'Added by Lydia 2023/11/14
    End If
Case 7 '後一筆
    If TBar1.Buttons(8).Enabled = True Then
        Computer.MoveNext
        If Computer.EOF Then
            DataErrorMessage (7)
            Computer.MoveLast
        End If
        If IsNull(Computer.Fields(0)) Then
            Text2.Text = ""
        Else
            Text2.Text = Computer.Fields(0).Value
        End If
        If IsNull(Computer.Fields(1).Value) Then
            Text3.Text = ""
        Else
            Text3.Text = Computer.Fields(1).Value
        End If
        If IsNull(Computer.Fields(2).Value) Then
            Text4.Text = ""
        Else
            Text4.Text = Computer.Fields(2).Value
        End If
        textSPT04.Text = "" & Computer.Fields(3).Value  'Added by Lydia 2023/11/14
    End If
Case 8 '最後一筆
    If TBar1.Buttons(9).Enabled = True Then
        Computer.MoveLast
        If IsNull(Computer.Fields(0)) Then
            Text2.Text = ""
        Else
            Text2.Text = Computer.Fields(0).Value
        End If
        If IsNull(Computer.Fields(1).Value) Then
            Text3.Text = ""
        Else
            Text3.Text = Computer.Fields(1).Value
        End If
        If IsNull(Computer.Fields(2).Value) Then
            Text4.Text = ""
        Else
            Text4.Text = Computer.Fields(2).Value
        End If
        textSPT04.Text = "" & Computer.Fields(3).Value  'Added by Lydia 2023/11/14
    End If
Case 9 '確定
    If TBar1.Buttons(11).Enabled = True Then
        If Text2.Text = "" Then MsgBox "專利或商標不可為空值", vbInformation: Text2.SetFocus: Exit Function
        If EditSelect = 1 Or EditSelect = 2 Then
            If Text3.Text = "" Then MsgBox "特殊專利商標不可為空值", vbInformation: Text3.SetFocus: Exit Function
            'Modified by Lydia 2023/11/14 「說明」改為「特殊專利商標代碼名稱」
            If Text4.Text = "" Then MsgBox "特殊專利商標代碼名稱不可為空值", vbInformation: Text4.SetFocus: Exit Function
        End If
        If EditSelect = 1 Then
            If cp.State = adStateOpen Then cp.Close
            cp.CursorLocation = adUseClient
            strExc(1) = "SELECT COUNT(SPT01) FROM SpecialPatentTrademark WHERE SPT01='" & Text2.Text & "' And SPT02='" & Me.Text3.Text & "' "
            cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
            If Val("" & cp.Fields(0).Value) <> 0 Then
                MsgBox "相同特殊專利商標資料已存在!!!"
                Text2.SetFocus
                cp.Close
                Exit Function
            End If
            cp.Close
            If Not (Text2.Text = "" Or (Text2.Text >= "1" And Text2.Text <= "2")) Then
                MsgBox "輸入錯誤", vbInformation
                Text2.SetFocus
                Exit Function
            End If
            If Not (Text3.Text = "" Or (Text3.Text >= "1" And Text3.Text <= "9")) Then
                MsgBox "輸入錯誤", vbInformation
                Text3.SetFocus
                Exit Function
            End If
        End If
        If EditSelect = 1 Or EditSelect = 2 Then
            If Me.Text4.Text = "" Then
                MsgBox "輸入錯誤", vbInformation
                Text4.SetFocus
                Exit Function
            End If
        End If
        Select Case EditSelect
        Case 1
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            'Modified by Lydia 2023/11/14 +ChgSQL(),textSPT04.Text
            cnnConnection.Execute "Insert Into SpecialPatentTrademark Values('" & Text2.Text & "','" & Text3.Text & "','" & ChgSQL(Text4.Text) & "','" & ChgSQL(textSPT04.Text) & "')"
            EditSelect = 0
        Case 2
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            'Modified by Lydia 2023/11/14 +ChgSQL(),textSPT04.Text
            cnnConnection.Execute "Update SpecialPatentTrademark Set SPT03='" & ChgSQL(Text4.Text) & "', SPT04='" & ChgSQL(textSPT04.Text) & "' Where SPT01='" & Text2.Text & "' And SPT02='" & Me.Text3.Text & "' "
            EditSelect = 0
        Case 4
            Computer.Find "SPT01='" & Text2 & "'", 0, adSearchForward, 1
            If Me.Text3.Text <> "" Then Computer.Find "SPT02='" & Text3 & "'", 0, adSearchForward, 1
            If Computer.EOF Then
                MsgBox "找不到搜尋之資料", vbInformation
                For intI = 1 To 4
                    TBar1.Buttons(intI).Enabled = True
                Next intI
                For intI = 6 To 9
                    TBar1.Buttons(intI).Enabled = True
                Next intI
                TBar1.Buttons(11).Enabled = False
                TBar1.Buttons(12).Enabled = False
                TBar1.Buttons(14).Enabled = True
                EditSelect = 0
                If Computer.State = adStateOpen Then Computer.Close
                'Modified by Lydia 2023/11/14 +SPT04
                strExc(0) = "Select SPT01, SPT02, SPT03, SPT04 From SpecialPatentTrademark Order By 1, 2 "
                Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
                Computer.MoveFirst
                Text2.Text = Computer.Fields(0).Value
                Text3.Text = Computer.Fields(1).Value
                Text4.Text = Computer.Fields(2).Value
                textSPT04.Text = "" & Computer.Fields(3).Value 'Added by Lydia 2023/11/14
                Exit Function
            Else
                If IsNull(Computer.Fields(0)) Then
                    Text2.Text = ""
                Else
                    Text2.Text = Computer.Fields(0).Value
                End If
                If IsNull(Computer.Fields(1).Value) Then
                    Text3.Text = ""
                Else
                    Text3.Text = Computer.Fields(1).Value
                End If
                If IsNull(Computer.Fields(2).Value) Then
                    Text4.Text = ""
                Else
                    Text4.Text = Computer.Fields(2).Value
                End If
                textSPT04.Text = Computer.Fields(3).Value 'Added by Lydia 2023/11/14
            End If
        End Select
        If Computer.State = adStateOpen Then Computer.Close
        'Modified by Lydia 2023/11/14 +SPT04
        strExc(0) = "Select SPT01, SPT02, SPT03,SPT04 From SpecialPatentTrademark Order By 1, 2 "
        Computer.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
        Computer.Find "SPT01='" & Text2.Text & "'", 0, adSearchForward, 1
        If Me.Text3.Text <> "" Then Computer.Find "SPT02='" & Text3.Text & "'", 0, adSearchForward, 1
        If IsNull(Computer.Fields(0)) Then
            Text2.Text = ""
        Else
            Text2.Text = Computer.Fields(0).Value
        End If
        If IsNull(Computer.Fields(1).Value) Then
            Text3.Text = ""
        Else
            Text3.Text = Computer.Fields(1).Value
        End If
        If IsNull(Computer.Fields(2).Value) Then
            Text4.Text = ""
        Else
            Text4.Text = Computer.Fields(2).Value
        End If
        textSPT04.Text = Computer.Fields(3).Value 'Added by Lydia 2023/11/14
        For intI = 1 To 4
            TBar1.Buttons(intI).Enabled = True
        Next intI
        For intI = 6 To 9
            TBar1.Buttons(intI).Enabled = True
        Next intI
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
                Text2.Text = Computer.Fields(0).Value
                Text3.Text = Computer.Fields(1).Value
                Text4.Text = Computer.Fields(2).Value
                textSPT04.Text = "" & Computer.Fields(3).Value  'Added by Lydia 2023/11/14
            End If
            For intI = 1 To 4
                TBar1.Buttons(intI).Enabled = True
            Next intI
            For intI = 6 To 9
                TBar1.Buttons(intI).Enabled = True
            Next intI
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

'Added by Lydia 2023/11/14
If Me.textSPT04.Enabled = True Then
   Cancel = False
   textSPT04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2023/11/14

TxtValidate = True
End Function

'Added by Lydia 2023/11/14
Private Sub textSPT04_GotFocus()
   TextInverse textSPT04
End Sub
'Added by Lydia 2023/11/14
Private Sub textSPT04_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Added by Lydia 2023/11/14
Private Sub textSPT04_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Me.textSPT04.Text, 20) = False Then
      Cancel = True
      Me.textSPT04.SetFocus
      textSPT04_GotFocus
   End If
End Sub
