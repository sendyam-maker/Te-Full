VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075010 
   BorderStyle     =   1  '單線固定
   Caption         =   "機關單位資料維護"
   ClientHeight    =   5172
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9336
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   9336
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1008
      Top             =   4872
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
            Picture         =   "frm075010.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075010.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9336
      _ExtentX        =   16468
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   90
      TabIndex        =   6
      Top             =   720
      Width           =   9210
      _ExtentX        =   16235
      _ExtentY        =   7684
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "單筆"
      TabPicture(0)   =   "frm075010.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAddr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOrgName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtOrgNum"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtFax"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtZipCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm075010.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid1"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3945
         Left            =   -74820
         TabIndex        =   7
         Top             =   300
         Width           =   8850
         _ExtentX        =   15600
         _ExtentY        =   6964
         _Version        =   393216
         BackColor       =   14942187
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   16772048
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtZipCode 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   264
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2424
         Width           =   1215
      End
      Begin VB.TextBox txtTel 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   264
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1488
         Width           =   2175
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   264
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1944
         Width           =   2172
      End
      Begin VB.TextBox txtOrgNum 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   264
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   564
         Width           =   1212
      End
      Begin MSForms.TextBox txtOrgName 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   996
         Width           =   6570
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "11589;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAddr 
         Height          =   585
         Left            =   1320
         TabIndex        =   8
         Top             =   2910
         Width           =   6570
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "11589;1032"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "機關名稱："
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   1020
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "機關編號："
         Height          =   252
         Left            =   252
         TabIndex        =   13
         Top             =   570
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "電　　話："
         Height          =   252
         Left            =   228
         TabIndex        =   12
         Top             =   1488
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2424
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "傳　　真："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1944
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "地　　址："
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2904
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm075010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、txtAddr、txtOrgName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim Rs As New ADODB.Recordset, blnIsSave As Boolean, blnIsSearch As Boolean, blnIsNew As Boolean, blnisEdit As Boolean
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean, intSaveKind As Integer
Dim m_OR01 As String
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_stat As Integer
Dim m_EDIT As Integer

Private Sub Form_Load()
'*****************
   m_bInsert = IsUserHasRightOfFunction("frm075010", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075010", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075010", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075010", strFind, False)
'****************
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/17
   m_EDIT = 0
   MoveFormToCenter Me
   blnIsSave = False
   TxtCanTUse
   CmdEnabled
   GetAllData
   
'**********************
   If m_bInsert Then
       tlbar.Buttons(1).Enabled = True
   Else
       tlbar.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       tlbar.Buttons(2).Enabled = True
   Else
       tlbar.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       tlbar.Buttons(3).Enabled = True
   Else
       tlbar.Buttons(3).Enabled = False
   End If
'************************
'Add By Cheng 2002/01/30
If Rs.EOF And Rs.BOF Then
   SetButtonNotEnabled
End If
End Sub
' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      
      Case vbKeyF2:        ' 新增
         If m_bInsert Then
           New_Click
           m_EDIT = 1
           KeyCode = 0
         End If
      Case vbKeyF3:        '修改
          If m_bUpdate Then
           Edit_Click
           m_EDIT = 2
           KeyCode = 0
          End If
      Case vbKeyF4:        '查詢
         m_EDIT = 4
         Rs.MoveFirst
         CmdUnabled
         blnIsSearch = True
         Cleartxt
         txtOrgNum.Locked = False
         txtOrgNum.SetFocus
         intSaveKind = 3
         KeyCode = 0
      Case vbKeyF5:        '刪除
       If m_bDelete Then
           Delete_Click
           m_EDIT = 0
           KeyCode = 0
       End If
      Case vbKeyHome:      '第一筆
        If m_EDIT = 0 Then
           Rs.MoveFirst
           PutDataInObject
           KeyCode = 0
        End If
      Case vbKeyPageUp:    '上一筆
        If m_EDIT = 0 Then
           Rs.MovePrevious
           If Rs.BOF Then
              Rs.MoveFirst
              PutDataInObject
              DataErrorMessage (6)
           End If
           PutDataInObject
           KeyCode = 0
         End If
      Case vbKeyPageDown:  '下一筆
      If m_EDIT = 0 Then
           Rs.MoveNext
           If Rs.EOF Then
              Rs.MoveLast
              PutDataInObject
              DataErrorMessage (7)
           End If
           PutDataInObject
           KeyCode = 0
        End If
      Case vbKeyEnd:       '最後一筆
      If m_EDIT = 0 Then
           Rs.MoveLast
           PutDataInObject
           KeyCode = 0
        End If
      Case vbKeyReturn:
          If m_EDIT <> 0 Then
             cmdok_Click
              If m_stat <> 1 Then
                 m_EDIT = 0
                 m_stat = 0
              End If
          End If
      Case vbKeyF9:        '確定
          If m_EDIT <> 0 Then
             cmdok_Click
             If m_stat <> 1 Then
                m_EDIT = 0
                m_stat = 0
             End If
          End If
      Case vbKeyF10:       '取消
          If m_EDIT <> 0 Or m_stat = 1 Then
           
            If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
             m_EDIT = 0
            m_stat = 0
            CmdEnabled
            If blnIsSearch Then
               blnIsSearch = False
               tlbar.Buttons(11).Enabled = False
            End If
            If blnIsNew Then blnIsNew = False
            If blnisEdit Then blnisEdit = False
            TxtCanTUse
            PutDataInObject
            m_EDIT = 0
          End If
      Case vbKeyEscape:    '離開
            Rs.Close
            m_EDIT = 0
            Unload Me
   End Select
    
   If m_stat = 1 Then
      Exit Sub
    End If
   
      ' Ken 90.07.16 -- Start
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Or _
      KeyCode = vbKeyF10 Then
      
      If m_bInsert Then
          tlbar.Buttons(1).Enabled = True
      Else
          tlbar.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          tlbar.Buttons(2).Enabled = True
      Else
          tlbar.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          tlbar.Buttons(3).Enabled = True
      Else
          tlbar.Buttons(3).Enabled = False
      End If
   End If
   ' Ken 90.07.16 -- End

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm075010 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
        .col = 0
        m_OR01 = .Text
        Rs.Find "or01='" + m_OR01 + "'", 0, adSearchForward
        PutDataInObject
   End With
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim yn As Boolean

   Select Case Button.Index
      Case 1
         New_Click
         m_EDIT = 1
      Case 2
         m_EDIT = 2
         Edit_Click
      Case 3
         m_EDIT = 0
         Delete_Click
      Case 4
         m_EDIT = 4
         'Add By Cheng 2002/01/30
         If Rs.EOF And Rs.BOF Then Exit Sub
         
         Rs.MoveFirst
         CmdUnabled
         blnIsSearch = True
         Cleartxt
         txtOrgNum.Locked = False
         txtOrgNum.SetFocus
         intSaveKind = 3
      Case 6
       If m_EDIT = 0 Then
         Rs.MoveFirst
         PutDataInObject
       End If
      Case 7
       If m_EDIT = 0 Then
         Rs.MovePrevious
         If Rs.BOF Then
            Rs.MoveFirst
            PutDataInObject
            DataErrorMessage (6)
         End If
         PutDataInObject
       End If
      Case 8
       If m_EDIT = 0 Then
         Rs.MoveNext
         If Rs.EOF Then
            Rs.MoveLast
            PutDataInObject
            DataErrorMessage (7)
         End If
         PutDataInObject
       End If
      Case 9
       If m_EDIT = 0 Then
         Rs.MoveLast
         PutDataInObject
       End If
      Case 11
          cmdok_Click
          If m_stat <> 1 Then
             m_EDIT = 0
             m_stat = 0
          End If
      Case 12
        
         yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
         If yn = vbNo Then
            Exit Sub
         End If
          m_EDIT = 0
          m_stat = 0
         CmdEnabled
         If blnIsSearch Then
            blnIsSearch = False
            tlbar.Buttons(11).Enabled = False
         End If
         If blnIsNew Then blnIsNew = False
         If blnisEdit Then blnisEdit = False
         
         TxtCanTUse
         PutDataInObject
         m_EDIT = 0
      Case 14
          m_EDIT = 0
          Rs.Close
          Unload Me
   End Select
   If m_stat = 1 Then
      Exit Sub
   End If
   '*********************
   If Button.Index <> 14 And Button.Index <> 1 And _
      Button.Index <> 2 And Button.Index <> 3 And _
      Button.Index <> 4 Then

      If m_bInsert Then
          tlbar.Buttons(1).Enabled = True
      Else
          tlbar.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          tlbar.Buttons(2).Enabled = True
      Else
          tlbar.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          tlbar.Buttons(3).Enabled = True
      Else
          tlbar.Buttons(3).Enabled = False
      End If
   End If
'****************************
'Add By Cheng 2002/01/30
If Rs.State = adStateOpen Then
   If Rs.EOF And Rs.BOF Then
      SetButtonNotEnabled
   End If
End If
End Sub

Private Sub cmdok_Click()
 Dim yn As Integer, i As Integer
   If intSaveKind <> 3 Then
      If AllTextBeforeSaveCheck Then
         m_stat = 1
         Exit Sub
      End If
         If intSaveKind = 1 Then
            blnIsNew = False
         ElseIf intSaveKind = 2 Then
            blnisEdit = False
         End If
      'Add By Cheng 2002/05/24
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      If Not SaveData(intSaveKind) Then
         DataErrorMessage (3)
         m_stat = 1
         Exit Sub
      End If
   Else
      cmdSearch_Click
   End If
   CmdEnabled
   TxtCanTUse
   m_stat = 0
End Sub

Private Sub cmdSearch_Click()
   blnIsSearch = False
   Rs.Find "or01='" + txtOrgNum + "'", 0, adSearchForward
   If Not Rs.EOF Then
      PutDataInObject
   Else
      MsgBox "無此資料", vbOKOnly, "訊息"
      Rs.MoveFirst
      PutDataInObject
   End If
   txtOrgNum.Locked = True
End Sub

Private Sub CmdEnabled()
 Dim i As Integer
    tlbar.Buttons(11).Enabled = False
    tlbar.Buttons(12).Enabled = False
    For i = 1 To 9
        tlbar.Buttons(i).Enabled = True
    Next
        tlbar.Buttons(14).Enabled = True
End Sub

Private Sub CmdUnabled()
 Dim i As Integer
    tlbar.Buttons(11).Enabled = True
    tlbar.Buttons(12).Enabled = True
    For i = 1 To 9
        tlbar.Buttons(i).Enabled = False
    Next
        tlbar.Buttons(14).Enabled = False
End Sub

Private Sub TxtCanTUse()
   txtOrgNum.Locked = True
   txtOrgName.Locked = True
   txtTel.Locked = True
   txtFax.Locked = True
   txtZipCode.Locked = True
   txtAddr.Locked = True
End Sub

Private Sub TxtCanUse()
txtOrgNum.Locked = False
txtOrgName.Locked = False
txtTel.Locked = False
txtFax.Locked = False
txtZipCode.Locked = False
txtAddr.Locked = False

End Sub

Private Function GetAllData() As Boolean
   strExc(1) = "select or01,or02,or03,or04,or05,or06 from organization order by or01"
   intI = 0
   Set Rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      GridHead
      GridData
      Rs.MoveFirst
      PutDataInObject
      GetAllData = True
   End If
   blnIsSave = False
End Function
Private Function GetData() As Boolean
   strExc(1) = "select or01,or02,or03,or04,or05,or06 from organization order by or01"
   intI = 0
   Set Rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If Not Rs.EOF Then
      GetData = True
   Else
      GetData = False
   End If
   blnIsSave = False
End Function

Private Sub PutDataInObject()
   'Modify By Cheng 2002/01/30
   '避色無資料時出現錯誤訊息
   If Not Rs.EOF Then
      txtOrgNum = IIf(IsNull(Rs.Fields!or01), "", Rs.Fields!or01)
      txtOrgName = IIf(IsNull(Rs.Fields!or02), "", Rs.Fields!or02)
      txtTel = IIf(IsNull(Rs.Fields!or03), "", Rs.Fields!or03)
      txtFax = IIf(IsNull(Rs.Fields!or04), "", Rs.Fields!or04)
      txtZipCode = IIf(IsNull(Rs.Fields!or05), "", Rs.Fields!or05)
      txtAddr = IIf(IsNull(Rs.Fields!or06), "", Rs.Fields!or06)
   Else
      txtOrgNum = ""
      txtOrgName = ""
      txtTel = ""
      txtFax = ""
      txtZipCode = ""
      txtAddr = ""
   End If
End Sub

Private Sub New_Click()
   CmdUnabled
   TxtCanUse
   Cleartxt
   txtOrgNum.SetFocus
   blnIsNew = True
   intSaveKind = 1
End Sub

Private Sub Edit_Click()
   TxtCanUse
   CmdUnabled
   txtOrgNum.Locked = True
   txtOrgName.SetFocus
   blnisEdit = True
   intSaveKind = 2
End Sub

Private Sub Delete_Click()
 Dim Del As Boolean
 Dim nBookMark As Variant
 
 nBookMark = Rs.Bookmark
   CmdUnabled
   If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
      strExc(1) = "delete from organization where or01=" & CNULL(txtOrgNum)
      'edit by nickc 2007/02/07 不用 dll 了
      'Del = objLawDll.ExecSQL(1, strExc)
      Del = ClsLawExecSQL(1, strExc)
      If Del Then
        ' MsgBox "'" + txtOrgNum + "'  刪除成功"
         If GetData Then
            If nBookMark > Rs.RecordCount Then
               Rs.MoveFirst
            Else
               Rs.Bookmark = nBookMark
            End If
            PutDataInObject
            GridData
         Else
            MSHFlexGrid1.Clear
            MSHFlexGrid1.Rows = 2
            Cleartxt
         End If
      Else
         MsgBox "'" + txtOrgNum + "'  刪除不成功"
      End If
   End If
   CmdEnabled
   TxtCanTUse
End Sub

Private Function SaveData(ByVal i As Integer) As Boolean
   Dim nBookMark As Variant
   Dim strSql As String
   Err.Clear
   On Error Resume Next
   Select Case i
      Case 1
        
         strExc(1) = "insert into organization(or01,or02,or03,or04,or05,or06) " & _
            "values (" & CNULL(txtOrgNum) & "," & CNULL(txtOrgName) & "," & _
            CNULL(txtTel) & "," & CNULL(txtFax) & "," & CNULL(txtZipCode) & _
            "," + CNULL(txtAddr) & ")"
         'edit by nickc 2007/02/07 不用 dll 了
         'SaveData = objLawDll.ExecSQL(1, strExc)
         SaveData = ClsLawExecSQL(1, strExc)
         blnIsNew = False
         If SaveData Then
           ' GetAllData
            Rs.ReQuery
            Rs.Find "or01='" + txtOrgNum.Text + "'", 0, adSearchForward
            PutDataInObject
            GridData
         End If
      Case 2
         nBookMark = Rs.Bookmark
'         strSQL = "update organization set or02=" + CNULL(txtOrgName) + ",or03=" + CNULL(txtTel) + ",or04=" + CNULL(txtFax) + _
'            ",or05=" + CNULL(txtZipCode) + ",or06=" + CNULL(txtAddr) + " where or01=" + CNULL(txtOrgNum) + ""
         strSql = "UPDATE ORGANIZATION SET OR02 ='" & txtOrgName.Text & "'," & _
                  "OR03 ='" & txtTel.Text & "'," & _
                  "OR04 ='" & txtFax.Text & "'," & _
                  "OR05 ='" & txtZipCode.Text & "'," & _
                  "OR06 ='" & txtAddr.Text & "' WHERE OR01='" & txtOrgNum & "'"
         'SaveData = objLawDll.ExecSQL(1, strExc)
         cnnConnection.Execute strSql
         blnisEdit = False
         If Err.Number = 0 Then
            SaveData = True
            'GetAllData
            Rs.ReQuery
            Rs.Bookmark = nBookMark
            PutDataInObject
            GridData
         Else
            SaveData = False
         End If
         
      Case 3
         cmdSearch_Click
   End Select
End Function

Private Sub Cleartxt()
   txtOrgNum = ""
   txtOrgName = ""
   txtTel = ""
   txtFax = ""
   txtZipCode = ""
   txtAddr = ""
End Sub

Private Sub txtAddr_GotFocus()
   TextInverse txtAddr
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAddr.IMEMode = 1
   OpenIme
End Sub

'Added by Lydia 2021/09/17
Private Sub txtAddr_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then Forms(0).PopupMenu2 txtAddr  'Form 2.0的TextBox增加右鍵選單功能; 經過測試MouseMove無效,要放在MouseDown
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtAddr, 60) = False Then
      Cancel = True
      txtAddr_GotFocus
   End If

End Sub

Private Sub txtFax_GotFocus()
   TextInverse txtFax
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtFax.IMEMode = 2
   CloseIme
End Sub

Private Sub txtOrgName_GotFocus()
   TextInverse txtOrgName
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtOrgName.IMEMode = 1
   OpenIme
End Sub

Private Sub txtOrgName_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtOrgName, 40) = False Then
      Cancel = True
      txtOrgName_GotFocus
   End If

End Sub

Private Sub txtOrgNum_GotFocus()
   TextInverse txtOrgNum
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtOrgNum.IMEMode = 2
   CloseIme
End Sub

Private Sub txtOrgNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And blnIsSearch Then
      cmdok_Click
   Else
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtOrgNum_Validate(Cancel As Boolean)
Dim strTempName As String
If txtOrgNum <> "" Then
  If blnIsNew Then
      'Added by Lydia 2025/11/18
      If Len(txtOrgNum) <> 3 Then
         MsgBox "目前使用3碼編號，若要變更長度，請詢問電腦中心！", vbInformation + vbOKOnly
         Cancel = True
         Exit Sub
      End If
      'end 2025/11/18
      If Not Rs.EOF And Not Rs.BOF Then
         Rs.MoveFirst
         Rs.Find "or01='" + txtOrgNum + "'", 0, adSearchForward
         If Not Rs.EOF Then MsgBox "此機關編號已存在", vbCritical: Cancel = True
      End If
  End If
  If blnIsSearch Then tlbar.Buttons(11).Enabled = True
End If
If Cancel Then txtOrgNum_GotFocus
End Sub
Private Sub ChktxtOrgNum()
If txtOrgNum = "" Then
   MsgBox "機關編號不可空白", vbOKOnly, "錯誤"
   Exit Sub
End If
End Sub
Private Sub GridHead()
With MSHFlexGrid1
.row = 0
.col = 0
.ColWidth(0) = 1000
.Text = "機關編號"
.col = 1
.ColWidth(1) = 2000
.Text = "機關名稱"
.col = 2
.ColWidth(2) = 1500
.Text = "電話"
.col = 3
.ColWidth(3) = 1500
.Text = "傳真"
.col = 4
.ColWidth(4) = 900
.Text = "郵遞區號"
.col = 5
.ColWidth(5) = 2500
.Text = "地址"
End With
End Sub

Private Sub GridData()
   Set MSHFlexGrid1.Recordset = Rs
   GridHead
End Sub

Private Sub txtTel_GotFocus()
   TextInverse txtTel
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtTel.IMEMode = 2
   CloseIme
End Sub

Private Sub txtZipCode_GotFocus()
   TextInverse txtZipCode
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtZipCode.IMEMode = 2
   CloseIme
End Sub

Private Sub txtZipCode_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Function AllTextBeforeSaveCheck() As Boolean
If txtOrgNum = "" Then
   MsgBox "機關編號不可空白", vbCritical
   AllTextBeforeSaveCheck = True
   txtOrgNum.SetFocus
   Exit Function
End If
If txtOrgName = "" Then
   MsgBox "機關名稱不可空白", vbCritical
   AllTextBeforeSaveCheck = True
   txtOrgName.SetFocus
   Exit Function
End If
If txtZipCode = "" Then
   MsgBox "郵遞區號不可空白", vbCritical
   AllTextBeforeSaveCheck = True
   txtZipCode.SetFocus
   Exit Function
End If
If txtAddr = "" Then
   MsgBox "地址不可空白", vbCritical
   AllTextBeforeSaveCheck = True
   txtAddr.SetFocus
   Exit Function
End If

AllTextBeforeSaveCheck = False
End Function

Private Sub SetButtonNotEnabled()
   tlbar.Buttons(2).Enabled = False
   tlbar.Buttons(3).Enabled = False
   tlbar.Buttons(4).Enabled = False
   tlbar.Buttons(6).Enabled = False
   tlbar.Buttons(7).Enabled = False
   tlbar.Buttons(8).Enabled = False
   tlbar.Buttons(9).Enabled = False
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.txtAddr.Enabled = True Then
   Cancel = False
   txtAddr_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtOrgName.Enabled = True Then
   Cancel = False
   txtOrgName_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtOrgNum.Enabled = True Then
   Cancel = False
   txtOrgNum_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/09/17 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

