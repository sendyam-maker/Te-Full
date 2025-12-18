VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090214 
   BorderStyle     =   1  '單線固定
   Caption         =   "期刊索引資料維護"
   ClientHeight    =   5730
   ClientLeft      =   1815
   ClientTop       =   1920
   ClientWidth     =   9315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin TabDlg.SSTab SSTab1 
      Height          =   5100
      Left            =   36
      TabIndex        =   3
      Top             =   576
      Width           =   9252
      _ExtentX        =   16325
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm090214.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm090214.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090214.frx":0038
         Height          =   4752
         Left            =   -74964
         TabIndex        =   6
         Top             =   324
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   8387
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   1110
         TabIndex        =   0
         Top             =   495
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   1110
         TabIndex        =   1
         Top             =   975
         Width           =   5895
         VariousPropertyBits=   671107099
         MaxLength       =   52
         Size            =   "10398;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "索引代號："
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   528
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "說明："
         Height          =   180
         Left            =   195
         TabIndex        =   4
         Top             =   1008
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5865
      Top             =   795
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   840
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
            Picture         =   "frm090214.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090214.frx":1E25
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
      Width           =   9315
      _ExtentX        =   16431
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
End
Attribute VB_Name = "frm090214"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (GRD1,Text1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, p As New ADODB.Recordset
Dim i As Integer, EDITSELECT As Integer, s As Integer
Dim NEXTSTR As String, str As String, NOWSTR As String

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyF2
        If TBar1.Buttons(1).Enabled = True Then
            EDITTOOL (1)
            Text1(0).SetFocus
            SSTab1.Tab = 0
        End If
            KeyCode = 0
       Case vbKeyF3
         If TBar1.Buttons(2).Enabled = True Then
            EDITTOOL (2)
            SSTab1.Tab = 0
         End If
         KeyCode = 0
       Case vbKeyF5
         If TBar1.Buttons(3).Enabled = True Then
            EDITTOOL (3)
            SSTab1.Tab = 0
         End If
         KeyCode = 0
       Case vbKeyF4
         If TBar1.Buttons(4).Enabled = True Then
            EDITTOOL (4)
            Text1(0).SetFocus
            SSTab1.Tab = 0
         End If
         KeyCode = 0
       Case vbKeyF9, vbKeyReturn
        If TBar1.Buttons(11).Enabled = True Then
            EDITTOOL (9)
        End If
        KeyCode = 0
       Case vbKeyHome
            If TBar1.Buttons(6).Enabled = True Then
                EDITTOOL (5)
            End If
            KeyCode = 0
       Case vbKeyEnd
         If TBar1.Buttons(9).Enabled = True Then
            EDITTOOL (8)
         End If
         KeyCode = 0
       Case vbKeyPageUp
         If TBar1.Buttons(7).Enabled = True Then
            EDITTOOL (6)
         End If
         KeyCode = 0
       Case vbKeyPageDown
         If TBar1.Buttons(8).Enabled = True Then
            EDITTOOL (7)
         End If
         KeyCode = 0
       Case vbKeyF10
         If TBar1.Buttons(12).Enabled = True Then
            EDITTOOL (10)
         End If
         KeyCode = 0
       Case vbKeyEscape
         If TBar1.Buttons(14).Enabled = True Then
            EDITTOOL (11)
         End If
         KeyCode = 0
End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
      If EDITSELECT > 4 Then
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
   End If

End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm090214", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090214", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090214", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090214", strFind, False)
 MoveFormToCenter Me
For i = 0 To 1
    Text1(0).Locked = True
Next i
If pemain.State = adStateOpen Then pemain.Close
pemain.CursorLocation = adUseClient
p.CursorLocation = adUseClient

strExc(0) = "SELECT PI01,PI02 FROM PERIODICALINDEX ORDER BY PI01"
pemain.CursorLocation = adUseClient
pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If Not pemain.RecordCount = 0 Then
    Set Adodc1.Recordset = pemain
    SetGrd
    MouseClick (1)
    Adodc1.Recordset.ReQuery
    For i = 0 To 1
    If IsNull(pemain.Fields(i).Value) Then
        Text1(i).Text = ""
    Else
        Text1(i).Text = pemain.Fields(i).Value
    End If
    Next i
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
    locktext (1)
    SSTab1.Tab = 1
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090214 = Nothing
End Sub

Private Sub Grd1_Click()
GetData
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
       Case 1
          EDITTOOL (1)
          Text1(0).SetFocus
       Case 2
          EDITTOOL (2)
       Case 3
          EDITTOOL (3)
       Case 4
          EDITTOOL (4)
          Text1(0).SetFocus
       Case 6
          EDITTOOL (5)
       Case 7
          EDITTOOL (6)
       Case 8
          EDITTOOL (7)
       Case 9
          EDITTOOL (8)
       Case 11
          EDITTOOL (9)
       Case 12
          EDITTOOL (10)
       Case 14
          EDITTOOL (11)
End Select
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

End Sub

Private Function EDITTOOL(Index As Integer)
Select Case Index
       Case 1 'NEW
          EDITSELECT = 1
          locktext (2)
          For i = 0 To 1
            Text1(i).Text = ""
          Next i
          For i = 1 To 4
          TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
          TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          SSTab1.Tab = 0
       Case 2 'UPDATA
          EDITSELECT = 2
          locktext (3)
          Text1(0).Locked = True
          Text1(1).Locked = False
          Text1(1).SetFocus
          For i = 1 To 4
          TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
          TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          SSTab1.Tab = 0
       Case 3 'DELETE
          locktext (1)
          EDITSELECT = 3
          For i = 1 To 4
          TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
          TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          If MsgBox("是否要刪除此筆資料", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
          NOWSTR = pemain.Fields(0).Value
          pemain.MoveNext
          If pemain.EOF Then
            pemain.MoveFirst
            NEXTSTR = pemain.Fields(0).Value
          Else
            NEXTSTR = pemain.Fields(0).Value
          End If
            pemain.MovePrevious
            cnnConnection.Execute "delete periodicalindex where pi01='" & NOWSTR & "'"
            pemain.ReQuery
            Set Adodc1.Recordset = pemain
            Set GRD1.Recordset = Adodc1.Recordset
            pemain.Find "PI01='" & NEXTSTR & "'"
            SetGrd
            MouseClick (pemain.Bookmark)
            For i = 0 To 1
            Text1(i).Text = pemain.Fields(i).Value
            Next i
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
              SSTab1.Tab = 0
       Case 4 'QUTION
          EDITSELECT = 4
          locktext (4)
          For i = 0 To 1
          Text1(i).Text = ""
          Next i
          For i = 1 To 4
          TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
          TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          SSTab1.Tab = 0
       Case 5 'FIRST
          If TBar1.Buttons(6).Enabled = True Then
            If Not pemain.RecordCount = 0 Then
               pemain.MoveFirst
               For i = 0 To 1
                  If IsNull(pemain.Fields(i).Value) Then
                      Text1(i) = ""
                  Else
                      Text1(i).Text = pemain.Fields(i).Value
                  End If
               Next i
               'SetGrd
               MouseClick (pemain.Bookmark)
            End If
          End If
       Case 6 'PRIVATE
          If TBar1.Buttons(7).Enabled = True Then
            If Not pemain.RecordCount = 0 Then
                pemain.MovePrevious
                If pemain.BOF Then
                    DataErrorMessage (6)
                    pemain.MoveFirst
                End If
                For i = 0 To 1
                  If IsNull(pemain.Fields(i).Value) Then
                      Text1(i) = ""
                  Else
                      Text1(i).Text = pemain.Fields(i).Value
                  End If
                Next i
                'SetGrd
                MouseClick (pemain.Bookmark)
            End If
          End If
       Case 7 'NEXT
          If TBar1.Buttons(8).Enabled = True Then
            If Not pemain.RecordCount = 0 Then
              pemain.MoveNext
              If pemain.EOF Then
                  DataErrorMessage (7)
                  pemain.MoveLast
              End If
              For i = 0 To 1
                  If IsNull(pemain.Fields(i).Value) Then
                      Text1(i) = ""
                  Else
                      Text1(i).Text = pemain.Fields(i).Value
                  End If
               Next i
               'SetGrd
                MouseClick (pemain.Bookmark)
            End If
          End If
       Case 8 'LAST
          If TBar1.Buttons(9).Enabled = True Then
            If Not pemain.RecordCount = 0 Then
                pemain.MoveLast
                For i = 0 To 1
                  If IsNull(pemain.Fields(i).Value) Then
                      Text1(i) = ""
                  Else
                      Text1(i).Text = pemain.Fields(i).Value
                  End If
               Next i
               'SetGrd
                MouseClick (pemain.Bookmark)
            End If
          End If
       Case 9 'ENTER
          If EDITSELECT = 1 Then
          If p.State = adStateOpen Then p.Close
            If Len(Trim(Text1(0))) = 0 Then
               s = MsgBox("索引代號不可空白！！", , "User 輸入錯誤")
               Text1(0).SetFocus
               Exit Function
            End If
          strExc(1) = "select count(pi01) from periodicalindex where pi01='" & Text1(0) & "'"
          p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
          If p.Fields(0).Value <> "0" Then
          MsgBox "此資料已存在"
          Text1(0).SetFocus
          Exit Function
          End If
          End If
          Select Case EDITSELECT
                 Case 1
                     str = Text1(0).Text
                     cnnConnection.Execute "INSERT INTO PERIODICALINDEX (pi01,pi02) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "')"
                     pemain.ReQuery
                     Set Adodc1.Recordset = pemain
                     Set GRD1.Recordset = Adodc1.Recordset
                     pemain.Find "pi01='" & str & "'", 0, adSearchForward, 1
                     SetGrd
                     MouseClick (pemain.Bookmark)
                 Case 2
                     str = Text1(0).Text
                     cnnConnection.Execute "UPDATE PERIODICALINDEX SET PI02='" & Text1(1) & "' WHERE PI01='" & Text1(0) & "'"
                     pemain.ReQuery
                     Set Adodc1.Recordset = pemain
                     Set GRD1.Recordset = Adodc1.Recordset
                     pemain.Find "pi01='" & str & "'", 0, adSearchForward, 1
                     SetGrd
                     MouseClick (pemain.Bookmark)
                 Case 4
                     str = pemain.Fields(0).Value
                     pemain.Find "PI01= '" & Text1(0).Text & "'", 0, adSearchForward, 1
                     If pemain.EOF Then
                        MsgBox "查無資料"
                        EDITSELECT = 0
                        pemain.Find "pi01='" & str & "'", 0, adSearchForward, 1
                        For i = 1 To 4
                           TBar1.Buttons(i).Enabled = True
                        Next i
                        For i = 6 To 9
                           TBar1.Buttons(i).Enabled = True
                        Next i
                        TBar1.Buttons(11).Enabled = False
                        TBar1.Buttons(12).Enabled = False
                        TBar1.Buttons(14).Enabled = True
                        EDITSELECT = 0
                        Exit Function
                     End If
                     SetGrd
                     MouseClick (pemain.Bookmark)
                     For i = 0 To 1
                        If IsNull(pemain.Fields(i).Value) Then
                        Text1(i).Text = ""
                     Else
                        Text1(i).Text = pemain.Fields(i).Value
                     End If
                     Next i
'                     Set Adodc1.Recordset = pemain
          End Select
              For i = 1 To 4
                  TBar1.Buttons(i).Enabled = True
              Next i
              For i = 6 To 9
                  TBar1.Buttons(i).Enabled = True
              Next i
              TBar1.Buttons(11).Enabled = False
              TBar1.Buttons(12).Enabled = False
              TBar1.Buttons(14).Enabled = True
              locktext (1)
              EDITSELECT = 0
       Case 10 'CANCEL
             If MsgBox("你尚未存檔,確定離開?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
             If EDITSELECT = 1 Then pemain.MoveFirst
             EDITSELECT = 0
             For i = 0 To 1
             Text1(i).Text = pemain.Fields(i).Value
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
              locktext (1)
          End If
          EDITSELECT = 0
       Case 11 'END
           Unload Me
End Select
End Function

Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
Case 1
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1(Index).IMEMode = 1
   OpenIme
Case Else
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1(Index).IMEMode = 2
   CloseIme
End Select
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
           Case 0
            If EDITSELECT = 4 And KeyAscii = 13 Then
                EDITTOOL (9)
            End If
           Case 1
            
    End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
              If EDITSELECT = 1 Then
                If p.State = adStateOpen Then p.Close
                    strExc(1) = "select count(pi01) from periodicalindex where pi01='" & Text1(0) & "'"
                    p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                If p.Fields(0).Value <> "0" Then
                    MsgBox "此資料已存在"
                    Text1(0).SetFocus
                    Text1(0).SelStart = 0
                    Text1(0).SelLength = Len(Text1(0))
                    Cancel = True
                    Exit Sub
                Else
                    Cancel = False
                End If
              End If
        Case 1
         If Not CheckLengthIsOK(Text1(1), 26) Then
         Text1(1).SetFocus
         Text1(1).SelStart = 0
         Text1(1).SelLength = Len(Text1(1))
         Cancel = True
         Else
         Cancel = False
         End If
    End Select
End Sub
Private Sub locktext(Index As Integer) '鎖住輸入項
Dim j As Integer
Select Case Index
       Case 1 '初值
          For j = 0 To 1
             Text1(j).Locked = True
          Next j
       Case 2 '新增
          For j = 0 To 1
             Text1(j).Locked = False
          Next j
       Case 3 '修改
          Text1(0).Locked = True
          Text1(1).Locked = False
       Case 4 '查詢
            Text1(0).Locked = False
            Text1(1).Locked = True
End Select
End Sub
Sub GetData()
GRD1.row = GRD1.MouseRow
If GRD1.row <> 0 Then
   MouseClick (GRD1.row)
   GRD1.col = 0
   pemain.Find "PI01='" & GRD1.Text & "'", 0, adSearchForward, 1
   For i = 0 To 1
   If IsNull(pemain.Fields(i).Value) Then
       Text1(i).Text = ""
   Else
       Text1(i).Text = pemain.Fields(i).Value
   End If
   Next i
   'SSTab1.Tab = 0
End If
End Sub
Sub MouseClick(Strindex As Integer)
With GRD1
   For s = 1 To .Rows - 1
      .row = s
      .col = 0
      .CellBackColor = QBColor(15)
      .col = 1
      .CellBackColor = QBColor(15)
   Next s
   .row = Strindex
   .col = 0
   .CellBackColor = &HFFC0C0
   .col = 1
   .CellBackColor = &HFFC0C0
End With
End Sub

Sub SetGrd()
With GRD1
   
   .Cols = 2
   .row = 0
   .col = 0: .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
   .Text = "索引代號"
   .col = 1: .ColWidth(1) = 6000
    .CellAlignment = flexAlignCenterCenter
   .Text = "說明"
End With
End Sub

