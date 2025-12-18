VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050718 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人指定國外代理人維護功能"
   ClientHeight    =   6075
   ClientLeft      =   420
   ClientTop       =   4410
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9150
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   4
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   6
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   3
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   2
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3465
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3465
      Width           =   375
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   795
      Width           =   975
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "新增"
      Height          =   285
      Index           =   1
      Left            =   6570
      TabIndex        =   2
      Top             =   3020
      Width           =   735
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "刪除"
      Height          =   285
      Index           =   3
      Left            =   8100
      TabIndex        =   8
      Top             =   3020
      Width           =   735
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "加入"
      Height          =   285
      Index           =   2
      Left            =   7335
      TabIndex        =   7
      Top             =   3020
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8415
      Top             =   30
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
            Picture         =   "frm050718.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050718.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7425
      Top             =   2280
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm050718.frx":20F4
      Height          =   1245
      Left            =   270
      TabIndex        =   11
      Top             =   1680
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   2196
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CAA01N"
         Caption         =   "種類"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CAA03N"
         Caption         =   "申請國家"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CAA05"
         Caption         =   "順序"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CAA04"
         Caption         =   "CF代理人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "FANAME"
         Caption         =   "代理人名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   4844.977
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textCUID 
      Height          =   270
      Left            =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5760
      Width           =   8220
      VariousPropertyBits=   671105055
      Size            =   "14499;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12091;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   28
      Top             =   4920
      Width           =   6975
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   4
      Left            =   1320
      TabIndex        =   27
      Top             =   4560
      Width           =   6975
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   3
      Left            =   1320
      TabIndex        =   26
      Top             =   4200
      Width           =   6975
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   2
      Left            =   5640
      TabIndex        =   25
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   1
      Left            =   6120
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   300
      Index           =   0
      Left            =   5160
      TabIndex        =   23
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "順序："
      Height          =   300
      Index           =   11
      Left            =   360
      TabIndex        =   22
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "　　(日)："
      Height          =   300
      Index           =   10
      Left            =   360
      TabIndex        =   21
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "　　(英)："
      Height          =   300
      Index           =   9
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "名稱(中)："
      Height          =   300
      Index           =   8
      Left            =   360
      TabIndex        =   19
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "代理人國籍："
      Height          =   300
      Index           =   7
      Left            =   4440
      TabIndex        =   18
      Top             =   3840
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "CF代理人編號："
      Height          =   300
      Index           =   6
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   300
      Index           =   5
      Left            =   4440
      TabIndex        =   16
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "(1專利 2商標 3法務)"
      Height          =   300
      Index           =   4
      Left            =   2160
      TabIndex        =   15
      Top             =   3480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "案件種類："
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   14
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "名　　稱："
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "國籍："
      Height          =   300
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "申請人編號："
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   1185
   End
End
Attribute VB_Name = "frm050718"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (textCUID,Combo1,lblData)
'Create by Lydia 2016/10/24 申請人指定國外代理人維護功能
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim oText As TextBox
Dim oLbl As LABEL

Dim rsAssign As ADODB.Recordset
Dim rsAssignOld As ADODB.Recordset

Dim m_iCAAEditMode As Integer 'CF代理人狀態 1:新增 2:修改
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料
Dim stFormName As String
Dim m_FA69 As String '代理人狀態
Dim m_CU80 As String '申請人狀態
Dim m_SK01 As String '案件種類
Dim m_ST03 As String '建檔者的部門
Dim bolAutoNo As String  '是否自動取得順序
Dim m_Rows As Integer '目前修改的記錄位置

Private Sub cmdDgrid_Click(Index As Integer)
   
   Select Case Index
      Case 1 '新增
         ClearData2
         txtData(1).Text = m_SK01
         txtData(4).Text = "01"
         txtData(2).SetFocus
         m_iCAAEditMode = 1
         UpdateCUID 0
         
      Case 2 '加入
         bolAutoNo = False '使用者可輸入順序
         If TxtValidate1 = True Then
            UpdateCAA
            ClearData2
            m_iCAAEditMode = 0
            cmdDgrid(1).SetFocus '移到新增
         End If
            
      Case 3 '刪除
        If txtData(2) <> "" And txtData(3) <> "" Then
            '同一部門新增的資料才能修改或刪除；電腦中心人員除外；
            If m_ST03 <> "" And Pub_StrUserSt03 <> "M51" Then
               If (m_ST03 <> "M51" And m_ST03 <> Pub_StrUserSt03) Or (Pub_StrUserSt03 <> "M51" And m_SK01 <> txtData(1)) Then
                  MsgBox "同一部門新增的資料才能修改或刪除!", vbCritical + vbOKOnly, "檢核資料"
                  Exit Sub
               End If
            End If
            
            If Not (rsAssign.EOF Or rsAssign.BOF) Then
               If txtData(2) = rsAssign.Fields("CAA03") And txtData(3) = rsAssign.Fields("CAA04") Then
                  rsAssign.Delete
                  rsAssign.UpdateBatch
                  ClearData2
               End If
            End If
        Else
            MsgBox "無資料可刪除!", vbCritical + vbOKOnly
        End If
   End Select
End Sub

Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadAgent
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadAgent
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me

   stFormName = Me.Caption
   OnAction vbKeyF4
   
   If InStr(UCase(App.EXEName), "LAW") > 0 Then
      m_SK01 = "3"
   ElseIf InStr(UCase(App.EXEName), "TRADEMARK") > 0 Then
      m_SK01 = "2"
   Else
      m_SK01 = "1"
   End If
   
   textCUID.BackColor = &H8000000F
   
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
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
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
         
      Case vbKeyInsert
         If cmdDgrid(2).Enabled = True Then
            cmdDgrid_Click 2
         End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050718 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 修改
      Case 2: OnAction vbKeyF3
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
Private Sub ClearData()
   For Each oText In txtData
      oText.Text = Empty
   Next
   
   For Each oLbl In lblData
      oLbl.Caption = ""
   Next
   
   Combo1.Clear
   
   m_CU80 = ""
   m_FA69 = ""
   m_iCAAEditMode = 0
   bolAutoNo = True
End Sub
Private Sub ClearData2()
   For Each oText In txtData
      If oText.Index > 0 Then oText.Text = Empty
   Next
   
   For Each oLbl In lblData
      If oLbl.Index > 0 Then oLbl.Caption = ""
   Next
   m_FA69 = ""
   bolAutoNo = True
End Sub
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         cmdDgrid(1).SetFocus '移到新增
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         SetCtrlReadOnly True
         ClearData
         OpenCAATable '為了清除grid
         UpdateToolbarState
         SetInputEntry
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetCtrlReadOnly True
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtData(0) = txtData(0).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               txtData(0) = txtData(0).Tag
               If txtData(0) <> "" Then
                  SetInputEntry
                  ShowRecord
               Else
                  ClearData
               End If
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
         Exit Sub
   End Select
   
   Select Case m_EditMode
      Case 1
         Me.Caption = stFormName & "(新增)"
      Case 2
         Me.Caption = stFormName & "(修改)"
      Case 4
         Me.Caption = stFormName & "(查詢)"
      Case Else
         Me.Caption = stFormName
   End Select
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   cmdDgrid(1).Enabled = Not bLocked
   cmdDgrid(2).Enabled = Not bLocked
   cmdDgrid(3).Enabled = Not bLocked
   For Each oText In txtData
      oText.Locked = bLocked
   Next
   
   If m_EditMode = 4 Then
      txtData(0).Locked = False
   ElseIf m_EditMode = 2 Then
      txtData(0).Locked = True
   End If
   
   If Pub_StrUserSt03 = "M51" Then
      txtData(1).Enabled = True
   Else
      txtData(1).Enabled = False
   End If
End Sub


'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtData(0) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtData(0) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtData(0) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   
   '不允許上下筆
   TBar1.Buttons(6).Enabled = False
   TBar1.Buttons(7).Enabled = False
   TBar1.Buttons(8).Enabled = False
   TBar1.Buttons(9).Enabled = False
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 2
            txtData(0).Locked = True
         Case 4
            txtData(0).Locked = False
            txtData(0).SetFocus
         Case Else
            txtData(0).Locked = True
            txtData(0).SetFocus
      End Select
   End If
End Sub

Private Function TxtValidate() As Boolean
   
   Dim Cancel As Boolean, ii As Integer, jj As Integer

   '查詢
   If m_EditMode = 4 Then
      If txtData(0) = "" Then
         ShowMsg "請輸入欲查詢之客戶編號 !"
         txtData(0).SetFocus
         Txtdata_GotFocus 0
         Exit Function
     
      Else
        If Not IsEmptyText(txtData(0)) Then
           If Mid(txtData(0), 1, 1) <> "X" Then
              Cancel = True
              MsgBox "申請人編號必須為X開頭", vbCritical + vbOKOnly, "檢核資料"
              txtData(0).Text = ""
              Txtdata_GotFocus 0
              Exit Function
           End If
           
           If Len(txtData(0)) < 6 Then
              Cancel = True
              MsgBox "申請人編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
              Txtdata_GotFocus 0
              Exit Function
           End If
           
           If m_CU80 <> "" And (InStr(m_CU80, "不再使用") > 0 Or InStr(m_CU80, "不得代理") > 0) Then
              Cancel = True
              MsgBox "狀態欄為:" & m_CU80, vbCritical + vbOKOnly, "檢核資料"
              Txtdata_GotFocus 0
              Exit Function
           End If
           
           txtData(0).Text = Mid(txtData(0) & "00", 1, 8)
           
        End If
      End If
   End If
      
   If m_iCAAEditMode > 0 Then
      If MsgBox("CF代理人尚未加入,是否加入?", vbYesNo + vbInformation, "檢核資料") = vbYes Then
         Call cmdDgrid_Click(2)
         If m_iCAAEditMode > 0 Then
           Txtdata_GotFocus 3
           Exit Function
         End If
      End If
   End If

   
   TxtValidate = True
   
End Function


Private Function ModRecord() As Boolean
   Dim stSQL As String
   Dim stDiff As String
   Dim bolExist As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '有資料 => 無資料
   If rsAssign.RecordCount = 0 Then
      If rsAssignOld.RecordCount > 0 Then
         '刪除資料
         stSQL = "delete from CustAssignAgent where CAA01='" & rsAssignOld.Fields("CAA01") & "' and CAA02='" & rsAssignOld.Fields("CAA02") & "' "
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
   Else
      '刪除資料(原來的資料在新的資料中找不到的)
      With rsAssignOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsAssign.MoveFirst
            bolExist = False
            Do While Not (rsAssign.EOF Or bolExist = True)
               If rsAssign.Fields("CAA01") & rsAssign.Fields("CAA02") & rsAssign.Fields("CAA03") & rsAssign.Fields("CAA04") _
                 = .Fields("CAA01") & .Fields("CAA02") & .Fields("CAA03") & .Fields("CAA04") Then
                  bolExist = True
               End If
               rsAssign.MoveNext
            Loop
            
            If bolExist = False Then
                '刪除資料
                stSQL = "delete from CustAssignAgent where CAA01='" & .Fields("CAA01") & "' and CAA02='" & .Fields("CAA02") & "' and CAA03='" & .Fields("CAA03") & "' and CAA04='" & .Fields("CAA04") & "' "
                Pub_SeekTbLog stSQL
                cnnConnection.Execute stSQL, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      '新增/變更資料
      With rsAssign
      .MoveFirst
      Do While Not .EOF
         If rsAssignOld.RecordCount = 0 Then
            bolExist = False
         Else
            rsAssignOld.MoveFirst
            bolExist = False
            stDiff = ""
            Do While Not (rsAssignOld.EOF Or bolExist = True)
               If rsAssignOld.Fields("CAA01") & rsAssignOld.Fields("CAA02") & rsAssignOld.Fields("CAA03") & rsAssignOld.Fields("CAA04") _
                 = .Fields("CAA01") & .Fields("CAA02") & .Fields("CAA03") & .Fields("CAA04") Then
                  bolExist = True
                  If rsAssignOld.Fields("CAA05") <> .Fields("CAA05") Then
                     stDiff = "UPDATE CUSTASSIGNAGENT set CAA05=" & .Fields("CAA05") & ", CAA09='" & .Fields("CAA09") & "', CAA10=" & .Fields("CAA10") & ",CAA11=" & .Fields("CAA11") & _
                              " WHERE CAA01='" & .Fields("CAA01") & "' AND CAA02='" & .Fields("CAA02") & "' AND CAA03='" & .Fields("CAA03") & "' AND CAA04='" & .Fields("CAA04") & "' "
                  End If
               End If
               rsAssignOld.MoveNext
            Loop
         End If
         
         '新增資料
         If bolExist = False Then
            stSQL = "INSERT INTO CUSTASSIGNAGENT(CAA01,CAA02,CAA03,CAA04,CAA05,CAA06,CAA07,CAA08) VALUES ('" & .Fields("CAA01") & "'" & _
                    ",'" & .Fields("CAA02") & "','" & .Fields("CAA03") & "','" & .Fields("CAA04") & "'," & Val(.Fields("CAA05")) & ",'" & .Fields("CAA06") & "'," & .Fields("CAA07") & "," & .Fields("CAA08") & ")"
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         End If
         '更改順序
         If stDiff <> "" Then
            Pub_SeekTbLog stDiff
            cnnConnection.Execute stDiff, intI
         End If
         
         .MoveNext
      Loop
      End With
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtData(0).SetFocus
               Txtdata_GotFocus 2
            End If
         End If
         
   End Select
End Function

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stCAA01 As String
   Dim stCAA02 As String
   Dim adoRst As New ADODB.Recordset
   
   stCAA01 = Left(txtData(0) & "000", 8)
   stCAA02 = "0"
      
   Select Case p_iWay
      Case 0 '當筆
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN,CU80,CU10,(NA03) CU10N" & _
               " FROM Customer,Nation WHERE CU01 = '" & stCAA01 & "' AND CU02 = '0' AND CU10=NA01(+)"
        
      Case -2 '首筆
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN,CU80,CU10,(NA03) CU10N" & _
               " FROM Customer,Nation WHERE CU02='0' AND CU10=NA01(+) order by CU01 ASC"

      Case -1 '前筆
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN,CU80,CU10,(NA03) CU10N" & _
               " FROM Customer,Nation WHERE CU01<'" & stCAA01 & "' AND CU02='0' AND CU10=NA01(+) order by CU01 DESC"

      Case 1 '後筆
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN,CU80,CU10,(NA03) CU10N" & _
               " FROM Customer,Nation WHERE CU01 >'" & stCAA01 & "' AND CU02='0' AND CU10=NA01(+) order by CU01 ASC"

      Case 2 '末筆
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN,CU80,CU10,(NA03) CU10N" & _
               " FROM Customer,Nation WHERE CU02='0' AND CU10=NA01(+) order by CU01 DESC"

   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearData
      txtData(0) = "" & adoRst.Fields("NO")
      txtData(0).Tag = txtData(0)
      m_CU80 = "" & adoRst.Fields("CU80")
      lblData(0).Caption = Mid(adoRst.Fields("CU10"), 1, 3) & " " & adoRst.Fields("CU10N")
      Combo1.Clear
      Combo1.AddItem "中: " & adoRst.Fields("CN")
      Combo1.AddItem "英: " & adoRst.Fields("EN")
      Combo1.AddItem "日: " & adoRst.Fields("JN")
      Combo1.ListIndex = 0
      OpenCAATable
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         OpenCAATable '為了清除grid
         ClearData
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtData(0).SetFocus
      Txtdata_GotFocus 2
   End If
End Function

Private Sub OpenCAATable()
Dim strQ As String
On Error GoTo Checking

   If txtData(0) <> "" Then
      strQ = " CAA02='" & txtData(0) & "' "
   Else
      strQ = " 0=1 "
   End If
   
   strExc(0) = "SELECT CAA01,CAA03,CAA05,CAA04,NVL(FA04,NVL(FA05,FA06)) FANAME,CAA02,CAA06,CAA07,CAA08,CAA09,CAA10,CAA11" & _
                ",DECODE(CAA01,'1','專利','2','商標','3','法務',CAA01) CAA01N,N1.NA03 CAA03N,FA04,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) FA05,FA06,FA10,N2.NA03 FA10N,FA69 " & _
                "From CUSTASSIGNAGENT, FAGENT, Nation N1,NATION N2 " & _
                "WHERE" & strQ & "AND CAA04=FA01(+) AND FA02='0' AND CAA03=N1.NA01(+) AND FA10=N2.NA01(+) " & _
                "ORDER BY CAA01,CAA03,CAA05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set rsAssign = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set rsAssignOld = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set Adodc1.Recordset = rsAssign
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   If rsAssign.RecordCount > 0 Then
      ReadAgent
   End If
   
Checking:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

Private Function getNewNo() As String
   Dim myTemp As ADODB.Recordset
   Dim iUsableNo As Integer
   
   Set myTemp = rsAssign.Clone
   With myTemp
      .Sort = "CAA05 DESC"
      iUsableNo = 1
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            If Trim(txtData(1)) = "" & .Fields("CAA01") And Trim(txtData(2)) = "" & .Fields("CAA03") Then
               iUsableNo = Val(.Fields(2)) + 1
               Exit Do
            End If
            .MoveNext
         Loop
      End If
      getNewNo = Format(iUsableNo, "00")
   End With
   Set myTemp = Nothing
End Function

Private Sub UpdateCAA()

   With rsAssign
      If m_iCAAEditMode = 1 Then
        .AddNew
      Else
        If .RecordCount > 0 And m_Rows > 0 Then
           .MoveFirst
           '移動到要修改的資料
           Do While Not .EOF
              If .AbsolutePosition = m_Rows Then
                 m_Rows = 0
                 Exit Do
              End If
              .MoveNext
           Loop
           If m_Rows > 0 Then .AddNew
        Else
           .AddNew
        End If
      End If
      
          .Fields("CAA01") = Trim(txtData(1).Text)
          .Fields("CAA02") = Trim(txtData(0).Text)
          .Fields("CAA03") = Trim(txtData(2).Text)
          .Fields("CAA04") = Trim(txtData(3).Text)
          .Fields("CAA05") = Val(txtData(4).Text)
          .Fields("CAA06") = strUserNum
          .Fields("CAA07") = strSrvDate(1)
          .Fields("CAA08") = Left(Format(ServerTime, "000000"), 4)
          .Fields("CAA09") = strUserNum
          .Fields("CAA10") = strSrvDate(1)
          .Fields("CAA11") = Left(Format(ServerTime, "000000"), 4)
          .Fields("FANAME") = IIf(Trim(lblData(3)) = "", IIf(Trim(lblData(4)) = "", lblData(5), lblData(4)), lblData(3))
          .Fields("CAA01N") = IIf(Trim(txtData(1).Text) = "1", "專利", IIf(Trim(txtData(1).Text) = "2", "商標", "法務"))
          .Fields("CAA03N") = Trim(lblData(1).Caption)
          .Fields("FA04") = Trim(lblData(3).Caption)
          .Fields("FA05") = Trim(lblData(4).Caption)
          .Fields("FA06") = Trim(lblData(5).Caption)
          .Fields("FA10") = Trim(Mid(lblData(2).Caption, 1, 3))
          .Fields("FA10N") = Trim(Mid(lblData(2).Caption, 5))
          .Fields("FA69") = m_FA69
        .UPDATE
   End With
End Sub

Private Sub ReadAgent()
   
   ClearData2
   With rsAssign
    If Not .EOF Then
       txtData(1) = "" & .Fields("CAA01")
       txtData(2) = "" & .Fields("CAA03")
       lblData(1) = "" & .Fields("CAA03N")
       txtData(3) = "" & .Fields("CAA04")
       txtData(4) = Format("" & .Fields("CAA05"), "00")
       lblData(2) = "" & Mid(.Fields("FA10"), 1, 3) & " " & .Fields("FA10N")
       lblData(3) = "" & .Fields("FA04")
       lblData(4) = "" & .Fields("FA05")
       lblData(5) = "" & .Fields("FA06")
       m_FA69 = "" & .Fields("FA69")
    End If
   End With
   
   m_iCAAEditMode = 0
   
   UpdateCUID 1, rsAssign
End Sub

Private Function TxtValidate1() As Boolean
Dim idx As Integer
Dim Cancel As Boolean

   For Each oText In txtData
      If oText.Index > 0 Then
         idx = oText.Index
         Cancel = False
         Txtdata_Validate idx, Cancel
         If Cancel = True Then
            Txtdata_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
   If Trim(txtData(2)) = "" Or lblData(1).Caption = "" Then
      MsgBox "申請國家不可空白", vbCritical + vbOKOnly, "檢核資料"
      Txtdata_GotFocus 2
      Exit Function
   End If
   
   If lblData(3) = "" And lblData(4) = "" And lblData(5) = "" Then
      MsgBox "請輸入CF代理人", vbCritical + vbOKOnly, "檢核資料"
      Txtdata_GotFocus 3
      Exit Function
   End If
   
   If m_FA69 <> "" And (InStr(m_FA69, "不再使用") > 0 Or InStr(m_FA69, "不得代理") > 0) Then
      MsgBox "狀態欄為:" & m_FA69, vbCritical + vbOKOnly, "檢核資料"
      Txtdata_GotFocus 3
      Exit Function
   End If
    
   
   '同一部門新增的資料才能修改或刪除；電腦中心人員除外；
   If m_ST03 <> "" And Pub_StrUserSt03 <> "M51" Then
      If (m_ST03 <> "M51" And m_ST03 <> Pub_StrUserSt03) Or (Pub_StrUserSt03 <> "M51" And m_SK01 <> txtData(1)) Then
         MsgBox "同一部門新增的資料才能修改或刪除!", vbCritical + vbOKOnly, "檢核資料"
         Exit Function
      End If
   End If
   
   DataGrid1_Validate False
   m_Rows = rsAssign.AbsolutePosition
    With rsAssign
       If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             '檢查同一案件種類的代理人是否重複(以順序區別修改)
             If .Fields("CAA01") & .Fields("CAA03") & .Fields("CAA04") & Format(.Fields("CAA05"), "00") = Trim(txtData(1)) & Trim(txtData(2)) & Trim(txtData(3)) & Format(Val(txtData(4)), "00") Then
                MsgBox "同一申請國家的CF代理人重複！", vbCritical, "檢核資料"
                Exit Function
             End If
             '檢查同一案件種類和國家的順序是否重複
             If .Fields("CAA01") & .Fields("CAA03") & Format(.Fields("CAA05"), "00") = Trim(txtData(1)) & Trim(txtData(2)) & Format(Val(txtData(4)), "00") Then
                MsgBox "同一申請國家的順序重複！", vbCritical, "檢核資料"
                Exit Function
             End If
             .MoveNext
          Loop
       End If
    End With
   
   TxtValidate1 = True
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer, Optional ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True, , m_ST03)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   Else
        If IsNull(rsSrcTmp.Fields("CAA06")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA06")) = False Then
              strCName = GetStaffName(rsSrcTmp.Fields("CAA06"), True, , m_ST03)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("CAA07")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA07")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("CAA07"))
              strCDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("CAA08")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA08")) = False Then
              strTemp = rsSrcTmp.Fields("CAA08")
              strCTime = Format(strTemp, "00:00")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("CAA09")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA09")) = False Then
              strUName = GetStaffName(rsSrcTmp.Fields("CAA09"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("CAA10")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA10")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("CAA10"))
              strUDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("CAA11")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("CAA11")) = False Then
              strTemp = rsSrcTmp.Fields("CAA11")
              strUTime = Format(strTemp, "00:00")
           End If
        End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & " " & vbTab & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
   CloseIme
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer
   Select Case Index
       Case 2
            If txtData(Index) = "000" Then
               MsgBox "申請國家不可為臺灣!", vbCritical + vbOKOnly
               Cancel = True
               Exit Sub
            End If
            
            If txtData(Index).Text <> "" Then
               strExc(0) = PUB_GetNationName(txtData(Index).Text)
               If strExc(0) <> "" Then
                 lblData(1).Caption = strExc(0)
                 If bolAutoNo Then txtData(4) = getNewNo
               Else
                 lblData(1).Caption = ""
                 Cancel = True
               End If
            Else
                 lblData(1).Caption = ""
            End If
       Case 1
            If txtData(Index).Locked = False And txtData(Index).Enabled = True And _
                             txtData(Index).Text <> "1" And txtData(Index).Text <> "2" And txtData(Index).Text <> "3" Then
               MsgBox "請輸入1-3!", vbCritical + vbOKOnly
               Cancel = True
            End If
       Case 3
            If txtData(Index).Text <> "" Then
               txtData(Index).Text = Mid(txtData(Index).Text & "00000000", 1, 8)
               strExc(0) = "select fa04,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) fa05,fa06,na01,na03,fa69 from fagent,nation where fa01='" & txtData(Index).Text & "' and fa02='0' and fa10=na01(+) "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  lblData(2).Caption = "" & Mid(RsTemp.Fields("na01"), 1, 3) & " " & RsTemp.Fields("na03")
                  lblData(3).Caption = "" & RsTemp.Fields("fa04")
                  lblData(4).Caption = "" & RsTemp.Fields("fa05")
                  lblData(5).Caption = "" & RsTemp.Fields("fa06")
                  m_FA69 = "" & RsTemp.Fields("fa69")
               Else
                  MsgBox "查無資料！", vbInformation
                  m_FA69 = ""
                  lblData(2).Caption = ""
                  lblData(3).Caption = ""
                  lblData(4).Caption = ""
                  lblData(5).Caption = ""
                  Cancel = True
               End If
            End If
   End Select
   
   If Not CheckLengthIsOK(txtData(Index), iLen) Then
      Cancel = True
   End If
   
End Sub

