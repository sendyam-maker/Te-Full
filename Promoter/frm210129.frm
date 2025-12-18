VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210129 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內往來記錄維護"
   ClientHeight    =   5265
   ClientLeft      =   420
   ClientTop       =   4410
   ClientWidth     =   9195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9195
   Begin VB.ListBox lstAtt 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "frm210129.frx":0000
      Left            =   1080
      List            =   "frm210129.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3690
      Width           =   7305
   End
   Begin VB.TextBox txtCF 
      Height          =   270
      Index           =   2
      Left            =   600
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3990
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "-> 移除"
      Height          =   255
      Left            =   8430
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4260
      Width           =   735
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "<- 新增"
      Height          =   285
      Left            =   8430
      TabIndex        =   7
      Top             =   3990
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   8430
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtCF 
      Height          =   270
      Index           =   6
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   4560
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8190
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7530
      Top             =   60
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
            Picture         =   "frm210129.frx":0013
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":032F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":064B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":0827
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":0B43
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":0E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":117B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":1497
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":17B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":1ACF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210129.frx":1DEB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
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
   Begin MSForms.TextBox txtCOR 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   1140
      Width           =   1125
      VariousPropertyBits=   671107099
      MaxLength       =   7
      Size            =   "1984;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR 
      Height          =   1080
      Index           =   5
      Left            =   1080
      TabIndex        =   4
      Top             =   2580
      Width           =   7770
      VariousPropertyBits=   -1466941413
      MaxLength       =   4000
      ScrollBars      =   2
      Size            =   "13705;1905"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR 
      Height          =   750
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   7770
      VariousPropertyBits=   -1467989989
      MaxLength       =   200
      Size            =   "13705;1323"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR 
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   1470
      Width           =   1125
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1984;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR 
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   810
      Width           =   1125
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1984;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4920
      Width           =   6765
      VariousPropertyBits=   671105055
      Size            =   "11933;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   3750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來日期："
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   16
      Top             =   1185
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   2250
      TabIndex        =   15
      Top             =   1500
      Width           =   5100
      VariousPropertyBits=   27
      Caption         =   "lbl1"
      Size            =   "8996;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "內　　容："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "主　　旨："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "記錄編號："
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1500
      Width           =   900
   End
End
Attribute VB_Name = "frm210129"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (txtCOR,lbl1)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_COR06 As String          '2008/12/10 ADD BY SONIA

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim TF_COR As Integer
Dim strTmp As String
Dim oText As Object
Dim idx As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Private Const cTableName As String = "CONTACTFILE" 'Add By Sindy 2020/5/17 指定FTP資料夾名稱


'Add By Sindy 2020/5/17
'開啟附件
Private Sub cmdOpenAtt_Click()
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      If txtCF(6).Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(txtCF(6).Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, "(") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      End If
   End If
End Sub
'可多選, 顯示檔案大小
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strMid As String, strList As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.*"
   strList = txtCF(6).Text
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '存FTP檔名
               strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
               AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
            Next
         Else
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '存FTP檔名
            strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
            AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
         End If
         '上傳到FTP,故只需留檔名
         txtCF(2) = ComposeAttList(lstAtt)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub
Private Sub cmdRemAtt_Click()
   If InStr(lstAtt, "\") = 0 And Pub_StrUserSt03 <> "M51" Then
         MsgBox "已上傳檔案不可移除！"
   ElseIf RemoveList(lstAtt) = True Then
      txtCF(2) = ComposeList(lstAtt)
      cmdAddAtt.SetFocus
   End If
End Sub
'附件
Private Function ComposeAttList(oList As ListBox) As String
   Dim iPos As Integer, stItem As String, stRtn As String, idx As Integer
   If oList.ListCount > 0 Then
      stItem = oList.List(0)
      stRtn = GetFileName(stItem)
      For idx = 1 To oList.ListCount - 1
         stItem = oList.List(idx)
         stRtn = stRtn & "," & GetFileName(stItem)
      Next
   End If
   ComposeAttList = stRtn
End Function
Private Function AddListX(oList As ListBox, stNewItem As String, stFtpName As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
      cmdAddAtt.SetFocus
      Exit Function
   End If
   
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件[" & stFileName & "]已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         AddListX = True
         
         'Added by Lydia 2017/08/09 存FTP檔名 (堆疊)
         txtCF(6) = stFtpName & IIf(txtCF(6) <> "", ",", "") & txtCF(6)
      End If
   End If
End Function
Private Function RemoveList(oList As ListBox) As Boolean
   Dim ii As Integer
   Dim tmpArr As Variant 'Added by Lydia 2017/08/09
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            'Added by Lydia 2017/08/09 移除FTP檔名
            If txtCF(6) <> "" Then
               txtCF(6) = Replace(txtCF(6), ",,", ",")
               If Left(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 2)
               If Right(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 1, Len(txtCF(6)) - 1)
               tmpArr = Empty
               tmpArr = Split(txtCF(6), ",")
               If Trim(tmpArr(ii)) <> "" Then txtCF(6) = Replace(txtCF(6), Trim(tmpArr(ii)), "")
            End If
            'end 2017/08/09
            
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
      
      'Added by Lydia 2017/08/09 重整FTP檔名
      txtCF(6) = Replace(txtCF(6), ",,", ",")
      If Left(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 2)
      If Right(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 1, Len(txtCF(6)) - 1)
      'end 2017/08/09
      
   End If
End Function
Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   GetFileName = stItem
End Function
'解析檔案大小
Private Function GetFileSize(ByVal strFileOName As String, ByRef strFileNName As String) As Long
Dim strCF02 As String
Dim intEnd As Integer, intStar As Integer
   
   GetFileSize = 0
   strCF02 = UCase(strFileOName)
   If InStr(strCF02, "KB)") > 0 Then
      intEnd = InStrRev(strCF02, "KB)")
      intStar = InStrRev(strCF02, "(")
      strFileNName = Trim(Mid(strFileOName, 1, intStar - 1))
      GetFileSize = Val(Mid(strCF02, intStar + 1, Len(strCF02) - intEnd + 1))
   End If
End Function
'2020/5/17 END

Private Sub Form_Initialize()
   strExc(0) = "select * from ContactRecord1 where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_COR = RsTemp.Fields.Count
   ReDim m_FieldList(TF_COR) As FIELDITEM
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = True 'IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = True 'IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = True 'IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   InitialField
   m_EditMode = 0
   ShowRecord 99
   SetInputEntry
   UpdateToolbarState
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1 '新增
            txtCOR(1).Locked = True
            txtCOR(2).SetFocus
            
         Case 2 '修改
            txtCOR(1).Locked = True
            txtCOR(2).SetFocus
         
         Case 4 '查詢
            txtCOR(1).Locked = False
            txtCOR(1).SetFocus
            
         Case Else
            txtCOR(1).Locked = True
            txtCOR(1).SetFocus
      End Select
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
         If m_bUpdate And txtCOR(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtCOR(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtCOR(1) <> "" Then
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210129 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtCOR
      idx = oText.Index
      m_FieldList(idx).fiName = "COR" & Format(idx, "00")
   Next
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim adoRst As New ADODB.Recordset
'Dim intQ As Integer 'Added by Lydia 2019/09/10

Top:
   '2008/12/10 MODIFY BY SONIA 因做語文控制,故此處txtCOR(1)改為txtCOR(1).Tag,否則最後筆或第一筆時會因權限控制而改txtCOR(1)值
   Select Case p_iWay
      Case 0   '還原
         strExc(0) = "SELECT * FROM ContactRecord1" & _
            " WHERE COR01 = '" & txtCOR(1).Tag & "'"
      Case -2  '第一筆
         'Added by Lydia 2020/11/03 限操作者=建檔人
         If Pub_StrUserSt03 <> "M51" Then
             strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b WHERE COR06=" & CNULL(strUserNum) & ")"
         Else
         'end 2020/11/03
             strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b)"
         End If 'Added by Lydia 2020/11/03
      Case -1 '前一筆
         'Added by Lydia 2020/11/03 限操作者=建檔人
         If Pub_StrUserSt03 <> "M51" Then
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select max(b.COR01) from ContactRecord1 b where b.COR01<'" & txtCOR(1).Tag & "' and COR06=" & CNULL(strUserNum) & ")"
         Else
         'end 2020/11/03
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select max(b.COR01) from ContactRecord1 b where b.COR01<'" & txtCOR(1).Tag & "')"
         End If    'Added by Lydia 2020/11/03
      Case 1 '後一筆
         'Added by Lydia 2020/11/03 限操作者=建檔人
         If Pub_StrUserSt03 <> "M51" Then
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b where b.COR01>'" & txtCOR(1).Tag & "' and COR06=" & CNULL(strUserNum) & ")"
         Else
         'end 2020/11/03
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b where b.COR01>'" & txtCOR(1).Tag & "')"
         End If 'Added by Lydia 2020/11/03
      Case 2 '最後一筆
         'Added by Lydia 2020/11/03 限操作者=建檔人
         If Pub_StrUserSt03 <> "M51" Then
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select max(b.COR01) from ContactRecord1 b where COR06=" & CNULL(strUserNum) & ")"
         Else
         'end 2020/11/03
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select max(b.COR01) from ContactRecord1 b)"
         End If 'Added by Lydia 2020/11/03
      Case 99  '全部
         If Pub_StrUserSt03 <> "M51" Then
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b WHERE b.COR06='" & strUserNum & "')"
         Else
            strExc(0) = "SELECT * FROM ContactRecord1 a where COR01=(select min(b.COR01) from ContactRecord1 b)"
         End If
   End Select
   '2008/12/10 END
      
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Lydia 2019/09/10
      'If CheckModifyLimit(adoRst.Fields("COR06"), False) = False Then
      'Modify By Sindy 2020/5/21
      'If CheckModifyLimit(adoRst.Fields("COR06"), False, "" & adoRst.Fields("COR03")) = False Then
      'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
      'If PUB_GetCustData_frm100101_19("" & adoRst.Fields("COR03"), "" & adoRst.Fields("COR06")) = False Then
      ''2020/5/21 END
      If PUB_GetCustData_frm100101_19("" & adoRst.Fields("COR03"), "" & adoRst.Fields("COR06"), False) = False Then
         txtCOR(1).Tag = adoRst.Fields("COR01")
         Set adoRst = Nothing
         If p_iWay = -2 Then p_iWay = 1
         If p_iWay = 2 Then p_iWay = -1
         If p_iWay = 0 Then
            MsgBox "您沒有此筆潛在客戶往來記錄維護權限 !!!", vbInformation
            'Modified by Lydia 2020/11/11 重新載入全記錄
            'p_iWay = 1
            p_iWay = 99
            Call ShowRecord(p_iWay)
            'end 2020/11/11
         End If
      'Modify By Sindy 2020/5/27 Mark,intQ因會彈2次訊息
      'Modified by Lydia 2019/09/10
'          intQ = intQ + 1
'          If intQ = 1 Then GoTo Top
      Else
          UpdateCtrlData adoRst
      'end 2019/09/10
      End If
      'UpdateCtrlData adoRst 'Remove by Lydia 2019/09/10
      ShowRecord = True
   Else
      'Modified by Lydia 2019/09/10
      'If p_iWay = 0 Then
      '   MsgBox "查無資料！", vbInformation
      'ElseIf p_iWay = -1 Then
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtCOR(1).Tag = txtCOR(1)
         p_iWay = 0
         'GoTo Top
         '2008/12/10 END
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtCOR(1).Tag = txtCOR(1)
         p_iWay = 0
         'GoTo Top
         '2008/12/10 END
      Else
         ClearField
         MsgBox "查無資料！", vbInformation
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtCOR(1).SetFocus
      txtCOR_GotFocus 1
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
Dim CUID(1 To 6) As String
Dim AdoRs As New ADODB.Recordset 'Add By Sindy 2020/5/18

   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtCOR
            idx = oText.Index
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29

            If idx = 2 Then '往來日期
               oText.Text = ChangeWStringToTString(m_FieldList(idx).fiOldData)
            Else
               oText.Text = m_FieldList(idx).fiOldData
            End If
         Next
         CUID(1) = "" & .Fields("COR06")
         m_COR06 = "" & .Fields("COR06")   '2008/12/10 ADD BY SONIA
         CUID(2) = "" & .Fields("COR07")
         CUID(3) = "" & .Fields("COR08")
         CUID(4) = "" & .Fields("COR09")
         CUID(5) = "" & .Fields("COR10")
         CUID(6) = "" & .Fields("COR11")
         txtCOR_Validate 3, False
         
         'Add By Sindy 2020/5/17
         strExc(0) = "SELECT cf02,cf06,cf07 FROM ContactFile where CF01='" & .Fields("COR01") & "'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            AdoRs.MoveFirst
            Do While Not AdoRs.EOF
               txtCF(2) = txtCF(2) & "," & AdoRs.Fields("cf02") & IIf("" & AdoRs.Fields("cf07") <> "", " (" & AdoRs.Fields("cf07") & " KB)", "")
               txtCF(6) = txtCF(6) & "," & AdoRs.Fields("cf06")
               AdoRs.MoveNext
            Loop
            txtCF(2) = Mid(txtCF(2), 2)
            txtCF(6) = Mid(txtCF(6), 2)
         Else
            txtCF(2) = ""
            txtCF(6) = ""
         End If
         SetList lstAtt, txtCF(2)
         txtCF(6).Tag = txtCF(6).Text
         '2020/5/17 END
      End If
   End With
   UpdateCUID CUID, textCUID
   txtCOR(1).Tag = txtCOR(1)
   
   Set AdoRs = Nothing 'Add By Sindy 2020/5/17
End Sub

Private Sub ClearField()
   Dim oLabel As LABEL
   
   For Each oText In txtCOR
      oText.Text = Empty
      'Modified by Lydia 2020/11/03 排除流水號
      'oText.Tag = "" 'Added by Lydia 2019/09/10
      If oText.Index <> 1 Then oText.Tag = ""
   Next
   lbl1 = Empty
   
   If m_EditMode = 1 Then
      '新增時開發日期預設當天
      txtCOR(2) = ChangeWStringToTString(strSrvDate(1))
   End If
   For intI = 1 To TF_COR
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   'Add By Sindy 2020/5/17
   lstAtt.Clear
   For Each oText In txtCF
      oText.Text = Empty
      oText.Tag = ""
   Next
   '2020/5/17 END
End Sub

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ItemData(intI), "00")
         End If
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & "," & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function

'Private Function GetCustData(p_stCust As String) As Boolean
'Dim strName As String
'
'   GetCustData = False
'
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,cu13 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "' "
'      Case "R"
'         strExc(0) = "select '',NVL(POC03,NVL(RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26),POC27)),'','',poc04 N3,poc13 " & _
'                          "from potcustomer1 where poc01='" & Left(p_stCust, 8) & "' and poc02='" & Right(p_stCust, 1) & "' "
'      Case Else
'         MsgBox "往來對象必須為 X 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(intI)) Then
'            strName = RsTemp(intI)
'            Exit For
'         End If
'      Next
'      '2010/8/20 modify by sonia 開放電腦中心人員
'      'If Not IsNull(RsTemp(5)) And RsTemp(5) <> strUserNum Then
'      'modify by sonia 2017/8/10 開放主管也可以維護
'      'If Not IsNull(RsTemp(5)) And RsTemp(5) <> strUserNum And GetStaffDepartment(strUserNum) <> "M51" Then
'      'Modified by Lydia 2019/09/10 傳入客戶編號
'      'If Not IsNull(RsTemp(5)) And RsTemp(5) <> strUserNum And GetStaffDepartment(strUserNum) <> "M51" And CheckModifyLimit(m_COR06, False) = False Then
'      If Not IsNull(RsTemp(5)) And RsTemp(5) <> strUserNum And Pub_StrUserSt03 <> "M51" _
'               And CheckModifyLimit(m_COR06, False, p_stCust) = False Then
'
'         MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
'         Exit Function
'      End If
'   Else
'      MsgBox "往來對象輸入錯誤！"
'      Exit Function
'   End If
'   lbl1 = strName
'
'   GetCustData = True
'End Function

'Add by Sindy 2020/5/17
Private Sub lstAtt_DblClick()
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub

Private Sub txtCOR_Change(Index As Integer)
   If Index = 3 Then
      txtCOR(3).Tag = txtCOR(3).Text
   End If
End Sub

Private Sub txtCOR_GotFocus(Index As Integer)
   Select Case Index
      Case 4, 5
         OpenIme
         
      Case Else
         CloseIme
         
   End Select
   TextInverse txtCOR(Index)
End Sub

Private Sub txtCOR_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCOR_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer
Dim strName As String 'Add By Sindy 2020/5/21
   
   Select Case Index
      Case 3
         If txtCOR(Index) <> "" Then
            If Len(txtCOR(Index)) > 5 Then
               txtCOR(Index) = Left(txtCOR(Index) & "000", 9)
               'Modify By Sindy 2020/5/21
               'If GetCustData(txtCOR(Index)) = False Then
               lbl1 = "" 'Add By Sindy 2020/5/21
               'Added by Lydia 2020/11/03 載入記錄不用檢查
               If m_EditMode = 0 Then
                    lbl1.Caption = Pub_GetNameBYnation(txtCOR(Index))
               Else
               'end 2020/11/03
                    'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
                    'If PUB_GetCustData_frm100101_19(txtCOR(Index), , strName) = False Then
                    '2020/5/21 END
                    If PUB_GetCustData_frm100101_19(txtCOR(Index), "ADD", False, strName) = False Then
                       If m_EditMode = "1" Or m_EditMode = "2" Then
                          Cancel = True
                          txtCOR_GotFocus Index
                       End If
                    'Modify By Sindy 2020/5/21
                    Else
                       lbl1 = strName
                    '2020/5/21 END
                    End If
               End If 'Added by Lydia 2020/11/03
            Else
               Cancel = True
               MsgBox "往來對象編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtCOR_GotFocus Index
            End If
         End If
         
      Case 2
         If txtCOR(Index) <> "" Then
            If ChkDate(txtCOR(Index)) = False Then
               txtCOR_GotFocus Index
               Cancel = True
            End If
         End If
   End Select
   
   If Cancel = False Then
      If txtCOR(Index).MaxLength > 0 Then
         Select Case Index
            '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
            Case 4, 5
               iLen = txtCOR(Index).MaxLength - 1
            Case Else
               iLen = txtCOR(Index).MaxLength
         End Select
         If Not CheckLengthIsOK(txtCOR(Index), iLen) Then
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

' 執行指令
'Private Sub OnAction(ByVal KeyCode As Integer)
Public Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF3 ' 修改
         'Modified by Lydia 2019/09/10
         'If CheckModifyLimit(m_COR06, True) = False Then Exit Sub
         'Modify By Sindy 2020/5/21
         'If CheckModifyLimit(m_COR06, True, txtCOR(3).Text) = False Then Exit Sub
         'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
         'If PUB_GetCustData_frm100101_19(txtCOR(3).Text, m_COR06) = False Then Exit Sub
         ''2020/5/21 END
         If PUB_GetCustData_frm100101_19(txtCOR(3).Text, m_COR06, False) = False Then Exit Sub
         
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF5 ' 刪除
         If Pub_StrUserSt03 <> "M51" Then
            MsgBox "無刪除權限 !!!", vbInformation
            Exit Sub
         End If
         'If CheckModifyLimit(m_COR06, True) = False Then Exit Sub
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         SetCtrlReadOnly True
         ClearField
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
         If m_EditMode = 4 Then txtCOR(1).Tag = txtCOR(1) '2008/12/10 ADD BY SONIA

         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtCOR(1) = txtCOR(1).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               txtCOR(1) = txtCOR(1).Tag
               m_EditMode = 0
               SetInputEntry
               'Modified by Lydia 2019/09/10
               'ShowRecord
               If txtCOR(1).Tag <> "" Then ShowRecord
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtCOR(1).SetFocus
               txtCOR_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean, ii As Integer, jj As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant 'Add By Sindy 2020/5/18
   
   'Added by Morgan 2022/1/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/20
   
   For Each oText In txtCOR
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         Cancel = False
         txtCOR_Validate oText.Index, Cancel
         If Cancel = True Then
            oText.SetFocus
            txtCOR_GotFocus oText.Index
            Exit Function
         End If
      End If
   Next
   '查詢
   If m_EditMode = 4 Then
      If txtCOR(1) = "" Then
         ShowMsg "請輸入欲查詢之往來記錄編號 !"
         txtCOR(1).SetFocus
         txtCOR_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      If txtCOR(2).Text = "" Then
         ShowMsg "往來日期不可為空白 !"
         txtCOR(2).SetFocus
         Exit Function
      End If
     
      If txtCOR(3).Text = "" Then
         ShowMsg "往來對象不可為空白 !"
         txtCOR(3).SetFocus
         Exit Function
      End If
      
      If txtCOR(4).Text = "" Then
         ShowMsg "主旨不可為空白 !"
         txtCOR(4).SetFocus
         Exit Function
      End If
      
      If txtCOR(5).Text = "" Then
         ShowMsg "內容不可為空白 !"
         txtCOR(5).SetFocus
         Exit Function
      End If
      
      'Add By Sindy 2020/5/18
      '檢查長度
      If CheckLengthIsOK(txtCF(2), 800, False) = False Then
         MsgBox "全部的附件檔名超過最大長度！" & vbCrLf & "(1個中文=2個字元)", vbCritical
         Exit Function
      End If
      '檢查List和FTP檔名的數量是否一致
      strExc(1) = "附件順序有誤，請全部移除後再新增附件"
      If (txtCF(2) = "" And txtCF(6) <> "") Or (txtCF(2) <> "" And txtCF(6) = "") Then
          ShowMsg strExc(1)
          Exit Function
      End If
      tmpArr1 = Empty: tmpArr2 = Empty
      tmpArr1 = Split(txtCF(2), ",")
      tmpArr2 = Split(txtCF(6), ",")
      If UBound(tmpArr1) <> UBound(tmpArr2) Then
          ShowMsg strExc(1)
          Exit Function
      End If
      '預估一個ftp路徑約50字
      If UBound(tmpArr2) > Format(1100 / 50, "0") Then
         MsgBox "附件數量超過最大上限(" & Format(1100 / 50, "0") & ")！", vbCritical
         Exit Function
      End If
      For intI = 0 To UBound(tmpArr1)
         If (Trim(tmpArr1(intI)) = "" And Trim(tmpArr2(intI)) <> "") Or (Trim(tmpArr1(intI)) <> "" And Trim(tmpArr2(intI)) = "") Then
            ShowMsg strExc(1)
            Exit Function
         End If
      Next intI
      '2020/5/18
   End If
   
   TxtValidate = True
   
End Function

Private Sub UpdateFieldNewData()
   For Each oText In txtCOR
      idx = oText.Index
      Select Case idx
         Case 2 '往來日期
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim stSQL As String, stCols As String, stValues As String
'Add By Sindy 2020/5/18
Dim iErr As Integer, sErrMsg As String
Dim varTemp1, varTemp2
Dim j As Integer
Dim strNewFile As String, longSize As Long
'2020/5/18 END
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans

   If txtCOR(1) = "" Then
      m_FieldList(1).fiNewData = AutoNo("K", 6)
   End If

   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtCOR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO ContactRecord1 (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL
   
   'Add By Sindy 2020/5/17
   If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
      GoTo ErrHand
   Else
      If txtCF(6).Text <> txtCF(6).Tag Then
         '先全部刪除附件,再新增附件資訊
         varTemp1 = Split(txtCF(2), ",")
         varTemp2 = Split(txtCF(6), ",")
         If UBound(varTemp1) = UBound(varTemp2) Then
            For j = 0 To UBound(varTemp1)
               longSize = GetFileSize(varTemp1(j), strNewFile)
               If longSize = 0 Then
                  MsgBox "檔案大小為 0 有問題，請確認附件內容！"
                  GoTo ErrHand
               End If
               stSQL = "INSERT INTO CONTACTFILE(cf01,cf02,cf06,cf07)" & _
                        "VALUES('" & m_FieldList(1).fiNewData & "','" & ChgSQL(strNewFile) & "','" & ChgSQL(varTemp2(j)) & "','" & longSize & "')"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL
            Next j
         Else
            MsgBox "附件資料檔名和路徑個數有誤，無法儲存！"
            GoTo ErrHand
         End If
      End If
   End If
   '2020/5/17 END
   
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtCOR(1) = m_FieldList(1).fiNewData
   txtCOR(1).Tag = txtCOR(1)     '2008/12/10 ADD BY SONIA
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim stSQL As String
Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   'Add By Sindy 2020/5/17
   If MsgBox(IIf(txtCF(2) <> "", "有附件", "") & "是否要刪除此筆往來記錄資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
   '2020/5/17 End
      cnnConnection.BeginTrans
      
      'Add By Sindy 2020/5/17 判斷移檔日期
      If txtCF(6) <> "" Then
         txtCF(6) = "" '刪除全部附件
         If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
            GoTo ErrHand
         End If
      End If
      '刪除附件
      stSQL = "delete from ContactFile where cf01='" & txtCOR(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL
      '2020/5/17 END
      
      stSQL = "delete from ContactRecord1 where cor01='" & txtCOR(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL
      
      cnnConnection.CommitTrans
      
      DelRecord = True
      ClearField
      txtCOR(1).Tag = ""
   End If
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function ModRecord() As Boolean
Dim stSQL As String, stSet As String, stCols As String, stValues As String
Dim bDifference As Boolean, bAddNew As Boolean
'Add By Sindy 2020/5/18
Dim iErr As Integer, sErrMsg As String
Dim ii As Integer
Dim arrTmp, arrOldTmp, varTemp1
Dim strNewFile As String, longSize As Long
'2020/5/18 END
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE ContactRecord1 SET "
   stSet = ""
   For Each oText In txtCOR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where cor01='" & txtCOR(1) & "'; end; "
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   If txtCF(6).Text <> txtCF(6).Tag Then
      If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
         GoTo ErrHand
      Else
         '異動附件資訊
         If bDifference = False Then
            'Update人,日,時
            stSQL = "UPDATE ContactRecord1 SET cor09='" & strUserNum & "',cor10=" & strSrvDate(1) & ",cor11=" & Left(Format(ServerTime, "000000"), 4) & " WHERE COR01='" & txtCOR(1) & "'"
            cnnConnection.Execute stSQL
         End If
         arrTmp = Empty: arrOldTmp = Empty: varTemp1 = Empty
         varTemp1 = Split(txtCF(2), ",")
         arrTmp = Split(txtCF(6), ",")
         arrOldTmp = Split(txtCF(6).Tag, ",")
         '先：刪除附件
         If txtCF(6).Tag <> "" Then
            For ii = 0 To UBound(arrOldTmp)
               If Trim(arrOldTmp(ii)) <> "" And InStr(txtCF(6), Trim(arrOldTmp(ii))) = 0 Then
                  stSQL = "delete from CONTACTFILE where cf01='" & txtCOR(1) & "' and upper(cf06)='" & ChgSQL(UCase(arrOldTmp(ii))) & "'"
                  Pub_SeekTbLog stSQL
                  cnnConnection.Execute stSQL
               End If
            Next ii
         End If
         '後：新增附件
         If txtCF(6) <> "" Then
            For ii = 0 To UBound(arrTmp)
               If Trim(arrTmp(ii)) <> "" And InStr(txtCF(6).Tag, Trim(arrTmp(ii))) = 0 Then
                  longSize = GetFileSize(varTemp1(ii), strNewFile)
                  If longSize = 0 Then
                     MsgBox "檔案大小為 0 有問題，請確認附件內容！"
                     GoTo ErrHand
                  End If
                  stSQL = "INSERT INTO CONTACTFILE(cf01,cf02,cf06,cf07)" & _
                          "VALUES('" & txtCOR(1) & "','" & ChgSQL(strNewFile) & "','" & ChgSQL(arrTmp(ii)) & "','" & longSize & "')"
                  Pub_SeekTbLog stSQL
                  cnnConnection.Execute stSQL
               End If
            Next ii
         End If
      End If
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCOR
      oText.Locked = bLocked
   Next
   
   'Add By Sindy 2020/5/18
   cmdOpenAtt.Enabled = bLocked
   cmdAddAtt.Enabled = Not bLocked
   cmdRemAtt.Enabled = Not bLocked
   '2020/5/18 END
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

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
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
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
   End Select
End Sub

''檢查維護權限
''Modified by Lydia 2019/09/10 +傳入客戶編號
'Private Function CheckModifyLimit(strChkID As String, bType As Boolean, bCustNo As String) As Boolean
'
'   CheckModifyLimit = True
'
'   If Trim(strChkID) = "" Then Exit Function
'
'   '2009/5/14 add by sonia 開放M51權限
'   If Pub_StrUserSt03 = "M51" Then
'      Exit Function
'   End If
'   '2009/5/14 end
'
'   'LoginUser須為智權人員或其案件主管, 方可維護此筆資料
'   If strUserNum = Trim(strChkID) Then
'      Exit Function
'   Else
'      'modify by sonia 2017/8/10 A0909->A0908
'      strExc(0) = "SELECT A0908 FROM STAFF,ACC090 " & _
'                         "WHERE ST03=A0901(+) and ST01= '" & Trim(strChkID) & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If strUserNum = RsTemp(0) Then Exit Function
'      End If
'   End If
'
'   CheckModifyLimit = False
'   If bType = True Then
'      MsgBox "無修改權限 !!!", vbInformation
'   End If
'End Function
