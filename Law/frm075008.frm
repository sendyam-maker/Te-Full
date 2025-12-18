VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075008 
   BorderStyle     =   1  '單線固定
   Caption         =   "合約基本資料維護"
   ClientHeight    =   4320
   ClientLeft      =   615
   ClientTop       =   810
   ClientWidth     =   7830
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7830
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
      Height          =   780
      ItemData        =   "frm075008.frx":0000
      Left            =   705
      List            =   "frm075008.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3090
      Width           =   5900
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "-> 移除"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   3630
      Width           =   735
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "<- 新增"
      Height          =   285
      Left            =   6720
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3090
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   630
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
            Picture         =   "frm075008.frx":0013
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":032F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":064B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":0827
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":0B43
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":0E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":117B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":1497
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":17B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":1ACF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075008.frx":1DEB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7200
      Top             =   1000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox textCT 
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3690
      Visible         =   0   'False
      Width           =   360
      VariousPropertyBits=   671105051
      Size            =   "635;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   288
      Index           =   6
      Left            =   3030
      TabIndex        =   4
      Top             =   1560
      Width           =   800
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1411;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   285
      Index           =   8
      Left            =   60
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
      VariousPropertyBits=   671105051
      Size            =   "635;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   285
      Index           =   5
      Left            =   3030
      TabIndex        =   5
      Top             =   1890
      Width           =   300
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "529;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   288
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   900
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1587;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   288
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   5500
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "9701;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   285
      Index           =   2
      Left            =   3540
      TabIndex        =   1
      Top             =   840
      Width           =   1020
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1799;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   288
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   900
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1587;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCT 
      Height          =   690
      Index           =   7
      Left            =   150
      TabIndex        =   6
      Top             =   2220
      Width           =   7335
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "12938;1217"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblStaffName 
      Height          =   285
      Left            =   3900
      TabIndex        =   24
      Top             =   1560
      Width           =   1050
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "1852;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCusName 
      Height          =   285
      Left            =   4620
      TabIndex        =   23
      Top             =   840
      Width           =   3045
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "5371;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label23 
      Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3990
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   210
      Index           =   5
      Left            =   2070
      TabIndex        =   18
      Top             =   1599
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   150
      TabIndex        =   17
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "機密：　　   （空:一般 C:國內 F:國外)"
      Height          =   210
      Index           =   4
      Left            =   2430
      TabIndex        =   16
      Top             =   1950
      Width           =   3345
   End
   Begin VB.Label Label1 
      Caption         =   "簽約日期："
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1599
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "合約名稱："
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1239
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "合約編號："
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   879
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "客戶/代理人編號："
      Height          =   210
      Index           =   0
      Left            =   2070
      TabIndex        =   9
      Top             =   879
      Width           =   1500
   End
   Begin VB.Label Label7 
      Caption         =   "備註："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1950
      Width           =   975
   End
End
Attribute VB_Name = "frm075008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; lblCusName、lblStaffName、textCT(index)
'Create by Amy 2017/12/20
Option Explicit

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料
Dim m_FirstKEY As String
' 最後一筆資料
Dim m_LastKEY As String
' 目前正在顯示
Dim m_CurrKEY As String
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean, m_bQuery As Boolean
Private Const stFolderN As String = "CONTRACT" '指定FTP資料夾名稱
Dim RbMain As New ADODB.Recordset, bp As New ADODB.Recordset
Dim Max_CT As Integer
Dim i As Integer
'Modified by Lydia 2021/09/15
'Dim oText As TextBox
Dim oText As Control
Dim stOldCT0809 As String '原始檔名-DB實體路徑
Dim strCDate As String, strQ As String
Dim ActionEdit As Integer '按鍵itemIndex(不記錄確定/取消)

' 初始化欄位陣列
Private Sub InitialField()
    Dim nIndex As Integer
    Dim strTmp As String
    
    ' 初始化欄位陣列
    For nIndex = 1 To Max_CT
        strTmp = Format(nIndex, "00")
        m_FieldList(nIndex).fiName = "CT" & strTmp
        m_FieldList(nIndex).fiOldData = Empty
        m_FieldList(nIndex).fiNewData = Empty
        m_FieldList(nIndex).fiType = 0  '文字型態
        Select Case nIndex
            Case 4, 11, 12, 14, 15
            m_FieldList(nIndex).fiType = 1  '數值型態
        End Select
    Next nIndex
End Sub

Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
    Dim arrCT08 As Variant, arrCT09 As Variant
    Dim ii As Integer
    
    stOldCT0809 = ""
    For Each oText In textCT
        If m_FieldList(oText.Index).fiName <> Empty Then
            If IsNull(rsTmp.Fields(m_FieldList(oText.Index).fiName)) = False Then
                m_FieldList(oText.Index).fiOldData = rsTmp.Fields(m_FieldList(oText.Index).fiName)
                m_FieldList(oText.Index).fiNewData = rsTmp.Fields(m_FieldList(oText.Index).fiName)
            Else
                m_FieldList(oText.Index).fiOldData = Empty
                m_FieldList(oText.Index).fiNewData = Empty
            End If
        End If
    Next
    
    '記錄原始檔名-DB實體路徑
    If "" & rsTmp.Fields(8) <> MsgText(601) Then
        arrCT08 = Split(rsTmp.Fields(7), ",")
        arrCT09 = Split(rsTmp.Fields(8), ",")
        For ii = LBound(arrCT08) To UBound(arrCT08)
            stOldCT0809 = stOldCT0809 & "," & arrCT08(ii) & "##" & arrCT09(ii)
        Next ii
        stOldCT0809 = Mid(stOldCT0809, 2)
    End If
End Sub

Private Sub UpdateFieldNewData()
    
    For Each oText In textCT
        Select Case oText.Index
            Case 4 '簽約日期
                m_FieldList(oText.Index).fiNewData = Val(oText.Text) + 19110000
            Case Else
                m_FieldList(oText.Index).fiNewData = oText.Text
        End Select
    Next
End Sub

Private Sub cmdAddAtt_Click()
    Dim sFile, fs, f, s
    Dim ii As Integer
    Dim stFileName As String, strMid As String, strList As String
On Error GoTo ErrHnd
    
    stFileName = "*.*"
    
    With CommonDialog1
        .CancelError = True
        .FileName = stFileName
        .Filter = "All Files (*.*)|*.*"
        .InitDir = PUB_Getdesktop
        .MaxFileSize = 3000
        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .ShowOpen
        If .FileName <> MsgText(601) Then
            '實體路徑:合約年度\檔名
            If InStr(.FileName, ChrW$(0)) > 0 Then
                sFile = Split(.FileName, ChrW$(0))
                For i = 1 To UBound(sFile)
                    If InStr(CStr(sFile(ii)), "#") > 0 Then
                        MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                        Exit Sub
                    End If
                    If InStr(sFile(i), "\") > 0 Then
                        stFileName = sFile(ii)
                    Else
                        stFileName = sFile(0) & "\" & sFile(ii)
                    End If
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set f = fs.GetFile(stFileName)
                    If f.Size = 0 Then
                        ShowMsg sFile(ii) & MsgText(9221)
                        Exit Sub
                    End If
                    strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
                   AddList lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
                Next i
            Else
                stFileName = .FileName
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(stFileName)
                strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
                If InStr(CStr(strMid), "#") > 0 Then
                    MsgBox CStr(strMid) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                    Exit Sub
                ElseIf f.Size = 0 Then
                    ShowMsg sFile(ii) & MsgText(9221)
                    Exit Sub
                End If
                AddList lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
            End If
        End If
    End With
    Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdOpenAtt_Click()
    Dim tmpArr As Variant, ii As Integer, jj As Integer
    Dim hLocalFile As Long
    Dim stPath As String, stFileName As String, bolSelect As Boolean
    Dim stSavePath As String 'Added by Morgan 2022/8/11
    
    If lstAtt.Text = "" Then
        MsgBox "請選擇欲開啟的附件！"
        Exit Sub
    End If
    
    stSavePath = App.path & "\" & strUserNum 'Added by Morgan 2022/8/11
   
    Screen.MousePointer = vbHourglass
    tmpArr = Empty
    tmpArr = Split(stOldCT0809, ",")
    For ii = 0 To lstAtt.ListCount - 1
        If lstAtt.Selected(ii) Then
            bolSelect = True
            stFileName = lstAtt.List(ii)
            
            'Removed by Morgan 2022/8/11 檔名會" ("時會連附檔名也被截掉導致無法開啟
            'If InStrRev(stFileName, " (") > 0 Then
            '   stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            'End If
            'end 2022/8/11
            
            For jj = LBound(tmpArr) To UBound(tmpArr)
                If InStr(tmpArr(jj), stFileName) > 0 Then
                    stPath = Mid(tmpArr(jj), InStr(tmpArr(jj), "##") + 2)
                End If
            Next jj
            stFileName = stSavePath & "\" & stFileName 'Added by Morgan 2022/8/11
            
            If PUB_GetFtpFile(stPath, stFileName, stFolderN) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
        End If
    Next ii
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRemAtt_Click()
    If RemoveList(lstAtt) = True Then
        cmdAddAtt.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    CheckOC3
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open "Select * From Contract Where Rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
    Max_CT = AdoRecordSet3.Fields.Count
    ReDim m_FieldList(1 To Max_CT) As FIELDITEM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2: '新增
            Call OnAction(1)
            KeyCode = 0
        Case vbKeyF3: '修改
            Call OnAction(2)
            KeyCode = 0
        Case vbKeyF4: '查詢
            Call OnAction(4)
            KeyCode = 0
        Case vbKeyF5: '刪除
            Call OnAction(3)
            KeyCode = 0
        Case vbKeyHome: '第一筆
            Call OnAction(6)
            KeyCode = 0
        Case vbKeyPageDown: '下一筆
            Call OnAction(7)
            KeyCode = 0
        Case vbKeyPageUp: '上一筆
            Call OnAction(8)
            KeyCode = 0
        Case vbKeyEnd: '最後一筆
            Call OnAction(9)
            KeyCode = 0
        Case vbKeyF9: '確定
            Call OnAction(11)
            KeyCode = 0
        Case vbKeyF10: '取消
            Call OnAction(12)
            KeyCode = 0
        Case vbKeyEscape:
        Case vbKeyReturn
            '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
            KeyCode = 0
            If ActionEdit <> 11 Then
                Call OnAction(11)
            End If
    End Select
End Sub

Private Sub Form_Load()
    
    MoveFormToCenter Me
    TxtClear
    InitialField
    ActionEdit = 12
    RefreshRange
    Call GetRecordVal("First", False)
    Call SetToolBarAndBT  '設定ToolBar按鈕顯示
    Call LockTxt
End Sub

Private Sub RefreshRange()
    Dim rsTmp As New ADODB.Recordset
    
    strQ = "Select Min(CT01) as CT01 From Contract Having Min(CT01)>0"
                                         
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("CT01")) = False Then: m_FirstKEY = rsTmp.Fields("CT01")
    End If
    rsTmp.Close
    
    strQ = "Select Max(CT01) as CT01 From Contract Having Max(CT01)>0"
                                         
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("CT01")) = False Then: m_LastKEY = rsTmp.Fields("CT01")
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Sub

Private Sub GetRecordVal(ByVal strStatus As String, Optional ByVal bolShowMsg As Boolean = True)
    Dim rsTmp As New ADODB.Recordset
    Dim strSign As String, strFun As String
    
    Select Case strStatus
        Case "Pre" '上一筆資料
            If m_CurrKEY = m_FirstKEY Then
                If bolShowMsg = True Then ShowMsg MsgText(9008)
                Exit Sub
            End If
            strSign = "<"
            strFun = "Max"
        Case "Next" '下一筆資料
            If m_CurrKEY = m_LastKEY Then
                If bolShowMsg = True Then ShowMsg MsgText(9009)
                Exit Sub
            End If
            strSign = ">"
            strFun = "Min"
        Case "First", "Last" '第一/最後 一筆
            If strStatus = "First" Then
                m_CurrKEY = m_FirstKEY
            Else
                m_CurrKEY = m_LastKEY
            End If
            Call QueryRecord(m_CurrKEY, False, True)
            Exit Sub
    End Select
    
    strQ = "Select " & strFun & "(CT01) as CT01 From Contract  Where CT01 " & strSign & m_CurrKEY
                                         
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("CT01")) = False Then: m_CurrKEY = rsTmp.Fields("CT01")
    End If
    rsTmp.Close
    Set rsTmp = Nothing
    Call QueryRecord(m_CurrKEY, False, True)
End Sub

' 顯示資料
Private Sub GetCurrRecordVal(ByVal strKey As String)
    
    If QueryRecord(strKey, False, True) = True Then
        'QueryRecord 設 m_CurrKEY
    Else
        Call GetRecordVal("Pre", False)
        If QueryRecord(m_FirstKEY, False, True) = False Then
            Call GetRecordVal("Next", False)
            If QueryRecord(m_FirstKEY, False, True) = False Then
                Call QueryRecord(m_LastKEY, False, True)
            End If
        End If
    End If
    RefreshRange
End Sub

Private Sub SetTxtValue(ByRef rsTmp As ADODB.Recordset)
    Dim idx As Integer
  
    For Each oText In textCT
        idx = oText.Index - 1
        If IsNull(rsTmp(idx)) Then
            oText = ""
        Else
            Select Case idx
                Case 3
                    oText = ChangeWStringToTString("" & rsTmp(idx))
                Case Else
                    oText = "" & rsTmp(idx)
            End Select
        End If
    Next
    '客戶/代理人
    If textCT(2) <> MsgText(601) Then
        Call textCT_Validate(2, False)
    End If
    '智權人員
    If textCT(6) <> MsgText(601) Then
        Call textCT_Validate(6, False)
    End If
    'Add by Amy 2019/05/02
    textCT(5).Tag = textCT(5) & textCT(6)
  
    '檔案列表
    SetList lstAtt, textCT(8)
    '更新CUID
    UpdateCUID rsTmp
    ' 更新暫存區的資料
    UpdateFieldOldData rsTmp
    
End Sub

'更新 Create 及 Update 資訊
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
    Dim strTemp As String, strCName As String, strCTime As String
    Dim strUName As String, strUDate As String, strUTime As String
   
    strCDate = ""
    If IsNull(rsSrcTmp.Fields("CT10")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("CT10")) = False Then
            strCName = GetStaffName(rsSrcTmp.Fields("CT10"), True)
        End If
    End If
    If IsNull(rsSrcTmp.Fields("ct11")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("ct11")) = False Then
            strTemp = TAIWANDATE(rsSrcTmp.Fields("ct11"))
            strCDate = Format(strTemp, "###/##/##")
        End If
    End If
    If IsNull(rsSrcTmp.Fields("ct12")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("ct12")) = False Then
            strTemp = rsSrcTmp.Fields("ct12")
            strCTime = Format(strTemp, "##:##")
        End If
    End If
    If IsNull(rsSrcTmp.Fields("ct13")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("ct13")) = False Then
            strUName = GetStaffName(rsSrcTmp.Fields("ct13"), True)
        End If
    End If
    If IsNull(rsSrcTmp.Fields("ct14")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("ct14")) = False Then
            strTemp = TAIWANDATE(rsSrcTmp.Fields("ct14"))
            strUDate = Format(strTemp, "###/##/##")
        End If
    End If
    If IsNull(rsSrcTmp.Fields("ct15")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("ct15")) = False Then
            strTemp = rsSrcTmp.Fields("ct15")
            strUTime = Format(strTemp, "##:##")
        End If
    End If
   
    ' 設定CUID中的文字
    Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

'Add by Amy 2019/05/02
Private Sub textCT_Change(Index As Integer)
    If Index <> 5 And Index <> 6 Then Exit Sub
    textCT(Index).Tag = ""
End Sub
'end 2019/05/02

'Modified by Lydia 2021/09/15 改成Form 2.0
'Private Sub textCT_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub textCT_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If Not (Index = 2 Or Index = 5 Or Index = 6) Then Exit Sub
    
    KeyAscii = UpperCase(KeyAscii)
    If Index = 5 Then
        'Modify by Amy 2019/05/02 原:機密=Y,拆成國內機密(C)/國外機密(F)
        If KeyAscii <> 8 And KeyAscii <> Asc("C") And KeyAscii <> Asc("F") Then
            KeyAscii = 0
            Beep
         End If
    End If
End Sub

Private Sub textCT_Validate(Index As Integer, Cancel As Boolean)
    Dim stTP As String, stName As String, stCU13 As String
    Dim bolMsg As Boolean
    Dim stMsg As String  'Add by Amy 2019/05/02
        
    bolMsg = True
    If ActionEdit >= 3 = True Then bolMsg = False
    If textCT(Index) = MsgText(601) Then Exit Sub
    
    Select Case Index
        Case 1 '合約編號(key)
            If bolMsg = True And Not IsNumeric(Val(textCT(Index))) Then
                Cancel = True
                MsgBox "合約編號輸入錯誤", , MsgText(5)
                textCT(Index).SetFocus
                TextInverse textCT(Index)
                Exit Sub
            End If
        Case 2 '客戶/代理人編號(8 碼)
            If bolMsg = True Then
                'Modify by Amy 2019/05/02 增加代理人合約
                If Left(Trim(textCT(Index)), 1) <> "X" And Left(Trim(textCT(Index)), 1) <> "Y" Then
                    Cancel = True
                    MsgBox "客戶/代理人編號輸入錯誤", , MsgText(5)
                    textCT(Index).SetFocus
                    TextInverse textCT(Index)
                    Exit Sub
                End If
            End If
            textCT(Index).Text = GetNewFagent(textCT(Index))
            textCT(Index) = Mid(textCT(Index), 1, 8) 'Added by Lydia 2021/09/15 只存前8碼
            stCU13 = "Y"
            'Modify by Amy 2021/09/22 ChkCusAgent搬至basQuery
            If ChkCusAgent(textCT(2).Text, stName, stCU13) = False Then
                If bolMsg = True Then
                    Cancel = True
                    'Modify by Amy 2019/05/02 增加代理人合約
                    MsgBox "無此" & IIf(Left(Trim(textCT(Index)), 1) = "X", "客戶", "代理人") & "請確認", , MsgText(5)
                    textCT(Index).SetFocus
                    TextInverse textCT(Index)
                    Exit Sub
                End If
            End If
            lblCusName = stName
            '新增時客戶合約時,預帶客戶檔當下智權人員
            'Modify by Amy 2019/05/02 增加代理人合約 +if
            If Left(Trim(textCT(Index)), 1) = "X" Then
                If ActionEdit = 1 And textCT(6) = MsgText(601) And stCU13 <> MsgText(601) Then
                    textCT(6) = stCU13
                    Call textCT_Validate(6, False)
                End If
            End If
        Case 3, 7 '合約名稱/備註
             If bolMsg = True And Not CheckLengthIsOK(textCT(Index), textCT(Index).MaxLength) Then
                Cancel = True
                textCT(Index).SetFocus
                TextInverse textCT(Index)
                Exit Sub
             End If
        Case 4 '簽約日期
            If bolMsg = True And CheckIsTaiwanDate(textCT(Index)) = False Then
                Cancel = True
                textCT(Index).SetFocus
                TextInverse textCT(Index)
                Exit Sub
            End If
        Case 5 '機密
           
        Case 6 '智權人員
            'Modify by Amy 2019/05/02 修改時提醒離職人員但可以過,新增時不可輸離職人員
            lblStaffName = ""
            stMsg = GetStaffData(textCT(Index), stName)
            If bolMsg = True Then
                If stMsg <> MsgText(601) Then
                    '新增
                    If ActionEdit = 1 Then
                        MsgBox stMsg, vbExclamation + vbOKOnly
                        Cancel = True
                        textCT(Index).SetFocus
                        TextInverse textCT(Index)
                        Exit Sub
                    '修改
                    Else
                        MsgBox stMsg, vbExclamation + vbOKOnly
                        If InStr(stMsg, "此代號員工已離職") > 0 Then
                            textCT(Index).Tag = textCT(Index)
                        ElseIf InStr(stMsg, "此代號員工已離職") = 0 Then
                            Cancel = True
                            textCT(Index).SetFocus
                            TextInverse textCT(Index)
                            Exit Sub
                        End If
                    End If
                End If
            End If
            lblStaffName = stName
            'end 2019/05/02
    End Select
   
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    textCT(1).TabStop = True
    Call OnAction(Button.Index)
End Sub

Private Function OnAction(ByVal BtIdx) As Boolean
    Dim stLocalPath As String, bolCancel As Boolean
    
    OnAction = False
    If BtIdx = 11 Then
        If TxtValidate = False Then Exit Function
    End If
    
    If BtIdx <= 9 Then
        SetToolBarAndBT (BtIdx)
        ActionEdit = BtIdx '不記錄按下確定及取消鈕
    End If
    
    LockTxt
    Select Case BtIdx
        Case 1 'Add
            TxtClear
            textCT(2).SetFocus
        Case 2 'Update
            textCT(3).SetFocus
            textCT(1).TabStop = False
        Case 3 'Del
            If MsgBox(IIf(m_FieldList(8).fiOldData <> "", "有附件", "") & "是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                If DelRecord = True Then
                    RefreshRange
                    ActionEdit = 12
                    Call GetRecordVal("Pre")
                    SetToolBarAndBT
                    LockTxt
                    textCT(8).Tag = "": textCT(9).Tag = ""
                Else
                    Exit Function
               End If
            End If
        Case 4 'Query
            TxtClear
            textCT(1).SetFocus
            textCT(1).TabStop = True
        Case 6 'MoveFirst
            GetRecordVal ("First")
        Case 7 'MovePrv
            GetRecordVal ("Pre")
        Case 8 'MoveNext
            GetRecordVal ("Next")
        Case 9 'MoveLast
            GetRecordVal ("Last")
        Case 11 'OK
            If ActionEdit <= 2 Then
                '新增
                If ActionEdit = 1 Then
                    '取得合約編號(民國年+3碼流水號)
                    textCT(1) = GetMaxCT01
                    If Len(textCT(1)) <> 6 Then
                        textCT(1) = ""
                        ShowMsg "讀取合約自動編號錯誤，請洽系統管理者 ！"
                        Exit Function
                    End If
                End If
                If QueryRecord(textCT(1), False, False) = True Then
                    If ActionEdit = 1 Then
                        ShowMsg "合約編號「" & textCT(1) & "」已存在 ！"
                        textCT(1) = ""
                        Exit Function
                    End If
                End If
                'Added by Lydia 2021/09/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
                If PUB_ChkUniText(Me, , True, "TextBox") = False Then
                     Exit Function
                End If
                'end 2021/09/15
                If lstAtt.ListCount > 0 Then
                    '抓取檔案資料
                    Call ComposeAttList(lstAtt)
                End If
                UpdateFieldNewData
                UpdateCT09NewData
                If ActionEdit = 1 Then
                    bolCancel = AddRecord
                Else
                    bolCancel = ModRecord
                End If
                If bolCancel = False Then Exit Function
                If ActionEdit = 1 Then RefreshRange
                textCT(8).Tag = "": textCT(9).Tag = ""
            End If
            '查詢
            If ActionEdit = 4 Then
                If QueryRecord(textCT(1), True, True) = False Then
                    Call QueryRecord(m_CurrKEY, True, True)
                End If
                ActionEdit = 12
            Else
                ActionEdit = 12 '先設,離職才不會彈訊息
                Call QueryRecord(textCT(1), False, True)
            End If
            SetToolBarAndBT
            LockTxt
        Case 12 'Cancel
            If ActionEdit <= 2 Then
                If MsgBox("未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
                    Exit Function
                End If
            End If
            ActionEdit = 12
            Call QueryRecord(m_CurrKEY, False, True)
            SetToolBarAndBT
            LockTxt
            textCT(1).SetFocus
        Case 14 'Exit
            Unload Me
    End Select
    OnAction = True
End Function

Private Sub SetToolBarAndBT(Optional ByVal BtIdx As Integer = 0)
    If BtIdx = 0 Then BtIdx = ActionEdit
    Select Case BtIdx
        Case 1 To 4 '新增/修改/刪除/查詢
            For i = 1 To 4
                tlbar.Buttons(i).Enabled = False
                tlbar.Buttons(i + 5).Enabled = False
            Next
            tlbar.Buttons(11).Enabled = True
            tlbar.Buttons(12).Enabled = True
            tlbar.Buttons(14).Enabled = False
            If BtIdx <= 2 Then
                cmdOpenAtt.Enabled = False
                cmdAddAtt.Enabled = True
                cmdRemAtt.Enabled = True
            End If
        Case 6 To 9, 11 To 12  '第一、上、下、最後一筆/確定/取消
            For i = 1 To 4
                tlbar.Buttons(i).Enabled = True
                tlbar.Buttons(i + 5).Enabled = True
            Next
            tlbar.Buttons(11).Enabled = False
            tlbar.Buttons(12).Enabled = False
            tlbar.Buttons(14).Enabled = True
            cmdOpenAtt.Enabled = True
            cmdAddAtt.Enabled = False
            cmdRemAtt.Enabled = False
    End Select
End Sub

Private Sub LockTxt()
    Dim idx As Integer
    
    For Each oText In textCT
        idx = oText.Index
        '新增/修改
        If ActionEdit <= 2 Then
            If idx = 1 Then
                oText.Enabled = False
                oText.Locked = True
            Else
                oText.Enabled = True
                oText.Locked = False
            End If
        '查詢
        ElseIf ActionEdit = 4 Then
            If idx = 1 Then
                oText.Enabled = True
                oText.Locked = False
            Else
                oText.Enabled = False
                oText.Locked = True
            End If
        '第一/上一/下一/最後一筆/確定/取消
        ElseIf ActionEdit >= 6 Then
            If idx = 1 Then
                oText.Enabled = True
                oText.Locked = False
            Else
                oText.Enabled = True
                oText.Locked = True
            End If
        End If
    Next
End Sub

Private Sub TxtClear()
    For i = 1 To Max_CT
        m_FieldList(i).fiOldData = Empty
        m_FieldList(i).fiNewData = Empty
    Next
    For Each oText In textCT
        oText.Text = ""
    Next
    'Add by Amy 2019/05/02
    textCT(5).Tag = ""
    If ActionEdit = 1 Then
        Label23.Caption = "Create ID: 　　　　　　　　　　　　　Update ID:"
    End If
    'end 2019/05/02
    lstAtt.Clear
    lblStaffName = ""
    lblCusName = ""
End Sub

Private Function DelRecord() As Boolean
    Dim strDel As String, m_CT01 As String, stErrMsg As String
    Dim iErr As Integer
On Error GoTo ErrHand

    DelRecord = False
    m_CT01 = textCT(1)
    
    strDel = "Delete From Contract Where CT01='" & m_CT01 & "' "
    cnnConnection.BeginTrans
    
    '先刪附件再刪DB資料
    If m_FieldList(9).fiOldData <> "" Then
        If UploadAttFile(iErr, stErrMsg) = False Then GoTo ErrHand
    End If
    cnnConnection.Execute strDel
    cnnConnection.CommitTrans
    
    DelRecord = True
    ' 只有刪除的是第一或最後一筆才須重新第一筆及最後一筆
    If (m_CT01 = m_LastKEY) Or (m_CT01 = m_FirstKEY) Then
        RefreshRange
    End If
    Call GetCurrRecordVal(m_CT01)
    Exit Function
    
ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical
    ElseIf iErr <> 0 Then
       MsgBox stErrMsg, vbCritical
    End If
End Function

Private Function AddRecord() As Boolean
    Dim stSQL As String, stCols As String, stValues As String
    Dim idx As Integer, iErr As Integer, stErrMsg As String

On Error GoTo ErrHand

    '畫面有的欄位才更新
    stCols = "": stValues = ""
    For Each oText In textCT
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
    stSQL = "INSERT INTO Contract (" & stCols & ") Values (" & stValues & ")"
    
    cnnConnection.BeginTrans
    
    If m_FieldList(9).fiNewData <> MsgText(601) Then
        '上傳附件檔
        If UploadAttFile(iErr, stErrMsg) = False Then GoTo ErrHand
    End If
    
    Pub_SeekTbLog stSQL
    cnnConnection.Execute stSQL
    
    cnnConnection.CommitTrans
    AddRecord = True
    Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical
    ElseIf iErr <> 0 Then
       MsgBox stErrMsg, vbCritical
    End If
End Function

Private Function ModRecord() As Boolean
    Dim strMod As String, stColVal As String, stErrMsg As String
    Dim idx As Integer, iErr As Integer, bDifference As Boolean
    
On Error GoTo ErrHand
    
    strMod = "begin user_data.user_enabled:=1; UPDATE Contract SET "
    
    For Each oText In textCT
        idx = oText.Index
        If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
            bDifference = True
            '文字
            If m_FieldList(idx).fiType = 0 Then
                stColVal = stColVal & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
            '數字
            Else
                stColVal = stColVal & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
            End If
        End If
    Next
    If bDifference = True Then
        stColVal = Mid(stColVal, 2)
        strMod = strMod & stColVal & " Where CT01='" & textCT(1) & "'; end; "
        
        cnnConnection.BeginTrans
        
        If m_FieldList(9).fiNewData <> m_FieldList(9).fiOldData Then
            '上傳附件檔
            If UploadAttFile(iErr, stErrMsg) = False Then GoTo ErrHand
        End If
        Pub_SeekTbLog strMod
        cnnConnection.Execute strMod, intI
        
        cnnConnection.CommitTrans
   End If
    
    ModRecord = True
    Exit Function
    
ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical
    ElseIf iErr <> 0 Then
       MsgBox stErrMsg, vbCritical
    End If
End Function

Private Sub UpdateCT09NewData()
    Dim stLocalPath As Variant
    Dim stFileName As String, stNewPath As String, NowTime As String
    Dim stTmp As String, stTmpFN As String
    
    If textCT(8).Tag = MsgText(601) Then Exit Sub
    
    '有檔案
    stLocalPath = Split(textCT(8).Tag, ",")
    NowTime = ServerTime
    For i = LBound(stLocalPath) To UBound(stLocalPath)
        stFileName = ""
        If InStr(stLocalPath(i), "\") > 0 Then
            If Len(NowTime) = 5 Then NowTime = "0" & NowTime
            '未上傳檔案
            stFileName = Left(textCT(1), 3) & "/" & textCT(1) & "/" & _
                            textCT(1) & "_" & strSrvDate(1) & "." & NowTime & Mid(stLocalPath(i), InStr(stLocalPath(i), "."))
            stTmpFN = textCT(1) & "_" & strSrvDate(1) & "." & NowTime & Mid(stLocalPath(i), InStr(stLocalPath(i), "."))
            NowTime = Val(NowTime) + 1
        Else
            '抓舊CT09位置
            stFileName = Mid(stOldCT0809, InStr(stOldCT0809, stLocalPath(i)))
            If InStr(stFileName, ",") > 0 Then
                stFileName = Mid(stFileName, 1, InStr(stFileName, ",") - 1)
            End If
            stFileName = Mid(stFileName, InStr(stFileName, "##") + 2)
            stTmpFN = stLocalPath(i)
        End If
        stTmp = stTmp & "," & stFileName
        '記錄尚未上傳檔案路徑##FTP檔名**CT09
        stNewPath = stNewPath & "," & stLocalPath(i) & "##" & stTmpFN & "**" & stFileName
    Next i
    textCT(9) = Mid(stTmp, 2)
    m_FieldList(9).fiNewData = textCT(9)
    textCT(9).Tag = Mid(stNewPath, 2)
End Sub

Private Function TxtValidate() As Boolean
    Dim bolCancel As Boolean
    Dim stName As String, stDept As String, stMsg As String 'Add by Amy 2019/05/02
    
    TxtValidate = False: bolCancel = False
    
    '查詢
    If ActionEdit = 4 Then
        If textCT(1) = MsgText(601) Then
            MsgBox "請輸入查詢條件", , MsgText(5)
            Exit Function
        End If
        Call textCT_Validate(1, bolCancel)
        If bolCancel = True Then Exit Function
        TxtValidate = True
        Exit Function
    ElseIf ActionEdit = 3 Then
        If textCT(1) = MsgText(601) Then
            MsgBox "無資料可刪除！", , MsgText(5)
            ActionEdit = 12
            Call QueryRecord(m_CurrKEY, False, False)
            SetToolBarAndBT
            Exit Function
        End If
        TxtValidate = True
        Exit Function
    '新增/修改
    Else
        If ActionEdit = 2 And textCT(1) = MsgText(601) Then
            MsgBox "無資料可修改！", , MsgText(5)
            ActionEdit = 12
            Call QueryRecord(m_CurrKEY, False, False)
            SetToolBarAndBT
            Exit Function
        End If
        If textCT(2) = MsgText(601) Then
            'Modify by Amy 2019/05/02 增加代理人合約
            MsgBox "客戶/代理人編號不可為空", , MsgText(5)
            Exit Function
        End If
        If textCT(3) = MsgText(601) Then
            MsgBox "合約名稱不可為空", , MsgText(5)
            Exit Function
        End If
        If textCT(4) = MsgText(601) Then
            MsgBox "簽約日期不可為空", , MsgText(5)
            Exit Function
        End If
        If textCT(6) = MsgText(601) Then
            MsgBox "智權人員不可為空", , MsgText(5)
            Exit Function
        End If
    End If
    
    For i = 2 To 4
        Call textCT_Validate(i, bolCancel)
        If bolCancel = True Then Exit For
    Next i
    If bolCancel = True Then Exit Function
    
    'Add by Amy 2019/05/02 智權人員檢查(因修改時Validate判斷離職人員離職會彈,此不需再彈)
    stMsg = GetStaffData(textCT(6), stName, stDept)
    If stMsg <> MsgText(601) Then
        '新增
        If ActionEdit = 1 Then
            MsgBox stMsg, vbExclamation + vbOKOnly
            textCT(6).SetFocus
            TextInverse textCT(6)
            Exit Function
        '修改
        ElseIf (InStr(stMsg, "此代號員工已離職") > 0 And textCT(6).Tag <> textCT(6)) Or InStr(stMsg, "此代號員工已離職") = 0 Then
            MsgBox stMsg, vbExclamation + vbOKOnly
            If InStr(stMsg, "此代號員工已離職") = 0 Then
                textCT(6).SetFocus
                TextInverse textCT(6)
                Exit Function
            End If
        End If
    End If
    
    '機密
    stMsg = ""
    If textCT(5) <> MsgText(601) And textCT(5) & textCT(6) <> textCT(5).Tag Then
        '智權人員為國內人員,若設國外機密彈訊息
        If Left(Trim(textCT(2)), 1) = "X" Then
            If (Left(stDept, 1) <> "F" And textCT(5) = "F") Or (Left(stDept, 1) = "F" And textCT(5) = "C") Then
                stMsg = "智權人員為國內人員,機密確定設為「國外機密」嗎?"
                If Left(stDept, 1) = "F" And textCT(5) = "C" Then
                    stMsg = "智權人員為國外人員,機密確定設為「國內機密」嗎?"
                End If
                If MsgBox(stMsg, vbYesNo) = vbNo Then
                    textCT(5).Tag = ""
                    textCT(5).SetFocus
                    TextInverse textCT(5)
                    Exit Function
                End If
            End If
        '若為代理人合約,若設國內機密彈訊息
        ElseIf Left(Trim(textCT(2)), 1) = "Y" And textCT(5) = "C" Then
            If MsgBox("此為代理人合約,機密確定設為「國內機密」嗎?", vbYesNo) = vbNo Then
                textCT(5).Tag = ""
                textCT(5).SetFocus
                TextInverse textCT(5)
                Exit Function
            End If
        End If
        textCT(5).Tag = textCT(5) & textCT(6) '記錄設定
    End If
    'end 2019/05/02
    TxtValidate = True
End Function

' 檢查記錄是否已經存在
Private Function QueryRecord(ByVal strKey As String, ByVal bolShowMsg As Boolean, Optional ByVal bolSetDt As Boolean = False) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strField As String
   
    QueryRecord = False
    If bolSetDt = True And ActionEdit >= 4 Then Call TxtClear
    
    strQ = "Select * From Contract Where CT01=" & Val(strKey)
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    ' 檢查讀取的資料筆數
    If RsQ.RecordCount = 0 Then
        QueryRecord = False
        If bolShowMsg = True Then
            MsgBox "無此資料！", , MsgText(5)
        End If
    Else
        QueryRecord = True
        If bolSetDt = True Then
            m_CurrKEY = strKey
            SetTxtValue RsQ
        End If
    End If
    RsQ.Close
    Set RsQ = Nothing
End Function

Private Function GetMaxCT01() As String
    Dim rsCT As New ADODB.Recordset
    Dim intCT As Integer
    
    strQ = "Select  Nvl(Max(CT01),0) as CT01 From Contract "
    intCT = 1
    Set rsCT = ClsLawReadRstMsg(intCT, strQ)
    If intCT = 1 Then
        If Val(rsCT.Fields("CT01")) = 0 Then
            GetMaxCT01 = Left(strSrvDate(2), 3) & ZeroBeforeNo(0, 3)
        Else
            GetMaxCT01 = Left(rsCT.Fields("CT01"), 3) & ZeroBeforeNo(Right(Val(rsCT.Fields("CT01")), 3), 3)
        End If
    End If
    rsCT.Close
    Set rsCT = Nothing
End Function

Private Function AddList(oList As ListBox, stNewItem As String, stSaveName As String) As Boolean
    Dim idx As Integer, bFound As Boolean, stFileName As String
    
    If InStr(stNewItem, ",") > 0 Then
        MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
        cmdAddAtt.SetFocus
        Exit Function
    End If
   
    If stNewItem <> "" Then
        For idx = 0 To oList.ListCount - 1
            stFileName = GetFileName(oList.List(idx))
            If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
                MsgBox "附件[" & stFileName & "]已存在！"
                AddList = False
                bFound = True
                Exit For
            End If
        Next
        
        If bFound = False Then
            oList.AddItem stNewItem, 0
            SetListScroll oList
            AddList = True
        End If
    End If
End Function

'取得原始附件檔名
Private Sub ComposeAttList(oList As ListBox)
    Dim stItem As String, stFileN As String, stRtn1 As String, stRtn2 As String
    Dim idx As Integer, iPos As Integer

    If oList.ListCount > 0 Then
        For idx = 0 To oList.ListCount - 1
            stItem = oList.List(idx)
            stRtn1 = stRtn1 & "," & GetFileName(stItem)
            stFileN = stItem
            If InStrRev(stFileN, " (") > 0 Then
                stFileN = Trim(Mid(stFileN, 1, InStrRev(stFileN, "(") - 1))
            End If
            stRtn2 = stRtn2 & "," & stFileN
        Next
    End If
    textCT(8) = Mid(stRtn1, 2)
    textCT(8).Tag = Mid(stRtn2, 2) '記錄lstAtt實體路徑
End Sub

'刪除lstAtt 點選項目
Private Function RemoveList(oList As ListBox) As Boolean
    Dim ii As Integer
    If oList.ListCount > 0 Then
        ii = 0
        Do While ii < oList.ListCount
            If oList.Selected(ii) = True Then
                oList.RemoveItem ii
                SetListScroll oList
                RemoveList = True
                ii = ii - 1
            End If
            ii = ii + 1
        Loop
    End If
End Function

Private Sub SetList(oList As ListBox, p_stList As String)
    Dim arrID
    oList.Clear
    If p_stList = MsgText(601) Then Exit Sub
    
    If p_stList <> "" Then
        arrID = Split(p_stList, ",")
        For intI = UBound(arrID) To LBound(arrID) Step -1
            oList.AddItem arrID(intI), 0
        Next
    End If
End Sub

Private Sub SetListScroll(oList As ListBox)
    Dim ii As Integer
    Dim lWnow As Long, lWmax As Long
   
    lWmax = 0
    For ii = 0 To oList.ListCount - 1
        lWnow = TextWidth(oList.List(ii) & " ")
        If lWnow > lWmax Then
            lWmax = lWnow
        End If
    Next
  
    If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
    SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function UploadAttFile(ByRef iErrNo As Integer, ByRef stErrMsg As String) As Boolean
    Dim arrNewPath As Variant, arrOldTmp As Variant
    Dim ii As Integer
    Dim strMid As String
    Dim stFilePath As String, stFileName As String, stLocalPath As String

On Error GoTo OutPort
   
    iErrNo = 0: stErrMsg = ""

    arrOldTmp = Empty
    arrNewPath = Split(textCT(9).Tag, ",") '上傳用
    arrOldTmp = Split(stOldCT0809, ",") '刪檔用
    
    If m_FieldList(9).fiOldData <> Empty Then
        For ii = LBound(arrOldTmp) To UBound(arrOldTmp)
            stFilePath = Mid(arrOldTmp(ii), InStr(arrOldTmp(ii), "##") + 2)
            If InStr(m_FieldList(9).fiNewData, stFilePath) = 0 Or ActionEdit = 3 Then
                '檔案放於 FTP,必須在DB資料刪除前執行刪除附件
                If PUB_DelFtpFile2(textCT(1), stFilePath, stFolderN) = False Then GoTo OutPort
            End If
        Next ii
    End If
    
    '上傳檔案
    For ii = LBound(arrNewPath) To UBound(arrNewPath)
        stLocalPath = Mid(arrNewPath(ii), 1, InStr(arrNewPath(ii), "##") - 1) '實體路徑
        strMid = Replace(arrNewPath(ii), stLocalPath & "##", "")
        stFileName = Mid(strMid, 1, InStr(strMid, "**") - 1) 'FTP檔名
        stFilePath = Replace(arrNewPath(ii), stLocalPath & "##" & stFileName & "**", "") 'FTP路徑
        If InStr(m_FieldList(9).fiOldData, stFilePath) = 0 Then
            If PUB_PutFtpFile(stLocalPath, textCT(1), stFileName, , stFolderN) = False Then GoTo OutPort
        End If
    Next ii
   
    UploadAttFile = True
   
    Exit Function
   
OutPort:
   iErrNo = Err.Number
   stErrMsg = Err.Description
   
End Function

'傳入員工編號,回傳名稱/部門
Private Function GetStaffData(ByVal stST01 As String, Optional ByRef stName As String, Optional ByRef stDept As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String
    
    GetStaffData = ""
    stName = "": stDept = ""
    
    stQ = "Select * From Staff,Acc090 Where ST03=A0901(+) And ST01='" & stST01 & "' "
    RsQ.CursorLocation = adUseClient
    RsQ.Open stQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        If RsQ.Fields("ST04") <> "1" Then
            GetStaffData = "此代號員工已離職！"
        End If
        stName = "" & RsQ.Fields("ST02").Value
        stDept = "" & RsQ.Fields("A0901").Value
    Else
        GetStaffData = "此代號不存在於員工檔！"
    End If

    RsQ.Close
End Function

