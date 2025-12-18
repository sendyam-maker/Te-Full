VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090226 
   BorderStyle     =   1  '單線固定
   Caption         =   "商申承辦人責任業務區分配維護"
   ClientHeight    =   6420
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7524
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7524
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00800000&
      Height          =   2040
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frm090226.frx":0000
      Top             =   4380
      Width           =   7395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "智權"
      Height          =   180
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   1260
      Value           =   -1  'True
      Width           =   800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "客戶"
      Height          =   180
      Index           =   1
      Left            =   3705
      TabIndex        =   6
      Top             =   1260
      Width           =   800
   End
   Begin VB.TextBox txtDB2 
      Height          =   270
      Left            =   3750
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "X1234567"
      Top             =   930
      Width           =   900
   End
   Begin VB.TextBox txtDB1 
      Height          =   270
      Left            =   870
      MaxLength       =   5
      TabIndex        =   1
      Top             =   930
      Width           =   700
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090226.frx":01D8
      Height          =   2805
      Left            =   120
      TabIndex        =   3
      Top             =   1530
      Width           =   7275
      _ExtentX        =   12827
      _ExtentY        =   4953
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   2
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
      _Band(0).Cols   =   7
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   150
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
            Picture         =   "frm090226.frx":01ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":0509
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":0825
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":0A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":0D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":1039
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":1355
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":1671
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":198D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":1CA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090226.frx":1FC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   660
      Width           =   3165
      VariousPropertyBits=   27
      Caption         =   "CREATE：ID  Date  Time"
      Size            =   "5583;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   12
      Top             =   660
      Width           =   3165
      VariousPropertyBits=   27
      Caption         =   "UPDATE：ID  Date  Time"
      Size            =   "5583;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門："
      Height          =   180
      Index           =   2
      Left            =   4710
      TabIndex        =   11
      Top             =   1260
      Width           =   540
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   10
      Top             =   1260
      Width           =   1725
      VariousPropertyBits=   27
      Caption         =   "label3(0)"
      Size            =   "3043;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權/客戶："
      Height          =   180
      Index           =   1
      Left            =   2790
      TabIndex        =   9
      Top             =   975
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   975
      Width           =   720
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   7
      Top             =   960
      Width           =   975
      VariousPropertyBits=   27
      Caption         =   "label3(1)"
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   4695
      TabIndex        =   4
      Top             =   930
      Width           =   2745
      VariousPropertyBits=   27
      Caption         =   "label3(2)"
      Size            =   "4851;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm090226"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Amy 2022/09/21
Option Explicit
Dim i As Integer, m_blnColOrderAsc As Boolean, dblPrevRow As Double, bolQuery As Boolean
Dim m_EditMode As Integer '0:瀏覽1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean, m_bQuery As Boolean
Dim strAllField As String, strWField As String, m_DZA01 As String, m_DZA02 As String, arrF, arrW
Dim m_FirstKEY(2) As String, m_LastKEY(2) As String, m_CurrKEY(2) As String ' 第一/最後/目前 資料
Dim arrDZAOld() As String 'Add by Amy 2023/05/31

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        ' 新增
        Case vbKeyF2
            If m_bInsert Then
                If m_EditMode = 0 Then
                    OnAction KeyCode
                    KeyCode = 0
                End If
            End If
        ' 修改
        Case vbKeyF3
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
            If m_EditMode = 0 Then
                OnAction KeyCode
            Else
                OnAction vbKeyF10
            End If
    End Select
End Sub

Private Sub Form_Load()
    '取得使用者執行各項功能的權限
    m_bInsert = IsUserHasRightOfFunction("frm090226", strAdd, False)
    m_bUpdate = IsUserHasRightOfFunction("frm090226", strEdit, False)
    m_bDelete = IsUserHasRightOfFunction("frm090226", strDel, False)
    m_bQuery = IsUserHasRightOfFunction("frm090226", strFind, False)
    
    MoveFormToCenter Me
    SetarrDZAOld 'Add by Amy 2023/05/31
    FormClear
    strAllField = "承辦人|智權/客戶編號|名稱|部門|新增人員時間|更新人員時間|DZA01|DZA02|DeptNo|Sort"
    arrW = Split("800,1200,2300,800,1500,1500,0,0,0,0", ",")
    arrF = Split(strAllField, "|")
    
    RefreshRange
    QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090226 = Nothing
End Sub

'Add by Amy 2023/05/31
Private Sub SetarrDZAOld()
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "Select * From DutyZoneAssign Where RowNum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   ReDim arrDZAOld(AdoRecordSet3.Fields.Count)
End Sub

Private Sub OnAction(ByVal KeyCode As Integer)
    Dim strTit As String, strMsg As String, nResponse
    
    Select Case KeyCode
        '新增
        Case vbKeyF2:
            m_EditMode = 1
            FormClear
            'SetLock 'Mark by Amy 2023/05/31  不需使用
            UpdateToolbarState
            txtDB1.SetFocus
        '修改
        Case vbKeyF3:
            m_EditMode = 2
            UpdateCtrlData
            'SetLock'Mark by Amy 2023/05/31  不需使用
            UpdateToolbarState
            txtDB2.SetFocus
        '刪除
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
        '查詢
        Case vbKeyF4:
            m_EditMode = 4
            'SetLock'Mark by Amy 2023/05/31  不需使用
            FormClear
            UpdateToolbarState
            txtDB1.SetFocus
        '第一筆
        Case vbKeyHome:
            If bolQuery = True Then
                FormClear
                QueryData
                bolQuery = False
            End If
            ShowFirstRecord
        '前一筆
        Case vbKeyPageUp:
            If bolQuery = True Then
                FormClear
                QueryData
                bolQuery = False
            End If
            ShowPrevRecord
        '後一筆
        Case vbKeyPageDown:
            If bolQuery = True Then
                FormClear
                QueryData
                bolQuery = False
            End If
            ShowNextRecord
        '最後一筆
        Case vbKeyEnd:
            If bolQuery = True Then
                FormClear
                QueryData
                bolQuery = False
            End If
            ShowLastRecord
        '確定
        Case vbKeyF9:
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
        '取消
        Case vbKeyF10:
            If m_EditMode = 1 Or m_EditMode = 2 Then
                strTit = "詢問"
                strMsg = "你並未存檔, 確定離開嗎?"
                nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
                If nResponse = vbYes Then
                    m_EditMode = 0
                    UpdateCtrlData
                    'SetLock 'Mark by Amy 2023/05/31  不需使用
                    UpdateToolbarState
               End If
            Else
                m_EditMode = 0
                UpdateCtrlData
                'SetLock 'Mark by Amy 2023/05/31  不需使用
                UpdateToolbarState
            End If
        '離開
        Case vbKeyEscape:
            Unload Me
    End Select
    
End Sub

'使用者按下確定的按紐
Private Function OnWork() As Boolean
    Dim strMsg As String, strTit As String, nResponse
    
    OnWork = False
    Select Case m_EditMode
        Case 1, 2: '新增/修改
            If FormCheck() = False Then
                Exit Function
            End If
            'Modify by Amy 2023/05/31 +if
            If m_EditMode = 1 Or (m_EditMode = 2 And arrDZAOld(0) & arrDZAOld(1) <> txtDB1 & txtDB2) Then
               If FormSave = True Then
                   FormClear
                   RefreshRange
                   QueryData
                   bolQuery = False
               Else
                   Exit Function
               End If
            End If
        Case 3: '刪除
            If FormCheck() = False Then
                Exit Function
            End If
            If DelRecord = True Then
                FormClear
                RefreshRange
                QueryData
                bolQuery = False
            Else
                Exit Function
            End If
        Case 4: '查詢
            If Trim(txtDB1 & txtDB2) = MsgText(601) Then
                MsgBox "請輸入查詢條件", vbExclamation + vbOKOnly
                Exit Function
            End If
            bolQuery = True
            dblPrevRow = 0
            If QueryData = False Then
                MsgBox "查無資料", vbExclamation + vbOKOnly
                Exit Function
            End If
    End Select
    m_EditMode = 0
    OnWork = True
End Function

Private Function FormSave() As Boolean
    Dim strCmd As String, stNowTime As String, intExe As Integer
    Dim bolTrans As Boolean 'Add by Ａｍy 2023/05/31
On Error GoTo ErrHand

    FormSave = False
    stNowTime = ServerTime
    
    '新增
    If m_EditMode = 1 Then
        strCmd = "Insert Into DutyZoneAssign (DZA01,DZA02,DZA03,DZA04,DZA05) Values" & _
                        "('" & txtDB1 & "','" & txtDB2 & "','" & strUserNum & "'," & strSrvDate(1) & "," & stNowTime & ")"
    Else
        'Modify by Amy 2023/05/31 開放承辦人也可修改,故改為先刪除後再新增[建立人員/日期/時間維持舊資料]
        'strCmd = "Update DutyZoneAssign Set DZA02='" & txtDB2 & "',DZA06='" & strUserNum & "',DZA07=" & strSrvDate(1) & ",DZA08=" & stNowTime & " " & _
                        "Where DZA01='" & txtDB1 & "' And DZA02='" & txtDB2.Tag & "' "
        cnnConnection.BeginTrans
        bolTrans = True
        strCmd = "Delete From DutyZoneAssign Where DZA01='" & arrDZAOld(0) & "' And DZA02='" & arrDZAOld(1) & "' "
        Pub_SeekTbLog strCmd
        cnnConnection.Execute strCmd, intExe
        strCmd = "Insert Into DutyZoneAssign (DZA01,DZA02,DZA03,DZA04,DZA05,DZA06,DZA07,DZA08) VALUES" & _
                        "('" & txtDB1 & "','" & txtDB2 & "','" & arrDZAOld(2) & "'," & arrDZAOld(3) & "," & arrDZAOld(4) & _
                        ",'" & strUserNum & "'," & strSrvDate(1) & "," & stNowTime & ")"
        Pub_SeekTbLog strCmd
    End If
    cnnConnection.Execute strCmd, intExe
    If bolTrans = True Then cnnConnection.CommitTrans 'Add by Amy 2023/05/31
    If intExe > 0 Then
        FormSave = True
    Else
        MsgBox "未有資料" & IIf(m_EditMode = 1, "新增", "修改") & "！請洽電腦中心"
    End If
    Exit Function
    
ErrHand:
    If bolTrans = True Then cnnConnection.RollbackTrans 'Add by Amy 2023/05/31
    MsgBox IIf(m_EditMode = 1, "新增", "修改") & "失敗！" & vbCrLf & Err.Description
End Function

Private Function DelRecord() As Boolean
    Dim strCmd As String, strDZA01  As String, strDZA02 As String
On Error GoTo ErrHand

    DelRecord = False
    strDZA01 = m_CurrKEY(0)
    strDZA02 = m_CurrKEY(1)
   
    strCmd = "Delete From DutyZoneAssign Where DZA01='" & strDZA01 & "' And DZA02='" & strDZA02 & "' "
    Pub_SeekTbLog strCmd
    cnnConnection.Execute strCmd
    
    '只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
    If (strDZA01 = m_LastKEY(0) And strDZA02 = m_LastKEY(1)) Or (strDZA01 = m_FirstKEY(0) And strDZA02 = m_FirstKEY(1)) Then
        RefreshRange
    End If
    ShowCurrRecord strDZA01, strDZA02
    DelRecord = True
    Exit Function
    
ErrHand:
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

'檢查記錄是否已經存在
'intChoose:0-查詢/1-檢查
Private Function IsRecordExist(ByVal intChoose As Integer, ByVal strKEY01 As String, ByVal strKEY02 As String, Optional ByRef stMsg As String) As Boolean
    Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
    Dim strWhere As String, strTmp As String 'Add by Amy 2023/05/31
    
    IsRecordExist = False: stMsg = ""
    
    '查詢
    If intChoose = 0 Then
        strQ = "Select DZA01,DZA02,'1' as State From DutyZoneAssign Where DZA01 = '" & strKEY01 & "' And DZA02 = '" & strKEY02 & "' "
    '檢查
    Else
        'Modify by Amy 2023/05/31 +strWhere開放承辦人也可修改,原2句 Union,ex:已有X84335 再建X8433501 不應該再彈訊息
        If m_EditMode = 2 Then
            strWhere = "And DZA01||DZA02<>'" & arrDZAOld(0) & arrDZAOld(1) & "' "
        End If
        'Modify by Amy 2024/01/11 智權人員要可輸MCTF0X
        If Left(strKEY02, 1) = 客戶編號 Then
            stMsg = "客戶 [" & strKEY02 & "] " & Len(strKEY02) & "碼編號 已設於"
            If Len(strKEY02) <> 6 Then
               strQ = "Select DZA01,DZA02,'1' as State From DutyZoneAssign Where SubStr(DZA02,1,8) = '" & Left(ChangeCustomerL(strKEY02), 8) & "' And Length(DZA02)=8 " & _
                            "And SubStr(DZA02,1,1)='" & 客戶編號 & "' " & strWhere
            Else
               strQ = "Select DZA01,DZA02,'2' as State From DutyZoneAssign Where SubStr(DZA02,1,6) = '" & Left(strKEY02, 6) & "' And Length(DZA02)=6 " & _
                            "And SubStr(DZA02,1,1)='" & 客戶編號 & "' " & strWhere
            End If
        Else
            stMsg = "智權人員 " & GetPrjSalesNM(strKEY02) & "(" & strKEY02 & ") 已設於"
            strQ = strQ & "Select DZA01,DZA02,'1' as State From DutyZoneAssign Where DZA02 = '" & strKEY02 & "' " & _
                        "And SubStr(DZA02,1,1)<>'" & 客戶編號 & "' " & strWhere
        End If
        'end 2024/01/11
        stMsg = stMsg & vbCrLf
        'end 2023/05/31
    End If
                  
    '讀取資料庫
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        If intChoose > 0 Then
            Do While Not RsQ.EOF
                'Modify by Amy 2023/05/31 開放承辦人也可修改
'                If "" & RsQ.Fields("State") = "1" And RsQ.Fields("DZA01") <> txtDB1 Then
'                    '新增/修改時,智權或客戶已存在且不為同一承辦人才可新增or修改
'                    stMsg = stMsg & "," & RsQ.Fields("DZA01") & " " & GetPrjSalesNM("" & RsQ.Fields("DZA01"))
'                ElseIf RsQ.Fields("DZA02") = txtDB2 And RsQ.Fields("DZA01") = txtDB1 Then
'                    '新增/修改時,智權/客戶已存在同一承辦人不可再新增/修改
'                    If m_EditMode = 1 Then
'                        stMsg = ",同一承辦人不可再新增同客戶6碼編號！"
'                    Else
'                        stMsg = ",同一承辦人已有設定此編號，不可修改！"
'                    End If
'                    Exit Do
'                Else
'                    stMsg = stMsg & "," & RsQ.Fields("DZA01") & " " & GetPrjSalesNM("" & RsQ.Fields("DZA01"))
'                    If "" & RsQ.Fields("State") = "2" Then
'                        stMsg = stMsg & " 設於(" & RsQ.Fields("DZA02") & ") 6碼" & vbCrLf
'                    Else
'                        stMsg = stMsg & vbCrLf
'                    End If
'                End If
                strTmp = strTmp & "," & GetPrjSalesNM("" & RsQ.Fields("DZA01")) & "(" & RsQ.Fields("DZA01") & ")"
                'end 2023/05/31
                RsQ.MoveNext
            Loop
            'Modify by Amy 2023/05/31 調整訊息
            If strTmp <> MsgText(601) Then
                stMsg = stMsg & Mid(strTmp, 2)
                If m_EditMode = 1 Then
                  stMsg = stMsg & ",不可再新增！"
                Else
                  stMsg = stMsg & ",不可修改！"
                End If
            End If
        End If
        IsRecordExist = True
    Else
        If intChoose > 0 Then
            stMsg = "此編號不存在！"
        End If
        IsRecordExist = False
    End If
    Set RsQ = Nothing
End Function

Private Sub RefreshRange()
    Dim rsTmp As New ADODB.Recordset, strSql As String, intA As Integer
    
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = (Select Min(DZA01) From DutyZoneAssign ) " & _
                   "And DZA02 =(Select Min(DZA02) From DutyZoneAssign Where DZA01= (Select Min(DZA01) From DutyZoneAssign ) ) "
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("DZA02")
    End If
    rsTmp.Close
    
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = (Select Max(DZA01) From DutyZoneAssign ) " & _
                   "And DZA02 =(Select Max(DZA02) From DutyZoneAssign Where DZA01= (Select Max(DZA01) From DutyZoneAssign ) ) "
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_LastKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_LastKEY(1) = rsTmp.Fields("DZA02")
    End If
    rsTmp.Close
    
    Set rsTmp = Nothing
End Sub

Private Function QueryData() As Boolean
    Dim rsTmp As New ADODB.Recordset, strSql As String, strWhere1 As String, strWhere2 As String, intA As Integer
    
    QueryData = False
    If txtDB1 <> MsgText(601) Then
        strWhere1 = " And DZA01='" & txtDB1 & "' "
        strWhere2 = " And DZA01='" & txtDB1 & "' "
    End If
    If txtDB2 <> MsgText(601) Then
        strWhere1 = strWhere1 & " And SubStr(DZA02,1," & Len(txtDB2) & ")='" & txtDB2 & "' "
        strWhere2 = strWhere2 & " And DZA02='" & txtDB2 & "' "
    End If
    
    'Modify by Amy 2024/01/11 智權人員要可輸MCTF0X
    '客戶8碼
    strSql = "Select W.st02 as WP,DZA02 as ANo,Decode(cu04,null,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as SName,'' as Dept" & _
                ",C.st02||' '||SqlDatet(Dza04)||' '||SqlTime(Dza05) AddDT,M.st02||' '||SqlDatet(Dza07)||' '||SqlTime(Dza08) ModDT" & _
                ",DZA01,DZA02,'' as DeptNo,1 as Sort " & _
                "From DutyZoneAssign,Customer,Staff W,Staff C,Staff M " & _
                "Where DZA01=W.ST01(+) " & strWhere1 & " And Dza03=C.ST01(+) And Dza06=M.ST01(+) " & _
                "And DZA02=CU01(+) And CU02='0' And Length(DZA02)=8 And SubStr(DZA02,1,1)='" & 客戶編號 & "' And CU01 is not null "
    '客戶6碼(顯示8碼00的名稱)
    strSql = strSql & " Union " & _
                "Select WP,ANo,Decode(cu04,null,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as SName,'' as Dept,AddDT,ModDT,DZA01,DZA02,'' as DeptNo,2 as Sort " & _
                "From (Select W.st02 as WP,DZA02 as ANo,DZA01,DZA02" & _
                            ",C.st02||' '||SqlDatet(Dza04)||' '||SqlTime(Dza05) AddDT,M.st02||' '||SqlDatet(Dza07)||' '||SqlTime(Dza08) ModDT " & _
                            "From DutyZoneAssign,Staff W,Staff C,Staff M " & _
                            "Where DZA01=W.ST01(+) " & strWhere1 & " And Dza03=C.ST01(+) And Dza06=M.ST01(+) And Length(DZA02)=6 And SubStr(DZA02,1,1)='" & 客戶編號 & "' " & _
                "),Customer Where DZA02||'00'=CU01(+) And CU02='0' And CU01 is not null "
    '智權
    strSql = strSql & " Union " & _
                "Select W.st02 as WP,DZA02 as ANo,S.st02 as SName,a0902 as Dept" & _
                ",C.st02||' '||SqlDatet(Dza04)||' '||SqlTime(Dza05) AddDT,M.st02||' '||SqlDatet(Dza07)||' '||SqlTime(Dza08) ModDT" & _
                ",DZA01,DZA02,S.st15 as DeptNo,3 as Sort " & _
                "From DutyZoneAssign,Staff W,Staff S,Acc090,Staff C,Staff M " & _
                "Where DZA01=W.ST01(+) And DZA02=S.ST01(+) And SubStr(DZA02,1,1)<>'" & 客戶編號 & "' " & strWhere2 & _
                " And Dza03=C.ST01(+) And Dza06=M.ST01(+) And S.st15=a0901(+) And S.ST01 is not null " & _
                "Order by DZA01,DZA02"
    'end 2024/01/11
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        Set GRD1.Recordset = rsTmp
        QueryData = True
        rsTmp.MoveFirst
        m_CurrKEY(0) = rsTmp.Fields("DZA01")
        m_CurrKEY(1) = rsTmp.Fields("DZA02")
        UpdateCtrlData
    End If
    SetGridWidth
    GetSelChage
    UpdateToolbarState
    RefreshRange
    Set rsTmp = Nothing
End Function

Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
    Dim rsTmp As New ADODB.Recordset, strSql As String, intA As Integer

    If IsRecordExist(0, strKEY01, strKEY02) = True Then
        m_CurrKEY(0) = strKEY01
        m_CurrKEY(1) = strKEY02
    End If
    
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = '" & m_CurrKEY(0) & "' " & _
                   "And DZA02 =(Select Min(DZA02) From DutyZoneAssign Where DZA01= '" & m_CurrKEY(0) & "' And DZA02 > '" & m_CurrKEY(1) & "') "
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
        UpdateCtrlData
        GoTo EXITSUB
    End If
    rsTmp.Close
    
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = (Select Min(DZA01) From DutyZoneAssign Where DZA01 > '" & m_CurrKEY(0) & "') " & _
                   "And DZA02 =(Select Min(DZA02) From DutyZoneAssign " & _
                                            "Where DZA01= (Select Min(DZA01) From DutyZoneAssign Where DZA01 > '" & m_CurrKEY(0) & "') )"
                                            
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
    Else
        ShowLastRecord
        GoTo EXITSUB
    End If
    rsTmp.Close
    UpdateCtrlData
    
EXITSUB:
    Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
    Dim rsTmp As New ADODB.Recordset, strSql As String, intA As Integer
    
    If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
        ShowMsg MsgText(9008)
        GoTo EXITSUB
    End If
   
   strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = '" & m_CurrKEY(0) & "' " & _
                   "And DZA02 =(Select Max(DZA02) From DutyZoneAssign Where DZA01= '" & m_CurrKEY(0) & "' And DZA02 < '" & m_CurrKEY(1) & "') "
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
        rsTmp.Close
        UpdateCtrlData
        GoTo EXITSUB
    End If
    rsTmp.Close
   
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = (Select Max(DZA01) From DutyZoneAssign Where DZA01 < '" & m_CurrKEY(0) & "') " & _
                   "And DZA02 =(Select Max(DZA02) From DutyZoneAssign " & _
                                            "Where DZA01= (Select Max(DZA01) From DutyZoneAssign Where DZA01 < '" & m_CurrKEY(0) & "') )"
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
    End If
    UpdateCtrlData
    rsTmp.Close
  
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
    Dim rsTmp As New ADODB.Recordset, strSql As String, intA As Integer
   
    If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
        ShowMsg MsgText(9009)
        GoTo EXITSUB
    End If
    
    strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = '" & m_CurrKEY(0) & "' " & _
                   "And DZA02 =(Select Min(DZA02) From DutyZoneAssign Where DZA01= '" & m_CurrKEY(0) & "' And DZA02 > '" & m_CurrKEY(1) & "') "
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
        rsTmp.Close
        UpdateCtrlData
        GoTo EXITSUB
   End If
   rsTmp.Close
    
  strSql = "Select * From DutyZoneAssign " & _
                "Where DZA01 = (Select Min(DZA01) From DutyZoneAssign Where DZA01> '" & m_CurrKEY(0) & "') " & _
                "And DZA02 =(Select Min(DZA02) From DutyZoneAssign Where DZA01= (Select Min(DZA01) From DutyZoneAssign Where DZA01> '" & m_CurrKEY(0) & "') )"
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        If IsNull(rsTmp.Fields("DZA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DZA01")
        If IsNull(rsTmp.Fields("DZA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DZA02")
    End If
    UpdateCtrlData
    rsTmp.Close
   
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   
   UpdateCtrlData
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
    Dim rsTmp As New ADODB.Recordset, strSql As String, stField As String, stTB As String, stWhere As String, intA As Integer

    m_DZA01 = m_CurrKEY(0)
    m_DZA02 = m_CurrKEY(1)
    
    'Modify by Amy 2024/01/11 智權人員要可輸MCTF0X
    If Left(m_DZA02, 1) = 客戶編號 Then
        stField = ",Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as SP,'' as Dept,'1' as SC"
        stTB = ",Customer "
        stWhere = "And SubStr(DZA02,1," & Len(m_DZA02) & ") ='" & m_DZA02 & "' "
        If Len(m_DZA02) = 6 Then
            '6碼的顯示8碼名稱
            stWhere = stWhere & "And SubStr(DZA02,1," & Len(m_DZA02) & ")=SubStr(CU01,1," & Len(m_DZA02) & ") And CU01='" & m_DZA02 & "00" & "' "
        Else
            stWhere = stWhere & "And DZA02=CU01(+) "
        End If
        stWhere = stWhere & "And CU02='0' And SubStr(DZA02,1,1)='" & 客戶編號 & "' And cu01 is not null "
    Else
        stField = ",S.st02 as SP,a0901||' '||a0902 as Dept,'2' as SC"
        stTB = ",Staff S,Acc090 "
        stWhere = "And DZA02=S.st01(+) And S.st15=a0901(+) And SubStr(DZA02,1,1)<>'" & 客戶編號 & "' And S.st01 is not null "
    End If
    'end 2024/01/11
    strSql = "Select DutyZoneAssign.*,P.st02 as WP" & stField & " " & _
                  ",C.st02||' '||SqlDatet(Dza04)||' '||SqlTime(Dza05) AddDT,M.st02||' '||SqlDatet(Dza07)||' '||SqlTime(Dza08) ModDT " & _
                  "From DutyZoneAssign,Staff P,Staff C,Staff M" & stTB & _
                   "Where DZA01 = '" & m_DZA01 & "' And DZA02 ='" & m_DZA02 & "' And DZA01=P.st01(+) " & _
                    "And Dza03=C.ST01(+) And Dza06=M.ST01(+) " & stWhere
    
    intA = 1
    Set rsTmp = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
        '來源-客戶
        If "" & rsTmp.Fields("SC") = "1" Then
            Option1(1).Value = 1
        '智權
        Else
            Option1(0).Value = 1
        End If
        txtDB1 = "" & rsTmp.Fields("DZA01")
        Label3(1) = "" & rsTmp.Fields("WP")
        txtDB2 = "" & rsTmp.Fields("DZA02")
        'txtDB2.Tag = txtDB2 'Mark by Amy 2023/05/31 不使用
        Label3(2) = "" & rsTmp.Fields("SP")
        Label3(0) = "" & rsTmp.Fields("Dept")
        Label23(0) = "CREATE : " & rsTmp.Fields("AddDT")
        Label23(1) = "UPDATE : " & rsTmp.Fields("ModDT")
        'Add by Amy 2023/05/31
        For i = LBound(arrDZAOld) To UBound(arrDZAOld)
            arrDZAOld(i) = "" & rsTmp.Fields(i)
        Next i
        'end 2023/05/31
        GetSelChage
    End If
End Sub

Private Sub FormClear()
    Dim oLab
    
    txtDB1 = ""
    txtDB2 = ""
    'txtDB2.Tag = "" 'Mark by Amy 2023/05/31  不使用
    For Each oLab In Label3
        oLab.Caption = ""
    Next
    For Each oLab In Label23
        oLab.Caption = ""
    Next
    'Add by Amy 2023/05/31
    For i = LBound(arrDZAOld) To UBound(arrDZAOld)
      arrDZAOld(i) = ""
    Next i
End Sub

'更新各控制項的狀態
'Mark by Amy 2023/05/31 開放也可改承辦人,故不需使用
Private Sub SetLock()
'    txtDB1.Locked = False
'    txtDB2.Locked = False
'    If m_EditMode = 0 Then
'        txtDB1.Locked = True
'        txtDB2.Locked = True
'    '修改
'    ElseIf m_EditMode = 2 Then
'        txtDB1.Locked = True
'    End If
End Sub

'更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
    Select Case m_EditMode
        '無任何動作
        Case 0:
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
            If m_bQuery Then
                TBar1.Buttons(4).Enabled = True
            Else
                TBar1.Buttons(4).Enabled = False
            End If
            '上/下/第一/最後 筆
            If m_bQuery Then
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
        Case 1, 2, 3, 4:
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

Private Function FormCheck() As Boolean
    Dim strTmp As String, strTmp1 As String, stF As String
    
    FormCheck = False
    '承辦人
    If Trim(txtDB1) = MsgText(601) Then
        MsgBox "承辦人不可為空！", vbExclamation + vbOKOnly
        txtDB1_GotFocus
        Exit Function
    ElseIf PUB_GetStaffNameDept(txtDB1, strTmp, strTmp1, True, True) = False Then
        txtDB1_GotFocus
        Exit Function
    End If
    
    '智權/客戶
    If Option1(0).Value = True Then
        stF = Option1(0).Caption
    Else
        stF = Option1(1).Caption
    End If
    If Trim(txtDB2) = MsgText(601) Then
        MsgBox stF & "不可為空！", vbExclamation + vbOKOnly
        txtDB2_GotFocus
        Exit Function
    End If
    If stF = "智權" Then
        'Modify by Amy 2024/01/11 智權人員要可輸MCTF0X
        If Len(txtDB2) = 5 And Len(txtDB2) = 6 Then
            MsgBox stF & "編號只能為 5 or 6 碼！", vbExclamation + vbOKOnly
            txtDB2_GotFocus
            Exit Function
        Else
            strTmp = "": strTmp1 = ""
            If PUB_GetStaffNameDept(txtDB2, strTmp, strTmp1, True, True) = False Then
                txtDB2_GotFocus
                Exit Function
            End If
        End If
    End If
    If stF = "客戶" Then
        If Len(txtDB2) <> 6 And Len(txtDB2) <> 8 Then
            MsgBox stF & "編號只能為 6碼或 8碼！", vbExclamation + vbOKOnly
            txtDB2_GotFocus
            Exit Function
        Else
            strTmp = ""
            If GetCusName(txtDB2, strTmp, True) = False Then
                '新增/修改
                If (m_EditMode = 1 Or m_EditMode = 2) And strTmp = MsgText(601) Then
                    MsgBox "無此" & stF & "編號，請確認！", vbExclamation + vbOKOnly
                End If
                txtDB2_GotFocus
                Exit Function
            End If
        End If
    End If
    '新增/修改
    If m_EditMode = 1 Or m_EditMode = 2 Then
        strTmp = ""
        If IsRecordExist(1, txtDB1, txtDB2, strTmp) = True Then
            'Modify by Amy 2023/05/31
'            If InStr(strTmp, "同一承辦人") > 0 Then
'                MsgBox strTmp, vbExclamation + vbOKOnly
'                Exit Function
'            Else
'                strTmp = "此編號已於設定於下列承辦人：" & vbCrLf & strTmp & vbCrLf & "要繼續？"
'                If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
'                    Exit Function
'                End If
'            End If
             MsgBox strTmp, vbExclamation + vbOKOnly
             Exit Function
             'end 2023/05/31
        End If
    End If
  
    FormCheck = True
End Function

Private Sub Grd1_Click()
    Dim nRow As Long, nCol As Long
    
    GRD1.Visible = False
    GRD1.row = GRD1.MouseRow
    GRD1.col = GRD1.MouseCol
    nRow = GRD1.row
    nCol = GRD1.col
    If nRow = 0 Then
        If Me.GRD1.Text <> "V" Then
            If m_blnColOrderAsc = True Then
                Me.GRD1.Sort = 5 '字串昇冪
                m_blnColOrderAsc = False
            Else
                Me.GRD1.Sort = 6 '字串降冪
                m_blnColOrderAsc = True
            End If
            bolQuery = True
        End If
    End If
    
    GRD1.Visible = True
End Sub

Private Sub GetSelChage()
    Dim j As Integer
    
    GRD1.Visible = False
    If GRD1.Rows - 1 > 0 Then
        '上一筆資料列清除反白
        If dblPrevRow > 0 Then
            GRD1.col = 2
            GRD1.row = dblPrevRow
            For i = 0 To 1
                GRD1.col = i
                GRD1.CellBackColor = &H8000000F
            Next i
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = QBColor(15)
            Next i
        End If
        '尋找目前資料列
        For j = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(j, GetValue("DZA01")) = m_CurrKEY(0) And GRD1.TextMatrix(j, GetValue("DZA02")) = m_CurrKEY(1) Then
                GRD1.col = 0
                GRD1.row = j
                dblPrevRow = GRD1.row
                For i = 0 To GRD1.Cols - 1
                    GRD1.col = i
                    GRD1.CellBackColor = &HFFC0C0
                Next i
                Exit For
            End If
        Next j
    End If
    GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nRow As Long, nCol As Long
    getGrdColRow GRD1, x, y, nCol, nRow
    If nCol < 0 Then Exit Sub
    If nRow < 0 Then Exit Sub
    GRD1.col = nCol
    GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
    GRD1.Visible = False
    If GRD1.row <> 0 Then
        m_CurrKEY(0) = GRD1.TextMatrix(GRD1.row, GetValue("DZA01")) '承辦人
        m_CurrKEY(1) = GRD1.TextMatrix(GRD1.row, GetValue("DZA02")) '智權/客戶
        UpdateCtrlData
    End If
    GRD1.Visible = True
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        '新增
        Case 1: OnAction vbKeyF2
        '修改
        Case 2: OnAction vbKeyF3
        '刪除
        Case 3: OnAction vbKeyF5
        '查詢
        Case 4: OnAction vbKeyF4
        '第一筆
        Case 6: OnAction vbKeyHome
        '前一筆
        Case 7: OnAction vbKeyPageUp
        '後一筆
        Case 8: OnAction vbKeyPageDown
        '最後一筆
        Case 9: OnAction vbKeyEnd
        '確定
        Case 11: OnAction vbKeyF9
        '取消
        Case 12: OnAction vbKeyF10
        '離開
        Case 14: OnAction vbKeyEscape
    End Select
End Sub

Private Function GetCusName(ByVal stNo As String, ByRef stName As String, bolShowMsg As Boolean) As Boolean
    Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
    
    GetCusName = False: stName = ""
    If Left(stNo, 1) <> 客戶編號 Then
        If bolShowMsg = True Then
            MsgBox "請輸入正確客戶編號！", vbExclamation + vbOKOnly
        End If
        Exit Function
    ElseIf Len(stNo) <> 6 And Len(stNo) <> 8 Then
        If bolShowMsg = True Then
            MsgBox "請輸入正確的 6碼或 8碼客戶編號！", vbExclamation + vbOKOnly
        End If
        Exit Function
    End If
    
    strQ = "Select Decode(cu04,null,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as CusName " & _
                "From Customer Where SubStr(cu01,1," & Len(stNo) & ")='" & stNo & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        stName = "" & RsQ.Fields("CusName")
        GetCusName = True
    End If
    Set RsQ = Nothing
End Function

Private Sub SetGridWidth()
    Dim ii As Integer
    
    With GRD1
        .FormatString = strAllField
        For ii = LBound(arrF) To UBound(arrF)
            .ColWidth(ii) = arrW(ii)
            .ColAlignment(ii) = flexAlignLeftCenter
        Next ii
    End With
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = LBound(arrF) To UBound(arrF)
        If UCase(arrF(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function

Private Sub txtDB1_GotFocus()
    TextInverse txtDB1
End Sub

Private Sub txtDB1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDB1_Validate(Cancel As Boolean)
    Dim strTmp As String, strTmp1 As String
    
    If txtDB1 = MsgText(601) Then Exit Sub
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    If PUB_GetStaffNameDept(txtDB1, strTmp, strTmp1, True, False) = False Then
        strTmp = ""
    End If
    Label3(1).Caption = strTmp
End Sub

Private Sub txtDB2_GotFocus()
    TextInverse txtDB1
End Sub

Private Sub txtDB2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDB2_Validate(Cancel As Boolean)
    Dim strTmp As String, strTmp1 As String, bolCus As Boolean
    
    If txtDB2 = MsgText(601) Then Exit Sub
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    '智權/客戶欄
    If Left(txtDB2, 1) = 客戶編號 Then
        Option1(1).Value = True
        bolCus = True
    Else
        Option1(0).Value = True
    End If
   
    strTmp = "": strTmp1 = ""
    If bolCus = True Then
        If GetCusName(txtDB2, strTmp, True) = False Then
            strTmp = "": strTmp1 = ""
        End If
    Else
        If PUB_GetStaffNameDept(txtDB2, strTmp, strTmp1, True, False) = False Then
            strTmp = "": strTmp1 = ""
        End If
    End If
    Label3(2).Caption = strTmp
    Label3(0).Caption = strTmp1
End Sub
