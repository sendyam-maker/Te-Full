VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040157 
   BorderStyle     =   1  '單線固定
   Caption         =   "特殊客戶/代理人收文費用維護"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8170
   Begin VB.TextBox textSG08 
      Height          =   270
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3900
      Width           =   1275
   End
   Begin VB.TextBox textSG07 
      Height          =   270
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3570
      Width           =   1275
   End
   Begin VB.TextBox textSG02 
      Height          =   270
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1620
      Width           =   1485
   End
   Begin VB.TextBox textSG06 
      Height          =   270
      Left            =   1935
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2940
      Width           =   1485
   End
   Begin VB.TextBox textSG03 
      Height          =   270
      Left            =   1935
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1950
      Width           =   675
   End
   Begin VB.TextBox textSG05 
      Height          =   270
      Left            =   1935
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2610
      Width           =   675
   End
   Begin VB.TextBox textSG04 
      Height          =   270
      Left            =   1935
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2280
      Width           =   675
   End
   Begin VB.TextBox textSG01 
      Height          =   270
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1290
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
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
            Picture         =   "frm12040157.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040157.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   470
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8170
      _ExtentX        =   14411
      _ExtentY        =   829
      ButtonWidth     =   1076
      ButtonHeight    =   794
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
      BorderStyle     =   1
   End
   Begin VB.Label Label3 
      Caption         =   "備註：僅限某申請人或某代理人，則於申請人編號與代理人編號欄輸入相同編號"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   330
      TabIndex        =   21
      Top             =   810
      Width           =   6435
   End
   Begin VB.Label LabelSG02 
      Height          =   240
      Left            =   3450
      TabIndex        =   20
      Top             =   1650
      Width           =   4635
   End
   Begin VB.Label LabelSG01 
      Height          =   240
      Left            =   3450
      TabIndex        =   19
      Top             =   1320
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請款上限："
      Height          =   180
      Index           =   3
      Left            =   975
      TabIndex        =   18
      Top             =   3930
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "預設服務費："
      Height          =   180
      Index           =   2
      Left            =   795
      TabIndex        =   17
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   1
      Left            =   795
      TabIndex        =   16
      Top             =   1665
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8010
      Y1              =   3360
      Y2              =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "啟用日期："
      Height          =   180
      Index           =   17
      Left            =   975
      TabIndex        =   15
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      Height          =   180
      Left            =   1155
      TabIndex        =   14
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label LabelSG05 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   2640
      TabIndex        =   13
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label LabelSG04 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   2640
      TabIndex        =   12
      Top             =   2310
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "國家代碼："
      Height          =   180
      Left            =   975
      TabIndex        =   11
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   975
      TabIndex        =   10
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   0
      Left            =   795
      TabIndex        =   9
      Top             =   1335
      Width           =   1080
   End
End
Attribute VB_Name = "frm12040157"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0不用改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Create By Sindy 2012/11/21
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的Key值
Dim m_FirstKEY(6) As String
' 最後一筆資料的Key值
Dim m_LastKEY(6) As String
' 目前正在顯示的Key值
Dim m_CurrKEY(6) As String
Dim StrSQLa As String
'Memo By Sonia 2021/12/24 Form2.0不用改
Dim rsA As New ADODB.Recordset
Dim tf_SG As Integer
Dim strText As String, arrKey As Variant
Dim m_QuerySystem As String
Const m_strNoRightMsg As String = "您無權限查詢或維護此系統類別+案件性質資料"

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from SpecGuestFee where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SG = rsA.Fields.Count
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
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
   
   ReDim m_FieldList(tf_SG) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSG01.BackColor = &H8000000F
   textSG02.BackColor = &H8000000F
   textSG03.BackColor = &H8000000F
   textSG04.BackColor = &H8000000F
   textSG05.BackColor = &H8000000F
   textSG06.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   FilterSystem
End Sub

Private Sub FilterSystem()
   Dim nIndex As Integer
   Dim nCount As Integer
   Dim strSys As String
   Dim strTemp As String
   m_QuerySystem = Empty
   
   strSys = GetUserSystemKind
   nCount = GetSubStringCount(strSys)
   For nIndex = 1 To nCount
      strTemp = GetSubString(strSys, nIndex)
      If IsEmptyText(m_QuerySystem) = False Then m_QuerySystem = m_QuerySystem & ","
      m_QuerySystem = m_QuerySystem & "'" & strTemp & "'"
NextRecord:
   Next nIndex
   
   m_QuerySystem = "(" & m_QuerySystem & ")"
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm12040157 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textSG01.Enabled = True Then
      Cancel = False
      textSG01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSG02.Enabled = True Then
      Cancel = False
      textSG02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSG03.Enabled = True Then
      Cancel = False
      textSG03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSG04.Enabled = True Then
      Cancel = False
      textSG04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSG05.Enabled = True Then
      Cancel = False
      textSG05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSG06.Enabled = True Then
      Cancel = False
      textSG06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_SG - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_SG - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
   
   AddRecord = False
   
   ' 檢查記錄是否已存在
   If IsRecordExist(textSG01, textSG02, textSG03, textSG04, textSG05, DBDATE(textSG06)) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO SpecGuestFee ("
   For nIndex = 0 To tf_SG - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_SG - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If ((textSG01 & textSG02 & textSG03 & textSG04 & textSG05 & DBDATE(textSG06)) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2) & m_FirstKEY(3) & m_FirstKEY(4) & m_FirstKEY(5))) Or _
      ((textSG01 & textSG02 & textSG03 & textSG04 & textSG05 & DBDATE(textSG06)) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2) & m_LastKEY(3) & m_LastKEY(4) & m_LastKEY(5))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord textSG01, textSG02, textSG03, textSG04, textSG05, DBDATE(textSG06)
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
       
   ModRecord = False
   
   strSql = "begin user_data.user_enabled:=1; UPDATE SpecGuestFee SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SG - 1
      strTmp = Empty
      'If nIndex < 7 Or nIndex > 12 Then
         If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
            If m_FieldList(nIndex).fiType = 0 Then
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
               End If
            Else
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
               End If
            End If
         End If
         If strTmp <> Empty Then
            bDifference = True
            If bFirst = True Then
               strSql = strSql & strTmp
               bFirst = False
            Else
               strSql = strSql & "," & strTmp
            End If
         End If
      'End If
   Next nIndex

   strSql = strSql & " " & _
            "WHERE sg01='" & m_CurrKEY(0) & "' and sg02='" & m_CurrKEY(1) & "' and sg03='" & m_CurrKEY(2) & _
            "' and sg04='" & m_CurrKEY(3) & "' and sg05='" & m_CurrKEY(4) & "' and sg06='" & m_CurrKEY(5) & "'; end;"
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2), m_CurrKEY(3), m_CurrKEY(4), m_CurrKEY(5)
      
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM SpecGuestFee " & _
            "WHERE sg01 = '" & m_CurrKEY(0) & "'  and sg02='" & m_CurrKEY(1) & "' " & _
              "and sg03 = '" & m_CurrKEY(2) & "'  and sg04='" & m_CurrKEY(3) & "' " & _
              "and sg05 = '" & m_CurrKEY(4) & "'  and sg06='" & m_CurrKEY(5) & "' "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) And _
       m_CurrKEY(3) = m_LastKEY(3) And m_CurrKEY(4) = m_LastKEY(4) And m_CurrKEY(5) = m_LastKEY(5)) Or _
      (m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) And _
       m_CurrKEY(3) = m_FirstKEY(3) And m_CurrKEY(4) = m_FirstKEY(4) And m_CurrKEY(5) = m_FirstKEY(5)) Then
      RefreshRange
   End If
   ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2), m_CurrKEY(3), m_CurrKEY(4), m_CurrKEY(5)
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If textSG06 = "" Then
      '若沒有輸入啟用日期時,抓最大日期
      strSql = "select max(SG06) from SpecGuestFee " & _
                "WHERE sg01='" & textSG01 & "' and sg02='" & textSG02 & "' " & _
                  "and sg03='" & textSG03 & "' and sg04='" & textSG04 & "' " & _
                  "and sg05='" & textSG05 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         textSG06 = Val(rsTmp.Fields(0)) - 19110000
      Else
         rsTmp.Close
         Set rsTmp = Nothing
         QueryRecord = False
         Exit Function
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   
   QueryRecord = False
   
   If IsRecordExist(textSG01, textSG02, textSG03, textSG04, textSG05, DBDATE(textSG06)) = True Then
      m_CurrKEY(0) = textSG01
      m_CurrKEY(1) = textSG02
      m_CurrKEY(2) = textSG03
      m_CurrKEY(3) = textSG04
      m_CurrKEY(4) = textSG05
      m_CurrKEY(5) = DBDATE(textSG06)
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            If textSG01 = "" And textSG02 <> "" Then textSG01 = textSG02
'            If textSG01 <> "" And textSG02 = "" Then textSG02 = textSG01
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2), m_CurrKEY(3), m_CurrKEY(4), m_CurrKEY(5)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSG01 = "" And textSG02 <> "" Then textSG01 = textSG02
         If textSG01 <> "" And textSG02 = "" Then textSG02 = textSG01
         If textSG01 <> "" And textSG02 <> "" And textSG03 <> "" And textSG04 <> "" And textSG05 <> "" Then 'And textSG06 <> ""
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If textSG01 = "" Or textSG02 = "" Or textSG03 = "" Or textSG04 = "" Or textSG05 = "" Then 'Or textSG06 = ""
               MsgBox "查詢條件必須全部輸齊，才可進行查詢動作！", vbInformation
            End If
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textSG01.SetFocus
      Case 2: If Me.Visible = True Then textSG07.SetFocus
      Case 4: If Me.Visible = True Then textSG01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, _
                               ByVal strKEY04 As String, ByVal strKEY05 As String, ByVal strKEY06 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM SpecGuestFee " & _
             "WHERE sg01='" & strKEY01 & "' and sg02='" & strKEY02 & "' and sg03='" & strKEY03 & "' " & _
               "and sg04='" & strKEY04 & "' and sg05='" & strKEY05 & "' and sg06='" & strKEY06 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, _
                           ByVal strKEY04 As String, ByVal strKEY05 As String, ByVal strKEY06 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from SpecGuestFee where rownum <2"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   If IsRecordExist(strKEY01, strKEY02, strKEY03, strKEY04, strKEY05, strKEY06) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
      m_CurrKEY(3) = strKEY04
      m_CurrKEY(4) = strKEY05
      m_CurrKEY(5) = strKEY06
   Else
      strSql = "SELECT sg01,sg02,sg03,sg04,sg05,sg06 FROM SpecGuestFee " & _
                "WHERE sg01='" & m_CurrKEY(0) & "' and sg02='" & m_CurrKEY(1) & "' " & _
                  "and sg03='" & m_CurrKEY(2) & "' and sg04='" & m_CurrKEY(3) & "' " & _
                  "and sg05='" & m_CurrKEY(4) & "' and sg06='" & m_CurrKEY(5) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sg01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sg01")
         If IsNull(rsTmp.Fields("sg02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sg02")
         If IsNull(rsTmp.Fields("sg03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sg03")
         If IsNull(rsTmp.Fields("sg04")) = False Then: m_CurrKEY(3) = rsTmp.Fields("sg04")
         If IsNull(rsTmp.Fields("sg05")) = False Then: m_CurrKEY(4) = rsTmp.Fields("sg05")
         If IsNull(rsTmp.Fields("sg06")) = False Then: m_CurrKEY(5) = rsTmp.Fields("sg06")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT min(sg01||'-'||sg02||'-'||sg03||'-'||sg04||'-'||sg05||'-'||sg06) FROM SpecGuestFee "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strText = "" & rsTmp.Fields(0)
         If Trim(strText) > "" Then
            arrKey = Split(strText, "-")
            m_CurrKEY(0) = arrKey(0)
            m_CurrKEY(1) = arrKey(1)
            m_CurrKEY(2) = arrKey(2)
            m_CurrKEY(3) = arrKey(3)
            m_CurrKEY(4) = arrKey(4)
            m_CurrKEY(5) = arrKey(5)
         End If
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   m_CurrKEY(3) = m_FirstKEY(3)
   m_CurrKEY(4) = m_FirstKEY(4)
   m_CurrKEY(5) = m_FirstKEY(5)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) _
      And m_CurrKEY(2) = m_FirstKEY(2) And m_CurrKEY(3) = m_FirstKEY(3) _
      And m_CurrKEY(4) = m_FirstKEY(4) And m_CurrKEY(5) = m_FirstKEY(5) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "select max(sg01||'-'||sg02||'-'||sg03||'-'||sg04||'-'||sg05||'-'||sg06) From SpecGuestFee " & _
             "where sg01||sg02||sg03||sg04||sg05||sg06<'" & m_CurrKEY(0) & m_CurrKEY(1) & m_CurrKEY(2) & m_CurrKEY(3) & m_CurrKEY(4) & m_CurrKEY(5) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_CurrKEY(0) = arrKey(0)
         m_CurrKEY(1) = arrKey(1)
         m_CurrKEY(2) = arrKey(2)
         m_CurrKEY(3) = arrKey(3)
         m_CurrKEY(4) = arrKey(4)
         m_CurrKEY(5) = arrKey(5)
      End If
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) _
      And m_CurrKEY(2) = m_LastKEY(2) And m_CurrKEY(3) = m_LastKEY(3) _
      And m_CurrKEY(4) = m_LastKEY(4) And m_CurrKEY(5) = m_LastKEY(5) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "select min(sg01||'-'||sg02||'-'||sg03||'-'||sg04||'-'||sg05||'-'||sg06) From SpecGuestFee " & _
             "where sg01||sg02||sg03||sg04||sg05||sg06>'" & m_CurrKEY(0) & m_CurrKEY(1) & m_CurrKEY(2) & m_CurrKEY(3) & m_CurrKEY(4) & m_CurrKEY(5) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_CurrKEY(0) = arrKey(0)
         m_CurrKEY(1) = arrKey(1)
         m_CurrKEY(2) = arrKey(2)
         m_CurrKEY(3) = arrKey(3)
         m_CurrKEY(4) = arrKey(4)
         m_CurrKEY(5) = arrKey(5)
      End If
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   m_CurrKEY(3) = m_LastKEY(3)
   m_CurrKEY(4) = m_LastKEY(4)
   m_CurrKEY(5) = m_LastKEY(5)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         textSG06 = strSrvDate(2)
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
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
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from SpecGuestFee where rownum <2"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   strSql = "SELECT min(sg01||'-'||sg02||'-'||sg03||'-'||sg04||'-'||sg05||'-'||sg06) FROM SpecGuestFee "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_FirstKEY(0) = arrKey(0)
         m_FirstKEY(1) = arrKey(1)
         m_FirstKEY(2) = arrKey(2)
         m_FirstKEY(3) = arrKey(3)
         m_FirstKEY(4) = arrKey(4)
         m_FirstKEY(5) = arrKey(5)
      End If
   End If
   rsTmp.Close
   
   strSql = "SELECT max(sg01||'-'||sg02||'-'||sg03||'-'||sg04||'-'||sg05||'-'||sg06) FROM SpecGuestFee "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_LastKEY(0) = arrKey(0)
         m_LastKEY(1) = arrKey(1)
         m_LastKEY(2) = arrKey(2)
         m_LastKEY(3) = arrKey(3)
         m_LastKEY(4) = arrKey(4)
         m_LastKEY(5) = arrKey(5)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM SpecGuestFee " & _
            "WHERE sg01='" & m_CurrKEY(0) & "' and sg02='" & m_CurrKEY(1) & "' " & _
              "and sg03='" & m_CurrKEY(2) & "' and sg04='" & m_CurrKEY(3) & "' " & _
              "and sg05='" & m_CurrKEY(4) & "' and sg06='" & m_CurrKEY(5) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sg01")) = False Then: textSG01 = rsTmp.Fields("sg01")
      If IsNull(rsTmp.Fields("sg02")) = False Then: textSG02 = rsTmp.Fields("sg02")
      If IsNull(rsTmp.Fields("sg03")) = False Then: textSG03 = rsTmp.Fields("sg03")
      If IsNull(rsTmp.Fields("sg04")) = False Then: textSG04 = rsTmp.Fields("sg04")
      If IsNull(rsTmp.Fields("sg05")) = False Then: textSG05 = rsTmp.Fields("sg05")
      If IsNull(rsTmp.Fields("sg06")) = False Then: textSG06 = TAIWANDATE(rsTmp.Fields("sg06"))
      If IsNull(rsTmp.Fields("sg07")) = False Then: textSG07 = rsTmp.Fields("sg07")
      If IsNull(rsTmp.Fields("sg08")) = False Then: textSG08 = rsTmp.Fields("sg08")
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      If Left(textSG01, 1) = "X" Then
         LabelSG01 = GetPrjPeople1(textSG01, "1")
      Else
         LabelSG01 = GetPrjName1(textSG01)
      End If
      If Left(textSG02, 1) = "X" Then
         LabelSG02 = GetPrjPeople1(textSG02, "1")
      Else
         LabelSG02 = GetPrjName1(textSG02)
      End If
      LabelSG04 = GetPrjNationName(textSG04)
      LabelSG05 = GetPrjState6(textSG03, textSG05)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
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
         ' 新增
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

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid = False
   
   If textSG01.Text = "" Then
      MsgBox "申請人編號不可空白！", vbExclamation
      textSG01.SetFocus
      Exit Function
   End If
   
   If textSG02.Text = "" Then
      MsgBox "代理人編號不可空白！", vbExclamation
      textSG02.SetFocus
      Exit Function
   End If
   
   If textSG03.Text = "" Then
      MsgBox "系統別不可空白！", vbExclamation
      textSG03.SetFocus
      Exit Function
   End If
   
   If textSG04.Text = "" Then
      MsgBox "國家代碼不可空白！", vbExclamation
      textSG04.SetFocus
      Exit Function
   End If
   
   If textSG05.Text = "" Then
      MsgBox "案件性質不可空白！", vbExclamation
      textSG05.SetFocus
      Exit Function
   End If
   
   If textSG06.Text = "" Then
      MsgBox "啟用日期不可空白！", vbExclamation
      textSG06.SetFocus
      Exit Function
   End If
   
   If Val(textSG08.Text) < Val(textSG07.Text) Then
      MsgBox "請款上限不可小於預設服務費！", vbExclamation
      textSG08.SetFocus
      Exit Function
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSG01.Locked = bEnable
   textSG02.Locked = bEnable
   textSG03.Locked = bEnable
   textSG04.Locked = bEnable
   textSG05.Locked = bEnable
   textSG06.Locked = bEnable
   If bEnable Then textSG01.BackColor = &H8000000F Else textSG01.BackColor = &H80000005
   If bEnable Then textSG02.BackColor = &H8000000F Else textSG02.BackColor = &H80000005
   If bEnable Then textSG03.BackColor = &H8000000F Else textSG03.BackColor = &H80000005
   If bEnable Then textSG04.BackColor = &H8000000F Else textSG04.BackColor = &H80000005
   If bEnable Then textSG05.BackColor = &H8000000F Else textSG05.BackColor = &H80000005
   If bEnable Then textSG06.BackColor = &H8000000F Else textSG06.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSG01.Locked = bEnable
   textSG02.Locked = bEnable
   textSG03.Locked = bEnable
   textSG04.Locked = bEnable
   textSG05.Locked = bEnable
   textSG06.Locked = bEnable
   If bEnable Then textSG01.BackColor = &H8000000F Else textSG01.BackColor = &H80000005
   If bEnable Then textSG02.BackColor = &H8000000F Else textSG02.BackColor = &H80000005
   If bEnable Then textSG03.BackColor = &H8000000F Else textSG03.BackColor = &H80000005
   If bEnable Then textSG04.BackColor = &H8000000F Else textSG04.BackColor = &H80000005
   If bEnable Then textSG05.BackColor = &H8000000F Else textSG05.BackColor = &H80000005
   If bEnable Then textSG06.BackColor = &H8000000F Else textSG06.BackColor = &H80000005
   textSG07.Locked = bEnable
   textSG08.Locked = bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSG01 = Empty
   LabelSG01 = Empty
   textSG02 = Empty
   LabelSG02 = Empty
   textSG03 = Empty
   textSG04 = Empty
   LabelSG04 = Empty
   textSG05 = Empty
   LabelSG05 = Empty
   textSG06 = Empty
   textSG07 = Empty
   textSG08 = Empty
   
   For nIndex = 0 To tf_SG - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SG01", textSG01
      SetFieldNewData "SG02", textSG02
      SetFieldNewData "SG03", textSG03
      SetFieldNewData "SG04", textSG04
      SetFieldNewData "SG05", textSG05
      SetFieldNewData "SG06", DBDATE(textSG06)
   End If
   SetFieldNewData "SG07", IIf(Trim(textSG07) = "", 0, textSG07)
   SetFieldNewData "SG08", textSG08
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SG
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SG" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 7, 8:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

Private Sub textSG01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG01
   End If
End Sub

Private Sub textSG01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSG01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelSG01 = Empty
   If IsEmptyText(textSG01) = False Then
      textSG01 = Left(textSG01 & "00000000", 8)
      If Left(textSG01, 1) = "X" Then
         LabelSG01 = GetPrjPeople1(textSG01, "1")
      Else
         LabelSG01 = GetPrjName1(textSG01)
      End If
      Select Case m_EditMode
         Case 1, 4:
'            If Left(textSG01, 1) <> "X" Then
'               Cancel = True
'               strTit = "檢核資料"
'               strMsg = "必須輸入客戶編號"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textSG01_GotFocus
'               GoTo EXITSUB
'            End If
            If IsEmptyText(LabelSG01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG01_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textSG02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG02
   End If
End Sub

Private Sub textSG02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSG02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelSG02 = Empty
   If IsEmptyText(textSG02) = False Then
      textSG02 = Left(textSG02 & "00000000", 8)
      If Left(textSG02, 1) = "X" Then
         LabelSG02 = GetPrjPeople1(textSG02, "1")
      Else
         LabelSG02 = GetPrjName1(textSG02)
      End If
      Select Case m_EditMode
         Case 1, 4:
'            If Left(textSG02, 1) <> "Y" Then
'               Cancel = True
'               strTit = "檢核資料"
'               strMsg = "必須輸入代理人編號"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textSG02_GotFocus
'               GoTo EXITSUB
'            End If
            If IsEmptyText(LabelSG02) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG02_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textSG03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG03
   End If
End Sub

Private Sub textSG03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textSG03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSG03) = False Then
      Select Case m_EditMode
         Case 1, 4:
            If IsAlphabetic(textSG03) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG03_GotFocus
               GoTo EXITSUB
            End If
            If IsUserHasRightOfSystem(strUserNum, textSG03) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "您沒有使用該系統類別的權限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG03_GotFocus
            End If
            If IsCorrectSysKind(textSG03) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG03_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textSG04_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG04
   End If
End Sub

' 國家代碼
Private Sub textSG04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelSG04 = Empty
   If IsEmptyText(textSG04) = False Then
      LabelSG04 = GetNationName(textSG04, 0)
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(LabelSG04) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "國家代碼不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG04_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

Private Sub textSG05_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG05
   End If
End Sub

Private Sub textSG05_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textSG05) = False Then
      If m_EditMode = 1 Then
         '若檢查使用者無權限新增此系統類別
         If IsRightExist(textSG03, textSG05) = False Then
            strMsg = m_strNoRightMsg
            nResponse = MsgBox(strMsg, vbOKOnly)
            textSG05.SetFocus
            Exit Sub
         End If
         If IsRecordExist(textSG01, textSG02, textSG03, textSG04, textSG05, DBDATE(textSG06)) = True Then
            'Cancel = True
            strTit = "檢核資料"
            strMsg = "該筆資料已經存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSG05.SetFocus
         End If
      End If
   End If
End Sub

' 檢查使用者是否有權限
Private Function IsRightExist(ByVal strCF01 As String, ByVal strCF03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRightExist = False
   strSql = "SELECT SG01,SG02,SG03 FROM Staff,Staff_Group " & _
            " WHERE ST11=SG01(+) AND SG02 IN " & m_QuerySystem & _
            " AND SG02='" & strCF01 & "' And ST01='" & strUserNum & "' " & _
            " And SG03='" & strCF03 & "'"
                  
   ' 讀取資料庫
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRightExist = True
   Else
      IsRightExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function


' 案件性質代號
Private Sub textSG05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelSG05 = Empty
   If IsEmptyText(textSG05) = False Then
      If textSG04 > "010" Then
         LabelSG05 = GetCaseTypeName(textSG03, textSG05, 1)
      Else
         LabelSG05 = GetCaseTypeName(textSG03, textSG05, 0)
      End If
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(LabelSG05) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "案件性質代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSG05_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

Private Sub textSG06_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG06
      CloseIme
   End If
End Sub

Private Sub textSG06_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textSG06_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSG06 <> "" Then
      If CheckIsTaiwanDate(textSG06, False) = False Then
         Call textSG06_GotFocus
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
'      ElseIf ChkWork(ChangeTStringToWString(textSG06)) = False Then
'         Call textSG06_GotFocus
'         Cancel = True
'         Exit Sub
      End If
   End If
End Sub

Private Sub textSG07_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG07
      CloseIme
   End If
End Sub

Private Sub textSG07_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textSG08_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSG08
      CloseIme
   End If
End Sub

Private Sub textSG08_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
