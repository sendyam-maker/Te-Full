VERSION 5.00
Begin VB.Form frm040104_1_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "作業失誤資料輸入"
   ClientHeight    =   4395
   ClientLeft      =   285
   ClientTop       =   2775
   ClientWidth     =   8040
   ControlBox      =   0   'False
   DrawMode        =   12  '沒有動作
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8040
   Begin VB.TextBox Text1 
      Height          =   1005
      Index           =   3
      Left            =   1140
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   1  '水平捲軸
      TabIndex        =   3
      Top             =   3270
      Width           =   6585
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   2910
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   1
      Top             =   2550
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2190
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   6708
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5730
      TabIndex        =   4
      Top             =   70
      Width           =   930
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   3300
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "失誤金額："
      Height          =   180
      Left            =   150
      TabIndex        =   23
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label lblStaffName 
      Height          =   255
      Left            =   2370
      TabIndex        =   22
      Top             =   2580
      Width           =   2475
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "失誤人員："
      Height          =   180
      Left            =   150
      TabIndex        =   21
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "失誤日期："
      Height          =   180
      Left            =   150
      TabIndex        =   20
      Top             =   2220
      Width           =   900
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      Index           =   1
      X1              =   150
      X2              =   7800
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   7800
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   4080
      TabIndex        =   19
      Top             =   1590
      Width           =   720
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   18
      Top             =   1590
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   16
      Top             =   1590
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   4080
      TabIndex        =   15
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   14
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   930
      Width           =   720
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   930
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm040104_1_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text1,lblCaseField)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit
'edit by nickc 2007/02/02
'Dim cp(1 To T_CP) As String
Dim cp() As String

Dim intLeaveKind As Integer

Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
    Case 0 '確定
        Screen.MousePointer = vbHourglass
        If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
        If FormSave = False Then
            Screen.MousePointer = vbDefault
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        Unload Me
    Case 1 '回前畫面
        Unload Me
    End Select
End Sub

Private Function FormSave() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

On Error GoTo ErrorHandler
cnnConnection.BeginTrans
FormSave = True
StrSQLa = "Select * From MissData Where MD01='" & Me.lblCaseField(0).Caption & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    strSql = "Update MissData Set MD02=" & DBDATE(Me.Text1(0).Text) & ", MD03='" & Me.Text1(1).Text & "', MD04=" & Val(Me.Text1(2).Text) & ", MD05='" & Me.Text1(3).Text & "' Where MD01='" & Me.lblCaseField(0).Caption & "' "
    cnnConnection.Execute strSql
Else
    strSql = "Insert Into MissData Values ('" & Me.lblCaseField(0).Caption & "'," & DBDATE(Me.Text1(0).Text) & ",'" & Me.Text1(1).Text & "', " & Val(Me.Text1(2).Text) & ",'" & Me.Text1(3).Text & "' )"
    cnnConnection.Execute strSql
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
cnnConnection.CommitTrans
Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Sub ReadAllData()
 On Error GoTo ErrHnd
    '總收文號
    cp(9) = Me.lblCaseField(0).Caption
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.ReadCaseProgressDatabase(cp(), intPWhere) Then
    If ClsPDReadCaseProgressDatabase(cp(), intPWhere) Then
        '本所案號
        If cp(1) = 馬德里案 Then
            Me.lblCaseField(1).Caption = cp(1) + " - " + Left(cp(2), 5) + _
                IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
                IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
                IIf(cp(4) = "00", "", " - " + cp(4))
        Else
            Me.lblCaseField(1).Caption = cp(1) + " - " + cp(2) + _
                IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
                IIf(cp(4) = "00", "", " - " + cp(4))
        End If
        '申請人
        Me.lblCaseField(2).Caption = PUB_GetCustName(Replace(Me.lblCaseField(1).Caption, "-", ""))
        '案件性質
        Me.lblCaseField(3).Caption = GetCasePropertyName(cp(9))
        '申請案號
        Me.lblCaseField(4).Caption = GetFileNo(Replace(Me.lblCaseField(1).Caption, "-", ""))
        '智權人員
        Me.lblCaseField(5).Caption = cp(13) & " " & GetStaffName(cp(13), True)
        '承辦人
        Me.lblCaseField(6).Caption = cp(14) & " " & GetStaffName(cp(14), True)
        '失誤日期
        Me.Text1(0).Text = strSrvDate(2)
    Else
        MsgBox "讀取CaseProgress檔時發生錯誤!!", vbCritical
        Unload Me
    End If
    Exit Sub
ErrHnd:
   MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    ReadAllData
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()


    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm040104_1_3 = Nothing
End Sub

'取得案件性質名稱
Private Function GetCasePropertyName(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCasePropertyName = ""
'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
StrSQLa = "Select Decode(PA09,'000',CPM03,CPM04) From CaseProgress, Patent ,CasePropertyMap Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 And CP09='" & strCP09 & "' "
StrSQLa = StrSQLa & " Union Select Decode(TM10,'000',CPM03,CPM04) From CaseProgress, Trademark ,CasePropertyMap Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 And CP09='" & strCP09 & "' "
StrSQLa = StrSQLa & " Union Select Decode(LC15,'000',CPM03,CPM04) From CaseProgress, Lawcase ,CasePropertyMap Where CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 And CP09='" & strCP09 & "' "
StrSQLa = StrSQLa & " Union Select Decode('000','000',CPM03,CPM04) From CaseProgress, Hirecase ,CasePropertyMap Where CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 And CP09='" & strCP09 & "' "
StrSQLa = StrSQLa & " Union Select Decode(SP09,'000',CPM03,CPM04) From CaseProgress, Servicepractice ,CasePropertyMap Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 And CP09='" & strCP09 & "' "
'end 2018/06/05
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCasePropertyName = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'取得申請案號
Private Function GetFileNo(strCaseNo As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetFileNo = ""
StrSQLa = "Select PA11 From Patent Where " & ChgPatent(strCaseNo)
StrSQLa = StrSQLa & " Union Select TM12 From Trademark Where " & ChgTradeMark(strCaseNo)
StrSQLa = StrSQLa & " Union Select SP11 From Servicepractice Where " & ChgService(strCaseNo)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetFileNo = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case 0 '失誤日期
        If Me.Text1(Index).Text <> "" Then
            If CheckIsTaiwanDate(Me.Text1(Index).Text) = False Then
                Cancel = True
            End If
        End If
    Case 1 '失誤人員
        Me.lblStaffName.Caption = GetStaffName(Me.Text1(Index).Text, False)
        If Me.Text1(Index).Text <> "" And Me.lblStaffName.Caption = "" Then
            Cancel = True
            MsgBox "失誤人員輸入錯誤!!!", vbExclamation + vbOKOnly
        End If
    End Select
    If Cancel = True Then TextInverse Me.Text1(Index)
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text1(0).Text = "" Then
    MsgBox "請輸入失誤日期!!!", vbExclamation + vbOKOnly
    Me.Text1(0).SetFocus
    TextInverse Me.Text1(0)
    Exit Function
End If
If Me.Text1(1).Text = "" Then
    MsgBox "請輸入失誤人員!!!", vbExclamation + vbOKOnly
    Me.Text1(1).SetFocus
    TextInverse Me.Text1(1)
    Exit Function
End If
If Me.Text1(2).Text = "" Then
    MsgBox "請輸入失誤金額!!!", vbExclamation + vbOKOnly
    Me.Text1(2).SetFocus
    TextInverse Me.Text1(2)
    Exit Function
End If
For Each objTxt In Me.Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text1(objTxt.Index).SetFocus
         Text1_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next
TxtValidate = True
End Function
