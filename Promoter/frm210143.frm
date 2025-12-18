VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm210143 
   BorderStyle     =   1  '單線固定
   Caption         =   "價目表"
   ClientHeight    =   5460
   ClientLeft      =   6090
   ClientTop       =   1550
   ClientWidth     =   9140
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9140
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(11)"
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   11
      Left            =   1500
      TabIndex        =   26
      Top             =   4740
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(10)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   10
      Left            =   4140
      TabIndex        =   24
      Top             =   3930
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(9)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   9
      Left            =   1500
      TabIndex        =   23
      Top             =   3930
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(8)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   8
      Left            =   4140
      TabIndex        =   20
      Top             =   3120
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(7)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   7
      Left            =   1500
      TabIndex        =   19
      Top             =   3120
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(6)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   6
      Left            =   6750
      TabIndex        =   16
      Top             =   2400
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(5)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   5
      Left            =   4140
      TabIndex        =   15
      Top             =   2400
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(4)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   4
      Left            =   1500
      TabIndex        =   14
      Top             =   2400
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(3)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   3
      Left            =   1500
      TabIndex        =   10
      Top             =   1590
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(2)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   2
      Left            =   6750
      TabIndex        =   8
      Top             =   870
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(1)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   4140
      TabIndex        =   7
      Top             =   870
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF2 
      Caption         =   "ChkPLF2(0)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   6
      Top             =   870
      Width           =   2505
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   7770
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(9)"
      Height          =   315
      Index           =   9
      Left            =   1260
      TabIndex        =   21
      Top             =   3690
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(10)"
      Height          =   315
      Index           =   10
      Left            =   3900
      TabIndex        =   22
      Top             =   3690
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(11)"
      Enabled         =   0   'False
      Height          =   315
      Index           =   11
      Left            =   1260
      TabIndex        =   25
      Top             =   4500
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(8)"
      Height          =   315
      Index           =   8
      Left            =   3900
      TabIndex        =   18
      Top             =   2880
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(7)"
      Height          =   315
      Index           =   7
      Left            =   1260
      TabIndex        =   17
      Top             =   2880
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(6)"
      Height          =   315
      Index           =   6
      Left            =   6510
      TabIndex        =   13
      Top             =   2160
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(5)"
      Height          =   315
      Index           =   5
      Left            =   3900
      TabIndex        =   12
      Top             =   2160
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(4)"
      Height          =   315
      Index           =   4
      Left            =   1260
      TabIndex        =   11
      Top             =   2160
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(3)"
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   9
      Top             =   1350
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(2)"
      Height          =   315
      Index           =   2
      Left            =   6510
      TabIndex        =   5
      Top             =   630
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(1)"
      Height          =   315
      Index           =   1
      Left            =   3900
      TabIndex        =   4
      Top             =   630
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(0)"
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   3
      Top             =   630
      Width           =   2505
   End
   Begin VB.CommandButton cmdPLF 
      Caption         =   "價目表"
      Height          =   375
      Left            =   5310
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdPLB 
      Caption         =   "公告公文記錄"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   90
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "外　商："
      Height          =   255
      Left            =   270
      TabIndex        =   30
      Top             =   3750
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "法務處："
      Height          =   255
      Left            =   270
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "商標處："
      Height          =   255
      Left            =   270
      TabIndex        =   28
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "專利處："
      Height          =   255
      Left            =   270
      TabIndex        =   27
      Top             =   690
      Width           =   915
   End
End
Attribute VB_Name = "frm210143"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已檢查 (無需修改的物件)
'Memo by Lydia 2019/07/01 表單名稱:價目表查詢=>價目表
'Create By Sindy 2014/3/6
Option Explicit

' 變數宣告區
Dim m_AttachPath As String
Private Declare Function SendMessageByNum Lib "user32" _
   Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
   wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim ii As Integer


'取得操作人員的價目表權限
Private Function GetPLFQueryLimit() As String
Dim Rs As New ADODB.Recordset
   
   GetPLFQueryLimit = ""
   If Left(Pub_StrUserSt15, 1) = "S" Or Left(Pub_StrUserSt15, 1) = "D" Then
      GetPLFQueryLimit = "01,02,03,04,05,06,07,08,09,10,11,12"
      Exit Function
   End If
   
   '讀取有權限的系統類別
   If Rs.State <> adStateClosed Then Rs.Close
   Rs.CursorLocation = adUseClient
    'Modified by Lydia 2015/11/30 排除CFP常辦國家年費(延展費)預估報價
'   rs.Open "Select plq01 From pricelistquery" & _
'           " Where instr(plq03,'" & strUserNum & "')>0 or instr(plq03,'" & Pub_StrUserSt03 & "')>0" & _
'           " group by plq01", _
'            cnnConnection, adOpenStatic, adLockReadOnly
   Rs.Open "Select plq01 From pricelistquery" & _
           " Where instr(plq03,'" & strUserNum & "')>0 or instr(plq03,'" & Pub_StrUserSt03 & "')>0 and PLQ01<>'13'" & _
           " group by plq01", _
            cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      Rs.MoveFirst
      Do While Not Rs.EOF
         GetPLFQueryLimit = GetPLFQueryLimit & "," & Rs.Fields(0).Value
         Rs.MoveNext
      Loop
      GetPLFQueryLimit = Mid(GetPLFQueryLimit, 2)
   End If
   Set Rs = Nothing
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

'公告公文記錄
Private Sub cmdPLB_Click()
   If ChkPLF(0).Enabled = False _
      And ChkPLF(1).Enabled = False _
      And ChkPLF(2).Enabled = False _
      And ChkPLF(3).Enabled = False _
      And ChkPLF(4).Enabled = False _
      And ChkPLF(5).Enabled = False _
      And ChkPLF(6).Enabled = False _
      And ChkPLF(7).Enabled = False _
      And ChkPLF(8).Enabled = False _
      And ChkPLF(9).Enabled = False _
      And ChkPLF(10).Enabled = False Then
'      And ChkPLF(11).Enabled = False Then  'cancel by sonia 2022/11/11 杜協理通知關閉
      MsgBox "無使用權限！", vbExclamation
      Exit Sub
   End If
   
   If ChkPLF(0).Value = 0 _
      And ChkPLF(1).Value = 0 _
      And ChkPLF(2).Value = 0 _
      And ChkPLF(3).Value = 0 _
      And ChkPLF(4).Value = 0 _
      And ChkPLF(5).Value = 0 _
      And ChkPLF(6).Value = 0 _
      And ChkPLF(7).Value = 0 _
      And ChkPLF(8).Value = 0 _
      And ChkPLF(9).Value = 0 _
      And ChkPLF(10).Value = 0 _
      And ChkPLF2(0).Value = 0 _
      And ChkPLF2(1).Value = 0 _
      And ChkPLF2(2).Value = 0 _
      And ChkPLF2(3).Value = 0 _
      And ChkPLF2(4).Value = 0 _
      And ChkPLF2(5).Value = 0 _
      And ChkPLF2(6).Value = 0 _
      And ChkPLF2(7).Value = 0 _
      And ChkPLF2(8).Value = 0 _
      And ChkPLF2(9).Value = 0 _
      And ChkPLF2(10).Value = 0 Then
'      And ChkPLF(11).Value = 0 _      'cancel by sonia 2022/11/11 杜協理通知關閉
'      And ChkPLF2(11).Value = 0 Then  'cancel by sonia 2022/11/11 杜協理通知關閉
      MsgBox "請至少點選一項價目表！", vbExclamation
      Exit Sub
   End If
   
   For ii = 0 To 10     'modify by sonia 2022/11/11 杜協理通知關閉法務處 11->10
      frm210143_1.ChkPLF(ii).Caption = Me.ChkPLF(ii).Caption
      frm210143_1.ChkPLF(ii).Enabled = Me.ChkPLF(ii).Enabled
      frm210143_1.ChkPLF(ii).Value = Me.ChkPLF(ii).Value
      If Me.ChkPLF2(ii).Visible = True And Me.ChkPLF2(ii).Value = 1 Then
         frm210143_1.ChkPLF(ii).Value = 1
      End If
   Next ii
   frm210143_1.QueryData
   'If frm210143_1.QueryData = True Then
      frm210143_1.Show vbModal
   'End If
End Sub

'最新價目表
Private Sub cmdPLF_Click()
Dim hLocalFile As Long
Dim stFileName As String
Dim strKEY01 As String, strKEY02 As String
Dim jj As Integer
Dim strTemp As String
   
   If ChkPLF(0).Enabled = False _
      And ChkPLF(1).Enabled = False _
      And ChkPLF(2).Enabled = False _
      And ChkPLF(3).Enabled = False _
      And ChkPLF(4).Enabled = False _
      And ChkPLF(5).Enabled = False _
      And ChkPLF(6).Enabled = False _
      And ChkPLF(7).Enabled = False _
      And ChkPLF(8).Enabled = False _
      And ChkPLF(9).Enabled = False _
      And ChkPLF(10).Enabled = False Then
'      And ChkPLF(11).Enabled = False Then    'cancel by sonia 2022/11/11 杜協理通知關閉
      MsgBox "無使用權限！", vbExclamation
      Exit Sub
   End If
   
   If ChkPLF(0).Value = 0 _
      And ChkPLF(1).Value = 0 _
      And ChkPLF(2).Value = 0 _
      And ChkPLF(3).Value = 0 _
      And ChkPLF(4).Value = 0 _
      And ChkPLF(5).Value = 0 _
      And ChkPLF(6).Value = 0 _
      And ChkPLF(7).Value = 0 _
      And ChkPLF(8).Value = 0 _
      And ChkPLF(9).Value = 0 _
      And ChkPLF(10).Value = 0 _
      And ChkPLF2(0).Value = 0 _
      And ChkPLF2(1).Value = 0 _
      And ChkPLF2(2).Value = 0 _
      And ChkPLF2(3).Value = 0 _
      And ChkPLF2(4).Value = 0 _
      And ChkPLF2(5).Value = 0 _
      And ChkPLF2(6).Value = 0 _
      And ChkPLF2(7).Value = 0 _
      And ChkPLF2(8).Value = 0 _
      And ChkPLF2(9).Value = 0 _
      And ChkPLF2(10).Value = 0 Then
 '     And ChkPLF(11).Value = 0 _        'cancel by sonia 2022/11/11 杜協理通知關閉法務處
 '     And ChkPLF2(11).Value = 0 Then    'cancel by sonia 2022/11/11 杜協理通知關閉法務處
      MsgBox "請至少點選一項價目表！", vbExclamation
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For ii = 0 To 10    'modify by sonia 2022/11/11 杜協理通知關閉法務處 11->10
      If ChkPLF(ii).Enabled = True Then
         For jj = 1 To 2
            If jj = 1 Then
               '最新價目表
               If ChkPLF(ii).Value = 1 Then
                  strTemp = ChkPLF(ii).Caption
                  strKEY01 = Format(ii + 1, "00")
                  strKEY02 = Mid(ChkPLF(ii).Caption, InStr(ChkPLF(ii).Caption, "(") + 1, 9)
               Else
                  GoTo ReadNext_jj
               End If
            Else
               '前一版價目表
               If ChkPLF2(ii).Visible = True And ChkPLF2(ii).Value = 1 Then
                  strTemp = ChkPLF2(ii).Caption
                  strKEY01 = Format(ii + 1, "00")
                  strKEY02 = Mid(ChkPLF2(ii).Caption, InStr(ChkPLF2(ii).Caption, "(") + 1, 9)
               Else
                  GoTo ReadNext_jj
               End If
            End If
            stFileName = SetFileName(strKEY01, DBDATE(strKEY02))
            If InStrRev(strTemp, "-") > 0 Then
               stFileName = stFileName & Mid(strTemp, InStrRev(strTemp, "-"))
            End If
            stFileName = "$$" & stFileName & ServerTime & ".pdf"
            If GetAttachFile(stFileName, strKEY01, strKEY02) = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
ReadNext_jj:
         Next jj
      End If
   Next ii
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_AttachPath = App.path '& Pub_GetSpecMan("EmpFlowAttPath")
   
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   MenuEnabled
   Set frm210143 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\$$*.pdf"
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim rsQuery As ADODB.Recordset
Dim strSql As String
Dim strQueryLimit As String
Dim intR As Integer
   
   ClearField
   strQueryLimit = GetPLFQueryLimit
   If strQueryLimit = "" Then Exit Sub
   strQueryLimit = Replace(strQueryLimit, ",", "','")
   'Modified by Lydia 2015/11/30 排除CFP常辦國家年費(延展費)預估報價
'   strSql = "SELECT PLF01,max(PLF02) FROM pricelistfile" & _
'            " WHERE PLF01 in('" & strQueryLimit & "')" & _
'            " group by PLF01"
   strSql = "SELECT PLF01,max(PLF02) FROM pricelistfile" & _
            " WHERE PLF01 in('" & strQueryLimit & "') and PLF01<>'13'" & _
            " group by PLF01"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '***** 最新價目表 *****
         ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption = Mid(ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption, 1, InStr(ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption, "(") - 1) & "(" & CFDate(TransDate(rsTmp.Fields(1), 1)) & ")"
         ChkPLF(CInt(rsTmp.Fields(0)) - 1).Enabled = True
         '加版本數
'         strSql = "SELECT count(*) FROM pricelistbulletin" & _
'                  " WHERE PLB01='" & rsTmp.Fields(0) & "' and PLB02>" & rsTmp.Fields(1)
         strSql = "SELECT count(*) FROM pricelistbulletin" & _
                  " WHERE PLB01='" & rsTmp.Fields(0) & "' and PLB12=" & rsTmp.Fields(1)
         intR = 1
         Set RsTemp = ClsLawReadRstMsg(intR, strSql)
         If intR = 1 Then
            If Val(rsTmp.Fields(1)) >= 20140401 Then
               If (RsTemp.Fields(0) - 1) > 0 Then
                  ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption = ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption & "-" & (RsTemp.Fields(0) - 1)
               End If
            Else
               If RsTemp.Fields(0) > 0 Then
                  ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption = ChkPLF(CInt(rsTmp.Fields(0)) - 1).Caption & "-" & RsTemp.Fields(0)
               End If
            End If
         End If
         
         '***** 前一版 *****
         strSql = "SELECT PLF01,PLF02 FROM pricelistfile" & _
                  " WHERE PLF01='" & rsTmp.Fields(0) & "' and PLF02<" & rsTmp.Fields(1) & _
                  " order by PLF02 desc"
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, strSql)
         If intR = 1 Then
            rsQuery.MoveFirst
            ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption = Mid(ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption, 1, InStr(ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption, "(") - 1) & "(" & CFDate(TransDate(rsQuery.Fields(1), 1)) & ")"
'            'Modify By Sindy 2020/7/16 夏慧珠說馬德里隱藏
'            If CInt(rsTmp.Fields(0)) <> 7 Then
               ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Visible = True
'            End If
            
            '加版本數
'            strSql = "SELECT count(*) FROM pricelistbulletin" & _
'                     " WHERE PLB01='" & rsTmp.Fields(0) & "' and PLB02<" & rsTmp.Fields(1) & " and PLB02>" & rsQuery.Fields(1)
            strSql = "SELECT count(*) FROM pricelistbulletin" & _
                     " WHERE PLB01='" & rsTmp.Fields(0) & "' and PLB12=" & rsQuery.Fields(1)
            intR = 1
            Set RsTemp = ClsLawReadRstMsg(intR, strSql)
            If intR = 1 Then
               If Val(rsQuery.Fields(1)) >= 20140401 Then
                  If (RsTemp.Fields(0) - 1) > 0 Then
                     ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption = ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption & "-" & (RsTemp.Fields(0) - 1)
                  End If
               Else
                  If RsTemp.Fields(0) > 0 Then
                     ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption = ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Caption & "-" & RsTemp.Fields(0)
                  End If
               End If
            End If
         End If
         
         'Add By Sindy 2024/8/5 下架日期,時間到再下架
         strSql = "SELECT PLF04 FROM pricelistfile" & _
                  " WHERE PLF01='" & rsTmp.Fields(0) & "'" & _
                  " AND nvl(PLF04,0)>0 AND nvl(PLF04,0)<=" & strSrvDate(1)
         intR = 1
         Set RsTemp = ClsLawReadRstMsg(intR, strSql)
         If intR = 1 Then
            ChkPLF(CInt(rsTmp.Fields(0)) - 1).Enabled = False
            ChkPLF2(CInt(rsTmp.Fields(0)) - 1).Enabled = False
         End If
         '2024/8/5 END
         
         rsTmp.MoveNext
      Loop
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
   Set rsQuery = Nothing
End Sub

Private Sub ClearField()
   ChkPLF(0).Caption = "國內專利(  /  /  )"
   ChkPLF(1).Caption = "大陸專利(  /  /  )"
   ChkPLF(2).Caption = "香港澳門專利(  /  /  )"
   ChkPLF(3).Caption = "CFP(  /  /  )"
   ChkPLF(4).Caption = "國內商標(  /  /  )"
   ChkPLF(5).Caption = "大陸商標(  /  /  )"
   ChkPLF(6).Caption = "馬德里商標(  /  /  )"
   ChkPLF(7).Caption = "國內著作權(  /  /  )"
   ChkPLF(8).Caption = "大陸著作權(  /  /  )"
   ChkPLF(9).Caption = "CFT(  /  /  )"
   ChkPLF(10).Caption = "美國著作權(  /  /  )"
'   ChkPLF(11).Caption = "顧問及法務(  /  /  )"   'cancel by sonia 2022/11/11 杜協理通知關閉
   For ii = 0 To 10   'modify by sonia 2022/11/11 杜協理通知關閉法務處 11->10
      ChkPLF(ii).Enabled = False
      ChkPLF(ii).Value = 0
      ChkPLF2(ii).Caption = "前一版(  /  /  )"
      ChkPLF2(ii).Visible = False '***
      ChkPLF2(ii).Value = 0
   Next ii
End Sub

Private Function GetAttachFile(ByRef pFileName As String, ByVal strKEY01 As String, ByVal strKEY02 As String, _
                               Optional pSavePath As String, Optional pFileSize As Integer = 0) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   strExc(0) = "select * from pricelistfile where PLF01='" & strKEY01 & "' and PLF02=" & Val(DBDATE(strKEY02))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pSavePath = "" Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
         stAttPath = m_AttachPath & "\" & pFileName
         '檔案已存在時
         If Dir(stAttPath) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stAttPath) = True Then
               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            Kill stAttPath
         End If
      Else
         stAttPath = pSavePath
      End If
      
      If Dir(stAttPath) <> "" Then Kill stAttPath
      
      'Add By Sindy 2017/5/31
      If "" & RsTemp.Fields("plf11") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("plf11"), stAttPath, UCase("PRICELISTFILE"))
      Else
      '2017/5/31 END
         With RsTemp
            lngSize = Val(.Fields("PLF03").Value)
            ReDim bytes(lngSize)
            If lngSize > 0 Then bytes() = .Fields("PLF04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      pFileName = stAttPath
      If pFileSize = 1 Then
         pFileName = pFileName & " (" & Round(RsTemp.Fields("PLF03") / 1024, 2) & " KB)"
      End If
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'檔案名稱：啟用日期(民國日期)＋系統類別中文名稱＋價目表.pdf
Private Function SetFileName(ByVal strSysKind As String, ByVal strDate As String) As String
   Select Case strSysKind
      Case "01"
         strSysKind = "國內專利"
      Case "02"
         strSysKind = "大陸專利"
      Case "03"
         strSysKind = "香港澳門專利"
      Case "04"
         strSysKind = "CFP"
      Case "05"
         strSysKind = "國內商標"
      Case "06"
         strSysKind = "大陸商標"
      Case "07"
         strSysKind = "馬德里商標"
      Case "08"
         strSysKind = "國內著作權"
      Case "09"
         strSysKind = "大陸著作權"
      Case "10"
         strSysKind = "CFT"
      Case "11"
         strSysKind = "美國著作權"
'      Case "12"                      'cancel by sonia 2022/11/11 杜協理通知關閉
'         strSysKind = "顧問及法務"   'cancel by sonia 2022/11/11 杜協理通知關閉
   End Select
   SetFileName = TransDate(strDate, 1) & strSysKind & "價目表"
End Function
