VERSION 5.00
Begin VB.Form frm030203_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   1990
   ClientLeft      =   170
   ClientTop       =   1340
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1990
   ScaleWidth      =   4530
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   210
      TabIndex        =   8
      Top             =   1200
      Width           =   3915
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   4
      Top             =   840
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   3
      Top             =   840
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2580
      TabIndex        =   5
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3552
      TabIndex        =   6
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "注意：分割案申請案號請由核准進入"
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   540
      Width           =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "frm030203_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/11 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 總收文號
Dim m_CP09 As String
'Add By Cheng 2003/02/17
Dim SeekPrint As Integer, SeekPrintL As Integer
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   
Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2003/02/17
    '列印地址條
'move to unload by nick 2004/10/22
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
    Unload Me
End Sub

Private Sub Form_Load()
  
MoveFormToCenter Me

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   ' 本所案號的系統別
   If IsEmptyText(textTM01) = True Then
      strTit = "檢核資料"
      strMsg = "請先本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textTM02) = True Then
      strTit = "檢核資料"
      strMsg = "請先本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Public Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String

   ' 檢查欄位的資料是否都已經輸入正確
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   strTM01 = Trim(textTM01)
   strTM02 = Trim(textTM02)
   strTM03 = Trim(textTM03)
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   strTM04 = Trim(textTM04)
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
'edit by nickc 2006/10/18  分割案已經不從此處進入
'   'add by nick  2004/12/23
'   If CheckDC = True Then
'        frm030203_03.SetData strTM01, strTM02, strTM03, strTM04
'        Me.Hide
'        frm030203_03.Show
'        frm030203_03.QueryData
'        Exit Sub
'   End If
   
   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case textTM01
      Case "FCT":
         ' 讀取商標基本檔
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         rsTmp.Close
         
         ' 讀取案件進度檔(案件性質為申請)
'edit by nick 2004/12/23 加入分割等於申請
'         StrSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strTM01 & "' AND " & _
                        "CP02 = '" & strTM02 & "' AND " & _
                        "CP03 = '" & strTM03 & "' AND " & _
                        "CP04 = '" & strTM04 & "' AND " & _
                        "CP10 = '101' "
         'modify by sonia 2023/8/18 加入已發文條件 CP27>0
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strTM01 & "' AND " & _
                        "CP02 = '" & strTM02 & "' AND " & _
                        "CP03 = '" & strTM03 & "' AND " & _
                        "CP04 = '" & strTM04 & "' AND " & _
                        "cp10='101' and CP27>0 " 'edit by nickc 2006/10/18 不抓分割案 "CP10 in ('101','308') "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         m_CP09 = rsTmp.Fields("CP09")
         rsTmp.Close
      Case Else:
         GoTo EXITSUB
   End Select
         
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   ' 本所案號
   frm030203_02.SetData 0, Trim(textTM01), True
   frm030203_02.SetData 1, Trim(textTM02), False
   If IsEmptyText(textTM03) = True Then
      frm030203_02.SetData 2, "0", False
   Else
      frm030203_02.SetData 2, Trim(textTM03), False
   End If
   If IsEmptyText(textTM04) = True Then
      frm030203_02.SetData 3, "00", False
   Else
      frm030203_02.SetData 3, Trim(textTM04), False
   End If
   ' 總收文號
   frm030203_02.SetData 4, m_CP09, False
   Me.Hide
   frm030203_02.Show
   frm030203_02.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PUB_PrintAddressList strUserNum, Me.Combo1.Text
    '刪除地址條列表資料
    PUB_DeleteAddressList strUserNum
    '初始化序號
    pub_AddressListSN = 0
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    'Add By Cheng 2003/02/17
    '還原預設印表機
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    'Add By Cheng 2002/07/19
    Set frm030203_01 = Nothing
End Sub

' 本所案號中的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "FCT":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub

Public Sub Clear()
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
End Sub
Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub
'add by nick 2004/12/23 檢查是否有分割
Public Function CheckDC() As Boolean
CheckOC3
Dim strSql As String
strSql = "select count(*) from divisioncase where dc05='" & strTM01 & "' and dc06='" & strTM02 & "' and dc07='" & strTM03 & "' and dc08='" & strTM04 & "' "
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .Fields(0).Value <> 0 Then
        CheckDC = True
    Else
        CheckDC = False
    End If
End With
CheckOC3
End Function
