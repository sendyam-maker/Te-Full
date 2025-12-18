VERSION 5.00
Begin VB.Form frm030001 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF 指示信"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4125
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "CFT"
      Top             =   480
      Width           =   500
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   585
      Left            =   30
      TabIndex        =   6
      Top             =   810
      Width           =   4035
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   210
         Width           =   3150
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   3180
      MaxLength       =   2
      TabIndex        =   2
      Top             =   480
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2835
      MaxLength       =   1
      TabIndex        =   1
      Top             =   480
      Width           =   180
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1815
      MaxLength       =   6
      TabIndex        =   0
      Top             =   480
      Width           =   850
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   1
      Left            =   2910
      TabIndex        =   4
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   30
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   3210
      X2              =   1680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   150
      TabIndex        =   5
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm030001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/03 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Dim strTit As String
         Dim strMsg As String
         Dim nResponse
          If Trim(txt1(0).Text) = "" Then
              strTit = "資料檢核"
              strMsg = "本所案號中的系統別不正確"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              txt1(0).SetFocus
              Exit Sub
          End If
          If Trim(txt1(1).Text) = "" Then
              strTit = "資料檢核"
              strMsg = "本所案號中的案號不正確"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              txt1(1).SetFocus
              Exit Sub
          End If
          strSql = "select tm01 from trademark where tm01='" & txt1(0).Text & "' and tm02='" & txt1(1).Text & "' and tm03='" & IIf(Trim(txt1(2).Text) = "", "0", Trim(txt1(2).Text)) & "' and tm04='" & IIf(Trim(txt1(3).Text) = "", "00", Trim(txt1(3).Text)) & "' "
          strSql = strSql & " union select sp01 from servicepractice where sp01='" & txt1(0).Text & "' and sp02='" & txt1(1).Text & "' and sp03='" & IIf(Trim(txt1(2).Text) = "", "0", Trim(txt1(2).Text)) & "' and sp04='" & IIf(Trim(txt1(3).Text) = "", "00", Trim(txt1(3).Text)) & "' "
          CheckOC
          With adoRecordset
              .CursorLocation = adUseClient
              .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
              If .RecordCount <> 0 Then
                  Select Case Trim(txt1(0).Text)
                  Case "CFT"
                      frm030001_1.SetData 0, txt1(0)
                      frm030001_1.SetData 1, txt1(1)
                      frm030001_1.SetData 2, IIf(Trim(txt1(2)) = "", "0", txt1(2))
                      frm030001_1.SetData 3, IIf(Trim(txt1(3)) = "", "00", txt1(3))
                      If frm030001_1.QueryData = True Then
                          frm030001_1.Show
                          Me.Hide
                          
                           'Add By Sindy 2018/2/1 增加案件申請人地址視窗彈跳
                           frm020102_23.Hide
                           Set frm020102_23.UpForm = frm030001_1
                           frm020102_23.m_TM01 = txt1(0)
                           frm020102_23.m_TM02 = txt1(1)
                           frm020102_23.m_TM03 = IIf(Trim(txt1(2)) = "", "0", txt1(2))
                           frm020102_23.m_TM04 = IIf(Trim(txt1(3)) = "", "00", txt1(3))
                           'Me.Hide
                           frm020102_23.QueryData
                           frm020102_23.Show vbModal
                           '2018/2/1 End
                      Else
                          Unload frm030001_1
                          txt1(0).SetFocus
                      End If
                     
                  Case "CFC"
                      strTit = "資料檢核"
                      strMsg = "外商還沒給 CFC 的定稿格式！"
                      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                      txt1(0).SetFocus
                      Exit Sub
                  Case Else
                  End Select
              Else
                  strTit = "資料檢核"
                  strMsg = "本所案號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  txt1(0).SetFocus
                  Exit Sub
              End If
          End With
      Case 1
              Unload Me
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   
   SeekPrintL = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo1, m_OriPrinterName, False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm030001 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 2
          KeyAscii = UpperCase(KeyAscii)
      Case Else
          Select Case KeyAscii
          Case 48 To 57
          Case Else
                  KeyAscii = 0
          End Select
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

   Cancel = False
   Select Case Index
      Case 0
              If txt1(Index) <> "CFT" And txt1(Index) <> "CFC" And txt1(Index) <> "" Then
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "本所案號中的系統別不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  txt1(Index).SetFocus
              End If
      Case Else
   End Select
End Sub
