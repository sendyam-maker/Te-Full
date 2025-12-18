VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040109_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "一案兩申請案件資料維護"
   ClientHeight    =   5760
   ClientLeft      =   135
   ClientTop       =   1050
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7128
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtChoose 
      Height          =   270
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   372
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4272
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7541
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2250
      X2              =   2370
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label lblCustNo 
      Height          =   255
      Index           =   1
      Left            =   2610
      TabIndex        =   11
      Top             =   720
      Width           =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6960
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "客戶代碼："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "功能代號：           (2.修改  4.刪除  5.查詢 )"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   3372
   End
   Begin VB.Label lblCustNo 
      Height          =   255
      Index           =   0
      Left            =   1125
      TabIndex        =   8
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   0
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   972
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   6
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Index           =   1
      Left            =   4905
      TabIndex        =   5
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm040109_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/21 改成Form2.0 (grdDataList)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Public m_CM10 As String 'Added by Morgan 2015/9/10
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim intNowRow As Integer

Select Case Index
   Case 0
         'Added by Morgan 2012/2/9
         With frm040109_1
         If Not ((txtChoose = "2" And .m_bUpdate) Or (txtChoose = "4" And .m_bDelete) Or (txtChoose = "5" And .m_bQuery)) Then
            MsgBox "無權限!!", vbExclamation
            Exit Sub
         End If
         End With
         'end 2012/2/9
   
           If grdDataList.Rows > 1 Then
              intNowRow = grdDataList.row
              frm040109_2.strCode1 = grdDataList.TextMatrix(intNowRow, 0)
              frm040109_2.strCode2 = grdDataList.TextMatrix(intNowRow, 1)
              frm040109_2.strCode3 = grdDataList.TextMatrix(intNowRow, 2)
              frm040109_2.strCode4 = grdDataList.TextMatrix(intNowRow, 3)
              frm040109_2.strCode5 = grdDataList.TextMatrix(intNowRow, 8)
              frm040109_2.strCode6 = grdDataList.TextMatrix(intNowRow, 9)
              frm040109_2.strCode7 = grdDataList.TextMatrix(intNowRow, 10)
              frm040109_2.strCode8 = grdDataList.TextMatrix(intNowRow, 11)
              frm040109_2.intChoose = Val(txtChoose)
              frm040109_2.m_CM10 = Me.m_CM10 'Added by Morgan 2015/9/10
              frm040109_2.Caption = Me.Caption 'Added by Morgan 2015/9/10
              Set frm040109_2.frmParent = Me
              frm040109_2.Show
              Me.Hide
           Else
              MsgBox "資料庫無資料 !", vbInformation
           End If
   Case 1
              intLeaveKind = 1
              Unload Me
   Case 2
              intLeaveKind = 0
              Unload Me
End Select
End Sub
Private Sub Form_Activate()
   Dim varSaveCursor As Variant
   
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If bReadCaseRelationRst = True Then
      SetDataListWidth
      intLastRow = 0
      If grdDataList.Rows > 1 Then
         ShowBar grdDataList, intLastRow, 15
      End If
      txtChoose = "5"
      txtChoose.SetFocus
   End If
   Screen.MousePointer = varSaveCursor
   
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If intLeaveKind = 1 Then
      frm040109_1.Show
   Else
      Unload frm040109_1
   End If
   intLeaveKind = 0
   'Add By Cheng 2002/07/18
   Set frm040109_3 = Nothing
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 15
      blnOKtoShow = True
   End If
End If
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(400, 650, 200, 250, 2100, 750, 750, 650, 400, 650, 200, 250, 2100, 750, 750, 650)
SetGridDataListWidth grdDataList, varGridWidth()
SetDataListVision grdDataList, , True
blnOKtoShow = True
End Sub

Private Sub txtChoose_GotFocus()
txtChoose.SelStart = 0
txtChoose.SelLength = Len(txtChoose)
End Sub

'Added by Morgan 2012/2/9
Private Sub txtChoose_KeyPress(KeyAscii As Integer)
   With frm040109_1
   If Not (KeyAscii = 8 Or (Chr(KeyAscii) = "2" And .m_bUpdate) Or (Chr(KeyAscii) = "4" And .m_bDelete) Or (Chr(KeyAscii) = "5" And .m_bQuery)) Then
      KeyAscii = 0
      Beep
   End If
   End With
End Sub

Private Sub txtChoose_Validate(Cancel As Boolean)
If Val(txtChoose) <> 2 And Val(txtChoose) <> 4 And Val(txtChoose) <> 5 Then
   ShowMsg MsgText(9198)
   txtChoose_GotFocus
   Cancel = True
End If
End Sub

Private Function bReadCaseRelationRst() As Boolean

   Dim stCon As String, stCustCon(1 To 5) As String
   
On Error GoTo ErrHnd

   'Modified by Morgan 2013/7/9 開放FCP也可使用加控制可用專利系統別
   'stCon = ""
   stCon = " and cm01 in (" & GetAddStr(Systemkind_g_P) & ")"
   'end 2013/7/9
   
   If lblCustNo(0) <> "" And lblCustNo(1) <> "" Then
      stCon = stCon & " AND ( ( PA1.PA26>='" & lblCustNo(0) & "' and PA1.PA26<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA1.PA27>='" & lblCustNo(0) & "' and PA1.PA27<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA1.PA28>='" & lblCustNo(0) & "' and PA1.PA28<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA1.PA29>='" & lblCustNo(0) & "' and PA1.PA29<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA1.PA30>='" & lblCustNo(0) & "' and PA1.PA30<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA2.PA26>='" & lblCustNo(0) & "' and PA2.PA26<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA2.PA27>='" & lblCustNo(0) & "' and PA2.PA27<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA2.PA28>='" & lblCustNo(0) & "' and PA2.PA28<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA2.PA29>='" & lblCustNo(0) & "' and PA2.PA29<='" & lblCustNo(1) & "')"
      stCon = stCon & " OR ( PA2.PA30>='" & lblCustNo(0) & "' and PA2.PA30<='" & lblCustNo(1) & "') )"
   End If
   If lblDate(0) <> "" Then
      stCon = stCon & " And CP1.CP05>=" & Format(Val(ChangeTDateStringToTString(lblDate(0))) + 19110000)
      stCon = stCon & " And CP2.CP05>=" & Format(Val(ChangeTDateStringToTString(lblDate(0))) + 19110000)
   End If
   If lblDate(1) <> "" Then
      stCon = stCon & " And CP1.CP05<=" & Format(Val(ChangeTDateStringToTString(lblDate(1))) + 19110000)
      stCon = stCon & " And CP2.CP05<=" & Format(Val(ChangeTDateStringToTString(lblDate(1))) + 19110000)
   End If
  'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If FMP2open = True Then
      strExc(0) = Replace(FMP2openSQL, "f0.CP", "PA1.PA")
      stCon = stCon & strExc(0)
      strExc(0) = Replace(FMP2openSQL, "f0.CP", "PA2.PA")
      stCon = stCon & strExc(0)
   End If
   strSql = "SELECT CM01 案件一,CM02 案件一,CM03 案件一,CM04 案件一, PA1.PA05 案件一名稱 , CP1.CP05-19110000 收文日, DECODE(CP1.CP27,NULL,NULL,CP1.CP27-19110000) 發文日, DECODE(PA1.PA16,'1','准','2','駁') 准駁" & _
         ",CM05 案件二,CM06 案件二,CM07 案件二,CM08 案件二, PA2.PA05 案件二名稱, CP2.CP05-19110000 收文日, DECODE(CP2.CP27,NULL,NULL,CP2.CP27-19110000) 發文日, DECODE(PA2.PA16,'1','准','2','駁') 准駁" & _
         " FROM CASEMAP, PATENT PA1, CASEPROGRESS CP1, PATENT PA2, CASEPROGRESS CP2" & _
         " WHERE CM10='" & IIf(m_CM10 <> "", m_CM10, "3") & "'" & _
         " AND CM01=PA1.PA01(+) AND CM02=PA1.PA02(+) AND CM03=PA1.PA03(+) AND CM04=PA1.PA04(+)" & _
         " AND CM01=CP1.CP01(+) AND CM02=CP1.CP02(+) AND CM03=CP1.CP03(+) AND CM04=CP1.CP04(+)" & _
         " AND CM05=PA2.PA01(+) AND CM06=PA2.PA02(+) AND CM07=PA2.PA03(+) AND CM08=PA2.PA04(+)" & _
         " AND CM05=CP2.CP01(+) AND CM06=CP2.CP02(+) AND CM07=CP2.CP03(+) AND CM08=CP2.CP04(+)" & _
         " AND CP1.CP10 IN ('101','102','103','104','105','125')" & _
         " AND CP2.CP10 IN ('101','102','103','104','105','125')" & stCon

   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList.Recordset = adoRecordset
   grdDataList.Refresh
   If adoRecordset.RecordCount > 0 Then
      bReadCaseRelationRst = True
   Else
      cmdOK(0).Visible = False
      MsgBox "資料庫無資料 !", vbInformation
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function
