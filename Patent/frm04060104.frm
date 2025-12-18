VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060104 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內代理人名稱查詢"
   ClientHeight    =   5760
   ClientLeft      =   -4152
   ClientTop       =   1476
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9348
   Begin VB.CommandButton buttonQuery 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7536
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8352
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4635
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   9075
      _ExtentX        =   16002
      _ExtentY        =   8170
      _Version        =   393216
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
   Begin MSForms.TextBox text01 
      Height          =   330
      Left            =   1260
      TabIndex        =   0
      Top             =   630
      Width           =   2715
      VariousPropertyBits=   671105051
      MaxLength       =   12
      Size            =   "4789;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "代理人名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   1095
   End
End
Attribute VB_Name = "frm04060104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、text01
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim m_Recordset As New ADODB.Recordset
'
Dim m_CurrSel As Integer

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub

Private Sub buttonQuery_Click()
   ExecuteQuery
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (m_Recordset.State <> adStateClosed) Then
      m_Recordset.Close
   End If
   Set m_Recordset = Nothing
   'Add By Cheng 2002/07/18
   Set frm04060104 = Nothing
End Sub

Public Sub ExecuteQuery()
Dim strSql As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   ' 檢查rsRecordset的狀態
   If (m_Recordset.State <> adStateClosed) Then
      m_Recordset.Close
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label1 & text01 'Add By Sindy 2010/12/2
   If IsEmpty(text01) = False Then
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA03 LIKE '%" & text01 & "%' AND " & _
                     "TA01 = 'P'"
   Else
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA01 = 'P'"
   End If
   
   ' 查詢
   m_Recordset.CursorLocation = adUseClient
   m_Recordset.Open strSql, cnnConnection, adOpenDynamic
   
   ' 檢查是否有資料傳回來
   If m_Recordset.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      InitialGridList
      strTit = "查詢"
      strMsg = "資料庫中沒有符合的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      InsertQueryLog (m_Recordset.RecordCount) 'Add By Sindy 2010/12/2
      UpdateGridList
   End If
End Sub

Private Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 5
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "代理人編號"
   grdList.ColWidth(1) = 1000
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "代理人名稱"
   grdList.ColWidth(2) = 1200
   grdList.ColAlignment(2) = flexAlignLeftCenter
   grdList.col = 3
   grdList.Text = "事務所名稱"
   grdList.ColWidth(3) = 1200
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "建檔時公告日"
   grdList.ColWidth(4) = 1400
   grdList.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub UpdateGridList()
   Dim strTA01, strTA02 As String
   Dim nRow As Integer
   
   grdList.Clear
   InitialGridList
   
   If m_Recordset.RecordCount > 0 Then
      grdList.Rows = m_Recordset.RecordCount + 1
      m_Recordset.MoveFirst
      nRow = 1
      While m_Recordset.EOF <> True
         grdList.row = nRow
         
         grdList.col = 1
         If IsNull(m_Recordset.Fields("TA02")) = False Then
            grdList.Text = m_Recordset.Fields("TA02")
         End If
         
         grdList.col = 2
         If IsNull(m_Recordset.Fields("TA03")) = False Then
            grdList.Text = m_Recordset.Fields("TA03")
         End If
         
         grdList.col = 3
         If IsNull(m_Recordset.Fields("TA04")) = False Then
            grdList.Text = m_Recordset.Fields("TA04")
         End If
         
         grdList.col = 4
         If IsNull(m_Recordset.Fields("TA05")) = False Then
            grdList.Text = ChangeWStringToTString(m_Recordset.Fields("TA05"))
         End If
         
         nRow = nRow + 1
         m_Recordset.MoveNext
      Wend
      'Added by Lydia 2022/02/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If grdList.Rows > 1 Then
         grdList.FixedRows = 1
      End If
   End If
End Sub

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub text01_GotFocus()
  TextInverse text01
  'add by nickc 2007/07/13 將輸入法改成使用API
  OpenIme
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub text01_Validate(Cancel As Boolean)
CloseIme
End Sub
