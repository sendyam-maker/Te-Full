VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04060103_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報查詢列印"
   ClientHeight    =   5745
   ClientLeft      =   15
   ClientTop       =   960
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4335
      Left            =   150
      TabIndex        =   4
      Top             =   990
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7646
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
   Begin VB.CommandButton bottonExit 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   7848
      TabIndex        =   0
      Top             =   144
      Width           =   1332
   End
   Begin VB.TextBox text01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   660
      Width           =   1572
   End
   Begin VB.Label labelCount 
      Caption         =   "發明合計:        新型合計:        設定合計:          合計:"
      Height          =   252
      Left            =   180
      TabIndex        =   3
      Top             =   5400
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "起始公告號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   1212
   End
End
Attribute VB_Name = "frm04060103_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
' 代理人
Dim m_DataKey1 As String
' 起始公告號
Dim m_DataKey2 As String
' 開始公告日
Dim m_DataKey3 As String
' 結束公告日
Dim m_DataKey4 As String
Dim m_ListDetail As Boolean
Dim m_Amount(3) As Integer
Dim m_Recordset As ADODB.Recordset
'
Dim m_CurrSel As Integer

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub bottonExit_Click()
   Unload Me
   frm04060103_1.Show
   frm04060103_1.SetInputEntry
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Recordset = Nothing
   frm04060103_1.Show
End Sub

Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & strNation & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly 'edit by nickc 2007/02/06 , adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetNation = rsTmp.Fields("NA03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

Public Sub UpdateCtrlData()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim Fld1, Fld2, Fld3, Fld4, Fld5, Fld6, Fld7, Fld8 As String
   Dim strTmp1 As String
   Dim strTmp2 As String
   Dim nRow, nCol
   Dim nIndex
   
   Set m_Recordset = New ADODB.Recordset
   
   text01.BackColor = &H8000000F
   text01 = m_DataKey2

   For nIndex = 0 To 2
      m_Amount(nIndex) = 0
   Next nIndex
   
   strSubSQL = Empty
   strSql = "SELECT TPB01,TPB02,NVL(TPB03 - 19110000, NULL) AS TPB03, " & _
                   "TPB04 || '-' || TPB05, NA03, TA03, PA01 || '-' || PA02 || '-' || PA03 || '-' || PA04, TA04 FROM TPBulletin, NATION, PATENT, TAGENT "
   If IsEmpty(m_DataKey3) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB03 >= " & ChangeTStringToWString(m_DataKey3) & " "
   End If
   If IsEmpty(m_DataKey4) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB03 <= " & ChangeTStringToWString(m_DataKey4) & " "
   End If
   If IsEmpty(m_DataKey3) = False Or IsEmpty(m_DataKey4) = False Then
      pub_QL05 = pub_QL05 & ";" & frm04060103_1.Label3 & frm04060103_1.text03_01 & "-" & frm04060103_1.text03_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(m_DataKey1) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB07 = '" & m_DataKey1 & "' "
      pub_QL05 = pub_QL05 & ";" & frm04060103_1.Label1 & frm04060103_1.text01_01 & frm04060103_1.text01_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(m_DataKey2) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB02 >= '" & m_DataKey2 & "' "
      pub_QL05 = pub_QL05 & ";" & frm04060103_1.Label5 & frm04060103_1.text05 'Add By Sindy 2010/12/2
   End If
   If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
'   strSubSQL = strSubSQL & "TPB06 = NA01(+) AND " & _
'                           "TPB07 = TA02(+) AND " & _
'                           "TPB01 = PA11(+) AND " & _
'                           "'P' = TA01(+) "
   strSubSQL = strSubSQL & "TPB06 = NA01(+) AND " & _
                           "TPB07 = TA02(+) AND " & _
                           "'P' = TA01(+) "
   'Modify By Sindy 2011/12/16
   If frm04060103_1.Check1 = 1 Then
      strSubSQL = strSubSQL & " AND TPB01 = PA11 "
      pub_QL05 = pub_QL05 & ";只顯示本所案件"
   Else
      strSubSQL = strSubSQL & " AND TPB01 = PA11(+) "
   End If
   '2011/12/16 End
   If strSubSQL <> Empty Then
      strSql = strSql & "WHERE " & strSubSQL
   End If
   strSql = strSql & "ORDER BY TPB04, TPB05, TPB02 ASC"
   
   m_Recordset.CursorLocation = adUseClient
   m_Recordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   grdList.Clear
   InitialGridList
   nRow = 1
   If m_Recordset.RecordCount > 0 Then
      InsertQueryLog (m_Recordset.RecordCount) 'Add By Sindy 2010/12/2
      m_Recordset.MoveFirst
      For nIndex = 0 To m_Recordset.RecordCount - 1
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
         If IsNull(m_Recordset.Fields(0)) = False Then
            grdList.TextMatrix(nRow, 1) = m_Recordset.Fields(0)
         End If
         If IsNull(m_Recordset.Fields(1)) = False Then
            grdList.TextMatrix(nRow, 2) = m_Recordset.Fields(1)
         End If
         If IsNull(m_Recordset.Fields(2)) = False Then
            grdList.TextMatrix(nRow, 3) = m_Recordset.Fields(2)
         End If
         If IsNull(m_Recordset.Fields(3)) = False Then
            grdList.TextMatrix(nRow, 4) = m_Recordset.Fields(3)
         End If
         If IsNull(m_Recordset.Fields(4)) = False Then
            grdList.TextMatrix(nRow, 5) = m_Recordset.Fields(4)
         End If
         If IsNull(m_Recordset.Fields(5)) = False Then
            grdList.TextMatrix(nRow, 6) = m_Recordset.Fields(5)
         End If
         If IsNull(m_Recordset.Fields(6)) = False Then
            If m_Recordset.Fields(6) <> "---" Then
               grdList.TextMatrix(nRow, 7) = m_Recordset.Fields(6)
            End If
         End If
         If IsNull(m_Recordset.Fields(7)) = False Then
            grdList.TextMatrix(nRow, 8) = m_Recordset.Fields(7)
         End If
         ' 計算 發明, 新型, 設計的總數
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'Select Case Mid(m_Recordset.Fields(0), 3, 1)
         Select Case Mid(m_Recordset.Fields(0), 4, 1)
            Case "1":
               m_Amount(0) = m_Amount(0) + 1
            Case "2":
               m_Amount(1) = m_Amount(1) + 1
            Case "3":
               m_Amount(2) = m_Amount(2) + 1
         End Select
         m_Recordset.MoveNext
      Next nIndex
      
      grdList.FixedRows = 1 'Added by Lydia 2022/02/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
   End If
   
   labelCount = "發明合計 : " & m_Amount(0) & "     " & _
                "新型合計 : " & m_Amount(1) & "     " & _
                "設計合計 : " & m_Amount(2) & "     " & _
                "合計 : " & m_Amount(0) + m_Amount(1) + m_Amount(2)
   
   Set rsTmp = Nothing
   m_Recordset.Close
   Set m_Recordset = Nothing
End Sub

Public Sub SetData(ByVal strKey1 As String, ByVal StrKey2 As String, ByVal strKey3 As String, ByVal strKey4 As String, ByVal bListDetail As Boolean)
   m_DataKey1 = strKey1
   m_DataKey2 = StrKey2
   m_DataKey3 = strKey3
   m_DataKey4 = strKey4
   m_ListDetail = bListDetail
End Sub

Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 9
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "申請案號"
   grdList.ColWidth(1) = 1200
   grdList.ColAlignment(1) = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "公告號"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "公告日"
   grdList.ColWidth(3) = 800
   grdList.ColAlignment(3) = flexAlignCenterCenter
   grdList.col = 4
   grdList.Text = "卷期"
   grdList.ColWidth(4) = 600
   grdList.ColAlignment(4) = flexAlignCenterCenter
   grdList.col = 5
   grdList.Text = "申請人國籍"
   grdList.ColWidth(5) = 1000
   grdList.ColAlignment(5) = flexAlignLeftCenter
   grdList.col = 6
   grdList.Text = "代理人"
   grdList.ColWidth(6) = 1000
   grdList.ColAlignment(6) = flexAlignLeftCenter
   grdList.col = 7
   grdList.Text = "本所案號"
   grdList.ColWidth(7) = 1400
   grdList.ColAlignment(7) = flexAlignLeftCenter
   grdList.col = 8
   grdList.Text = "事務所名稱"
   grdList.ColWidth(8) = 1200
   grdList.ColAlignment(8) = flexAlignLeftCenter
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

