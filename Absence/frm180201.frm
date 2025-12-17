VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180201 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核作業"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選(&A)"
      Height          =   360
      Index           =   0
      Left            =   5190
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Index           =   1
      Left            =   6030
      TabIndex        =   3
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "簽核(&O)"
      Height          =   360
      Index           =   2
      Left            =   7230
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   3
      Left            =   8055
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180201.frx":0000
      Height          =   5235
      Left            =   60
      TabIndex        =   2
      Top             =   450
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   9243
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
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
End
Attribute VB_Name = "frm180201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/5
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵


Public Sub PubShowNextData()
   Select Case cmdState
      Case 0 '全選
         GRD1.Visible = False
         If GRD1.Rows > 1 Then
            If GRD1.TextMatrix(1, 1) <> "" Then
               For j = 1 To GRD1.Rows - 1
                  GRD1.col = 0
                  GRD1.row = j
                  GRD1.Text = "V"
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = &HFFC0C0
                  Next i
               Next j
            End If
         End If
         GRD1.Visible = True
      Case 1 '查詢
         If QueryData = False Then ShowNoData
      Case 2 '簽核
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.col = 0
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
                GRD1.col = 4
                If Not IsNull(GRD1.Text) Then
                   Screen.MousePointer = vbHourglass
                   Me.Hide
                   frm180201_01.txtB1001 = Pub_RplStr(GRD1.Text)
                   frm180201_01.QueryData
                   frm180201_01.Show
                   Screen.MousePointer = vbDefault
                   Me.Enabled = True
                   Exit Sub
                End If
            End If
         Next i
         Me.Enabled = True
         Call QueryData
      Case 3 '結束
         Unload Me
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   
   PubShowNextData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2023/12/22
   'Modify By Sindy 2025/2/11 薛經理提,修正[假單]的順序.請以小到大排列及依序進行簽核
'   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "Select ' ' as V,a1.A0922 部門別,s1.ST01 員工代號,s1.ST02 姓名,B1001 表單編號," & B1002CName & " 表單類別,AC03 假別," & _
               "sqldateT(B1004)||' '||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) 起始日期時間," & _
               "sqldateT(decode(B1002,'02',B1004,B1006))||' '||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) 結束日期時間," & _
               "B1009 天數,decode(B1002,'02',nvl(B1012,B1013),B1010) 時數,nvl(s2.ST02,a2.A0922) 目前處理人員," & B1018CName & " 目前表單狀態 " & _
               "From ABS010,Staff s1,ACC090NEW a1,allcode,Staff s2,ACC090NEW a2 " & _
               "Where B1003<>B1017 and B1017='" & strUserNum & "' And B1019 Is Null and B1003=s1.ST01(+) and s1.ST93=a1.A0921(+) and B1017=a2.A0921(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) " & _
               "order by B1001 asc "
               '"order by B1001 desc "
               '"order by 2,3,5 "
'   Else
'   '2023/12/22 END
'      strSql = "Select ' ' as V,a1.A0902 部門別,s1.ST01 員工代號,s1.ST02 姓名,B1001 表單編號," & B1002CName & " 表單類別,AC03 假別," & _
'               "sqldateT(B1004)||' '||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) 起始日期時間," & _
'               "sqldateT(decode(B1002,'02',B1004,B1006))||' '||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) 結束日期時間," & _
'               "B1009 天數,decode(B1002,'02',nvl(B1012,B1013),B1010) 時數,nvl(s2.ST02,a2.A0902) 目前處理人員," & B1018CName & " 目前表單狀態 " & _
'               "From ABS010,Staff s1,ACC090 a1,allcode,Staff s2,ACC090 a2 " & _
'               "Where B1003<>B1017 and B1017='" & strUserNum & "' And B1019 Is Null and B1003=s1.ST01(+) and s1.ST03=a1.A0901(+) and B1017=a2.A0901(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) " & _
'               "order by B1001 desc "
'               '"order by 2,3,5 "
'   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   If rsTmp.RecordCount > 0 Then
      GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
Dim strData As String
Dim strTemp As Variant
   
   MoveFormToCenter Me
   
   'Add By Sindy 2014/2/18
   '檢查是否有F.打卡異常主管處理待確認
   strData = ChkIsAbsenceMustPro
   strTemp = Split(strData, ",")
   For i = 0 To UBound(strTemp)
      If strTemp(i) = "F" Then
         MsgBox "您尚有打卡異常主管待確認的資料！" & vbCrLf & vbCrLf & _
                "請進入「打卡異常主管處理」作業，進行處理。"
      End If
   Next i
   '2014/2/18 END
   
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strText As String
   
   'Me.Form=A
   '一進入系統,檢查是否有須要開啟此作業
   If pub_CallNextABSForm = True Then
      strText = ChkIsAbsenceMustPro
      Me.Hide
      If InStr(1, strText, "B") > 0 Then
         frm180101.Show
      ElseIf InStr(1, strText, "C") > 0 Then
         frm180203_1.Show
      ElseIf InStr(1, strText, "D") > 0 Then
         frm160102.intChoose = 1
         frm160102.Hide
         Call frm160102.cmdok_Click(0)
      'Add By Sindy 2015/7/2
      ElseIf InStr(1, strText, "G") > 0 Then
         If TypeName(Tmpfrm210148) <> "Nothing" Then
            Tmpfrm210148.Show
         End If
      ElseIf InStr(1, strText, "H") > 0 Then
         If TypeName(Tmpfrm210147) <> "Nothing" Then
            Tmpfrm210147.Show
         End If
      '2015/7/2 END
      Else
         pub_CallNextABSForm = False
      End If
   End If
   
   Set frm180201 = Nothing
   If pub_CallNextABSForm = False Then
      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "部門別", "員工代號", "姓名", "表單編號", "表單類別", "假別", "起始日期時間", "結束日期時間", "天數", "時數", "目前處理人員", "目前表單狀態")
   arrGridHeadWidth = Array(200, 0, 800, 800, 900, 800, 700, 1300, 1300, 600, 600, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "表單編號" Or Me.GRD1.Text = "天數" Or Me.GRD1.Text = "時數" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
