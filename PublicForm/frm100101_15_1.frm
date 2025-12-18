VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_15_1 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "往來記錄瀏覽區 - 搬移檔案"
   ClientHeight    =   5740
   ClientLeft      =   50
   ClientTop       =   290
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5740
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Left            =   5460
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1560
      Width           =   1635
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   7140
      Picture         =   "frm100101_15_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1560
      Width           =   350
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   6930
      TabIndex        =   7
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   8
      Top             =   90
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3525
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   2130
      Width           =   4395
      _ExtentX        =   7743
      _ExtentY        =   6227
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|往來記錄編號|往來日期|往來類別|主旨|地點|內容|聯絡人"
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
      _Band(0).Cols   =   8
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1110
      TabIndex        =   13
      Top             =   1290
      Width           =   7455
      Begin VB.OptionButton Option1 
         Caption         =   "其他對象"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   1
         Top             =   30
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "同對象"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   30
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3525
      Index           =   1
      Left            =   4500
      TabIndex        =   6
      Top             =   2130
      Width           =   4365
      _ExtentX        =   7691
      _ExtentY        =   6227
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|往來記錄編號|往來日期|往來類別|主旨|地點|內容|聯絡人"
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
      _Band(0).Cols   =   8
   End
   Begin MSForms.ListBox lstAtt 
      Height          =   960
      Left            =   1110
      TabIndex        =   4
      Top             =   300
      Width           =   5745
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "10134;1693"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "移動檔案："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   4
      Left            =   4500
      TabIndex        =   16
      Top             =   1590
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "名稱："
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   4500
      TabIndex        =   15
      Top             =   1860
      Width           =   960
   End
   Begin MSForms.Label lblCaseName2 
      Height          =   285
      Left            =   5475
      TabIndex        =   14
      Top             =   1860
      Width           =   3405
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6006;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "請勾選要移至那一道程序："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "移動方式："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   30
      Width           =   945
   End
   Begin MSForms.Label lblKey 
      Height          =   285
      Left            =   1140
      TabIndex        =   9
      Top             =   30
      Width           =   5715
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_15_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/06 Form2.0已修改
'Create By Sindy 2019/11/22
Option Explicit

Public m_strSaveFiles As String
Dim m_MousePointer As Integer
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim m_PrevForm As Form '前一畫面
Dim strKey As String
Dim m_mouseRow As Long


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim intChkCnt As Integer
Dim strNewCF01 As String
Dim strNewCF02 As String
Dim strOldCF01 As String, strOldCF02 As String
Dim strTemp As String
Dim arrData
Dim adoRst As ADODB.Recordset
Dim intGrd_idx As Integer
Dim bolConn As Boolean
Dim ii As Integer, jj As Integer
Dim bolSave As Boolean

On Error GoTo ErrHnd
      
   If Option1(0).Value = True Then '本案件
      intGrd_idx = 0
   Else '其他案件
      intGrd_idx = 1
   End If
   
   '確定
   If Index = 0 Then
      intChkCnt = 0
      For ii = 1 To GRD1(intGrd_idx).Rows - 1
         GRD1(intGrd_idx).row = ii
         GRD1(intGrd_idx).col = 1
         If GRD1(intGrd_idx).CellBackColor = &HFFC0C0 Then
            intChkCnt = intChkCnt + 1
            '新記錄
            strNewCF01 = Trim(GRD1(intGrd_idx).TextMatrix(ii, 1))
            For jj = 0 To lstAtt.ListCount - 1
               strTemp = lstAtt.List(jj)
               arrData = Split(strTemp, "  ")
               If UBound(arrData) <> 1 Then
                  MsgBox "欲搬移的電子檔資料有誤!!!"
                  Exit Sub
               End If
               '舊記錄
               strOldCF01 = Trim(arrData(0))
               strOldCF02 = Trim(arrData(1))
               
               If IsRecordExist(strOldCF01, strOldCF02, False) = False Then
                  MsgBox strOldCF02 & " 此檔案不存在,無法移檔!!!"
                  Exit Sub
               End If
               
               '更新資料
               strNewCF02 = Replace(strOldCF02, strOldCF01, strNewCF01)
               If IsRecordExist(strNewCF01, strNewCF02) = False Then
                  strSql = "update contactFile set cf01='" & strNewCF01 & "',cf02='" & strNewCF02 & "'" & _
                     " where cf01='" & strOldCF01 & "' and upper(cf02)='" & UCase(strOldCF02) & "'"
                  Pub_SeekTbLog strSql 'Add By Sindy 2025/8/13
                  cnnConnection.Execute strSql
                  bolSave = True
               Else
                  Exit Sub
               End If
            Next jj
            If bolSave = True Then Call m_PrevForm.cmdok_Click(3)
            Exit For
         End If
      Next ii
      If intChkCnt = 0 Then
         MsgBox "請勾選要移至那一道程序!!!"
         Exit Sub
      End If
   End If

   Set adoRst = Nothing
   Unload Me
   Screen.MousePointer = m_MousePointer

   Exit Sub

ErrHnd:
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   Set adoRst = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCF01 As String, ByVal strFileName As String, _
   Optional ByVal bolShowMsg As Boolean = True) As Boolean
Dim adoRst As ADODB.Recordset

   IsRecordExist = False
   
   strSql = "SELECT cf01 FROM contactfile WHERE cf01='" & strCF01 & "' and upper(cf02)=upper('" & strFileName & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      IsRecordExist = True
      If bolShowMsg = True Then
         MsgBox "往來記錄編號：" & strCF01 & " 附件：" & strFileName & " 已存在！"
      End If
   End If

   Set adoRst = Nothing
End Function

Private Sub Command5_Click()
   If Text1 = "" Then
      MsgBox "往來對象不可空白!", vbExclamation
      If Text1.Enabled = True Then Text1.SetFocus
      Exit Sub
   End If
   
   Call QueryData(1)
End Sub

Private Sub Form_Load()
Dim sFile
Dim ii As Integer

   MoveFormToCenter Me
   m_MousePointer = Screen.MousePointer
   lstAtt.Clear
   If m_strSaveFiles <> "" Then
      sFile = Split(m_strSaveFiles, "&")
      For ii = 0 To UBound(sFile)
         lstAtt.AddItem sFile(ii), 0
         'SetListScroll lstAtt
      Next ii
   End If

   Me.lblKey = m_PrevForm.Label3
   
   Call SetGrd(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm100101_15_1 = Nothing
End Sub

Public Function QueryData(Index As Integer) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   QueryData = False
   '清空及預設欄位值
   GRD1(Index).Clear
   Call SetGrd(Index)
   
   If Index = 0 Then
      strKey = lblKey.Tag
   Else
      strKey = Left(ChangeCustomerL(Text1.Text), 8)
      lblCaseName2.Caption = ""
   End If
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Modify By Sindy 2020/5/15
   If m_PrevForm.m_quyDataKind = 1 Then '1.國內
      strSql = "select ' ' AS V,Cor01 AS 往來記錄編號" & _
         "," & SQLDate("Cor02") & " 往來日期,'' 往來類別,Cor04 主旨" & _
         ",'' 地點,Cor05 內容,'' 聯絡人" & _
         " from contactrecord1" & _
         " where SUBSTR(cor03,1,8)='" & strKey & "'" & _
         " order by Cor01 desc"
   Else '0.國外
   '2020/5/15 END
      strSql = "select ' ' AS V,CR01 AS 往來記錄編號" & _
         "," & SQLDate("CR02") & " 往來日期,AC03 往來類別,CR06 主旨" & _
         ",CR07 地點,CR08 內容,CR04 聯絡人" & _
         " from contactrecord,allcode" & _
         " where SUBSTR(cr03,1,8)='" & strKey & "'" & _
         " and ac01(+)='11' and cr05=ac02(+)" & _
         " order by cr01 desc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1(Index).Recordset = rsTmp
      QueryData = True
      GRD1(Index).col = 0
      GRD1(Index).row = 1
      
      If Index = 1 Then
         'Modify By Sindy 2020/5/15
         If m_PrevForm.m_quyDataKind = 1 Then '1.國內
            strExc(0) = "select N1,nvl(pcc05,nvl(pcc03,pcc04)) N2, N3, NO1, NO2, PCU51" & _
               " from (select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) N1,CU01||CU02 NO1,CU01 NO2,CU10 N3,'' PCU51" & _
               " from customer where cu01='" & strKey & "' and cu02='0'" & _
               " union all select NVL(Poc03,DECODE(Poc23,NULL,Poc27,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)) N1,Poc01||Poc02 NO1,Poc01 NO2,Poc04 N3,'' PCU51" & _
               " from potcustomer1 where poc01='" & strKey & "' and poc02='0'" & _
               ") A,potcustcont where A.NO2=pcc01(+)"
            If strKey1 <> "" Then
               strExc(0) = strExc(0) & " and pcc02(+)='" & strKey1 & "'"
            Else
               strExc(0) = strExc(0) & " and pcc02(+)='ZZ'"
            End If
         Else '0.國外
         '2020/5/15 END
            strExc(0) = "select N1,nvl(pcc05,nvl(pcc03,pcc04)) N2, N3, NO1, NO2, PCU51" & _
               " from (select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) N1,CU01||CU02 NO1,CU01 NO2,CU10 N3,'' PCU51" & _
               " from customer where cu01='" & strKey & "' and cu02='0'" & _
               " union all select NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) N1,FA01||FA02 NO1,FA01 NO2,FA10 N3,'' PCU51" & _
               " from fagent where fa01='" & strKey & "' and fa02='0'" & _
               " union all select NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) N1,PCU01||PCU02 NO1,PCU01 NO2,PCU09 N3,PCU51" & _
               " from potcustomer where pcu01='" & strKey & "' and pcu02='0'" & _
               ") A,potcustcont where A.NO2=pcc01(+)"
            If strKey1 <> "" Then
               strExc(0) = strExc(0) & " and pcc02(+)='" & strKey1 & "'"
            Else
               strExc(0) = strExc(0) & " and pcc02(+)='ZZ'"
            End If
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lblCaseName2.Caption = RsTemp.Fields(0)
         End If
      End If
      
      cmdOK(Index).Default = True
      rsTmp.Close
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      If Index = 0 Then
         Unload Me
         Screen.MousePointer = vbDefault
         Exit Function
      End If
   End If
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Set rsTmp = Nothing
End Function

Private Sub SetGrd(Index As Integer, Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1               2           3           4       5       6       7
   arrGridHeadText = Array("V", "往來記錄編號", "往來日期", "往來類別", "主旨", "地點", "內容", "聯絡人")
   'Modify By Sindy 2020/5/15
   If m_PrevForm.m_quyDataKind = 1 Then '1.國內
      arrGridHeadWidth = Array(200, 950, 800, 0, 800, 0, 500, 0)
   Else '0.國外
   '2020/5/15 END
      arrGridHeadWidth = Array(200, 950, 800, 1250, 800, 500, 500, 500)
   End If
   GRD1(Index).Visible = False
   GRD1(Index).Cols = UBound(arrGridHeadText) + 1
   If bolSetRow = True Then
      GRD1(Index).Rows = 2
   End If
   For iRow = 0 To GRD1(Index).Cols - 1
      GRD1(Index).row = 0
      GRD1(Index).col = iRow
      If bolSetRow = True Then
         GRD1(Index).Text = arrGridHeadText(iRow)
      End If
      GRD1(Index).ColWidth(iRow) = arrGridHeadWidth(iRow)
      If bolSetRow = True Then
         GRD1(Index).CellAlignment = flexAlignCenterCenter
      End If
   Next
   GRD1(Index).Visible = True
End Sub

Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long

   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next

   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim oldRow As Long
Dim jj As Integer

getGrdColRow GRD1(Index), x, y, nCol, nRow
GRD1(Index).col = nCol
GRD1(Index).row = nRow
oldRow = m_mouseRow

GRD1(Index).Visible = False
If GRD1(Index).MouseRow <> 0 And Trim(GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 1)) <> "" Then
   m_mouseRow = GRD1(Index).MouseRow
   If oldRow <> m_mouseRow Then
      GRD1(Index).row = oldRow
      GRD1(Index).col = 1
      If GRD1(Index).CellBackColor = &HFFC0C0 Then
         '清除反白
         GRD1(Index).TextMatrix(oldRow, 0) = ""
         For jj = 1 To GRD1(Index).Cols - 1
            GRD1(Index).col = jj
            GRD1(Index).CellBackColor = QBColor(15)
         Next jj
      End If
   End If

   GRD1(Index).row = GRD1(Index).MouseRow
   GRD1(Index).col = 1
   If GRD1(Index).CellBackColor = &HFFC0C0 Then
      '清除反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = ""
      For jj = 1 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = QBColor(15)
      Next jj
   Else
      '資料列反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = "V"
      For jj = 1 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = &HFFC0C0
      Next jj
   End If
End If
GRD1(Index).Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 1 Then
      Text1.Enabled = True
      Command5.Enabled = True
      Call QueryData(0)
      GRD1(0).Enabled = False
      Text1.SetFocus
      Command5.Default = True
   Else
      Text1.Enabled = False: Text1 = ""
      Command5.Enabled = False
      '清空及預設欄位值
      GRD1(1).Clear
      Call SetGrd(1)
      lblCaseName2.Caption = ""
      GRD1(0).Enabled = True
   End If
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
