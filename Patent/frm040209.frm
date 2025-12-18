VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040209 
   BorderStyle     =   1  '單線固定
   Caption         =   "CFP申請文件齊備維護"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   690
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3336
      TabIndex        =   4
      Top             =   570
      Width           =   800
   End
   Begin VB.TextBox txtcp04 
      Height          =   285
      Left            =   2916
      MaxLength       =   2
      TabIndex        =   3
      Top             =   615
      Width           =   375
   End
   Begin VB.TextBox txtcp03 
      Height          =   285
      Left            =   2592
      MaxLength       =   1
      TabIndex        =   2
      Top             =   615
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   285
      Left            =   1584
      MaxLength       =   6
      TabIndex        =   1
      Top             =   615
      Width           =   975
   End
   Begin VB.TextBox txtcp01 
      Height          =   285
      Left            =   1056
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "CFP"
      Top             =   615
      Width           =   495
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   8376
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdSure 
      Caption         =   "確定(&O)"
      Height          =   405
      Left            =   7548
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3825
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6747
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1050
      TabIndex        =   14
      Top             =   1425
      Width           =   8070
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14235;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      Caption         =   "lblCancel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8400
      TabIndex        =   18
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8400
      TabIndex        =   17
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label5 
      Caption         =   "分所案號："
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   660
      Width           =   900
   End
   Begin VB.Label lbePA47 
      Height          =   255
      Left            =   5250
      TabIndex        =   15
      Top             =   660
      Width           =   2055
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Left            =   2235
      TabIndex        =   12
      Top             =   1065
      Width           =   6900
      VariousPropertyBits=   27
      Size            =   "12171;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCusNum 
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label lbeNum 
      Height          =   255
      Left            =   1485
      TabIndex        =   10
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "申  請  人："
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1455
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "frm040209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (MSHFlexGrid1,cboCaseName,lbeCusName)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2011/5/20 CREATE BY SONIA
Option Explicit

Dim rs As New ADODB.Recordset
Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
Dim pa() As String
Dim intWhere As Integer

Private Sub GridHead()
Dim i As Integer

   With MSHFlexGrid1
      blnOKtoShow = False
      .Cols = 15
      .Visible = False
      .row = 0
      .col = 0
      .Visible = True
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .ColWidth(0) = 200: .Text = "V"
      .col = 1: .ColWidth(1) = 800: .Text = "收文日"
      .CellAlignment = flexAlignRightCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "相關收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 600: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 600: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignRightCenter
      .col = 8: .ColWidth(8) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignRightCenter
      .col = 9: .ColWidth(9) = 1000: .Text = "文件齊備日"
      .CellAlignment = flexAlignRightCenter
      .col = 10: .ColWidth(10) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1000: .Text = "取消收文日"
      .CellAlignment = flexAlignRightCenter
      .col = 12: .ColWidth(12) = 1000: .Text = "收款情形"
      .CellAlignment = flexAlignCenterCenter
      .col = 13: .ColWidth(13) = 0
      .col = 14: .ColWidth(14) = 0
      
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      If .Rows > 1 Then
         '將第一列反白
         .row = 1
      End If
      .Visible = True
   End With
End Sub

Private Sub cmdBack_Click()
   Unload Me
End Sub

Public Sub cmdSearch_Click()
Dim LcTmp As String
Dim i As Integer
Dim IntTemp1 As Long
Dim IntTemp2 As Long

   lbePA47 = ""
   lbeCusNum = ""
   lbeCusName = ""
   cboCaseName.Clear
   'cboCaseName.Text = "" 'Removed by Morgan 2022/1/3
   MSHFlexGrid1.Rows = 2
   Me.lblClose.Caption = ""
   Me.lblCancel.Caption = ""
   
   If QueryDB = False Then
      txtcp02.SetFocus
      TextInverse txtcp02
      MSHFlexGrid1.Clear
      MSHFlexGrid1.Rows = 2
      Exit Sub
   End If
 
   MSHFlexGrid1.Clear
   If txtcp03.Text = "" Then txtcp03.Text = "0"
   If txtcp04.Text = "" Then txtcp04.Text = "00"
   LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
   'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
   strExc(1) = "select ' ' AS V," & SQLDate("CP05") & " as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號," + _
     "NVL(S2.ST02,CP14) as 承辦人,NVL(S1.ST02,CP13) as 智權人員," & SQLDate("CP06") & " as 本所期限," & SQLDate("CP07") & " as 法定期限," & SQLDate("CP143") & " as 文件齊備日," & SQLDate("CP27") & " as 發文日," & SQLDate("CP57") & " as 取消收文日" & _
     ",decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60" + _
     " from caseprogress, patent,STAFF S1,STAFF S2, CASEPROPERTYMAP " + _
     "where " & ChgPatent(LcTmp) & " AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) " + _
     "AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) " + _
     "and cp09<'C' order by cp05 desc,CP09"

   intI = 0
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      MSHFlexGrid1.Visible = False
      Set MSHFlexGrid1.Recordset = rs
      MSHFlexGrid1.Visible = True
   End If

   '檢查資料筆數
   For i = 1 To rs.RecordCount
      MSHFlexGrid1.row = i
      '收款情形
      IntTemp1 = 0
      IntTemp2 = 0
      Me.MSHFlexGrid1.col = 12
      If Not IsNull(MSHFlexGrid1.Text) Then
          'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
          'If Mid(MSHFlexGrid1.Text, 1, 1) = "E" Then
          '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & MSHFlexGrid1.Text & "'"
          'Else
          If Mid(MSHFlexGrid1.Text, 1, 1) = "X" Then
          'end 2011/8/12
             strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0 FROM ACC1K0 WHERE A1K01='" & MSHFlexGrid1.Text & "'"
          'End If 'Remove by Morgan 2011/8/12
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                If Not IsNull(adoRecordset1.Fields(0)) Then
                    IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                End If
                If Not IsNull(adoRecordset1.Fields(1)) Then
                    IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                End If
                If Not IsNull(adoRecordset1.Fields(4)) Then
                    IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                End If
                If Not IsNull(adoRecordset1.Fields(5)) Then
                    IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                End If
                If IntTemp1 = IntTemp2 Then
                     MSHFlexGrid1.Text = "收回"
                Else
                     If IntTemp2 = 0 Then
                         MSHFlexGrid1.Text = "未收"
                     Else
                         If IntTemp1 > IntTemp2 Then
                             MSHFlexGrid1.Text = "部分收回"
                         End If
                     End If
                 End If
            End If
          End If 'Add by Morgan 2011/8/12
      End If
   Next i

   GridHead
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
End Sub

Private Sub cmdSure_Click()
Dim i As Integer, blnChoese As Boolean

   blnChoese = False
   
   If txtcp02.Text = "" Then
      MsgBox "本所案號不可空白!"
      txtcp02.SetFocus
      Exit Sub
   End If
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .col = 0
         If .Text = "V" Then
            .col = 10
            If .Text <> "" Then
               MsgBox "此程序已發文,不可點選 !"
               Exit Sub
            End If
            .col = 11
            If .Text <> "" Then
               MsgBox "此程序已取消收文,不可點選 !"
               Exit Sub
            End If
            Set frm040209_1.UpForm = Me
            blnChoese = True
            .Text = ""
            .col = 2
            frm040209_1.Show
            frm040209_1.Tag = .Text
            frm040209_1.GetData (0)
            Exit For
         End If
      Next
   End With
   
   If Not blnChoese Then
      MsgBox "請點選輸入資料", vbCritical
      Exit Sub
   End If
   
   Me.Hide
End Sub

Private Sub Form_Activate()
   CmdSearch.Default = True
End Sub

Private Sub Form_Initialize()
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
   txtcp01 = "CFP"
   Me.lblClose.Caption = ""
   Me.lblCancel.Caption = ""
   
   intWhere = 國外_CF
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040209 = Nothing
End Sub

Private Sub lbeCusNum_Change()
Dim StrCusName As String
   
   If Len(lbeCusNum) > 7 Then
      If ClsPDGetCustomer(lbeCusNum, StrCusName) Then lbeCusName = StrCusName
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
      intCols = MSHFlexGrid1.Cols - 1
      ShowBar MSHFlexGrid1, intLastRow, intCols
      .col = 0
      ClearGrid
      .row = intLastRow
      If .Text = "V" Then
         .Text = ""
      Else
         .Text = "V"
         cmdSure.SetFocus
      End If
   End With
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If txtcp02 <> "" Then
      blnCom2 = True
   End If
   If Cancel Then TextInverse txtcp02
End Sub

Private Sub ChkCmd()
   If blnCom1 And blnCom2 And blnCom3 And blnCom4 Then
      CmdSearch.Enabled = True
      CmdSearch.SetFocus
   End If
End Sub

Private Sub txtcp03_GotFocus()
  TextInverse txtcp03
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
   If txtcp03 <> "" Then
      blnCom3 = True
   End If
   If Cancel Then TextInverse txtcp03
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtcp04 <> "" Then
      blnCom4 = True
   Else
      blnCom4 = True
      ChkCmd
   End If
   If Cancel Then TextInverse txtcp04
End Sub

Private Function QueryDB() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim bQuery As Boolean
   
   QueryDB = False
   pa(1) = txtcp01.Text
   pa(2) = txtcp02.Text
   If txtcp03.Text <> Empty Then
      pa(3) = txtcp03.Text
   Else
      pa(3) = "0"
   End If
   If txtcp04.Text <> Empty Then
      pa(4) = txtcp04.Text
   Else
      pa(4) = "00"
   End If
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      ' 案件名稱
      AddCboName cboCaseName, pa(5), pa(6), pa(7)
      ' 申請人
      lbeCusName.Caption = GetCustomerName(pa(26), 0)
      lbeCusNum.Caption = pa(26)
      ' 分所案號
      lbePA47.Caption = pa(47)
      ' 閉卷
      If Not pa(57) = "" Then
          Me.lblClose.Caption = "已閉卷"
      End If
      ' 銷卷
      If pub_strUserOffice = "1" And Not pa(108) = "" Then
          Me.lblCancel.Caption = "已銷卷"
      ElseIf pub_strUserOffice <> "1" And Not pa(136) = "" Then
          Me.lblCancel.Caption = "已銷卷"
      End If
      QueryDB = True
   End If
   
End Function

Private Sub ClearGrid()
Dim i As Integer
   
   With MSHFlexGrid1
      .Visible = False
      For i = 1 To .Rows - 1
         .col = 0
         .row = i
         .Text = ""
      Next
      .Visible = True
   End With
End Sub
