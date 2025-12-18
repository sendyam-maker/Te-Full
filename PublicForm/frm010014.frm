VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010014 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問案件電話諮詢"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   690
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Top             =   449
      Width           =   800
   End
   Begin VB.TextBox txtAccept 
      Height          =   264
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1470
      Width           =   1095
   End
   Begin VB.TextBox txtcp04 
      Height          =   288
      Left            =   2916
      MaxLength       =   2
      TabIndex        =   3
      Top             =   492
      Width           =   375
   End
   Begin VB.TextBox txtcp03 
      Height          =   288
      Left            =   2592
      MaxLength       =   1
      TabIndex        =   2
      Top             =   492
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   288
      Left            =   1584
      MaxLength       =   6
      TabIndex        =   1
      Top             =   492
      Width           =   975
   End
   Begin VB.TextBox txtcp01 
      Height          =   288
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "LA"
      Top             =   492
      Width           =   495
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8376
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7548
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3945
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6959
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
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   1145
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2220
      TabIndex        =   17
      Top             =   820
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "分所案號："
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   509
      Width           =   900
   End
   Begin VB.Label lbeHC07 
      Height          =   255
      Left            =   5250
      TabIndex        =   15
      Top             =   509
      Width           =   2055
   End
   Begin VB.Label Label21 
      Caption         =   "來電日期："
      Height          =   252
      Left            =   144
      TabIndex        =   14
      Top             =   1476
      Width           =   972
   End
   Begin VB.Label lbeCusNum 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   820
      Width           =   1092
   End
   Begin VB.Label lbeNum 
      Height          =   252
      Left            =   1488
      TabIndex        =   11
      Top             =   540
      Width           =   1692
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   144
      TabIndex        =   10
      Top             =   836
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   1160
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   144
      TabIndex        =   6
      Top             =   510
      Width           =   972
   End
End
Attribute VB_Name = "frm010014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim rs As New ADODB.Recordset
Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean


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
      .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 800: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "相對人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 800: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      'Modified by Lydia 2015/10/05
'      .col = 6: .ColWidth(6) = 800: .Text = "承辦律師"
      .col = 6: .ColWidth(6) = 800: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      'Modified by Lydia 2015/10/05
'      .col = 7: .ColWidth(7) = 800: .Text = "承辦法務"
      .col = 7: .ColWidth(7) = 800: .Text = "協辦人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 800: .Text = "回執日"
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .ColWidth(10) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1000: .Text = "取消收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 12: .ColWidth(12) = 0
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
   
   lbeHC07 = ""
   lbeCusNum = ""
   lbeCusName = ""
   cboCaseName.Clear
   cboCaseName.Text = ""
   MSHFlexGrid1.Rows = 2
   
   If QueryDB = False Then
      MsgBox "本所案號不存在!"
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
   strExc(1) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
     "cp09,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64),NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))),decode(CP13,S1.ST01,S1.ST02)," + _
     "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02)," + _
     "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
     "decode(cp46,null,'',SUBSTR(CP46, 1, 4)- 1911 || '/' || SUBSTR(CP46, 5, 2)|| '/' || SUBSTR(CP46, 7, 2))," + _
     "cpm03,decode(cp57,null,'',SUBSTR(CP57, 1, 4)- 1911 || '/' || SUBSTR(CP57, 5, 2)|| '/' || SUBSTR(CP57, 7, 2)) " + _
     "from caseprogress, hirecase,STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP, CUSTOMER " + _
     "where " & ChgHirecase(LcTmp) & " AND HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " + _
     "AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) " + _
     "and cp09<'C' AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) order by cp05 DESC,CP09"

   intI = 0
   Set rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      PutDataInGrid
   End If
   
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
   
   If txtAccept = "" Then
      MsgBox "來電日期不可空白", vbCritical
      txtAccept.SetFocus
      Exit Sub
   Else
      strExc(0) = txtcp01 & txtcp02 & txtcp03 & txtcp04
   End If
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .col = 0
         If .Text = "v" Then
            Set frm010015.UpForm = Me
            '2011/5/26 ADD BY SONIA
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            '2011/5/26 END
            blnChoese = True
            .Text = ""
            .col = 2
            frm010015.Show
            frm010015.Tag = .Text & txtAccept
            frm010015.GetData (0)
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
   cmdSearch.Default = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
   txtAccept.Text = ChangeWStringToTString(GetTodayDate)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010014 = Nothing
End Sub

Private Sub lbeCusNum_Change()
Dim StrCusName As String
   
   If Len(lbeCusNum) > 7 Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCustomer(lbeCusNum, StrCusName) Then lbeCusName = StrCusName
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
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
         CmdSure.SetFocus
      End If
   End With
End Sub

Private Sub PutDataInGrid()
Dim i As Integer, strTempName As String, strTemp As String

   With MSHFlexGrid1
      .Visible = False
      
      If Not (rs.EOF And rs.BOF) Then
         Set .Recordset = rs
      End If
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click
End Sub

Private Sub txtAccept_GotFocus()
   TextInverse txtAccept
End Sub

Private Sub txtAccept_Validate(Cancel As Boolean)
   If txtAccept <> "" Then
      If CheckIsTaiwanDate(txtAccept) Then
         If Val(GetTaiwanTodayDate) - Val(txtAccept) < 0 Then
             MsgBox "來電日期不可大於系統日", vbCritical
             Cancel = True
         Else
             If MSHFlexGrid1.Text <> "" Then
                CmdSure.Enabled = True
                CmdSure.SetFocus
              End If
          End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtAccept
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
      cmdSearch.Enabled = True
      cmdSearch.SetFocus
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
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
Dim bQuery As Boolean
   
   strTM01 = txtcp01.Text
   strTM02 = txtcp02.Text
   If txtcp03.Text <> Empty Then
      strTM03 = txtcp03.Text
   Else
      strTM03 = "0"
   End If
   If txtcp04.Text <> Empty Then
      strTM04 = txtcp04.Text
   Else
      strTM04 = "00"
   End If
   
   bQuery = QueryHireCase(strTM01, strTM02, strTM03, strTM04)
   QueryDB = bQuery
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cboCaseName.AddItem rsTmp.Fields("HC06")
      End If
      ' 顯示案件名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("HC05"), 0)
         lbeCusNum.Caption = rsTmp.Fields("HC05")
      End If
      ' 分所案號
      If IsNull(rsTmp.Fields("HC07")) = False Then
         lbeHC07.Caption = rsTmp.Fields("HC07")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
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
