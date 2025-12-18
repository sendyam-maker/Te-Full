VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071010 
   BorderStyle     =   1  '單線固定
   Caption         =   "回執"
   ClientHeight    =   5850
   ClientLeft      =   90
   ClientTop       =   690
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9345
   Begin VB.TextBox txtcp01 
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtcp02 
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtcp03 
      Height          =   285
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtcp04 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3552
      TabIndex        =   4
      Top             =   315
      Width           =   800
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8328
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7500
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4245
      Left            =   150
      TabIndex        =   7
      Top             =   1485
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   7488
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Top             =   1080
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
      Left            =   2430
      TabIndex        =   15
      Top             =   720
      Width           =   6375
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11245;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbldestroy 
      AutoSize        =   -1  'True
      Caption         =   "11/01/01表示回執未回"
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
      Index           =   1
      Left            =   5500
      TabIndex        =   14
      Top             =   480
      Width           =   1920
   End
   Begin VB.Label lbldestroy 
      AutoSize        =   -1  'True
      Caption         =   "回執日11/11/11表示退件"
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
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   376
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1095
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   736
      Width           =   972
   End
   Begin VB.Label lbeNum 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lbeCusNum 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "frm071010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
Dim rs As New ADODB.Recordset
Dim m_CP As String

Private Sub cmdEnd_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim LcTmp As String
   
   cmdSure.Enabled = True 'Add By Sindy 2011/6/27
   
   MSHFlexGrid1.Clear
   lbeCusNum.Caption = ""
   lbeCusName.Caption = ""
   cboCaseName.Clear
   cboCaseName.Text = ""
   If txtcp03 = "" Then
      LcTmp = txtcp01 + txtcp02 + "000"
   Else
      If txtcp04 = "" Then
         LcTmp = txtcp01 + txtcp02 + txtcp03 + "00"
      Else
         LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
      End If
   End If
   m_CP = LcTmp
   If txtcp01 <> "LA" Then
      '2006/1/4 MODIFY BY SONIA 不限制B類收文號,改抓CP50收件人有值且已發文的進度資料
'      strExc(1) = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
'     " cp09 AS 收文號,decode(lc15,020,cpm04,cpm03) AS 案件性質,decode(CP13," + _
'     " S1.ST01,S1.ST02) AS 智權人員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
'     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
'     " cp50 as 收件人,decode(cp71,or01,or02) as 機關代號,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
'     " STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
'     ChgLawcase(LcTmp) & " AND cp13 = s1.st01(+)  AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
'     " cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and (cp09>='B' and cp09<'C') and cp27 is not null"
    'Modified by Lydia 2015/10/05 '承辦律師'改為'承辦人'、'承辦法務'改為'協辦人員'
    'Modified by Lydia 2016/05/30 +回執退件日/郵局送達日(cp47)
      strExc(1) = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(lc15,020,cpm04,cpm03) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 智權人員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 協辦人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp46,null,'',SUBSTR(CP46, 1, 4)- 1911 || '/' || SUBSTR(CP46, 5, 2)|| '/' || SUBSTR(CP46, 7, 2)) as 回執日," + _
     " cp50 as 收件人,CP64 AS 進度備註,sqldatet(cp47) as cp47 ,decode(cp71,or01,or02) as 機關名稱,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
     " STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
     ChgLawcase(LcTmp) & " AND cp13 = s1.st01(+)  AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
     " cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and CP50 IS NOT NULL and cp27 is not null"
      '2006/1/4 END
   Else
      '2006/1/4 MODIFY BY SONIA 不限制B類收文號,改抓CP50收件人有值且已發文的進度資料
'      strExc(1) = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
'     " cp09 AS 收文號,cpm03 AS 案件性質,decode(CP13," + _
'     " S1.ST01,S1.ST02) AS 智權人員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
'     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
'     " cp50 as 收件人,decode(cp71,or01,or02) as 機關代號,hc05,hc06 from caseprogress, hirecase,STAFF S1," + _
'     " STAFF S2,STAFF S3, CASEPROPERTYMAP,organization  where " & ChgCaseprogress(LcTmp) + " and " + _
'     ChgHirecase(LcTmp) & " AND cp13 = s1.st01(+) and (cp09>='B' AND CP09<'C') " + _
'     " AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and cp27 is not null"
    'Modified by Lydia 2015/10/05
    'Modified by Lydia 2016/05/30 +回執退件日/郵局送達日(cp47)
      strExc(1) = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,cpm03 AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 智權人員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 協辦人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp46,null,'',SUBSTR(CP46, 1, 4)- 1911 || '/' || SUBSTR(CP46, 5, 2)|| '/' || SUBSTR(CP46, 7, 2)) as 回執日," + _
     " cp50 as 收件人,CP64 AS 進度備註,sqldatet(cp47) as cp47 ,decode(cp71,or01,or02) as 機關名稱,hc06,null,null,hc05 from caseprogress, hirecase,STAFF S1," + _
     " STAFF S2,STAFF S3, CASEPROPERTYMAP,organization  where " & ChgCaseprogress(LcTmp) + " and " + _
     ChgHirecase(LcTmp) & " AND cp13 = s1.st01(+) and CP50 IS NOT NULL " + _
     " AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and cp27 is not null"
      '2006/1/4 END
   End If
   strExc(1) = strExc(1) & " AND CP10<>'0' ORDER BY CP27 DESC,CP09" '2009/9/9 ADD BY SONIA 加剔除顧問聘任及排序條件
   intI = 0
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   GridHead
   If intI = 1 Then
      PutDataInGrid
   Else
      lbeCusNum = ""
      lbeCusName = ""
      cboCaseName.Clear
      MSHFlexGrid1.Rows = 2
    '  cmdSure.Enabled = False
   End If
   'cmdSearch.Enabled = False
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
End Sub

Private Sub cmdSure_Click()
Dim i As Integer, blnChoese As Boolean
   blnChoese = False
   With MSHFlexGrid1
   For i = 1 To .Rows - 1
      '.Col = 0
      If .TextMatrix(i, 0) = "v" Then
         'Add By Sindy 2011/6/27
         '2011/6/28 add by sonia
         If .TextMatrix(i, 8) = "11/11/11" Then
            MsgBox "已退回, 不可再點選此筆資料 !", vbCritical
            cmdSure.Enabled = False
            Exit Sub
         '2011/6/28 end
         ElseIf .TextMatrix(i, 8) <> "" Then
            MsgBox "已輸入回執, 不可再點選此筆資料 !", vbCritical
            cmdSure.Enabled = False
            Exit Sub
         End If
         '2011/6/27 End
         blnChoese = True
         Exit For
      End If
   Next
   End With
   If Not blnChoese Then
      MsgBox "請點選輸入資料", vbCritical
      cmdSure.Enabled = False
      Exit Sub
   End If
   
   frm071011.Show
   'cmdSearch.SetFocus
   Me.Hide
End Sub

Private Sub Form_Activate()
   txtcp01.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'cmdSearch.Enabled = False
   cmdSure.Enabled = False
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False

End Sub

Private Sub GridHead()
Dim i As Integer

   With MSHFlexGrid1
      .Cols = 16
      blnOKtoShow = False
      .Visible = False
      .row = 0
      .col = 0
      .ColWidth(0) = 200
      .Text = "v"
      .col = 1
      .ColWidth(1) = 800
      .col = 2
      .ColWidth(2) = 1000
      .col = 3
      .ColWidth(3) = 1000
      .col = 4
      .ColWidth(4) = 800
      .col = 5
      .ColWidth(5) = 800
      .col = 6
      .ColWidth(6) = 800
      .col = 7
      .ColWidth(7) = 800
      .col = 8
      .ColWidth(8) = 800
      .col = 9
      .ColWidth(9) = 1200
      .col = 10
      .ColWidth(10) = 1200
      .col = 11
      'Modified by Lydia 2016/05/30
      '.ColWidth(11) = 1500
      .ColWidth(11) = 2000
      .Text = "回執退件日/郵局送達日"
      'end 2016/05/30
      .col = 12
      'Modified by Lydia 2016/05/30
      '.ColWidth(12) = 0
      .ColWidth(12) = 1500
      .col = 13
      .ColWidth(13) = 0
      .col = 14
      .ColWidth(14) = 0
      .col = 15
      .ColWidth(15) = 0
      
      .CellAlignment = flexAlignCenterCenter
      
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      If .Rows > 1 Then .row = 1
      
      .Visible = True
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071010 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Integer

   intCols = MSHFlexGrid1.Cols - 1
   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols) Then Exit Sub
   cmdSure.Enabled = True
   cmdSure.SetFocus

End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click

End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)

End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
  
   txtcp01 = UCase(txtcp01)
   If IsEmptyText(txtcp01) = False Then
      blnCom1 = True
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtcp01) <> "3" And CheckSys(txtcp01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         blnCom1 = False
         TextInverse txtcp01
         Exit Sub
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         blnCom1 = False
         TextInverse txtcp01
         Exit Sub
      End If
   End If

   'If txtcp01 <> "" Then
   '    txtcp01 = UCase(txtcp01)
   '    If txtcp01 = "L" Or txtcp01 = "FCL" Or txtcp01 = "LA" Then
   '       blnCom1 = True
   '    Else
   '       DataErrorMessage 1, "系統類別"
   '       blnCom1 = False
   '       Cancel = True
   '    End If
   'End If
   
   If Cancel Then TextInverse txtcp01

End Sub

Private Sub txtcp02_Change()
   If Len(txtcp02) = 6 Then
      blnCom2 = True
      CmdSearch.Enabled = True
   End If

End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If txtcp02 <> "" Then
      blnCom2 = True
      CmdSearch.Enabled = True
   End If
   If Cancel Then TextInverse txtcp02
End Sub
Private Sub txtcp03_Change()
   blnCom3 = True
   ChkCmd
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

Private Sub PutDataInGrid()
Dim i As Integer, strTempName As String, t As Integer, strTemp() As String
Dim strCus As String
   
   With MSHFlexGrid1
      .Visible = False
      If Not (rs.EOF And rs.BOF) Then
      If txtcp01 <> "LA" Then
         strCus = IIf(IsNull(rs.Fields!LC11), "", rs.Fields!LC11)
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
         If ClsPDGetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
         If Not IsNull(rs.Fields!lc05) Then cboCaseName.AddItem "中:" + rs.Fields!lc05
         If Not IsNull(rs.Fields!lc06) Then cboCaseName.AddItem "英:" + rs.Fields!lc06
         If Not IsNull(rs.Fields!lc07) Then cboCaseName.AddItem "日:" + rs.Fields!lc07
      Else
         strCus = rs.Fields!hc05
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
         If ClsPDGetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
         If Not IsNull(rs.Fields!hc06) Then cboCaseName.AddItem "中:" + rs.Fields!hc06
      End If
      
      cboCaseName.ListIndex = 0
      i = 1
      rs.MoveFirst
      Set .Recordset = rs
      Call GridHead 'Added by Lydia 2016/05/30
      End If
      .Visible = True
   
   End With
    
End Sub
Private Function DisArray(strAcceptMail As String) As String()
Dim i As Integer, j As Integer, strTemp() As String, t As Integer

   j = 1
   For i = 1 To Len(rs.Fields!cp50)
       If Mid(rs.Fields!cp50, i, 1) = ";" Then
            ReDim Preserve strTemp(t)
            strTemp(t) = Mid(rs.Fields!cp50, j, i - j)
            j = i + 1
            t = t + 1
       End If
    Next
    DisArray = strTemp
 
End Function

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
  blnCom4 = True
  ChkCmd
End Sub
' 設定該筆收文資料已做完存檔的工作
Public Sub SetDataComplete(ByVal strCP09 As String)
Dim nIndex As Integer
   
   For nIndex = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(nIndex, 2) = strCP09 Then
         MSHFlexGrid1.TextMatrix(nIndex, 0) = Empty
         Exit For
      End If
   Next nIndex

End Sub
