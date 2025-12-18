VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100118_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "監視系統案件查詢 "
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   1005
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8532
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   20
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   20
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6084
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7308
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   20
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5232
      Left            =   36
      TabIndex        =   0
      Top             =   468
      Width           =   9252
      _ExtentX        =   16325
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   12
   End
End
Attribute VB_Name = "frm100118_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim strSQL5 As String, StrSQL7 As String
Dim s As Integer, i As Integer, j As Integer, intK As Integer
Dim strSql  As String, strTemp As Variant, StrTest As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


Private Sub SetDataListWidth()
grdDataList.Cols = 13
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 2: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(2) = 0
Else
    grdDataList.ColWidth(2) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(3) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "收文日"
grdDataList.ColWidth(4) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "總收文號"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "案件性質"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "發文字號"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "發文日"
grdDataList.ColWidth(8) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "承辦人"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "智權人員"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = "BTTM"
grdDataList.ColWidth(11) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = "CCC Code"
grdDataList.ColWidth(12) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
        Dim Str01 As String
        grdDataList.col = 1
        If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_8.Show
            frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_8.StrMenu
            Screen.MousePointer = vbDefault
             Me.Enabled = True
             Exit Sub
        End If
     End If
     Next i
     Me.Enabled = True
Case 1
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
             Me.Enabled = True
             Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 1
        If Not IsNull(grdDataList.Text) Then
            Screen.MousePointer = vbHourglass
            frm100101_8.Show
            'frm100101_8.Hide
             
            'Modify By Cheng 2002/04/26
'            frm100101_8.Tag = grdDataList.Text
            frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_8.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm100101_8.Show
            Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
            Loop Until Not frm100101_8.Visible
            Unload frm100101_8
        End If
        grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
     End If
     Next i
     Me.Enabled = True
     Me.Show
Case 1
     Me.Enabled = False
     
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            'frm100101_2.Hide
             
            'Modify By Cheng 2002/04/26
'            frm100101_2.Tag = grdDataList.Text
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm100101_2.Show
            Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
            Loop Until Not frm100101_2.Visible
            Unload frm100101_2
         End If
        grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
     End If
     Next i
     Me.Enabled = True
     Me.Show
Case 2
     Me.Hide
Case 3
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()
Dim ii As Integer

Me.Enabled = False
strSQL5 = ""

'本所案號
If frm100118_1.Option1(0).Value = True Then
   'strSQL = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,CP05 AS 收文日,CP09 AS 總收文號,CP10 AS 案件性質,CP28 AS 發文字號,CP27 AS 發文日,CP14 AS 承辦人,CP13 AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP01='TM' AND SP02='" & frm100118_1.txt1(1) & "' AND SP03='" & frm100118_1.txt1(2) & "' AND SP04='" & frm100118_1.txt1(3) & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) "
   strSQL5 = strSQL5 & " AND SP02='" & frm100118_1.txt1(1) & "' "
   pub_QL05 = pub_QL05 & ";" & frm100118_1.Option1(0).Caption & "TM" & "-" & frm100118_1.txt1(1) 'Add By Sindy 2010/11/16
   If Len(Trim(frm100118_1.txt1(2))) <> 0 Then
      strSQL5 = strSQL5 & " AND SP03='" & frm100118_1.txt1(2) & "' "
      pub_QL05 = pub_QL05 & "-" & frm100118_1.txt1(2) 'Add By Sindy 2010/11/16
   Else
      strSQL5 = strSQL5 & " and sp03='0' "
      pub_QL05 = pub_QL05 & "-0" 'Add By Sindy 2010/11/16
   End If
   If Len(Trim(frm100118_1.txt1(3))) <> 0 Then
      strSQL5 = strSQL5 & " AND SP04='" & frm100118_1.txt1(3) & "' "
      pub_QL05 = pub_QL05 & "-" & frm100118_1.txt1(3) 'Add By Sindy 2010/11/16
   Else
      strSQL5 = strSQL5 & " and sp04='00' "
      pub_QL05 = pub_QL05 & "-00" 'Add By Sindy 2010/11/16
   End If
Else
    '發文號
    If frm100118_1.Option1(1).Value = True Then
        'strSQL = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,CP05 AS 收文日,CP09 AS 總收文號,CP10 AS 案件性質,CP28 AS 發文字號,CP27 AS 發文日,CP14 AS 承辦人,CP13 AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS WHERE CP28='" & frm100118_1.txt1(4) & "' AND CP01='TM' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) "
        strSQL5 = strSQL5 & " and CP28='" & frm100118_1.txt1(4) & "' "
        pub_QL05 = pub_QL05 & ";" & frm100118_1.Option1(1).Caption & frm100118_1.txt1(4) 'Add By Sindy 2010/11/16
    Else
        'BTTM
        If frm100118_1.Option1(2).Value = True Then
            'strSQL = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,CP05 AS 收文日,CP09 AS 總收文號,CP10 AS 案件性質,CP28 AS 發文字號,CP27 AS 發文日,CP14 AS 承辦人,CP13 AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP50='" & frm100118_1.txt1(5) & "' AND SP01='TM' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) "
            strSQL5 = strSQL5 & " and SP50='" & frm100118_1.txt1(5) & "' "
            pub_QL05 = pub_QL05 & ";" & frm100118_1.Option1(2).Caption & frm100118_1.txt1(5) 'Add By Sindy 2010/11/16
        Else
            'CCC CODE
            If frm100118_1.Option1(3).Value = True Then
                'strSQL = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,CP05 AS 收文日,CP09 AS 總收文號,CP10 AS 案件性質,CP28 AS 發文字號,CP27 AS 發文日,CP14 AS 承辦人,CP13 AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS WHERE "
                strSQL5 = strSQL5 + " and ( "
                strTemp = Split(frm100118_1.txt1(6), ",")
                For i = 0 To UBound(strTemp)
                    strSQL5 = strSQL5 + " instr(SP24, '" & strTemp(i) & "')>0 "
                    If i <> UBound(strTemp) Then
                        strSQL5 = strSQL5 + " OR "
                    End If
                Next i
                strSQL5 = strSQL5 + " ) "
                'strSQL = strSQL + " AND SP01='TM' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) "
                pub_QL05 = pub_QL05 & ";" & frm100118_1.Option1(3).Caption & frm100118_1.txt1(6) 'Add By Sindy 2010/11/16
            End If
        End If
    End If
End If

'Modify By Sindy 2012/5/17 +if
'查詢發文號時由進度檔串回基本檔
If frm100118_1.Option1(1).Value = True Then
   StrSQL7 = " CP01='TM' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
'否則均由基本檔串回進度檔
Else
   StrSQL7 = " SP01='TM' AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
End If
'Modify By Cheng 2002/04/26
'若已閉卷, 則在本所案號後加"*"號
'strSQL = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP09 AS 總收文號,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質,CP28 AS 發文字號,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS,casepropertymap,staff s1,staff s2 WHERE " & StrSQL7 & StrSQL5
'2010/9/15 MODIFY BY SONIA 日期欄改百年日期排序問題
strSql = "SELECT '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP09 AS 總收文號,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質,CP28 AS 發文字號,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,SP50 AS BTTM,SP24 AS CCC_Code FROM SERVICEPRACTICE,CASEPROGRESS,casepropertymap,staff s1,staff s2 WHERE " & StrSQL7 & strSQL5
strSql = strSql + " ORDER BY 本所案號,收文日 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/16
    cmdOK(0).Enabled = True
    cmdOK(1).Enabled = True
Else
    InsertQueryLog (0)  'Add By Sindy 2010/11/16
    cmdOK(0).Enabled = False
    cmdOK(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
CheckOC
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 6) = Me.grdDataList.TextMatrix(ii, 6) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 5), "1")
Next ii
Me.grdDataList.Visible = True
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100118_2 = Nothing
End Sub

Private Sub grdDataList_SelChange()
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
grdDataList.Visible = True
End Sub
