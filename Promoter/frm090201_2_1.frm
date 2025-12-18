VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090201_2_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護_專利相關案件"
   ClientHeight    =   4656
   ClientLeft      =   3984
   ClientTop       =   2268
   ClientWidth     =   7944
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4656
   ScaleWidth      =   7944
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   1
      Left            =   5040
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完整卷宗"
      Height          =   345
      Index           =   4
      Left            =   5820
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4125
      Left            =   45
      TabIndex        =   3
      Top             =   450
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   7260
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(U)"
      Height          =   345
      Index           =   0
      Left            =   6765
      TabIndex        =   2
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frm090201_2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/28 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strSql As String
Dim RS090201 As New ADODB.Recordset
Dim s As Integer
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2014/1/14
Public cmdState As Integer '紀錄作用按鍵


'Add By Sindy 2014/1/14
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Modify By Sindy 2016/2/22
Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'Modify By Sindy 2016/2/22
Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim StrTag As String

Select Case cmdState
Case 0 '回前畫面
   'Modify By Sindy 2014/1/14
   'Modify By Sindy 2023/10/17
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
      m_PrevForm.Show
   End If
'   If UCase(m_PrevForm.Name) = UCase("frm090202_2") Then
'      frm090202_2.Show
'   Else
'   '2014/1/14 END
'      frm090201_2.Show
'   End If
   '2023/10/17 END
   Unload Me
Case 1 '進度
   Me.Enabled = False
   StrTag = ""
   For i = 1 To grd1.Rows - 1
      grd1.col = 0
      grd1.row = i
      If Trim(grd1.Text) = "V" Then
         grd1.col = 0
         grd1.Text = ""
         For j = 0 To grd1.Cols - 1
             grd1.col = j
             grd1.CellBackColor = QBColor(15)
         Next j
         grd1.col = 1
         StrTag = grd1.Text
         If Not IsNull(grd1.Text) Then
            fnCloseAllFrm100 'Added by Morgan 2016/2/22
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(StrTag)
            frm100101_2.cmdOK(15).Visible = False '承辦歷程-聯絡
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
Case 4 '完整卷宗
   Me.Enabled = False
   StrTag = ""
   For i = 1 To grd1.Rows - 1
      grd1.col = 0
      grd1.row = i
      If Trim(grd1.Text) = "V" Then
         grd1.col = 0
         grd1.Text = ""
         For j = 0 To grd1.Cols - 1
             grd1.col = j
             grd1.CellBackColor = QBColor(15)
         Next j
         grd1.col = 1
         StrTag = grd1.Text
         If Not IsNull(grd1.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_L.m_strKey = Pub_RplStr(StrTag)
            frm100101_L.SetParent Me
            If frm100101_L.QueryData = True Then
               frm100101_L.Show
               Me.Hide
            Else
               Unload frm100101_L
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_PrevForm = Nothing 'Add By Sindy 2014/1/14
Set frm090201_2_1 = Nothing
End Sub

'傳入本所案號
'用本所案號串 caseMap 的  cm05~08 且 cm10='0' (國外案)，再用 cm01~04 串案件進度檔串基本檔
Sub StrMenu(strText As String)
'Modified by Morgan 2011/12/22 改和共同查詢的專利相關他國案一樣 Ex.P-098685無法看到P-098684(一案二請案件)
'strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent                                          WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and cm01 in (" & SQLGrpStr("", 1) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=tm01(+) and cm02=tm02(+) and cm03=tm03(+) and cm04=tm04(+) and cm01 in (" & SQLGrpStr("", 2) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=lc01(+) and cm02=lc02(+) and cm03=lc03(+) and cm04=lc04(+) and cm01 in (" & SQLGrpStr("", 3) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=hc01(+) and cm02=hc02(+) and cm03=hc03(+) and cm04=hc04(+) and cm01 in (" & SQLGrpStr("", 4) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=sp01(+) and cm02=sp02(+) and cm03=sp03(+) and cm04=sp04(+) and cm01 in (" & SQLGrpStr("", 5) & ")  "
''add by nick 2005/02/17 陳玲玲填請做單要國內外皆可
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent             WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and cm05 in (" & SQLGrpStr("", 1) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=tm01(+) and cm06=tm02(+) and cm07=tm03(+) and cm08=tm04(+) and cm05 in (" & SQLGrpStr("", 2) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=lc01(+) and cm06=lc02(+) and cm07=lc03(+) and cm08=lc04(+) and cm05 in (" & SQLGrpStr("", 3) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=hc01(+) and cm06=hc02(+) and cm07=hc03(+) and cm08=hc04(+) and cm05 in (" & SQLGrpStr("", 4) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=sp01(+) and cm06=sp02(+) and cm07=sp03(+) and cm08=sp04(+) and cm05 in (" & SQLGrpStr("", 5) & ") "
''2008/6/2 add by sonia 香港大陸案之關聯也要可查詢
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent                                          WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='4' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and cm01 in (" & SQLGrpStr("", 1) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='4' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=tm01(+) and cm02=tm02(+) and cm03=tm03(+) and cm04=tm04(+) and cm01 in (" & SQLGrpStr("", 2) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='4' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=lc01(+) and cm02=lc02(+) and cm03=lc03(+) and cm04=lc04(+) and cm01 in (" & SQLGrpStr("", 3) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='4' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=hc01(+) and cm02=hc02(+) and cm03=hc03(+) and cm04=hc04(+) and cm01 in (" & SQLGrpStr("", 4) & ") "
'strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='4' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=sp01(+) and cm02=sp02(+) and cm03=sp03(+) and cm04=sp04(+) and cm01 in (" & SQLGrpStr("", 5) & ")  "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent                                          WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='4' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and cm05 in (" & SQLGrpStr("", 1) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='4' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=tm01(+) and cm06=tm02(+) and cm07=tm03(+) and cm08=tm04(+) and cm05 in (" & SQLGrpStr("", 2) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='4' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=lc01(+) and cm06=lc02(+) and cm07=lc03(+) and cm08=lc04(+) and cm05 in (" & SQLGrpStr("", 3) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='4' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=hc01(+) and cm06=hc02(+) and cm07=hc03(+) and cm08=hc04(+) and cm05 in (" & SQLGrpStr("", 4) & ") "
'strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='4' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=sp01(+) and cm06=sp02(+) and cm07=sp03(+) and cm08=sp04(+) and cm05 in (" & SQLGrpStr("", 5) & ")   order by 1  "
'
''因為會有多個工程師，故暫不秀
''strSQL = "SELECT distinct pa01||'-'||pa02||'-'||pa03||'-'||pa04,pa05 FROM CASEPROGRESS,patent,r100101_h                                          WHERE r001001=pa01 and r001002=pa02 and r001003=pa03 and r001004=pa04 and r001001=cp01 and r001002=cp02 and r001003=cp03 and r001004=cp04 and id='" & strUserNum & "' and pa01 in (" & SQLGrpStr("", 1) & ") "
''strSQL = strSQL & " union SELECT tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm05 FROM CASEPROGRESS,trademark,r100101_h       WHERE r001001=tm01 and r001002=tm02 and r001003=tm03 and r001004=tm04 and r001001=cp01 and r001002=cp02 and r001003=cp03 and r001004=cp04  and id='" & strUserNum & "'  and tm01 in (" & SQLGrpStr("", 2) & ") "
''strSQL = strSQL & " union SELECT lc01||'-'||lc02||'-'||lc03||'-'||lc04,lc05 FROM CASEPROGRESS,lawcase,r100101_h            WHERE r001001=lc01 and r001002=lc02 and r001003=lc03 and r001004=lc04 and r001001=cp01 and r001002=cp02 and r001003=cp03 and r001004=cp04  and id='" & strUserNum & "'  and lc01 in (" & SQLGrpStr("", 3) & ") "
''strSQL = strSQL & " union SELECT hc01||'-'||hc02||'-'||hc03||'-'||hc04,hc06 FROM CASEPROGRESS,hirecase,r100101_h           WHERE r001001=hc01 and r001002=hc02 and r001003=hc03 and r001004=hc04 and r001001=cp01 and r001002=cp02 and r001003=cp03 and r001004=cp04  and id='" & strUserNum & "' and hc01 in (" & SQLGrpStr("", 4) & ") "
''strSQL = strSQL & " union SELECT sp01||'-'||sp02||'-'||sp03||'-'||sp04,sp05 FROM CASEPROGRESS,servicepractice,r100101_h WHERE r001001=sp01 and r001002=sp02 and r001003=sp03 and r001004=sp04 and r001001=cp01 and r001002=cp02 and r001003=cp03 and r001004=cp04  and id='" & strUserNum & "' and sp01 in (" & SQLGrpStr("", 5) & ") "
'Modified by Morgan 2023/10/24
'cnnConnection.Execute "delete from r100101_h where id='" & strUserNum & "' "
'cnnConnection.Execute "insert into r100101_h select '" & SystemNumber(strText, 1) & "','" & SystemNumber(strText, 2) & "','" & SystemNumber(strText, 3) & "','" & SystemNumber(strText, 4) & "',0,'1','" & strUserNum & "' from dual "
'cnnConnection.Execute "insert into r100101_h select '" & SystemNumber(strText, 1) & "','" & SystemNumber(strText, 2) & "','" & SystemNumber(strText, 3) & "','" & SystemNumber(strText, 4) & "',0,'2','" & strUserNum & "' from dual "
'cnnConnection.Execute "begin db_r100101_h('" & strUserNum & "'); end;"
''Modify By Sindy 2016/2/22 +'',
''Modified by Morgan 2023/10/23 承辦人改先抓A類收文之新申請案的案件性質,沒有才抓CP31=Y的進度--秀玲
''strSql = "select distinct '',replace(decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),'N---','')" & _
   ",pa05,ST02,nvl(ptm03,ptm04),nvl(na03,na04),r001005 as bysort,r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A" & _
   " from r100101_h,patent,patenttrademarkmap,nation,caseprogress,staff" & _
   " where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+)" & _
   " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp31(+)='Y' and st01(+)=cp14"
'strSql = "select distinct '',replace(decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),'N---','')" & _
   ",pa05,nvl(s2.ST02,s1.ST02) st02,nvl(ptm03,ptm04),nvl(na03,na04),r001005 as bysort,r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A" & _
   " from r100101_h,patent,patenttrademarkmap,nation,caseprogress c1,staff s1,caseprogress c2,staff s2" & _
   " where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+)" & _
   " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04 and c1.cp31(+)='Y' and s1.st01(+)=c1.cp14" & _
   " and c2.cp01(+)=pa01 and c2.cp02(+)=pa02 and c2.cp03(+)=pa03 and c2.cp04(+)=pa04 and c2.cp09(+)<'B' and instr('" & NewCasePtyList & "',c2.cp10(+))>0 and s2.st01(+)=c2.cp14"
''end 2023/10/23
'strSql = strSql & " order by bysort,A "
strSql = PUB_GetPatRefCaseSQL(strText)
'end 2011/12/22
Set RS090201 = New ADODB.Recordset
RS090201.CursorLocation = adUseClient
RS090201.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If RS090201.RecordCount <> 0 Then
    Set grd1.Recordset = RS090201
    SetGrd1
Else
    s = MsgBox("沒有國內外案關聯資料！", , "沒有資料！")
    'Modify By Sindy 2023/10/17
    'frm090201_2.Show
    If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
      m_PrevForm.Show
    End If
    '2023/10/17 END
    Unload Me
End If
Set RS090201 = Nothing
End Sub

Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 6 '5
    .row = 0
    'Add By Sindy 2016/2/22
    .col = 0:   .Text = "V"
    .ColWidth(0) = 200
    .CellAlignment = flexAlignCenterCenter
    '2016/2/22 END
    .col = 1:   .Text = "本所案號"
    .ColWidth(1) = 1600
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "案件名稱"
    .ColWidth(2) = 3300
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "承辦人"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    'Added by Morgan 2011/12/22
    .col = 4:   .Text = "種類"
    .ColWidth(4) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "申請國家"
    .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
    'end 2011/12/22
    .Visible = True
End With
End Sub

'Add By Sindy 2016/2/22
Private Sub grd1_SelChange()
Dim i As Integer

grd1.Visible = False
grd1.row = grd1.MouseRow
'空白不勾
grd1.col = 1
If Trim(grd1.Text) <> "" Then
   grd1.col = 0
   If grd1.row <> 0 Then
      If grd1.Text = "V" Then
         grd1.Text = ""
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = QBColor(15)
         Next i
      Else
         grd1.Text = "V"
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
grd1.Visible = True
End Sub
