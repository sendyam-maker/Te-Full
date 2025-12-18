VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090112_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "組群統計"
   ClientHeight    =   4764
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7608
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4764
   ScaleWidth      =   7608
   Begin VB.CommandButton Command2 
      Caption         =   "每個組群重新編順序"
      Height          =   345
      Left            =   2190
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm090112_1.frx":0000
      Left            =   2700
      List            =   "frm090112_1.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   540
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3795
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   7485
      _ExtentX        =   13208
      _ExtentY        =   6689
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm090112_1.frx":001A
      Left            =   780
      List            =   "frm090112_1.frx":001C
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   375
      Left            =   6420
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組群："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frm090112_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'add by nick 2004/11/10
Option Explicit
Public UpForm As Form    '上一層
Public strTM09 As String    '接受組群
Public StrDateStart As String    '查名時間起   (西元格式)
Public StrDateEnd As String    '查名時間迄   (西元格式)
Dim SeekCbo1 As String
Dim SeekCbo2 As String

Private Sub cmdOK_Click()
UpForm.Show
Unload Me
End Sub

Private Sub Combo1_Click()
If SeekCbo1 <> Combo1.Text And Combo1.Enabled = True Then
    'Modified by Lydia 2015/06/16 使用查名新規則
    'QueryData
    NewQueryData
    
    SeekCbo1 = Combo1.Text
End If
End Sub

Private Sub Combo2_Click()
If SeekCbo2 <> Combo2.Text And Combo2.Enabled = True Then
    'Modified by Lydia 2015/06/16 使用查名新規則
    'QueryData
    NewQueryData
    
    SeekCbo2 = Combo2.Text
End If
End Sub

'add by nickc 2008/03/06  葉大給順序(3/5 mail)
'林嘉雯84027
'林忱瑾78027
'程序
'林淑鈴69001
'簡玉滿69004
Private Sub Command2_Click()
Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_Point As Integer
Dim m_seek(1 To 10) As String
Dim m_i As Integer
Dim m_k As Integer
Screen.MousePointer = vbHourglass
'2010/6/14 modify by sonia 因加入99020故重排順序,原為84027,78027,70003,69001,69004
m_seek(1) = "78027"
m_seek(2) = "99020"
m_seek(3) = "69001"
m_seek(4) = "84027"
m_seek(5) = "69004"
m_seek(6) = "70003"
If m_rs.State = 1 Then m_rs.Close
'Memo by Lydia 2023/01/12 電子化查名單已不使用tmqctl，僅供查詢到2015
m_str = "select distinct tmqc02 from tmqctl where length(tmqc02)=2 order by tmqc02 "
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    m_rs.MoveFirst
    m_Point = 1
    Do While Not m_rs.EOF
        For m_i = 1 To 6
            m_k = (m_i - 1 + m_Point) Mod 6
            If m_k = 0 Then m_k = 6
            m_str = "update tmqctl set tmqc04='" & Trim(m_i) & "',tmqc13='" & Trim(m_i) & "' where tmqc02='" & CheckStr(m_rs.Fields(0)) & "' and tmqc01='" & m_seek(m_k) & "' "
            'Debug.Print m_str
            cnnConnection.Execute m_str
        Next m_i
        m_Point = m_Point + 1
        m_rs.MoveNext
    Loop
End If
If m_rs.State = 1 Then m_rs.Close
m_str = "select distinct  tmqc02 from tmqctl where length(tmqc02)=4 order by tmqc02 "
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    m_rs.MoveFirst
    m_Point = 1
    Do While Not m_rs.EOF
        For m_i = 1 To 6
            m_k = (m_i - 1 + m_Point) Mod 6
            If m_k = 0 Then m_k = 6
            m_str = "update tmqctl set tmqc04='" & Trim(m_i) & "',tmqc13='" & Trim(m_i) & "' where tmqc02='" & CheckStr(m_rs.Fields(0)) & "' and tmqc01='" & m_seek(m_k) & "' "
            'Debug.Print m_str
            cnnConnection.Execute m_str
        Next m_i
        m_Point = m_Point + 1
        m_rs.MoveNext
    Loop
End If
Screen.MousePointer = vbDefault
MsgBox "OK!!"
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me '將畫面移至中央
   'add by nickc 2007/08/15
   'Modified by Lydia 2015/06/16 隱藏按鈕
'   If GetStaffDepartment(strUserNum) = "M51" Then
'        '2011/8/22 cancel by sonia 更新特定人的加減次數tmqc03,但此欄已不再使用故取消顯示
'        'Me.Command1.Visible = True
'        Me.Command2.Visible = True
'    End If
   ' Me.Command1.Visible = False: Me.Command2.Visible = False 'Mark by Lydia 2023/01/12
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090112_1 = Nothing
End Sub

Public Sub QueryData()
Dim strSql As String
'edit by nickc 2008/03/06  修改成新規則
'Dim TxtOrPic1 As String
'Dim TxtOrPic2 As String
'Dim TxtOrPic3 As String
'Dim TxtOrPic4 As String
Dim ClassStr As String
'edit by nickc 2008/03/06  修改成新規則
'Dim ClassStr21 As String
'Dim ClassStr22 As String
'Dim SpecStr11 As String
'Dim SpecStr12 As String
'Dim SpecStr13 As String
'Dim SpecStr14 As String
'Dim SpecStr15 As String
'Dim SpecStr16 As String
'Dim SpecStr17 As String
'Dim SpecStr18 As String
'Dim SpecStr21 As String
'Dim SpecStr22 As String
'Dim SpecStr23 As String
'Dim SpecStr24 As String
'Dim SpecStr25 As String
'Dim SpecStr26 As String
'add by nickc 2008/03/06
Dim SpecStr As String

   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   'ClassStr = Trim(Combo1.Text)
   'edit by nickc 2008/03/06  修改成新規則
   'ClassStr = IIf(Combo2.Text = "文字", Trim(Combo1.Text), GetRelClass(Trim(Combo1.Text)))
   'SpecStr11 = ""
   'SpecStr12 = ""
   'SpecStr13 = ""
   'SpecStr14 = ""
   'SpecStr15 = ""
   'SpecStr16 = ""
   'SpecStr17 = ""
   'SpecStr18 = ""
   'SpecStr21 = ""
   'SpecStr22 = ""
   'SpecStr23 = ""
   'SpecStr24 = ""
   'edit by nickc 2008/03/06  修改成新規則
   'TxtOrPic1 = IIf(Combo2.Text = "文字", "sum(nvl(G1.tmq07,0)+nvl(G1.tmq08,0))", "sum(nvl(G1.tmq09,0))")
   'TxtOrPic2 = IIf(Combo2.Text = "文字", "sum(nvl(G3.tmq07,0)+nvl(G3.tmq08,0))", "sum(nvl(G3.tmq09,0))")
   'TxtOrPic3 = IIf(Combo2.Text = "文字", "sum(nvl(G5.tmq07,0)+nvl(G5.tmq08,0))", "sum(nvl(G5.tmq09,0))")
   'TxtOrPic4 = IIf(Combo2.Text = "文字", "sum(nvl(G7.tmq07,0)+nvl(G7.tmq08,0))", "sum(nvl(G7.tmq09,0))")
   'SpecStr11 = SpecStr11 & IIf(Combo2.Text = "文字", " and (nvl(G1.tmq07,0)+nvl(G1.tmq08,0)) > 0 ", " and (nvl(G1.tmq09,0)) > 0 ")
   'SpecStr12 = SpecStr12 & IIf(Combo2.Text = "文字", " and (nvl(H1.tmq07,0)+nvl(H1.tmq08,0)) > 0 ", " and (nvl(H1.tmq09,0)) > 0 ")
   'SpecStr13 = SpecStr13 & IIf(Combo2.Text = "文字", "  (nvl(G3.tmq07,0)+nvl(G3.tmq08,0)) > 0 ", "  (nvl(G3.tmq09,0)) > 0 ")
   'SpecStr14 = SpecStr14 & IIf(Combo2.Text = "文字", " and (nvl(G4.tmq07,0)+nvl(G4.tmq08,0)) > 0 ", " and (nvl(G4.tmq09,0)) > 0 ")
   'SpecStr15 = SpecStr15 & IIf(Combo2.Text = "文字", " and (nvl(G5.tmq07,0)+nvl(G5.tmq08,0)) > 0 ", " and (nvl(G5.tmq09,0)) > 0 ")
   'SpecStr16 = SpecStr16 & IIf(Combo2.Text = "文字", " and (nvl(H2.tmq07,0)+nvl(H2.tmq08,0)) > 0 ", " and (nvl(H2.tmq09,0)) > 0 ")
   'SpecStr17 = SpecStr17 & IIf(Combo2.Text = "文字", "  (nvl(G7.tmq07,0)+nvl(G7.tmq08,0)) > 0 ", "  (nvl(G7.tmq09,0)) > 0 ")
   'SpecStr18 = SpecStr18 & IIf(Combo2.Text = "文字", " and (nvl(G8.tmq07,0)+nvl(G8.tmq08,0)) > 0 ", " and (nvl(G8.tmq09,0)) > 0 ")
   
   If StrDateStart <> "" Then
   'edit by nickc 2008/03/06  修改成新規則
   '    SpecStr21 = SpecStr21 & " and G1.tmq11>=" & StrDateStart & " "
   '    SpecStr22 = SpecStr22 & " and G4.tmq11>=" & StrDateStart & " "
   '    SpecStr23 = SpecStr23 & " and G5.tmq11>=" & StrDateStart & " "
   '    SpecStr24 = SpecStr24 & " and G8.tmq11>=" & StrDateStart & " "
   SpecStr = SpecStr & " and DD.tmq13>=" & StrDateStart & " "
   End If
   If StrDateEnd <> "" Then
   'edit by nickc 2008/03/06  修改成新規則
   '    SpecStr21 = SpecStr21 & " and G1.tmq11<=" & StrDateEnd & " "
   '    SpecStr22 = SpecStr22 & " and G4.tmq11<=" & StrDateEnd & " "
   '    SpecStr23 = SpecStr23 & " and G5.tmq11<=" & StrDateEnd & " "
   '    SpecStr24 = SpecStr24 & " and G8.tmq11<=" & StrDateEnd & " "
   SpecStr = SpecStr & " and DD.tmq13<=" & StrDateEnd & " "
   End If
   If StrDateStart <> "" Or StrDateEnd <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090112.Label2 & frm090112.Txt1(1) & "-" & frm090112.Txt1(2) 'Add By Sindy 2010/12/13
   End If
   pub_QL05 = pub_QL05 & ";" & frm090112.Label1 & frm090112.Txt1(0)  'Add By Sindy 2010/12/13
   'edit by nickc 2008/03/06  修改成新規則
   'ClassStr21 = IIf(Combo2.Text = "文字", " and G3.tmq03 like '" & Mid(ClassStr, 1, 2) & "%' ", " and G3.tmq03 like '" & ClassStr & "%'  ")
   'ClassStr22 = IIf(Combo2.Text = "文字", " and G7.tmq03 like '" & Mid(ClassStr, 1, 2) & "%' ", " and G7.tmq03 like '" & ClassStr & "%'  ")
   ClassStr = IIf(Combo2.Text = "文字", " and DD.tmq03 like '" & Mid(Trim(Combo1.Text), 1, 2) & "%' and (nvl(DD.tmq07,0)+nvl(DD.tmq08,0)) > 0", " and DD.tmq03 like '" & Trim(Combo1.Text) & "%'  and (nvl(DD.tmq09,0)) > 0 ")
   
   'strSQL = "select st02,nvl(D.YesCountDT,0) + decode(TMQC03,null,0,tmqc03) ,nvl(A.YesCount,0),B.firstCount,C.noFlag,tmqc12,C.noCount,decode(TMQC03,null,0,tmqc03),st03,st01 from (" & _
   '            " select tmq10," & TxtOrPic & " as YesCount from trademarkquery where tmq03 like '%" & ClassStr & "%'  and tmq04>=20040301 and tmq11 is not null " & SpecStr1 & SpecStr2 & " group by tmq10) A, " & _
   '            " (select tmq10,'*' firstCount  from trademarkquery where tmq01 in (select min(tmq01) from trademarkquery where tmq04>=20040301 and tmq03 like '%" & ClassStr & "%' and tmq11 is not null " & SpecStr1 & "  )) B," & _
   '            " (select Db.tmq10 as tmq10,decode(Db.Da,0,'','#') as noFlag,Db.Da as noCount from (select tmq10," & TxtOrPic & " as Da from trademarkquery where tmq11 is null and tmq04>=20040301 " & SpecStr1 & ClassStr2 & "  group by tmq10 ) Db) C," & _
   '            " (select tmq10,count(*) as YesCountDT from (select distinct tmq10,tmq11 from trademarkquery where tmq03 like '" & ClassStr & "%'  and tmq04>=20040301 and tmq11 is not null  " & SpecStr1 & SpecStr2 & " ) DC group by tmq10) D, " & _
   '            " tmqctl , staff " & _
   '            " where st01=b.tmq10(+) and st01=C.tmq10(+) and st01=tmqc01(+) and st01=D.tmq10(+)  and st01 = A.tmq10(+) and '" & IIf(Combo2.Text = "文字", Mid(ClassStr, 1, 2), ClassStr) & "'=tmqc02(+) and st04='1' and st05 in ('93','95') and st01 <> 'TM4' "
   'strSQL = strSQL & "union select '程序人員',sum(D.YesCountDT) + max(decode(TMQC03,null,0,tmqc03)),sum(A.YesCount),max(B.firstCount),max(C.noFlag),max(tmqc12),sum(C.noCount),max(decode(TMQC03,null,0,tmqc03)),'P22' as st03,'XXXXXX' as st01 from (" & _
   '            " select tmq10," & TxtOrPic & " as YesCount from trademarkquery where tmq03 like '%" & ClassStr & "%'  and tmq04>=20040301 and tmq11 is not null " & SpecStr1 & SpecStr2 & " group by tmq10) A, " & _
   '            " (select tmq10,'*' firstCount  from trademarkquery where tmq01 in (select min(tmq01) from trademarkquery where tmq04>=20040301 and tmq03 like '%" & ClassStr & "%' and tmq11 is not null " & SpecStr1 & "  )) B," & _
   '            " (select Db.tmq10 as tmq10,decode(Db.Da,0,'','#') as noFlag,Db.Da as noCount from (select tmq10," & TxtOrPic & " as Da from trademarkquery where tmq11 is null and tmq04>=20040301 " & SpecStr1 & ClassStr2 & "  group by tmq10 ) Db) C," & _
   '            " (select tmq10,count(*) as YesCountDT from (select distinct tmq10,tmq11 from trademarkquery where tmq03 like '" & ClassStr & "%'  and tmq04>=20040301 and tmq11 is not null  " & SpecStr1 & SpecStr2 & " ) DC group by tmq10) D, " & _
   '            " tmqctl , staff " & _
   '            " where st01=b.tmq10(+) and st03='P22' and st01=C.tmq10(+) and st01=tmqc01(+) and st01=D.tmq10(+) and st01=A.tmq10(+) and '" & IIf(Combo2.Text = "文字", Mid(ClassStr, 1, 2), ClassStr) & "'=tmqc02(+) and st04='1' " & _
   '            " order by st03,st01 "
   'edit by nickc 2008/03/06  修改成新規則
   'strSQL = "select st02,nvl(D.YesCountDT,0) + decode(TMQC03,null,0,tmqc03) ,nvl(A.YesCount,0),B.firstCount,C.noFlag,tmqc12,C.noCount,decode(TMQC03,null,0,tmqc03),st03,st01 from (" & _
   '            " select G1.tmq10," & TxtOrPic1 & " as YesCount from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G1 where G1.tmq03 like '" & ClassStr & "%'  and G1.tmq04>=20040301 and G1.tmq11 is not null " & SpecStr11 & SpecStr21 & " group by G1.tmq10) A, " & _
   '            " (select G2.tmq10,'*' firstCount  from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G2 where G2.tmq01 in (select min(H1.tmq01) from trademarkquery H1 where H1.tmq04>=20040301 and H1.tmq03 like '" & ClassStr & "%' and H1.tmq11 is not null " & SpecStr12 & "  )) B," & _
   '            " (select Db1.tmq10 as tmq10,decode(Db1.Da,0,'','#') as noFlag,Db1.Da as noCount from (select G3.tmq10," & TxtOrPic2 & " as Da from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) and tmq11 is null and tmq04>=20040301) G3 where  " & SpecStr13 & ClassStr21 & "  group by G3.tmq10 ) Db1) C," & _
   '            " (select Dc1.tmq10,count(*) as YesCountDT from (select distinct G4.tmq10,G4.tmq11 from (select tmq01,tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G4 where G4.tmq03 like '" & ClassStr & "%'  and G4.tmq04>=20040301 and G4.tmq11 is not null  " & SpecStr14 & SpecStr22 & " ) DC1 group by Dc1.tmq10) D, " & _
   '            " tmqctl , staff " & _
   '            " where st01=b.tmq10(+) and st01=C.tmq10(+) and st01=tmqc01(+) and st01=D.tmq10(+)  and st01 = A.tmq10(+) and '" & IIf(Combo2.Text = "文字", Mid(ClassStr, 1, 2), ClassStr) & "'=tmqc02(+) and st04='1' and st05 in ('93','95') and st01 <> 'TM4' "
   'strSQL = strSQL & "union select '程序人員',sum(nvl(D.YesCountDT,0)) + max(decode(TMQC03,null,0,tmqc03)),sum(A.YesCount),max(B.firstCount),max(C.noFlag),max(tmqc12),sum(C.noCount),max(decode(TMQC03,null,0,tmqc03)),'P22' as st03,'XXXXXX' as st01 from (" & _
   '            " select G5.tmq10," & TxtOrPic3 & " as YesCount from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G5 where G5.tmq03 like '" & ClassStr & "%'  and G5.tmq04>=20040301 and G5.tmq11 is not null " & SpecStr15 & SpecStr23 & " group by G5.tmq10) A, " & _
   '            " (select G6.tmq10,'*' firstCount  from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G6 where G6.tmq01 in (select min(H2.tmq01) from trademarkquery H2 where H2.tmq04>=20040301 and H2.tmq03 like '" & ClassStr & "%' and H2.tmq11 is not null " & SpecStr16 & "  )) B," & _
   '            " (select Db2.tmq10 as tmq10,decode(Db2.Da,0,'','#') as noFlag,Db2.Da as noCount from (select G7.tmq10," & TxtOrPic4 & " as Da from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) and tmq11 is null and tmq04>=20040301) G7 where  " & SpecStr17 & ClassStr22 & "  group by G7.tmq10 ) Db2) C," & _
   '            " (select Dc2.tmq10,count(*) as YesCountDT from (select distinct G8.tmq10,G8.tmq11 from (select tmq01, tmq10,tmq11,decode(sign(nvl(tmq09,0)),0,substr(tmq03,1,4),nvl(tmqm02,substr(tmq03,1,4))) tmq03,tmq04,tmq07,tmq08,tmq09 from  trademarkquery,tmqmap where substr(tmq03,1,4)=tmqm01(+) ) G8 where G8.tmq03 like '" & ClassStr & "%'  and G8.tmq04>=20040301 and G8.tmq11 is not null  " & SpecStr18 & SpecStr24 & " ) DC2 group by Dc2.tmq10) D, " & _
   '            " tmqctl , staff " & _
   '            " where st01=b.tmq10(+) and st03='P22' and st01=C.tmq10(+) and st01=tmqc01(+) and st01=D.tmq10(+) and st01=A.tmq10(+) and '" & IIf(Combo2.Text = "文字", Mid(ClassStr, 1, 2), ClassStr) & "'=tmqc02(+) and st04='1' " & _
   '            "  order by st03,st01"
   '2011/9/13 modify by sonia 加A0036至70003
   strSql = "select decode(tmqc01,'70003','程序人員',st02)||'('||st01||')',nvl(BB.tmq01,0),nvl(BB.tmqcnt,0)," & IIf(Combo2.Text = "文字", "tmqc13", "tmqc04") & " from tmqctl,(select tmq10,count(tmq01) tmq01,tmq03,sum(tmqcnt) tmqcnt from (select decode(CC.tmq10,'79041','70003','73014','70003','A0036','70003',CC.tmq10) tmq10,DD.tmq01," & IIf(Combo2.Text = "文字", "substr(DD.tmq03,1,2)", "substr(DD.tmq03,1,4)") & "  tmq03," & IIf(Combo2.Text = "文字", "nvl(DD.tmq07,0)+nvl(DD.tmq08,0)", "nvl(DD.tmq09,0)") & "  tmqcnt from trademarkquery DD,(select distinct tmq10 from trademarkquery ,staff where tmq10=st01 and st04='1') CC where CC.tmq10=DD.tmq10(+) " & ClassStr & SpecStr & " ) AA group by tmq10,tmq03) BB,staff " & _
                        " Where  " & IIf(Combo2.Text = "文字", " '" & Mid(Trim(Combo1.Text), 1, 2) & "'", " '" & Trim(Combo1.Text) & "'") & "=tmqc02(+) and tmqc02=BB.tmq03(+) and tmqc01=BB.tmq10(+) and tmqc01=st01(+)  "
   strSql = strSql & "union all select st02||'('||st01||')',nvl(BB.tmq01,0),nvl(BB.tmqcnt,0),decode(st01,'73014',6,'79041',6,7) from (select tmq10,count(tmq01) tmq01,tmq03,sum(tmqcnt) tmqcnt from (select CC.tmq10 tmq10,DD.tmq01," & IIf(Combo2.Text = "文字", "substr(DD.tmq03,1,2)", "substr(DD.tmq03,1,4)") & "  tmq03," & IIf(Combo2.Text = "文字", "nvl(DD.tmq07,0)+nvl(DD.tmq08,0)", "nvl(DD.tmq09,0)") & "  tmqcnt from trademarkquery DD,(select distinct tmq10 from trademarkquery ,staff where tmq10=st01 and st04='1' and st03='P22') CC where CC.tmq10=DD.tmq10(+) " & ClassStr & SpecStr & " ) AA group by tmq10,tmq03) BB,staff " & _
                        " Where  st01=tmq10(+) and st03='P22' and st04='1'  and (" & IIf(Combo2.Text = "文字", " '" & Mid(Trim(Combo1.Text), 1, 2) & "'", " '" & Trim(Combo1.Text) & "'") & ",'70003') in (select tmqc02,tmqc01 from tmqctl ) "
   strSql = strSql & " order by 4 "
   
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
       Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/13
       End If
       'If .RecordCount <> 0 Then
           Set grdDataList.Recordset = adoRecordset
       'End If
       SetDataListWidth
   End With
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub SetDataListWidth()
'edit by nickc 2008/03/06 改成新規則
'grdDataList.Cols = 8
grdDataList.Cols = 4
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "查名人"
grdDataList.ColWidth(0) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "查名總次數"
grdDataList.ColWidth(1) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "查名總筆數"
grdDataList.ColWidth(2) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
'edit by nickc 2008/03/06 改成新規則
'grdDataList.col = 3: grdDataList.Text = "第一筆"
'grdDataList.ColWidth(3) = 800
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 4: grdDataList.Text = "未查覆"
'grdDataList.ColWidth(4) = 800
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 5: grdDataList.Text = "查覆中"
'grdDataList.ColWidth(5) = 800
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 6: grdDataList.Text = "未查覆筆數"
'grdDataList.ColWidth(6) = 1000
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 7: grdDataList.Text = "加減次數"
'grdDataList.ColWidth(7) = 1000
'grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = ""
grdDataList.ColWidth(3) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Public Sub InToCombo()
Dim arrTM09 As Variant
Dim CboStrTm09 As String
Dim ii As Integer
Dim jj As Integer
Combo1.Clear
jj = 0
arrTM09 = Split(strTM09, ".")
For ii = 0 To UBound(arrTM09)
    If arrTM09(ii) <> "" Then
        Combo1.AddItem arrTM09(ii)
        jj = jj + 1
    End If
Next ii
Combo1.Enabled = False
Combo2.Enabled = False
Combo1.ListIndex = 0
Combo2.Enabled = True
Combo2.ListIndex = 1
If jj > 1 Then
    Combo1.Enabled = True
End If
End Sub
'Remove by Lydia 2015/06/16 使用查名新規則
'Function GetRelClass(oClass As String) As String
'CheckOC3
'With AdoRecordSet3
'    .CursorLocation = adUseClient
'    .Open "Select tmqm02 from tmqmap where tmqm01='" & oClass & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 Then
'        GetRelClass = CheckStr(.Fields(0).Value)
'    Else
'        GetRelClass = oClass
'    End If
'End With
'CheckOC3
'End Function

Public Sub NewQueryData()
Dim strMid As String
Dim ClassStr As String
Dim SpecStr As String

   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   
   If StrDateStart <> "" Then
      SpecStr = SpecStr & " and tmq13>=" & StrDateStart & " "
   End If
   If StrDateEnd <> "" Then
      SpecStr = SpecStr & " and tmq13<=" & StrDateEnd & " "
   End If
   If StrDateStart <> "" Or StrDateEnd <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090112.Label2 & frm090112.Txt1(1) & "-" & frm090112.Txt1(2)
   End If
   pub_QL05 = pub_QL05 & ";" & frm090112.Label1 & frm090112.Txt1(0)
  
   If Combo2.Text = "文字" Then
      ClassStr = " and tmq03 like '" & Mid(Trim(Combo1.Text), 1, 2) & "%' and (nvl(tmq07,0)+nvl(tmq08,0)) > 0"
      strExc(2) = ",sum(tmq07+tmq08) "
   Else
      ClassStr = " and tmq03 like '" & Trim(Combo1.Text) & "%'  and (nvl(tmq09,0)) > 0 "
      strExc(2) = ",sum(tmq09) "
   End If
   
   'Modified by Lydia 2024/11/18 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   'strMid = " select decode(tmqm02,'70003','程序人員',st02)||'('||st01||')',sum(x1.cnt)" & strExc(2) & _
            ",st01 from tmqmember,staff,(select tmq10,1 as cnt,nvl(tmq07,0) tmq07,nvl(tmq08,0) tmq08,nvl(tmq09,0) tmq09 " & _
            "from trademarkquery where 1=1 " & SpecStr & ClassStr & ") x1 " & _
            "where tmqm01=x1.tmq10(+) and tmqm02=st01(+) and st04='1' " & _
            "group by decode(tmqm02,'70003','程序人員',st02)||'('||st01||')',st01 order by st01 "
   strMid = " select decode(tmqm02,'70003','程序人員',st02)||'('||st01||')' as tmq10n ,sum(x1.cnt) as tot1 " & strExc(2) & _
            " as tot2,st01 from tmqmember,staff,(select tmq10,1 as cnt,nvl(tmq07,0) tmq07,nvl(tmq08,0) tmq08,nvl(tmq09,0) tmq09 " & _
            "from trademarkquery where 1=1 " & SpecStr & ClassStr & _
            "union all select tmq10,cnt,tmq07,tmq08,tmq09 from (select tma10 as tmq10, 1 as cnt, nvl(tma36,0) as tmq07, nvl(tma37,0) as tmq08, nvl(tma38,0) as tmq09 " & _
            ", " & PUB_GetTMAforClass & " as tmq03,to_char(tma04,'YYYYMMDD') as tmq13 from tmqappform where TO_CHAR(TMA04,'YYYYMMDD')>='20240601') where 1=1 " & SpecStr & ClassStr & _
            " ) x1 " & _
            "where tmqm01=x1.tmq10(+) and tmqm02=st01(+) and st04='1' " & _
            "group by decode(tmqm02,'70003','程序人員',st02)||'('||st01||')',st01 "
   strMid = strMid & " order by st01 "
   'end 2024/11/18
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strMid, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 Then
         InsertQueryLog (.RecordCount)
       Else
         InsertQueryLog (0)
       End If
       'If .RecordCount <> 0 Then
           Set grdDataList.Recordset = adoRecordset
       'End If
       SetDataListWidth
   End With
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

