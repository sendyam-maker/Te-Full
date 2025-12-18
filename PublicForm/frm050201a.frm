VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050201a 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人新案案件統計"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3480
      Left            =   1080
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6138
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG1 
      Height          =   4968
      Left            =   36
      TabIndex        =   2
      Top             =   720
      Width           =   9252
      _ExtentX        =   16325
      _ExtentY        =   8758
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "代理人資料(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6696
      TabIndex        =   1
      Top             =   70
      Width           =   1300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8016
      TabIndex        =   0
      Top             =   70
      Width           =   1200
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   1272
      TabIndex        =   4
      Top             =   480
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "排行順序:"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "frm050201a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; MFG1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, strSQL11 As String, strSQL22 As String, strSQL33 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, strTemp As Variant, StrTest As String, strTemp1 As Variant, StrTest2 As String
Dim Int01 As Integer, Int02 As Integer, Int03 As Integer, Int04 As Integer, Int05 As Integer, StrSQL6 As String
Dim IntT01 As Integer, IntT02 As Integer, IntT03 As Integer, IntT04 As Integer, IntT05 As Integer
Dim StrR08001 As String
Dim StrR08002 As String
Dim StrR08003 As String
Dim StrR08004 As String
Dim StrR08005 As String
Dim StrR08006 As String
Dim StrR08007 As String
Dim StrR08008 As String
Dim StrR08009 As String
Dim StrR08010 As String
'Dim Rss As ADODB.Recordset
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub cmdok_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
  Select Case Index
  Case 1
    Unload Me
    frm050201.Show
    'Rss.Close
  Case 0
        Me.Enabled = False
        For i = 1 To MFG1.Rows - 1
            MFG1.row = i
            MFG1.col = 0
            If MFG1.Text = "V" Then
'               'Modify By Cheng 2002/03/14
''                MFG1.Col = 9
'                MFG1.Col = 10
'                bolToEndByNick = False
'                Screen.MousePointer = vbHourglass
'                frm100101_10.Show
'                frm100101_10.Hide
'
'                frm100101_10.Tag = MFG1.Text ' StrTag  傳代理人代號
'                frm100101_10.StrMenu
'                Screen.MousePointer = vbDefault
'                Me.Hide
'                frm100101_10.Show
'                Do
'                DoEvents
'                If bolToEndByNick = True Then Unload Me: Me.Enabled = True: Me.Show: Exit Sub
'                Loop Until Not frm100101_10.Visible
'                Unload frm100101_10
                MFG1.col = 0
                MFG1.Text = ""
                For j = 0 To MFG1.Cols - 1
                  MFG1.row = i
                  MFG1.col = j
                  MFG1.CellBackColor = QBColor(15)
                Next j

                Me.MFG1.col = 10
                If Not IsNull(Me.MFG1.Text) Then
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                   Screen.MousePointer = vbHourglass
                   frm100101_10.Show
                   frm100101_10.Tag = Me.MFG1.Text
                   frm100101_10.StrMenu
                   Screen.MousePointer = vbDefault
                   Me.Enabled = True
                   Exit Sub
                End If

            End If
        Next i
        Me.Enabled = True
        Me.Show
  Case Else
  End Select
End Sub

Private Sub Form_Load()
''edit by nickc 2007/02/06 不用 dll 了 Dim obj01 As Object

    MoveFormToCenter Me
    'MFG1.CellAlignment = 9
    MFG1.Rows = 2
    'MFG1.Cols = 8
    MFG1.FixedRows = 1
    MFG1.FixedCols = 0
    MFG1.ColWidth(0) = 200
    MFG1.ColWidth(1) = 1000
    MFG1.ColWidth(2) = 1500
    MFG1.ColWidth(3) = 800
    MFG1.ColWidth(4) = 800
    MFG1.ColWidth(5) = 1000
    MFG1.ColWidth(6) = 1000
    MFG1.ColWidth(7) = 800
    MFG1.ColWidth(8) = 800
    MFG1.ColWidth(9) = 2500
    MFG1.ColWidth(10) = 0
    'Add By Cheng 2002/03/14
    Me.MFG1.ColWidth(11) = 0
    Me.MFG1.ColWidth(12) = 0
    With MFG1
        .TextMatrix(0, 0) = "V"
        .TextMatrix(0, 1) = "代理人國籍"
        .TextMatrix(0, 2) = "代理人"
        .TextMatrix(0, 3) = "發明件數"
        .TextMatrix(0, 4) = "新型件數"
        .TextMatrix(0, 5) = "設計件數"
         
         'Add By Cheng 2002/01/04
        .TextMatrix(0, 6) = "專利總件數"
        
        .TextMatrix(0, 7) = "商標件數"
        .TextMatrix(0, 8) = "法務件數"
        .TextMatrix(0, 9) = "給案備註"
   End With
    Select Case frm050201.txt1(11)
    Case 1
        lbl1 = "專利件數"
    Case 2
        lbl1 = "商標件數"
    Case 3
        lbl1 = "法務件數"
    Case 4
        lbl1 = "人工排名"
    End Select
'strmenu
StrMenu1
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu1()
'Add By Cheng 2002/07/08
Dim StrSQLa As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050201 where id='" & strUserNum & "' "
'只有專利,商標,法務
'檢查收發文
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL6 = ""
If Len(frm050201.txt1(10)) <> 0 Then
   strSQL1 = strSQL1 + " and cp01 in (" & SQLGrpStr(frm050201.txt1(10), 1) & ") "
   strSQL2 = strSQL2 + " and cp01 in (" & SQLGrpStr(frm050201.txt1(10), 2) & ") "
   StrSQL3 = StrSQL3 + " and cp01 in (" & SQLGrpStr(frm050201.txt1(10), 3) & ") "
   pub_QL05 = pub_QL05 & ";" & frm050201.Label7 & frm050201.txt1(10) 'Add By Sindy 2010/01/22
End If
strSQL1 = strSQL1 + " AND cp31='Y' AND CP57 IS NULL and (CP44 IS NOT NULL OR PA75 IS NOT NULL) "
strSQL2 = strSQL2 + " AND cp31='Y' AND CP57 IS NULL and (CP44 IS NOT NULL OR tm44 IS NOT NULL) "
StrSQL3 = StrSQL3 + " AND cp31='Y' AND CP57 IS NULL and (CP44 IS NOT NULL OR lc22 IS NOT NULL) "
If frm050201.txt1(4) = "1" Then
   If Len(Trim(frm050201.txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp05>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp05>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
   End If
   If Len(Trim(frm050201.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp05<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp05<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "收文" & frm050201.Label3(0) & frm050201.txt1(5) & "-" & frm050201.txt1(6) 'Add By Sindy 2010/01/22
Else
   If Len(Trim(frm050201.txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp27>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp27>=" & Val(ChangeTStringToWString(frm050201.txt1(5))) & " "
   End If
   If Len(Trim(frm050201.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp27<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp27<=" & Val(ChangeTStringToWString(frm050201.txt1(6))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "發文" & frm050201.Label3(0) & frm050201.txt1(5) & "-" & frm050201.txt1(6) 'Add By Sindy 2010/01/22
End If
'Modify By Cheng 2002/03/08
'If frm050201.Op1(0).Value = True Then
   If Len(Trim(frm050201.txt1(0))) <> 0 Then
      strSQL1 = strSQL1 + " and PA09>='" & frm050201.txt1(0) & "' "
      strSQL2 = strSQL2 + " and TM10>='" & frm050201.txt1(0) & "' "
      StrSQL3 = StrSQL3 + " and LC15>='" & frm050201.txt1(0) & "' "
   End If
   If Len(Trim(frm050201.txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " and PA09<='" & frm050201.txt1(1) & "' "
      strSQL2 = strSQL2 & " and TM10<='" & frm050201.txt1(1) & "' "
      StrSQL3 = StrSQL3 & " and LC15<='" & frm050201.txt1(1) & "' "
   End If
   pub_QL05 = pub_QL05 & ";" & frm050201.Label6 & frm050201.txt1(0) & "-" & frm050201.txt1(1) 'Add By Sindy 2010/01/22
'Else
   If Len(Trim(frm050201.txt1(2))) <> 0 Then
      StrSQL6 = StrSQL6 + " and fa10>='" & frm050201.txt1(2) & "' "
   End If
   If Len(Trim(frm050201.txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 & " and fa10<='" & frm050201.txt1(3) & "z' "
   End If
   If Len(Trim(frm050201.txt1(2))) <> 0 Or Len(Trim(frm050201.txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm050201.Label10 & frm050201.txt1(2) & "-" & frm050201.txt1(3) 'Add By Sindy 2010/01/22
   End If
'End If
If Len(frm050201.txt1(7)) <> 0 Then
    strSQL1 = strSQL1 & " and decode(pa09,'000',pa75,cp44)='" & GetNewFagent(frm050201.txt1(7)) & "' "
    strSQL2 = strSQL2 & " and decode(tm10,'000',tm44,cp44)='" & GetNewFagent(frm050201.txt1(7)) & "' "
    StrSQL3 = StrSQL3 & " and decode(lc15,'000',lc22,cp44)='" & GetNewFagent(frm050201.txt1(7)) & "' "
    pub_QL05 = pub_QL05 & ";" & frm050201.Label4 & frm050201.txt1(7) 'Add By Sindy 2010/01/22
End If
If Len(frm050201.txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10>='" & frm050201.txt1(8) & "' "
    strSQL2 = strSQL2 + " and cp10>='" & frm050201.txt1(8) & "' "
    StrSQL3 = StrSQL3 + " and cp10>='" & frm050201.txt1(8) & "' "
End If
If Len(frm050201.txt1(9)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10<='" & frm050201.txt1(9) & "' "
    strSQL2 = strSQL2 + " and cp10<='" & frm050201.txt1(9) & "' "
    StrSQL3 = StrSQL3 + " and cp10<='" & frm050201.txt1(9) & "' "
End If
If Len(frm050201.txt1(8)) <> 0 Or Len(frm050201.txt1(9)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm050201.Label3(1) & frm050201.txt1(8) & "-" & frm050201.txt1(9) 'Add By Sindy 2010/01/22
End If
'Modify By Cheng 2002/03/08
'                    strSQL = "select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),count(*),0,0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='101' or (pa08='1' and cp10='104')) " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,count(*),0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='102' or (pa08='2' and cp10='104'))  " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,count(*),0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='103' or cp10='105') " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,0,count(*),0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(tm10,'000',tm44,cp44),1,8) as a,decode(substr(decode(tm10,'000',tm44,cp44),9,1),'','0',substr(decode(tm10,'000',tm44,cp44),9,1)) as b from caseprogress, trademark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp10='101' " & strSQL2 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '2'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,0,0,count(*),''   ,fa01||decode(fa02,'','0',fa02),0    ,'" & strUserNum & "'  from (select substr(decode(lc15,'000',lc22,cp44),1,8) as a,decode(substr(decode(lc15,'000',lc22,cp44),9,1),'','0',substr(decode(lc15,'000',lc22,cp44),9,1)) as b from caseprogress, lawcase   where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & StrSQL3 & ") new,fagent,nation where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and FA10=NA01(+) group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),'',fa01||decode(fa02,'','0',fa02),0,'" & strUserNum & "' "
'Modify By Cheng 2002/07/08
'若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'                    strSQL = "select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),count(*),0,0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='101' or (pa08='1' and cp10='104')) " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) " & StrSQL6 & " group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,count(*),0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='102' or (pa08='2' and cp10='104'))  " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) " & StrSQL6 & " group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,count(*),0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b from caseprogress, patent    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='103' or cp10='105') " & strSQL1 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) " & StrSQL6 & " group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,0,count(*),0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(tm10,'000',tm44,cp44),1,8) as a,decode(substr(decode(tm10,'000',tm44,cp44),9,1),'','0',substr(decode(tm10,'000',tm44,cp44),9,1)) as b from caseprogress, trademark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp10='101' " & strSQL2 & ") new,fagent,nation,facase where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '2'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) " & StrSQL6 & " group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSQL = strSQL + " union all select na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),0,0,0,0,count(*),''   ,fa01||decode(fa02,'','0',fa02),0    ,'" & strUserNum & "'  from (select substr(decode(lc15,'000',lc22,cp44),1,8) as a,decode(substr(decode(lc15,'000',lc22,cp44),9,1),'','0',substr(decode(lc15,'000',lc22,cp44),9,1)) as b from caseprogress, lawcase   where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & StrSQL3 & ") new,fagent,nation where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and FA10=NA01(+) " & StrSQL6 & " group by na03,nvl(FA05||' '||FA63||' '||FA64||' '||FA65,nvl(fa04,fa06)),'',fa01||decode(fa02,'','0',fa02),0,'" & strUserNum & "' "
'Modified by Lydia 2018/05/02 拿掉facase
'StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) ,"
'                    strSql = "select na03," & StrSQLa & "count(*),0,0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='101' or (pa08='1' and cp10='104')) " & strSQL1 & ") new,fagent,nation,facase,SystemKind  where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & "fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSql = strSql + " union all select na03," & StrSQLa & "0,count(*),0,0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='102' or (pa08='2' and cp10='104'))  " & strSQL1 & ") new,fagent,nation,facase,SystemKind where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & "fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSql = strSql + " union all select na03," & StrSQLa & "0,0,count(*),0,0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='103' or cp10='105' or cp10='125') " & strSQL1 & ") new,fagent,nation,facase,SystemKind                 where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '1'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & "fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSql = strSql + " union all select na03," & StrSQLa & "0,0,0,count(*),0,fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "'  from (select substr(decode(tm10,'000',tm44,cp44),1,8) as a,decode(substr(decode(tm10,'000',tm44,cp44),9,1),'','0',substr(decode(tm10,'000',tm44,cp44),9,1)) as b,CP01 from caseprogress, trademark  where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp10='101' " & strSQL2 & ") new,fagent,nation,facase,SystemKind                                 where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and '2'=fac01(+) and fa01=fac02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & "fac04,fa01||decode(fa02,'','0',fa02),fac03,'" & strUserNum & "' "
'strSql = strSql + " union all select na03," & StrSQLa & "0,0,0,0,count(*),''   ,fa01||decode(fa02,'','0',fa02),0    ,'" & strUserNum & "'  from (select substr(decode(lc15,'000',lc22,cp44),1,8) as a,decode(substr(decode(lc15,'000',lc22,cp44),9,1),'','0',substr(decode(lc15,'000',lc22,cp44),9,1)) as b,CP01 from caseprogress, lawcase    where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & StrSQL3 & ") new,fagent,nation,SystemKind                                                       where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) and FA10=NA01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & "'',fa01||decode(fa02,'','0',fa02),0,'" & strUserNum & "' "
StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) ,"
                    strSql = "select na03," & StrSQLa & "count(*),0,0,0,0,'' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='101' or (pa08='1' and cp10='104')) " & strSQL1 & ") new,fagent,nation,SystemKind  where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & " '' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "' "
strSql = strSql + " union all select na03," & StrSQLa & "0,count(*),0,0,0,'' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='102' or (pa08='2' and cp10='104'))  " & strSQL1 & ") new,fagent,nation,SystemKind where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & " '' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "' "
strSql = strSql + " union all select na03," & StrSQLa & "0,0,count(*),0,0,'' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "'  from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01 from caseprogress, patent     where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and (cp10='103' or cp10='105' or cp10='125') " & strSQL1 & ") new,fagent,nation,SystemKind                 where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & " '' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "' "
strSql = strSql + " union all select na03," & StrSQLa & "0,0,0,count(*),0,'' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "'  from (select substr(decode(tm10,'000',tm44,cp44),1,8) as a,decode(substr(decode(tm10,'000',tm44,cp44),9,1),'','0',substr(decode(tm10,'000',tm44,cp44),9,1)) as b,CP01 from caseprogress, trademark  where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp10='101' " & strSQL2 & ") new,fagent,nation,SystemKind                                 where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & " '' ,fa01||decode(fa02,'','0',fa02),0 , '" & strUserNum & "' "
strSql = strSql + " union all select na03," & StrSQLa & "0,0,0,0,count(*),''   ,fa01||decode(fa02,'','0',fa02),0    ,'" & strUserNum & "'  from (select substr(decode(lc15,'000',lc22,cp44),1,8) as a,decode(substr(decode(lc15,'000',lc22,cp44),9,1),'','0',substr(decode(lc15,'000',lc22,cp44),9,1)) as b,CP01 from caseprogress, lawcase    where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & StrSQL3 & ") new,fagent,nation,SystemKind                                                       where new.a=fa01(+) and new.b=fa02(+) and fa10=na01(+) AND NEW.CP01=SK01(+) " & StrSQL6 & " group by na03," & StrSQLa & " '',fa01||decode(fa02,'','0',fa02),0,'" & strUserNum & "' "
'end 2018/05/02

strSql = "INSERT INTO R050201 " & strSql
cnnConnection.Execute strSql
CheckOC
'Modify By Cheng 2002/01/04
'strSQL = "SELECT '' AS V,R08001 AS 代理人國籍,R08002 AS 代理人,SUM(R08003) AS 發明件數,SUM(R08004) AS 新型件數,SUM(R08005) AS 設計件數,SUM(R08006) AS 商標件數,SUM(R08007) AS 法務件數,R08008 AS 給案備註,R08009,SUM(R08003)+SUM(R08004)+SUM(R08005) AS TOTLE  FROM R050201 WHERE ID='" & strUserNum & "' "
'Modify By Cheng 2002/03/14
'多顯示"排行層級"欄為人工排名排序用
'strSQL = "SELECT '' AS V,R08001 AS 代理人國籍,R08002 AS 代理人,SUM(R08003) AS 發明件數,SUM(R08004) AS 新型件數,SUM(R08005) AS 設計件數, SUM(R08003) + SUM(R08004) + SUM(R08005) AS 專利總件數,SUM(R08006) AS 商標件數,SUM(R08007) AS 法務件數,R08008 AS 給案備註,R08009,SUM(R08003)+SUM(R08004)+SUM(R08005) AS TOTLE  FROM R050201 WHERE ID='" & strUserNum & "' "
strSql = "SELECT '' AS V,R08001 AS 代理人國籍,R08002 AS 代理人,SUM(R08003) AS 發明件數,SUM(R08004) AS 新型件數,SUM(R08005) AS 設計件數, SUM(R08003) + SUM(R08004) + SUM(R08005) AS 專利總件數,SUM(R08006) AS 商標件數,SUM(R08007) AS 法務件數,R08008 AS 給案備註,R08009,SUM(R08003)+SUM(R08004)+SUM(R08005) AS TOTLE ,nvl(R08010,0) as 排行層級 FROM R050201 WHERE ID='" & strUserNum & "' "
strSql = strSql & " GROUP BY R08001,R08002,R08008,R08009,R08010 "

'排名件數
Select Case frm050201.txt1(11)
Case "1" '專利件數
     strSql = strSql + " ORDER BY TOTLE desc"
Case "2" '商標件數
     strSql = strSql + " ORDER BY 商標件數 desc"
Case "3" '法務案件
     strSql = strSql + " ORDER BY 法務件數 desc"
'Remove by Lydia 2018/05/02 拿掉facase
'Case "4" '人工排名
'      'Modify By Cheng 2002/03/14
''     strSQL = strSQL = " ORDER BY R08010 "
'     strSql = strSql + " ORDER BY R08010 desc"
'end 2018/05/02
Case Else
End Select
pub_QL05 = pub_QL05 & ";" & frm050201.Label8 & frm050201.txt1(11) & frm050201.Label9 'Add By Sindy 2010/01/22
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   Me.Hide
   frm050201.Show
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Set MFG1.Recordset = adoRecordset
'Add By Cheng 2002/03/14
'欄位資料對齊方式
Me.MFG1.ColAlignment(3) = flexAlignRightCenter
Me.MFG1.ColAlignment(4) = flexAlignRightCenter
Me.MFG1.ColAlignment(5) = flexAlignRightCenter
Me.MFG1.ColAlignment(6) = flexAlignRightCenter
Me.MFG1.ColAlignment(7) = flexAlignRightCenter
Me.MFG1.ColAlignment(8) = flexAlignRightCenter

CheckOC
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm050201a = Nothing
End Sub

Private Sub MFG1_Click()
MFG1.Visible = False
MFG1.col = 0
MFG1.row = MFG1.MouseRow
If MFG1.MouseRow <> 0 Then
If MFG1.Text = "V" Then
     MFG1.Text = ""
     For i = 0 To MFG1.Cols - 1
          MFG1.col = i
          MFG1.CellBackColor = QBColor(15)
    Next i
Else
     MFG1.Text = "V"
     For i = 0 To MFG1.Cols - 1
         MFG1.col = i
         MFG1.CellBackColor = &HFFC0C0
     Next i

End If
End If
MFG1.Visible = True
End Sub

Public Sub PubShowNextData()
  Select Case cmdState
  Case 1
    Unload Me
    frm050201.Show
    'Rss.Close
  Case 0
        Me.Enabled = False
        For i = 1 To MFG1.Rows - 1
            MFG1.row = i
            MFG1.col = 0
            If MFG1.Text = "V" Then
                MFG1.col = 0
                MFG1.Text = ""
                For j = 0 To MFG1.Cols - 1
                  MFG1.row = i
                  MFG1.col = j
                  MFG1.CellBackColor = QBColor(15)
                Next j

                Me.MFG1.col = 10
                If Not IsNull(Me.MFG1.Text) Then
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                   Screen.MousePointer = vbHourglass
                   frm100101_10.Show
                   frm100101_10.Tag = Me.MFG1.Text
                   frm100101_10.StrMenu
                   Screen.MousePointer = vbDefault
                   Me.Enabled = True
                   Exit Sub
                End If

            End If
        Next i
        Me.Enabled = True
        Me.Show
  Case Else
  End Select
End Sub
