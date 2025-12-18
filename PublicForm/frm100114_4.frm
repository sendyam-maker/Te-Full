VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100114_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件明細"
   ClientHeight    =   5388
   ClientLeft      =   2160
   ClientTop       =   2940
   ClientWidth     =   11136
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5388
   ScaleWidth      =   11136
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   315
      Index           =   1
      Left            =   10035
      TabIndex        =   2
      Top             =   75
      Width           =   1080
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   315
      Index           =   0
      Left            =   8895
      TabIndex        =   1
      Top             =   75
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4905
      Left            =   15
      TabIndex        =   0
      Top             =   435
      Width           =   11070
      _ExtentX        =   19516
      _ExtentY        =   8657
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm100114_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'91.08.01   nick 第二階段 編號 805 新撰寫
Option Explicit
Dim strSql As String
Dim tmpFa As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
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
     Me.Hide
Case 1
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
'92.04.16 nick
cmdState = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100114_4 = Nothing
End Sub

Sub StrMenu(oStrFA As String, From100114 As Boolean, oStrCp10 As String, oStrCp01 As String, Optional oStrCFFC As String = "")
tmpFa = oStrFA
'edit by nickc 2007/12/21
'If From100114 = True Then
If Mid(UCase(tmpFa), 1, 1) = "Y" Then
'代理人來
    StrMenu1 tmpFa, oStrCp10, oStrCp01, oStrCFFC
Else
'申請人來
    StrMenu2 tmpFa, oStrCp10, oStrCp01
End If

End Sub

'從代理人來
'edit by nickc 2007/12/21
'Sub StrMenu1(oStrFA As String, StrCp10 As String, oStrCp01 As String)
Sub StrMenu1(oStrFA As String, strCP10 As String, oStrCp01 As String, Optional oStrCFFC As String = "")
Dim ii As Integer

Me.Enabled = False
'顯示表單上頭資料
'開始搜尋
Dim strSQL11 As String
Dim strSQL22 As String
Dim strSQL33 As String
Dim strSQL44 As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim StrSQL6 As String
Dim strSQL8 As String

'add by nickc 2007/12/21
Dim frm As Form
Dim IsFrom100114 As Boolean
IsFrom100114 = False
For Each frm In Forms
    Select Case UCase(frm.Name)
    Case "FRM100114_1"
            IsFrom100114 = True
    End Select
Next

strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
strSQL8 = ""
'add by nickc 2007/12/21
If IsFrom100114 = True Then
                If Len(Trim(oStrCp01)) <> 0 Then
                   'strsql1 = strsql1 & " and tm01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
                   'strsql2 = strsql2 & " and pa01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
                   'strsql3 = strsql3 & " and sp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
                   'strsql4 = strsql4 & " and lc01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
                   'strsql11 = strsql11 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
                   'strsql22 = strsql22 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
                   'strsql33 = strsql33 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
                   'strsql44 = strsql44 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
                   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 2) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") "
                   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 1) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") "
                   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 5) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") "
                   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 3) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") "
                   strSQL8 = strSQL8 & " and HC01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 4) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") "
                '   strSQL11 = strSQL11 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 2) & ") "
                '   StrSQL22 = StrSQL22 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 1) & ") "
                '   StrSQL33 = StrSQL33 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 5) & ") "
                '   strsql44 = strsql44 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 3) & ") "
                End If
                If Len(Trim(frm100114_1.txt1(8))) <> 0 Then           '檢查申請國家
                   strSQL1 = strSQL1 + " AND TM10='" & frm100114_1.txt1(8) & "' "
                   strSQL2 = strSQL2 + " AND PA09='" & frm100114_1.txt1(8) & "' "
                   StrSQL3 = StrSQL3 + " AND SP09='" & frm100114_1.txt1(8) & "' "
                   StrSQL4 = StrSQL4 + " AND LC15='" & frm100114_1.txt1(8) & "' "
                '   strSQL11 = strSQL11 + " AND TM10='" & frm100114_1.txt1(8) & "' "
                '   StrSQL22 = StrSQL22 + " AND PA09='" & frm100114_1.txt1(8) & "' "
                '   StrSQL33 = StrSQL33 + " AND SP09='" & frm100114_1.txt1(8) & "' "
                '   strsql44 = strsql44 + " AND LC15='" & frm100114_1.txt1(8) & "' "
                End If
                If Len(Trim(frm100114_1.txt1(6))) <> 0 Then            '檢查案件性質
                    strSQL1 = strSQL1 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    strSQL2 = strSQL2 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    StrSQL3 = StrSQL3 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    StrSQL4 = StrSQL4 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    strSQL8 = strSQL8 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    strSQL11 = strSQL11 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    StrSQL22 = StrSQL22 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    StrSQL33 = StrSQL33 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    strsql44 = strsql44 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                End If
                If Len(Trim(frm100114_1.txt1(7))) <> 0 Then
                    strSQL1 = strSQL1 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    strSQL2 = strSQL2 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    StrSQL3 = StrSQL3 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    StrSQL4 = StrSQL4 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    strSQL8 = strSQL8 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    strSQL11 = strSQL11 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    StrSQL22 = StrSQL22 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    StrSQL33 = StrSQL33 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    strsql44 = strsql44 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                End If
                If frm100114_1.txt1(3) = "1" Then        '收文
                   If Len(frm100114_1.txt1(4)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
                   End If
                   If Len(frm100114_1.txt1(5)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
                   Else
                      If Len(frm100114_1.txt1(4).Text) > 0 Then
                         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                      End If
                   End If
                Else
                   If Len(frm100114_1.txt1(4)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
                   End If
                   If Len(frm100114_1.txt1(5)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
                   Else
                      If Len(frm100114_1.txt1(4).Text) > 0 Then
                         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                      End If
                   End If
                End If
                strSQL5 = strSQL5 & " and cp10='" & strCP10 & "' "
Else
        '系統類別
        If Len(Trim(oStrCp01)) <> 0 Then
           strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") "
           strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") "
           StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") "
           StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") "
           strSQL8 = strSQL8 & " and hc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") "
        End If
        If Len(Trim(frm100102_1.Text6)) <> 0 Then            '檢查案件性質
            strSQL1 = strSQL1 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL2 = strSQL2 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL3 = StrSQL3 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL4 = StrSQL4 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL8 = strSQL8 + " AND CP10>='" & frm100102_1.Text6 & "' "
        End If
        If Len(Trim(frm100102_1.Text7)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL2 = strSQL2 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL3 = StrSQL3 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL4 = StrSQL4 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL8 = strSQL8 + " AND CP10<='" & frm100102_1.Text7 & "' "
        End If
        If Len(frm100102_1.Text4) <> 0 Then
           strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100102_1.Text4)) & " "
        End If
        If Len(frm100102_1.Text5) <> 0 Then
           strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100102_1.Text5)) & " "
        Else
           If Len(frm100102_1.Text4) > 0 Then
              strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
           End If
        End If
        strSQL5 = strSQL5 & " and cp10='" & strCP10 & "' "

End If

                'Modify By Cheng 2003/08/15
'    strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (SP58='" & oStrFA & "')  and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL5
'    strSQL = strSQL & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL5
'    strSQL = strSQL + " ORDER BY 本所案號 "
'edit by nickc 2005/05/13
'    strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP26='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'
'    strSQL = strSQL & " Union SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA26='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA27='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA28='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA29='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA30='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP08='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'  and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號, HC06 AS 案件名稱,'' AS 商品類別, CPM03 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM Hirecase ,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'  and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL8 & strSQL5
'
'    strSQL = strSQL + " ORDER BY 本所案號 "
    '2010/9/15 MODIFY BY SONIA 所有日期欄若需排序改百年日期排序問題
'Mark by Lydia 2019/12/26 利益衝突案件：因為已經在前一畫面frm100114_3，改成先丟暫存檔逐案件比對，所以直接抓前一畫面的暫存檔
'    If InStr(1, oStrCFFC, "FC") > 0 Then
'        strSql = "SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'        strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'        strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP26='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'        strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    ElseIf InStr(1, oStrCFFC, "CF") > 0 Then
'        strSql = "select decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'        strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'        strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'        strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    End If
'end 2019/12/26
'edit by nickc 2007/12/21 秀玲說，這邊只要抓代理人就好
'    strSQL = strSQL & " Union SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA26='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA27='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA28='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA29='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA30='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP08='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'  and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號, HC06 AS 案件名稱,'' AS 商品類別, CPM03 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM Hirecase ,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'  and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL8 & strSQL5

'Added by Lydia 2019/12/26 利益衝突案件：因為已經在前一畫面frm100114_3，改成先丟暫存檔逐案件比對，所以直接抓前一畫面的暫存檔
    strExc(1) = ""
    If InStr(1, oStrCFFC, "FC") > 0 Then '判斷FC/CF代理
        strExc(1) = " AND R021003='FC代理' "
    ElseIf InStr(1, oStrCFFC, "CF") > 0 Then
        strExc(1) = " AND R021003='CF代理' "
    End If
    strSql = " SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, TRADEMARK, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04"
    strSql = strSql & " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, PATENT, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04"
    strSql = strSql & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, ServicePractice, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04"
    strSql = strSql & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, LawCase, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04"
'end 2019/12/26
    strSql = strSql + " ORDER BY FSort,本所案號 "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
     cmdok(0).Enabled = True
     cmdok(1).Enabled = True
Else
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 4) = Me.grdDataList.TextMatrix(ii, 4) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 8), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True

End Sub

'由申請人畫面來   910801
Sub StrMenu2(oStrFA As String, strCP10 As String, oStrCp01 As String)
Dim ii As Integer

Me.Enabled = False
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim StrSQL6 As String
Dim strSQL8 As String


'add by nickc 2007/12/21
Dim frm As Form
Dim IsFrom100114 As Boolean
IsFrom100114 = False
For Each frm In Forms
    Select Case UCase(frm.Name)
    Case "FRM100114_1"
            IsFrom100114 = True
    End Select
Next

'開始搜尋
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
strSQL8 = ""
If IsFrom100114 = True Then
                If Len(Trim(oStrCp01)) <> 0 Then
                   'strsql1 = strsql1 & " and tm01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
                   'strsql2 = strsql2 & " and pa01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
                   'strsql3 = strsql3 & " and sp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
                   'strsql4 = strsql4 & " and lc01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
                   'strsql11 = strsql11 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
                   'strsql22 = strsql22 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
                   'strsql33 = strsql33 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
                   'strsql44 = strsql44 & " and cp01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
                   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 2) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") "
                   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 1) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") "
                   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 5) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") "
                   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 3) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") "
                   strSQL8 = strSQL8 & " and HC01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 4) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") "
                '   strSQL11 = strSQL11 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 2) & ") "
                '   StrSQL22 = StrSQL22 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 1) & ") "
                '   StrSQL33 = StrSQL33 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 5) & ") "
                '   strsql44 = strsql44 & " and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, oStrCp01), 3) & ") "
                End If
                If Len(Trim(frm100114_1.txt1(8))) <> 0 Then           '檢查申請國家
                   strSQL1 = strSQL1 + " AND TM10='" & frm100114_1.txt1(8) & "' "
                   strSQL2 = strSQL2 + " AND PA09='" & frm100114_1.txt1(8) & "' "
                   StrSQL3 = StrSQL3 + " AND SP09='" & frm100114_1.txt1(8) & "' "
                   StrSQL4 = StrSQL4 + " AND LC15='" & frm100114_1.txt1(8) & "' "
                '   strSQL11 = strSQL11 + " AND TM10='" & frm100114_1.txt1(8) & "' "
                '   StrSQL22 = StrSQL22 + " AND PA09='" & frm100114_1.txt1(8) & "' "
                '   StrSQL33 = StrSQL33 + " AND SP09='" & frm100114_1.txt1(8) & "' "
                '   strsql44 = strsql44 + " AND LC15='" & frm100114_1.txt1(8) & "' "
                End If
                If Len(Trim(frm100114_1.txt1(6))) <> 0 Then            '檢查案件性質
                    strSQL1 = strSQL1 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    strSQL2 = strSQL2 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    StrSQL3 = StrSQL3 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    StrSQL4 = StrSQL4 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                    strSQL8 = strSQL8 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    strSQL11 = strSQL11 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    StrSQL22 = StrSQL22 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    StrSQL33 = StrSQL33 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                '    strsql44 = strsql44 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                End If
                If Len(Trim(frm100114_1.txt1(7))) <> 0 Then
                    strSQL1 = strSQL1 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    strSQL2 = strSQL2 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    StrSQL3 = StrSQL3 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    StrSQL4 = StrSQL4 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                    strSQL8 = strSQL8 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    strSQL11 = strSQL11 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    StrSQL22 = StrSQL22 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    StrSQL33 = StrSQL33 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                '    strsql44 = strsql44 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                End If
                If frm100114_1.txt1(3) = "1" Then        '收文
                   If Len(frm100114_1.txt1(4)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
                   End If
                   If Len(frm100114_1.txt1(5)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
                   Else
                      If Len(frm100114_1.txt1(4).Text) > 0 Then
                         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                      End If
                   End If
                Else
                   If Len(frm100114_1.txt1(4)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
                   End If
                   If Len(frm100114_1.txt1(5)) <> 0 Then
                      strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
                   Else
                      If Len(frm100114_1.txt1(4).Text) > 0 Then
                         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                      End If
                   End If
                End If
                strSQL5 = strSQL5 & " and cp10='" & strCP10 & "' "
Else
        '系統類別
        If Len(Trim(oStrCp01)) <> 0 Then
           strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 2) & ") "
           strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 1) & ") "
           StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 5) & ") "
           StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 3) & ") "
           strSQL8 = strSQL8 & " and hc01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") and cp01 in (" & SQLGrpStr(IIf(oStrCp01 <> "ALL", oStrCp01, (oStrCp01)), 4) & ") "
        End If
        '2010/3/15 ADD BY SONIA
        If Len(Trim(frm100102_1.txtCountry(0))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10>='" & frm100102_1.txtCountry(0) & "' "
           strSQL2 = strSQL2 + " AND PA09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL3 = StrSQL3 + " AND SP09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL4 = StrSQL4 + " AND LC15>='" & frm100102_1.txtCountry(0) & "' "
        End If
        If Len(Trim(frm100102_1.txtCountry(1))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10<='" & frm100102_1.txtCountry(1) & "' "
           strSQL2 = strSQL2 + " AND PA09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL3 = StrSQL3 + " AND SP09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL4 = StrSQL4 + " AND LC15<='" & frm100102_1.txtCountry(1) & "' "
        End If
        '2010/3/15 END
        If Len(Trim(frm100102_1.Text6)) <> 0 Then            '檢查案件性質
            strSQL1 = strSQL1 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL2 = strSQL2 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL3 = StrSQL3 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL4 = StrSQL4 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL8 = strSQL8 + " AND CP10>='" & frm100102_1.Text6 & "' "
        End If
        If Len(Trim(frm100102_1.Text7)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL2 = strSQL2 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL3 = StrSQL3 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL4 = StrSQL4 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL8 = strSQL8 + " AND CP10<='" & frm100102_1.Text7 & "' "
        End If
        If Len(frm100102_1.Text4) <> 0 Then
           strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100102_1.Text4)) & " "
        End If
        If Len(frm100102_1.Text5) <> 0 Then
           strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100102_1.Text5)) & " "
        Else
           If Len(frm100102_1.Text4) > 0 Then
              strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
           End If
        End If
        strSQL5 = strSQL5 & " and cp10='" & strCP10 & "' "
End If

'Modify By Cheng 2003/08/15
'    strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (SP58='" & oStrFA & "')  and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL5
'    strSQL = strSQL & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL5
'    strSQL = strSQL + " ORDER BY 本所案號 "
'edit by nickc 2005/05/13
'    strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA26='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA27='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA28='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA29='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA30='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP08='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP58='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP59='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "' and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,'' AS 商品類別,CPM03 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM Hirecase, nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "' and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL8 & strSQL5
'
'    strSQL = strSQL & " Union SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP26='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'
'    strSQL = strSQL + " ORDER BY 本所案號 "
'Mark by Lydia 2019/12/26 利益衝突案件：因為已經在前一畫面frm100114_3，改成先丟暫存檔逐案件比對，所以直接抓前一畫面的暫存檔
'    strSql = "SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA26='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA27='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA28='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA29='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA30='" & oStrFA & "' and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP08='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP58='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP59='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP65='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP66='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "' and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSql = strSql + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,'' AS 商品類別,CPM03 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM Hirecase, nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "' and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL8 & strSQL5
'end 2019/12/26
'edit by nickc 2007/12/21 秀玲說，這邊只要抓申請人就好
'    strSQL = strSQL & " Union SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and PA75='" & oStrFA & "'and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and SP26='" & oStrFA & "' and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5
'    strSQL = strSQL & " union select decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM PATENT,nation,caseprogress,casepropertymap WHERE pa09=na01(+) and cp44='" & oStrFA & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE sp09=na01(+) and (cp44='" & oStrFA & "')  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & strSQL5

'Added by Lydia 2019/12/26 利益衝突案件：因為已經在前一畫面frm100114_3，改成先丟暫存檔逐案件比對，所以直接抓前一畫面的暫存檔
    strExc(1) = ""  'Added by Lydia 2024/01/17 從申請人查詢來,不用加額外判斷
    strSql = " SELECT decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM09 AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, TRADEMARK, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04"
    strSql = strSql & " union select decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, PATENT, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04"
    strSql = strSql & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, ServicePractice, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04"
    strSql = strSql & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, LawCase, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04"
    strSql = strSql & " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,'' AS 商品類別,R021007 AS 案件性質,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,DECODE(CP60,NULL,'','Y') AS 請款註記, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
                " FROM R100102_1, HireCase, CASEPROGRESS WHERE ID= '" & strUserNum & "@frm100114_3' AND R021006=CP09 AND R021014='" & oStrCp01 & "' AND R021018=" & CNULL(strCP10) & strExc(1) & _
                " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04"
'end 2019/12/26
    strSql = strSql + " ORDER BY FSort,本所案號 "
    
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cmdok(0).Enabled = True
    cmdok(1).Enabled = True
Else
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 4) = Me.grdDataList.TextMatrix(ii, 4) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 8), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True

End Sub

Private Sub SetDataListWidth()
With grdDataList
'edit by nickc
'.Cols = 8
.Cols = 10
.row = 0
.col = 0: .Text = "本所案號"
.ColWidth(0) = 1550
.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 1: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(1) = 0
Else
    grdDataList.ColWidth(1) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
.col = 2: .Text = "案件名稱"
.ColWidth(2) = 2500
.CellAlignment = flexAlignCenterCenter
.col = 3: .Text = "商品類別"
.ColWidth(3) = 800
.CellAlignment = flexAlignCenterCenter
.col = 4: .Text = "案件性質"
.ColWidth(4) = 1400
.CellAlignment = flexAlignCenterCenter
.col = 5: .Text = "收文日"
.ColWidth(5) = 850
.CellAlignment = flexAlignCenterCenter
.col = 6: .Text = "發文日"
.ColWidth(6) = 850
.CellAlignment = flexAlignCenterCenter
.col = 7: .Text = "請款註記"
.ColWidth(7) = 800
.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
.col = 8: .Text = "CP09"
.ColWidth(8) = 0
.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/13
.col = 9: .Text = ""
.ColWidth(9) = 0
.CellAlignment = flexAlignCenterCenter
End With
End Sub

