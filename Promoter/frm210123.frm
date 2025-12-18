VERSION 5.00
Begin VB.Form frm210123 
   BorderStyle     =   1  '單線固定
   Caption         =   "逾  預定收款日 15 日 未收款、未收齊清單列印"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3735
   Begin VB.TextBox tmpBox 
      Height          =   345
      Left            =   390
      TabIndex        =   2
      Text            =   "ALL"
      Top             =   300
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   435
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   180
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frm210123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim strTempS(1 To 3) As String
Dim iPgae As Integer, iLine As Integer
Dim m_i As Integer

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        Screen.MousePointer = vbHourglass
        doQuery
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210123 = Nothing
End Sub

Sub doQuery()
   Dim stCon As String
   Dim strCon As String
   
   ''Modified by Morgan 2011/11/23 考慮拆收據情形改抓acc0j0,另+rd06 is null條件
   'strCon = "SELECT st06,a0k01,a0k02,a0k22,a0k20,ST02,A0K03,SUBSTR(CU04,1,10) as cu04,DIFF ,Acase ,pay,min(cp01||'-'||cp02||'-'||cp03||'-'||cp04) as ocp FROM ("
   strCon = "SELECT st06,a0k01,a0k02,a0k22,a0k20,ST02,A0K03,SUBSTR(CU04,1,10) as cu04,DIFF ,Acase ,pay,min(a0j01) as a0j01 FROM ("
   'end 2011/11/23
   
   strCon = strCon & "  select a0k01,a0k02,A0K03,a0k20,a0k22,SUM(nvl(a0k06,0)+nvl(a0k07,0)-PAY) DIFF,SUM(nvl(a0k06,0)+nvl(a0k07,0)) Acase,sum(pay) pay FROM( "
   strCon = strCon & " SELECT A0K01,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0) as a0k06,nvl(A0K07,0) as a0k07,sum(nvl(a1u04,0))+sum(nvl(a1u07,0))-sum(nvl(a1u08,0))+sum(nvl(a1u05,0))+sum(nvl(a1u09,0))-sum(nvl(a1u10,0)) PAY"
   strCon = strCon & " from acc0k0,ACC1U0 WHERE (a0k09 is null or a0k09 = 0) AND A0K01=A1U02(+) "
   strCon = strCon & " and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0)) "
   strCon = strCon & " GROUP BY A0K01,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0),nvl(A0K07,0) ) AA"
   strCon = strCon & " where (nvl(a0k06,0)+nvl(a0k07,0)) > PAY GROUP BY a0k01,a0k02,A0K03,a0k20,a0k22"
   
   ''Modified by Morgan 2011/11/23 考慮拆收據情形改抓acc0j0,另+rd06 is null條件
   'strCon = strCon & " ) NEW,CUSTOMER,STAFF,caseprogress "
   'strCon = strCon & " WHERE a0k01=cp60(+) and SUBSTR(A0K03,1,8)=CU01(+) AND SUBSTR(A0K03,9,1)=CU02(+) AND a0k20=ST01(+) and diff<>0 "
   strCon = strCon & " ) NEW,CUSTOMER,STAFF,acc0j0 "
   strCon = strCon & " WHERE a0k01=a0j13(+) and SUBSTR(A0K03,1,8)=CU01(+) AND SUBSTR(A0K03,9,1)=CU02(+) AND a0k20=ST01(+) and diff<>0 "
   'end 2011/11/23
   
   strCon = strCon & " group by st06,A0K01,a0k22,a0k20,ST02,a0k02,A0K03,SUBSTR(CU04,1,10),DIFF,Acase,pay "
   
   ''Modified by Morgan 2011/11/23 考慮拆收據情形改抓acc0j0,另+rd06 is null條件
   'stCon = "select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,aa.st02 as 智權人員,aa.cu04 as 申請人,aa.ocp as 本所案號,sqldatet(AA.a0k02) as 收據日期,AA.a0k01 as 收據號碼,TO_CHAR(AA.Acase,'999,999,999') 收據金額,TO_CHAR(AA.diff,'999,999,999') 未收金額,sqldatet(rd05) 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'1' as osort from (" & strCon & ") AA,acc090,(select * from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd01,rd02) in (select rd01,max(rd02) from ReceivablesDay group by rd01) group by rd01,rd02)) BB,caseprogress where AA.a0k22=a0901(+) and AA.a0k01=cp60(+) and cp09=BB.RD01(+) and BB.rd05<=" & (CompWorkDay(16, strSrvDate(1), 1)) & "  "
   stCon = "select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,aa.st02 as 智權人員,aa.cu04 as 申請人,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sqldatet(AA.a0k02) as 收據日期,AA.a0k01 as 收據號碼,TO_CHAR(AA.Acase,'999,999,999') 收據金額,TO_CHAR(AA.diff,'999,999,999') 未收金額,sqldatet(rd05) 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'1' as osort from (" & strCon & ") AA,acc090,(select * from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd01,rd02) in (select rd01,max(rd02) from ReceivablesDay where rd06 is null group by rd01) group by rd01,rd02)) BB,caseprogress where AA.a0k22=a0901(+) and AA.a0j01=cp09(+) and cp09=BB.RD01(+) and BB.rd05<=" & (CompWorkDay(16, strSrvDate(1), 1)) & "  "
   'end 2011/11/23
   
   stCon = stCon & " order by oSort "
    If m_rs.State = 1 Then m_rs.Close
    m_rs.CursorLocation = adUseClient
    m_rs.Open stCon, cnnConnection, adOpenStatic, adLockReadOnly
    If Not m_rs.EOF And Not m_rs.BOF Then
        m_rs.MoveFirst
        strTempS(1) = ""
        strTempS(2) = "'"
        strTempS(3) = ""
        PrintTitle
        Do While Not m_rs.EOF
            For m_i = 1 To 10
                strTemp(m_i) = CheckStr(m_rs.Fields(m_i - 1))
            Next m_i
            If strTemp(1) <> strTempS(1) Then
                strTempS(1) = strTemp(1)
                strTempS(2) = strTemp(2)
                strTempS(3) = strTemp(3)
            Else
                strTemp(1) = ""
                If strTemp(2) <> strTempS(2) Then
                    strTempS(2) = strTemp(2)
                    strTempS(3) = strTemp(3)
                Else
                    strTemp(2) = ""
                    If strTemp(3) <> strTempS(3) Then
                        strTempS(3) = strTemp(3)
                    Else
                        strTemp(3) = ""
                    End If
                End If
            End If
            strTemp(2) = StrToStr(strTemp(2), 4)
            strTemp(3) = StrToStr(strTemp(3), 3)
            strTemp(4) = StrToStr(strTemp(4), 10)
            PrintDetail
            If iLine >= 53 Then
                If m_rs.AbsolutePosition <> m_rs.RecordCount Then
                    Printer.NewPage
                    PrintTitle
                End If
            End If
            m_rs.MoveNext
        Loop
    Else
        ShowNoData
        Exit Sub
    End If
    Printer.EndDoc
    ShowPrintOk
End Sub

Sub PrintTitle()
Dim oStr As String
oStr = "逾預定收款日，未收款、未全額收齊清單"
GetPleft
Printer.Font.Size = 18
Printer.Font.Underline = True
Printer.FontBold = True
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(oStr) / 2)
Printer.CurrentY = 300
Printer.Print oStr
Printer.Font.Size = 10
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "所別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "部門"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "智權人員"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "申請人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "收據日期"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "收據號碼"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "收據金額"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "未收金額"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "預定收款日"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(300, "-")
iLine = iLine + 1
End Sub
Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 10
    If m_j = 8 Or m_j = 9 Then
        Printer.CurrentX = PLeft(m_j + 1) - 300 - Printer.TextWidth(strTemp(m_j))
        Printer.CurrentY = iLine * 300
        Printer.Print strTemp(m_j)
    Else
        Printer.CurrentX = PLeft(m_j)
        Printer.CurrentY = iLine * 300
        Printer.Print strTemp(m_j)
    End If
Next m_j
iLine = iLine + 1
End Sub
Sub GetPleft()
PLeft(1) = 200
PLeft(2) = 750
PLeft(3) = 1750
PLeft(4) = 2500
PLeft(5) = 4650
PLeft(6) = 6150
PLeft(7) = 7100
PLeft(8) = 8100
PLeft(9) = 9100
PLeft(10) = 10100
End Sub
