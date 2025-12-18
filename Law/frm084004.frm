VERSION 5.00
Begin VB.Form frm084004 
   BorderStyle     =   1  '單線固定
   Caption         =   "逾期未結案統計表"
   ClientHeight    =   1890
   ClientLeft      =   1650
   ClientTop       =   2130
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3828
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1008
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1008
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2640
      Y1              =   1128
      Y2              =   1128
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   1056
      Width           =   900
   End
End
Attribute VB_Name = "frm084004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 3) As Integer, iLine As Integer

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   'Add By Cheng 2002/03/19
   If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
      Me.Text1(0).SetFocus
      Text1_GotFocus 0
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   
   If RunNick(Text1(0), Text1(1)) Then
      Text1(0).SetFocus
      Text1_GotFocus (0)
      Exit Sub
   End If
   DoEvents
   GetPrintLeft
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
   PrintCase
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub PrintCase()
 Dim i As Integer, StS(1 To 2) As String, Page As Integer, iPrint As Integer
 Dim TmpArea As String, rsTmp1 As New ADODB.Recordset, TmpArea1 As String
 Dim iCount(1 To 2) As Integer
 
On Error GoTo ErrHand:
   StS(1) = strGetcdnSQL
   
   If Me.Tag = 3 Then
      '910711 Sieg 107
'      strExc(0) = "select deptid,deptname,salesid,salesname,sum(count1) from ("
'      strExc(0) = strExc(0) & "select " & _
'         "decode(pa75,'',s1.ST15,s2.ST15) as deptid," & _
'         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
'         "decode(pa75,'',s1.st01,s2.st01) as salesid," & _
'         "decode(pa75,'',s1.st02,s2.st02) as salesname," & _
'         "count(*) as count1 from nextprogress,patent,customer,fagent,staff s1,staff s2," & _
'         "acc090 a1,acc090 a2,nation n1,nation n2 where " & _
'         StS(1) & " and np02='FCP' and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 and " & _
'         "substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and " & _
'         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
'         "substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and " & _
'         "fa10=n2.na01(+) and n2.na51=s2.st01(+) and s2.ST15=a2.a0901(+) " & _
'         "group by " & _
'         "decode(pa75,'',s1.ST15,s2.ST15)," & _
'         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
'         "decode(pa75,'',s1.st01,s2.st01)," & _
'         "decode(pa75,'',s1.st02,s2.st02) "
'      strExc(0) = strExc(0) & "Union select " & _
'         "decode(sp26,'',s1.ST15,s2.ST15) as deptid," & _
'         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
'         "decode(sp26,'',s1.st01,s2.st01) as salesid," & _
'         "decode(sp26,'',s1.st02,s2.st02) as salesname," & _
'         "count(*) as count1 from nextprogress,servicepractice,customer,fagent," & _
'         "staff s1,staff s2,acc090 a1,acc090 a2,nation n1,nation n2 " & _
'         "where" & StS(1) & " and np02='FG' and np02=sp01 and np03=sp02 and np04=sp03 and np05=sp04 and " & _
'         "substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and " & _
'         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
'         "substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and " & _
'         "fa10=n2.na01(+) and n2.na51=s2.st01(+) and s2.ST15=a2.a0901(+) " & _
'         "group by " & _
'         "decode(sp26,'',s1.ST15,s2.ST15)," & _
'         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
'         "decode(sp26,'',s1.st01,s2.st01)," & _
'         "decode(sp26,'',s1.st02,s2.st02) "
'      strExc(0) = strExc(0) & ") temp group by deptid,deptname,salesid,salesname order by deptid,salesid"
         'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
         '中國區基礎設定:簡欣儀(=Nation.Na51=A2012)，代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務")
         Dim midStr As String
        'Modified by Lydia 2016/02/03改成回傳case句
         'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
         midStr = Pub_GetSpecFCP
            
         'Modified by Lydia 2016/02/03
'      strExc(0) = "select deptid,deptname,salesid,salesname,sum(count1) from ("
'      strExc(0) = strExc(0) & "select " & _
'         "decode(pa75,'',s1.ST15,s2.ST15) as deptid," & _
'         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
'         "decode(pa75,'',s1.st01,s2.st01) as salesid," & _
'         "decode(pa75,'',s1.st02,s2.st02) as salesname," & _
'         "count(*) as count1 from nextprogress,patent,customer,fagent,staff s1,staff s2," & _
'         "acc090 a1,acc090 a2,nation n1,nation n2 where " & _
'         StS(1) & " and np02='FCP' and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 and " & _
'         "substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and " & _
'         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
'         "substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and " & _
'         "fa10=n2.na01(+) and decode(pa75,'Y51333010','" & midStr & "',n2.na51)=s2.st01 and s2.ST15=a2.a0901(+) " & _
'         "group by " & _
'         "decode(pa75,'',s1.ST15,s2.ST15)," & _
'         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
'         "decode(pa75,'',s1.st01,s2.st01)," & _
'         "decode(pa75,'',s1.st02,s2.st02) "
'      strExc(0) = strExc(0) & "Union select " & _
'         "decode(sp26,'',s1.ST15,s2.ST15) as deptid," & _
'         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
'         "decode(sp26,'',s1.st01,s2.st01) as salesid," & _
'         "decode(sp26,'',s1.st02,s2.st02) as salesname," & _
'         "count(*) as count1 from nextprogress,servicepractice,customer,fagent," & _
'         "staff s1,staff s2,acc090 a1,acc090 a2,nation n1,nation n2 " & _
'         "where" & StS(1) & " and np02='FG' and np02=sp01 and np03=sp02 and np04=sp03 and np05=sp04 and " & _
'         "substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and " & _
'         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
'         "substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and " & _
'         "fa10=n2.na01(+) and decode(sp26,'Y51333010','" & midStr & "',n2.na51)=s2.st01 and s2.ST15=a2.a0901(+) " & _
'         "group by " & _
'         "decode(sp26,'',s1.ST15,s2.ST15)," & _
'         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
'         "decode(sp26,'',s1.st01,s2.st01)," & _
'         "decode(sp26,'',s1.st02,s2.st02) "
      strExc(0) = "select deptid,deptname,salesid,salesname,sum(count1) from ("
      strExc(0) = strExc(0) & "select " & _
         "decode(pa75,'',s1.ST15,s2.ST15) as deptid," & _
         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
         "decode(pa75,'',s1.st01,s2.st01) as salesid," & _
         "decode(pa75,'',s1.st02,s2.st02) as salesname," & _
         "count(*) as count1 from nextprogress,patent,customer,fagent,staff s1,staff s2," & _
         "acc090 a1,acc090 a2,nation n1,nation n2 where " & _
         StS(1) & " and np02='FCP' and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 and " & _
         "substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and " & _
         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
         "substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and " & _
         "fa10=n2.na01(+) and decode(pa75," & midStr & ",n2.na51)=s2.st01 and s2.ST15=a2.a0901(+) " & _
         "group by " & _
         "decode(pa75,'',s1.ST15,s2.ST15)," & _
         "decode(pa75,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
         "decode(pa75,'',s1.st01,s2.st01)," & _
         "decode(pa75,'',s1.st02,s2.st02) "
      strExc(0) = strExc(0) & "Union select " & _
         "decode(sp26,'',s1.ST15,s2.ST15) as deptid," & _
         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903)) as deptname," & _
         "decode(sp26,'',s1.st01,s2.st01) as salesid," & _
         "decode(sp26,'',s1.st02,s2.st02) as salesname," & _
         "count(*) as count1 from nextprogress,servicepractice,customer,fagent," & _
         "staff s1,staff s2,acc090 a1,acc090 a2,nation n1,nation n2 " & _
         "where" & StS(1) & " and np02='FG' and np02=sp01 and np03=sp02 and np04=sp03 and np05=sp04 and " & _
         "substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and " & _
         "cu10=n1.na01(+) and n1.na51=s1.st01(+) and s1.ST15=a1.a0901(+) and " & _
         "substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and " & _
         "fa10=n2.na01(+) and decode(sp26," & midStr & ",n2.na51)=s2.st01 and s2.ST15=a2.a0901(+) " & _
         "group by " & _
         "decode(sp26,'',s1.ST15,s2.ST15)," & _
         "decode(sp26,'',nvl(a1.a0902,a1.a0903),nvl(a2.a0902,a2.a0903))," & _
         "decode(sp26,'',s1.st01,s2.st01)," & _
         "decode(sp26,'',s1.st02,s2.st02) "
      strExc(0) = strExc(0) & ") temp group by deptid,deptname,salesid,salesname order by deptid,salesid"
   
   'Modify by Sindy 2011/3/15 FCT,T,TF延展(102)和第二期(716)專用權須存在(TM17=Y)
   ElseIf Me.Tag = 5 Or Me.Tag = 6 Then
      strExc(0) = "SELECT ST15,nvl(A0902,a0903),NP10,nvl(st02,np10),count(*) " & _
         "FROM NEXTPROGRESS,STAFF,ACC090,Trademark WHERE" & StS(1) & " AND NP10=ST01(+) AND ST15=A0901(+) " & _
         "and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 " & _
         "and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' " & _
         "GROUP BY ST15,nvl(A0902,a0903),NP10,nvl(st02,np10) " & _
         "ORDER BY decode(ST15,null,'0',ST15),NP10 "
   '2011/3/15 End
   
   Else
      strExc(0) = "SELECT ST15,nvl(A0902,a0903),NP10,nvl(st02,np10),count(*) " & _
         "FROM NEXTPROGRESS,STAFF,ACC090 WHERE" & StS(1) & " AND NP10=ST01(+) AND " & _
         "ST15=A0901(+) GROUP BY ST15,nvl(A0902,a0903),NP10,nvl(st02,np10) ORDER BY decode(ST15,null,'0',ST15),NP10"
   End If
   
   Screen.MousePointer = vbHourglass
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      Screen.MousePointer = vbDefault
      InsertQueryLog (0) 'Add By Sindy 2010/10/19
      MsgBox "資料庫內無資料 !", vbInformation
      Exit Sub
   Else
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/10/19
   End If
   i = 1
   Page = 1
   Printer.KillDoc
   CaseTitle TmpArea, 1
   iPrint = 2700
   iCount(1) = 0
   iCount(2) = 0
   TmpArea1 = ""
   With RsTemp
   Do While Not .EOF
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      If IsNull(.Fields(1)) Then
         Printer.Print ""
         iCount(1) = Val(CheckStr(.Fields(4)))
         TmpArea = ""
      Else
      If TmpArea <> .Fields(1) Then
         Printer.Print CheckStr(.Fields(1))
         iCount(1) = Val(CheckStr(.Fields(4)))
         TmpArea = .Fields(1)
      Else
         Printer.Print ""
         iCount(1) = iCount(1) + Val(CheckStr(.Fields(4)))
      End If
      End If
      
      Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
      Printer.Print CheckStr(.Fields(3))
      Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(4), "&&&&")
      iCount(2) = iCount(2) + Val(CheckStr(.Fields(4)))
      
      .MoveNext
      If Not .EOF Then
         If iPrint >= 14000 Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle "", Page
            iPrint = 2700
         End If
         StS(2) = CheckStr(.Fields(1).Value)
         If StS(2) <> TmpArea Then
            iPrint = iPrint + 300
            Printer.CurrentX = 500:              Printer.CurrentY = iPrint
            Printer.Print String(195, "-")
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
            Printer.Print "小計"
            Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
            Printer.Print Format(iCount(1), "&&&&")
            iCount(1) = 0
            i = i + 3
            iPrint = iPrint + 300
         End If
         i = i + 1
         iPrint = iPrint + 300
         If iPrint >= 14000 Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle "", Page
            iPrint = 2700
         End If

      Else
         iPrint = iPrint + 300
         Printer.CurrentX = 500:              Printer.CurrentY = iPrint
         Printer.Print String(195, "-")
         iPrint = iPrint + 300
         Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
         Printer.Print "小計"
         Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
         Printer.Print Format(iCount(1), "&&&&")
         iPrint = iPrint + 300
         If iPrint >= 14000 Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle "", Page
            iPrint = 2700
         End If

      End If
   Loop
   End With
   iPrint = iPrint + 600
   Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
   Printer.Print "總計"
   Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
   Printer.Print Format(iCount(2), "&&&&")
   Screen.MousePointer = vbDefault
   MsgBox ("列印完成!!")
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 500:    PLeft(1) = 3500
   PLeft(2) = 6500:    PLeft(3) = 9000
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   'Printer.Orientation = vbPRORPortrait
   Printer.Font.Size = 22
   Printer.Font.Name = "細明體"
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "逾期未結案統計表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4200:         Printer.CurrentY = i + 500
   Printer.Print "本所期限:" & Format(ChangeTStringToTDateString(Text1(0)), "@@@@@@@@@") & _
      "-" & ChangeTStringToTDateString(Text1(1))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(195, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "件數"
   Printer.CurrentX = 500:              Printer.CurrentY = i + 2000
   Printer.Print String(195, "-")
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Printer.Orientation = 1
End Sub

Private Function strGetcdnSQL() As String
   If Text1(0).Text = "" And Text1(1).Text <> "" Then
      strGetcdnSQL = " NP08<='" + Text1(1) + "'"
   ElseIf Text1(0).Text <> "" And Text1(1).Text <> "" Then
      strGetcdnSQL = " (NP08 BETWEEN '" + ChangeTStringToWString(Text1(0)) + "' AND '" + ChangeTStringToWString(Text1(1)) + "')"
   End If
   If Text1(0).Text <> "" Or Text1(1).Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1(0) & "-" & Text1(1) 'Add By Sindy 2010/10/19
   End If
   '********************  89/03/17 nick 改
   'If Me.Tag = 0 Then
   '     strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('L','LA')"
   'Else
   '     strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('FCL','CFL')"
   'End If
   Select Case Me.Tag
   Case 0           '內法 LAWCASE HIRECASE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('L','LA')"
   Case 1           '外法 LAWCASE
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('FCL','CFL','LIN')"
   Case 2           'cfp  PATENT SERVICEPRACTICE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('CFP','CPS')"
   Case 3           'FCP  PATENT SERVICEPRACTICE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL"
   Case 4           'P    PATENT SERVICEPRACTICE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('P','PS')"
   Case 5           '內商 TRADEMARK,SERVICEPRACTICE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND SUBSTR(NP02,1,1)='T' "
   Case 6           '外商 TRADEMARK,SERVICEPRACTICE
      strGetcdnSQL = strGetcdnSQL & " AND NP06 IS NULL AND NP02 IN ('FCT','CFT','CFC','S')"
   End Select
   '****************************************************
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm084004 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 1
      If RunNick(Text1(0), Text1(1)) Then
         Text1(0).SetFocus
         Text1_GotFocus (0)
      End If
Case Else
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) = "" Then Exit Sub
   If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
   If Cancel Then TextInverse Text1(Index)
End Sub
