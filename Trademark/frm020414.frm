VERSION 5.00
Begin VB.Form frm020414 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件准駁統計表"
   ClientHeight    =   2400
   ClientLeft      =   4275
   ClientTop       =   2130
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   48
      TabIndex        =   14
      Top             =   1752
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   15
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3084
      TabIndex        =   8
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2292
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   2
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1116
      MaxLength       =   4
      TabIndex        =   1
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1104
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1116
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1104
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1116
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1428
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1116
      TabIndex        =   0
      Top             =   468
      Width           =   1908
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2088
      TabIndex        =   13
      Top             =   1476
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   1620
      X2              =   2370
      Y1              =   888
      Y2              =   888
   End
   Begin VB.Line Line1 
      X1              =   1548
      X2              =   2808
      Y1              =   1248
      Y2              =   1248
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   12
      Top             =   816
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   1140
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   10
      Top             =   1476
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   492
      Width           =   912
   End
End
Attribute VB_Name = "frm020414"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 15) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, PLeft1(1 To 7) As Integer, SeekPrint As Integer, SeekPrintL As Integer, k As Integer

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定

     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
     Printer.EndDoc 'Add By Sindy 2011/11/1
     'Modified by Moran 2015/6/1
     'Printer.PaperSize = 39
     Printer.PaperSize = PUB_GetPaperSize(15, 2)
     'end 2015/6/1
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(2)) = 0 Then
             s = MsgBox("申請國家區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
               Me.txt1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
             
             If Len(txt1(4)) = 0 Then
                 s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
                 txt1(3).SetFocus
                 txt1_GotFocus (3)
                 Exit Sub
             Else
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
                 ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
                 Process
                 Me.Enabled = True
                 Screen.MousePointer = vbDefault
             End If
         End If
     End If
Case 1
    '若印表機有變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020414 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
StrSQL6 = ""
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP25>=" & Val(ChangeTStringToWString(txt1(3)))
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP25<=" & Val(ChangeTStringToWString(txt1(4)))
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM44='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND SP26='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(5) & lbl1 'Add By Sindy 2010/10/19
End If
CheckOC
'Modify By Cheng 2003/03/26
'資料只含有代理人的資料
'strSQL = "SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10 from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='101' " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='101' " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='401' " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='401' " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='601' OR CP10='602') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='601' OR CP10='602') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='603' OR CP10='604') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='603' OR CP10='604') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='605' OR CP10='606') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='605' OR CP10='606') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606') " & strSQL2 + StrSQL6
'92.04.03 nick add left join
'strSQL = "SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10 from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND CP10='101' " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND CP10='101' " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND CP10='401' " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND CP10='401' " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='601' OR CP10='602') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='601' OR CP10='602') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='603' OR CP10='604') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='603' OR CP10='604') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='605' OR CP10='606') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10='605' OR CP10='606') " & strSQL2 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606') " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01 AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02 AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606') " & strSQL2 + StrSQL6
strSql = "SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10 from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='101' " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='101' " & strSQL2 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='401' " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND CP10='401' " & strSQL2 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='601' OR CP10='602') " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='601' OR CP10='602') " & strSQL2 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='603' OR CP10='604') " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='603' OR CP10='604') " & strSQL2 + StrSQL6
'modify by sonia 2017/8/30 以下四句都加623,624
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='605' OR CP10='606' OR CP10='623' OR CP10='624') " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10='605' OR CP10='606' OR CP10='623' OR CP10='624') " & strSQL2 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,trademark,nation,fagent where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606' AND CP10<>'623' AND CP10<>'624') " & strSQL1 + StrSQL6
strSql = strSql + " UNION all SELECT NVL(na03,na04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),0,0,0,0,0,0,0,0,0,0,decode(cp24,'1',1,0),decode(cp24,'2',1,0),decode(cp24,'1',1,0),decode(cp24,'2',1,0),CP09,CP10  from caseprogress,SERVICEPRACTICE,nation,fagent where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(cp44,1,8) = FA01(+) AND decode(SUBSTR(cp44,9,1),'','0',SUBSTR(cp44,9,1))=FA02(+) AND (CP10<>'101' AND CP10<>'401' AND CP10<>'601' AND CP10<>'602' AND CP10<>'603' AND CP10<>'604' AND CP10<>'605' AND CP10<>'606' AND CP10<>'623' AND CP10<>'624') " & strSQL2 + StrSQL6
'end 2017/8/30
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 15
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "INSERT INTO R020414 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT R082001,R082002,SUM(R082003),SUM(R082004),SUM(R082005),SUM(R082006),SUM(R082007),SUM(R082008),SUM(R082009),SUM(R082010),SUM(R082011),SUM(R082012),SUM(R082013),SUM(R082014),SUM(R082015),SUM(R082016) FRoM R020414 WHERE ID='" & strUserNum & "' GROUP BY R082001,R082002 ORDER BY R082001,R082002 "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        strTemp3 = CheckStr(.Fields(0))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 15
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(0) <> strTemp3 Then
                Page = Page + 1
                ShowLine
                Printer.NewPage
                strTemp3 = strTemp(0)
                PrintTitle
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            'Add By Cheng 2003/03/26
            '代理人名稱取前10碼
            strTemp(1) = StrToStr(strTemp(1), 10)
            PrintDatil
            If iPrint >= 14000 Then
                Page = Page + 1
                ShowLine
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
ShowLine
Printer.EndDoc

End Sub

Sub PrintTitle()
GetPleft
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "代理人案件准駁統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
iPrint = iPrint + 300
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & txt1(0)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & strTemp3
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(3) - 184 - (Printer.TextWidth("商申") / 2)
Printer.CurrentY = iPrint
Printer.Print "商申"
Printer.CurrentX = PLeft(5) - 184 - (Printer.TextWidth("復審") / 2)
Printer.CurrentY = iPrint
Printer.Print "復審"
Printer.CurrentX = PLeft(7) - 184 - (Printer.TextWidth("異議異答") / 2)
Printer.CurrentY = iPrint
Printer.Print "異議異答"
Printer.CurrentX = PLeft(9) - 184 - (Printer.TextWidth("裁定裁答") / 2)
Printer.CurrentY = iPrint
Printer.Print "裁定裁答"
Printer.CurrentX = PLeft(11) - 184 - (Printer.TextWidth("廢止廢答") / 2)
Printer.CurrentY = iPrint
Printer.Print "廢止廢答"
Printer.CurrentX = PLeft(13) - 184 - (Printer.TextWidth("其他") / 2)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(15) - 184 - (Printer.TextWidth("小計") / 2)
Printer.CurrentY = iPrint
Printer.Print "小計"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Line (PLeft(2), iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "代理人"
Printer.CurrentX = PLeft(2) + 400
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(3) + 400
Printer.CurrentY = iPrint
Printer.Print "駁"
For k = 4 To 12 Step 2
    Printer.CurrentX = PLeft(k) + 400
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1) + 400
    Printer.CurrentY = iPrint
    Printer.Print "敗"
Next k
Printer.CurrentX = PLeft(14) + 400
Printer.CurrentY = iPrint
Printer.Print "准/勝"
Printer.CurrentX = PLeft(15) + 400
Printer.CurrentY = iPrint
Printer.Print "駁/敗"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 15
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 0
PLeft(2) = 2500
PLeft(3) = 3644
PLeft(4) = 4788
PLeft(5) = 5932
PLeft(6) = 7076
PLeft(7) = 8220
PLeft(8) = 9364
PLeft(9) = 10508
PLeft(10) = 11652
PLeft(11) = 12796
PLeft(12) = 13940
PLeft(13) = 15084
PLeft(14) = 16228
PLeft(15) = 17372
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickT

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020414 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add By Cheng 2002/07/09
Dim strTempName As String
Dim strTmp

Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNickT), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 2
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 3, 4
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 4 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case 5
    'Modify By Cheng 2003/03/26
    '若有輸入代理人
    If Me.txt1(5).Text <> "" Then
          'Modify By Cheng 2002/07/09
    '      lbl1.Caption = GetPrjName2(txt1(5))
          strTmp = Split(Me.txt1(0).Text, ",")
          If PUB_GetAgentName(IIf(Me.txt1(0).Text = "", "", strTmp(0)), Me.txt1(5).Text, strTempName) Then
             lbl1.Caption = "" & strTempName
          Else
             s = MsgBox("無此代理人", , "USER 輸入錯誤")
             lbl1.Caption = ""
             txt1(5).SetFocus
             txt1(5).SelStart = 0
             txt1(5).SelLength = Len(txt1(5))
          End If
    '若未輸入代理人
    Else
        lbl1.Caption = ""
    End If
Case Else
End Select

End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

