VERSION 5.00
Begin VB.Form frm090639 
   BorderStyle     =   1  '單線固定
   Caption         =   "支援記錄獎金統計"
   ClientHeight    =   1920
   ClientLeft      =   2415
   ClientTop       =   1725
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4395
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2376
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   280
      Index           =   1
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   2
      Top             =   975
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   280
      Index           =   0
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   1
      Top             =   975
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   2430
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "支援日期："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   1095
   End
End
Attribute VB_Name = "frm090639"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/14 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2016/01/25
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         MsgBox "支援日期範圍起不可空白!!", , "USER 輸入錯誤"
         txt1(0).SetFocus
         Exit Sub
     ElseIf Len(txt1(1)) = 0 Then
         MsgBox "支援日期範圍止不可空白!!", , "USER 輸入錯誤"
         txt1(1).SetFocus
         Exit Sub
     ElseIf txt1(0) > txt1(1) Then
         MsgBox "支援日期範圍起不可大於支援日期範圍止!!", , "USER 輸入錯誤"
         txt1(0).SetFocus
         Exit Sub
     Else
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         ClearQueryLog (Me.Name)
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
Dim strCon As String
Dim rsRd As New ADODB.Recordset
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim strAt As String
Dim strR As Integer
Dim tmpArr As Variant
Dim xlsSalesPoint As New Excel.Application
Dim wks639 As New Worksheet
Dim xRows As Integer '目前列位置

   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1)
   strCon = " and sh01>=" & DBDATE(txt1(0)) & " and sh01<=" & DBDATE(txt1(1))
   
   'Modified by Lydia 2016/07/06 計算支援獎金的次數也要考慮一案只算一次
   'strExc(0) = "select st03,sh02,st02,count(distinct nvl(sh06||sh07||sh08||sh09,sh01||sh02||sh03||sh04)) cnt1" & _
               ",count(sh20) cnt2 from supporthour,staff where substr(st03,1,2)='P1' and sh03<>'71011' " & strCon & _
               " and sh02=st01(+) group by st03,sh02,st02 order by 1,2 "
   'Modified by Lydia 2016/08/11 改變計算方式=>報表中的次數應為”計算支援”之統計,計算支援獎金統計則是自”計算支援”再過濾同一本所案號只計1次by郭雅娟
    'strExc(1) = "select st03,sh02,st02,nvl(sh06||sh07||sh08||sh09,sh01||sh02||sh03||sh04) tc1,decode(sh20,null,0,1) tc2 " & _
               " from supporthour,staff where substr(st03,1,2)='P1' and sh03<>'71011' " & strCon & _
               " and sh02=st01(+)"
    'strExc(0) = "select st03,sh02,st02,tc1,decode(sum(tc2)-1,-1,0,1) tc3 " & _
                "from (" & strExc(1) & ") group by st03,sh02,st02,tc1"
    'strExc(0) = "select st03,sh02,st02,count(tc1) cnt1,sum(tc3) cnt2 " & _
                "From (" & strExc(0) & ") group by st03,sh02,st02 order by 1,2 "
    'Modified by Lydia 2017/01/06 + st01
    strExc(1) = "select st03,sh02,st02,nvl(sh06||sh07||sh08||sh09,sh01||sh02||sh03||sh04) tc1, 1 qty,st01 " & _
               " from supporthour,staff where substr(st03,1,2)='P1' and sh03<>'71011' " & strCon & _
               " and sh02=st01(+) and sh20='V'"
    strExc(0) = "select st03,sh02,st02,sum(qty) cnt1,count(distinct tc1) cnt2,st01 " & _
                "From (" & strExc(1) & ") group by st03,sh02,st02,st01 order by 1,2 "
   intI = 0
   Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (rsRd.RecordCount) 'Added by Lydia 2023/04/20
      strAt = "支援人員,次數,計算支援獎金統計"
      xRows = 1
      strTempFile = txt1(0) & "-" & txt1(1) & Me.Caption & ACDate(ServerDate) & MsgText(43)
      strPath = strExcelPath & strTempFile
      
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
         MkDir strExcelPath
      End If
      If Dir(strPath) <> "" Then
         Kill strPath
      End If
      
      xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks639 = xlsSalesPoint.Worksheets(1)
      wks639.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks639.PageSetup.PrintTitleRows = "$1:$3"
       'Modified by Lydia 2017/01/06 欄寬從10->13
       wks639.Columns("a:a").ColumnWidth = 13
       wks639.Columns("b:b").ColumnWidth = 10
       wks639.Columns("c:c").ColumnWidth = 20
       wks639.Range("a1").Value = txt1(0) & "-" & txt1(1) & Me.Caption
       wks639.Range("a1:c1").Merge
       wks639.Range("a2").Value = "列印人:" & strUserName
       wks639.Range("a2:c2").Merge
       xRows = 3
       tmpArr = Split(strAt, ",")
       For intI = 0 To UBound(tmpArr)
           If tmpArr(intI) <> "" Then
              wks639.Range(Chr(Asc("a") + intI) & xRows).Value = tmpArr(intI)
           End If
       Next
       strR = xRows + 1
       rsRd.MoveFirst
       With rsRd
           Do While Not .EOF
             xRows = xRows + 1
             'Modified by Lydia 2017/01/06 +姓名前面加員工編號
             'wks639.Range("a" & xRows).Value = "" & .Fields("st02")
             wks639.Range("a" & xRows).Value = .Fields("st01") & " " & .Fields("st02")
             wks639.Range("b" & xRows).Value = "" & .Fields("cnt1")
             wks639.Range("c" & xRows).Value = "" & .Fields("cnt2")
             .MoveNext
           Loop
       End With
       
       xRows = xRows + 1
       wks639.Range("a" & xRows).Value = "合　計:"
       wks639.Range("b" & xRows).Formula = "=SUM(b" & strR & ":b" & xRows - 1 & ")"
       wks639.Range("c" & xRows).Formula = "=SUM(c" & strR & ":c" & xRows - 1 & ")"
       With wks639.Range("a1:c" & xRows)
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
       End With
       
        '判斷若版本2007以上改變存格式
        If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
        Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        'Modify by Amy 2021/06/22 原:strPath 改中文字顯示
        MsgBox "檔案已產生！" & vbCrLf & "檔案存於 " & strExcelPathN & " " & strTempFile, vbInformation
   Else
      InsertQueryLog (0)
      Exit Sub
   End If

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090639 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Dim strTemp1 As String
    If Index < 2 And txt1(Index) <> "" Then
        strTemp1 = txt1(Index)
        If CheckIsTaiwanDate(strTemp1) = False Then
           MsgBox "請輸入民國日期!", vbCritical
           txt1(Index).SetFocus
           txt1_GotFocus Index
           Exit Sub
        End If
    End If
    
End Sub
