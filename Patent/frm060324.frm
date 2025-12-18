VERSION 5.00
Begin VB.Form frm060324 
   BorderStyle     =   1  '單線固定
   Caption         =   "FCP/FG 作業失誤清單"
   ClientHeight    =   1380
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   4572
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4572
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   3030
      MaxLength       =   7
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   1
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3105
      TabIndex        =   3
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2070
      TabIndex        =   2
      Top             =   30
      Width           =   912
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2790
      X2              =   2910
      Y1              =   653
      Y2              =   653
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   3030
      TabIndex        =   7
      Top             =   570
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   2070
      TabIndex        =   6
      Top             =   570
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2790
      X2              =   2910
      Y1              =   965
      Y2              =   965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外專上次列印發文日："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   570
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本次列印發文日："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   882
      Width           =   1440
   End
End
Attribute VB_Name = "frm060324"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Const cntX As Long = 500
Const cntY As Long = 500
Const cntL As Long = 300
Dim iPrint As Long
Dim PLeft(0 To 3) As Integer  'Added by Lydia 2016/12/12 明細欄位左邊界

'Added by Lydia 2016/12/12 設定欄位寬度
Private Sub GetPleft()
'strTmp = Right(Space(3) & .Fields(0), 15) & Space(2) & Right(Space(10) & Format(.Fields(1), DDollar), 10)
   Printer.Font = 22
   PLeft(0) = cntX '本所案號
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(8, "　")) + 100 '規費
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + 100 '案件性質
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(5, "　")) + 100 '進度備註
End Sub

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
   Set frm060324 = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   Dim ii As Integer
   
   For ii = 1 To 2
      If Text1(ii).Text = "" Then
         MsgBox "發文日條件不可空白！"
         Cancel = True
         Text1(ii).SetFocus
         Text1_GotFocus ii
         Exit For
      End If
      Cancel = False
      Text1_Validate ii, Cancel
      If Cancel = True Then Exit For
   Next
   TxtValidate = Not Cancel
End Function

Private Sub cmdPrint_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      If DoPrint = True Then
         'Modify by Amy 2014/07/14
'         SaveSetting "TAIE", "FCP", Me.Name & "#DATE01", Text1(1).Text
'         SaveSetting "TAIE", "FCP", Me.Name & "#DATE02", Text1(2).Text
         'modify by sonia 2016/11/15 因財務處也要看故加入控制外專人員操作才更新
         'PUB_SaveLastDate Me.Name, "Text1(1)", Text1(1)
         'PUB_SaveLastDate Me.Name, "Text1(2)", Text1(2)
         If PUB_GetST05(strUserNum) <> "00" Then
            PUB_SaveLastDate Me.Name, "Text1(1)", Text1(1)
            PUB_SaveLastDate Me.Name, "Text1(2)", Text1(2)
         End If
         'end 2016/11/15
         'end 2014/07/14
         MsgBox "列印完成！"
      End If
   End If
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub PrintHead(stTitle As String, iPage As Integer, iPageTot As Integer)

   iPrint = cntY
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5500 - (Printer.TextWidth(stTitle) / 2)
   Printer.CurrentY = iPrint
   Printer.Print stTitle
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   iPrint = Printer.CurrentY
   Printer.CurrentX = cntX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   
   iPrint = iPrint + cntL
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iPrint
   Printer.Print "列印時間：" & Format(ServerTime, Tformat)
   
   iPrint = iPrint + cntL
   Printer.CurrentX = cntX
   Printer.CurrentY = iPrint
   Printer.Print "發文日：" & Text1(1) & " － " & Text1(2)
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & Format(iPage) & "/" & Format(iPageTot)
   
   iPrint = iPrint + cntL
   Printer.CurrentX = cntX
   Printer.Print String(200, "-")
   
   iPrint = iPrint + cntL
   'Modified by Lydia 2016/12/12
   'Printer.CurrentX = cntX
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   'Modified by Lydia 2016/12/12
   'Printer.CurrentX = cntX + 2000
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "規費"
   
   'Added by 2016/12/12 增加案件性質,備註
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "進度備註"
   'end 2016/12/12
   
   iPrint = iPrint + cntL
   Printer.CurrentX = cntX
   Printer.Print String(200, "-")
      
End Sub

Private Sub PrintTail(stData As String)

      Printer.FontSize = 12
      Printer.CurrentX = cntX
      Printer.Print String(200, "-")
      Printer.CurrentX = cntX
      Printer.Print stData
      
End Sub

Private Function DoPrint() As Boolean
   Dim strTitle As String, strTmp As String
   Dim iPage As Integer, iPageTot As Integer
   Dim iRec As Integer, iRecs As Integer, lngTot As Long
   Const PageRec As Integer = 40
   Dim inP As Integer 'Added by Lydia 2016/12/12
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/12/13
   'Modified by Lydia 2016/12/12 增加案件性質,備註
   'strSql = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01, CP17 X02 from caseprogress " & _
      " where cp01 in ('FCP','FG') AND CP27>=19110000+" & Text1(1).Text & " AND CP27<=19110000+" & Text1(2) & " AND CP18<0 AND CP57 IS NULL" & _
      " ORDER BY 1,2"
   strSql = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01, CP17 X02,DECODE(PA09,'000',CPM03,CPM04) X03,CP64 X04" & _
      " from caseprogress,patent,casepropertymap" & _
      " where cp01 in ('FCP','FG') AND CP158>=19110000+" & Text1(1).Text & " AND CP158<=19110000+" & Text1(2) & " AND CP159=0 AND CP18<0" & _
      " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) ORDER BY 1,2"

On Error GoTo ErrHnd
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
         Printer.Orientation = 1
         Printer.Font = "細明體"
         GetPleft 'Added by Lydia 2016/12/12 設定欄位寬度
         .MoveFirst
         iPage = 1: lngTot = 0: iRec = 0: iRecs = 0
         iPageTot = .RecordCount \ PageRec + IIf((.RecordCount Mod PageRec) = 0, 0, 1)
         strTitle = "FCP/FG 作業失誤清單"
         PrintHead strTitle, iPage, iPageTot
         Do While Not .EOF
            
            iRec = iRec + 1: iRecs = iRecs + 1: lngTot = lngTot + .Fields(1)
            If iRec > PageRec Then
               Printer.NewPage
               iPage = iPage + 1
               PrintHead strTitle, iPage, iPageTot
               iRec = 0
            End If
            'Modified by Lydia 2016/12/12 改欄位寬度設定
            'strTmp = Right(Space(3) & .Fields(0), 15) & Space(2) & Right(Space(10) & Format(.Fields(1), DDollar), 10)
            'iPrint = iPrint + cntL
            'Printer.CurrentX = cntX
            'Printer.CurrentY = iPrint
            'Printer.Print strTmp
            iPrint = iPrint + cntL
            For inP = 0 To 3
                If inP <> 1 Then
                   Printer.CurrentX = PLeft(inP)
                   strTmp = "" & .Fields(inP)
                   If inP = 2 Then
                      strTmp = PUB_StrToStr(strTmp, 10)
                   ElseIf inP = 3 Then
                      strTmp = PUB_StrToStr(strTmp, 56)
                   End If
                Else '規費靠右
                   Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(Format("" & .Fields(inP), DDollar)) - 100
                   strTmp = Format("" & .Fields(inP), DDollar)
                End If
                Printer.CurrentY = iPrint
                Printer.Print strTmp
            Next inP
            'end 2016/12/12
            
            .MoveNext
         Loop
         strTmp = Left("共 " & iRecs & " 筆" & Space(12), 13) & Space(2) & Right(Space(10) & Format(lngTot, DDollar), 10)
         PrintTail strTmp
         
         Printer.EndDoc
         DoPrint = True
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/13
         MsgBox "無可列印資料！", vbInformation
         
         'Addded by Morgan 2024/1/12 沒資料
         If PUB_GetST05(strUserNum) <> "00" Then
            PUB_SaveLastDate Me.Name, "Text1(1)", Text1(1)
            PUB_SaveLastDate Me.Name, "Text1(2)", Text1(2)
         End If
         'end 2024/1/12
      End If
   End With

ErrHnd:

   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   CheckOC

End Function

Private Sub Form_Load()
 
   MoveFormToCenter Me
   'Modify by Amy 2014/07/14 第一次DB可能沒資料所以先抓client
   If PUB_GetLastDate(Me.Name, "Text1(1)") <> "" Then
        Label2(0).Caption = PUB_GetLastDate(Me.Name, "Text1(1)")
   Else
        Label2(0).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE01", "")
   End If
   If PUB_GetLastDate(Me.Name, "Text1(2)") <> "" Then
        Label2(1).Caption = PUB_GetLastDate(Me.Name, "Text1(2)")
   Else
        Label2(1).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE02", "")
   End If
   'end 2014/07/14
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) <> "" Then
      If ChkDate(Text1(Index)) = False Then
         Cancel = True
         Text1_GotFocus Index
      End If
   End If
End Sub
