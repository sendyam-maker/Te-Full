VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090630 
   BorderStyle     =   1  '單線固定
   Caption         =   "加乘註記修改歷史查詢列印"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8025
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   8
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1065
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   7
      Left            =   2670
      MaxLength       =   2
      TabIndex        =   3
      Top             =   450
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   6
      Left            =   2370
      MaxLength       =   1
      TabIndex        =   2
      Top             =   450
      Width           =   225
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   1
      Top             =   450
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3900
      Left            =   30
      TabIndex        =   16
      Top             =   1380
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   6879
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
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
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   4365
      MaxLength       =   6
      TabIndex        =   7
      Top             =   750
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   6
      Top             =   750
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1035
      MaxLength       =   7
      TabIndex        =   5
      Top             =   750
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   4365
      MaxLength       =   1
      TabIndex        =   4
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1035
      MaxLength       =   3
      TabIndex        =   0
      Top             =   450
      Width           =   390
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   330
      Index           =   2
      Left            =   7005
      TabIndex        =   11
      Top             =   30
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印"
      Height          =   330
      Index           =   1
      Left            =   6000
      TabIndex        =   10
      Top             =   30
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   4965
      TabIndex        =   9
      Top             =   30
      Width           =   960
   End
   Begin MSForms.Label lbl3 
      Height          =   255
      Left            =   2010
      TabIndex        =   20
      Top             =   1080
      Width           =   2055
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl2 
      Height          =   255
      Left            =   5340
      TabIndex        =   19
      Top             =   780
      Width           =   2055
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "承辦/繪圖："
      Height          =   180
      Left            =   75
      TabIndex        =   18
      Top             =   1110
      Width           =   945
   End
   Begin VB.Line Line2 
      X1              =   1230
      X2              =   2880
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(1:承辦；2:草圖；3:墨圖)"
      Height          =   180
      Left            =   4740
      TabIndex        =   17
      Top             =   495
      Width           =   1965
   End
   Begin VB.Line Line1 
      X1              =   1695
      X2              =   2520
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "修改人員："
      Height          =   180
      Left            =   3240
      TabIndex        =   15
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "類別："
      Height          =   180
      Left            =   3600
      TabIndex        =   14
      Top             =   495
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   435
      TabIndex        =   13
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   75
      TabIndex        =   12
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm090630"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grd1改字型=新細明體-ExtB、lbl2、lbl3 ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Create by nickc 2005/03/16

Option Explicit
'列印控制
'Modified by Lydia 2017/06/02 調整列印版面
'Dim PLeft(0 To 7) As Integer
Dim PLeft(0 To 8) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim bolInsert As Boolean, bolUpdate As Boolean, bolDelete As Boolean, bolSelect As Boolean, bolPrint As Boolean
Private Const colGap As Integer = 180  'Added by Lydia 2017/06/02 行間距
Private Const nDot As Integer = 132  'Added by Lydia 2017/06/02 劃線的數量
Sub SetGrid()
With grd1
      .Cols = 9
      .row = 0
      .col = 0: .Text = "本所案號"
      .ColWidth(0) = 1500
      .CellAlignment = flexAlignCenterCenter
      'add by nickc 2005/07/25 加入承辦人
      'Modified by Lydia 2017/06/02
      '.col = 1: .Text = "承辦人"
      .col = 1: .Text = "承辦/繪圖"
      .ColWidth(1) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .Text = "修改人員"
      .ColWidth(2) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .Text = "修改日期"
      .ColWidth(3) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .Text = "修改時間"
      .ColWidth(4) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .Text = "類別"
      .ColWidth(5) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .Text = "原加乘註記"
      .ColWidth(6) = 1400
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .Text = "修改後加乘註記"
      .ColWidth(7) = 1400
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .Text = "理由"
      .ColWidth(8) = 3000
      .CellAlignment = flexAlignCenterCenter

End With
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strSQL1 As String
Dim arrTxt As TextBox
Dim bolDataIsOk As Boolean
Select Case Index
Case 0    '查詢
         bolDataIsOk = False
         For Each arrTxt In txt1
            txt1_Validate arrTxt.Index, bolDataIsOk
            If bolDataIsOk = True Then Exit Sub
         Next
         Screen.MousePointer = vbHourglass
         grd1.MousePointer = flexHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
         strSQL1 = ""
         If Trim(txt1(0)) <> "" Then
            strSQL1 = strSQL1 & " and cp01='" & txt1(0) & "' "
         End If
         If Trim(txt1(5)) <> "" Then
            strSQL1 = strSQL1 & " and cp02='" & txt1(5) & "' "
         End If
         If Trim(txt1(6)) <> "" Then
            strSQL1 = strSQL1 & " and cp03='" & txt1(6) & "' "
         End If
         If Trim(txt1(7)) <> "" Then
            strSQL1 = strSQL1 & " and cp04='" & txt1(7) & "' "
         End If
         If Trim(txt1(0)) <> "" Or Trim(txt1(5)) <> "" Or _
            Trim(txt1(6)) <> "" Or Trim(txt1(7)) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/20
         End If
         If Trim(txt1(1)) <> "" Then
            strSQL1 = strSQL1 & " and fs04='" & txt1(1) & "' "
            pub_QL05 = pub_QL05 & ";" & Label3 & txt1(1) & "(1:承辦；2:草圖；3:墨圖)" 'Add By Sindy 2010/12/20
         End If
         If Trim(txt1(2)) <> "" Then
            strSQL1 = strSQL1 & " and fs02>=" & ChangeTStringToWString(txt1(2)) & " "
         End If
         If Trim(txt1(3)) <> "" Then
            strSQL1 = strSQL1 & " and fs02<=" & ChangeTStringToWString(txt1(3)) & " "
         End If
         If Trim(txt1(2)) <> "" Or Trim(txt1(3)) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/20
         End If
         If Trim(txt1(4)) <> "" Then
            strSQL1 = strSQL1 & " and fs08='" & txt1(4) & "' "
            pub_QL05 = pub_QL05 & ";" & Label4 & txt1(4) & lbl2 'Add By Sindy 2010/12/20
         End If
         'add by nickc 2005/07/25 加入承辦人
         If Trim(txt1(8)) <> "" Then
            'Added by Lydia 2017/06/02 +判斷承辦人/繪圖人
            'strSQL1 = strSQL1 & " and cp14='" & txt1(8) & "' "
            'pub_QL05 = pub_QL05 & ";" & Label7 & txt1(8) & lbl3 'Add By Sindy 2010/12/20
            Select Case Trim(txt1(1))
                Case "1"
                    strSQL1 = strSQL1 & " and cp14='" & txt1(8) & "' "
                    pub_QL05 = pub_QL05 & ";" & "承辦人:" & txt1(8) & lbl3
                Case "2", "3"
                    strSQL1 = strSQL1 & " and cp29='" & txt1(8) & "' "
                    pub_QL05 = pub_QL05 & ";" & "繪圖人:" & txt1(8) & lbl3
                Case Else
                    strSQL1 = strSQL1 & " and (cp14='" & txt1(8) & "' or cp29='" & txt1(8) & "') "
                    pub_QL05 = pub_QL05 & ";" & "承辦人/繪圖人:" & txt1(8) & lbl3
            End Select
            'end 2017/06/02
         End If
         'edit by nickc 2005/07/25 加入承辦人
         'strSQL = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04,st02," & SQLDate("fs02") & ",substr(rtrim(ltrim(to_char(fs03,'000000'))),1,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),3,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),5,2),decode(fs04,'1','承辦人','2','草圖','3','墨圖',''),fs05,fs06,fs07 from flagstory,staff,caseprogress where fs08=st01(+) and fs01=cp09(+) " & strSQL1
         'Modified by Lydia 2017/06/02
         'strSql = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04,s2.st02,s1.st02," & SQLDate("fs02") & ",substr(rtrim(ltrim(to_char(fs03,'000000'))),1,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),3,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),5,2),decode(fs04,'1','承辦人','2','草圖','3','墨圖',''),fs05,fs06,fs07 from flagstory,staff s1,caseprogress,staff s2 where fs08=s1.st01(+) and fs01=cp09(+) and cp14=s2.st01(+) " & strSQL1
         'Modified by Lydia 2017/06/06 備註去掉換行符號 fs07=> replace(fs07,chr(13)||chr(10),'');
         strSql = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04,decode(fs04,'1',s2.st02,nvl(s3.st02,s2.st02)) tname,s1.st02," & _
                  SQLDate("fs02") & ",substr(rtrim(ltrim(to_char(fs03,'000000'))),1,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),3,2)||':'||substr(rtrim(ltrim(to_char(fs03,'000000'))),5,2)," & _
                  "decode(fs04,'1','承辦人','2','草圖','3','墨圖',''),fs05,fs06,replace(fs07,chr(13)||chr(10),'') fs07 from flagstory,staff s1,caseprogress,staff s2,staff s3 " & _
                  "where fs08=s1.st01(+) and fs01=cp09(+) and cp14=s2.st01(+) and cp29=s3.st01(+) " & strSQL1
         'end 2017/06/02
         CheckOC3
         With AdoRecordSet3
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount <> 0 Then
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
                  'add by nick 2007/12/12
                  If bolPrint Then
                     cmdok(1).Enabled = True
                  End If
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/20
                  ShowNoData
                  cmdok(1).Enabled = False
               End If
               Set grd1.Recordset = AdoRecordSet3
               SetGrid
         End With
         CheckOC3
         grd1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
Case 1    '列印
         PrintData
Case 2
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'add by nickc 2007/12/12
      bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
      bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
      bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
      bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
      bolPrint = IsUserHasRightOfFunction(Me.Name, strPrint, False)
      
      
   cmdok(1).Enabled = False
   SetGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090630 = Nothing
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.grd1.MouseRow < 1 Then
         If m_blnColOrderAsc = True Then
             Me.grd1.Sort = 5 '昇冪
             m_blnColOrderAsc = False
         Else
             Me.grd1.Sort = 6 '降冪
             m_blnColOrderAsc = True
         End If
    End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0, 6, 4, 8
      KeyAscii = UpperCase(KeyAscii)
Case 1
      If Not (KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 49 Or KeyAscii = 50 Or KeyAscii = 51) Then
         KeyAscii = 0
      End If
Case 2, 3
      If Not (KeyAscii = 13 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
         KeyAscii = 0
      End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
If Index = 4 Then lbl2.Caption = ""
If Index = 8 Then lbl3.Caption = ""
If Len(Trim(txt1(Index))) = 0 Then Exit Sub
Select Case Index
Case 2
         If PUB_CheckKeyInDate(txt1(Index)) < 0 Then Cancel = True: txt1(Index).SetFocus: txt1_GotFocus Index
Case 3
         If PUB_CheckKeyInDate(txt1(Index)) = 0 Then
            'Modify by Morgan 2010/8/17 百年蟲
            'If txt1(2) <> "" And (txt1(3) < txt1(2)) Then
            If txt1(2) <> "" And Val(txt1(3)) < Val(txt1(2)) Then
               MsgBox "結束日期必需大於開始日期！", vbCritical
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
            End If
         Else
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
         End If
Case 1
         If Val(txt1(Index)) > 3 Or Val(txt1(Index)) < 1 Then
            MsgBox "類別請介於 1 ~3 ！", vbCritical
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
         End If
Case 4
      lbl2.Caption = GetStaffName(txt1(Index))
      If Trim(lbl2.Caption) = "" Then
          MsgBox "請輸入正確修改人員編號！", vbCritical
          txt1(Index).SetFocus
          txt1_GotFocus Index
          Cancel = True
      End If
Case 8
      lbl3.Caption = GetStaffName(txt1(Index))
      If Trim(lbl3.Caption) = "" Then
          'Modified by Lydia 2017/06/02
          'MsgBox "請輸入正確承辦人編號！", vbCritical
          MsgBox "請輸入正確承辦/繪圖人員編號！", vbCritical
          txt1(Index).SetFocus
          txt1_GotFocus Index
          Cancel = True
      End If
Case Else
End Select
End Sub

'報表列印
Private Sub PrintData()
   
   Dim ii As Integer
   
   GetPleft
   Page = 1
   PrintTitle

   With grd1
      For ii = 1 To .Rows - 1
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 0)
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 1)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 2)
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 3)
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 4)
         'Modified by Lydia 2017/06/02
         'Printer.CurrentX = PLeft(5) + 1000 - Printer.TextWidth(Format(.TextMatrix(ii, 5), "0.0"))
         'Printer.CurrentY = iPrint
         'Printer.Print Format(.TextMatrix(ii, 5), "0.0")
         'Printer.CurrentX = PLeft(6) + 1000 - Printer.TextWidth(Format(.TextMatrix(ii, 6), "0.0"))
         'Printer.CurrentY = iPrint
         'Printer.Print Format(.TextMatrix(ii, 6), "0.0")
         'Printer.CurrentX = PLeft(7)
         'Printer.CurrentY = iPrint
         'Printer.Print .TextMatrix(ii, 7)
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint
         Printer.Print "" & .TextMatrix(ii, 5)
         Printer.CurrentX = PLeft(6) + ((PLeft(7) - PLeft(6) - colGap) / 2) '置中
         Printer.CurrentY = iPrint
         Printer.Print Format(.TextMatrix(ii, 6), "0.0")
         Printer.CurrentX = PLeft(7) + ((PLeft(8) - PLeft(7) - colGap) / 2)  '置中 '(PLeft(8) - PLeft(7) - Printer.TextWidth(Format(.TextMatrix(ii, 7), "0.0")) - colGap)
         Printer.CurrentY = iPrint
         Printer.Print Format(.TextMatrix(ii, 7), "0.0")
         'Added by lydia 2017/06/02 理由
         Printer.CurrentX = PLeft(8)
         Printer.CurrentY = iPrint
         Printer.Print convForm("" & .TextMatrix(ii, 8), 38)
         'end 2017/06/02
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            'Modified by Lydia 2017/06/02 200改常數
            Printer.Print String(nDot, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
         End If
      Next ii
   End With
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Lydia 2017/06/02 200改常數
   Printer.Print String(nDot, "-")
   Printer.EndDoc
   
End Sub

Sub GetPleft()

   Erase PLeft
   
   PLeft(0) = 500
   'Modified by Lydia 2017/06/02 調整列印版面
'   PLeft(1) = PLeft(0) + 1250
'   PLeft(2) = PLeft(1) + 1250
'   PLeft(3) = PLeft(2) + 1250
'   PLeft(4) = PLeft(3) + 1250
'   PLeft(5) = PLeft(4) + 1250
'   PLeft(6) = PLeft(5) + 1500
'   PLeft(7) = PLeft(6) + 1900
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 12
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(15, "A")) + colGap
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + colGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(4, "　")) + colGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(5, "　")) + colGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + colGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(3, "　")) + colGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(5, "　")) + colGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(7, "　")) + colGap
   'end 2017/06/02
End Sub

Sub PrintTitle()

'GetPleft 'Remove by Lydia 2017/06/02
   
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "加乘註記歷史明細表"

   iPrint = iPrint + 500
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "修改日期：" & Format(ChangeTStringToTDateString(Me.txt1(2).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txt1(3).Text)

   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & GetPrjSalesNM(strUserNum)
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   iPrint = iPrint + 300
   'Modified by Lydia 2017/06/02 移動到最前面
   'Printer.CurrentX = 9000
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "類別：" & IIf(txt1(1) = "1", "承辦人", IIf(txt1(1) = "2", "草圖", IIf(txt1(1) = "3", "墨圖", "全部")))
   
   'Modified by Lydia 2017/06/02 移動到類別後面
   'Printer.CurrentX = 500
   Printer.CurrentX = 4500
   Printer.CurrentY = iPrint
   'Modified by Lydia 2017/06/02+判斷
   If Trim(Me.txt1(4).Text) <> "" Then Printer.Print "修改人員：" & Me.txt1(4).Text & " " & lbl2.Caption
   'end 2017/06/02
   
   'Modified by Lydia 2017/06/02 移動到類別後面
   'Printer.Print "收文號：" & txt1(0).Text
   'Printer.CurrentX = 6000
   Printer.CurrentX = 8000
   Printer.CurrentY = iPrint
   strExc(1) = txt1(0) & "-" & txt1(5) & IIf(Val(Trim(txt1(6) & txt1(7))) = 0, "", "-" & txt1(6) & "-" & txt1(7))
   If Replace(strExc(1), "-", "") <> "" Then
       Printer.Print "本所案號：" & strExc(1)
   End If
   'end 2017/06/02
   
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Lydia 2017/06/02 200改常數
   Printer.Print String(nDot, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   'Modified by Lydia 2017/06/02
   'Printer.Print "收文號"
   Printer.Print "本所案號"
   'Added by Lydia 2017/06/02
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "承辦/繪圖"
   'end 2017/06/02
   'Modified by Lydia 2017/06/02 index + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "修改人員"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "修改日期"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "修改時間"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "類別"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "原加乘註記"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "修改後加乘註記"
   Printer.CurrentX = PLeft(8)
   'end 2017/06/02
   Printer.CurrentY = iPrint
   Printer.Print "理由"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Lydia 2017/06/02 200改常數
   Printer.Print String(nDot, "-")
   iPrint = iPrint + 300
   
End Sub
