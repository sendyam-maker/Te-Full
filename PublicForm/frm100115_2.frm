VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100115_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以國籍查詢代理人/申請人"
   ClientHeight    =   5715
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8448
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7224
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4965
      Left            =   45
      TabIndex        =   7
      Top             =   735
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8758
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin VB.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   6972
      TabIndex        =   10
      Top             =   468
      Width           =   408
   End
   Begin VB.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   468
      Width           =   2256
   End
   Begin VB.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   852
      TabIndex        =   8
      Top             =   468
      Width           =   1788
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(1.代理人 2.申請人)"
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   468
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "往來日期："
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   468
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍 :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   468
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查詢別 :"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   468
      Width           =   630
   End
End
Attribute VB_Name = "frm100115_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String
Dim s As Integer, i As Integer, j As Integer, intK As Integer
Dim strSql As String, StrTest As String, StrTest2 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
GrdDataList.row = 0
GrdDataList.col = 0: GrdDataList.Text = "V"
GrdDataList.ColWidth(0) = 200
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 1: GrdDataList.Text = "國籍"
GrdDataList.ColWidth(1) = 1000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 2: GrdDataList.Text = "編號"
GrdDataList.ColWidth(2) = 1000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 3: GrdDataList.Text = "名稱"
GrdDataList.ColWidth(3) = 4500
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 4: GrdDataList.Text = ""
GrdDataList.ColWidth(4) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 5: GrdDataList.Text = ""
GrdDataList.ColWidth(5) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter

End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      If lbl1(2).Caption = "2" Then
          Me.Enabled = False
          For i = 1 To GrdDataList.Rows - 1
             GrdDataList.row = i
             GrdDataList.col = 1
             StrTest = GrdDataList.Text
             GrdDataList.col = 2
             If Len(Trim(StrTest)) <> 0 And Len(Trim(GrdDataList.Text)) <> 0 Then
                GrdDataList.col = 0
                If Trim(GrdDataList.Text) = "V" Then
                    GrdDataList.col = 0
                    GrdDataList.Text = ""
                    For j = 0 To GrdDataList.Cols - 1
                       GrdDataList.col = j
                       GrdDataList.CellBackColor = QBColor(15)
                    Next j
                    If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                    End If
                    GrdDataList.col = 2
                    Screen.MousePointer = vbHourglass
                    frm100101_11.Show
                    frm100101_11.Tag = Pub_RplStr(GrdDataList.Text)
                    frm100101_11.StrMenu
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                    Exit Sub
                End If
             End If
          Next i
          Me.Enabled = True
      Else
          Me.Enabled = False
          For i = 1 To GrdDataList.Rows - 1
             GrdDataList.row = i
             GrdDataList.col = 1
             StrTest = GrdDataList.Text
             GrdDataList.col = 2
             If Len(Trim(StrTest)) <> 0 And Len(Trim(GrdDataList.Text)) <> 0 Then
                GrdDataList.col = 0
                If Trim(GrdDataList.Text) = "V" Then
                    GrdDataList.col = 0
                     GrdDataList.Text = ""
                     For j = 0 To GrdDataList.Cols - 1
                          GrdDataList.col = j
                          GrdDataList.CellBackColor = QBColor(15)
                     Next j
                    If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                    End If
                    GrdDataList.col = 2
                    Screen.MousePointer = vbHourglass
                    frm100101_10.Show
                    frm100101_10.Tag = GrdDataList.Text ' StrTag  傳代理人代號
                    frm100101_10.StrMenu
                    Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                End If
            End If
          Next i
          Me.Enabled = True
      End If
Case 1
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2
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
      If lbl1(2).Caption = "2" Then
          Me.Enabled = False
          For i = 1 To GrdDataList.Rows - 1
             GrdDataList.row = i
             GrdDataList.col = 1
             StrTest = GrdDataList.Text
             GrdDataList.col = 2
             If Len(Trim(StrTest)) <> 0 And Len(Trim(GrdDataList.Text)) <> 0 Then
                GrdDataList.col = 0
                If Trim(GrdDataList.Text) = "V" Then
                    GrdDataList.col = 2
                    Screen.MousePointer = vbHourglass
                    frm100101_11.Show
                    'frm100101_11.Hide
                     
                    frm100101_11.Tag = GrdDataList.Text
                    frm100101_11.StrMenu
                    Screen.MousePointer = vbDefault
                    Me.Hide
                    'frm100101_11.Show
                    Do
                    DoEvents
                    If bolToEndByNick = True Then Unload Me: Exit Sub
                    Loop Until Not frm100101_11.Visible
                    Unload frm100101_11
                End If
                GrdDataList.col = 0
               GrdDataList.Text = ""
               For j = 0 To GrdDataList.Cols - 1
                  GrdDataList.col = j
                  GrdDataList.CellBackColor = QBColor(15)
               Next j

             End If
          Next i
          Me.Enabled = True
          Me.Show
      Else
          Me.Enabled = False
          For i = 1 To GrdDataList.Rows - 1
             GrdDataList.row = i
             GrdDataList.col = 1
             StrTest = GrdDataList.Text
             GrdDataList.col = 2
             If Len(Trim(StrTest)) <> 0 And Len(Trim(GrdDataList.Text)) <> 0 Then
                GrdDataList.col = 0
                If Trim(GrdDataList.Text) = "V" Then
                    GrdDataList.col = 2
                    Screen.MousePointer = vbHourglass
                    frm100101_10.Show
                    'frm100101_10.Hide
                     
                    frm100101_10.Tag = GrdDataList.Text ' StrTag  傳代理人代號
                    frm100101_10.StrMenu
                    Screen.MousePointer = vbDefault
                    Me.Hide
                    'frm100101_10.Show
                    Do
                        DoEvents
                        If bolToEndByNick = True Then Unload Me: Exit Sub
                    Loop Until Not frm100101_10.Visible
                    Unload frm100101_10
                    GrdDataList.col = 0
                     GrdDataList.Text = ""
                     For j = 0 To GrdDataList.Cols - 1
                          GrdDataList.col = j
                          GrdDataList.CellBackColor = QBColor(15)
                     Next j
                End If
            End If
          Next i
          Me.Enabled = True
          Me.Show
      End If
Case 1
     Me.Hide
Case 2
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
If bolFNation = False Then
    Label3.Caption = "(1.申請人)"
    Me.Caption = "以國籍查詢申請人"
Else
    Label3.Caption = "(1.代理人 2.申請人)"
End If
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()          '代理人
Me.Enabled = False
lbl1(0).Caption = frm100115_1.txt1(0) + "-" + frm100115_1.txt1(1)
lbl1(1).Caption = frm100115_1.txt1(2) + "-" + frm100115_1.txt1(3)
lbl1(2).Caption = frm100115_1.txt1(4)
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
'國籍
If Len(Trim(frm100115_1.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " AND FA10>='" & frm100115_1.txt1(0) & "' "
   StrSQL6 = StrSQL6 & " AND FA10>='" & frm100115_1.txt1(0) & "' "
   strSQL2 = strSQL2 & " AND FA10>='" & frm100115_1.txt1(0) & "' "
End If
If Len(Trim(frm100115_1.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND FA10<='" & frm100115_1.txt1(1) & "z' "
   StrSQL6 = StrSQL6 & " AND FA10<='" & frm100115_1.txt1(1) & "z' "
   strSQL2 = strSQL2 & " AND FA10<='" & frm100115_1.txt1(1) & "z' "
End If
If Len(Trim(frm100115_1.txt1(0))) <> 0 Or Len(Trim(frm100115_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100115_1.Label1 & frm100115_1.txt1(0) & "-" & frm100115_1.txt1(1) 'Add By Sindy 2010/11/4
End If
'add by nick 2004/07/26 加快速度
'往來日期
If Len(Trim(frm100115_1.txt1(2))) <> 0 And Len(Trim(frm100115_1.txt1(3))) <> 0 Then
       strSQL1 = strSQL1 & " AND ((CP05>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " and CP05<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & ") or (CP27>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " AND CP27<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & ")) "
       pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & frm100115_1.txt1(2) & "-" & frm100115_1.txt1(3) 'Add By Sindy 2010/11/4
Else
    'edit by nick 2004/07/26 加快速度
    If Len(Trim(frm100115_1.txt1(2))) <> 0 And Len(Trim(frm100115_1.txt1(3))) = 0 Then
        strSQL1 = strSQL1 & " AND ((CP05>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " AND CP05<=" & ChangeTStringToWString(ServerDate - 19110000) & " ) or (CP27>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " AND CP27<=" & ChangeTStringToWString(ServerDate - 19110000) & "))  "
        pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & frm100115_1.txt1(2) & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/4
    Else
         If Len(Trim(frm100115_1.txt1(2))) = 0 And Len(Trim(frm100115_1.txt1(3))) <> 0 Then
            strSQL1 = strSQL1 & " AND ((CP05<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " ) or (CP27<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " ))  "
            pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & "<=" & frm100115_1.txt1(3) 'Add By Sindy 2010/11/4
         End If
    End If
'    If Len(Trim(frm100115_1.txt1(2))) <> 0 Then
'       strSQL1 = strSQL1 & " AND CP05>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " "
'       StrSQL6 = StrSQL6 & " and CP27>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " "
'    End If
'    If Len(Trim(frm100115_1.txt1(3))) <> 0 Then
'       strSQL1 = strSQL1 & " AND CP05<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " "
'       StrSQL6 = StrSQL6 & " AND CP27<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " "
'    'Add By Cheng 2002/03/18
'    Else
'       If Len(frm100115_1.txt1(2).Text) > 0 Then
'          strSQL1 = strSQL1 & " AND CP05<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
'          StrSQL6 = StrSQL6 & " AND CP27<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
'       End If
'    End If
End If
'StrSQL2 = StrSQL1
'StrSQL3 = StrSQL1
'StrSQL4 = StrSQL1
'StrSQL5 = StrSQL1

If Len(Trim(frm100115_1.txt1(2))) = 0 And Len(Trim(frm100115_1.txt1(3))) = 0 Then
   strSql = "SELECT DISTINCT '' AS V,NVL(NA03,FA10) AS 國籍,FA01||FA02 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,'','' FROM FAGENT,NATION WHERE FA10=NA01(+) " & strSQL2
Else
          strSql = "SELECT DISTINCT '' AS V,NVL(NA03,FA10) AS 國籍,TM44 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,FAGENT,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND " & SQLNewFag("TM44", "FA") & " AND FA10=NA01(+) " & strSQL1
   'Modify By Cheng 200/03/18
   '將SQL語法中的 ALL 去掉
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,PA75 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,FAGENT,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND " & SQLNewFag("PA75", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,SP26 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,FAGENT,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND " & SQLNewFag("SP26", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,LC22 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,FAGENT,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND " & SQLNewFag("LC22", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,CP44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM CASEPROGRESS,FAGENT,NATION WHERE " & SQLNewFag("CP44", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL & " union select '' AS V,NVL(NA03,FA10) AS 國籍,TM44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,FAGENT,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND " & SQLNewFag("TM44", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,PA75 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,FAGENT,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND " & SQLNewFag("PA75", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,SP26 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,FAGENT,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND " & SQLNewFag("SP26", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,LC22 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,FAGENT,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND " & SQLNewFag("LC22", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,CP44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM CASEPROGRESS,FAGENT,NATION WHERE " & SQLNewFag("CP44", "FA") & " AND FA10=NA01(+) " & StrSQL6
'edit by nick 2004/07/26 加快速度
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,PA75 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,FAGENT,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND " & SQLNewFag("PA75", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,SP26 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,FAGENT,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND " & SQLNewFag("SP26", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,LC22 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,FAGENT,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND " & SQLNewFag("LC22", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,CP44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM CASEPROGRESS,FAGENT,NATION WHERE " & SQLNewFag("CP44", "FA") & " AND FA10=NA01(+) " & strSQL1
'   strSQL = strSQL & " union select '' AS V,NVL(NA03,FA10) AS 國籍,TM44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,FAGENT,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND " & SQLNewFag("TM44", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,PA75 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,FAGENT,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND " & SQLNewFag("PA75", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,SP26 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,FAGENT,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND " & SQLNewFag("SP26", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,LC22 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,FAGENT,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND " & SQLNewFag("LC22", "FA") & " AND FA10=NA01(+) " & StrSQL6
'   strSQL = strSQL + " union select '' AS V,NVL(NA03,FA10) AS 國籍,CP44 AS 編號,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) AS 名稱,'','' FROM CASEPROGRESS,FAGENT,NATION WHERE " & SQLNewFag("CP44", "FA") & " AND FA10=NA01(+) " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,FA10) AS 國籍,PA75 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,FAGENT,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND " & SQLNewFag("PA75", "FA") & " AND FA10=NA01(+) " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,FA10) AS 國籍,SP26 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,FAGENT,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND " & SQLNewFag("SP26", "FA") & " AND FA10=NA01(+) " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,FA10) AS 國籍,LC22 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,FAGENT,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND " & SQLNewFag("LC22", "FA") & " AND FA10=NA01(+) " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,FA10) AS 國籍,CP44 AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,'','' FROM CASEPROGRESS,FAGENT,NATION WHERE " & SQLNewFag("CP44", "FA") & " AND FA10=NA01(+) " & strSQL1
End If
strSql = strSql + " ORDER BY 國籍,名稱,編號 "

CheckOC
StrTest2 = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/4
    Me.Enabled = True
    cmdOK(0).Enabled = False
    Screen.MousePointer = vbDefault
    ShowNoData
    'Modify By Cheng 2003/07/30
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set GrdDataList.Recordset = adoRecordset
CheckOC
SetDataListWidth
Me.Enabled = True
End Sub

Sub StrMenu1()          '申請人
Me.Enabled = False
lbl1(0).Caption = frm100115_1.txt1(0) + "-" + frm100115_1.txt1(1)
lbl1(1).Caption = frm100115_1.txt1(2) + "-" + frm100115_1.txt1(3)
lbl1(2).Caption = frm100115_1.txt1(4)
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""

If Len(Trim(frm100115_1.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " AND CU10>='" & frm100115_1.txt1(0) & "' "
   StrSQL6 = StrSQL6 & " AND CU10>='" & frm100115_1.txt1(0) & "' "
   strSQL2 = strSQL2 & " AND CU10>='" & frm100115_1.txt1(0) & "' "
End If
If Len(Trim(frm100115_1.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND CU10<='" & frm100115_1.txt1(1) & "z' "
   StrSQL6 = StrSQL6 & " AND CU10<='" & frm100115_1.txt1(1) & "z' "
   strSQL2 = strSQL2 & " AND CU10<='" & frm100115_1.txt1(1) & "z' "
End If
If Len(Trim(frm100115_1.txt1(0))) <> 0 Or Len(Trim(frm100115_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100115_1.Label1 & frm100115_1.txt1(0) & "-" & frm100115_1.txt1(1) 'Add By Sindy 2010/11/4
End If

If Len(Trim(frm100115_1.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " "
   StrSQL6 = StrSQL6 & " and CP27>=" & ChangeTStringToWString(frm100115_1.txt1(2)) & " "
End If
If Len(Trim(frm100115_1.txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " "
   StrSQL6 = StrSQL6 & " AND CP27<=" & ChangeTStringToWString(frm100115_1.txt1(3)) & " "
   If Len(frm100115_1.txt1(2).Text) > 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & frm100115_1.txt1(2) & "-" & frm100115_1.txt1(3) 'Add By Sindy 2010/11/4
   Else
      pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & "<=" & frm100115_1.txt1(3)  'Add By Sindy 2010/11/4
   End If
Else
   If Len(frm100115_1.txt1(2).Text) > 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
      StrSQL6 = StrSQL6 & " AND CP27<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
      pub_QL05 = pub_QL05 & ";" & frm100115_1.Label4 & frm100115_1.txt1(2) & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/4
   End If
End If

If Len(Trim(frm100115_1.txt1(2))) = 0 And Len(Trim(frm100115_1.txt1(3))) = 0 Then
   strSql = "SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,CU01||CU02 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM CUSTOMER,NATION WHERE CU10=NA01(+) " & strSQL2
Else
   strSql = "SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,TM23 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM23", "CU") & " " & strSQL1
   strSql = strSql + " union SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,TM78 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM78", "CU") & " " & strSQL1
   strSql = strSql + " union SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,TM79 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM79", "CU") & " " & strSQL1
   strSql = strSql + " union SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,TM80 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM80", "CU") & " " & strSQL1
   strSql = strSql + " union SELECT DISTINCT '' AS V,NVL(NA03,CU10) AS 國籍,TM81 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM81", "CU") & " " & strSQL1
   
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA26 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA27 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA27", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA28 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA28", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA29 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA29", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA30 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA30", "CU") & " " & strSQL1
   
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP08 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP08", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP58 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP58", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP59 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP59", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP65 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP65", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP66 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP66", "CU") & " " & strSQL1
   
   'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC11 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC11", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC43 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC43", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC44 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC44", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC45 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC45", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC46 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC46", "CU") & " " & strSQL1
   
   'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC05 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC05", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC24 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC24", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC25 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC25", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC26 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC26", "CU") & " " & strSQL1
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC27 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC27", "CU") & " " & strSQL1
   
   strSql = strSql & " union select '' AS V,NVL(NA03,CU10) AS 國籍,TM23 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM23", "CU") & " " & StrSQL6
   strSql = strSql & " union select '' AS V,NVL(NA03,CU10) AS 國籍,TM78 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM78", "CU") & " " & StrSQL6
   strSql = strSql & " union select '' AS V,NVL(NA03,CU10) AS 國籍,TM79 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM79", "CU") & " " & StrSQL6
   strSql = strSql & " union select '' AS V,NVL(NA03,CU10) AS 國籍,TM80 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM80", "CU") & " " & StrSQL6
   strSql = strSql & " union select '' AS V,NVL(NA03,CU10) AS 國籍,TM81 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CU10=NA01(+) AND " & SQLNewFag("TM81", "CU") & " " & StrSQL6
   
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA26 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA27 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA27", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA28 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA28", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA29 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA29", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,PA30 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CU10=NA01(+) AND " & SQLNewFag("PA30", "CU") & " " & StrSQL6
   
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP08 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP08", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP58 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP58", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP59 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP59", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP65 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP65", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,SP66 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) aS 名稱,'','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CU10=NA01(+) AND " & SQLNewFag("SP66", "CU") & " " & StrSQL6
   
   'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC11 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC11", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC43 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC43", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC44 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC44", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC45 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC45", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,LC46 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CU10=NA01(+) AND " & SQLNewFag("LC46", "CU") & " " & StrSQL6
   
   'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC05 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC05", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC24 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC24", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC25 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC25", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC26 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC26", "CU") & " " & StrSQL6
   strSql = strSql + " union select '' AS V,NVL(NA03,CU10) AS 國籍,HC27 AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,'','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CU10=NA01(+) AND " & SQLNewFag("HC27", "CU") & " " & StrSQL6
End If
strSql = strSql + " ORDER BY 國籍,名稱,編號 "
CheckOC
StrTest2 = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/4
    Me.Enabled = True
    cmdOK(0).Enabled = False
    ShowNoData
    Screen.MousePointer = vbDefault
    'Modify By Cheng 2003/07/30
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set GrdDataList.Recordset = adoRecordset
CheckOC
SetDataListWidth
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100115_2 = Nothing
End Sub

Private Sub grdDataList_SelChange()
GrdDataList.Visible = False
GrdDataList.row = GrdDataList.MouseRow
GrdDataList.col = 0
If GrdDataList.row <> 0 Then
If GrdDataList.Text = "V" Then
     GrdDataList.Text = ""
     For i = 0 To GrdDataList.Cols - 1
          GrdDataList.col = i
          GrdDataList.CellBackColor = QBColor(15)
    Next i
Else
     GrdDataList.Text = "V"
     For i = 0 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = &HFFC0C0
     Next i

End If
End If
GrdDataList.Visible = True
End Sub
