VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060101 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專分案"
   ClientHeight    =   5750
   ClientLeft      =   110
   ClientTop       =   990
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9340
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7560
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8388
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComAllData 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6336
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton ComUCase 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5295
      TabIndex        =   7
      Top             =   70
      Width           =   1020
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   4080
      TabIndex        =   6
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   1332
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   4752
      Begin VB.TextBox txtGDate1 
         Enabled         =   0   'False
         Height          =   312
         Index           =   1
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton Option1 
         Caption         =   "以前未分案 :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1430
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   1190
      End
      Begin VB.OptionButton Option1 
         Caption         =   "收文日期："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1190
      End
      Begin VB.TextBox txtGDate1 
         Enabled         =   0   'False
         Height          =   312
         Index           =   0
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   0
         Top             =   240
         Width           =   1092
      End
      Begin VB.TextBox txtcp01 
         Height          =   312
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtcp02 
         Height          =   312
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtcp03 
         Height          =   312
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtcp04 
         Height          =   312
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1332
      Left            =   4980
      TabIndex        =   11
      Top             =   540
      Width           =   4212
      Begin VB.OptionButton Option6 
         Caption         =   "接洽及內部收文單"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option7 
         Caption         =   "主管機關來函"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   540
         Width           =   2055
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   1950
      Width           =   9075
      _ExtentX        =   15998
      _ExtentY        =   6368
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm060101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/11 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean
' 搜尋的方式 1:所有資料 2:未分案
Dim m_QueryType As Integer
Dim m_bCross As Boolean '跨部門 Added by Morgan 2012/5/16


Private Sub cmdSearch_Click()
 Dim i As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = "v"
      Next
   End With
End Sub

Private Sub ComAllData_Click()
   If CheckCP02 = False Then Exit Sub 'Add by Morgan 2004/10/21
   ' 91.05.16
   Screen.MousePointer = vbHourglass
   ' 90.07.06 modify by louis
   m_QueryType = 1
    If CheckChoese(2) Then
        PutDataInGrid
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
   GridHead
   ' 91.05.16
   Screen.MousePointer = vbDefault
End Sub

Private Sub ComBack_Click()
   blnIsFormBack = False
   Unload Me
End Sub

Private Sub ComSure_Click()
 Dim i As Integer
 'Add By Cheng 2001/12/25
 Dim ii As Integer '回圈流水號
 'Added by Lydia 2018/05/21 改從SetParent傳收文號
 Dim IntTot As Integer '筆數
 Dim StrCaseList  As String '本所案號
 Dim StrCPList As String '收文號
 
   With MSHFlexGrid1
      If .Rows < 2 Then Exit Sub
      For i = 1 To .Rows - 1
         .Visible = False
        'Modify By Cheng 2003/09/17
'         If .TextMatrix(i, 0) = "v" Then
         If .TextMatrix(i, 0) = "v" And .TextMatrix(i, 1) <> "" Then
             'Modified by Lydia 2018/05/21 改從SetParent傳收文號
             'Exit For
             StrCaseList = StrCaseList & "," & .TextMatrix(i, 7)   '本所案號
             StrCPList = StrCPList & "," & .TextMatrix(i, 1)   '收文號
             IntTot = IntTot + 1
             'end 2018/05/21
         Else
            'Modified by Lydia 2018/05/21 最後一筆判斷
            'If i = .Rows - 1 Then
            If i = .Rows - 1 And (StrCaseList = "" Or StrCPList = "") Then
               MsgBox "請點選欲分案資料"
               .Visible = True
               Exit Sub
             End If
         End If
      Next
      .Visible = True
   End With
   
   'Added by Morgan 2012/4/26
   If Option1(1).Value = True And (txtcp01 = "P" Or txtcp01 = "PS" Or txtcp01 = "CFP" Or txtcp01 = "CPS") Then
      If m_bCross Then
         Call frm060101_3.SetParent(Me, IntTot, Mid(StrCaseList, 2), Mid(StrCPList, 2)) 'Added by Lydia 2018/05/21  改從SetParent傳收文號
         frm060101_3.Show
         Me.Hide
      Else
         MsgBox "權限不足，無法作業！"
         Exit Sub
      End If
   Else
   'end 2012/4/26
      
      Call frm060101_1.SetParent(Me, IntTot, Mid(StrCaseList, 2), Mid(StrCPList, 2)) 'Added by Lydia 2018/05/21  改從SetParent傳收文號
      frm060101_1.Show
      
'Removed by Morgan 2012/6/15 下一畫面也有控制
'      'Add by Morgan 2004/4/20
'      '若為主管機關來函時，轉本所案號不可輸入
'      If Option7.Value = True Then
'         frm060101_1.Text1(2).Enabled = False
'         frm060101_1.Text1(25).Enabled = False
'         frm060101_1.Text1(26).Enabled = False
'         frm060101_1.Text1(27).Enabled = False
'      End If
'end 2012/6/15

      Me.Hide
      
      'Add By Cheng 2001/12/25
      DoEvents
      For ii = 0 To Forms.Count - 1
         '專利案件基本資料維護(frm050701)
         If Forms(ii).Name = "frm050701" Then
            frm060101_1.ZOrder 1
            'Add By Cheng 2002/01/03
            'Removed by Morgan 2012/6/18
            'frm050701.SelectToolbarButtom
            
            Exit For
         End If
      Next ii
      
   End If 'Added by Morgan 2012/4/26
   
   

End Sub

'Add by Morgan 2004/10/21 檢查本所號
Private Function CheckCP02() As Boolean
   If Option1(1).Value = True Then
      If txtcp02.Enabled = True And Len(txtcp02.Text) <> 6 Then
         MsgBox "本所案號輸入錯誤！"
         txtcp02.SetFocus
         txtcp02_GotFocus
         CheckCP02 = False
         Exit Function
      End If
   End If
   CheckCP02 = True
End Function

Private Sub ComUCase_Click()
   If CheckCP02 = False Then Exit Sub 'Add by Morgan 2004/10/21
   
   ' 90.07.06 modify by louis
   m_QueryType = 2
   ' 91.05.16
   Screen.MousePointer = vbHourglass
    If CheckChoese(1) Then
        PutDataInGrid
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
   GridHead
   'Add By Cheng 2002/04/23
   '若只有搜尋到一筆資料, 則直接進入下一畫面
   'Modify by Morgan 2003/12/23
   'If Me.MSHFlexGrid1.Rows = 2 Then
   If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
   'Modify end 2003/12/23
      cmdSearch_Click
      ComSure_Click
   End If
   ' 91.05.16
   Screen.MousePointer = vbDefault
End Sub

Sub ComUCase2()
   If m_QueryType = 2 Then
      If CheckChoese(1) Then PutDataInGrid
   Else
      If CheckChoese(2) Then PutDataInGrid
   End If
   GridHead
End Sub

Private Sub Form_Activate()
 Dim i As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .TextMatrix(i, 0) = ""
      Next
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtGDate1(0).Text = strSrvDate(2)
   txtGDate1(1) = txtGDate1(0)
    'Modify By Cheng 2002/12/10
'   Option1_Click 1
   txtcp01 = "FCP"
   'Add By Cheng 2002/12/10
   SendKeys "{Tab}"
    'Add By Cheng 2003/09/17
    'Begin
    Me.MSHFlexGrid1.Cols = 18
    GridHead
    'End
    
    m_bCross = IsUserHasRightOfFunction(Me.Name, strCrossDept, False)
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "案件名稱 "
      For i = 7 To 16
         .col = i: .ColWidth(i) = 0
      Next
      'Added by Morgan 2012/8/8
      .col = 17: .ColWidth(17) = 800: .Text = "已分案"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(17) = flexAlignCenterCenter
      'end 2012/8/8
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   'Add By Cheng 2002/07/18
   Set frm060101 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0, 1
   ComSure.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   GridClick MSHFlexGrid1, intLastRow, 0, 1
End Sub

Private Sub Option1_Click(Index As Integer)
   txtGDate1(0).Enabled = False
   txtGDate1(1).Enabled = False
   txtcp01.Enabled = False
   txtcp02.Enabled = False
   txtcp03.Enabled = False
   txtcp04.Enabled = False
   Select Case Index
      Case 0
         txtGDate1(0).Enabled = True
         txtGDate1(1).Enabled = True
         'Add By Cheng 2002/12/10
         txtGDate1(0).SetFocus
      Case 1
         txtcp01.Enabled = True
         txtcp02.Enabled = True
         txtcp03.Enabled = True
         txtcp04.Enabled = True
         'Add By Cheng 2002/12/10
         If Me.txtcp01.Text = "" Then
            Me.txtcp01.SetFocus
         Else
            Me.txtcp02.SetFocus
         End If
      Case 2
         
   End Select
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtcp01 <> "" Then
      txtcp01 = UCase(txtcp01)
      'Removed by Morgan 2012/6/7
      'If ChkSysName(txtcp01) = True Then
         'Modified by Morgan 2012/5/15 +P,PS,CFP,CPS
         'If txtcp01 <> "FCP" And txtcp01 <> "FG" Then
         'Modified by Lydia 2016/06/16 改成模組
         'If txtcp01 <> "FCP" And txtcp01 <> "FG" And txtcp01 <> "P" And txtcp01 <> "PS" And txtcp01 <> "CFP" And txtcp01 <> "CPS" Then
         If PUB_CheckFCPsys(txtcp01.Text) = False Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      'Else
      '   Cancel = True
      'End If
   Else
      If Option1(1).Value = True Then
         MsgBox "系統別不可空白，請重新輸入 !"
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub PutDataInGrid()
 Dim i As Integer, strPropertyName As String, strTempName As String, j As Integer
   With MSHFlexGrid1
      .Visible = False
      For i = 1 To .Rows - 1
         .row = i
         For j = 8 To 12
            strExc(j - 8) = .TextMatrix(i, j)
         Next
         .col = 6
         If strExc(0) <> "1" Then
            If strExc(1) = "Y" Then
               If strExc(2) <> "" Then
                  .Text = strExc(2)
               Else
                  If strExc(3) <> "" Then
                     .Text = strExc(3)
                  Else
                     .Text = strExc(4)
                  End If
               End If
            End If
         End If
         '逾本所
         If .TextMatrix(i, 13) <> "" And .TextMatrix(i, 13) <= strSrvDate(1) And .TextMatrix(i, 16) = "" Then
            For j = 0 To 16
               .col = j
               .CellBackColor = &H8080FF
            Next
         '閉卷
         ElseIf .TextMatrix(i, 14) = "Y" Then
            For j = 0 To 16
               .col = j
               .CellBackColor = &HFFFF&
            Next
         '取消收文
         ElseIf .TextMatrix(i, 15) <> "" Then
            For j = 0 To 16
               .col = j
               .CellBackColor = &HE0E0E0
            Next
         End If
      Next
      .Visible = True
   End With
End Sub

Private Function CheckChoese(ByRef i As Integer) As Boolean
Dim LcTmp As String
    
    CheckChoese = False
   '若按下未分案按鈕且不是選擇"以前未分案"條件
   If i = 1 And Option1(2).Value = False Then
      'Modified by Morgan 2012/8/8
      'strExc(0) = " AND CP14 IS NULL AND CP10 <> '907' AND CP10<>'913' "
      strExc(0) = " AND CP122||CP27 IS NULL AND CP10 <> '907' AND CP10<>'913' "
   Else
      strExc(0) = " AND CP10 <> '907' AND CP10<>'913' "
   End If
   '若條件為"接洽及內部收文單"
   If Option6.Value Then
      'Modify By Cheng 2002/04/12
'      strExc(0) = strExc(0) & " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B')"
      strExc(0) = strExc(0) & " and ( cp09<'C' )"
   '若條件為"主管機關來函"
   ElseIf Option7.Value Then
      'Modify By Cheng 2002/04/12
'      strExc(0) = strExc(0) & " and substr(cp09,1,1)='C'"
      strExc(0) = strExc(0) & " and cp09>'C' "
   End If
   
   'Modify by Morgan 2010/8/12 百年蟲 " & SQLDate("CP05") & "-->substrb(' '||sqldatet(cp05),-9)
   '若條件為"收文日期"
   If Option1(0).Value Then
        If Me.txtGDate1(0).Text <> "" Then
            If CheckIsTaiwanDate(Me.txtGDate1(0).Text) = False Then Exit Function
        End If
      If txtGDate1(1) = "" Then MsgBox "請輸入日期 !", vbInformation: Exit Function
        If CheckIsTaiwanDate(Me.txtGDate1(1).Text) = False Then Exit Function
      If txtGDate1(0) = "" Then
         strExc(0) = strExc(0) & " and cp05<=" & TransDate(txtGDate1(1), 2)
      Else
         strExc(0) = strExc(0) & " and cp05 between " & TransDate(txtGDate1(0), 2) & _
            " and " + TransDate(txtGDate1(1), 2)
      End If
      'Modify By Cheng 2001/12/20
'      strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37,CP38," & _
'         "CP39,CP06,PA57,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14=ST01(+) AND CP01 IN ('FCP','') AND CP01=CPM01(+) and CP10=CPM02(+) AND " & _
'         "CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0) & " UNION " & _
'         "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'         "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14=ST01(+) AND CP01 IN ('FG','') AND CP01=CPM01(+) and CP10=CPM02(+) AND " & _
'         "CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
      'Modified by Morgan 2012/8/8 +CP122
      strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
         "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37,CP38," & _
         "CP39,CP06,PA57,CP57,CP27,CP122 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
         "CP14=ST01(+) AND CP01 IN ('FCP','') AND CP01=CPM01(+) and CP10=CPM02(+) AND " & _
         "CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0) & " UNION " & _
         "SELECT '',CP09," & SQLDate("CP05") & " As Column2," & ChgCaseprogress("", 1) & "," & _
         "CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
         "CP06,SP15,CP57,CP27,CP122 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
         "CP14=ST01(+) AND CP01 IN ('FG','') AND CP01=CPM01(+) and CP10=CPM02(+) AND " & _
         "CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
   '若條件為"本所案號"
   ElseIf Option1(1).Value Then
      
      If txtcp03.Text = "" Then txtcp03 = "0"
      If txtcp04.Text = "" Then txtcp04.Text = "00"
      LcTmp = ChgCaseprogress(txtcp01 & txtcp02 & txtcp03 & txtcp04)
      If txtcp01 = "FCP" Then
         'Modify By Cheng 2001/12/20
'         strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'            "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37," & _
'            "CP38,CP39,CP06,PA57,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'            "CP01 IN ('FCP','') AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
'            "CP10=CPM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0)
         'Modified by Morgan 2012/8/8 +CP122
         strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
            "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37," & _
            "CP38,CP39,CP06,PA57,CP57,CP27,CP122 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
            "CP01 IN ('FCP','') AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
            "CP10=CPM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0)
      ElseIf txtcp01 = "FG" Then
         'Modify By Cheng 2001/12/20
'         strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'            "CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'            "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'            "CP01 IN ('FG','') AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
'            "CP10=CPM02(+) AND CP01=SP01 AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
         'Modified by Morgan 2012/8/8 +CP122
         strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
            "CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
            "CP06,SP15,CP57,CP27,CP122 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
            "CP01 IN ('FG','') AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
            "CP10=CPM02(+) AND CP01=SP01 AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
            
      'Added by Morgan 2012/4/26
      '本所案號可輸入FMP案
      ElseIf txtcp01 = "P" Or txtcp01 = "CFP" Then
         'Modified by Morgan 2012/8/8 +CP122
         strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
            "decode(pa01,'CFP',CPM03,DECODE(pa09,'000',CPM03,CPM04)) CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37," & _
            "CP38,CP39,CP06,PA57,CP57,CP27,CP122 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
            "CP01 in ('P','CFP') AND substr(cp12,1,1)='F' AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
            "CP10=CPM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0)
            
      ElseIf txtcp01 = "PS" Or txtcp01 = "CPS" Then
         'Modified by Morgan 2012/8/8 +CP122
         strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
            "CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
            "CP06,SP15,CP57,CP27,CP122 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
            "CP01 IN ('PS','CPS') AND substr(cp12,1,1)='F' AND " & LcTmp & " AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
            "CP10=CPM02(+) AND CP01=SP01 AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
      'end 2012/4/26
      End If
   '若條件為"以前未分案 "
   ElseIf Option1(2).Value Then
      'Modify By Cheng 2001/12/20
'      strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37,CP38," & _
'         "CP39,CP06,PA15,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14 IS NULL AND CP01 IN ('FCP','') AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
'         "CP10=CPM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0) & " UNION " & _
'         "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & _
'         ",CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'         "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14 IS NULL AND CP01 IN ('FG','') AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
'         "CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
      'Modified by Morgan 2012/8/8 +CP122
      strExc(0) = "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & "," & _
         "CPM03,ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23,CP31,CP37,CP38," & _
         "CP39,CP06,PA15,CP57,CP27,CP122 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
         "CP14 IS NULL AND CP01 IN ('FCP','') AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
         "CP10=CPM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & strExc(0) & " UNION " & _
         "SELECT '',CP09,substrb(' '||sqldatet(cp05),-9) As Column2," & ChgCaseprogress("", 1) & _
         ",CPM03,ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
         "CP06,SP15,CP57,CP27,CP122 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
         "CP14 IS NULL AND CP01 IN ('FG','') AND CP14=ST01(+) AND CP01=CPM01(+) AND " & _
         "CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)" & strExc(0)
   End If
   
   'Add By Cheng 2001/12/20
   '依收文日由大到小, 收文號由大到小
   strExc(0) = strExc(0) & " Order By Column2 Desc, Cp09 Desc"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   If intI = 1 Then
        CheckChoese = True
   Else
        CheckChoese = False
        GridHead
   End If

End Function

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtGDate1_GotFocus(Index As Integer)
   TextInverse txtGDate1(Index)
End Sub

Private Sub txtGDate1_Validate(Index As Integer, Cancel As Boolean)
   If txtGDate1(Index).Text <> "" Then
        'Modify By Cheng 2003/09/17
'      If Not ChkDate(txtGDate1(Index)) Then
      If Not CheckIsTaiwanDate(txtGDate1(Index)) Then
         Cancel = True
      Else
         If Index = 1 And txtGDate1(0) <> "" And txtGDate1(1) <> "" Then
            If Not ChkRange(txtGDate1(0), txtGDate1(1), "收文") Then Cancel = True
         End If
      End If
      If Cancel Then TextInverse txtGDate1(Index)
   End If
End Sub

' 90.07.06 modify by louis (會到該畫面以原有條件再重新查詢一次)
Public Sub RefreshData()
   Select Case m_QueryType
      Case 1:
         ComAllData_Click
      Case 2:
         ComUCase_Click
      Case Else:
   End Select
End Sub
