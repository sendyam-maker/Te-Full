VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm20 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品名稱查詢"
   ClientHeight    =   5724
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9204
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9204
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgList 
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   1950
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   6583
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox Check1 
      Caption         =   "日文商品名稱"
      Height          =   276
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   1080
      Width           =   1485
   End
   Begin VB.CheckBox Check1 
      Caption         =   "英文商品名稱"
      Height          =   276
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   780
      Width           =   1485
   End
   Begin VB.CheckBox Check1 
      Caption         =   "中文商品名稱"
      Height          =   276
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   348
      Index           =   6
      Left            =   6744
      TabIndex        =   4
      Top             =   696
      Width           =   996
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  '沒有框線
      Height          =   396
      Left            =   0
      ScaleHeight     =   396
      ScaleWidth      =   9132
      TabIndex        =   5
      Top             =   0
      Width           =   9132
      Begin VB.CommandButton cmd 
         Caption         =   "結束(&X)"
         Height          =   348
         Index           =   0
         Left            =   7908
         TabIndex        =   6
         Top             =   24
         Width           =   996
      End
      Begin VB.CommandButton cmd 
         Caption         =   "匯出檔案(&E)"
         Height          =   348
         Index           =   7
         Left            =   6828
         TabIndex        =   13
         Top             =   24
         Width           =   1080
      End
      Begin VB.CommandButton cmd 
         Caption         =   "整批刪除(&C)"
         Height          =   348
         Index           =   8
         Left            =   4704
         TabIndex        =   14
         Top             =   24
         Width           =   1110
      End
      Begin VB.CommandButton cmd 
         Caption         =   "單筆刪除(&D)"
         Height          =   348
         Index           =   4
         Left            =   3588
         TabIndex        =   10
         Top             =   24
         Width           =   1110
      End
      Begin VB.CommandButton cmd 
         Caption         =   "修改(&E)"
         Height          =   348
         Index           =   3
         Left            =   2592
         TabIndex        =   9
         Top             =   24
         Width           =   996
      End
      Begin VB.CommandButton cmd 
         Caption         =   "單筆新增(&A)"
         Height          =   348
         Index           =   2
         Left            =   1290
         TabIndex        =   8
         Top             =   24
         Width           =   1284
      End
      Begin VB.CommandButton cmd 
         Caption         =   "整批新增(&B)"
         Height          =   348
         Index           =   1
         Left            =   15
         TabIndex        =   7
         Top             =   24
         Width           =   1284
      End
      Begin VB.CommandButton cmd 
         Caption         =   "快顯(&S)"
         Height          =   348
         Index           =   5
         Left            =   5832
         TabIndex        =   11
         Top             =   24
         Width           =   996
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "國際分類"
      Height          =   276
      Index           =   2
      Left            =   180
      TabIndex        =   18
      Top             =   1410
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "(查詢多個國際分類, 以逗號區隔開來)"
      Height          =   270
      Left            =   5580
      TabIndex        =   19
      Top             =   1413
      Width           =   2925
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1380
      Width           =   3840
      VariousPropertyBits=   671105051
      Size            =   "6773;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   1050
      Width           =   4788
      VariousPropertyBits=   671105051
      Size            =   "8446;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   744
      Width           =   4788
      VariousPropertyBits=   671105051
      Size            =   "8446;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   432
      Width           =   4788
      VariousPropertyBits=   671105051
      Size            =   "8446;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢結果：符合條件的資料，共0筆!!!"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   2970
   End
End
Attribute VB_Name = "frm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/10 msgList=> MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
'Memo by Lydia 2021/09/24 改成Form2.0 ; msgList改字型=新細明體-ExtB、Text1(index)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit
Dim intLastRow As Integer 'Added by Lydia 2022/03/15

Private Sub Check1_Click(Index As Integer)
    Select Case Index
    Case 0
        If Me.Check1(Index).Value = vbChecked Then
            Me.Text1(0).SetFocus
        Else
            Me.Text1(0).Text = ""
        End If
    Case 1
        If Me.Check1(Index).Value = vbChecked Then
            Me.Text1(1).SetFocus
        Else
            Me.Text1(1).Text = ""
        End If
    Case 2
        If Me.Check1(Index).Value = vbChecked Then
            Me.Text1(2).SetFocus
        Else
            Me.Text1(2).Text = ""
        End If
    'add by nick 2004/10/12
    Case 3
        If Me.Check1(Index).Value = vbChecked Then
            Me.Text1(3).SetFocus
        Else
            Me.Text1(3).Text = ""
        End If
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
'Add By Cheng 2003/03/27
Dim ii As Long
Dim strFileName As String
Dim strMid As String 'Added by Lydia 2022/03/15

   Select Case Index
   Case 0 '離開
      Unload Me
   Case 1 '整批新增
      frm20_1.Show
      frm20_1.Tag = Index
      frm20_1.Caption = "商品名稱查詢--整批新增"
      Me.Hide
   Case 2 '單筆新增
      frm20_2.Show
      frm20_2.m_strFormStatus = "1"
      frm20_2.DisplayProperty
      Me.Hide
   Case 3 '修改
      'edit by nick 2004/10/12
      'If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 0) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "") Then
      If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 3) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 4) <> "") Then
         frm20_2.Show
         frm20_2.m_strFormStatus = "2"
'add by nick 2004/10/12
'         frm20_2.DisplayProperty Me.msgList.TextMatrix(Me.msgList.RowSel, 0), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 1), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 2)
         'Modified by Lydia 2022/03/15 +流水號Me.msgList.TextMatrix(Me.msgList.RowSel, 5)
         frm20_2.DisplayProperty Me.msgList.TextMatrix(Me.msgList.RowSel, 1), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 2), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 3), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 4), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 5)
         Me.Hide
      Else
         MsgBox "請先點選欲修改的資料!!!", vbExclamation + vbOKOnly
      End If
   Case 4 '刪除
      DeleteData
   Case 5 '快顯
      'edit by nick 2004/10/12
      'If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 0) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "") Then
      If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 3) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 4) <> "") Then
         frm20_2.Show
         frm20_2.m_strFormStatus = "3"
         frm20_2.m_row = Me.msgList.RowSel
'edit by nick 2004/10/12
'         frm20_2.DisplayProperty Me.msgList.TextMatrix(Me.msgList.RowSel, 0), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 1), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 2)
         'Modified by Lydia 2022/03/15 +流水號Me.msgList.TextMatrix(Me.msgList.RowSel, 5)
         frm20_2.DisplayProperty Me.msgList.TextMatrix(Me.msgList.RowSel, 1), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 2), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 3), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 4), _
                                 Me.msgList.TextMatrix(Me.msgList.RowSel, 5)
         Me.Hide
      Else
         MsgBox "請先點選欲快顯的資料!!!", vbExclamation + vbOKOnly
      End If
   Case 6 '查詢
      QueryData
    'Add By Cheng 2003/03/27
    Case 7 '匯出檔案
      'edit by nick 2004/10/12
      'If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 0) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "") Then
      If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 3) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 4) <> "") Then
'            'Add By Cheng 2003/05/19
'            Dim strSQLA As String
'            Dim rsA As New ADODB.Recordset
'            Dim strKind As String '國際分類
'            Dim strFile As String '檔案名稱
'            strSQLA = "Select * From TrademarkMerchandiseName Order By 1, 3 ,2 "
'            rsA.CursorLocation = adUseClient
'            rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'                strKind = "" & rsA.Fields(0).Value
'                Open App.Path & "\Class" & strKind & ".txt" For Append As #1
'                While Not rsA.EOF
'                    If strKind <> "" & rsA.Fields(0).Value Then
'                        strKind = "" & rsA.Fields(0).Value
'                        Close #1
'                        Open App.Path & "\Class" & strKind & ".txt" For Append As #1
'                    End If
'                    Print #1, "" & rsA.Fields(2).Value & ":" & rsA.Fields(1).Value
'                    rsA.MoveNext
'                Wend
'            End If
'            Close #1
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
            strFileName = InputBox("請輸入檔案名稱???")
            If Trim("" & strFileName) <> "" Then
                Screen.MousePointer = vbHourglass
                'Modified by Lydia 2022/03/15 改變輸出格式
                'Open App.path & "\" & strFileName & ".doc" For Append As #12
                If Dir(strExcelPath & strFileName & ".txt") = MsgText(601) Then
                    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                        MkDir strExcelPath
                    End If
                Else
                    Kill strExcelPath & strFileName & ".txt"
                End If
                'end 2022/03/15
                For ii = 1 To Me.msgList.Rows - 1
                    'Modify By Cheng 2004/02/11
                    '匯出格式44:英文:中文
'                    Print #12, Me.msgList.TextMatrix(ii, 0)
'                    Print #12, Me.msgList.TextMatrix(ii, 1)
'                    Print #12, Me.msgList.TextMatrix(ii, 2)
                    'edit by nick 2004/10/12
                    'Print #12, Format(Me.msgList.TextMatrix(ii, 0), "00") & ":" & Me.msgList.TextMatrix(ii, 2) & ":" & Me.msgList.TextMatrix(ii, 1)
                    'Modified by Lydia 2022/03/15
                    'Print #12, Format(Me.msgList.TextMatrix(ii, 0), "00") & ":" & Me.msgList.TextMatrix(ii, 2) & ":" & Me.msgList.TextMatrix(ii, 1) & ":" & Me.msgList.TextMatrix(ii, 3)
                    strMid = strMid & Format(Me.msgList.TextMatrix(ii, 1), "00") & ":" & Me.msgList.TextMatrix(ii, 3) & ":" & Me.msgList.TextMatrix(ii, 2) & ":" & Me.msgList.TextMatrix(ii, 4) & vbCrLf
                    'Ed
                Next ii
                'Modified by Lydia 2022/03/15
                'Close #12
                'Screen.MousePointer = vbDefault
                'MsgBox "檔案 " & strFileName & ".doc 匯出成功!!!", vbExclamation + vbOKOnly
                Screen.MousePointer = vbDefault
                If strMid = "" Then
                    MsgBox "無資料可輸出!!!", vbExclamation + vbOKOnly
                Else
                    Call PUB_SaveTextAsUTF8(strExcelPath & strFileName & ".txt", strMid)
                    If Dir(strExcelPath & strFileName & ".txt") <> "" Then
                        MsgBox "檔案已產生於" & vbCrLf & _
                                      " [" & strExcelPath & strFileName & ".txt] "
                    End If
                End If
                'end 2022/03/15
            Else
                MsgBox "請輸入檔案名稱!!!", vbExclamation + vbOKOnly
            End If
      Else
         MsgBox "無資料可匯出!!!", vbExclamation + vbOKOnly
      End If
   'Add By Cheng 2003/06/25
   Case 8 '整批刪除
      frm20_1.Show
      frm20_1.Tag = Index
      frm20_1.Caption = "商品名稱查詢--整批刪除"
      Me.Hide
   End Select
End Sub

Private Sub DeleteData()
Dim ii As Integer
'edit by nick 2004/10/12
'   If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 0) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "") Then
'      If MsgBox("國際分類：" & Me.msgList.TextMatrix(Me.msgList.RowSel, 0) & vbcrlf & _
'               "商品中文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 1), 20) & "..." & vbcrlf & _
'               "商品英文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 2), 60) & "..." & vbcrlf & vbcrlf & _
'               "您是否要刪除此筆資料???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
'         strSQL = IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 0) = "", " TMN01 IS NULL ", " TMN01='" & Me.msgList.TextMatrix(Me.msgList.RowSel, 0) & "' ")
'         strSQL = strSQL & IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 1) = "", " AND TMN02 IS NULL ", " AND TMN02='" & ChgSQL(Me.msgList.TextMatrix(Me.msgList.RowSel, 1)) & "' ")
'         strSQL = strSQL & IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 2) = "", " AND TMN03 IS NULL ", " AND TMN03='" & ChgSQL(Me.msgList.TextMatrix(Me.msgList.RowSel, 2)) & "' ")
'         strSQL = "Delete From TrademarkMerchandiseName Where " & strSQL
    If Me.msgList.RowSel > 0 And (Me.msgList.TextMatrix(Me.msgList.RowSel, 1) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 2) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 3) <> "" Or Me.msgList.TextMatrix(Me.msgList.RowSel, 4) <> "") Then
      'Modified by Lydia 2022/03/15 改用可顯示 Unicode 的對話框
      'If MsgBox("國際分類：" & Me.msgList.TextMatrix(Me.msgList.RowSel, 0) & vbCrLf & _
               "商品中文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 1), 20) & "..." & vbCrLf & _
               "商品英文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 2), 60) & "..." & vbCrLf & _
               "商品日文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 3), 60) & "..." & vbCrLf & vbCrLf & _
               "您是否要刪除此筆資料???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
      strExc(0) = "國際分類：" & Me.msgList.TextMatrix(Me.msgList.RowSel, 1) & vbCrLf & _
               "商品中文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 2), 20) & "..." & vbCrLf & _
               "商品英文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 3), 60) & "..." & vbCrLf & _
               "商品日文：" & Left(Me.msgList.TextMatrix(Me.msgList.RowSel, 4), 60) & "..." & vbCrLf & vbCrLf & _
               "您是否要刪除此筆資料???"
      If UniMsgBox(strExc(0), vbYesNo + vbDefaultButton2) = vbYes Then
      'end 2022/03/15
         'Added by Lydia 2022/03/15
         If "" & Me.msgList.TextMatrix(Me.msgList.RowSel, 5) <> "" Then
             strSql = " tmn05=" & Me.msgList.TextMatrix(Me.msgList.RowSel, 5)
         Else
         'end 2022/03/15
            strSql = IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 1) = "", " TMN01 IS NULL ", " TMN01='" & Me.msgList.TextMatrix(Me.msgList.RowSel, 1) & "' ")
            strSql = strSql & IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 2) = "", " AND TMN02 IS NULL ", " AND TMN02='" & ChgSQL(Me.msgList.TextMatrix(Me.msgList.RowSel, 2)) & "' ")
            strSql = strSql & IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 3) = "", " AND TMN03 IS NULL ", " AND TMN03='" & ChgSQL(Me.msgList.TextMatrix(Me.msgList.RowSel, 3)) & "' ")
            strSql = strSql & IIf(Me.msgList.TextMatrix(Me.msgList.RowSel, 4) = "", " AND TMN04 IS NULL ", " AND TMN04='" & ChgSQL(Me.msgList.TextMatrix(Me.msgList.RowSel, 4)) & "' ")
         End If 'Added by Lydia 2022/03/15
         strSql = "Delete From TrademarkMerchandiseName Where " & strSql
         cnnConnection.Execute strSql, ii
         '若有刪除資料, 則重新顯示查詢畫面
         If ii <> 0 Then cmd_Click 6
      End If
   Else
      MsgBox "請先點選欲刪除的資料!!!", vbExclamation + vbOKOnly
   End If
End Sub

Public Sub QueryData()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Double
Dim arrTMN01
Dim strTMN01 As String

    On Error GoTo ErrorHandler
    Me.Enabled = False
    Me.msgList.Visible = False: DoEvents
    Screen.MousePointer = vbHourglass
    StrSQLa = ""
    If Me.Check1(0).Value = vbUnchecked And Me.Check1(1).Value = vbUnchecked And Me.Check1(3).Value = vbUnchecked Then
        MsgBox "請先勾選欲查詢的項目!!!", vbExclamation + vbOKOnly
        GoTo ErrorHandler
    End If
    If Me.Check1(0).Value = vbChecked Then
        If Me.Text1(0).Text = "" Then
            MsgBox "請輸入中文商品名稱!!!", vbExclamation + vbOKOnly
            Me.Enabled = True
            Me.Text1(0).SetFocus
            GoTo ErrorHandler
        End If
        StrSQLa = StrSQLa & " And TMN02 Like '%" & ChgSQL(Me.Text1(0).Text) & "%' "
    End If
    If Me.Check1(1).Value Then
        If Me.Text1(1).Text = "" Then
            MsgBox "請輸入英文商品名稱!!!", vbExclamation + vbOKOnly
            Me.Enabled = True
            Me.Text1(1).SetFocus
            GoTo ErrorHandler
        End If
        StrSQLa = StrSQLa & " And Upper(TMN03) Like '%" & ChgSQL(UCase(Me.Text1(1).Text)) & "%' "
     End If
     'add by nick 2004/10/12
    If Me.Check1(3).Value Then
        If Me.Text1(3).Text = "" Then
            MsgBox "請輸入日文商品名稱!!!", vbExclamation + vbOKOnly
            Me.Enabled = True
            Me.Text1(3).SetFocus
            GoTo ErrorHandler
        End If
        StrSQLa = StrSQLa & " And Upper(TMN04) like '%" & ChgSQL(UCase(Me.Text1(3).Text)) & "%' "
     End If
     If Me.Check1(2).Value = vbChecked Then
        If Me.Text1(2).Text <> "" Then
            'Modify By Cheng 2004/03/24
'            strSQLA = strSQLA & " And TMN01='" & ChgSQL(Me.Text1(2).Text) & "' "
            arrTMN01 = Split(Me.Text1(2).Text, ",")
            strTMN01 = ""
            For ii = LBound(arrTMN01) To UBound(arrTMN01)
                If arrTMN01(ii) = "" Then
                    strTMN01 = strTMN01 & " TMN01 Is Null Or "
                Else
                    strTMN01 = strTMN01 & " TMN01='" & ChgSQL(arrTMN01(ii)) & "' Or "
                End If
            Next ii
            strTMN01 = Left(strTMN01, Len(strTMN01) - 3)
            StrSQLa = StrSQLa & " And ( " & strTMN01 & " ) "
        Else
            StrSQLa = StrSQLa & " And TMN01 Is Null "
        End If
    End If
    'Modify By Cheng 2004/03/24
'    strSQLA = "Select * From TrademarkMerchandiseName Where TMN01=TMN01 " & strSQLA
    'Modified by Lydia 2022/03/15
    'StrSQLa = "Select * From TrademarkMerchandiseName Where '01'='01' " & StrSQLa
    StrSQLa = "Select '' as V,tmn01,tmn02,tmn03,tmn04,tmn05 From TrademarkMerchandiseName Where '01'='01' " & StrSQLa
    'End
'edit by nick 2004/10/12
'    If Me.Check1(0).Value = vbChecked Then
'        strSQLA = strSQLA & " Order By TMN01, TMN02, TMN03 "
'    ElseIf Me.Check1(1).Value = vbChecked Then
'        strSQLA = strSQLA & " Order By TMN01, TMN03, TMN02 "
'    Else
'        strSQLA = strSQLA & " Order By TMN01, TMN02, TMN03 "
'    End If
    If Me.Check1(0).Value = vbChecked Then
        StrSQLa = StrSQLa & " Order By TMN01, TMN02, TMN03,tmn04 "
    ElseIf Me.Check1(1).Value = vbChecked Then
        StrSQLa = StrSQLa & " Order By TMN01, TMN03, TMN02,tmn04 "
    ElseIf Me.Check1(2).Value = vbChecked Then
        StrSQLa = StrSQLa & " Order By TMN01, TMN02, TMN03,tmn04 "
    Else
        StrSQLa = StrSQLa & " Order By TMN01, tmn04,TMN02, TMN03 "
    End If
   SetGridTitle
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      For ii = 1 To rsA.RecordCount
'edit by nick 2004/10/12
'         Me.msgList.TextMatrix(ii, 0) = "" & rsA.Fields(0).Value
'         Me.msgList.TextMatrix(ii, 1) = "" & rsA.Fields(1).Value
'         Me.msgList.TextMatrix(ii, 2) = "" & rsA.Fields(2).Value
         'Modified by Lydia 2022/03/15
         'Me.msgList.TextMatrix(ii, 0) = "" & rsA.Fields("TMN01").Value
         'Me.msgList.TextMatrix(ii, 1) = "" & rsA.Fields("TMN02").Value
         'Me.msgList.TextMatrix(ii, 2) = "" & rsA.Fields("TMN03").Value
         'Me.msgList.TextMatrix(ii, 3) = "" & rsA.Fields("TMN04").Value
         For intI = 0 To 5
             Me.msgList.TextMatrix(ii, intI) = "" & rsA(intI)
         Next intI
         'end 2022/03/15
         Me.msgList.Rows = Me.msgList.Rows + 1
         Me.msgList.RowHeight(ii) = 648
         rsA.MoveNext
      Next ii
      Me.msgList.Rows = Me.msgList.Rows - 1
      Me.Label1.Caption = "查詢結果：符合條件的資料，共 " & Format(rsA.RecordCount, "#,##0") & " 筆!!!"
      'Added by Lydia 2022/03/10 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If Me.msgList.Rows > 1 Then
         Me.msgList.FixedRows = 1
      End If
   Else
      Me.Label1.Caption = "查詢結果：符合條件的資料，共 0 筆!!!"
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Screen.MousePointer = vbDefault
   Me.msgList.Visible = True: DoEvents
   Me.Enabled = True
   
   Exit Sub
ErrorHandler:
   If Err.Number <> 0 Then MsgBox "(" & Err.Number & ") " & Err.Description, vbExclamation + vbOKOnly
   Screen.MousePointer = vbDefault
   Me.msgList.Visible = True: DoEvents
   Me.Enabled = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.cmd(1).Enabled = IsUserHasRightOfFunction("frm20", strAdd, False)
   Me.cmd(2).Enabled = IsUserHasRightOfFunction("frm20", strAdd, False)
   Me.cmd(3).Enabled = IsUserHasRightOfFunction("frm20", strEdit, False)
   Me.cmd(4).Enabled = IsUserHasRightOfFunction("frm20", strDel, False)
   Me.cmd(5).Enabled = IsUserHasRightOfFunction("frm20", strFind, False)
   Me.cmd(6).Enabled = IsUserHasRightOfFunction("frm20", strFind, False)
    'Add By Cheng 2003/03/28
    '匯出檔案
   Me.cmd(7).Enabled = IsUserHasRightOfFunction("frm20", strPrint, False)
   '整批刪除
   Me.cmd(8).Enabled = IsUserHasRightOfFunction("frm20", strDel, False)
   SetGridTitle
End Sub

Private Sub SetGridTitle()
   Me.msgList.Clear
   Me.msgList.Rows = 2
   'add by nick 2004/10/12
   Me.msgList.Cols = 6   'Modified by Lydia 2022/03/15 4=>6 ;增加勾選項,tmn05=流水號
   Me.msgList.FixedRows = 1
   'Modified by Lydia 2022/03/15
   'Me.msgList.FixedCols = 1
   'cFixed = 2
   'Me.msgList.FixedCols = cFixed
   'end 2022/03/15
   Me.msgList.RowHeight(0) = 648
   'Modified by Lydia 2022/03/15 原本Code整合如下
   intI = 0
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "V"
   Me.msgList.ColWidth(intI) = 300
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
   intI = intI + 1
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "國際分類"
   Me.msgList.ColWidth(intI) = 500
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
   intI = intI + 1
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "中文商品名稱"
   Me.msgList.ColWidth(intI) = 2800
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
   intI = intI + 1
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "英文商品名稱"
   Me.msgList.ColWidth(intI) = 2800
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
   intI = intI + 1
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "日文商品名稱"
   Me.msgList.ColWidth(intI) = 2800
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
   intI = intI + 1
   Me.msgList.col = intI
   Me.msgList.TextMatrix(0, intI) = "流水號"
   If Pub_StrUserSt03 = "M51" Then
      Me.msgList.ColWidth(intI) = 1500
   Else
      Me.msgList.ColWidth(intI) = 0
   End If
   Me.msgList.ColAlignment(intI) = flexAlignLeftCenter
'   Me.msgList.TextMatrix(0, 0) = "國際分類"
'   Me.msgList.TextMatrix(0, 1) = "中文商品名稱"
'   Me.msgList.TextMatrix(0, 2) = "英文商品名稱"
'   'add by nick 2004/10/12
'   Me.msgList.TextMatrix(0, 3) = "日文商品名稱"
'   Me.msgList.TextMatrix(0, 4) = "流水號"  'Added by Lydia 2022/03/15
'
'   Me.msgList.ColWidth(0) = 500
'   'Me.msgList.ColWidth(1) = 3000
'   Me.msgList.ColWidth(1) = 2800
'   'Me.msgList.ColWidth(2) = 8000
'   Me.msgList.ColWidth(2) = 2800
'   'add by nick 2004/10/12
'   'Me.msgList.ColWidth(3) = 8000
'   Me.msgList.ColWidth(3) = 2800
'   Me.msgList.ColAlignment(0) = flexAlignLeftCenter
'   Me.msgList.ColAlignment(1) = flexAlignLeftCenter
'   Me.msgList.ColAlignment(2) = flexAlignLeftCenter
'   'add by nick 2004/10/12
'   Me.msgList.ColAlignment(3) = flexAlignLeftCenter
'   'Added by Lydia 2022/03/15
'   Me.msgList.ColAlignment(4) = flexAlignLeftCenter
'   If Pub_StrUserSt03 = "M51" Then
'       Me.msgList.ColWidth(4) = 1500
'   Else
'       Me.msgList.ColWidth(4) = 0
'   End If
'   'end 2022/03/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm20 = Nothing
End Sub

'Added by Lydai 2022/03/15
Private Sub msgList_Click()
Dim lngColor As Long
   With msgList
       If .MouseRow > 0 Then
          lngColor = .CellBackColor
          GridClick msgList, intLastRow, 0, 0, 0, "V", lngColor
       End If
   End With
End Sub

Private Sub msgList_DblClick()
   '執行快顯功能
   cmd_Click 5
End Sub

Private Sub msgList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
        If msgList.MouseRow > 0 And msgList.MouseCol > 0 And msgList.MouseRow <= msgList.Rows - 1 And msgList.MouseCol <= msgList.Cols - 1 Then
            msgList.row = msgList.MouseRow
            msgList.col = msgList.MouseCol
            mdiMain.CopyWord = msgList.Text
            PopupMenu mdiMain.mnuMouseR
        End If
   End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Index = 0 Then OpenIme
End Sub

Private Sub Text1_LostFocus(Index As Integer)
'edit by nickc 2007/06/06 切換輸入法改用API
If Index = 0 Then CloseIme
End Sub
