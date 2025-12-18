VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030301 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部商標處期限通知"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9396
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9396
   Begin VB.CommandButton cmdOK 
      Caption         =   "發催審Email"
      Height          =   400
      Index           =   4
      Left            =   8240
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   810
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "商標延展案件結案通知(&M)"
      Height          =   400
      Index           =   0
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   810
      Width           =   2280
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail(S)"
      Height          =   400
      Index           =   1
      Left            =   3870
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "事件說明(&H)"
      Height          =   400
      Left            =   2745
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   60
      Width           =   1110
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隱藏白色(&H)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   1380
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有未發文(&A)"
      Height          =   400
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   60
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm030301.frx":0000
      Left            =   6150
      List            =   "frm030301.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   510
      Width           =   3195
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4455
      Left            =   30
      TabIndex        =   2
      Top             =   1260
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7853
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限|法定期限|承辦期限|核稿期限|管制人|承辦人|事件　|本所案號　　　|案件性質|備註　　　　|案件名稱　　　"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   525
      Width           =   1710
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件性質前面有 @ 代表發過結案通知"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2400
      TabIndex        =   16
      Top             =   930
      Width           =   2940
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   13
      Top             =   5160
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4860
      TabIndex        =   5
      Top             =   570
      Width           =   1260
   End
End
Attribute VB_Name = "frm030301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblUserName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim bLvlX As Boolean, bLvl4 As Boolean, bLvl5 As Boolean
Dim stNumList1(1 To 5) As String
Dim StrToMail(7) As String
'Modified by Lydia 2025/05/07 改成變數
Private Const ExpNa01 = "'011','101','239'"   'Added by Lydia 2016/11/16 分區處理: 日本、美國、歐盟
Dim strExpNa01 As String 'Added by Lydia 2025/05/07 改成變數
Dim strTemplatePath As String ', strTempFolder As String 'Add By Sindy 2024/7/22

Private Sub cmdHelp_Click()
   frm030301_2.Show vbModal
End Sub

Private Sub cmdHide_Click()

   'Modified by Morgan 2021/9/3 隱藏後會無法恢復且再點選進度會變沒權限
   'SetColor cmdHide.Tag
   Dim ii As Integer, bHide As Boolean, intCnt As Integer
   
   If Left(cmdHide.Caption, 2) = "隱藏" Then
      bHide = True
      cmdHide.Caption = "顯示白色(&S)"
   Else
      bHide = False
      cmdHide.Caption = "隱藏白色(&H)"
   End If
   
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      intCnt = 0
      For ii = 1 To .Rows - 1
         .row = ii
         .col = 1
         If .CellBackColor = &HFFFFFF Then
            .col = 0
            .CellBackColor = &HFFFFFF
            If bHide = True Then
               .RowHeight(ii) = 0
               .TextMatrix(ii, 0) = "H"
            ElseIf .TextMatrix(ii, 0) = "H" Then
               .TextMatrix(ii, 0) = ""
               .RowHeight(ii) = -1
            End If
         End If
         If .RowHeight(ii) > 0 Then
            intCnt = intCnt + 1
         End If
      Next
      .Visible = True
      lblCnt.Caption = "共 " & intCnt & " 筆"
   End If
   End With
   'end 2021/9/3
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
Dim Str01 As String
Dim StrTmpCp01020304 As String, StrTmpCp09 As String, StrTmpNp22 As String
Dim j As Integer, tmpArr As Variant, TmpArrNp22 As Variant
On Error GoTo ErrorHandler
Dim Str02 As String, Str03 As String, Str04 As String, strContent As String 'Add By Sindy 2024/7/22
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Add By Sindy 2010/01/04
   Select Case cmdState
      Case 0 '外商(延展、第二期註冊費) 案件結案通知
         StrTmpCp01020304 = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 9
               '外商且為逾期案件
               If Left(Trim(grdDataList.TextMatrix(i, 9)), 1) < "A" Or Left(Trim(grdDataList.TextMatrix(i, 9)), 1) > "Z" Then
                  StrTag = Right(Trim(grdDataList.TextMatrix(i, 9)), Len(Trim(grdDataList.TextMatrix(i, 9))) - 1)
               Else
                  StrTag = Trim(grdDataList.TextMatrix(i, 9))
               End If
               Str01 = SystemNumber(StrTag, 1)
               If Str01 = "FCT" And Mid(Trim(grdDataList.TextMatrix(i, 9)), 1, 1) = "*" Then
                  If Trim(grdDataList.TextMatrix(i, 28)) <> "@" Then
                     grdDataList.col = 10
                     'Modified by Lydia 2023/05/16 已無第二期註冊費之案件; 拿掉Or Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "第二期註冊費"
                     If Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "延展" Or Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "續展" Then
'                        '阿蓮要求增加:僅法定期限<系統日的案件才可以按結案通知
'                        If ChangeTDateStringToTString(Trim(grdDataList.TextMatrix(i, 2))) < strSrvDate(2) Then
                           StrTmpCp01020304 = StrTmpCp01020304 & StrTag & vbCrLf
'                        Else
'                           MsgBox "含有法定期限大於或等於系統日的案件，請重新點選！", vbExclamation, "發生錯誤！"
'                           grdDataList.MousePointer = flexDefault
'                           Screen.MousePointer = vbDefault
'                           Me.Enabled = True
'                           Exit Sub
'                        End If
                     Else
                        'Modified by Lydia 2023/05/16 拿掉”或非外商第二期註冊費”
                        MsgBox "含有非外商延展案，請重新點選！", vbExclamation, "發生錯誤！"
                        grdDataList.MousePointer = flexDefault
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                     End If
                  Else
                     'Modified by Lydia 2023/05/16 拿掉”或第二期註冊費資料”
                     MsgBox "含有已通知結案延展，請重新點選！", vbExclamation, "發生錯誤！"
                     grdDataList.MousePointer = flexDefault
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               Else
                    MsgBox "含有非外商逾期案件，請重新點選！", vbExclamation, "發生錯誤！"
                    grdDataList.MousePointer = flexDefault
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                    Exit Sub
               End If
            End If
         Next i
          
         If StrTmpCp01020304 <> "" Then
            'Modified by Lydia 2023/05/16
            'If MsgBox(StrTmpCp01020304 & vbCrLf & vbCrLf & "確定結案？", vbYesNo, "警告！") = vbNo Then
            If MsgBox("已選取下列案件：" & vbCrLf & StrTmpCp01020304 & vbCrLf & "若案件有未結清之帳款會逐案通知，" & vbCrLf & "是否繼續？", vbInformation + vbYesNo, "商標延展案件結案通知") = vbNo Then
               grdDataList.MousePointer = flexDefault
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         End If
         
         StrTmpCp09 = ""
         StrTmpNp22 = ""
         StrTmpCp01020304 = ""
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 9
               If Left(Trim(grdDataList.TextMatrix(i, 9)), 1) < "A" Or Left(Trim(grdDataList.TextMatrix(i, 9)), 1) > "Z" Then
                  StrTag = Right(Trim(grdDataList.TextMatrix(i, 9)), Len(Trim(grdDataList.TextMatrix(i, 9))) - 1)
               Else
                  StrTag = Trim(grdDataList.TextMatrix(i, 9))
               End If
               Str01 = SystemNumber(StrTag, 1)
               If Str01 = "FCT" And Mid(Trim(grdDataList.TextMatrix(i, 9)), 1, 1) = "*" And Trim(grdDataList.TextMatrix(i, 28)) <> "@" Then
                  grdDataList.col = 10
                  'Modified by Lydia 2023/05/16 已無第二期註冊費之案件; 拿掉Or Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "第二期註冊費"
                  If Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "延展" Or Mid(grdDataList.Text, 2, Len(grdDataList.Text)) = "續展" Then
                     StrTmpCp09 = StrTmpCp09 & grdDataList.TextMatrix(i, 26) & ","
                     StrTmpNp22 = StrTmpNp22 & grdDataList.TextMatrix(i, 29) & ","
                     StrTmpCp01020304 = StrTmpCp01020304 & StrTag & vbCrLf
                     'Modify By Sindy 2016/7/20
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open "select np02,np03,np04,np05 from nextprogress where np01='" & grdDataList.TextMatrix(i, 26) & "' and np22=" & grdDataList.TextMatrix(i, 29) & " and np12 is null", cnnConnection, adOpenStatic, adLockReadOnly
                     If AdoRecordSet3.RecordCount = 1 Then
                        'Modified by Lydia 2023/05/16 grdDataList.TextMatrix(i, 29))=>grdDataList.TextMatrix(i, 29), i)
                        If UpdCloseOk(AdoRecordSet3.Fields("np02"), AdoRecordSet3.Fields("np03"), AdoRecordSet3.Fields("np04"), AdoRecordSet3.Fields("np05"), grdDataList.TextMatrix(i, 26), grdDataList.TextMatrix(i, 29), i) = False Then
                           grdDataList.MousePointer = flexDefault
                           Screen.MousePointer = vbDefault
                           Me.Enabled = True
                           Exit Sub
                        End If
                     End If
                     'Mark by Lydia 2023/05/16 改在UpdCloseOk
                     'grdDataList.RowHeight(i) = 0
                     ''2016/7/20 END
                     'grdDataList.col = 28
                     'grdDataList.Text = "@"
                     'grdDataList.col = 10
                     'grdDataList.Text = "@" & Trim(grdDataList.TextMatrix(i, 10))
                     'grdDataList.col = 0
                     'grdDataList.Text = ""
                     'end 2023/05/16
                     For j = 0 To grdDataList.Cols - 1
                        grdDataList.col = j
                        If grdDataList.CellBackColor = &HFFC0C0 Then
                           grdDataList.CellBackColor = grdDataList.BackColor '&H80000018
                        Else
                           grdDataList.CellBackColor = &H8080FF
                        End If
                     Next j
                  End If
               End If
            End If
         Next i
         'Modify By Sindy 2016/7/20 Mark
'         If StrTmpCp09 <> "" Then
'              '紀錄
'              TmpArr = Split(StrTmpCp09, ",")
'              TmpArrNp22 = Split(StrTmpNp22, ",")
'              For i = 0 To UBound(TmpArr)
'                 If CheckStr(TmpArr(i)) <> "" Then
'                    CheckOC3
'                    AdoRecordSet3.CursorLocation = adUseClient
'                    AdoRecordSet3.Open "select * from t102inform where ti01=to_number(to_char(sysdate, 'YYYYMMDD')) and ti02='" & TmpArr(i) & "' and ti04=" & TmpArrNp22(i), cnnConnection, adOpenStatic, adLockReadOnly
'                    If AdoRecordSet3.RecordCount = 0 Then
'                       cnnConnection.Execute "insert into t102inform (ti01,ti02,ti03,ti04) values (to_number(to_char(sysdate, 'YYYYMMDD')),'" & TmpArr(i) & "','" & strUserNum & "'," & TmpArrNp22(i) & ") "
'                    End If
'                 End If
'              Next i
'         End If

         grdDataList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
   End Select
   '2010/01/04 End
   
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         bolRefresh = False
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next
         
         '案號
         StrTag = grdDataList.TextMatrix(i, 9)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         
         'Added by Lydia 2024/01/12 增加對執行”基本資料”和”案件進度查詢”的限閱案件控制
         If cmdState = 1 Or cmdState = 2 Or cmdState = 3 Then
            'Modified by Lydia 2024/04/10 +pType = "1"
            If PUB_ChkCufaByCaseNo(strUserNum, Me.Name, Replace(StrTag, "-", ""), "1") = False Then
               Exit For
            End If
         End If
         'end 2024/01/12
         
         Me.Show
         
         Select Case cmdState
            Case 1 '發E-Mail
               If Len(Trim(StrTag)) <> 0 Then
                  StrToMail(1) = Trim(StrTag) '本所案號
                  If Not IsNull(StrTag) Then
                     Me.Enabled = False
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     StrToMail(2) = Trim(grdDataList.TextMatrix(i, 12)) '案件名稱
                     StrToMail(3) = Trim(grdDataList.TextMatrix(i, 26)) '總收文號
                     StrToMail(4) = Mid(grdDataList.TextMatrix(i, 10), 2, Len(grdDataList.TextMatrix(i, 10))) '案件性質
                     StrToMail(5) = Trim(grdDataList.TextMatrix(i, 1)) '本所期限
                     StrToMail(6) = Trim(grdDataList.TextMatrix(i, 2)) '法定期限
                     StrToMail(7) = Trim(grdDataList.TextMatrix(i, 3)) '承辦期限
                     Screen.MousePointer = vbHourglass
                     '期限
                     If Trim(grdDataList.TextMatrix(i, 14)) = "A" Then
                        frm030301_1.strLimitKind = "本所"
                     ElseIf Trim(grdDataList.TextMatrix(i, 14)) = "B" Then
                        frm030301_1.strLimitKind = "承辦"
                     ElseIf Trim(grdDataList.TextMatrix(i, 14)) = "H" Then
                        frm030301_1.strLimitKind = "法定"
                     ElseIf Trim(grdDataList.TextMatrix(i, 14)) = "D" Then
                        frm030301_1.strLimitKind = "本所"
                        If Left(StrTag, 3) = "CFT" Or Left(StrTag, 3) = "CFC" Or _
                           (Left(StrTag, 1) = "S" And Trim(grdDataList.TextMatrix(i, 27)) <> "000") Then
                           frm030301_1.strLimitKind = "法定"
                        End If
                        '智權人員
                        frm030301_1.StrMailNum2 = Trim(grdDataList.TextMatrix(i, 18))
                        frm030301_1.lbl1(1).Caption = Trim(grdDataList.TextMatrix(i, 7))
                        frm030301_1.strNP22 = Trim(grdDataList.TextMatrix(i, 29)) 'Add By Sindy 2015/4/9 +NP22
                     End If
'                     frm030301_1.txt1(1) = "                本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
'                                                          "                案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf + _
'                                                          "                總收文號：" + Left(Trim(StrToMail(3)) & "                              ", 30) + "案件性質：" + StrToMail(4) + vbCrLf + vbCrLf + _
'                                                          "                本所期限：" + Left(Trim(StrToMail(5)) & "                              ", 30) + "法定期限：" + StrToMail(6) + vbCrLf + vbCrLf + _
'                                                          "                承辦期限：" + Left(Trim(StrToMail(7)) & "                              ", 30)
                     frm030301_1.txt1(1) = "                本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
                                                          "                案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf + _
                                                          "                總收文號：" + StrToMail(3) + vbCrLf + vbCrLf + _
                                                          "                案件性質：" + StrToMail(4) + vbCrLf + vbCrLf + _
                                                          "                本所期限：" + StrToMail(5) + vbCrLf + vbCrLf + _
                                                          "                法定期限：" + StrToMail(6) + vbCrLf + vbCrLf + _
                                                          "                承辦期限：" + StrToMail(7)
                     frm030301_1.Show
                     frm030301_1.Tag = StrTag
                     frm030301_1.strCP09 = Trim(StrToMail(3)) '總收文號
                     frm030301_1.strEvents = Trim(grdDataList.TextMatrix(i, 8)) '事件
                     frm030301_1.StrMenu
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
               
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  'Modified by Lydia 2025/05/07 +TF案
                  Case "FCT", "CFT", "TF" '商標
                     Screen.MousePointer = vbHourglass
                     frm100101_4.Show
                     frm100101_4.Tag = StrTag
                     frm100101_4.StrMenu
                     Screen.MousePointer = vbDefault
                  Case "CFC"
                     Screen.MousePointer = vbHourglass
                     frm100101_A.Show
                     frm100101_A.Tag = StrTag
                     frm100101_A.StrMenu
                     Screen.MousePointer = vbDefault
                  Case "S"
                     Screen.MousePointer = vbHourglass
                     frm100101_B.Show
                     frm100101_B.Tag = StrTag
                     frm100101_B.StrMenu
                     Screen.MousePointer = vbDefault
                  'Added by Lydia 2025/05/07
                  Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
                     Screen.MousePointer = vbHourglass
                     frm100101_5.Show
                     frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_5.StrMenu
                     Screen.MousePointer = vbDefault
                  'end 2025/05/07
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
            
            'Add By Sindy 2024/7/22
            Case 4 '發催審Email
               Str02 = SystemNumber(StrTag, 2)
               Str03 = SystemNumber(StrTag, 3)
               Str04 = SystemNumber(StrTag, 4)
               'Modify By Sindy 2024/8/1
               strContent = "Dear Colleagues," & vbCrLf & vbCrLf & _
                            "We refer to your email as attached." & vbCrLf & vbCrLf & _
                            "As our client is very concerned about the development of the subject mark, we would appreciate it if you could advise us of the relevant current status so that we can report the same to the client." & vbCrLf & vbCrLf & _
                            "We look forward to hearing from you soon."
               Call PUB_SettingF11eMail(strTemplatePath, Str01, Str02, Str03, Str04, strContent)
            '2024/7/22 END
         End Select
         Exit For
      End If
   Next i
   
   If bolRefresh = True Then
      cmdQuery_Click 0
   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub
'Modified by Lydia 2023/05/16 +iRow as Integer
Private Function UpdCloseOk(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
                            strNP01 As String, strNP22 As String, iRow As Integer) As Boolean
Dim intP As Integer
Dim strAutoNum As String
Dim SCp(1 To 79) As String
Dim strSql As String
Dim bolConn As Boolean, intAns As Integer 'Added by Lydia 2023/05/16

   On Error GoTo CheckingErr
   
   UpdCloseOk = False
   
   'Added by Lydia 2023/05/16 增刪商標延展案件結案通知功能：逐筆詢問選擇「是」結案閉卷：原本的處理，選擇「否」更新並管制下次延展期限，選擇「取消」放棄該案的處理
   strExc(0) = "select tm22,tm23,na14 from trademark, nation where tm01='" & strCP01 & "' and tm02='" & strCP02 & "' and tm03='" & strCP03 & "' and tm04='" & strCP04 & "' and tm10=na01(+) "
   intI = 1
   strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = "" & RsTemp.Fields("tm22") '專用期間止日
      strExc(2) = "" & RsTemp.Fields("tm23") '申請人1
      strExc(3) = "" & RsTemp.Fields("na14") '延展年度
   End If
   If Val(strExc(1)) = 0 Or Val(strExc(3)) = 0 Then
       MsgBox strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & "無" & Mid(IIf(Val(strExc(1)) = 0, "、專用期間止日", "") & IIf(Val(strExc(3)) = 0, "、延展年度", ""), 2) & "！"
       UpdCloseOk = True
       Exit Function
   Else
      intAns = MsgBox(strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 <> "000", "-" & strCP03 & "-" & strCP04, "") & _
               vbCrLf & "選擇「是」將結案閉卷；" & vbCrLf & "選擇「否」將更新並管制下次延展期限；" & vbCrLf & _
               "選擇「取消」該案不處理。", vbInformation + vbYesNoCancel + vbDefaultButton2, "商標延展案件結案通知")
      If intAns = 2 Then  'Cancel 取消
         UpdCloseOk = True
         Exit Function
      End If
   End If
   'end 2023/05/16
   Screen.MousePointer = vbHourglass
   bolConn = True 'Added by Lydia 2023/05/16
   cnnConnection.BeginTrans
      'Added by Lydia 2023/05/16 選擇「否」更新並管制下次延展期限
      If intAns = 7 Then  'vbNo
         'Added by Lydia 2025/11/12
         Dim strNewTM22 As String
         strNewTM22 = CompDate(0, Val(strExc(3)), strExc(1))
         'end 2025/11/12
         
         '將原期限更新NP11=系統日,NP12=19,NP06=N
         strSql = "Update NextProgress set np11=to_char(sysdate,'yyyymmdd'), np12='19', np06='N' Where np01='" & strNP01 & "' and np22=" & strNP22
         cnnConnection.Execute strSql
         '將基本檔專用期止日TM22增加10年(NA14)
         'Modified by Lydia 2025/11/12 CompDate(0, Val(strExc(3)), strExc(1))改用變數strNewTM22
         strSql = "Update TradeMark set TM22=" & strNewTM22 & " Where TM01='" & strCP01 & "' and TM02='" & strCP02 & "' and TM03='" & strCP03 & "' and TM04='" & strCP04 & "' "
         cnnConnection.Execute strSql
         '新增下一程序之延展期限
         strExc(0) = GetNextProgressNo()
         'Modified by Lydia 2025/11/12 CompDate(0, Val(strExc(3)), strExc(1))改用變數strNewTM22; 所限改抓最近工作天
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                 "VALUES ('" & strNP01 & "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "','102'," & PUB_GetWorkDay1(strNewTM22, True) & "," & strNewTM22 & _
                 ",'" & PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04) & "'," & strExc(0) & ")"
         cnnConnection.Execute strSql
      Else '選擇「是」結案閉卷：原本的處理
      'end 2023/05/16
         'UPDATE 是否續辦為 N 和解除期限日期和解除期限原因
         '是不算案件數(CP26)設為"N"
         strSql = "UPDATE NEXTPROGRESS SET NP11=" & strSrvDate(1) & ",NP12='52' " & ",NP06='N' " & _
                  " WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22
         cnnConnection.Execute strSql
         strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & strSrvDate(1) & ",TM31='52' WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "' "
         cnnConnection.Execute strSql
         
         If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
            CheckOC
            strSql = "select au01||(au02-1911) from autonumber where au01='B'"
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If Not adoRecordset.BOF Then adoRecordset.MoveFirst
            If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Function
            '比對自動編號年度
            strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNum
            CheckOC
            strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20,cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30,cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50,cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60,cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76,cp77,cp78,cp79) values "
            For intP = 1 To 79
               Select Case intP
               Case 8, 11, 21, 22, 23, 24, 28, 29, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63, 64
                    SCp(intP) = "null "
               Case 14 '承辦人代號
                    SCp(intP) = "'" & strUserNum & "'"
               Case 1
                    SCp(intP) = "'" & strCP01 & "'"
               Case 2
                    SCp(intP) = "'" & strCP02 & "'"
               Case 3
                    SCp(intP) = "'" & strCP03 & "'"
               Case 4
                    SCp(intP) = "'" & strCP04 & "'"
               Case 12
                    SCp(intP) = "'" & GetSalesArea(PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)) & "'"
               Case 13
                    SCp(intP) = "'" & PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04) & "'"
               Case 5, 27
                    SCp(intP) = GetTodayDate
               Case 9
                    SCp(intP) = "'" & strAutoNum & "'"
               Case 20
                     SCp(intP) = "'N'"
               Case 26, 32
                    SCp(intP) = "'N'"
               Case 43
                    SCp(intP) = "'" & strNP01 & "'"
               Case 30
                    SCp(intP) = "'" & strNP22 & "'"
               Case 10
                     SCp(intP) = "'704'"
               Case 65, 66, 67, 68, 69, 70
                    SCp(intP) = ""
               Case 57
                    SCp(intP) = GetTodayDate
               Case 58
                    SCp(intP) = "'52'"  '延展未收到指示，逾期結案通知系統閉卷
               '數字
               Case Else
                    SCp(intP) = "null "
               End Select
            Next intP
            strSql = strSql & " ("
            For intP = 1 To 79
                Select Case intP
                Case 65, 66, 67, 68, 69, 70
                Case Else
                     strSql = strSql & SCp(intP)
                     If intP <> 79 Then
                        strSql = strSql & ","
                     End If
                End Select
            Next intP
            strSql = strSql & ") "
            cnnConnection.Execute strSql
         Else
            MsgBox ("延展閉卷自動給號錯誤!")
            GoTo CheckingErr
         End If
      End If '-----'Added by Lydia 2023/05/16 選擇「否」更新並管制下次延展期限
   cnnConnection.CommitTrans
   'Added by Lydia 2023/05/16 增刪商標延展案件結案通知功能：有應收帳款->發email通知
   bolConn = False
   grdDataList.RowHeight(iRow) = 0
   grdDataList.TextMatrix(iRow, 28) = "@"
   grdDataList.TextMatrix(iRow, 10) = "@" & grdDataList.TextMatrix(iRow, 10)
   grdDataList.TextMatrix(iRow, 0) = ""
   If PUB_GetBillDataAll("1", strExc(2), strCP01, "102", txtUsernum, , , , , strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04) = True Then
      strExc(4) = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04, "102")
      strExc(5) = ""
      strExc(0) = "select st52,st53 from staff where st01='" & strExc(4) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("st52") <> "" Then strExc(5) = strExc(5) & ";" & RsTemp.Fields("st52")
         If "" & RsTemp.Fields("st53") <> "" Then strExc(5) = strExc(5) & ";" & RsTemp.Fields("st53")
      End If
      If InStr(strExc(4) & ";" & strExc(5), strUserNum) = 0 Then
         strExc(4) = strExc(4) & ";" & strUserNum '操作者非承辦或主管
      Else
         strExc(4) = ""
      End If
      If strExc(5) <> "" Then
         strExc(5) = Mid(strExc(5), 2)
         PUB_SendMail strUserNum, strExc(5), strNP01, "未結清帳款案件通知(" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 <> "000", "-" & strCP03 & "-" & strCP04, "") & "延展結案)", "本案尚有未結清之帳款，請確認無誤後，發函向客戶催款。", , , , , , strExc(4)
      End If
   End If
   'end 2023/05/16
   UpdCloseOk = True
   
   Screen.MousePointer = vbDefault
   Exit Function
      
CheckingErr:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans 'Modified by Lydia 2023/05/16
   MsgBox Err.Description
End Function
'2016/7/20 END

Public Sub cmdQuery_Click(Index As Integer)

   Screen.MousePointer = vbHourglass
   If pub_bolInformCheck = False Then
      If MsgBox("是否確定要查詢？", vbYesNo + vbDefaultButton2) = vbNo Then
         GoTo SubOut
      End If
   End If
   
   Me.Enabled = False
   doQuery Index
   Me.Enabled = True
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   'Modify By Sindy 2025/9/3 從 Form_Load 移來:有使用者不關閉此作業,記錄不到Log; 所以改在有啟動查詢時就記錄Log
   PUB_AddExcuteLog Me.Name
   
SubOut:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String
   
   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery(idx As Integer)
Dim stVTB As String
Dim stDate0 As String, stDate1 As String, stDate2 As String, stDate_3 As String
Dim stDate4 As String
Dim stDate5 As String, stDate6 As String, stDate7 As String, stDate_10 As String
Dim stNumList As String, stDept As String, stDeptST03 As String
Dim ii As Integer, stIdList
Dim stUserID As String
Dim strOtherUser As String
Dim txtData As Variant, strWhSql As String, strUser As String
Dim strF4103TSql As String, strF4103SSql As String, strF4103LSql As String
Dim iRow As Long 'Add By Sindy 2014/9/17
Dim rsTmp As New ADODB.Recordset
Dim stConCP142 As String 'Add By Sindy 2023/12/11
Dim stDate102105_B As String, stDate102105_E As String  'Added by Lydia 2025/05/06 延展(102)/使用宣誓(105)未收文期限
Dim stDate_B1 As String, strExCond As String 'Added by Lydia 2025/05/06

   stVTB = ""
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   stUserID = txtUsernum
   '使用者收文智權人員所屬部門
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
      stDeptST03 = Pub_StrUserSt03
   Else
      stDept = GetST15(stUserID)
      stDeptST03 = GetStaffDepartment(stUserID)
   End If
   
   '抓員工外譯對照資料
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   '期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stDate0 = strSrvDate(1) - 10000 '系統日-1年
   stDate1 = CompWorkDay(3, strSrvDate(1))
   stDate2 = CompWorkDay(4, strSrvDate(1))
   stDate4 = CompWorkDay(6, strSrvDate(1)) '+5工作天，用在<表示4個工作天內
   stDate_3 = CompWorkDay(4, strSrvDate(1), 1) '減3個工作天
   stDate5 = CompWorkDay(7, strSrvDate(1))
   stDate6 = CompWorkDay(8, strSrvDate(1))
   stDate7 = CompWorkDay(9, strSrvDate(1))
   stDate_10 = CompWorkDay(11, strSrvDate(1), 1) '減10個工作天
   stConCP142 = " AND CP142>=" & stDate0 & " AND CP142< " & stDate2 'Add By Sindy 2023/12/11 指定日期＜＝系統日＋２個工作天之未發文案件。
   'Added by Lydia 2025/05/06 延展(102)/使用宣誓(105)：法定期限＜＝系統日+22個工作天之未收文案件，只跳2工作天提醒
   stDate102105_B = CompWorkDay(21, strSrvDate(1))
   stDate102105_E = CompWorkDay(23, strSrvDate(1))
   stDate_B1 = CompWorkDay(2, strSrvDate(1), 1) '減1個工作天
   'Added by Lydia 2025/05/07 特殊管制國家另外用「CFT特殊管制國家」控制
   strExpNa01 = "'AAA'"
   strSql = "select na01 from setspecman,nation where ocode='CFT特殊管制國家' and instr(oman,na01) > 0 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         'CFT_101239已被分成單一國家設定
         'Modified by Lydia 2025/07/21 debug: 判斷主管底下的人員有設定「特殊管制國家」
         'strExc(0) = "select ocode from setspecman where ocode not like '%101239%' and ocode like 'CFT_" & RsTemp.Fields("na01") & "%' and instr(oman,'" & stUserID & "') > 0 "
         txtData = Empty
         txtData = Split(stNumList, ",")
         For ii = 0 To UBound(txtData)
            If Trim(txtData(ii)) <> "" Then
               strExc(0) = "select ocode from setspecman where ocode not like '%101239%' and ocode like 'CFT_" & RsTemp.Fields("na01") & "%' and instr(oman," & Trim(txtData(ii)) & ") > 0 "
         'end 2025/07/21
               If rsTmp.State <> 0 Then rsTmp.Close
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  If strExpNa01 = "'AAA'" Then strExpNa01 = ""
                  strExpNa01 = strExpNa01 & IIf(strExpNa01 <> "", ",", "") & "'" & RsTemp.Fields("na01") & "'"
                  strExCond = strExCond & rsTmp.GetString(adClipString, , , ",")
               End If
         'Added by Lydia 2025/07/21
            End If
         Next ii
         'end 2025/07/21
         RsTemp.MoveNext
      Loop
      If strExCond <> "" Then strExCond = Mid(strExCond, 1, Len(strExCond) - 1)
   End If
   'end 2025/05/06
   
   '特殊權限
   bLvlX = CheckLevel(stUserID, "M") '未交稿,已完稿無核稿管制人
   bLvl4 = CheckLevel(stUserID, "V") '第四級管制人
   'Modified by Lydia 2018/11/13 改設定
   'bLvl5 = CheckLevel(stUserID, "O") '第五級管制人
   bLvl5 = CheckLevel(stUserID, "V1")
   
   'Modify By Sindy 2015/2/9
   '清除暫存檔
   strSql = "delete R030301 where ID='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2010/7/29 另外依F4103期限做區分
   '日本地區-葉易雲主任 78011
   '其他地區-洪琬姿副理 80030
   strF4103TSql = ""
   strF4103SSql = ""
   strF4103LSql = "" 'Add By Sindy 2015/10/23
   If txtUsernum = "78011" Then
      strF4103TSql = " AND ((TM44 is not null AND exists (select * from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9) and substr(fa10,1,3)='011'))" & _
                       " or (TM44 is null AND exists (select * from customer where cu01=substr(tm23,1,8) and cu02=substr(tm23,9) and substr(cu10,1,3)='011'))) "
      'Add By Sindy 2015/10/23
      strF4103LSql = " AND ((LC22 is not null AND exists (select * from fagent where fa01=substr(LC22,1,8) and fa02=substr(LC22,9) and substr(fa10,1,3)='011'))" & _
                       " or (LC22 is null AND exists (select * from customer where cu01=substr(LC11,1,8) and cu02=substr(LC11,9) and substr(cu10,1,3)='011'))) "
      '2015/10/23 END
      strF4103SSql = " AND ((sp26 is not null AND exists (select * from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9) and substr(fa10,1,3)='011'))" & _
                       " or (sp26 is null AND exists (select * from customer where cu01=substr(sp08,1,8) and cu02=substr(sp08,9) and substr(cu10,1,3)='011'))) "
   ElseIf txtUsernum = "80030" Then
      strF4103TSql = " AND ((TM44 is not null AND exists (select * from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (TM44 is null AND exists (select * from customer where cu01=substr(tm23,1,8) and cu02=substr(tm23,9) and substr(cu10,1,3)<>'011'))) "
      'Add By Sindy 2015/10/23
      strF4103LSql = " AND ((LC22 is not null AND exists (select * from fagent where fa01=substr(LC22,1,8) and fa02=substr(LC22,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (LC22 is null AND exists (select * from customer where cu01=substr(LC11,1,8) and cu02=substr(LC11,9) and substr(cu10,1,3)<>'011'))) "
      '2015/10/23 END
      strF4103SSql = " AND ((sp26 is not null AND exists (select * from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (sp26 is null AND exists (select * from customer where cu01=substr(sp08,1,8) and cu02=substr(sp08,9) and substr(cu10,1,3)<>'011'))) "
   End If
   
   '代碼1:A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=達法定
   '      I=達指會,J=今送件,K=需請款,N=達指定
   '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   '【FCT、T、S台灣案】
   '承辦組
   If Trim(stDept) = "F11" Or Trim(stDept) = "F10" Or bLvl4 = True Or bLvl5 = True Then
'***********************
''F' EV1,'3' EV2
'***********************
         'Add By Sindy 2010/5/25
         '未請款,(承辦人為非F1*非外商人員)：以發文日次日起算10個工作天
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, staff" & _
               " WHERE CP01 in ('FCT','T','S') AND CP05>=20100101" & _
               " AND CP27<" & stDate_10 & _
               " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in ('FCT','T') AND CP05>=20100101" & _
                     " AND CP27<" & stDate_10 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S' AND CP05>=20100101" & _
                     " AND CP27<" & stDate_10 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2010/5/25
         '未請款,(承辦人為F1*外商人員)：以發文日次日起算3個工作天
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件; CP01 in ('FCT','T','S')=>CP01 in ('FCT','T','S'" & IIf(Left(stDept, 2) = "F1", "'CFT','CFC'", "") & ")
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, staff" & _
               " WHERE CP01 in ('FCT','T','S'" & IIf(Left(stDept, 2) = "F1", ",'CFT','CFC'", "") & ")  AND CP05>=20100101" & _
               " AND CP27<" & stDate_3 & _
               " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in ('FCT','T') AND CP05>=20100101" & _
                     " AND CP27<" & stDate_3 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S' AND CP05>=20100101" & _
                     " AND CP27<" & stDate_3 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2015/10/23 +法務:未請款
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS" & _
               " WHERE CP01 in ('FCL','CFL','LIN','ACS') AND CP05>=20100101" & _
               " AND CP27<" & stDate_3 & _
               " AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
'***********************
''A' EV1,'3' EV2
'***********************
         '已收文未發文,(承辦人為非F1*非外商人員)2個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => NVL(CP27,'0')='0' AND NVL(CP57,'0')='0' AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP13051027) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP13051027) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND NVL(CP27,'0')='0' AND NVL(CP57,'0')='0'" & _
               " AND NVL(CP14,'0') > '0' AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND NVL(CP14,'0') > '0' AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         '已收文未發文,(承辦人為F1*外商人員)1個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件; CP01 in ('FCT','T') =>CP01 in ('FCT','T'" & IIf(Left(stDept, 2) = "F1", ",'CFT'", "") & ")
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in ('FCT','T'" & IIf(Left(stDept, 2) = "F1", ",'CFT'", "") & ")" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
        ' strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01 in ('S'" & IIf(Left(stDept, 2) = "F1", ",'CFC'", "") & ")" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:達本所
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         
'***********************
''A' EV1,'0' EV2
'***********************
         '已收文未發文,[未分案]2個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:達本所
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         
'***********************
''B' EV1,'3' EV2
'***********************
         '已收文未發文之722外商發文,2個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
'2010/4/19 modify by sonia 發現FCT-027964的核駁前先行通知逾承辦未出現,
                          '故不限制722外商發文,只要承辦人為F1部門者都出現
         'Modify By Sindy 2015/8/24 剔除901.催款
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
               " AND CP10 Not in('720','901') AND CP01||CP10<>'FCT1201'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2015/8/24 剔除901.催款
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                     " AND CP10 Not in('720','901') AND CP01||CP10<>'FCT1201'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
'2010/4/19 END
         '已收文未發文之720回覆代理人,1個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
         'Modify By Sindy 2015/8/24 增加901.催款期限控管
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND TM29 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 in('720','901')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/11/19 +法務:達承辦
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND LC08 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 ='901'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/11/19 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND SP15 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 in('720','901')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2015/8/24 增加901.催款期限控管
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND TM29 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 in('720','901')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/11/19 +法務:達承辦
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND LC08 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 ='901'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/11/19 END
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND SP15 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 in('720','901')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2015/7/30 已收文未發文之FCT.1201審查報告,4個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
         'Modified by Lydia 2016/10/14
         'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01||CP10='FCT1201'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP57 is null AND CP27 is null" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01='FCT' AND CP10='1201' AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01||CP10='FCT1201'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/7/30 END
                           
'***********************
''D' EV1,'3' EV2
'***********************
         '未收文且 7個工作天 後達本所期限者(不含當日) --智權人員-D3(未收文,智權人員)
         'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
         'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
         'Modify By Sindy 2015/7/30 FCT的208補優先權證明期限要剔除,獨立控管 + and NP02||NP07<>'FCT208'
         'Modify By Sindy 2015/8/21 FCT的901催款要剔除,獨立控管
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02 in ('FCT','T') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                     " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
         If Left(stDept, 2) = "F1" Then
            'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                      " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                         "From NEXTPROGRESS, TradeMark, Caseprogress" & _
                         " WHERE NP02 in ('CFT') AND NP06 is null" & _
                         " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                         " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                         " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                         " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                         " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                         " AND TM29 IS NULL AND NP01=CP09(+) AND SUBSTR(CP12,1,2)='F1'" & _
                         " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
             cnnConnection.Execute strSql, intI
         End If
         'end 2022/11/21
         'Add By Sindy 2015/10/23 +法務:未收文
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Lawcase, Staff" & _
                     " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('901') AND NP10=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, SERVICEPRACTICE" & _
                     " WHERE NP02='S' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
         If Left(stDept, 2) = "F1" Then
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                        "From NEXTPROGRESS, SERVICEPRACTICE,Caseprogress " & _
                        " WHERE NP02 IN ('S','CFC') AND NP06 is null" & _
                        " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                        " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901')" & _
                        " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                        " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                        " AND SP15 IS NULL AND NP01=CP09(+) AND SUBSTR(CP12,1,2)='F1'" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
            cnnConnection.Execute strSql, intI
         End If
         'end 2022/11/21

         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
               'Modify By Sindy 2015/7/30 FCT的208補優先權證明期限要剔除,獨立控管 + and NP02||NP07<>'FCT208'
               'Modify By Sindy 2015/8/21 FCT的901催款要剔除,獨立控管
               'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02 in('FCT','T') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103TSql & " AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                           " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:未收文
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, Lawcase" & _
                           " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103LSql & " AND np07 NOT IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                           " AND LC08 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, SERVICEPRACTICE" & _
                           " WHERE NP02='S' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103SSql & " AND np07 NOT IN ('305','901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                           " AND SP15 IS NULL" & _
                           " AND SP09='000'" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         
         'Add By Sindy 2015/8/21 FCT的901催款=>改控制為本所期限<=系統日+2個工作天
         'T:外商收的大陸案
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02 in ('FCT','T') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:未收文
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Lawcase, Staff" & _
                     " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901') AND NP10=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, SERVICEPRACTICE" & _
                     " WHERE NP02='S' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02 in('FCT','T') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103TSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:未收文
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, Lawcase" & _
                           " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103LSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                           " AND LC08 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, SERVICEPRACTICE" & _
                           " WHERE NP02='S' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103SSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                           " AND SP15 IS NULL" & _
                           " AND SP09='000'" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/8/21 END
         
         'Add By Sindy 2015/7/30 FCT的208補優先權證明期限=>改控制為本所期限<=系統日+30日曆天
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02||NP07='FCT208' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08<=" & DBDATE(DateAdd("d", 30, ChangeWStringToWDateString(strSrvDate(1)))) & _
                     " AND NP10 in (" & stNumList & ")" & _
                     " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02||NP07='FCT208' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08<=" & DBDATE(DateAdd("d", 30, ChangeWStringToWDateString(strSrvDate(1)))) & _
                           " AND NP10 ='F4103'" & strF4103TSql & _
                           " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/7/30 END
   End If
   
'***********************
''A' EV1,'1' EV2
'***********************
   '【FCT、S台灣案】
   '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01='FCT'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2015/10/23 +法務:達本所
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2015/10/23 END
   
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2010/7/29 另外依F4103期限做區分
   If txtUsernum = "78011" Or txtUsernum = "80030" Then
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01='FCT'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103TSql & " AND CP158=0 AND CP159=0" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103LSql & " AND CP158=0 AND CP159=0" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103SSql & " AND CP158=0 AND CP159=0" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
   End If
   
'***********************
''K' EV1,'1' EV2
'***********************
   'Add By Sindy 2021/6/3
   '需請款：S案之「查名」進度有承辦期限時，承辦期限當日提醒承辦人，事件為「需請款」，顏色為紫色
   'Modify By Sindy 2021/7/1:6/4上線有關查名案期限通知之需求，請改為以本所期限管控，即：
   'S案「查名」進度之本所期限當日提醒承辦人，事件為「需請款」，顏色為紫色
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','K' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S' AND CP10='001' AND CP16>0 AND CP60 IS NULL" & _
               " AND ((CASEPROGRESS.CP06 is not null AND CASEPROGRESS.CP06<=" & strSrvDate(1) & ") or (CASEPROGRESS.CP07 is not null AND CASEPROGRESS.CP07<=" & strSrvDate(1) & "))" & _
               " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='K' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2023/3/24 + FCT案的001.查名
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','K' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Trademark" & _
               " WHERE CP01='FCT' AND CP10='001' AND CP16>0 AND CP60 IS NULL" & _
               " AND ((CASEPROGRESS.CP06 is not null AND CASEPROGRESS.CP06<=" & strSrvDate(1) & ") or (CASEPROGRESS.CP07 is not null AND CASEPROGRESS.CP07<=" & strSrvDate(1) & "))" & _
               " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM10='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='K' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2021/6/3 END
   
'***********************
''D' EV1,'2' EV2
'***********************
   '程序組
   If Trim(stDept) = "F12" Then
         '未收文且 1個工作天 後達本所期限者(不含當日) --管制人-D2(未收文,管制人)
         'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
         'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, TradeMark" & _
                     " WHERE NP02||NP06='FCT'" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND np07 NOT IN ('305')" & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and decode(np02||np07,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:未收文
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, Lawcase" & _
                     " WHERE NP02||NP06 in('FCL','CFL','LIN','ACS')" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, SERVICEPRACTICE" & _
                     " WHERE NP02||NP06='S'" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND np07 NOT IN ('305')" & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         
'***********************
''J' EV1,'1' EV2
'***********************
         'Add By Sindy 2014/12/11 將FCT分案時輸入之承辦期限納入外商程序組之期限自動通知系統:J.今送件
         'Modified by Lydia 2016/10/14 CP01||CP27||CP57='FCT' => CP01='FCT' AND CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','J' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01='FCT' AND CP158=0 AND CP159=0" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48<=" & strSrvDate(1) & _
                     " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='J' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP01||CP27||CP57='S' => CP01='S' AND CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','J' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S' AND CP158=0 AND CP159=0" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48<=" & strSrvDate(1) & _
                     " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='J' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
   End If
   
'***********************
''H' EV1,'1' EV2
'***********************
   '【CFT、CFC、S非台灣案】
   '已收文未發文,5個工作天後達法定期限者(不含當日) --承辦人-H1(法定期限,承辦人)
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','H' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & _
               " AND CASEPROGRESS.CP07>=" & stDate0 & " AND CASEPROGRESS.CP07< " & stDate5 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','H' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP07>=" & stDate0 & " AND CASEPROGRESS.CP07< " & stDate5 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2023/12/11 +達指定
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & stConCP142 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29||TM57 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & stConCP142 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '達指定:未分案,管制人
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'2' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,NA69 NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, Nation" & _
               " WHERE CP01='CFT'" & stConCP142 & _
               " AND CASEPROGRESS.CP14 is null AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='2' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'2' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,NA69 NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE CP01 in('S','CFC')" & stConCP142 & _
               " AND CASEPROGRESS.CP14 is null AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='2' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2023/12/11 END
   
   'Add By Sindy 2021/7/5
   '已收文未發文,1個工作天後達本所期限者(不含當日) --承辦人-H1(法定期限,承辦人)
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2021/7/5 END
      
   'Added By Lydia 2022/11/21
   '已收文未發文,承辦期限＜＝系統日＋１個工作天之未發文案件。 --達承辦 B1 (承辦期限,承辦人)
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','B' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000'))" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','B' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'end 2022/11/21
   
'***********************
''D' EV1,'2' EV2
'***********************
   '未收文且 第5、6個工作天後達法定期限者(不含當日) --管制人-D2(未收文,管制人)
   'Modify By Sindy 2010/3/1 申請國家非日本時，管制人為國家檔CFT承辦人
   'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
   'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
   'Modified by Lydia 2016/03/11 國家檔的CFT承辦人(NA69)改成模組(DB.Functions)
   'Remove by Lydia 2016/03/24 外商反應未考慮好,先移除
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            "select * from ( SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,GETNA69(NP02,NP03,NP04,NP05,'','') NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 NOT IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10<>'011' AND TM10=NA01(+) AND TM10=NA01(+)" & _
               " and decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')='Y'" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)" & _
               " ) where na16 in (" & stNumList & ") "
   'Modified by Lydia 2016/11/15  TM10<>'011' => TM10 NOT IN (" & strExpNa01 & ")
   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
   'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
   'Modified by Lydia 2025/05/06 延展(102)：從系統日＋５或６個工作天前只彈跳2天提醒，改成從系統日＋６個工作天起持續跳至專用期限滿1個工作天後再解除彈跳。(原) CompWorkDay(6, strSrvDate(1)) => stDate_B1
   'Modified by Lydia 2025/05/07 合併「CFT特殊管制國家」的語法; (原) AND TM10 NOT IN (" & strExpNa01 & ") AND NA69 in (" & stNumList & ")" => AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))
                                 '(原),NA69 NA16, => ," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
            "From NEXTPROGRESS, TradeMark, Nation" & _
            " WHERE NP02||NP06='CFT'" & _
            " AND NP09>=" & stDate_B1 & " AND NP09< " & stDate6 & _
            " AND np07 IN ('102')" & _
            " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
            " AND TM29||TM57 IS NULL" & _
            " AND TM10=NA01(+) AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))" & _
            " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
            " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
            " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/10/14 NP02||NP06 in ('S','CFC') => NP02 in ('S','CFC') AND NVL(NP06,'0')='0' ; + /*+ INDEX(NEXTPROGRESS IDXNP09020706) */
   'Modified by Lydia 2016/11/15  SP09<>'011' => SP09 NOT IN (" & strExpNa01 & ")
   'Modified by Lydia 2018/12/04 CFC不抓NA69
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(NEXTPROGRESS IDXNP09020706) */ '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 NOT IN ('305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & strExpNa01 & ") AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
  'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
               ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=> SP15||SP61 IS NULL
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
   'Mark by Lydia 2025/05/06 S,CFC沒有102延展
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(NEXTPROGRESS IDXNP09020706) */ '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 IN ('102')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   'cnnConnection.Execute strSql, intI
   '-----'Mark by Lydia 2025/05/06 S,CFC沒有102延展

   'Add By Sindy 2014/9/12 催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
   'Modified by Lydia 2016/03/11 國家檔的CFT承辦人(NA69)改成模組(DB.Functions)
   'Remove by Lydia 2016/03/24 外商反應未考慮好,先移除
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            "select * from ( SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,GETNA69(NP02,NP03,NP04,NP05,'','') NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL AND TM10<>'011' AND TM10=NA01(+)" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)" & _
               " ) where na16 in (" & stNumList & ") "
   'Modified by Lydia 2016/11/15 TM10<>'011' => TM10 NOT IN (" & strExpNa01 & ")
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
   'Modified by Lydia 2025/05/07 合併「CFT特殊管制國家」的語法; (原) AND TM10 NOT IN (" & strExpNa01 & ") AND NA69 in (" & stNumList & ") => AND (TM10 IN (" & strExpNa01 & ") or NP10 in (" & stNumList & "))
                               '(原),NA69 NA16, => ," & IIf(strExpNa01 = "AAA", "NP10 NA16", "'" & stUserID & "' AS NA16") & ",
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08," & IIf(strExpNa01 = "AAA", "NP10 NA16", "'" & stUserID & "' AS NA16") & ",NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10=NA01(+) AND (TM10 IN (" & strExpNa01 & ") or NP10 in (" & stNumList & "))" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件 'Memo by Lydia 2025/05/07 TF案原本為內商管理，但分案到外商則由外商負責，所以判斷NP10 or CP14
   'Modified by Lydia 2025/05/07 不用判斷「CFT特殊管制國家」，拿掉TM10 NOT IN (" & strExpNa01 & ")；
                                 '(原),NA69 NA16, =>,NP10 NA16,
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10=NA01(+) AND NP10 in (" & stNumList & ")　" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/09/06
   'Modified by Lydia 2016/11/15  SP09<>'011' => SP09 NOT IN (" & strExpNa01 & ")
   'Modified by Lydia 2018/12/04 CFC不抓NA69
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & strExpNa01 & ") AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
   'Modified by Lydia 2025/05/07 不用判斷「CFT特殊管制國家」，拿掉SP09 NOT IN (" & strExpNa01 & ")；
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   '2014/9/12 END
   'Added by Lydia 2022/11/21 延展102仍保持只跳二天，目前所有中間程序從只跳二天(系統日＋５或６個工作天)，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
   'Memo by Lydia 2025/05/06 (2022/11/21)備註雖然有提到「持續跳到智權同仁有收文接洽單或填結案單為止」，實際只有跳2天和本所期限當日起有彈跳期限。
   'Modified by Lydia 2025/05/06 調整「所有中間程序(非102,305,997,998)從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
               '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
   'Modified by Lydia 2025/05/07 合併「CFT特殊管制國家」的語法; (原) AND TM10 NOT IN (" & strExpNa01 & ") AND NA69 in (" & stNumList & ") =>AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))
                                 '(原),NA69 NA16, => ," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10=NA01(+) AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件 'Memo by Lydia 2025/05/07 TF案原本為內商管理，但分案到外商則由外商負責，所以判斷NP10 or CP14
   'Modified by Lydia 2025/05/06 比照CFT案，調整「從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
               '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
   'Modified by Lydia 2025/05/07 不用判斷「CFT特殊管制國家」，拿掉TM10 NOT IN (" & strExpNa01 & ")；
                                 '(原),NA69 NA16, =>,NP10 NA16,
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE  NP02||NP06='TF' AND TM10<>'000' " & _
               " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/09/06
   'Modified by Lydia 2025/05/06 比照CFT案，調整「從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
               '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
   'Modified by Lydia 2025/05/07 不用判斷「CFT特殊管制國家」，拿掉SP09 NOT IN (" & strExpNa01 & ")；
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2022/11/21
   'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
   'Modified by Lydia 2025/05/06 修改為「系統日+22個工作天之未收文案件」，並僅彈跳兩天提醒。
               '(原) AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) => AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E
   'Modified by Lydia 2025/05/07 合併「CFT特殊管制國家」的語法; (原) AND TM10 NOT IN (" & strExpNa01 & ") AND NA69 in (" & stNumList & ") => AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))
                                 '(原),NA69 NA16, => ," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08," & IIf(strExpNa01 = "AAA", "NA69 NA16", "'" & stUserID & "' AS NA16") & ",NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E & _
               " AND np07 IN ('102','105')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10=NA01(+) AND (TM10 IN (" & strExpNa01 & ") or NA69 in (" & stNumList & "))" & _
               " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/07/31
   
   'Added by Lydia 2025/05/07 合併「CFT特殊管制國家」的語法：刪除非設定的業務區
   txtData = Empty
   If strExCond <> "" Then
      txtData = Split(strExCond, ",")
      strExc(0) = "": strExc(1) = "": strExc(2) = ""
      strWhSql = ""
      '處理CFT共同語法; 另外TF,CFC,S只判斷CP14,NP10; 因為承辦收文達期限的NA16=空白，所以要排除NA16=空白
      strSql = "DELETE FROM R030301 WHERE NVL(NA16,'N') <> 'N' AND (ID,EV1,EV2,CP09,NP22) IN ( " & _
               "SELECT A.ID,A.EV1,A.EV2,A.CP09,A.NP22 FROM R030301 A,CASEPROGRESS B " & _
               ",(SELECT VC01,VC02,VC03,VC04,SUBSTR(VC05,18,1) ST06 FROM ( " & _
               "SELECT CP01 AS VC01,CP02 AS VC02,CP03 AS VC03,CP04 AS VC04,MAX(CP05||CP09||ST06) VC05 FROM CASEPROGRESS,STAFF " & _
               "WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM R030301 A,CASEPROGRESS B,TRADEMARK " & _
               "WHERE ID='" & strUserNum & "' AND A.CP09=B.CP09(+) AND B.CP01=TM01(+) AND B.CP02=TM02(+) AND B.CP03=TM03(+) AND B.CP04=TM04(+) " & _
               "AND TM01='CFT' AND TM10='999') AND CP159=0 AND CP09<'D' AND CP13=ST01(+) GROUP BY CP01,CP02,CP03,CP04)) VTB01 " & _
               "WHERE ID='" & strUserNum & "' AND A.CP09=B.CP09(+) AND B.CP01=VC01(+) AND B.CP02=VC02(+) AND B.CP03=VC03(+) AND B.CP04=VC04(+) ZZZ)"
      For ii = 0 To UBound(txtData)
         If Trim(txtData(ii)) <> "" Then
            strExc(1) = Mid(Trim(txtData(ii)), 5, 3)
            If strExc(0) <> "" And strExc(0) <> strExc(1) Then
               If strWhSql <> "" Then
                  If InStr("101,239", strExc(0)) > 0 Then
                     strWhSql = " AND st06 not in (" & Mid(strWhSql, 2) & ") "
                  End If
               End If
               If strWhSql <> "" Then
                  strExc(2) = Replace(Replace(UCase(strSql), "999", strExc(0)), "ZZZ", strWhSql)
                  cnnConnection.Execute strExc(2), intI
               End If
               strWhSql = ""
            End If
            Select Case strExc(1)
               Case "011" '日本: A=中南高所、B=北所
                  If InStr(strExCond, strExc(1) & "A") > 0 And InStr(strExCond, strExc(1) & "B") > 0 Then
                     '同一國家全部屬於同一人
                  Else
                     If Right(Trim(txtData(ii)), 1) = "B" Then
                        strWhSql = " AND st06 in ('2','3','4') "
                     Else
                        strWhSql = " AND st06 in ('1') "
                     End If
                  End If
               Case "101", "239" '美國、歐盟:  A=南高所、B=北所、C=中所
                  If InStr(strExCond, strExc(1) & "A") > 0 And InStr(strExCond, strExc(1) & "B") > 0 And InStr(strExCond, strExc(1) & "C") > 0 Then
                     '同一國家全部屬於同一人
                  Else
                     If Right(Trim(txtData(ii)), 1) = "A" Then
                        strWhSql = strWhSql & ",'3','4'"
                     ElseIf Right(Trim(txtData(ii)), 1) = "B" Then
                        strWhSql = strWhSql & ",'1'"
                     ElseIf Right(Trim(txtData(ii)), 1) = "C" Then
                        strWhSql = strWhSql & ",'2'"
                     End If
                  End If
            End Select
            strExc(0) = strExc(1)
         End If
      Next ii
      If strWhSql <> "" And strExc(0) <> "" Then
         If InStr("101,239", strExc(0)) > 0 Then
            strWhSql = " AND st06 not in (" & Mid(strWhSql, 2) & ") "
         End If
         strExc(2) = Replace(Replace(UCase(strSql), "999", strExc(0)), "ZZZ", strWhSql)
         cnnConnection.Execute strExc(2), intI
         strWhSql = ""
      End If
   End If
   GoTo JumpToMerge
   'end 2025/05/06   -----合併「CFT特殊管制國家」的語法：刪除非設定的業務區
   
'Mark by Lydia 2025/05/06 CFT特殊管制國家：改成先暫存所有案件，再刪除非管制的案件
'**********CFT特殊管制國strExpNa01**********
'   '申請國家為日本時，管制人(中南高所)78011葉易雲
'   '                                                (北所)        98018蔡庭蓁
'   'Modify By Sindy 2010/5/31
'   '申請國家為日本時，管制人(中南高所)99011王婉
'   '                                                (北所)        98018蔡庭蓁
'   strSql = "select OMAN from SetSpecMan where OCODE='CFT_011A'"
'   intI = 1: strUser = ""
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If Not IsNull(RsTemp.Fields(0)) Then
'         strUser = Trim(RsTemp.Fields(0)) 'CFT承辦人日本(中南高所)管制人
'      End If
'   End If
'   txtData = Split(stNumList, strUser)
'   If UBound(txtData) = 1 Then
'      strWhSql = " AND st06 in ('2','3','4')"
'   Else
'      strSql = "select OMAN from SetSpecMan where OCODE='CFT_011B'"
'      intI = 1: strUser = ""
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If Not IsNull(RsTemp.Fields(0)) Then
'            strUser = Trim(RsTemp.Fields(0)) 'CFT承辦人日本(北所)管制人
'         End If
'      End If
'      txtData = Split(stNumList, strUser)
'      If UBound(txtData) = 1 Then
'         strWhSql = " AND st06 in ('1')"
'      End If
'   End If
'   If UBound(txtData) = 1 Then
'      'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
'      'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
'      'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
'                ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
'      'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
'      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
'      'Modified by Lydia 2025/05/06 延展(102)：從系統日＋５或６個工作天前只彈跳2天提醒，改成從系統日＋６個工作天起持續跳至專用期限滿1個工作天後再解除彈跳。(原) CompWorkDay(6, strSrvDate(1)) => stDate_B1
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND NP09>=" & stDate_B1 & " AND NP09< " & stDate6 & _
'                  " AND np07 IN ('102')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'Modified by Lydia 2016/10/14 NP02||NP06 in ('S','CFC') => NP02 in ('S','CFC') AND NVL(NP06,'0')='0' ; + /*+ INDEX(NEXTPROGRESS IDXNP09020706) */
'      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
'      'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                  " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
'                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
'                  " AND np07 NOT IN ('305','997','998')" & _
'                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                  " AND SP15 IS NULL" & _
'                  " AND substr(SP09,1,3)='011' " & _
'                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      'cnnConnection.Execute strSql, intI
'
'      'Add By Sindy 2014/9/12 催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
'      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29 IS NULL AND substr(TM10,1,3)='011'" & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29 IS NULL AND substr(TM10,1,3)='011'" & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'end 2024/09/06
'      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
'      'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                  " WHERE NP02||NP06 in ('S','CFC')" & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305')" & _
'                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                  " AND SP15 IS NULL" & _
'                  " AND substr(SP09,1,3)='011' " & _
'                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      'cnnConnection.Execute strSql, intI
'      '2014/9/12 END
'      'Added by Lydia 2022/11/21 延展102仍保持只跳二天，目前所有中間程序從只跳二天(系統日＋５或６個工作天)，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
'      'Memo by Lydia 2025/05/06 (2022/11/21)備註雖然有提到「持續跳到智權同仁有收文接洽單或填結案單為止」，實際只有跳2天和本所期限當日起有彈跳期限。
'      'Modified by Lydia 2025/05/06 調整「所有中間程序(非102,305,997,998)從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
'                 '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
'                  " AND np07 NOT IN ('102','305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'END 2022/11/21
'      'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
'      'Modified by Lydia 2025/05/06 修改為「系統日+22個工作天之未收文案件」，並僅彈跳兩天提醒。
'               '(原) AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) => AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E
'       strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E & _
'                  " AND np07 IN ('102','105')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+) " & strWhSql & _
'                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'end 2024/07/31
'   End If
'   '2010/3/1 End
'   'Added by Lydia 2018/12/04 CFC不判斷特殊設定
'   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
'              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=>SP15||SP61 IS NULL
'   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
'   'Modified by Lydia 2025/05/06 延展(102)：從系統日＋５或６個工作天前只彈跳2天提醒，改成從系統日＋６個工作天起持續跳至專用期限滿1個工作天後再解除彈跳。(原) CompWorkDay(6, strSrvDate(1)) => stDate_B1
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
'                " AND NP09>=" & stDate_B1 & " AND NP09< " & stDate6 & _
'                " AND np07 IN ('102')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15||SP61 IS NULL" & _
'                " AND substr(SP09,1,3)='011' " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02||NP06 in ('S','CFC')" & _
'                " AND NP08<=" & strSrvDate(1) & _
'                " AND np07 IN ('305','997','998')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15 IS NULL" & _
'                " AND substr(SP09,1,3)='011' " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'end 2018/12/04
'    'Added by Lydia 2022/11/21 延展102仍保持只跳二天，目前所有中間程序從只跳二天(系統日＋５或６個工作天)，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
'    'Memo by Lydia 2025/05/06 (2022/11/21)備註雖然有提到「持續跳到智權同仁有收文接洽單或填結案單為止」，實際只有跳2天和本所期限當日起有彈跳期限。
'    'Modified by Lydia 2025/05/06 調整「所有中間程序(非102,305,997,998)從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
'               '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02||NP06 in ('S','CFC')" & _
'                " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
'                " AND np07 NOT IN ('102','305','997','998')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15||SP61 IS NULL" & _
'                " AND substr(SP09,1,3)='011' " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'end 2022/11/21
'
'   'Added by Lydia 2016/11/16 美國和歐盟案件依所別區分管制人
'   strUser = Pub_GetSpecMan("CFT_101239A") '南高所
'   txtData = Split(stNumList, strUser)
'   'Modified by Lydia 2018/10/03 分成CFT_101239B(北所)和CFT_101239C(中所)兩個設定
''   If UBound(txtData) = 1 Then
''      strWhSql = " AND st06 in ('3','4')"
''   Else
''      strUser = Pub_GetSpecMan("CFT_101239B") '北中所
''      txtData = Split(stNumList, strUser)
''      If UBound(txtData) = 1 Then
''         strWhSql = " AND st06 in ('1','2')"
''      End If
''   End If
''   If UBound(txtData) = 1 Then
'   strExc(1) = ""
'   If UBound(txtData) = 1 Then
'        strExc(1) = strExc(1) & "3,4,"
'   End If
'   strUser = Pub_GetSpecMan("CFT_101239B") '北所
'   txtData = Split(stNumList, strUser)
'   If UBound(txtData) = 1 Then
'        strExc(1) = strExc(1) & "1,"
'   End If
'   strUser = Pub_GetSpecMan("CFT_101239C") '中所
'   txtData = Split(stNumList, strUser)
'   If UBound(txtData) = 1 Then
'        strExc(1) = strExc(1) & "2,"
'   End If
'   If strExc(1) <> "" Then
'      strWhSql = " AND st06 in (" & GetAddStr(strExc(1)) & ")"
''end 2018/10/03
'      '延展(102)和第二期(716)專用權須存在(TM17=Y)
'      '未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
'      'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
'              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
'      'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
'      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
'      'Modified by Lydia 2025/05/06 延展(102)：從系統日＋５或６個工作天前只彈跳2天提醒，改成從系統日＋６個工作天起持續跳至專用期限滿1個工作天後再解除彈跳。(原) CompWorkDay(6, strSrvDate(1)) => stDate_B1
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND NP09>=" & stDate_B1 & " AND NP09< " & stDate6 & _
'                  " AND np07 IN ('102')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
''      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
''               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
''                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
''                  " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
''                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
''                  " AND np07 NOT IN ('305','997','998')" & _
''                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
''                  " AND SP15 IS NULL" & _
''                  " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
''                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
''                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
''                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
''                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
''                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
''      cnnConnection.Execute strSql, intI
'
'      '催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
'      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT' " & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
'      'Modified by Lydia 2025/05/05 Debug: (修改前)NP02||NP06='CFT' -> NP02||NP06='TF'
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'end 2024/09/06
'      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
''      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
''               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
''                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
''                  " WHERE NP02||NP06 in ('S','CFC')" & _
''                  " AND NP08<=" & strSrvDate(1) & _
''                  " AND np07 IN ('305')" & _
''                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
''                  " AND SP15 IS NULL" & _
''                  " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
''                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
''                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
''                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
''                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
''                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
''      cnnConnection.Execute strSql, intI
'      'Added by Lydia 2022/11/21 延展102仍保持只跳二天，目前所有中間程序從只跳二天(系統日＋５或６個工作天)，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
'      'Memo by Lydia 2025/05/06 (2022/11/21)備註雖然有提到「持續跳到智權同仁有收文接洽單或填結案單為止」，實際只有跳2天和本所期限當日起有彈跳期限。
'      'Modified by Lydia 2025/05/06 調整「所有中間程序(非102,305,997,998)從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
'                  '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
'                  " AND np07 NOT IN ('102','305','997','998')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'END 2022/11/21
'      'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
'      'Modified by Lydia 2025/05/06 修改為「系統日+22個工作天之未收文案件」，並僅彈跳兩天提醒。
'               '(原) AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) => AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
'                  " WHERE NP02||NP06='CFT'" & _
'                  " AND NP09>" & stDate102105_B & " AND NP09<= " & stDate102105_E & _
'                  " AND np07 IN ('102','105')" & _
'                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
'                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
'                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+) " & strWhSql & _
'                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
'      'end 2024/07/31
'   End If
'    'Added by Lydia 2018/12/04 CFC不判斷特殊設定
'   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)。
'              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=>SP15||SP61 IS NULL
'    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
'    'Modified by Lydia 2025/05/06 延展(102)：從系統日＋５或６個工作天前只彈跳2天提醒，改成從系統日＋６個工作天起持續跳至專用期限滿1個工作天後再解除彈跳。(原) CompWorkDay(6, strSrvDate(1)) => stDate_B1
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
'                " AND NP09>=" & stDate_B1 & " AND NP09< " & stDate6 & _
'                " AND np07 IN ('102')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15||SP61 IS NULL" & _
'                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02||NP06 in ('S','CFC')" & _
'                " AND NP08<=" & strSrvDate(1) & _
'                " AND np07 IN ('305','997','998')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15 IS NULL" & _
'                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'end 2018/12/04
'   'end 2016/11/16
'    'Added by Lydia 2022/11/21 延展102仍保持只跳二天，目前所有中間程序從只跳二天(系統日＋５或６個工作天)，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
'    'Memo by Lydia 2025/05/06 (2022/11/21)備註雖然有提到「持續跳到智權同仁有收文接洽單或填結案單為止」，實際只有跳2天和本所期限當日起有彈跳期限。
'    'Modified by Lydia 2025/05/06 調整「所有中間程序(非102,305,997,998)從法限<=系統日＋６個工作天起，或是本所期限當日起彈跳提醒，持續跳到智權同仁有收文接洽單或填結案單為止。」
'                '(原) AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ") => AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")
'    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                " WHERE NP02||NP06 in ('S','CFC')" & _
'                " AND (NP09<= " & stDate5 & " or NP08<=" & strSrvDate(1) & ")" & _
'                " AND np07 NOT IN ('102','305','997','998')" & _
'                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                " AND SP15||SP61 IS NULL" & _
'                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
'                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
'                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'    cnnConnection.Execute strSql, intI
'    'end 2022/11/21
'**********CFT特殊管制國strExpNa01**********
JumpToMerge: 'Added by Lydia 2025/05/06

'***********************
''E' EV1,'1' EV2
'***********************
   '所有未發文--承辦人-E(未發文) -- 含T案
   If idx = 1 Then
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('T','FCT','CFT') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Add By Sindy 2015/10/23 +法務:未發文
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, Lawcase" & _
                  " WHERE CP01 in('FCL','CFL','LIN','ACS') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                  " AND LC08 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      '2015/10/23 END
      
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
      'Add By Sindy 2010/7/29 另外依F4103期限做區分
      If txtUsernum = "78011" Or txtUsernum = "80030" Then
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, TradeMark" & _
                        " WHERE CP01 in('T','FCT','CFT') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103TSql & " AND CP158=0 AND CP159=0" & _
                        " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                        " AND TM29 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
            'Add By Sindy 2015/10/23 +法務:未發文
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, Lawcase" & _
                        " WHERE CP01 in('FCL','CFL','LIN','ACS') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103LSql & " AND CP158=0 AND CP159=0" & _
                        " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                        " AND LC08 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
            '2015/10/23 END
            
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, SERVICEPRACTICE" & _
                        " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103SSql & " AND CP158=0 AND CP159=0" & _
                        " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                        " AND SP15 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
      End If
   End If
   
'***********************
''A' EV1,'0' EV2
'***********************
   '未分案-0 -- 不含T案
   If bLvl4 = True Or bLvl5 = True Then
      '已收文未發,2個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('FCT','CFT')" & _
                  " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC')" & _
                  " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
'***********************
''B' EV1,'0' EV2
'***********************
      '已收文未發,2個工作天後達承辦期限者(不含當日)-B0(承辦期限,未分案)
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('FCT','CFT')" & _
                  " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC')" & _
                  " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
'***********************
''E' EV1,'0' EV2
'***********************
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
         'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','CFT') AND CP05>20030000" & _
                     " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
         'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                     " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
'***********************
''B' EV1,'1' EV2
'***********************
   'Add By Sindy 2012/6/4
   '【T、FCT台灣商標爭議案逾承辦期限、逾指定會稿日】
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, Trademark, EngineerProgress" & _
               " WHERE CP05>=20120601" & _
               " AND CP01 in('T','FCT')" & _
               " AND CP10 in(" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)" & _
               " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
               " AND TM10='000' and tm29 is null" & _
               " AND CP09=EP02(+)" & _
               " AND (CP48<" & strSrvDate(1) & " and CP48 is not null)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2018/12/10 +T台灣案非爭議案
   'Remove by Lydia 2019/01/30
'   If strSrvDate(1) >= T案收文齊備啟用日 Then
'        strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'                 " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
'                    "From CASEPROGRESS, Trademark, EngineerProgress" & _
'                    " WHERE CP05>=" & T案收文齊備啟用日 & _
'                    " AND CP01 ='T'" & _
'                    " AND CP10 not in(" & TMdebate & ")" & _
'                    " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0 AND CP09 LIKE 'A%' " & _
'                    " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
'                    " AND TM10='000' and tm29 is null" & _
'                    " AND CP09=EP02(+)" & _
'                    " AND (CP48<" & strSrvDate(1) & " and CP48 is not null)"
'        cnnConnection.Execute strSql, intI
'   End If
   'end 2018/12/10
'***********************
''I' EV1,'1' EV2
'***********************
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','I' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, Trademark, EngineerProgress" & _
               " WHERE CP05>=20120601" & _
               " AND CP01 in('T','FCT')" & _
               " AND CP10 in(" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)" & _
               " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
               " AND TM10='000' and tm29 is null" & _
               " AND CP09=EP02(+)" & _
               " AND (EP28<" & strSrvDate(1) & " and EP28 is not null)"
   cnnConnection.Execute strSql, intI
   '2012/6/4 End
   'Added by Lydia 2018/12/10 +T台灣案非爭議案
   'Remove by Lydia 2019/01/30
'   If strSrvDate(1) >= T案收文齊備啟用日 Then
'        strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'                 " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','I' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
'                    "From CASEPROGRESS, Trademark, EngineerProgress" & _
'                    " WHERE CP05>=" & T案收文齊備啟用日 & _
'                    " AND CP01 ='T'" & _
'                    " AND CP10 not in(" & TMdebate & ")" & _
'                    " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0 AND CP09 LIKE 'A%' " & _
'                    " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
'                    " AND TM10='000' and tm29 is null" & _
'                    " AND CP09=EP02(+)" & _
'                    " AND (EP28<" & strSrvDate(1) & " and EP28 is not null)"
'        cnnConnection.Execute strSql, intI
'   End If
   'end 2018/12/10
   
   'Add By Sindy 2015/2/17 若同案有程序的J.今送件及A.達本所期限,就不要再顯示達本所
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='A' and R1.EV2='1'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='J' and R2.EV2='1' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '2015/2/17 END
   
   'Add By Sindy 2017/9/11
   '若同案有 'H達法定'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'H'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='H' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '若同案有 'A達本所'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'A'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='A' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '2017/9/11 END
   
   'Add By Sindy 2015/5/7 若為程序組人員時,只有程序主管需顯示未收文之延展案件
   'Modify By Sindy 2020/7/2 程序不管制未收文延展期限
   '程序組
   'If Trim(stDept) = "F12" And stUserID <> Pub_GetSpecMan("P_FCT") Then
   If Trim(stDept) = "F12" Then
      strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='D'" & _
               " AND R1.ID||R1.EV1||R1.EV2||R1.CP09||R1.NP22 in (select R2.ID||R2.EV1||R2.EV2||R2.CP09||R2.NP22 from R030301 R2,nextprogress where R2.ID='" & strUserNum & "' and R2.EV1='D' and R2.CP09=nextprogress.NP01 and R2.np22=nextprogress.NP22 and nextprogress.NP07='102')"
      cnnConnection.Execute strSql, intI
   Else
      'Modify By Sindy 2020/7/2 未收文延展期限，英文組改為法定期限＋１個工作天＋7個月才開始提醒，日文組維持原規則。
      strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='D'" & _
               " AND R1.ID||R1.EV1||R1.EV2||R1.CP09||R1.NP22 in (" & _
               "select R2.ID||R2.EV1||R2.EV2||R2.CP09||R2.NP22" & _
               " from R030301 R2,nextprogress,trademark,fagent" & _
               " where R2.ID='" & strUserNum & "' and R2.EV1='D' and R2.CP09=nextprogress.NP01 and R2.np22=nextprogress.NP22 and nextprogress.NP07='102'" & _
               " AND NP02=TM01 AND NP03=TM02 AND NP04=TM03 AND NP05=TM04" & _
               " AND substr(TM44,1,8)=fa01(+) AND substr(TM44,9,1)=fa02(+)" & _
               " AND substr(fa10,1,3)<>'011' AND fa10 IS NOT NULL" & _
               " AND WORKDAYADD(+1,to_char(add_months(to_date(np09,'YYYYMMDD'),7),'YYYYMMDD'))>" & strSrvDate(1) & _
               ")"
      cnnConnection.Execute strSql, intI
   End If
   '2015/5/7 END
   
   '案件進度
   'Modify By Sindy 2021/6/3 + ,'K','需請款'
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   'Modify By Sindy 2023/12/11 + ,'N','達指定'
   strExc(0) = "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "TM05 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,CP10,CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04" & _
      " FROM R030301,trademark,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   'Add By Sindy 2015/10/23
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(LC15,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "NVL(LC05,NVL(LC06,LC07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,LC01,LC02,LC03,LC04,'' 未收款,CP10,CP27,R030301.CP09,LC15,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04" & _
      " FROM R030301,Lawcase,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/10/23 END
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,CP10,CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04" & _
      " FROM R030301,SERVICEPRACTICE,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '下一程序
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "TM05 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,NP07,0 CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04" & _
      " FROM R030301,trademark,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   'Add By Sindy 2015/10/23
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(LC15,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "NVL(LC05,NVL(LC06,LC07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,LC01,LC02,LC03,LC04,'' 未收款,NP07,0 CP27,R030301.CP09,LC15,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04" & _
      " FROM R030301,Lawcase,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/10/23 END
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,NP07,0 CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04" & _
      " FROM R030301,SERVICEPRACTICE,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/2/12 END
   'Modify By Sindy 2015/1/30
'   strExc(0) = "SELECT V,本所期限,法定期限,承辦期限,核稿期限,管制人,承辦人,智權人員" & _
'                     ",事件,本所案號,案件性質,案件備註,案件名稱,代理人國籍," & _
'                     "EV1,EV2,NA16,CP14,CP13,TM01,TM02,TM03,TM04,未收款,CP10,CP27,CP09,TM10,ti01,NP22,CP06 from (" & _
'               strExc(0) & ")"
   Select Case stDept
      Case "F11", "F10" '承辦組
         'strExc(0) = strExc(0) & " order by 本所期限 asc,智權人員 asc,本所案號 asc"
         strExc(0) = strExc(0) & " order by sort asc,智權人員 asc,本所案號 asc"
      'Case "F12" '程序組
      Case Else
         'strExc(0) = strExc(0) & " order by 本所期限 asc,承辦人 asc,本所案號 asc"
         strExc(0) = strExc(0) & " order by sort asc,承辦人 asc,本所案號 asc"
   End Select
   If rsTmp.State <> 0 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grdDataList.Recordset = rsTmp
      SetXRecord '更新欄位值
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   Else
      Screen.MousePointer = vbDefault
      MsgBox "查無資料！", vbInformation
      rsTmp.Close
      Set rsTmp = Nothing
      cmdHide.Enabled = False
      lblCnt.Caption = "共 0 筆"
      Exit Sub 'Add By Sindy 2014/9/17
   End If
   rsTmp.Close
   
   'Modify By Sindy 2014/9/17 案件性質+相關總收文號的案件性質
   For iRow = 1 To grdDataList.Rows - 1
      grdDataList.TextMatrix(iRow, 10) = grdDataList.TextMatrix(iRow, 10) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(iRow, 26), grdDataList.TextMatrix(iRow, 29), "1")
   Next iRow
   '2014/9/17 END
End Sub

'更新欄位值
Private Function SetXRecord()
Dim iRow As Integer
Dim strNote As String
   
   For iRow = 1 To grdDataList.Rows - 1
      'Add By Sindy 2023/12/11
      If grdDataList.TextMatrix(iRow, 8) = "達指定" And grdDataList.TextMatrix(iRow, 26) <> "" Then '26=CP09
         strSql = "select * from caseprogress where cp09='" & grdDataList.TextMatrix(iRow, 26) & "'"
         intI = 1
         strNote = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Val("" & RsTemp.Fields("cp142")) > 0 Then
               strNote = "客戶指定" & ChangeWStringToTDateString(RsTemp.Fields("cp142"))
               If "" & RsTemp.Fields("cp164") <> "" Then
                  strNote = strNote & IIf(RsTemp.Fields("cp164") = "1", "當天", IIf(RsTemp.Fields("cp164") = "2", "之前", IIf(RsTemp.Fields("cp164") = "3", "之後", ""))) & "送件;"
               End If
               '備註
               grdDataList.TextMatrix(iRow, 11) = "達指定;" & strNote & grdDataList.TextMatrix(iRow, 11)
            End If
         End If
      End If
      '2023/12/11 END
   Next iRow
End Function

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      '                0 1          2          3          4          5       6       7         8       9            10        11           12
      .FormatString = "V|本所期限  |法定期限  |承辦期限  |核稿期限  |管制人 |承辦人 |智權人員 |事件　 |本所案號　　|案件性質 |備註　　　　|案件名稱　　　"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         'If (intI > 12 And intI < 23) Or intI > 23 Then
         If (intI > 3 And intI < 6) Or intI > 12 Then
            .ColWidth(intI) = 0
         End If
      Next
'      .ColWidth(23) = 700
'      .ColAlignment(23) = flexAlignRightTop
      .ColAlignment(1) = flexAlignRightTop
      .ColAlignment(2) = flexAlignRightTop
      .ColAlignment(3) = flexAlignRightTop
      .ColAlignment(4) = flexAlignRightTop
      .ColWidth(11) = 1300
      .ColWidth(12) = 1300
      .Visible = True
   End With
End Sub

Private Sub SetColor(Optional sHide As String = "N")
Dim lngToday As Long, lngCP06 As Long, lngCP48 As Long, lngEP08 As Long, stType As String
Dim lngCP07 As Long
Dim ii As Integer, jj As Integer, dblCnt As Double
Dim strBDate As String 'Added by Lydia 2024/07/31

   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(2))
      strBDate = TransDate(CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1), 1) 'Added by Lydia 2024/07/31
      For ii = 1 To .Rows - 1
         lngCP06 = Val(Replace(.TextMatrix(ii, 1), "/", "")) '本所期限
         lngCP07 = Val(Replace(.TextMatrix(ii, 2), "/", "")) '法定期限
         lngCP48 = Val(Replace(.TextMatrix(ii, 3), "/", "")) '承辦期限
         lngEP08 = Val(Replace(.TextMatrix(ii, 4), "/", "")) '核稿期限
         stType = .TextMatrix(ii, 14)
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
         
         'Move by Lydia 2025/05/06 ----- Added by Lydia 2024/07/31 延展(102)/使用宣誓(105)未收文案件於30日曆天前之期限提醒於事件欄位以「◆未收文」顯示，並僅彈跳兩天提醒。
         If stType = "D" And Mid("" & .TextMatrix(ii, 9), 1, 3) = "CFT" And _
               ("" & .TextMatrix(ii, 24) = "102" Or "" & .TextMatrix(ii, 24) = "105") And lngCP07 >= strBDate Then
               .TextMatrix(ii, 8) = "◆" & .TextMatrix(ii, 8)
         'end 2024/07/31
         '逾管控期限
         'Add By Sindy 2012/6/4 +stType = "I"
         '((stType = "B" Or stType = "G" Or stType = "J") And lngCP48 > 0 And lngCP48 < lngToday) Or
         'Modified by Lydia 2025/05/06 原本在下面「延展(102)/使用宣誓(105)未收文案件」的移上來
         'If ((stType = "A" Or stType = "D" Or stType = "F") And lngCP06 > 0 And lngCP06 < lngToday) Or
         ElseIf ((stType = "A" Or stType = "D" Or stType = "F") And lngCP06 > 0 And lngCP06 < lngToday) Or _
            ((stType = "B" Or stType = "G") And lngCP48 > 0 And lngCP48 < lngToday) Or _
            (stType = "C" And lngEP08 > 0 And lngEP08 < lngToday) Or _
            (stType = "H" And lngCP07 > 0 And lngCP07 < lngToday) Or _
            stType = "I" Then
            .TextMatrix(ii, 9) = "*" & Trim(.TextMatrix(ii, 9))
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         '((stType = "B" Or stType = "G" Or stType = "J") And lngCP48 = lngToday) Or
         ElseIf ((stType = "A" Or stType = "D" Or stType = "F") And lngCP06 = lngToday) Or _
            ((stType = "B" Or stType = "G") And lngCP48 = lngToday) Or _
            (stType = "C" And lngEP08 = lngToday) Or _
            (stType = "H" And lngCP07 = lngToday) Then
            .TextMatrix(ii, 9) = "v" & Trim(.TextMatrix(ii, 9))
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         'Add By Sindy 2015/5/7 今送件
         'Modify By Sindy 2021/6/3 + 需請款
         ElseIf stType = "J" Or stType = "K" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '紫
               .CellBackColor = &HFF80FF   '&HFF8080
            Next
         '2015/5/7 END
         '未分案
         ElseIf .TextMatrix(ii, 15) = "0" Then
            .TextMatrix(ii, 9) = "#" & Trim(.TextMatrix(ii, 9)) 'Add By Sindy 2015/5/7
            '第五級不看
            If bLvl5 = True Then
               .RowHeight(ii) = 0
            Else
               For jj = 1 To .Cols - 1
                  .col = jj
                  '黃
                  .CellBackColor = &HFFFF&
               Next
            End If
         ElseIf sHide <> "N" Then
            .RowHeight(ii) = 0
         Else
            strExc(1) = .TextMatrix(ii, 15)
            Select Case strExc(1)
               '承辦人,核稿人
               Case "1", "4"
                  strExc(2) = .TextMatrix(ii, 17)
               Case "2" '管制人
                  strExc(2) = .TextMatrix(ii, 16)
               Case "3" '智權人員
                  strExc(2) = .TextMatrix(ii, 18)
               Case Else
                  strExc(2) = ""
            End Select
            
            If strExc(2) <> "" Then
               'Add By Sindy 2010/7/29 例外情況
               If (Trim(txtUsernum) = "78011" Or Trim(txtUsernum) = "80030") And strExc(2) = "F4103" Then
                  '78011及80030為F4103的第二級主管
               Else
                  '本人或第二級才看
                  'Modify By Sindy 2021/7/5 外商的期限彈跳改2~4級主管全部彈(原本3級以上主管只彈逾期資料)
                  'If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
                  If InStr(stNumList1(1) & "," & stNumList1(2) & "," & stNumList1(3) & "," & stNumList1(4), strExc(2)) = 0 Then
                     .RowHeight(ii) = 0
                  End If
               End If
            End If
         End If
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With
   lblCnt.Caption = "共 " & dblCnt & " 筆"
   If sHide = "N" Then
      cmdHide.Tag = "Y"
      cmdHide.Caption = "隱藏白色(&H)"
   Else
      cmdHide.Tag = "N"
      cmdHide.Caption = "顯示白色(&S)"
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Combo1.Clear
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v)：表示當日期限"
   Combo1.AddItem "黃色(#)：表示未分案"
   'Modify By Sindy 2021/6/3 + 、需請款
   Combo1.AddItem "紫色：表示今送件、需請款" 'Add By Sindy 2015/5/7
   Combo1.AddItem "藍色：表示點選資料"
   Combo1.ListIndex = 0
   txtUsernum = strUserNum
   If Pub_StrUserSt03 = "M51" Then
      txtUsernum.Enabled = True
   End If
   'Added by Lydia 2023/05/10 開放輸入「員工編號」欄：總經理
   If InStr("01,08,", Pub_strUserST05 & ",") > 0 Then
      txtUsernum.Enabled = True
   End If
   'end 2023/05/10
   
'   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
'   PUB_AddExcuteLog Me.Name
   
   strTemplatePath = PUB_DownloadOftPath("F11", "") 'Add By Sindy 2024/7/22 下載郵件範本
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Memo by Lydia 2019/11/04   國外部自動通知順序: FMP案frm060206=> 國外部期限frm060204=> 外商frm030301=> 外法frm072005=>國外部行事曆frm060209
   If Not bolUnloading Then 'Add by Sindy 2016/7/22
   
      Dim strSql As String, bolRun As Boolean
   
      '電腦中心除外
      If Pub_StrUserSt03 <> "M51" Then
         '專利
         bolRun = False
         'Modified by Lydia 2016/10/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp14='" & strUserNum & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         If CheckUse("frm060204", strExec, False) = True Or bolRun = True Then
            strSql = "select * from executelog where el01='frm060204' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_bolInformCheck = True 'Add By Sindy 2009/09/21
               Load frm060204
               frm060204.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
         '法務
         bolRun = False
         'Modified by Lydia 2016/10/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select count(*) from caseprogress " & _
                        "where cp01 in ('CFL','FCL','LIN','ACS','L','LA') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 and cp06>0 " & _
                        "and cp14='" & strUserNum & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
'         If strGroup = "F1" Or strGroup = "F2" Or strGroup = "D4" Or bolRun = True Then
            If CheckUse("frm072005", strExec, False) = True Or bolRun = True Then
               strSql = "select * from executelog where el01='frm072005' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI <> 1 Then
                  pub_bolInformCheck = True 'Add By Sindy 2009/09/21
                  Load frm072005
                  frm072005.cmdQuery(0).Value = True
                  Exit Sub
               End If
            End If
'         End If

         'Added by Lydia 2019/11/04 國外部行事曆通知(每天早上和下午自動執行時才run)
         If Left(Pub_StrUserSt03, 2) = "F2" Then
            If CheckUse("frm060209", strExec, False) = True Then
                strSql = "select * from executelog where el01='frm060209' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI <> 1 Then
                   pub_bolInformCheck = True
                   Load frm060209
                   pub_bolInformCheck = False
                   Exit Sub
                End If
            End If
         End If
         'end 2019/11/04
      End If
      
      MenuEnabled
   End If
   
   Set frm030301 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

'Modify By Sindy 2015/1/30
Private Sub GrdDataList_Click()
Dim nCol As Integer, nRow As Integer
      
   With grdDataList
      .Visible = False
      nCol = .MouseCol
      If nCol = 9 Then nCol = 31 'Add By Sindy 2015/8/20 本所案號
      nRow = .MouseRow
      If nRow = 0 Then
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
      End If
      .Visible = True
   End With
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If grdDataList.MouseRow <> 0 And _
      (grdDataList.MouseCol = 11 Or grdDataList.MouseCol = 12) Then
      If iRow <> grdDataList.MouseRow Or iCol <> grdDataList.MouseCol Then
         If grdDataList.TextMatrix(grdDataList.MouseRow, grdDataList.MouseCol) <> "" Then
            CreateToolTip GetHWndForToolTip(grdDataList), grdDataList.TextMatrix(grdDataList.MouseRow, grdDataList.MouseCol)
            iRow = grdDataList.MouseRow
            iCol = grdDataList.MouseCol
         End If
      End If
   End If
End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 4
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            'Modify By Sindy 2010/01/04
            'For ii = 0 To 2
            For ii = 0 To 0
            '2010/01/04 End
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

'Add By Sindy 2010/11/26
Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Memo by Lydia 2025/05/06 先將修改前的程式備份，更名為doQuery_Old
'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery_Old(idx As Integer)
Dim stVTB As String
Dim stDate0 As String, stDate1 As String, stDate2 As String, stDate_3 As String
Dim stDate4 As String
Dim stDate5 As String, stDate6 As String, stDate7 As String, stDate_10 As String
Dim stNumList As String, stDept As String, stDeptST03 As String
Dim ii As Integer, stIdList
Dim stUserID As String
Dim strOtherUser As String
Dim txtData As Variant, strWhSql As String, strUser As String
Dim strF4103TSql As String, strF4103SSql As String, strF4103LSql As String
Dim iRow As Long 'Add By Sindy 2014/9/17
Dim rsTmp As New ADODB.Recordset
Dim stConCP142 As String 'Add By Sindy 2023/12/11
   
   stVTB = ""
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   stUserID = txtUsernum
   '使用者收文智權人員所屬部門
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
      stDeptST03 = Pub_StrUserSt03
   Else
      stDept = GetST15(stUserID)
      stDeptST03 = GetStaffDepartment(stUserID)
   End If
   
   '抓員工外譯對照資料
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   '期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stDate0 = strSrvDate(1) - 10000 '系統日-1年
   stDate1 = CompWorkDay(3, strSrvDate(1))
   stDate2 = CompWorkDay(4, strSrvDate(1))
   stDate4 = CompWorkDay(6, strSrvDate(1)) '4個工作天
   stDate_3 = CompWorkDay(4, strSrvDate(1), 1) '減3個工作天
   stDate5 = CompWorkDay(7, strSrvDate(1))
   stDate6 = CompWorkDay(8, strSrvDate(1))
   stDate7 = CompWorkDay(9, strSrvDate(1))
   stDate_10 = CompWorkDay(11, strSrvDate(1), 1) '減10個工作天
   stConCP142 = " AND CP142>=" & stDate0 & " AND CP142< " & stDate2 'Add By Sindy 2023/12/11 指定日期＜＝系統日＋２個工作天之未發文案件。
   
   '特殊權限
   bLvlX = CheckLevel(stUserID, "M") '未交稿,已完稿無核稿管制人
   bLvl4 = CheckLevel(stUserID, "V") '第四級管制人
   'Modified by Lydia 2018/11/13 改設定
   'bLvl5 = CheckLevel(stUserID, "O") '第五級管制人
   bLvl5 = CheckLevel(stUserID, "V1")
   
   'Modify By Sindy 2015/2/9
   '清除暫存檔
   strSql = "delete R030301 where ID='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2010/7/29 另外依F4103期限做區分
   '日本地區-葉易雲主任 78011
   '其他地區-洪琬姿副理 80030
   strF4103TSql = ""
   strF4103SSql = ""
   strF4103LSql = "" 'Add By Sindy 2015/10/23
   If txtUsernum = "78011" Then
      strF4103TSql = " AND ((TM44 is not null AND exists (select * from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9) and substr(fa10,1,3)='011'))" & _
                       " or (TM44 is null AND exists (select * from customer where cu01=substr(tm23,1,8) and cu02=substr(tm23,9) and substr(cu10,1,3)='011'))) "
      'Add By Sindy 2015/10/23
      strF4103LSql = " AND ((LC22 is not null AND exists (select * from fagent where fa01=substr(LC22,1,8) and fa02=substr(LC22,9) and substr(fa10,1,3)='011'))" & _
                       " or (LC22 is null AND exists (select * from customer where cu01=substr(LC11,1,8) and cu02=substr(LC11,9) and substr(cu10,1,3)='011'))) "
      '2015/10/23 END
      strF4103SSql = " AND ((sp26 is not null AND exists (select * from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9) and substr(fa10,1,3)='011'))" & _
                       " or (sp26 is null AND exists (select * from customer where cu01=substr(sp08,1,8) and cu02=substr(sp08,9) and substr(cu10,1,3)='011'))) "
   ElseIf txtUsernum = "80030" Then
      strF4103TSql = " AND ((TM44 is not null AND exists (select * from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (TM44 is null AND exists (select * from customer where cu01=substr(tm23,1,8) and cu02=substr(tm23,9) and substr(cu10,1,3)<>'011'))) "
      'Add By Sindy 2015/10/23
      strF4103LSql = " AND ((LC22 is not null AND exists (select * from fagent where fa01=substr(LC22,1,8) and fa02=substr(LC22,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (LC22 is null AND exists (select * from customer where cu01=substr(LC11,1,8) and cu02=substr(LC11,9) and substr(cu10,1,3)<>'011'))) "
      '2015/10/23 END
      strF4103SSql = " AND ((sp26 is not null AND exists (select * from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9) and substr(fa10,1,3)<>'011'))" & _
                       " or (sp26 is null AND exists (select * from customer where cu01=substr(sp08,1,8) and cu02=substr(sp08,9) and substr(cu10,1,3)<>'011'))) "
   End If
   
   '代碼1:A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=達法定
   '      I=達指會,J=今送件,K=需請款,N=達指定
   '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   '【FCT、T、S台灣案】
   '承辦組
   If Trim(stDept) = "F11" Or Trim(stDept) = "F10" Or bLvl4 = True Or bLvl5 = True Then
'***********************
''F' EV1,'3' EV2
'***********************
         'Add By Sindy 2010/5/25
         '未請款,(承辦人為非F1*非外商人員)：以發文日次日起算10個工作天
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, staff" & _
               " WHERE CP01 in ('FCT','T','S') AND CP05>=20100101" & _
               " AND CP27<" & stDate_10 & _
               " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in ('FCT','T') AND CP05>=20100101" & _
                     " AND CP27<" & stDate_10 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(11,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S' AND CP05>=20100101" & _
                     " AND CP27<" & stDate_10 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2010/5/25
         '未請款,(承辦人為F1*外商人員)：以發文日次日起算3個工作天
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件; CP01 in ('FCT','T','S')=>CP01 in ('FCT','T','S'" & IIf(Left(stDept, 2) = "F1", "'CFT','CFC'", "") & ")
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, staff" & _
               " WHERE CP01 in ('FCT','T','S'" & IIf(Left(stDept, 2) = "F1", ",'CFT','CFC'", "") & ")  AND CP05>=20100101" & _
               " AND CP27<" & stDate_3 & _
               " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in ('FCT','T') AND CP05>=20100101" & _
                     " AND CP27<" & stDate_3 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S' AND CP05>=20100101" & _
                     " AND CP27<" & stDate_3 & _
                     " AND CP10 not in ('727','303') AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2015/10/23 +法務:未請款
         'Modified by Lydia 2016/10/14 CP20 is null AND CP60 is null AND CP57 is null => NVL(CP20||CP60||CP57,'0')='0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','F' EV1,'3' EV2,CASEPROGRESS.CP09,workdayadd(4,cp27) CP06,0 CP07,nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS" & _
               " WHERE CP01 in ('FCL','CFL','LIN','ACS') AND CP05>=20100101" & _
               " AND CP27<" & stDate_3 & _
               " AND CP16>0 AND NVL(CP20||CP60||CP57,'0')='0' AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND CASEPROGRESS.CP14 is not null" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='F' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
'***********************
''A' EV1,'3' EV2
'***********************
         '已收文未發文,(承辦人為非F1*非外商人員)2個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => NVL(CP27,'0')='0' AND NVL(CP57,'0')='0' AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP13051027) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP13051027) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND NVL(CP27,'0')='0' AND NVL(CP57,'0')='0'" & _
               " AND NVL(CP14,'0') > '0' AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND NVL(CP14,'0') > '0' AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)<>'F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         '已收文未發文,(承辦人為F1*外商人員)1個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件; CP01 in ('FCT','T') =>CP01 in ('FCT','T'" & IIf(Left(stDept, 2) = "F1", ",'CFT'", "") & ")
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in ('FCT','T'" & IIf(Left(stDept, 2) = "F1", ",'CFT'", "") & ")" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0' ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null AND CASEPROGRESS.CP14 is not null => CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'
         'Modified by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
        ' strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01 in ('S'" & IIf(Left(stDept, 2) = "F1", ",'CFC'", "") & ")" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0 AND NVL(CP14,'0') > '0'" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:達本所
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is not null AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         
'***********************
''A' EV1,'0' EV2
'***********************
         '已收文未發文,[未分案]2個工作天後達本所期限者(不含當日) --智權人員-A3(本所期限,智權人員)
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14 is null" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:達本所
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP09<'C' AND CP10 not in('720','722')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14 is null" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         
'***********************
''B' EV1,'3' EV2
'***********************
         '已收文未發文之722外商發文,2個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
'2010/4/19 modify by sonia 發現FCT-027964的核駁前先行通知逾承辦未出現,
                          '故不限制722外商發文,只要承辦人為F1部門者都出現
         'Modify By Sindy 2015/8/24 剔除901.催款
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, staff" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
               " AND CP10 Not in('720','901') AND CP01||CP10<>'FCT1201'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null => CP158=0 AND CP159=0; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2015/8/24 剔除901.催款
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark, staff" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                     " AND CP10 Not in('720','901') AND CP01||CP10<>'FCT1201'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE, staff" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND CASEPROGRESS.CP14=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
'2010/4/19 END
         '已收文未發文之720回覆代理人,1個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
         'Modify By Sindy 2015/8/24 增加901.催款期限控管
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND TM29 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01 in('FCT','T')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 in('720','901')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/11/19 +法務:達承辦
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND LC08 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 ='901'" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/11/19 END
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND SP15 IS NULL
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CP10 in('720','901')" & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2015/8/24 增加901.催款期限控管
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND TM29 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','T')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 in('720','901')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/11/19 +法務:達承辦
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND LC08 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 ='901'" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103LSql & " AND CP57 is null AND CP27 is null" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
               '2015/11/19 END
               'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都要提醒 取消:AND SP15 IS NULL
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
                     " AND CP10 in('720','901')" & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103SSql & " AND CP57 is null AND CP27 is null" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         'Add By Sindy 2015/7/30 已收文未發文之FCT.1201審查報告,4個工作天後達承辦期限者(不含當日) --智權人員-B3(承辦期限,智權人員)
         'Modified by Lydia 2016/10/14
         'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01||CP10='FCT1201'" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ") AND CP57 is null AND CP27 is null" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01='FCT' AND CP10='1201' AND CP158=0 AND CP159=0" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
               " AND CASEPROGRESS.CP13 IN(" & stNumList & ")" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'3' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01||CP10='FCT1201'" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate4 & _
                     " AND CASEPROGRESS.CP13 ='F4103'" & strF4103TSql & " AND CP57 is null AND CP27 is null" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='3' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/7/30 END
                           
'***********************
''D' EV1,'3' EV2
'***********************
         '未收文且 7個工作天 後達本所期限者(不含當日) --智權人員-D3(未收文,智權人員)
         'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
         'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
         'Modify By Sindy 2015/7/30 FCT的208補優先權證明期限要剔除,獨立控管 + and NP02||NP07<>'FCT208'
         'Modify By Sindy 2015/8/21 FCT的901催款要剔除,獨立控管
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02 in ('FCT','T') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                     " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
         If Left(stDept, 2) = "F1" Then
            'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                      " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                         "From NEXTPROGRESS, TradeMark, Caseprogress" & _
                         " WHERE NP02 in ('CFT') AND NP06 is null" & _
                         " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                         " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                         " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                         " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                         " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                         " AND TM29 IS NULL AND NP01=CP09(+) AND SUBSTR(CP12,1,2)='F1'" & _
                         " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
             cnnConnection.Execute strSql, intI
         End If
         'end 2022/11/21
         'Add By Sindy 2015/10/23 +法務:未收文
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Lawcase, Staff" & _
                     " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('901') AND NP10=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, SERVICEPRACTICE" & _
                     " WHERE NP02='S' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                     " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2022/11/21 CFT彈跳期限增加顯示智權同仁為FCT承辦人並且為”１未收文,５達本所,９未請款”之案件;
         If Left(stDept, 2) = "F1" Then
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                        "From NEXTPROGRESS, SERVICEPRACTICE,Caseprogress " & _
                        " WHERE NP02 IN ('S','CFC') AND NP06 is null" & _
                        " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                        " AND NP10 in (" & stNumList & ") AND np07 NOT IN ('305','901')" & _
                        " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                        " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                        " AND SP15 IS NULL AND NP01=CP09(+) AND SUBSTR(CP12,1,2)='F1'" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
            cnnConnection.Execute strSql, intI
         End If
         'end 2022/11/21

         'Add By Sindy 2010/7/29 另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
               'Modify By Sindy 2015/7/30 FCT的208補優先權證明期限要剔除,獨立控管 + and NP02||NP07<>'FCT208'
               'Modify By Sindy 2015/8/21 FCT的901催款要剔除,獨立控管
               'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02 in('FCT','T') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103TSql & " AND np07 NOT IN ('305','901') and NP02||NP07<>'FCT208'" & _
                           " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:未收文
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, Lawcase" & _
                           " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103LSql & " AND np07 NOT IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                           " AND LC08 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, SERVICEPRACTICE" & _
                           " WHERE NP02='S' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate7 & _
                           " AND NP10 ='F4103'" & strF4103SSql & " AND np07 NOT IN ('305','901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                           " AND SP15 IS NULL" & _
                           " AND SP09='000'" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         
         'Add By Sindy 2015/8/21 FCT的901催款=>改控制為本所期限<=系統日+2個工作天
         'T:外商收的大陸案
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02 in ('FCT','T') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:未收文
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Lawcase, Staff" & _
                     " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901') AND NP10=ST01(+) AND substr(ST15,1,2)='F1'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, SERVICEPRACTICE" & _
                     " WHERE NP02='S' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                     " AND NP10 in (" & stNumList & ") AND np07 IN ('901')" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02 in('FCT','T') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103TSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2015/10/23 +法務:未收文
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, Lawcase" & _
                           " WHERE NP02 in('FCL','CFL','LIN','ACS') AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103LSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                           " AND LC08 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
               '2015/10/23 END
               
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, SERVICEPRACTICE" & _
                           " WHERE NP02='S' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08< " & stDate2 & _
                           " AND NP10 ='F4103'" & strF4103SSql & " AND np07 IN ('901')" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                           " AND SP15 IS NULL" & _
                           " AND SP09='000'" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/8/21 END
         
         'Add By Sindy 2015/7/30 FCT的208補優先權證明期限=>改控制為本所期限<=系統日+30日曆天
         'Modified by Lydia 2016/10/14 +  " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & "
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT " & IIf(Len(stNumList) > 10, "/*+ INDEX(NEXTPROGRESS IDXNP1008) */", "") & " '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, TradeMark" & _
                     " WHERE NP02||NP07='FCT208' AND NP06 is null" & _
                     " AND NP08>=" & stDate0 & " AND NP08<=" & DBDATE(DateAdd("d", 30, ChangeWStringToWDateString(strSrvDate(1)))) & _
                     " AND NP10 in (" & stNumList & ")" & _
                     " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '另外依F4103期限做區分
         If txtUsernum = "78011" Or txtUsernum = "80030" Then
               'Modified by Lydia 2023/05/16 已無第二期註冊費之案件and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')
               strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                        " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                           "From NEXTPROGRESS, TradeMark" & _
                           " WHERE NP02||NP07='FCT208' AND NP06 is null" & _
                           " AND NP08>=" & stDate0 & " AND NP08<=" & DBDATE(DateAdd("d", 30, ChangeWStringToWDateString(strSrvDate(1)))) & _
                           " AND NP10 ='F4103'" & strF4103TSql & _
                           " and decode(np02||np07,'T102',tm17,'FCT102',tm17,'Y')='Y'" & _
                           " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                           " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                           " AND TM29 IS NULL" & _
                           " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='3' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
               cnnConnection.Execute strSql, intI
         End If
         '2015/7/30 END
   End If
   
'***********************
''A' EV1,'1' EV2
'***********************
   '【FCT、S台灣案】
   '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE CP01='FCT'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2015/10/23 +法務:達本所
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Lawcase" & _
               " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2015/10/23 END
   
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S'" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2010/7/29 另外依F4103期限做區分
   If txtUsernum = "78011" Or txtUsernum = "80030" Then
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01='FCT'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103TSql & " AND CP158=0 AND CP159=0" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:達本所
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, Lawcase" & _
                     " WHERE CP01 in('FCL','CFL','LIN','ACS')" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103LSql & " AND CP158=0 AND CP159=0" & _
                     " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S'" & _
                     " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                     " AND CASEPROGRESS.CP14 ='F4103'" & strF4103SSql & " AND CP158=0 AND CP159=0" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
   End If
   
'***********************
''K' EV1,'1' EV2
'***********************
   'Add By Sindy 2021/6/3
   '需請款：S案之「查名」進度有承辦期限時，承辦期限當日提醒承辦人，事件為「需請款」，顏色為紫色
   'Modify By Sindy 2021/7/1:6/4上線有關查名案期限通知之需求，請改為以本所期限管控，即：
   'S案「查名」進度之本所期限當日提醒承辦人，事件為「需請款」，顏色為紫色
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','K' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01='S' AND CP10='001' AND CP16>0 AND CP60 IS NULL" & _
               " AND ((CASEPROGRESS.CP06 is not null AND CASEPROGRESS.CP06<=" & strSrvDate(1) & ") or (CASEPROGRESS.CP07 is not null AND CASEPROGRESS.CP07<=" & strSrvDate(1) & "))" & _
               " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP09='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='K' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2023/3/24 + FCT案的001.查名
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','K' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, Trademark" & _
               " WHERE CP01='FCT' AND CP10='001' AND CP16>0 AND CP60 IS NULL" & _
               " AND ((CASEPROGRESS.CP06 is not null AND CASEPROGRESS.CP06<=" & strSrvDate(1) & ") or (CASEPROGRESS.CP07 is not null AND CASEPROGRESS.CP07<=" & strSrvDate(1) & "))" & _
               " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM10='000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='K' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2021/6/3 END
   
'***********************
''D' EV1,'2' EV2
'***********************
   '程序組
   If Trim(stDept) = "F12" Then
         '未收文且 1個工作天 後達本所期限者(不含當日) --管制人-D2(未收文,管制人)
         'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
         'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
         'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'FCT716',tm17,'FCT102',tm17,'Y')=>decode(np02||np07,'FCT102',tm17,'Y')
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, TradeMark" & _
                     " WHERE NP02||NP06='FCT'" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND np07 NOT IN ('305')" & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and decode(np02||np07,'FCT102',tm17,'Y')='Y'" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2015/10/23 +法務:未收文
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, Lawcase" & _
                     " WHERE NP02||NP06 in('FCL','CFL','LIN','ACS')" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
                     " AND LC08 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         '2015/10/23 END
         
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,ST57 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                     "From NEXTPROGRESS, Staff, SERVICEPRACTICE" & _
                     " WHERE NP02||NP06='S'" & _
                     " AND NP08>=" & stDate0 & " AND NP08< " & stDate1 & _
                     " AND np07 NOT IN ('305')" & _
                     " AND NP10=ST01(+) AND ST57 in (" & stNumList & ")" & _
                     " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                     " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
         cnnConnection.Execute strSql, intI
         
'***********************
''J' EV1,'1' EV2
'***********************
         'Add By Sindy 2014/12/11 將FCT分案時輸入之承辦期限納入外商程序組之期限自動通知系統:J.今送件
         'Modified by Lydia 2016/10/14 CP01||CP27||CP57='FCT' => CP01='FCT' AND CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','J' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01='FCT' AND CP158=0 AND CP159=0" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48<=" & strSrvDate(1) & _
                     " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='J' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CP01||CP27||CP57='S' => CP01='S' AND CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','J' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01='S' AND CP158=0 AND CP159=0" & _
                     " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48<=" & strSrvDate(1) & _
                     " AND CASEPROGRESS.CP14 in (" & stNumList & ")" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND SP09='000'" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='J' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
   End If
   
'***********************
''H' EV1,'1' EV2
'***********************
   '【CFT、CFC、S非台灣案】
   '已收文未發文,5個工作天後達法定期限者(不含當日) --承辦人-H1(法定期限,承辦人)
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','H' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & _
               " AND CASEPROGRESS.CP07>=" & stDate0 & " AND CASEPROGRESS.CP07< " & stDate5 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','H' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP07>=" & stDate0 & " AND CASEPROGRESS.CP07< " & stDate5 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2023/12/11 +達指定
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & stConCP142 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29||TM57 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & stConCP142 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '達指定:未分案,管制人
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'2' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,NA69 NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark, Nation" & _
               " WHERE CP01='CFT'" & stConCP142 & _
               " AND CASEPROGRESS.CP14 is null AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='2' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','N' EV1,'2' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,NA69 NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE CP01 in('S','CFC')" & stConCP142 & _
               " AND CASEPROGRESS.CP14 is null AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='N' and EV2='2' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2023/12/11 END
   
   'Add By Sindy 2021/7/5
   '已收文未發文,1個工作天後達本所期限者(不含當日) --承辦人-H1(法定期限,承辦人)
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000')) " & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='H' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   '2021/7/5 END
      
   'Added By Lydia 2022/11/21
   '已收文未發文,承辦期限＜＝系統日＋１個工作天之未發文案件。 --達承辦 B1 (承辦期限,承辦人)
   'Modified by Lydia 2024/07/31 (公告1130828-01)+TF非台灣案;CP01='CFT'=>(CP01='CFT' OR (TM01='TF' AND TM10<>'000'))
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','B' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, TradeMark" & _
               " WHERE (CP01='CFT' OR (TM01='TF' AND TM10<>'000'))" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
               " AND TM29 IS NULL" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','B' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
               "From CASEPROGRESS, SERVICEPRACTICE" & _
               " WHERE CP01 in('S','CFC')" & _
               " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate1 & _
               " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
   cnnConnection.Execute strSql, intI
   'end 2022/11/21
   
'***********************
''D' EV1,'2' EV2
'***********************
   '未收文且 第5、6個工作天後達法定期限者(不含當日) --管制人-D2(未收文,管制人)
   'Modify By Sindy 2010/3/1 申請國家非日本時，管制人為國家檔CFT承辦人
   'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
   'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
   'Modified by Lydia 2016/03/11 國家檔的CFT承辦人(NA69)改成模組(DB.Functions)
   'Remove by Lydia 2016/03/24 外商反應未考慮好,先移除
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            "select * from ( SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,GETNA69(NP02,NP03,NP04,NP05,'','') NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 NOT IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10<>'011' AND TM10=NA01(+) AND TM10=NA01(+)" & _
               " and decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')='Y'" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)" & _
               " ) where na16 in (" & stNumList & ") "
   'Modified by Lydia 2016/11/15  TM10<>'011' => TM10 NOT IN (" & ExpNa01 & ")
   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
      'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 IN ('102')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/10/14 NP02||NP06 in ('S','CFC') => NP02 in ('S','CFC') AND NVL(NP06,'0')='0' ; + /*+ INDEX(NEXTPROGRESS IDXNP09020706) */
   'Modified by Lydia 2016/11/15  SP09<>'011' => SP09 NOT IN (" & ExpNa01 & ")
   'Modified by Lydia 2018/12/04 CFC不抓NA69
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(NEXTPROGRESS IDXNP09020706) */ '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 NOT IN ('305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
  'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
               ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=> SP15||SP61 IS NULL
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(NEXTPROGRESS IDXNP09020706) */ '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
               " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
               " AND np07 IN ('102')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2014/9/12 催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
   'Modified by Lydia 2016/03/11 國家檔的CFT承辦人(NA69)改成模組(DB.Functions)
   'Remove by Lydia 2016/03/24 外商反應未考慮好,先移除
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            "select * from ( SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,GETNA69(NP02,NP03,NP04,NP05,'','') NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL AND TM10<>'011' AND TM10=NA01(+)" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)" & _
               " ) where na16 in (" & stNumList & ") "
   'Modified by Lydia 2016/11/15 TM10<>'011' => TM10 NOT IN (" & ExpNa01 & ")
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/09/06
   'Modified by Lydia 2016/11/15  SP09<>'011' => SP09 NOT IN (" & ExpNa01 & ")
   'Modified by Lydia 2018/12/04 CFC不抓NA69
   'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND NP08<=" & strSrvDate(1) & _
               " AND np07 IN ('305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   '2014/9/12 END
   'Added by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE  NP02||NP06='TF' AND TM10<>'000' " & _
               " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/09/06
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, SERVICEPRACTICE, Nation" & _
               " WHERE NP02||NP06 in ('S','CFC')" & _
               " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
               " AND np07 NOT IN ('102','305','997','998')" & _
               " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
               " AND SP15||SP61 IS NULL" & _
               " AND SP09<>'000'" & _
               " AND SP09 NOT IN (" & ExpNa01 & ") AND SP09=NA01(+) AND NP10 in (" & stNumList & ")" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2022/11/21
   'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NA69 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
               "From NEXTPROGRESS, TradeMark, Nation" & _
               " WHERE NP02||NP06='CFT'" & _
               " AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) & _
               " AND np07 IN ('102','105')" & _
               " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
               " AND TM29||TM57 IS NULL" & _
               " AND TM10 NOT IN (" & ExpNa01 & ") AND TM10=NA01(+) AND NA69 in (" & stNumList & ")" & _
               " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
               " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
               " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
   cnnConnection.Execute strSql, intI
   'end 2024/07/31
   
   '申請國家為日本時，管制人(中南高所)78011葉易雲
   '                                                (北所)        98018蔡庭蓁
   'Modify By Sindy 2010/5/31
   '申請國家為日本時，管制人(中南高所)99011王婉
   '                                                (北所)        98018蔡庭蓁
   strSql = "select OMAN from SetSpecMan where OCODE='CFT_011A'"
   intI = 1: strUser = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Then
         strUser = Trim(RsTemp.Fields(0)) 'CFT承辦人日本(中南高所)管制人
      End If
   End If
   txtData = Split(stNumList, strUser)
   If UBound(txtData) = 1 Then
      strWhSql = " AND st06 in ('2','3','4')"
   Else
      strSql = "select OMAN from SetSpecMan where OCODE='CFT_011B'"
      intI = 1: strUser = ""
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields(0)) Then
            strUser = Trim(RsTemp.Fields(0)) 'CFT承辦人日本(北所)管制人
         End If
      End If
      txtData = Split(stNumList, strUser)
      If UBound(txtData) = 1 Then
         strWhSql = " AND st06 in ('1')"
      End If
   End If
   If UBound(txtData) = 1 Then
      'Modify By Sindy 2011/3/15 延展(102)和第二期(716)專用權須存在(TM17=Y)
      'Modify By Sindy 2012/3/7 未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
      'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
      'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
                  " AND np07 IN ('102')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2016/10/14 NP02||NP06 in ('S','CFC') => NP02 in ('S','CFC') AND NVL(NP06,'0')='0' ; + /*+ INDEX(NEXTPROGRESS IDXNP09020706) */
      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
      'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                  " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
                  " AND np07 NOT IN ('305','997','998')" & _
                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                  " AND SP15 IS NULL" & _
                  " AND substr(SP09,1,3)='011' " & _
                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      'cnnConnection.Execute strSql, intI
      
      'Add By Sindy 2014/9/12 催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND NP08<=" & strSrvDate(1) & _
                  " AND np07 IN ('305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29 IS NULL AND substr(TM10,1,3)='011'" & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
                  " AND NP08<=" & strSrvDate(1) & _
                  " AND np07 IN ('305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29 IS NULL AND substr(TM10,1,3)='011'" & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'end 2024/09/06
      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
      'strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                  " WHERE NP02||NP06 in ('S','CFC')" & _
                  " AND NP08<=" & strSrvDate(1) & _
                  " AND np07 IN ('305')" & _
                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                  " AND SP15 IS NULL" & _
                  " AND substr(SP09,1,3)='011' " & _
                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      'cnnConnection.Execute strSql, intI
      '2014/9/12 END
      'Added by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
                  " AND np07 NOT IN ('102','305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'END 2022/11/21
      'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) & _
                  " AND np07 IN ('102','105')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND substr(TM10,1,3)='011'" & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+) " & strWhSql & _
                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'end 2024/07/31
   End If
   '2010/3/1 End
   'Added by Lydia 2018/12/04 CFC不判斷特殊設定
   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=>SP15||SP61 IS NULL
    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
                " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
                " AND np07 IN ('102')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15||SP61 IS NULL" & _
                " AND substr(SP09,1,3)='011' " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02||NP06 in ('S','CFC')" & _
                " AND NP08<=" & strSrvDate(1) & _
                " AND np07 IN ('305','997','998')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15 IS NULL" & _
                " AND substr(SP09,1,3)='011' " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'end 2018/12/04
    'Added by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02||NP06 in ('S','CFC')" & _
                " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
                " AND np07 NOT IN ('102','305','997','998')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15||SP61 IS NULL" & _
                " AND substr(SP09,1,3)='011' " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'end 2022/11/21
    
   'Added by Lydia 2016/11/16 美國和歐盟案件依所別區分管制人
   strUser = Pub_GetSpecMan("CFT_101239A") '南高所
   txtData = Split(stNumList, strUser)
   'Modified by Lydia 2018/10/03 分成CFT_101239B(北所)和CFT_101239C(中所)兩個設定
'   If UBound(txtData) = 1 Then
'      strWhSql = " AND st06 in ('3','4')"
'   Else
'      strUser = Pub_GetSpecMan("CFT_101239B") '北中所
'      txtData = Split(stNumList, strUser)
'      If UBound(txtData) = 1 Then
'         strWhSql = " AND st06 in ('1','2')"
'      End If
'   End If
'   If UBound(txtData) = 1 Then
   strExc(1) = ""
   If UBound(txtData) = 1 Then
        strExc(1) = strExc(1) & "3,4,"
   End If
   strUser = Pub_GetSpecMan("CFT_101239B") '北所
   txtData = Split(stNumList, strUser)
   If UBound(txtData) = 1 Then
        strExc(1) = strExc(1) & "1,"
   End If
   strUser = Pub_GetSpecMan("CFT_101239C") '中所
   txtData = Split(stNumList, strUser)
   If UBound(txtData) = 1 Then
        strExc(1) = strExc(1) & "2,"
   End If
   If strExc(1) <> "" Then
      strWhSql = " AND st06 in (" & GetAddStr(strExc(1)) & ")"
'end 2018/10/03
      '延展(102)和第二期(716)專用權須存在(TM17=Y)
      '未收文增加過濾進度檔中若有該筆相關總收文號且未發文未取消收文者,不出現該筆未收文
      'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), TM29 IS NULL=> TM29||TM57 IS NULL
      'Modified by Lydia 2023/05/16 已無第二期註冊費之案件decode(np02||np07,'CFT716',tm17,'CFT102',tm17,'Y')=>decode(np02||np07,'CFT102',tm17,'Y')
      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
                  " AND np07 IN ('102')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                  " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
'                  " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
'                  " AND np07 NOT IN ('305','997','998')" & _
'                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                  " AND SP15 IS NULL" & _
'                  " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
'                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
      
      '催審期限另外抓,為本所期限當日開始顯示,至催審期限取消為止
      'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT' " & _
                  " AND NP08<=" & strSrvDate(1) & _
                  " AND np07 IN ('305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2024/09/06 (公告1130828-01)CFT承辦人承辦之TF案件
      'Modified by Lydia 2025/05/05 Debug: (修改前)NP02||NP06='CFT' -> NP02||NP06='TF'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='TF' AND TM10<>'000' " & _
                  " AND NP08<=" & strSrvDate(1) & _
                  " AND np07 IN ('305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'end 2024/09/06
      'Remove by Lydia 2018/12/04 CFC不判斷特殊設定
'      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
'                  "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
'                  " WHERE NP02||NP06 in ('S','CFC')" & _
'                  " AND NP08<=" & strSrvDate(1) & _
'                  " AND np07 IN ('305')" & _
'                  " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
'                  " AND SP15 IS NULL" & _
'                  " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
'                  " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
'                  " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
'                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
'                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
'                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
'      cnnConnection.Execute strSql, intI
      'Added by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
                  " AND np07 NOT IN ('102','305','997','998')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+)" & strWhSql & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'END 2022/11/21
      'Added by Lydia 2024/07/31 (公告1130828-01)延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件，並僅彈跳兩天提醒。
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,'" & strUser & "' NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                  "From NEXTPROGRESS, TradeMark,caseprogress,staff" & _
                  " WHERE NP02||NP06='CFT'" & _
                  " AND NP09>" & CompWorkDay(3, CompDate(1, 1, strSrvDate(1)), 1) & " AND NP09<= " & CompDate(1, 1, strSrvDate(1)) & _
                  " AND np07 IN ('102','105')" & _
                  " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05" & _
                  " AND TM29||TM57 IS NULL AND (substr(TM10,1,3)='101' OR TM10='239') " & _
                  " AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+)" & _
                  " AND CP05 in (select MAX(cp05) from caseprogress where TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04)" & _
                  " AND caseprogress.CP13=ST01(+) " & strWhSql & _
                  " and decode(np02||np07,'CFT102',tm17,'Y')='Y'" & _
                  " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
      cnnConnection.Execute strSql, intI
      'end 2024/07/31
   End If
    'Added by Lydia 2018/12/04 CFC不判斷特殊設定
   'Modified by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
              ' np07 NOT IN ('305','997','998')=> np07 IN ('102','997','998'), SP15 IS NULL=>SP15||SP61 IS NULL
    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('102','997','998')=> np07 IN ('102')
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02 in ('S','CFC') AND NVL(NP06,'0')='0'" & _
                " AND NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & _
                " AND np07 IN ('102')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15||SP61 IS NULL" & _
                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'Modified by Lydia 2024/08/06 (公告1130828-01)收達(997)/提申(998)改成和催審一樣 by Alice ; np07 IN ('305')=> np07 IN ('305','997','998')
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02||NP06 in ('S','CFC')" & _
                " AND NP08<=" & strSrvDate(1) & _
                " AND np07 IN ('305','997','998')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15 IS NULL" & _
                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'end 2018/12/04
   'end 2016/11/16
    'Added by Lydia 2022/11/21 延展102仍保持只跳二天(系統日＋５或６個工作天)，目前所有中間程序只跳二天，修改為持續跳到智權同仁有收文接洽單或填結案單為止。
    strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
             " SELECT distinct '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(NP08,0),nvl(NP09,0),0 CP48,0 EP08,NP10 NA16,NULL CP14,NP10,NEXTPROGRESS.NP22 " & _
                "From NEXTPROGRESS, SERVICEPRACTICE,caseprogress,staff" & _
                " WHERE NP02||NP06 in ('S','CFC')" & _
                " AND ((NP09>=" & CompWorkDay(6, strSrvDate(1)) & " AND NP09< " & stDate6 & ") or NP08<=" & strSrvDate(1) & ")" & _
                " AND np07 NOT IN ('102','305','997','998')" & _
                " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05" & _
                " AND SP15||SP61 IS NULL" & _
                " AND (substr(SP09,1,3)='101' OR SP09='239') " & _
                " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+)" & _
                " AND CP05 in (select MAX(cp05) from caseprogress where SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04)" & _
                " AND caseprogress.CP13=ST01(+) AND NP10 in (" & stNumList & ")" & _
                " and not exists (select * from caseprogress where cp43=np01 and cp27 is null and cp57 is null)" & _
                " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='D' and EV2='2' and CP09=NP01 and R030301.NP22=NEXTPROGRESS.NP22)"
    cnnConnection.Execute strSql, intI
    'end 2022/11/21
    
'***********************
''E' EV1,'1' EV2
'***********************
   '所有未發文--承辦人-E(未發文) -- 含T案
   If idx = 1 Then
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('T','FCT','CFT') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Add By Sindy 2015/10/23 +法務:未發文
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, Lawcase" & _
                  " WHERE CP01 in('FCL','CFL','LIN','ACS') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                  " AND LC08 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      '2015/10/23 END
      
      'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                  " AND CASEPROGRESS.CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
      'Add By Sindy 2010/7/29 另外依F4103期限做區分
      If txtUsernum = "78011" Or txtUsernum = "80030" Then
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, TradeMark" & _
                        " WHERE CP01 in('T','FCT','CFT') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103TSql & " AND CP158=0 AND CP159=0" & _
                        " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                        " AND TM29 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
            'Add By Sindy 2015/10/23 +法務:未發文
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, Lawcase" & _
                        " WHERE CP01 in('FCL','CFL','LIN','ACS') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103LSql & " AND CP158=0 AND CP159=0" & _
                        " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                        " AND LC08 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
            '2015/10/23 END
            
            'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
            strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                        "From CASEPROGRESS, SERVICEPRACTICE" & _
                        " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                        " AND CASEPROGRESS.CP14 ='F4103'" & strF4103SSql & " AND CP158=0 AND CP159=0" & _
                        " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                        " AND SP15 IS NULL" & _
                        " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='1' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
            cnnConnection.Execute strSql, intI
      End If
   End If
   
'***********************
''A' EV1,'0' EV2
'***********************
   '未分案-0 -- 不含T案
   If bLvl4 = True Or bLvl5 = True Then
      '已收文未發,2個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('FCT','CFT')" & _
                  " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC')" & _
                  " AND CASEPROGRESS.CP06>=" & stDate0 & " AND CASEPROGRESS.CP06< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='A' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
'***********************
''B' EV1,'0' EV2
'***********************
      '已收文未發,2個工作天後達承辦期限者(不含當日)-B0(承辦期限,未分案)
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, TradeMark" & _
                  " WHERE CP01 in('FCT','CFT')" & _
                  " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                  " AND TM29 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
      strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                  "From CASEPROGRESS, SERVICEPRACTICE" & _
                  " WHERE CP01 in('S','CFC')" & _
                  " AND CASEPROGRESS.CP48>=" & stDate0 & " AND CASEPROGRESS.CP48< " & stDate2 & _
                  " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                  " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                  " AND SP15 IS NULL" & _
                  " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='B' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
      cnnConnection.Execute strSql, intI
      
'***********************
''E' EV1,'0' EV2
'***********************
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
         'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, TradeMark" & _
                     " WHERE CP01 in('FCT','CFT') AND CP05>20030000" & _
                     " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                     " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04" & _
                     " AND TM29 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
         'Modified by Lydia 2016/10/14 CASEPROGRESS.CP14||CP27||CP57 IS NULL => CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'
         'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (CASEPROGRESS.cp06||CASEPROGRESS.cp48 is null or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48 is null) or (CASEPROGRESS.cp06 is null and CASEPROGRESS.cp48>=" & stDate1 & ") or (CASEPROGRESS.cp06>=" & stDate1 & " and CASEPROGRESS.cp48>=" & stDate1 & "))"
         strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CASEPROGRESS.CP09,nvl(CASEPROGRESS.CP06,0),nvl(CASEPROGRESS.CP07,0),nvl(CASEPROGRESS.CP48,0),0 EP08,'' NA16,CASEPROGRESS.CP14,CASEPROGRESS.CP13,0 " & _
                     "From CASEPROGRESS, SERVICEPRACTICE" & _
                     " WHERE CP01 in('S','CFC') AND CP05>20030000" & _
                     " AND CP158=0 AND CP159=0 AND NVL(CP14,'0')='0'" & _
                     " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
                     " AND SP15 IS NULL" & _
                     " AND not exists (select * from R030301 where ID='" & strUserNum & "' and EV1='E' and EV2='0' and R030301.CP09=CASEPROGRESS.CP09 and np22=0)"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
'***********************
''B' EV1,'1' EV2
'***********************
   'Add By Sindy 2012/6/4
   '【T、FCT台灣商標爭議案逾承辦期限、逾指定會稿日】
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, Trademark, EngineerProgress" & _
               " WHERE CP05>=20120601" & _
               " AND CP01 in('T','FCT')" & _
               " AND CP10 in(" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)" & _
               " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
               " AND TM10='000' and tm29 is null" & _
               " AND CP09=EP02(+)" & _
               " AND (CP48<" & strSrvDate(1) & " and CP48 is not null)"
   cnnConnection.Execute strSql, intI
   'Added by Lydia 2018/12/10 +T台灣案非爭議案
   'Remove by Lydia 2019/01/30
'   If strSrvDate(1) >= T案收文齊備啟用日 Then
'        strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'                 " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','B' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
'                    "From CASEPROGRESS, Trademark, EngineerProgress" & _
'                    " WHERE CP05>=" & T案收文齊備啟用日 & _
'                    " AND CP01 ='T'" & _
'                    " AND CP10 not in(" & TMdebate & ")" & _
'                    " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0 AND CP09 LIKE 'A%' " & _
'                    " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
'                    " AND TM10='000' and tm29 is null" & _
'                    " AND CP09=EP02(+)" & _
'                    " AND (CP48<" & strSrvDate(1) & " and CP48 is not null)"
'        cnnConnection.Execute strSql, intI
'   End If
   'end 2018/12/10
'***********************
''I' EV1,'1' EV2
'***********************
   'Modified by Lydia 2016/10/14 CP57 is null AND CP27 is null=> CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','I' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, Trademark, EngineerProgress" & _
               " WHERE CP05>=20120601" & _
               " AND CP01 in('T','FCT')" & _
               " AND CP10 in(" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)" & _
               " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0" & _
               " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
               " AND TM10='000' and tm29 is null" & _
               " AND CP09=EP02(+)" & _
               " AND (EP28<" & strSrvDate(1) & " and EP28 is not null)"
   cnnConnection.Execute strSql, intI
   '2012/6/4 End
   'Added by Lydia 2018/12/10 +T台灣案非爭議案
   'Remove by Lydia 2019/01/30
'   If strSrvDate(1) >= T案收文齊備啟用日 Then
'        strSql = "INSERT INTO R030301(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
'                 " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','I' EV1,'1' EV2,CP09,nvl(CP06,0),nvl(CP07,0),nvl(CP48,0),0 EP08,'' NA16,CP14,CP13,0 " & _
'                    "From CASEPROGRESS, Trademark, EngineerProgress" & _
'                    " WHERE CP05>=" & T案收文齊備啟用日 & _
'                    " AND CP01 ='T'" & _
'                    " AND CP10 not in(" & TMdebate & ")" & _
'                    " AND CP14 in(" & stNumList & ") AND CP158=0 AND CP159=0 AND CP09 LIKE 'A%' " & _
'                    " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
'                    " AND TM10='000' and tm29 is null" & _
'                    " AND CP09=EP02(+)" & _
'                    " AND (EP28<" & strSrvDate(1) & " and EP28 is not null)"
'        cnnConnection.Execute strSql, intI
'   End If
   'end 2018/12/10
   
   'Add By Sindy 2015/2/17 若同案有程序的J.今送件及A.達本所期限,就不要再顯示達本所
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='A' and R1.EV2='1'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='J' and R2.EV2='1' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '2015/2/17 END
   
   'Add By Sindy 2017/9/11
   '若同案有 'H達法定'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'H'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='H' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '若同案有 'A達本所'期限,其他的就不要再顯示
   strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'A'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='A' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   '2017/9/11 END
   
   'Add By Sindy 2015/5/7 若為程序組人員時,只有程序主管需顯示未收文之延展案件
   'Modify By Sindy 2020/7/2 程序不管制未收文延展期限
   '程序組
   'If Trim(stDept) = "F12" And stUserID <> Pub_GetSpecMan("P_FCT") Then
   If Trim(stDept) = "F12" Then
      strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='D'" & _
               " AND R1.ID||R1.EV1||R1.EV2||R1.CP09||R1.NP22 in (select R2.ID||R2.EV1||R2.EV2||R2.CP09||R2.NP22 from R030301 R2,nextprogress where R2.ID='" & strUserNum & "' and R2.EV1='D' and R2.CP09=nextprogress.NP01 and R2.np22=nextprogress.NP22 and nextprogress.NP07='102')"
      cnnConnection.Execute strSql, intI
   Else
      'Modify By Sindy 2020/7/2 未收文延展期限，英文組改為法定期限＋１個工作天＋7個月才開始提醒，日文組維持原規則。
      strSql = "delete from R030301 R1 where R1.ID='" & strUserNum & "' and R1.EV1='D'" & _
               " AND R1.ID||R1.EV1||R1.EV2||R1.CP09||R1.NP22 in (" & _
               "select R2.ID||R2.EV1||R2.EV2||R2.CP09||R2.NP22" & _
               " from R030301 R2,nextprogress,trademark,fagent" & _
               " where R2.ID='" & strUserNum & "' and R2.EV1='D' and R2.CP09=nextprogress.NP01 and R2.np22=nextprogress.NP22 and nextprogress.NP07='102'" & _
               " AND NP02=TM01 AND NP03=TM02 AND NP04=TM03 AND NP05=TM04" & _
               " AND substr(TM44,1,8)=fa01(+) AND substr(TM44,9,1)=fa02(+)" & _
               " AND substr(fa10,1,3)<>'011' AND fa10 IS NOT NULL" & _
               " AND WORKDAYADD(+1,to_char(add_months(to_date(np09,'YYYYMMDD'),7),'YYYYMMDD'))>" & strSrvDate(1) & _
               ")"
      cnnConnection.Execute strSql, intI
   End If
   '2015/5/7 END
   
   '案件進度
   'Modify By Sindy 2021/6/3 + ,'K','需請款'
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   'Modify By Sindy 2023/12/11 + ,'N','達指定'
   strExc(0) = "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "TM05 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,CP10,CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04" & _
      " FROM R030301,trademark,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   'Add By Sindy 2015/10/23
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(LC15,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "NVL(LC05,NVL(LC06,LC07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,LC01,LC02,LC03,LC04,'' 未收款,CP10,CP27,R030301.CP09,LC15,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04" & _
      " FROM R030301,Lawcase,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/10/23 END
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,CP10,CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04" & _
      " FROM R030301,SERVICEPRACTICE,caseprogress,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '下一程序
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',TM10),0,CPM03,CPM04) => DECODE(TM10,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(TM10,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "TM05 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,TM01,TM02,TM03,TM04,'' 未收款,NP07,0 CP27,R030301.CP09,TM10,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "TM01||'-'||TM02||'-'||TM03||'-'||TM04" & _
      " FROM R030301,trademark,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05 AND TM01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   'Add By Sindy 2015/10/23
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(LC15,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "NVL(LC05,NVL(LC06,LC07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,LC01,LC02,LC03,LC04,'' 未收款,NP07,0 CP27,R030301.CP09,LC15,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "LC01||'-'||LC02||'-'||LC03||'-'||LC04" & _
      " FROM R030301,Lawcase,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/10/23 END
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
               "SELECT '' V," & _
      "NVL(lpad(SQLDateT(R030301.CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(decode(R030301.CP07,0,null,R030301.CP07)),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(R030301.CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(R030301.EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "S2.ST02 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','達法定','I','達指會','J','今送件','K','需請款','N','達指定') 事件," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04 本所案號," & _
      "decode(ti01,null,' ','@')||DECODE(SP09,'000',CPM03,CPM04) 案件性質," & _
      "NP15 案件備註," & _
      "NVL(SP05,NVL(SP06,SP07)) 案件名稱," & _
      "'' 代理人國籍," & _
      "EV1,EV2,R030301.NA16,R030301.CP14,R030301.CP13,SP01,SP02,SP03,SP04,'' 未收款,NP07,0 CP27,R030301.CP09,SP09,decode(ti01,null,' ','@') ti01,R030301.NP22,decode(R030301.CP06,0,99999999,R030301.CP06) sort," & _
      "SP01||'-'||SP02||'-'||SP03||'-'||SP04" & _
      " FROM R030301,SERVICEPRACTICE,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,T102InForm" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R030301.NA16" & _
      " AND S2.ST01(+)=R030301.CP14" & _
      " AND S3.ST01(+)=R030301.CP13" & _
      " AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND R030301.CP09=ti02(+) and R030301.np22=ti04(+)"
   '2015/2/12 END
   'Modify By Sindy 2015/1/30
'   strExc(0) = "SELECT V,本所期限,法定期限,承辦期限,核稿期限,管制人,承辦人,智權人員" & _
'                     ",事件,本所案號,案件性質,案件備註,案件名稱,代理人國籍," & _
'                     "EV1,EV2,NA16,CP14,CP13,TM01,TM02,TM03,TM04,未收款,CP10,CP27,CP09,TM10,ti01,NP22,CP06 from (" & _
'               strExc(0) & ")"
   Select Case stDept
      Case "F11", "F10" '承辦組
         'strExc(0) = strExc(0) & " order by 本所期限 asc,智權人員 asc,本所案號 asc"
         strExc(0) = strExc(0) & " order by sort asc,智權人員 asc,本所案號 asc"
      'Case "F12" '程序組
      Case Else
         'strExc(0) = strExc(0) & " order by 本所期限 asc,承辦人 asc,本所案號 asc"
         strExc(0) = strExc(0) & " order by sort asc,承辦人 asc,本所案號 asc"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grdDataList.Recordset = rsTmp
      SetXRecord '更新欄位值
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   Else
      Screen.MousePointer = vbDefault
      MsgBox "查無資料！", vbInformation
      rsTmp.Close
      Set rsTmp = Nothing
      cmdHide.Enabled = False
      lblCnt.Caption = "共 0 筆"
      Exit Sub 'Add By Sindy 2014/9/17
   End If
   rsTmp.Close
   
   'Modify By Sindy 2014/9/17 案件性質+相關總收文號的案件性質
   For iRow = 1 To grdDataList.Rows - 1
      grdDataList.TextMatrix(iRow, 10) = grdDataList.TextMatrix(iRow, 10) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(iRow, 26), grdDataList.TextMatrix(iRow, 29), "1")
   Next iRow
   '2014/9/17 END
End Sub

