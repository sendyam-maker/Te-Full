VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090608 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人達成情形查詢"
   ClientHeight    =   4920
   ClientLeft      =   2580
   ClientTop       =   1536
   ClientWidth     =   5232
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5232
   Begin VB.CheckBox Check3 
      Caption         =   "含離職人員"
      Height          =   252
      Left            =   3504
      TabIndex        =   31
      Top             =   4080
      Width           =   1500
   End
   Begin VB.CheckBox Check2 
      Caption         =   "顯示加分基數分離欄位"
      Height          =   255
      Left            =   180
      TabIndex        =   29
      Top             =   4080
      Width           =   2400
   End
   Begin VB.CheckBox Check1 
      Caption         =   " 只統計實際工作的件數　　　　　     　　　　　　　  (不含加乘,支援,業務身份收文,法務分配等加分基數)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   180
      TabIndex        =   13
      Top             =   4380
      Width           =   4860
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3045
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1950
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1950
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2355
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1665
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2355
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   2
      Top             =   855
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   510
      Width           =   1920
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   855
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1215
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2220
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1215
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1575
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2715
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3405
      Width           =   330
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3165
      TabIndex        =   14
      Top             =   45
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3930
      TabIndex        =   15
      Top             =   45
      Width           =   1200
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   30
      Top             =   1575
      Width           =   1500
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "5. 完稿基數%)"
      Height          =   195
      Left            =   1560
      TabIndex        =   28
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "統計方式：         (1.新制  2.舊制  3.點數)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   11
      Left            =   180
      TabIndex        =   27
      Top             =   3048
      Width           =   3600
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   2835
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "部門別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   26
      Top             =   1950
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "EX：(9001-9012)"
      Height          =   195
      Left            =   3180
      TabIndex        =   25
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2100
      TabIndex        =   24
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   23
      Top             =   2355
      Width           =   1005
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1920
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line1 
      X1              =   1770
      X2              =   2655
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   21
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   19
      Top             =   1575
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式：          (1.螢幕  2.報表)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   180
      TabIndex        =   18
      Top             =   2712
      Width           =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   17
      Top             =   3405
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發文點數 % 2.發文基數 % 3.發文平均 % 4.承辦人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   10
      Left            =   1575
      TabIndex        =   16
      Top             =   3405
      Width           =   2400
   End
   Begin VB.Line Line2 
      X1              =   1740
      X2              =   2895
      Y1              =   1350
      Y2              =   1350
   End
End
Attribute VB_Name = "frm090608"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify by Morgan 2010/12/30 新舊制選項代碼對調
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String
Dim StrSQL7 As String, strSQL8 As String, strSQL9 As String, strSQL10 As String, strSQL11 As String, strSQL12 As String

Dim i As Integer, j As Integer, k As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String
Dim iPrint As Integer
Dim Page As Integer, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 7) As String
'Modifie by Morgan 2019/3/22
'Dim StrTemp7(0 To 13) As String
'Dim strTemp(0 To 25) As String
'Dim PLeft(0 To 13) As Integer
Dim StrTemp7(0 To 15) As String
Dim strTemp(0 To 27) As String
Dim PLeft(0 To 16) As Integer
'end 2019/3/22
Dim strTemp1 As Variant, strTemp2 As Variant, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim bol911001checkRange As Boolean
Public m_bolShowMemo As Boolean 'Added by Morgan 2014/9/24
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
Dim m_bol108Rule As Boolean 'Added by Morgan 2019/3/22 專利處108考核
Dim m_lngRptLineEnd As Long

Private Sub Check1_Click()
   'Added by Morgan 2019/1/4
   If Check1.Value = 1 Then
      Check2.Value = 0
      Check2.Enabled = False
   Else
      Check2.Enabled = True
   End If
   'end 2019/1/4
End Sub

Private Sub cmdok_Click(Index As Integer)

m_bolShowMemo = False 'Added by Morgan 2014/9/24
Select Case Index
Case 0 '確定
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         If PUB_CheckKeyInYYMM(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         'add by nickc 2005/03/04  加入判斷新舊制
         If Len(txt1(12)) = 0 Then
             s = MsgBox("統計方式不可空白!!", , "USER 輸入錯誤")
             If Len(txt1(12)) = 0 Then txt1(12).SetFocus
             Exit Sub
         '2012/1/20 ADD BY SONIA
         Else
            txt1_LostFocus (12)
         '2012/1/20 END
         End If
         If Len(txt1(3)) = 0 Or Len(txt1(4)) = 0 Then
             s = MsgBox("發文年月區間不可空白!!", , "USER 輸入錯誤")
             'If Len(txt1(4)) = 0 Then txt1(4).SetFocus
             If Len(txt1(3)) = 0 Then txt1(3).SetFocus
             Exit Sub
         Else
             If Len(txt1(8)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 txt1(8).SetFocus
                 Exit Sub
             Else
                 If Len(txt1(9)) = 0 Then
                     s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                     txt1(9).SetFocus
                     Exit Sub
                 Else
                    'Add By Cheng 2003/06/11
                    '檢查部門別範圍
                    If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
                        If Me.txt1(10).Text > Me.txt1(11).Text Then
                            MsgBox "部門別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                            Me.txt1(10).SetFocus
                            txt1_GotFocus 10
                            Exit Sub
                        End If
                    End If
                    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
                     Screen.MousePointer = vbHourglass
                     Me.Enabled = False
                     
                     If Not (txt1(12) = "1" And Check1.Value = vbUnchecked) Then Me.Check2.Value = 0 'Addd by Morgan 2018/12/11
                     
                     m_bol108Rule = False 'Added by Morgan 2019/3/22
                     If txt1(12) = "2" Then
                        pub_QL05 = pub_QL05 & ";" & Label1(11) & "2.舊制" 'Add By Sindy 2010/12/14
                        Process
                     
                     'Added by Morgan 2025/7/10
                     ElseIf txt1(12) = "3" Then
                        pub_QL05 = pub_QL05 & ";" & Label1(11) & "3.點數"
                        m_bol108Rule = True
                        ProcessNew4
                     'end 2025/7/10
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(11) & "1.新制" 'Add By Sindy 2010/12/14
                        'Added by Morgan 2019/3/22 專利處108考核
                        If Left(txt1(10), 2) = "P1" And DBDATE(txt1(3) & "01") >= PUB_108RuleDate Then
                           m_bol108Rule = True
                           ProcessNew3
                        Else
                        'end 2019/3/22
                        
                           ProcessNew2
                           
                        End If 'Added by Morgan 2019/3/22
                     End If
                     If Val(txt1(8)) = 1 Then
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.螢幕" 'Add By Sindy 2010/12/14
                        Me.Hide
                        frm090608_1.m_bol108Rule = m_bol108Rule 'Added by Morgan 2019/3/22
                        frm090608_1.Check1 = Me.Check2
                        frm090608_1.Show
                        Screen.MousePointer = vbDefault
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.報表" 'Add By Sindy 2010/12/14
                        PrintData
                     End If
                     Me.Enabled = True
                     Screen.MousePointer = vbDefault
                 End If
             End If
         End If
     End If
Case 1 '回前畫面
     Unload Me
Case Else
End Select
End Sub

Sub PrintData()
'取小數兩位
strSql = "select R102001,nvl(st02,r102002), Round(sum(r102003),2), Round(sum(r102004),2), Round(sum(r102005),2), Round(sum(r102006),2), Round(sum(r102007),2), Round(sum(r102008),2), Round(sum(r102009),2), Round(sum(r102010),2), Round(sum(r102011),2), Round(sum(r102012),2), Round(sum(r102013),2), Round(sum(r102014),2),r102002, ST06, ST03"
'Added by Morgan 2018/12/11
If Check2.Value = 1 Then
   strSql = strSql & ",round(sum(DECODE(r102015,'1',nvl(r102016,0),0)),2) X01,round(sum(DECODE(r102015,'1',nvl(r102017,0),0)),2) X02,round(sum(DECODE(r102015,'1',nvl(r102006,0),0)),2) X03,round(sum(DECODE(r102015,'2',nvl(r102006,0),0)),2) X04" & _
      ",round(sum(DECODE(r102015,'4',nvl(r102006,0),0)),2) X05,round(sum(DECODE(r102015,'1',nvl(r102018,0),0)),2) X06" & _
      ",round(sum(DECODE(r102015,'1',nvl(r102019,0),0)),2) X07,round(sum(DECODE(r102015,'1',nvl(r102020,0),0)),2) X08,round(sum(DECODE(r102015,'1',nvl(r102011,0),0)),2) X09,round(sum(DECODE(r102015,'2',nvl(r102011,0),0)),2) X10" & _
      ",round(sum(DECODE(r102015,'4',nvl(r102011,0),0)),2) X11,round(sum(DECODE(r102015,'1',nvl(r102021,0),0)),2) X12"
End If
'end 2018/12/11

'Added by Morgan 2019/3/25 108考核
If m_bol108Rule Then
   strSql = strSql & ", Round(sum(r102022),2) X13, Round(sum(r102023),2) X14"
End If
'end 2019/3/25

strSql = strSql & " from r090608,staff WHERE ID='" & strUserNum & "' and r102002=st01(+) group by r102001,r102002,nvl(st02,r102002), ST06, ST03 "
'排除皆為0的資料
strSql = strSql & " Having (Nvl(Round(sum(r102003),2),0)+Nvl(Round(sum(r102004),2),0)+Nvl(Round(sum(r102005),2),0)+Nvl(Round(sum(r102006),2),0)+Nvl(Round(sum(r102007),2),0)+Nvl(Round(sum(r102008),2),0)+Nvl(Round(sum(r102009),2),0)+Nvl(Round(sum(r102010),2),0)+Nvl(Round(sum(r102011),2),0)+Nvl(Round(sum(r102012),2),0)+Nvl(Round(sum(r102013),2),0)+Nvl(Round(sum(r102014),2),0))>0 "
'End
Select Case Val(frm090608.txt1(9))
Case 1
     pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "1.發文點數 %" 'Add By Sindy 2010/12/14
     strSql = strSql + " ORDER BY r102001,sum(R102007) Desc "
Case 2
     pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "2.發文基數 %" 'Add By Sindy 2010/12/14
     strSql = strSql + " ORDER BY R102001,sum(R102008) Desc "
Case 3
     pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "3.發文平均 %" 'Add By Sindy 2010/12/14
     strSql = strSql + " ORDER BY R102001,sum(R102009) Desc "
Case 4
      'Modify By Cheng 2003/06/11
'     strSQL = strSQL + " ORDER BY R102001,R102002 "
     pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "4.承辦人" 'Add By Sindy 2010/12/14
     strSql = strSql + " ORDER BY R102001, ST06, ST03, R102002 "
'Added by Lydia 2016/12/19 +完稿基數%
Case 5
     pub_QL05 = pub_QL05 & ";" & frm090608.Label1(6) & "5.完稿基數 %"
     strSql = strSql + " ORDER BY R102001,sum(R102013) Desc "
Case Else
End Select
CheckOC
Page = 1
SavDay1 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    '若有資料
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/14
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 13
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            
            'Added by Morgan 2018/12/11
            If Check2.Value = 1 Then
               For i = 14 To 25
                   strTemp(i) = CheckStr(.Fields(i + 3))
               Next i
            End If
            'end2018/12/11
            
            'Added by Morgan 2019/3/25
            If m_bol108Rule Then
               strTemp(26) = CheckStr(.Fields("X13"))
               strTemp(27) = CheckStr(.Fields("X14"))
            End If
            'end 2019/3/25
            
            If SavDay1 <> strTemp(0) Then
                ShowLine
                'Added by Morgan 2019/3/25
                If m_bol108Rule Then
                  PrintEnd3
                Else
                'end 2019/3/25
                
                  PrintEnd
                  
                End If 'Added by Morgan 2019/3/25
                
                ShowLine
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(0)
                PrintTitle
            End If
            
            'Added by Morgan 2019/3/25
            If m_bol108Rule Then
               PrintDatil3
            Else
            'end 2019/3/25
            
               PrintDatil
               
            End If 'Added by Morgan 2019/3/25
            
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    '若無資料
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/14
        CheckOC
        ShowNoData
        Exit Sub
    End If
End With
CheckOC
ShowLine
'Added by Morgan 2019/3/25
If m_bol108Rule Then
   PrintEnd3
Else
'end 2019/3/25
   PrintEnd
End If 'Added by Morgan 2019/3/25

ShowLine
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle() '列印抬頭

If m_bol108Rule Then PrintTitle3: Exit Sub 'Added by Morgan 2019/3/25

Dim strColName As String
'Modify by Morgan 2010/10/19
'Modified by Morgan 2013/5/22 新制未選實際工作件數的才用基數
If bolNewPromoterRule And txt1(12) = "1" And frm090608.Check1.Value = vbUnchecked Then
   'Modify by Morgan 2010/12/30 計點->基數
   strColName = "基數"
Else
   strColName = "件數"
End If

GetPleft
iPrint = 0
Printer.Orientation = vbPRORLandscape
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "承辦人達成情形統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
'Modify by Morgan 2011/2/16 修正百年問題
'Printer.Print "統計年月：" & Mid(Txt1(3), 1, 2) & "/" & Mid(Txt1(3), 3, 2) & "－" & Mid(Txt1(4), 1, 2) & "/" & Mid(Txt1(4), 3, 2)
Printer.Print "統計年月：" & (Val(txt1(3)) \ 100) & "/" & Right(txt1(3), 2) & "－" & (Val(txt1(4)) \ 100) & "/" & Right(txt1(4), 2)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modified by Morgan 2014/9/24
'   Printer.Print "系統類別：" & SavDay1
Printer.Print "系統類別：" & SavDay1 & IIf(m_bolShowMemo And SavDay1 = "ALL", "　※此頁之目標基數以王副總的設定值為主, 可能與P 與 CFP 的分別的目標基數計算合計有誤差 !", "")
'end 2014/9/24
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(m_lngRptLineEnd, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Underline = True
Printer.CurrentX = (PLeft(2) + (PLeft(3) - PLeft(2) + Printer.TextWidth(strColName)) / 2) - (Printer.TextWidth("目標") / 2)
Printer.CurrentY = iPrint
Printer.Print "目標"
Printer.CurrentX = (PLeft(4) + (PLeft(5) - PLeft(4) + Printer.TextWidth(strColName)) / 2) - (Printer.TextWidth("目標達成") / 2)
Printer.CurrentY = iPrint
Printer.Print "目標達成"
Printer.CurrentX = PLeft(7) + (Printer.TextWidth(strColName) / 2) - (Printer.TextWidth("發文達成率 %") / 2)
Printer.CurrentY = iPrint
Printer.Print "發文達成率 %"
Printer.CurrentX = (PLeft(9) + (PLeft(10) - PLeft(9) + Printer.TextWidth(strColName)) / 2) - (Printer.TextWidth("完稿") / 2)
Printer.CurrentY = iPrint
Printer.Print "完稿"
Printer.CurrentX = PLeft(12) + (Printer.TextWidth(strColName) / 2) - (Printer.TextWidth("完稿達成率 %") / 2)
Printer.CurrentY = iPrint
Printer.Print "完稿達成率 %"
Printer.Font.Underline = False
iPrint = iPrint + 300

Printer.Line (PLeft(2), iPrint + 150)-(PLeft(4) - 50, iPrint + 150)
Printer.Line (PLeft(4), iPrint + 150)-(PLeft(6) - 50, iPrint + 150)
Printer.Line (PLeft(6), iPrint + 150)-(PLeft(9) - 50, iPrint + 150)
Printer.Line (PLeft(9), iPrint + 150)-(PLeft(11) - 50, iPrint + 150)
Printer.Line (PLeft(11), iPrint + 150)-(PLeft(13) + PLeft(13) - PLeft(12), iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "基數"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strColName
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print strColName
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print strColName
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print strColName
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "平均"
iPrint = iPrint + 300

'Added by Morgan 2018/12/11
If Check2.Value = 1 Then
   intI = Printer.TextWidth("　")
   Printer.CurrentX = PLeft(2) + intI
   Printer.CurrentY = iPrint
   Printer.Print "分配-發"
   Printer.CurrentX = PLeft(3) + intI
   Printer.CurrentY = iPrint
   Printer.Print "原始-發"
   Printer.CurrentX = PLeft(4) + intI
   Printer.CurrentY = iPrint
   Printer.Print "加乘-發"
   Printer.CurrentX = PLeft(5) + intI
   Printer.CurrentY = iPrint
   Printer.Print "支援-發"
   Printer.CurrentX = PLeft(6) + intI
   Printer.CurrentY = iPrint
   Printer.Print "收文-發"
   Printer.CurrentX = PLeft(7) + intI
   Printer.CurrentY = iPrint
   Printer.Print "件數-發"
   Printer.CurrentX = PLeft(8) + intI
   Printer.CurrentY = iPrint
   Printer.Print "分配-完"
   Printer.CurrentX = PLeft(9) + intI
   Printer.CurrentY = iPrint
   Printer.Print "原始-完"
   Printer.CurrentX = PLeft(10) + intI
   Printer.CurrentY = iPrint
   Printer.Print "加乘-完"
   Printer.CurrentX = PLeft(11) + intI
   Printer.CurrentY = iPrint
   Printer.Print "支援-完"
   Printer.CurrentX = PLeft(12) + intI
   Printer.CurrentY = iPrint
   Printer.Print "收文-完"
   Printer.CurrentX = PLeft(13) + intI
   Printer.CurrentY = iPrint
   Printer.Print "件數-完"
   iPrint = iPrint + 300
End If
'end 2018/12/11

If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(m_lngRptLineEnd, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

'Added by Morgan 2019/3/22
Sub PrintTitle3() '列印抬頭
   
   Dim strColName As String
   Dim lngFix As Long, lngDx As Long

   If Check1.Value = vbUnchecked Then
      strColName = "基數"
   Else
      strColName = "件數"
   End If

   GetPleft2
   
   iPrint = 0
   Printer.Orientation = vbPRORLandscape
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5500
   Printer.CurrentY = iPrint
   Printer.Print "承辦人達成情形統計表"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "統計年月：" & (Val(txt1(3)) \ 100) & "/" & Right(txt1(3), 2) & "－" & (Val(txt1(4)) \ 100) & "/" & Right(txt1(4), 2)
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "系統類別：" & SavDay1 & "　※發文達成率平均=(發文實績點數達成率+發文基數達成率)/2"

   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(m_lngRptLineEnd, iPrint + 150)
   iPrint = iPrint + 300
   If iPrint >= 14000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle3
       Exit Sub
   End If
      
   Printer.Font.Underline = True
   
   lngFix = Printer.TextWidth("點數") - Printer.TextWidth("999.00")
   
   lngDx = (PLeft(3) + Printer.TextWidth("基數") - (PLeft(2) + lngFix) - Printer.TextWidth("目標")) / 2
   Printer.CurrentX = PLeft(2) + lngFix + lngDx
   Printer.CurrentY = iPrint
   Printer.Print "目標"
   
   lngDx = (PLeft(6) + Printer.TextWidth("基數") - (PLeft(4) + lngFix) - Printer.TextWidth("目標達成")) / 2
   Printer.CurrentX = PLeft(4) + lngFix + lngDx
   Printer.CurrentY = iPrint
   Printer.Print "目標達成"
   
   lngDx = (PLeft(10) + Printer.TextWidth("平均") - (PLeft(7) + lngFix) - Printer.TextWidth("發文達成率 %")) / 2
   Printer.CurrentX = PLeft(7) + lngFix + lngDx
   Printer.CurrentY = iPrint
   Printer.Print "發文達成率 %"
   
   lngDx = (PLeft(12) + Printer.TextWidth("基數") - (PLeft(11) + lngFix) - Printer.TextWidth("完稿")) / 2
   Printer.CurrentX = PLeft(11) + lngFix + lngDx
   Printer.CurrentY = iPrint
   Printer.Print "完稿"
   
   lngDx = (PLeft(15) + Printer.TextWidth("基數") - (PLeft(13) + lngFix) - Printer.TextWidth("完稿達成率 %")) / 2
   Printer.CurrentX = PLeft(13) + lngFix + lngDx
   Printer.CurrentY = iPrint
   Printer.Print "完稿達成率 %"
   
   Printer.Font.Underline = False
   iPrint = iPrint + 300

   Printer.Line (PLeft(2) + lngFix, iPrint + 150)-(PLeft(3) + Printer.TextWidth("基數"), iPrint + 150)
   Printer.Line (PLeft(4) + lngFix, iPrint + 150)-(PLeft(6) + Printer.TextWidth("基數"), iPrint + 150)
   Printer.Line (PLeft(7) + lngFix, iPrint + 150)-(PLeft(10) + Printer.TextWidth("平均"), iPrint + 150)
   Printer.Line (PLeft(11) + lngFix, iPrint + 150)-(PLeft(12) + Printer.TextWidth("基數"), iPrint + 150)
   Printer.Line (PLeft(13) + lngFix, iPrint + 150)-(PLeft(15) + Printer.TextWidth("平均"), iPrint + 150)

   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   '目標
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "基數"
   '目標達成
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("實績")
   Printer.CurrentY = iPrint
   Printer.Print "實績點數"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strColName
   '發文達成率
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("實績")
   Printer.CurrentY = iPrint
   Printer.Print "實績點數"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print strColName
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "平均"
   '完稿
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(12)
   Printer.CurrentY = iPrint
   Printer.Print strColName
   '完稿達成率
   Printer.CurrentX = PLeft(13)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(14)
   Printer.CurrentY = iPrint
   Printer.Print strColName
   Printer.CurrentX = PLeft(15)
   Printer.CurrentY = iPrint
   Printer.Print "平均"
   iPrint = iPrint + 300

   'Added by Morgan 2018/12/11
   If Check2.Value = 1 Then
      intI = Printer.TextWidth("　")
      Printer.CurrentX = PLeft(2) + intI
      Printer.CurrentY = iPrint
      Printer.Print "分配-發"
      Printer.CurrentX = PLeft(3) + intI
      Printer.CurrentY = iPrint
      Printer.Print "原始-發"
      Printer.CurrentX = PLeft(4) + intI
      Printer.CurrentY = iPrint
      Printer.Print "加乘-發"
      Printer.CurrentX = PLeft(5) + intI
      Printer.CurrentY = iPrint
      Printer.Print "支援-發"
      Printer.CurrentX = PLeft(6) + intI
      Printer.CurrentY = iPrint
      'Printer.Print "收文-發"
      Printer.CurrentX = PLeft(7) + intI
      Printer.CurrentY = iPrint
      Printer.Print "件數-發"
      Printer.CurrentX = PLeft(8) + intI
      Printer.CurrentY = iPrint
      Printer.Print "分配-完"
      Printer.CurrentX = PLeft(9) + intI
      Printer.CurrentY = iPrint
      Printer.Print "原始-完"
      Printer.CurrentX = PLeft(10) + intI
      Printer.CurrentY = iPrint
      Printer.Print "加乘-完"
      Printer.CurrentX = PLeft(11) + intI
      Printer.CurrentY = iPrint
      Printer.Print "支援-完"
      Printer.CurrentX = PLeft(12) + intI
      Printer.CurrentY = iPrint
      'Printer.Print "收文-完"
      Printer.CurrentX = PLeft(13) + intI
      Printer.CurrentY = iPrint
      Printer.Print "件數-完"
      iPrint = iPrint + 300
   End If
   'end 2018/12/11

   If iPrint >= 14000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle3
       Exit Sub
   End If
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(m_lngRptLineEnd, iPrint + 150)
   iPrint = iPrint + 300
   If iPrint >= 14000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle3
       Exit Sub
   End If
End Sub

Sub PrintDatil() '列印資料

'Added by Morgan 2018/12/11
If Check2.Value = 1 And iPrint > 2900 Then
   'intI = Printer.TextHeight("　")
   Printer.DrawStyle = vbDot
   Printer.Line (0, iPrint - 50)-(m_lngRptLineEnd, iPrint - 50)
   Printer.DrawStyle = vbSolid
End If
'end 2018/12/11

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)

Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(strTemp(2), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(2), "##0.00")
Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(strTemp(3), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(3), "##0.00")
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(4), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(4), "####0.00")
Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(strTemp(5), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(5), "##0.00")
For i = 6 To 13
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "##0.00")
Next i
iPrint = iPrint + 300

'Added by Morgan 2018/12/11
If Check2.Value = 1 Then
   intI = Printer.TextWidth("　分配-發")
   For i = 14 To 25
       Printer.CurrentX = PLeft(i - 12) + intI - Printer.TextWidth(Format(strTemp(i), "##0.00"))
       Printer.CurrentY = iPrint
       Printer.Print Format(strTemp(i), "##0.00")
   Next i
   iPrint = iPrint + 300
End If
End Sub

Sub PrintDatil3() '列印資料

'Added by Morgan 2018/12/11
If Check2.Value = 1 And iPrint > 2900 Then
   'intI = Printer.TextHeight("　")
   Printer.DrawStyle = vbDot
   Printer.Line (0, iPrint - 50)-(m_lngRptLineEnd, iPrint - 50)
   Printer.DrawStyle = vbSolid
End If
'end 2018/12/11

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)

Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(strTemp(2), "##0.0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(2), "##0.0")
Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(strTemp(3), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(3), "##0.00")
'點數
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(4), "####0.0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(4), "####0.0")
'實績點數
Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(strTemp(26), "####0.0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(26), "####0.0")
'基數
Printer.CurrentX = PLeft(6) + 500 - Printer.TextWidth(Format(strTemp(5), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(5), "##0.00")
'點數達成率
Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(strTemp(6), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(6), "##0.00")
'實績點數達成率
Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(strTemp(27), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(27), "##0.00")

For i = 7 To 13
   If i = 9 Then
      Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(strTemp(i), "##0.0"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(i), "##0.0")
   Else
      Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(i), "##0.00")
   End If
Next i
iPrint = iPrint + 300

'Added by Morgan 2018/12/11
If Check2.Value = 1 Then
   intI = Printer.TextWidth("　分配-發")
   For i = 14 To 25
      If i <> 18 And i <> 24 Then '剔除收文
         Printer.CurrentX = PLeft(i - 12) + intI - Printer.TextWidth(Format(strTemp(i), "##0.00"))
         Printer.CurrentY = iPrint
         Printer.Print Format(strTemp(i), "##0.00")
      End If
   Next i
   iPrint = iPrint + 300
End If
End Sub

Sub GetPleft()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
For i = 2 To 13
   PLeft(i) = 1500 + (i - 2) * 1160
Next i
m_lngRptLineEnd = 16000
End Sub

'Added by Morgan 2019/3/22
Sub GetPleft2()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0

For i = 2 To 16
   PLeft(i) = 1500 + (i - 2) * 1100
Next i
m_lngRptLineEnd = 16250
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(m_lngRptLineEnd, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    If m_bol108Rule Then
      PrintTitle3
    Else
      PrintTitle
   End If
End If
End Sub

Sub PrintEnd()
'列印結尾
'Modified by Morgan 2018/12/6 修正達成率問題
'strSql = "SELECT '','合  計',SUM(R102003),SUM(R102004),SUM(R102005),SUM(R102006),SUM(R102007),SUM(R102008),SUM(R102009),SUM(R102010),SUM(R102011),SUM(R102012),SUM(R102013),sum(r102014) FROM R090608 WHERE ID='" & strUserNum & "' AND R102001='" & SavDay1 & "' "
strSql = "SELECT '','合  計',SUM(R102003),SUM(R102004),SUM(R102005),SUM(R102006)" & _
   ",decode(SUM(R102003),0,0,round(100*SUM(R102005)/SUM(R102003),2)) as R102007" & _
   ",decode(SUM(R102004),0,0,round(100*SUM(R102006)/SUM(R102004),2)) as R102008" & _
   ",(decode(SUM(R102003),0,0,round(100*SUM(R102005)/SUM(R102003),2))+decode(SUM(R102004),0,0,round(100*SUM(R102006)/SUM(R102004),2)))/2 as R102009" & _
   ",SUM(R102010),SUM(R102011)" & _
   ",decode(SUM(R102003),0,0,round(100*SUM(R102010)/SUM(R102003),2)) as R102012" & _
   ",decode(SUM(R102004),0,0,round(100*SUM(R102011)/SUM(R102004),2)) as R102013" & _
   ",(decode(SUM(R102003),0,0,round(100*SUM(R102010)/SUM(R102003),2))+decode(SUM(R102004),0,0,round(100*SUM(R102011)/SUM(R102004),2)))/2 as R102014" & _
   " FROM R090608 WHERE ID='" & strUserNum & "' AND R102001='" & SavDay1 & "' "
'end 2018/12/6
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 13
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            StrTemp7(8) = str((Val(StrTemp7(6)) + Val(StrTemp7(7))) / 2)
            StrTemp7(13) = str((Val(StrTemp7(11)) + Val(StrTemp7(12))) / 2)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(StrTemp7(2), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(2), "####0.00")
            Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(StrTemp7(3), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(3), "####0.00")
            Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(StrTemp7(4), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "####0.00")
            Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(StrTemp7(5), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(5), "####0.00")
            For i = 6 To 13
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0.00")
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Added by Morgan 2019/3/25 108考核
Sub PrintEnd3()
'列印結尾
strSql = "SELECT '','合  計',SUM(R102003),SUM(R102004),SUM(R102005),SUM(R102006)" & _
   ",decode(SUM(R102003),0,0,round(100*SUM(R102005)/SUM(R102003),2)) as R102007" & _
   ",decode(SUM(R102004),0,0,round(100*SUM(R102006)/SUM(R102004),2)) as R102008" & _
   ",(decode(SUM(R102003),0,0,round(100*SUM(R102022)/SUM(R102003),2))+decode(SUM(R102004),0,0,round(100*SUM(R102006)/SUM(R102004),2)))/2 as R102009" & _
   ",SUM(R102010),SUM(R102011)" & _
   ",decode(SUM(R102003),0,0,round(100*SUM(R102010)/SUM(R102003),2)) as R102012" & _
   ",decode(SUM(R102004),0,0,round(100*SUM(R102011)/SUM(R102004),2)) as R102013" & _
   ",(decode(SUM(R102003),0,0,round(100*SUM(R102010)/SUM(R102003),2))+decode(SUM(R102004),0,0,round(100*SUM(R102011)/SUM(R102004),2)))/2 as R102014" & _
   ",SUM(R102022) as R102022" & _
   ",decode(SUM(R102003),0,0,round(100*SUM(R102022)/SUM(R102003),2)) as R102023" & _
   " FROM R090608 WHERE ID='" & strUserNum & "' AND R102001='" & SavDay1 & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 15
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            '點數-目標
            Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(StrTemp7(2), "####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(2), "####0.0")
            '基數-目標
            Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(StrTemp7(3), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(3), "####0.00")
            '點數-達成
            Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(StrTemp7(4), "####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "####0.0")
            '實績點數-達成
            Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(StrTemp7(14), "####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(14), "####0.0")
            '基數-達成
            Printer.CurrentX = PLeft(6) + 500 - Printer.TextWidth(Format(StrTemp7(5), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(5), "####0.00")
            '點數達成率
            Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(StrTemp7(6), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(6), "####0.00")
            '實績點數達成率
            Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(StrTemp7(15), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(15), "####0.00")
                        
            For i = 7 To 13
               If i = 9 Then
                  Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0.0"))
                  Printer.CurrentY = iPrint
                  Printer.Print Format(StrTemp7(i), "####0.0")
               Else
                  Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0.00"))
                  Printer.CurrentY = iPrint
                  Printer.Print Format(StrTemp7(i), "####0.00")
               End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle3
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'add by nickc 2005/03/04 新制算法
Sub ProcessNew()
cnnConnection.Execute "DELETE FROM R090608 WHERE ID='" & strUserNum & "' "
strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = ""
strSQL5 = "": StrSQL6 = "": StrSQL7 = "": strSQL8 = "": strSQL10 = ""
strSQL1 = strSQL1 + " AND pe03>=" & Val(txt1(3)) + 191100 & " AND pe03<=" & Val(txt1(4)) + 191100 & " "
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND pe01='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND ST06>='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
    strSQL1 = strSQL1 + " AND ST06<='" & txt1(7) & "' "
End If
'部門別
If Me.txt1(10).Text <> "" Then
    strSQL1 = strSQL1 + " AND ST03>='" & txt1(10) & "' "
End If
If Me.txt1(11).Text <> "" Then
    strSQL1 = strSQL1 + " AND ST03<='" & txt1(11) & "' "
End If
'只限在職人員的資料
If Check3.Value = vbUnchecked Then 'Added by Morgan 2024/4/24
   strSQL1 = strSQL1 + " and ST04='1' "
End If
CheckOC
'***目標***
'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,r102005,r102006,r102010,r102011,id) " & _
              " Select 'ALL',pe01,sum(nvl(pe06,0)+nvl(pe08,0)),sum(nvl(decode(pe02,'CFP',pe05*2,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)),sum(nvl(ma40,0))/count(pe01),sum(nvl(ma37,0))/count(pe01),sum(nvl(ma50,0))/count(pe01),sum(nvl(ma43,0))/count(pe01),'" & strUserNum & "' From Performance, Staff,monthassess " & _
              "Where PE01=ST01(+) and pe02 in ('P','CFP') and pe01=ma01(+) and pe03=ma02(+) " & strSQL1 & " group by pe01,'" & strUserNum & "'"
cnnConnection.Execute strSql
'***整理***
strSql = "SELECT 'ALL',R102002," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2,'" & strUserNum & "' " & _
         "FROM R090608 WHERE ID='" & strUserNum & "' GROUP BY R102002,'" & strUserNum & "' "
cnnConnection.Execute "INSERT INTO R090608 (R102001,R102002,R102007,R102008,R102009,R102012,R102013,R102014,ID) " & strSql
End Sub

Sub Process()
cnnConnection.Execute "DELETE FROM R090608 WHERE ID='" & strUserNum & "' "
strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = ""
strSQL5 = "": StrSQL6 = "": StrSQL7 = "": strSQL8 = "": strSQL10 = ""
If txt1(0) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/14
End If
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/14
End If
StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
StrSQL7 = StrSQL7 + " AND EP09>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP09<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
'Add By Cheng 2003/12/16
strSQL8 = strSQL8 + " AND SH01>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND SH01<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
'End
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/14
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & txt1(5) & "' "
    StrSQL7 = StrSQL7 + " AND EP05='" & txt1(5) & "' "
    'Add By Cheng 2003/12/16
    strSQL8 = strSQL8 + " AND SH02='" & txt1(5) & "' "
    'End
    pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0) 'Add By Sindy 2010/12/14
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06>='" & txt1(6) & "' "
    StrSQL7 = StrSQL7 + " AND ST06>='" & txt1(6) & "' "
    'Add By Cheng 2003/12/16
    strSQL8 = strSQL8 + " AND ST06>='" & txt1(6) & "' "
    'End
End If
If Len(txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06<='" & txt1(7) & "' "
    StrSQL7 = StrSQL7 + " AND ST06<='" & txt1(7) & "' "
    'Add By Cheng 2003/12/16
    strSQL8 = strSQL8 + " AND ST06<='" & txt1(7) & "' "
    'End
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/14
End If
'Add By Cheng 2003/06/11
'部門別
If Me.txt1(10).Text <> "" Then
    StrSQL6 = StrSQL6 + " AND ST03>='" & txt1(10) & "' "
    StrSQL7 = StrSQL7 + " AND ST03>='" & txt1(10) & "' "
    'Add By Cheng 2003/12/16
    strSQL8 = strSQL8 + " AND ST03>='" & txt1(10) & "' "
    'End
End If
If Me.txt1(11).Text <> "" Then
    StrSQL6 = StrSQL6 + " AND ST03<='" & txt1(11) & "' "
    StrSQL7 = StrSQL7 + " AND ST03<='" & txt1(11) & "' "
    'Add By Cheng 2003/12/16
    strSQL8 = strSQL8 + " AND ST03<='" & txt1(11) & "' "
    'End
End If
If Me.txt1(10).Text <> "" Or Me.txt1(11).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/14
End If
'Modify By Cheng 2003/12/10
'不論計不計件點數仍要算
'StrSQL6 = StrSQL6 + " and CP26 IS NULL "
'StrSQL7 = StrSQL7 + " and CP26 IS NULL "
'End
'Add By Cheng 2003/09/03
'只限在職人員的資料
If Check3.Value = vbUnchecked Then 'Added by Morgan 2024/4/24
   StrSQL6 = StrSQL6 + " and ST04='1' "
   StrSQL7 = StrSQL7 + " and ST04='1' "
   'Add By Cheng 2003/12/16
   strSQL8 = strSQL8 + " and ST04='1' "
End If
strSQL8 = strSQL8 + " and SH11='V' "
strSQL10 = strSQL10 + " and SCR03='V' "
'End
CheckOC
'900629 NICK 改
'strSQL = "SELECT CP01,ST02,PE06 AS A,PE05 AS B,SUM(CP18) AS C,COUNT(CP27) AS D, SUM(CP18)/PE05 * 100, COUNT(CP27)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(CP27)/PE06))/2)*100),0,0,0,0,0 FROM CASEPROGRESS,TRADEMARK,PERFORMANCE,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP14=PE01(+) AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03(+) " & StrSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,PE06 AS A,PE05 AS B,SUM(CP18) AS C,COUNT(CP27) AS D, SUM(CP18)/PE05 * 100, COUNT(CP27)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(CP27)/PE06))/2)*100),0,0,0,0,0 FROM CASEPROGRESS,PATENT,PERFORMANCE,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP14=PE01(+) AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03(+) " & StrSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,PE06 AS A,PE05 AS B,SUM(CP18) AS C,COUNT(CP27) AS D, SUM(CP18)/PE05 * 100, COUNT(CP27)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(CP27)/PE06))/2)*100),0,0,0,0,0 FROM CASEPROGRESS,LAWCASE,PERFORMANCE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) AND CP14=PE01(+) AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,PE06 AS A,PE05 AS B,SUM(CP18) AS C,COUNT(CP27) AS D, SUM(CP18)/PE05 * 100, COUNT(CP27)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(CP27)/PE06))/2)*100),0,0,0,0,0 FROM CASEPROGRESS,HIRECASE,PERFORMANCE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) AND CP14=PE01(+) AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,PE06 AS A,PE05 AS B,SUM(CP18) AS C,COUNT(CP27) AS D, SUM(CP18)/PE05 * 100, COUNT(CP27)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(CP27)/PE06))/2)*100),0,0,0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,PERFORMANCE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP14=PE01(+) AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03(+) " & StrSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,0,0,0,0,0,0,0,SUM(CP18) AS C,COUNT(EP09) AS D, SUM(CP18)/PE05 * 100, COUNT(EP09)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(EP09)/PE06))/2)*100) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,PERFORMANCE,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) AND EP05=PE01 AND CP01=PE02(+) AND SUBSTR(EP09,1,6)=PE03 " & StrSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,0,0,0,0,0,0,0,SUM(CP18) AS C,COUNT(EP09) AS D, SUM(CP18)/PE05 * 100, COUNT(EP09)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(EP09)/PE06))/2)*100) FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,PERFORMANCE,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) AND EP05=PE01 AND CP01=PE02(+) AND SUBSTR(EP09,1,6)=PE03 " & StrSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,0,0,0,0,0,0,0,SUM(CP18) AS C,COUNT(EP09) AS D, SUM(CP18)/PE05 * 100, COUNT(EP09)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(EP09)/PE06))/2)*100) FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,PERFORMANCE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) AND EP05=PE01 AND CP01=PE02(+) AND SUBSTR(EP09,1,6)=PE03 " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,0,0,0,0,0,0,0,SUM(CP18) AS C,COUNT(EP09) AS D, SUM(CP18)/PE05 * 100, COUNT(EP09)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(EP09)/PE06))/2)*100) FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,PERFORMANCE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) AND EP05=PE01 AND CP01=PE02(+) AND SUBSTR(EP09,1,6)=PE03 " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY CP01,ST02,PE05,PE06 "
'strSQL = strSQL + " UNION all  SELECT CP01,ST02,0,0,0,0,0,0,0,SUM(CP18) AS C,COUNT(EP09) AS D, SUM(CP18)/PE05 * 100, COUNT(EP09)/PE06 * 100,((((SUM(CP18)/PE05)+(COUNT(EP09)/PE06))/2)*100) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,PERFORMANCE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) AND EP05=PE01 AND CP01=PE02(+) AND SUBSTR(EP09,1,6)=PE03 " & StrSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY CP01,ST02,PE05,PE06 "
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        DoEvents
'        Do While .EOF = False
'            For i = 0 To 13
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSQL = "INSERT INTO R090608 VALUES ('" & strTemp(0) & "','" & strTemp(1) & "'," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & ",'" & strUserNum & "') "
'            cnnConnection.Execute strSQL
'            .MoveNext
'            DoEvents
'        Loop
'    End If
'End With
'CheckOC
'***目標***
'Modify By Cheng 2003/05/14
'strSQL = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' from performance wheRE pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(txt1(0)) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & " group by pe02,pe01,'" & strUserNum & "' "
'Modify By Cheng 2003/06/11
'strSQL = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(txt1(0)) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & " group by pe02,pe01,'" & strUserNum & "' "
'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
'2012/1/19 modify by sonia 若部門別為P2字頭則抓T的目標
'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) and pe02 in ('P','CFP') And pe03>=" & Val(Txt1(3)) + 191100 & " and pe03<=" & Val(Txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(Txt1(0)) & ") " & IIf(Len(Txt1(5)) <> 0, " AND PE01='" & Txt1(5) & "' ", "") & IIf(Me.Txt1(6).Text <> "", " AND ST06>='" & Txt1(6) & "' ", "") & IIf(Me.Txt1(7).Text <> "", " AND ST06<='" & Txt1(7) & "' ", "") & _
                IIf(Me.Txt1(10).Text <> "", " And ST03>='" & Me.Txt1(10).Text & "' ", "") & IIf(Me.Txt1(11).Text <> "", " And ST03<='" & Me.Txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
If Left(txt1(10), 2) = "P2" Then
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) and pe02='T' And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01,'" & strUserNum & "' "
Else
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) and pe02 in ('P','CFP') And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(txt1(0)) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01,'" & strUserNum & "' "
End If
'2012/1/19 END
cnnConnection.Execute strSql
'***發文目標達成***
'Modify By Cheng 2003/07/07
'                strSQL = "SELECT cp01,cp14,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT cp01,cp14,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT cp01,cp14,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT cp01,cp14,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT cp01,cp14,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & StrSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
                strSql = "SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id) " & strSql

If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22

   'Add By Cheng 2003/12/16
   '支援記錄
   'Modify By Cheng 2004/03/01
                   strSql = "SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'P' As SH06, SH02, 0, Sum(Round(Decode('P', 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & " AND 'P' IN (" & SQLGrpStr(txt1(0), 1) & ") And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   'End
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id) " & strSql

   'End
   'Add By Cheng 2003/12/17
   '特殊案件記錄
                   strSql = "SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL2 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL1 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL3 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL4 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL5 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id) " & strSql

End If 'Added by Morgan 2013/5/22

'End
'***完稿達成***
'Modify By Cheng 2003/07/07
'               strSQL = " SELECT CP01,ep05,SUM(CP18),COUNT(EP09),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT CP01,ep05,SUM(CP18),COUNT(EP09),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT CP01,ep05,SUM(CP18),COUNT(EP09),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT CP01,ep05,SUM(CP18),COUNT(EP09),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT CP01,ep05,SUM(CP18),COUNT(EP09),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
               strSql = " SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id) " & strSql

If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22

   'Add By Cheng 2003/12/16
   '支援記錄
   'Modify By Cheng 2004/03/01
   '                strSQL = "SELECT SH06, SH02, 0, Sum(Round(Nvl(SH05, 0)/4,2)),'" & strUserNum & "' FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01 AND SH07=TM02 AND SH08=TM03 AND SH09=TM04 AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'strSQL = strSQL + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Nvl(SH05, 0)/4,2)),'" & strUserNum & "' FROM SupportHour, PATENT, STAFF WHERE SH06=PA01 AND SH07=PA02 AND SH08=PA03 AND SH09=PA04 AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'strSQL = strSQL + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Nvl(SH05, 0)/4,2)),'" & strUserNum & "' FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01 AND SH07=LC02 AND SH08=LC03 AND SH09=LC04 AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'strSQL = strSQL + " UNION all  SELECT SH06, SH02, 0, Sum(Nvl(Round(SH05, 0)/4,2)),'" & strUserNum & "' FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01 AND SH07=HC02 AND SH08=HC03(+) AND SH09=HC04 AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'strSQL = strSQL + " UNION all  SELECT SH06, SH02, 0, Sum(Nvl(Round(SH05, 0)/4,2)),'" & strUserNum & "' FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01 AND SH07=SP02 AND SH08=SP03 AND SH09=SP04 AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
                   strSql = "SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'P' As SH06, SH02, 0, Sum(Round(Decode('P', 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "' FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & " AND 'P' IN (" & SQLGrpStr(txt1(0), 1) & ") And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   'End
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id) " & strSql
   '特殊案件記錄
                  strSql = " SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL1 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL2 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL3 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL4 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL5 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id) " & strSql
   'End
   
End If 'Added by Morgan 2013/5/22

'***整理***
'strSQL = "SELECT R102001,R102002,SUM"
strSql = "SELECT R102001,R102002," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2,'" & strUserNum & "' " & _
         "FROM R090608 WHERE ID='" & strUserNum & "' GROUP BY R102001,R102002,'" & strUserNum & "' "
cnnConnection.Execute "INSERT INTO R090608 (R102001,R102002,R102007,R102008,R102009,R102012,R102013,R102014,ID) " & strSql
End Sub

Private Sub Form_Activate()
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   MoveFormToCenter Me
   txt1(0) = Systemkind_g
   For i = 0 To 7
       StrTemp99(i) = ""
   Next i
   txt1(8) = "1"
   bol911001checkRange = True
   
   'Added by Morgan 2025/7/17
   If Pub_StrUserSt03 <> "M51" Then
      Label1(11) = "統計方式：         (1.新制  2.舊制)"
   End If
   'end 2025/7/17
      
   If strSrvDate(1) >= "20050315" And (Mid(Pub_StrUserSt03, 1, 1) = "P" Or Pub_StrUserSt03 = "M51") Then
      txt1(12).Text = "1"
      Label1(1).Visible = False
      txt1(1).Visible = False
      txt1(2).Visible = False
      Line1.Visible = False
      txt1(0).Visible = False
      Label1(0).Visible = False
   Else
      txt1(12).Text = "2"
      txt1(12).Enabled = False
      Check1.Enabled = False
      Check2.Enabled = False
   End If
   'Add By Cheng 2003/07/30
   '若為個人
   If ProState = "1" Then
       GetPersonalData
       Me.txt1(6).Enabled = False
       Me.txt1(7).Enabled = False
       Me.txt1(10).Enabled = False
       Me.txt1(11).Enabled = False
       Me.txt1(5).Enabled = False
   '2012/1/19 ADD BY SONIA
   Else
      Select Case Mid(GetStaffDepartment(strUserNum), 1, 2)
         Case "P1"
            If GetStaffDepartment(strUserNum) = "P12" Then
               txt1(10) = ""
               txt1(11) = ""
            Else
               txt1(10) = "P10"
               txt1(11) = "P11"
            End If
         Case "P2"
            txt1(10) = "P20"
            txt1(11) = "P21"
         Case Else
            txt1(10) = ""
            txt1(11) = ""
      End Select
   '2012/1/19 END
   End If
End Sub

Private Sub GetPersonalData()
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
   
   StrSQLa = "SELECT * FROM STAFF WHERE ST01='" & strUserNum & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Me.txt1(6).Text = "" & rsA("ST06").Value
      Me.txt1(7).Text = "" & rsA("ST06").Value
      Me.txt1(10).Text = "" & rsA("ST03").Value
      Me.txt1(11).Text = "" & rsA("ST03").Value
      Me.txt1(5).Text = "" & rsA("ST01").Value
      Me.lbl1(0).Caption = "" & rsA("ST02").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090608 = Nothing
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
   
   'Added by Morgan 2025/7/10
   If Index = 12 Then
      If KeyAscii = 51 Then
         txt1(8) = "1"
         txt1(8).Enabled = False
         Check1.Value = vbUnchecked
         Check1.Enabled = False
         Check2.Value = vbUnchecked
         Check2.Enabled = False
      Else
         txt1(8).Enabled = True
         Check1.Enabled = True
         Check2.Enabled = True
      End If
   End If
   'end 2025/7/10
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
      'Add By Cheng 2002/01/07
'      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
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
    If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
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
     lbl1(0) = GetPrjSalesNM(txt1(5))
     If Trim(txt1(Index)) <> "" Then
        If Trim(lbl1(0).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 6
     bol911001checkRange = True
     Select Case Trim(txt1(6))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 7
     If bol911001checkRange = True Then
          Select Case Trim(txt1(7))
          Case "1", "2", "3", "4", "5", ""
          Case Else
               s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
               txt1(7).SetFocus
               txt1(7).SelStart = 0
               txt1(7).SelLength = Len(txt1(7))
               Exit Sub
          End Select
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
   End If
   bol911001checkRange = True
Case 8
     Select Case Trim(txt1(8))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Exit Sub
     End Select
Case 9
     Select Case Trim(txt1(9))
     'Modified by Lydia 2016/12/19 + 5
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("列印順序只能輸入 1 ~ 5   !!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1(9).SelStart = 0
          txt1(9).SelLength = Len(txt1(9))
          Exit Sub
     End Select
'Add By Cheng 2003/06/11
Case 10, 11 '部門別區間
    If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
        If Me.txt1(10).Text > Me.txt1(11).Text Then
            MsgBox "部門別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(10).SetFocus
            txt1_GotFocus 10
            Exit Sub
        End If
    End If

'add by nickc 2005/03/04
Case 12
   txt1(8).Enabled = True 'Added by Morgan 2025/7/18
   Select Case txt1(Index)
   'Modified by Morgan 2025/7/17 +3
   Case "1", "3"
      'If (strSrvDate(1) >= "20050315" And Mid(Pub_StrUserSt03, 1, 1) = "P") Then
         Label1(1).Visible = False
         txt1(1).Visible = False
         txt1(2).Visible = False
         Line1.Visible = False
         txt1(0).Visible = False
         Label1(0).Visible = False
      'Else
      '   Label1(1).Visible = True
      '   txt1(1).Visible = True
      '   txt1(2).Visible = True
      '   Line1.Visible = True
      '   txt1(0).Visible = True
      '   Label1(0).Visible = True
      'End If
      If txt1(Index) = "3" Then txt1(8) = "1": txt1(8).Enabled = False 'Added by Morgan 2025/7/18
   Case "2"
      Label1(1).Visible = False
      txt1(1).Visible = False
      txt1(2).Visible = False
      Line1.Visible = False
      txt1(0).Visible = False
      Label1(0).Visible = False
      txt1(0) = Systemkind_g
   Case Else
          MsgBox "請輸入 1、2 或 3 ！", , "選擇統計方式！"
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
   End Select
Case Else
End Select
End Sub

'即時算
Sub ProcessNew2()
'add by nickc 2008/01/03
txt1(0) = "ALL"
cnnConnection.Execute "DELETE FROM R090608 WHERE ID='" & strUserNum & "' "
strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = ""
strSQL5 = "": StrSQL6 = "": StrSQL7 = "": strSQL8 = "": strSQL10 = "": strSQL12 = ""
If txt1(0) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/14
End If
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/14
End If
StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
StrSQL7 = StrSQL7 + " AND EP09>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP09<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
strSQL8 = strSQL8 + " AND SH01>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND SH01<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
strSQL12 = strSQL12 & " AND cp05>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND cp05<=" & Val(ChangeTStringToWString(txt1(4) & "31")) 'Add by Morgan 2010/10/18
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/14
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & txt1(5) & "' "
    StrSQL7 = StrSQL7 + " AND EP05='" & txt1(5) & "' "
    strSQL8 = strSQL8 + " AND SH02='" & txt1(5) & "' "
    strSQL12 = strSQL12 & " AND CP13='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0) 'Add By Sindy 2010/12/14
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06>='" & txt1(6) & "' "
    StrSQL7 = StrSQL7 + " AND ST06>='" & txt1(6) & "' "
    strSQL8 = strSQL8 + " AND ST06>='" & txt1(6) & "' "
    strSQL12 = strSQL12 & " AND ST06>='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06<='" & txt1(7) & "' "
    StrSQL7 = StrSQL7 + " AND ST06<='" & txt1(7) & "' "
    strSQL8 = strSQL8 + " AND ST06<='" & txt1(7) & "' "
    strSQL12 = strSQL12 & " AND ST06<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/14
End If
'部門別
If Me.txt1(10).Text <> "" Then
    StrSQL6 = StrSQL6 + " AND ST03>='" & txt1(10) & "' "
    StrSQL7 = StrSQL7 + " AND ST03>='" & txt1(10) & "' "
    strSQL8 = strSQL8 + " AND ST03>='" & txt1(10) & "' "
    strSQL12 = strSQL12 & " AND ST03>='" & txt1(10) & "' "
End If
If Me.txt1(11).Text <> "" Then
    StrSQL6 = StrSQL6 + " AND ST03<='" & txt1(11) & "' "
    StrSQL7 = StrSQL7 + " AND ST03<='" & txt1(11) & "' "
    strSQL8 = strSQL8 + " AND ST03<='" & txt1(11) & "' "
    strSQL12 = strSQL12 + " AND ST03<='" & txt1(11) & "' "
End If
If Me.txt1(10).Text <> "" Or Me.txt1(11).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/14
End If
'只限在職人員的資料
If Check3.Value = vbUnchecked Then 'Added by Morgan 2024/4/24
   StrSQL6 = StrSQL6 + " and ST04='1' "
   StrSQL7 = StrSQL7 + " and ST04='1' "
   strSQL8 = strSQL8 + " and ST04='1' "
End If
strSQL8 = strSQL8 + " and SH11='V' "
strSQL10 = strSQL10 + " and SCR03='V' "
'Modify by Morgan 2010/11/4 +智權人員收文不算(74018杜燕文)
'Modified by Morgan 2012/4/20 改只抓業務區為 P1 的(因為商標處不需要統計)
'strSQL12 = strSQL12 + " and ST04='1' and substr(cp12,1,1)<>'S'"
strSQL12 = strSQL12 + IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " and substr(cp12,1,2)='P1'"

CheckOC
'***目標***
'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
'2012/1/19 modify by sonia 若部門別為P2字頭則抓T的目標
'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
If Left(txt1(10), 2) = "P2" Then
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02='T' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01,'" & strUserNum & "' "
Else
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01,'" & strUserNum & "' "
End If
'2012/1/19 END
cnnConnection.Execute strSql
'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select 'ALL',pe01,sum(pe06),sum(pe05 * decode(pe02,'CFP',2,1)),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
'2012/1/19 modify by sonia 若部門別為P2字頭則抓T的目標
'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select 'ALL',pe01,sum(pe06),sum(pe05 * decode(pe02,'CFP',2,1)),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
If Left(txt1(10), 2) = "P2" Then
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select 'ALL',pe01,sum(pe06),sum(pe05 * decode(pe02,'CFP',2,1)),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02='T' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01,'" & strUserNum & "' "
Else
   'Modified by Morgan 2014/9/24 目標改抓基數設定值
   'strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select 'ALL',pe01,sum(pe06),sum(pe05 * decode(pe02,'CFP',2,1)),'" & strUserNum & "' From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & " And ST04='1' group by pe02,pe01,'" & strUserNum & "' "
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id) Select 'ALL',pe01,sum(pe06),sum(er03),'" & strUserNum & "' from (select pe01,pe03,sum(pe06) pe06 From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & IIf(Me.txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
                   IIf(Me.txt1(10).Text <> "", " And ST03>='" & Me.txt1(10).Text & "' ", "") & IIf(Me.txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe01,pe03),engradix where er01(+)=pe03 and er02(+)=pe01 group by pe01"
   m_bolShowMemo = True
   'end 2014/9/24
End If
'2012/1/19 END
cnnConnection.Execute strSql
If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22
   'Modified by Morgan 2018/11/6 +R102015:來源 >> 1=案件, 2=支援, 3=特殊, 4=收文
   'Modified by Morgan 2018/11/7 +R102016:分配點數, R102017 不含加乘的基數, R102018 件數

   '***發文目標達成***
   strSql = "SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1))),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   'strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(Txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(Decode(CP27, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql & " union all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1))),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   'strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(Txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "

   '支援記錄
   'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
   strSql = strSql & " union all SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'P' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   strSql = strSql & " union all SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'Modified by Morgan 2019/3/6
   'strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(Decode('P', 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   'end 2019/3/6
'end 2014/3/20
   
   '特殊案件記錄
   strSql = strSql & " union all SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL2 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL1 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL3 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL4 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL5 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql & " union all SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL2 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL1 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL3 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL4 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & strSQL5 & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "

   'Add by Morgan 2010/10/18
   '2012/1/20 MODIFY BY SONIA 加部門別條件,否則統計商標處的會出現P的資料
   'If bolNewPromoterRule Then
   'Modified by Morgan 2012/4/20 部門條件改放 strSQL12 控制否則只輸員工編號但沒有輸部門條件查詢時會沒統計到
   'If bolNewPromoterRule And Left(txt1(10), 2) = "P1" Then
   If bolNewPromoterRule Then
      'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
      'strSql = strSql + " UNION all  SELECT 'P' As cp01, cp13, 0, Sum(Round(nvl(a0n03/1000,cp18)*0.05 ,2)),'" & strUserNum & "' FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      'strSql = strSql + " UNION all  SELECT 'ALL' As cp01, cp13, 0, Sum(Round(nvl(a0n03/1000,cp18)*0.05 ,2)),'" & strUserNum & "' FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      strSql = strSql + " UNION all  SELECT 'P' As cp01, cp13, 0, Sum(Round(" & Pt2EPtCode & ",2)),'" & strUserNum & "','4',0,0,0 FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      strSql = strSql + " UNION all  SELECT 'ALL' As cp01, cp13, 0, Sum(Round(" & Pt2EPtCode & ",2)),'" & strUserNum & "','4',0,0,0 FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      'end 2014/3/20
   End If
   
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id,R102015,R102016,R102017,R102018) " & strSql
   
'Added by Morgan 2013/5/22
Else
   '***發文目標達成***
   strSql = "SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'Modified by Morgan 2018/11/6 不用抓分配點數
   'strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(nvl(a0n03/1000,CP18)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "','1' FROM CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'end 2018/11/6
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT cp01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql & " union all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+)  " & strSQL2 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'Modified by Morgan 2018/11/6 不用抓分配點數
   'strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(nvl(a0n03/1000,CP18)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "','1' FROM CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   'end 2018/11/6
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,STAFF WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=ST01(+) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,STAFF WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=ST01(+) " & StrSQL4 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14,'" & strUserNum & "' "

   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id) " & strSql
End If
'end 2013/5/22



If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22
   'Modified by Morgan 2018/11/6 +R102015:來源 >> 1=案件, 2=支援, 3=特殊, 4=收文
   'Modified by Morgan 2018/11/7 +R102016:分配點數, R102017 不含加乘的基數, R102018 件數

   '***完稿達成***
   'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   '               strSql = " SELECT CP01,ep05,SUM(CP18),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(Txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = " SELECT CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum( Decode(EP09, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1) * decode(cp01,'CFP',0.5,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)* decode(cp01,'CFP',0.5,1)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   'strSql = strSql & " union all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(Txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql & " union all  SELECT 'ALL' As CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   'edit by nickc 2006/02/22 可以不用
   'cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id) " & strSQL


   '支援記錄
   'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
   strSql = strSql & " union all SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'P' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0.5,1) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
                   strSql = strSql & " union all SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, TRADEMARK, STAFF WHERE SH06=TM01(+) AND SH07=TM02(+) AND SH08=TM03(+) AND SH09=TM04(+) AND SH02=ST01 " & strSQL2 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL1 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & StrSQL3 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, HIRECASE, STAFF WHERE SH06=HC01(+) AND SH07=HC02(+) AND SH08=HC03(+) AND SH09=HC04(+) AND SH02=ST01 " & StrSQL4 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL5 & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02, '" & strUserNum & "' "
   'Modified by Morgan 2019/3/6
   'strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(Decode('P', 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY 'P', SH02, '" & strUserNum & "' "
   'end 2019/3/6
   'end 2014/3/20
   
   '特殊案件記錄
   strSql = strSql & " union all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL1 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL2 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL3 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL4 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL5 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql & " union all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL1 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL2 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL3 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL4 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL5 & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "


   'Add by Morgan 2010/10/18
   '2012/1/20 MODIFY BY SONIA 加部門別條件,否則統計商標處的會出現P的資料
   'If bolNewPromoterRule Then
   'Modified by Morgan 2012/4/20 部門條件改放 strSQL12 控制否則只輸員工編號但沒有輸部門條件查詢時會沒統計到
   'If bolNewPromoterRule And Left(txt1(10), 2) = "P1" Then
   If bolNewPromoterRule Then
      'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
      'strSql = strSql + " UNION all  SELECT 'P' As cp01, cp13, 0, Sum(Round(nvl(a0n03/1000,cp18)*0.05 ,2)),'" & strUserNum & "' FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      'strSql = strSql + " UNION all  SELECT 'ALL' As cp01, cp13, 0, Sum(Round(nvl(a0n03/1000,cp18)*0.05 ,2)),'" & strUserNum & "' FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      strSql = strSql + " UNION all  SELECT 'P' As cp01, cp13, 0, Sum(Round(" & Pt2EPtCode & ",2)),'" & strUserNum & "','4',0,0,0 FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      strSql = strSql + " UNION all  SELECT 'ALL' As cp01, cp13, 0, Sum(Round(" & Pt2EPtCode & ",2)),'" & strUserNum & "','4',0,0,0 FROM caseprogress, STAFF,acc0n0 where a0n02(+)=cp09 and ST01(+)=cp13 " & strSQL12 & " and cp20 is null and cp57 is null GROUP BY 'P', cp13, '" & strUserNum & "' "
      'end 2014/3/20
   End If
   
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id,R102015,R102019,R102020,R102021) " & strSql
   
'Added by Morgan 2013/5/22
Else
   '***完稿達成***
   'Modified by Morgan 2018/11/6 不用抓分配點數
   'strSql = " SELECT CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "','1' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = " SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   'end 2018/11/6
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   'Modified by Morgan 2018/11/6 不用抓分配點數
   'strSql = strSql & " union all  SELECT 'ALL' As CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "','1' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql & " union all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL1 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   'end 2018/11/6
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,TRADEMARK,STAFF WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=ST01(+) " & strSQL2 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 2) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,LAWCASE,STAFF WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=ST01(+) " & StrSQL3 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,HIRECASE,STAFF WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=ST01(+) " & StrSQL4 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 4) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL5 & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05,'" & strUserNum & "' "

   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id) " & strSql
End If
'end 2013/5/22

'***整理***
strSql = "SELECT R102001,R102002," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102010,0,0,NULL,0,R102010))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)+ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102011,0,0,NULL,0,R102011))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2))/2,'" & strUserNum & "' " & _
         "FROM R090608 WHERE ID='" & strUserNum & "' GROUP BY R102001,R102002,'" & strUserNum & "' "
cnnConnection.Execute "INSERT INTO R090608 (R102001,R102002,R102007,R102008,R102009,R102012,R102013,R102014,ID) " & strSql

End Sub

'Added by Morgan 2019/3/22 專利處108考核
'改自ProcessNew2
Sub ProcessNew3()
   Dim stAcc1u0 As String
   
   cnnConnection.Execute "DELETE FROM R090608 WHERE ID='" & strUserNum & "' "
   
   StrSQL6 = "": StrSQL7 = "": strSQL8 = "": strSQL10 = ""
   txt1(0) = "ALL"
   If txt1(0) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/14
   End If
   strSQL1 = ""
   '專利處不會有國家條件,已取消,語法可簡化
   
   StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
   StrSQL7 = StrSQL7 + " AND EP09>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP09<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
   strSQL8 = strSQL8 + " AND SH01>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND SH01<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/14
   If Len(txt1(5)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND CP14='" & txt1(5) & "' "
       StrSQL7 = StrSQL7 + " AND EP05='" & txt1(5) & "' "
       strSQL8 = strSQL8 + " AND SH02='" & txt1(5) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0) 'Add By Sindy 2010/12/14
   End If
   If Len(txt1(6)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND ST06>='" & txt1(6) & "' "
       StrSQL7 = StrSQL7 + " AND ST06>='" & txt1(6) & "' "
       strSQL8 = strSQL8 + " AND ST06>='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND ST06<='" & txt1(7) & "' "
       StrSQL7 = StrSQL7 + " AND ST06<='" & txt1(7) & "' "
       strSQL8 = strSQL8 + " AND ST06<='" & txt1(7) & "' "
   End If
   If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/14
   End If
   '部門別
   If Me.txt1(10).Text <> "" Then
       StrSQL6 = StrSQL6 + " AND ST03>='" & txt1(10) & "' "
       StrSQL7 = StrSQL7 + " AND ST03>='" & txt1(10) & "' "
       strSQL8 = strSQL8 + " AND ST03>='" & txt1(10) & "' "
   End If
   If Me.txt1(11).Text <> "" Then
       StrSQL6 = StrSQL6 + " AND ST03<='" & txt1(11) & "' "
       StrSQL7 = StrSQL7 + " AND ST03<='" & txt1(11) & "' "
       strSQL8 = strSQL8 + " AND ST03<='" & txt1(11) & "' "
   End If
   If Me.txt1(10).Text <> "" Or Me.txt1(11).Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/14
   End If
   '只限在職人員的資料
   If Check3.Value = vbUnchecked Then 'Added by Morgan 2024/4/24
      StrSQL6 = StrSQL6 + " and ST04='1' "
      StrSQL7 = StrSQL7 + " and ST04='1' "
      strSQL8 = strSQL8 + " and ST04='1' "
   End If
   strSQL8 = strSQL8 + " and SH11='V' "
   strSQL10 = strSQL10 + " and SCR03='V' "

   CheckOC
   
   'Modified by Morgan 2019/8/7 還是統計 P,CFP 的達成情形 --王副總
   
   '***目標***
   'Modified by Morgan 2014/9/24 目標改抓基數設定值
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id)" & _
      " Select 'ALL',pe01,sum(pe06),sum(er03),'" & strUserNum & "' from (select pe01,pe03,sum(pe06) pe06" & _
      " From Performance, Staff Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & _
      " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & _
      IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & _
      IIf(txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
      IIf(txt1(10).Text <> "", " And ST03>='" & txt1(10).Text & "' ", "") & IIf(txt1(11).Text <> "", " And ST03<='" & txt1(11).Text & "' ", "") & _
      IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe01,pe03),engradix where er01(+)=pe03 and er02(+)=pe01 group by pe01"
   cnnConnection.Execute strSql
   
   'Added by Morgan 2019/8/7
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id)" & _
      " Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "' From Performance, Staff" & _
      " Where PE01=ST01(+) And pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & _
      " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & _
      IIf(txt1(6).Text <> "", " AND ST06>='" & txt1(6) & "' ", "") & IIf(Me.txt1(7).Text <> "", " AND ST06<='" & txt1(7) & "' ", "") & _
      IIf(txt1(10).Text <> "", " And ST03>='" & txt1(10).Text & "' ", "") & IIf(txt1(11).Text <> "", " And ST03<='" & Me.txt1(11).Text & "' ", "") & _
      IIf(Check3.Value = vbUnchecked, " And ST04='1'", "") & " group by pe02,pe01"
   cnnConnection.Execute strSql
   
   
   m_bolShowMemo = True
   
   'Modified by Morgan 2019/3/22 +R102022:實績點數,R102023:實績點數達成率(一併修正未扣除銷帳點數問題)
   'Added by Morgan 2019/3/25 銷帳金額
   'Modified by Morgan 2023/1/7 銷帳服務費可能為負(扣點數) A1U07>0->A1U07<>0
   stAcc1u0 = "SELECT A1U03,SUM(A1U07) AS A1U07 FROM CASEPROGRESS,STAFF,ACC1U0 WHERE ST01(+)=CP14 AND A1U03(+)=CP09 AND A1U07<>0 " & StrSQL6 & " GROUP BY A1U03"
   'end 2019/3/25
   If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22
      'Modified by Morgan 2018/11/6 +R102015:來源 >> 1=案件, 2=支援, 3=特殊, 4=收文
      'Modified by Morgan 2018/11/7 +R102016:分配點數,R102017:不含加乘的基數,R102018:件數
      
      '***發文目標達成***
      'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
      strSql = " SELECT CP01,cp14,SUM(nvl(a0n03/1000,CP18-nvl(a1u07/1000,0))),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),SUM(decode(cp26,'N',0,nvl(a0n03/1000,CP18-nvl(a1u07/1000,0)))) FROM CASEPROGRESS,PATENT,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT CP01,cp14,SUM(CP18-nvl(a1u07/1000,0)),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all SELECT 'ALL' As CP01,cp14,SUM(nvl(a0n03/1000,CP18-nvl(a1u07/1000,0))),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),SUM(decode(cp26,'N',0,nvl(a0n03/1000,CP18-nvl(a1u07/1000,0)))) FROM CASEPROGRESS,PATENT,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18-nvl(a1u07/1000,0)),Sum( Decode(CP27, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum(Decode(CP27, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
   
      '支援記錄
      'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
      'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
      'Modified by Morgan 2021/7/13 +ACS Ex:ACS-000108 (支援記錄:1100630001)
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'P', SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
      
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
      
      '特殊案件記錄
      strSql = strSql + " UNION all  SELECT CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
   
      '專利處成員的業務收文點數折算--108考核取消
      
      cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id,R102015,R102016,R102017,R102018,R102022) " & strSql
   
   '只統計件數
   Else
      '***發文目標達成***
      'Modified by Morgan 2018/11/6 不用抓分配點數
      strSql = " SELECT CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,PATENT,STAFF,(" & stAcc1u0 & ") A1U0 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,(" & stAcc1u0 & ") A1U0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
      
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,PATENT,STAFF,(" & stAcc1u0 & ") A1U0 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18),Sum(Decode(CP26, Null, Decode(CP27, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,(" & stAcc1u0 & ") A1U0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
   
      cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id,R102022) " & strSql
   End If

   If Check1.Value = vbUnchecked Then 'Added by Morgan 2013/5/22
      'Modified by Morgan 2018/11/6 +R102015:來源 >> 1=案件, 2=支援, 3=特殊, 4=收文
      'Modified by Morgan 2018/11/7 +R102016:分配點數, R102017 不含加乘的基數, R102018 件數
   
      '***完稿達成***
      'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
      strSql = " SELECT CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all SELECT CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
      
      strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum( Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(EP09, Null, 0 ,cp97 * cp98 * decode(cp112,'Y',cp111,1)) ),'" & strUserNum & "','1',0,Sum( Decode(EP09, Null, 0 ,cp97)),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
      '支援記錄
      'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
      'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
      'Modified by Morgan 2021/7/13 +ACS Ex:ACS-000108 (支援記錄:1100630001)
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'P', SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
      
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
      strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
      
      '特殊案件記錄
      strSql = strSql & " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
      strSql = strSql & " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
      '專利處成員的業務收文點數折算
      '108考核(取消)
      
      cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id,R102015,R102019,R102020,R102021) " & strSql
      
   '只統計件數
   Else
      '***完稿達成***
      'Modified by Morgan 2018/11/6 不用抓分配點數
      strSql = " SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all SELECT CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
      
      strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
      strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(Decode(CP26, Null, Decode(EP09, Null, 0 ,1) , 0)),'" & strUserNum & "',SUM(decode(cp26,'N',0,CP18)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
      cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id,R102021) " & strSql
      
   End If
   
   '***整理***
   '發文實績點數達成率(R102023)=發文實績點數(R102022)/點數目標(R102003)
   '發文點數達成率(R102007)=發文點數(R102005)/點數目標(R102003)
   '發文基數數達成率(R102008)=發文基數(R102006)/基數目標(R102004)
   '發文平均達成率(R102009)=(發文實績點數達成率(R102023)+發文基數數達成率(R102008))/2
   '完稿點數達成率(R102012)=完稿點數(R102010)/點數目標(R102003)
   '完稿基數數達成率(R102013)=完稿基數(R102011)/基數目標(R102004)
   '完稿平均達成率(R102014)=(完稿點數達成率(R102012)+完稿基數數達成率(R102013))/2
   strSql = "SELECT R102001,R102002," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102022,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102005,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102006,0))/SUM(NVL(R102004,0)))*100),2)," & _
            "(ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102022,0))/SUM(NVL(R102003,0)))*100),2)+ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102006,0))/SUM(NVL(R102004,0)))*100),2))/2," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102010,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102011,0))/SUM(NVL(R102004,0)))*100),2)," & _
            "(ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102010,0))/SUM(NVL(R102003,0)))*100),2)+ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102011,0))/SUM(NVL(R102004,0)))*100),2))/2," & _
            "ID FROM R090608 WHERE ID='" & strUserNum & "' GROUP BY ID,R102001,R102002 "
   cnnConnection.Execute "INSERT INTO R090608 (R102001,R102002,R102023,R102007,R102008,R102009,R102012,R102013,R102014,ID) " & strSql

End Sub

'Added by Morgan 2025/7/10
'點數
Sub ProcessNew4()
   Dim stAcc1u0 As String
   Dim stDate1 As String, stDate2 As String
   
   cnnConnection.Execute "DELETE FROM R090608 WHERE ID='" & strUserNum & "' ", intI
      
   stDate1 = Val(ChangeTStringToWString(txt1(3) & "01"))
   stDate2 = Val(ChangeTStringToWString(txt1(4) & "31"))
   
   StrSQL4 = "": strSQL5 = "": StrSQL6 = "": StrSQL7 = "": strSQL8 = "": strSQL10 = ""
   txt1(0) = "ALL"
   If txt1(0) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)
   End If
   strSQL1 = ""
   
   '發文年月
   strSQL5 = strSQL5 + " AND PE03>=" & Left(stDate1, 6) & " AND PE03<=" & Left(stDate2, 6)
   StrSQL6 = StrSQL6 + " AND CP27>=" & stDate1 & " AND CP27<=" & stDate2
   StrSQL7 = StrSQL7 + " AND EP09>=" & stDate1 & " AND EP09<=" & stDate2
   strSQL8 = strSQL8 + " AND SH01>=" & stDate1 & " AND SH01<=" & stDate2
   strSQL9 = strSQL9 + " AND EP07>=" & stDate1 & " AND EP07<=" & stDate2
   
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4)
   '承辦人
   If Len(txt1(5)) <> 0 Then
      strSQL5 = strSQL5 + " AND PE01='" & txt1(5) & "' "
      StrSQL6 = StrSQL6 + " AND CP14='" & txt1(5) & "' "
      StrSQL7 = StrSQL7 + " AND EP05='" & txt1(5) & "' "
      strSQL8 = strSQL8 + " AND SH02='" & txt1(5) & "' "
      strSQL9 = strSQL9 + " AND EP05='" & txt1(5) & "' "
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0)
   End If
   '所別
   If Len(txt1(6)) <> 0 Then
      StrSQL4 = StrSQL4 + " AND ST06>='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
      StrSQL4 = StrSQL4 + " AND ST06<='" & txt1(7) & "' "
   End If
   
   If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7)
   End If
   '部門別
   If Me.txt1(10).Text <> "" Then
      StrSQL4 = StrSQL4 + " AND ST03>='" & txt1(10) & "' "
   End If
   If Me.txt1(11).Text <> "" Then
      StrSQL4 = StrSQL4 + " AND ST03<='" & txt1(11) & "' "
   End If
   If Me.txt1(10).Text <> "" Or Me.txt1(11).Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(10) & "-" & txt1(11)
   End If
   
   '只限在職人員的資料
   If Check3.Value = vbUnchecked Then
      StrSQL4 = StrSQL4 + " and ST04='1' "
   End If
   
   strSQL5 = strSQL5 + StrSQL4
   StrSQL6 = StrSQL6 + StrSQL4
   StrSQL7 = StrSQL7 + StrSQL4
   strSQL8 = strSQL8 + " and SH11='V' " + StrSQL4
   strSQL9 = strSQL9 + StrSQL4
   strSQL10 = strSQL10 + " and SCR03='V' "
   
   CheckOC
   
   'R102001:系統別 >> P,CFP,ALL
   'R102002:承辦人
   'R102003:目標點數
   'R102004:目標基數
   
   '***目標***
   '目標改抓基數設定值
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id)" & _
      " Select 'ALL',pe01,sum(pe06),sum(er03),'" & strUserNum & "' from (select pe01,pe03,sum(pe06) pe06" & _
      " From Performance, Staff Where PE01=ST01(+)" & strSQL5 & _
      " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & _
      " group by pe01,pe03),engradix where er01(+)=pe03 and er02(+)=pe01 group by pe01"
   cnnConnection.Execute strSql, intI
   
   strSql = "INSERT INTO R090608 (r102001,r102002,r102003,r102004,id)" & _
      " Select pe02,pe01,sum(pe06),sum(pe05),'" & strUserNum & "'" & _
      " From Performance, Staff Where PE01=ST01(+)" & strSQL5 & _
      " and pe02 in ('P','CFP') and pe02 in (" & GetAddStr(GetAllSysKind(txt1(0))) & ") " & _
      " group by pe02,pe01"
   cnnConnection.Execute strSql, intI
   
   m_bolShowMemo = False
   
   '銷帳服務費
   '先統一都不扣銷帳
   'stAcc1u0 = "SELECT A1U03,SUM(A1U07) AS A1U07 FROM CASEPROGRESS,STAFF,ACC1U0 WHERE ST01(+)=CP14 AND A1U03(+)=CP09 AND A1U07<>0 " & StrSQL6 & " GROUP BY A1U03"
   stAcc1u0 = "SELECT '' A1U03,0 AS A1U07 FROM DUAL"
   
   'R102005:發文點數
   'R102006:發文基數
   'R102015:來源 >> 1=案件, 2=支援, 3=特殊
   'R102016:發文分配點數
   'R102017:發文不含加乘的基數
   'R102018:發文件數
   'R102022:發文計件點數
   
   '***發文目標達成***
   '若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   strSql = " SELECT CP01,cp14,SUM(nvl(a0n03/1000,CP18-nvl(a1u07/1000,0))),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26, Null,1,0)),SUM(decode(cp26,'N',0,nvl(a0n03/1000,CP18-nvl(a1u07/1000,0)))) FROM CASEPROGRESS,PATENT,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all  SELECT CP01,cp14,SUM(CP18-nvl(a1u07/1000,0)),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26, Null,1,0)),SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all SELECT 'ALL' As CP01,cp14,SUM(nvl(a0n03/1000,CP18-nvl(a1u07/1000,0))),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26, Null,1,0)),SUM(decode(cp26,'N',0,nvl(a0n03/1000,CP18-nvl(a1u07/1000,0)))) FROM CASEPROGRESS,PATENT,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14,SUM(CP18-nvl(a1u07/1000,0)),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26, Null,1,0)),SUM(decode(cp26,'N',0,CP18-nvl(a1u07/1000,0))) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,acc0n0,(" & stAcc1u0 & ") A1U0 WHERE a0n02(+)=cp09 and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND A1U03(+)=CP09 " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"

   '支援記錄
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'P', SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   '特殊案件記錄
   strSql = strSql + " UNION all  SELECT CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all  SELECT CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,cp14"
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,cp14, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0,0 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) And CP09=SCR01 " & StrSQL6 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,cp14"

   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102005,r102006,id,R102015,R102016,R102017,R102018,R102022) " & strSql, intI
      
      
   'R102010:完稿點數
   'R102011:完稿基數
   'R102015:來源 >> 1=案件, 2=支援, 3=特殊
   'R102019:完稿分配點數
   'R102020:完稿不含加乘的基數
   'R102021:完稿件數
   
   '***完稿達成***
   '若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   strSql = " SELECT CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26, Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all SELECT CP01,ep05,SUM(CP18),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26, Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
   strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26, Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26, Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & StrSQL7 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"

   '支援記錄
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'P', SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   '特殊案件記錄
   strSql = strSql & " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"

   strSql = strSql & " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & StrSQL7 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
      
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102010,r102011,id,R102015,R102019,R102020,R102021) " & strSql, intI
      
   'R102024:會稿點數
   'R102025:會稿基數
   'R102015:來源 >> 1=案件, 2=支援, 3=特殊
   'R102026:會稿分配點數
   'R102027:會稿不含加乘的基數
   'R102028:會稿件數
   
   '***會稿達成***
   '若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   strSql = " SELECT CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26,Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL9 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all SELECT CP01,ep05,SUM(CP18),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26, Null, 1 , 0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL9 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
   
   strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(nvl(a0n03/1000,CP18)),Sum(cp97 * cp98),'" & strUserNum & "','1',SUM(nvl(a0n03/1000,0)),Sum(cp97),Sum(Decode(CP26, Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF,acc0n0 WHERE a0n02(+)=cp09 and  EP02=CP09(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) " & strSQL9 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all SELECT 'ALL' As CP01,ep05,SUM(CP18),Sum(cp97 * cp98),'" & strUserNum & "','1',0,Sum(cp97),Sum(Decode(CP26,Null,1,0)) FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) " & strSQL9 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"

   '支援記錄
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'P', SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, PATENT, STAFF WHERE SH06=PA01(+) AND SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, SERVICEPRACTICE, STAFF WHERE SH06=SP01(+) AND SH07=SP02(+) AND SH08=SP03(+) AND SH09=SP04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, LAWCASE, STAFF WHERE SH06=LC01(+) AND SH07=LC02(+) AND SH08=LC03(+) AND SH09=LC04(+) AND SH02=ST01 " & strSQL8 & " AND SH06 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 3) & ") GROUP BY SH06, SH02"
   strSql = strSql + " UNION all  SELECT 'ALL' As SH06, SH02, 0, Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),'" & strUserNum & "','2',0,0,0 FROM SupportHour, STAFF WHERE SH02=ST01 " & strSQL8 & "  And SH06 Is Null GROUP BY SH06, SH02"
   
   '特殊案件記錄
   strSql = strSql & " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL9 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all  SELECT CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL9 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"

   strSql = strSql & " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL9 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 1) & ") GROUP BY cp01,ep05"
   strSql = strSql + " UNION all  SELECT 'ALL' As CP01,ep05, 0, Sum(Nvl(SCR02,0)),'" & strUserNum & "','3',0,0,0 FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF, SpecialCaseRecord WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=ST01(+) And EP02=SCR01 " & strSQL9 & strSQL10 & " AND CP01 IN (" & SQLGrpStr(GetAllSysKind(txt1(0)), 5) & ") GROUP BY cp01,ep05"
      
   cnnConnection.Execute "insert into r090608 (r102001,r102002,r102024,r102025,id,R102015,R102026,R102027,R102028) " & strSql, intI
   
   '***整理***
   '發文點數達成率(R102007)=發文點數(R102005)/點數目標(R102003)
   '發文基數數達成率(R102008)=發文基數(R102006)/基數目標(R102004)
   '完稿點數達成率(R102012)=完稿點數(R102010)/點數目標(R102003)
   '完稿基數數達成率(R102013)=完稿基數(R102011)/基數目標(R102004)
   '會稿點數達成率(R102029)=會稿點數(R102024)/點數目標(R102003)
   '會稿基數數達成率(R102030)=會稿基數(R102025)/基數目標(R102004)
   strSql = "SELECT R102001,R102002," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102005,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102006,0))/SUM(NVL(R102004,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102010,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102011,0))/SUM(NVL(R102004,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102003,0)),0,0,(SUM(NVL(R102024,0))/SUM(NVL(R102003,0)))*100),2)," & _
            "ROUND(DECODE(SUM(NVL(R102004,0)),0,0,(SUM(NVL(R102025,0))/SUM(NVL(R102004,0)))*100),2)," & _
            "ID FROM R090608 WHERE ID='" & strUserNum & "' GROUP BY ID,R102001,R102002 "
   cnnConnection.Execute "INSERT INTO R090608 (R102001,R102002,R102007,R102008,R102012,R102013,R102029,R102030,ID) " & strSql, intI

End Sub

