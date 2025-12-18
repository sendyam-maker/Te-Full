VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010613_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件狀況資料"
   ClientHeight    =   6170
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6170
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdNext 
      Cancel          =   -1  'True
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6180
      TabIndex        =   23
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOpenF 
      Caption         =   "開啟郵件"
      Default         =   -1  'True
      Height          =   360
      Left            =   7020
      TabIndex        =   0
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7980
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm06010613_1.frx":0000
      Height          =   3230
      Left            =   30
      TabIndex        =   2
      Top             =   2700
      Width           =   8870
      _ExtentX        =   15646
      _ExtentY        =   5697
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "收受者|讀取日期時間|讀取人員|刪除日期時間|刪除人員|轉寄日期時間|轉寄者"
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
      _Band(0).Cols   =   7
   End
   Begin VB.Frame FraII27 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   230
      Left            =   2952
      TabIndex        =   36
      Top             =   1920
      Width           =   5916
      Begin VB.Label Lblii29 
         Caption         =   "外商處理結果："
         Height          =   216
         Left            =   2808
         TabIndex        =   40
         Top             =   36
         Width           =   1320
      End
      Begin MSForms.TextBox txtII29 
         Height          =   288
         Left            =   4164
         TabIndex        =   39
         Top             =   0
         Width           =   1404
         VariousPropertyBits=   680544287
         BackColor       =   -2147483633
         ScrollBars      =   3
         Size            =   "2469;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Lblii27 
         Caption         =   "外專處理結果："
         Height          =   216
         Left            =   0
         TabIndex        =   38
         Top             =   36
         Width           =   1320
      End
      Begin MSForms.TextBox txtII27 
         Height          =   288
         Left            =   1356
         TabIndex        =   37
         Top             =   0
         Width           =   1404
         VariousPropertyBits=   680544287
         BackColor       =   -2147483633
         ScrollBars      =   3
         Size            =   "2469;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Label LblPI23 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Left            =   180
      TabIndex        =   41
      Top             =   2190
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSForms.TextBox txtPI23 
      Height          =   290
      Left            =   1110
      TabIndex        =   31
      Top             =   2190
      Visible         =   0   'False
      Width           =   7760
      VariousPropertyBits=   -1466939361
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "13688;512"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblReceiver 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "LblReceiver"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Left            =   1870
      TabIndex        =   29
      Top             =   2500
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.Label Label12 
      Caption         =   "註：信件自”可刪除日期”起算一個月後刪除。"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   80
      TabIndex        =   35
      Top             =   5970
      Width           =   7640
   End
   Begin VB.Label Lblii28 
      Caption         =   "回信沖銷亂碼："
      Height          =   210
      Left            =   180
      TabIndex        =   34
      Top             =   1950
      Width           =   1310
   End
   Begin MSForms.TextBox txtII28 
      Height          =   290
      Left            =   1540
      TabIndex        =   33
      Top             =   1900
      Width           =   1360
      VariousPropertyBits=   747653151
      BackColor       =   -2147483633
      BorderStyle     =   1
      ScrollBars      =   3
      Size            =   "2399;512"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtII11 
      Height          =   230
      Left            =   1110
      TabIndex        =   32
      Top             =   1050
      Width           =   7760
      VariousPropertyBits=   -1466939361
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "13688;406"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtII17 
      Height          =   560
      Left            =   1110
      TabIndex        =   42
      Top             =   1320
      Width           =   7760
      VariousPropertyBits=   -1466939361
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "13688;988"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblNote 
      Caption         =   "系統記錄"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   4590
      TabIndex        =   30
      Top             =   840
      Width           =   4280
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "其他信箱的收受者："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Left            =   80
      TabIndex        =   28
      Top             =   2500
      Visible         =   0   'False
      Width           =   1760
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "信件編號："
      Height          =   180
      Left            =   2340
      TabIndex        =   27
      Top             =   420
      Width           =   900
   End
   Begin VB.Label LblII03 
      AutoSize        =   -1  'True
      Caption         =   "LblII03"
      Height          =   180
      Left            =   3270
      TabIndex        =   26
      Top             =   420
      Width           =   540
   End
   Begin VB.Label LblPI18_T 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   420
      Width           =   900
   End
   Begin VB.Label LblPI18 
      AutoSize        =   -1  'True
      Caption         =   "LblPI18"
      Height          =   180
      Left            =   1110
      TabIndex        =   24
      Top             =   420
      Width           =   570
   End
   Begin VB.Label LblII08_T 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "未轉寄刪除日期："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   60
      TabIndex        =   22
      Top             =   90
      Width           =   1440
   End
   Begin VB.Label LblII08 
      Caption         =   "LblII08"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   1530
      TabIndex        =   21
      Top             =   90
      Width           =   825
   End
   Begin VB.Label LblII09 
      Caption         =   "LblII09"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   2400
      TabIndex        =   20
      Top             =   90
      Width           =   1035
   End
   Begin VB.Label LblII10_T 
      Alignment       =   1  '靠右對齊
      Caption         =   "刪除人員："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3510
      TabIndex        =   19
      Top             =   90
      Width           =   900
   End
   Begin VB.Label LblII10 
      Caption         =   "LblII10"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4440
      TabIndex        =   18
      Top             =   90
      Width           =   1665
   End
   Begin VB.Label LblII16 
      Caption         =   "LblII16"
      Height          =   180
      Left            =   6720
      TabIndex        =   17
      Top             =   630
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "寄件者："
      Height          =   180
      Left            =   350
      TabIndex        =   16
      Top             =   1050
      Width           =   740
   End
   Begin VB.Label LblII13 
      AutoSize        =   -1  'True
      Caption         =   "LblII13"
      Height          =   180
      Left            =   3270
      TabIndex        =   15
      Top             =   840
      Width           =   540
   End
   Begin VB.Label LblII12 
      AutoSize        =   -1  'True
      Caption         =   "LblII12"
      Height          =   180
      Left            =   1110
      TabIndex        =   14
      Top             =   840
      Width           =   540
   End
   Begin VB.Label LblII05 
      AutoSize        =   -1  'True
      Caption         =   "LblII05"
      Height          =   180
      Left            =   5520
      TabIndex        =   13
      Top             =   630
      Width           =   540
   End
   Begin VB.Label LblII04 
      AutoSize        =   -1  'True
      Caption         =   "LblII04"
      Height          =   180
      Left            =   5520
      TabIndex        =   12
      Top             =   420
      Width           =   540
   End
   Begin VB.Label LblII02 
      AutoSize        =   -1  'True
      Caption         =   "LblII02"
      Height          =   180
      Left            =   3270
      TabIndex        =   11
      Top             =   630
      Width           =   540
   End
   Begin VB.Label LblII01 
      AutoSize        =   -1  'True
      Caption         =   "LblII01"
      Height          =   180
      Left            =   1110
      TabIndex        =   10
      Top             =   630
      Width           =   540
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "主　　旨："
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "信件時間："
      Height          =   180
      Left            =   2340
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "轉入人員："
      Height          =   180
      Left            =   4590
      TabIndex        =   7
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "信件日期："
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "轉入日期："
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "分　　類："
      Height          =   180
      Left            =   4590
      TabIndex        =   4
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "轉入時間："
      Height          =   180
      Left            =   2340
      TabIndex        =   3
      Top             =   630
      Width           =   900
   End
End
Attribute VB_Name = "frm06010613_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/16 Form2.0已修改
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim m_PrevForm As Form '前一畫面
Public m_II01 As String
Public m_II02 As String
Public m_II03 As String
Public m_II19 As String
Dim m_II14 As String 'Add By Sindy 2016/10/4


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
   End If
   Unload Me
End Sub

'Modify By Sindy 2016/9/22
Private Function ReadIPDeptInput() As Boolean
Dim rsTmp As New ADODB.Recordset
   
'   LblPI18_T.Visible = False
'   LblPI18.Visible = False
   ReadIPDeptInput = True
   '國外部信件主檔
   strSql = "select IPDeptinput.*,decode(ii27," & 外專信件處理結果 & ",ii27) as ii27txt,decode(ii29," & 外專信件處理結果 & ",ii29) as ii29txt" & _
            " From IPDeptinput" & _
            " where ii01='" & m_II01 & "' and ii02='" & m_II02 & "' and ii03='" & ChgSQL(m_II03) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Me.LblII01 = ChangeWStringToTDateString(m_II01)
      Me.LblII02 = Format(m_II02, "##:##:##")
      Me.LblII03 = "" & rsTmp.Fields("ii03") & IIf("" & rsTmp.Fields("ii15") <> "", " (" & "" & rsTmp.Fields("ii15") & ")", "") 'Add By Sindy 2017/12/25
      Me.txtII17 = "" & rsTmp.Fields("ii17")
      Me.txtII28 = "" & rsTmp.Fields("ii28") 'Add By Sindy 2022/8/12
      Me.txtII11 = "" & rsTmp.Fields("ii11")
      Me.LblII04 = rsTmp.Fields("ii04") & " " & GetPrjSalesNM(rsTmp.Fields("ii04"))
      Me.txtII27 = Trim("" & rsTmp.Fields("ii27txt")) 'Add By Sindy 2023/5/23
      Me.txtII29 = Trim("" & rsTmp.Fields("ii29txt")) 'Add By Sindy 2023/7/13
      
      'Add By Sindy 2020/6/15
'      Me.LblNote.Visible = False
'      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         LblNote = "系統記錄:" & rsTmp.Fields("ii18")
'         LblNote.Visible = True
'      End If
      '2020/6/15 END
      
      If "" & rsTmp.Fields("ii05") = "" Then
         Me.LblII05 = ""
      Else
         Select Case rsTmp.Fields("ii05")
            Case "1"
               LblII05 = "個案"
            Case "2"
               LblII05 = "外商"
            Case "3"
               LblII05 = "外專"
            Case "4"
               LblII05 = "專利處"
            Case "5"
               LblII05 = "外法"
            Case "6"
               LblII05 = "新知"
            Case "7"
               LblII05 = "財務"
            'Add By Sindy 2016/6/15
            Case "8"
               LblII05 = "開拓"
            '2016/6/15 END
            Case Else
               LblII05 = "其他"
         End Select
      End If
'      If "" & rsTmp.Fields("ii07") = "Y" Then '刪除
'         'Modify By Sindy 2016/5/17
''         Me.LblII08_T = "刪除日期："
''         Me.LblII09_T = "刪除時間："
''         Me.LblII10_T = "刪除人員："
'         Me.LblII08_T.Visible = True
'         Me.LblII08.Visible = True
'         Me.LblII09.Visible = True
'         Me.LblII10_T.Visible = True
'         Me.LblII10.Visible = True
      'Modify By Sindy 2019/7/1
      If "" & rsTmp.Fields("ii07") = "Y" Then '刪除
         Me.LblII10_T = "刪除人員:"
         Me.LblII08_T = "未轉寄刪除日期:"
      Else
         Me.LblII10_T = "轉寄人員:"
         Me.LblII08_T = "轉寄日期:"
      End If
      '2019/7/1 END
         If Val("" & rsTmp.Fields("ii08")) = 0 Then
            Me.LblII08 = ""
         Else
            Me.LblII08 = ChangeWStringToTDateString(rsTmp.Fields("ii08"))
         End If
         If "" & rsTmp.Fields("ii09") = "" Then
            Me.LblII09 = ""
         Else
            Me.LblII09 = Format(rsTmp.Fields("ii09"), "##:##:##")
         End If
         If "" & rsTmp.Fields("ii10") = "" Then
            Me.LblII10 = ""
         Else
            Me.LblII10 = rsTmp.Fields("ii10") & " " & GetPrjSalesNM(rsTmp.Fields("ii10"))
         End If
         '2016/5/17 END
'      End If
      If Val("" & rsTmp.Fields("ii12")) = 0 Then
         Me.LblII12 = ""
      Else
         Me.LblII12 = ChangeWStringToTDateString(rsTmp.Fields("ii12"))
      End If
      If "" & rsTmp.Fields("ii13") = "" Then
         Me.LblII13 = ""
      Else
         Me.LblII13 = Format(rsTmp.Fields("ii13"), "##:##:##")
      End If
      If "" & rsTmp.Fields("ii14") = "" Then
         cmdOpenF.Enabled = False
      Else
         m_II14 = rsTmp.Fields("ii14") 'Add By Sindy 2016/10/4
         cmdOpenF.Enabled = True
      End If
      'Add By Sindy 2016/4/29 Msg檔可刪除日期
      Me.LblII16.Visible = False
      'If Pub_StrUserSt03 = "M51" Then
         Me.LblII16.Visible = True
         Me.LblII16 = ""
         If Val("" & rsTmp.Fields("ii16")) > 0 Then
            Me.LblII16 = "可刪除日期:" & ChangeWStringToTDateString(rsTmp.Fields("ii16"))
         End If
      'End If
      'Add By Sindy 2022/7/21
      If "" & rsTmp.Fields("ii23") = "" Then
         Me.LblPI18 = ""
      Else
         Me.LblPI18 = rsTmp.Fields("ii23") & "-" & rsTmp.Fields("ii24") & "-" & rsTmp.Fields("ii25") & "-" & rsTmp.Fields("ii26")
      End If
      '2022/7/21 END
   Else
      ReadIPDeptInput = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2016/9/22
Private Function ReadPatentInput() As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ReadPatentInput = True
   '專利處信件主檔
   strSql = "select *" & _
            " From patentinput" & _
            " where pi01='" & m_II01 & "' and pi02='" & m_II02 & "' and pi03='" & m_II03 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Me.LblII01 = ChangeWStringToTDateString(m_II01)
      Me.LblII02 = Format(m_II02, "##:##:##")
      Me.LblII03 = "" & rsTmp.Fields("pi03") & IIf("" & rsTmp.Fields("pi22") <> "", " (" & "" & rsTmp.Fields("pi22") & ")", "") 'Add By Sindy 2017/12/25
      'Add By Sindy 2024/4/22
      LblPI23.Visible = True
      txtPI23.Visible = True
      Me.txtPI23 = "" & rsTmp.Fields("PI23")
      '2024/4/22 END
      Me.txtII17 = "" & rsTmp.Fields("pi17")
      Me.txtII28 = "" 'Add By Sindy 2022/8/12
      Me.txtII11 = "" & rsTmp.Fields("pi11")
      Me.LblII04 = rsTmp.Fields("pi04") & " " & GetPrjSalesNM(rsTmp.Fields("pi04"))
      
      'Add By Sindy 2020/6/15
'      Me.LblNote.Visible = False
'      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         LblNote = "系統記錄:" & rsTmp.Fields("pi15")
'         LblNote.Visible = True
'      End If
      '2020/6/15 END
      
      If "" & rsTmp.Fields("pi05") = "" Then
         Me.LblII05 = ""
      Else
         Select Case rsTmp.Fields("pi05")
            Case "1"
               LblII05 = "P程序1"
            Case "2"
               LblII05 = "P程序2"
            'Modify By Sindy 2018/6/21
            Case "3"
               LblII05 = "美日(單)"
            Case "4"
               LblII05 = "美日(雙)"
            Case "5"
               LblII05 = "美日外(單)"
            Case "6"
               LblII05 = "美日外(雙)"
            '2018/6/21 END
            Case "7"
               LblII05 = "其他"
            Case "8"
               LblII05 = "垃圾信箱"
            'Add By Sindy 2018/6/21
            Case "A"
               LblII05 = "亞洲"
            Case "B"
               LblII05 = "歐洲"
            Case "C"
               LblII05 = "美洋非(單)"
            Case "D"
               LblII05 = "美洋非(雙)"
            '2018/6/21 END
            'Add By Sindy 2020/3/18
            Case Else
               LblII05 = rsTmp.Fields("pi05")
            '2020/3/18 END
         End Select
      End If
      'Modify By Sindy 2019/7/1
      If "" & rsTmp.Fields("pi07") = "Y" Then '刪除
'         Me.LblII08_T.Visible = True
'         Me.LblII08.Visible = True
'         Me.LblII09.Visible = True
'         Me.LblII10_T.Visible = True
'         Me.LblII10.Visible = True
         Me.LblII10_T = "刪除人員:"
         Me.LblII08_T = "未轉寄刪除日期:"
      Else
         Me.LblII10_T = "轉寄人員:"
         Me.LblII08_T = "轉寄日期:"
      End If
      '2019/7/1 END
         If Val("" & rsTmp.Fields("pi08")) = 0 Then
            Me.LblII08 = ""
         Else
            Me.LblII08 = ChangeWStringToTDateString(rsTmp.Fields("pi08"))
         End If
         If "" & rsTmp.Fields("pi09") = "" Then
            Me.LblII09 = ""
         Else
            Me.LblII09 = Format(rsTmp.Fields("pi09"), "##:##:##")
         End If
         If "" & rsTmp.Fields("pi10") = "" Then
            Me.LblII10 = ""
         Else
            Me.LblII10 = rsTmp.Fields("pi10") & " " & GetPrjSalesNM(rsTmp.Fields("pi10"))
         End If
      'End If
      If Val("" & rsTmp.Fields("pi12")) = 0 Then
         Me.LblII12 = ""
      Else
         Me.LblII12 = ChangeWStringToTDateString(rsTmp.Fields("pi12"))
      End If
      If "" & rsTmp.Fields("pi13") = "" Then
         Me.LblII13 = ""
      Else
         Me.LblII13 = Format(rsTmp.Fields("pi13"), "##:##:##")
      End If
      If "" & rsTmp.Fields("pi14") = "" Then
         cmdOpenF.Enabled = False
      Else
         m_II14 = rsTmp.Fields("pi14") 'Add By Sindy 2016/10/4
         cmdOpenF.Enabled = True
      End If
      'Msg檔可刪除日期
      Me.LblII16.Visible = False
      'If Pub_StrUserSt03 = "M51" Then
         Me.LblII16.Visible = True
         Me.LblII16 = ""
         If Val("" & rsTmp.Fields("pi16")) > 0 Then
            Me.LblII16 = "可刪除日期:" & ChangeWStringToTDateString(rsTmp.Fields("pi16"))
         End If
      'End If
      If "" & rsTmp.Fields("pi18") = "" Then
         Me.LblPI18 = ""
      Else
         Me.LblPI18 = rsTmp.Fields("pi18") & "-" & rsTmp.Fields("pi19") & "-" & rsTmp.Fields("pi20") & "-" & rsTmp.Fields("pi21")
      End If
   Else
      ReadPatentInput = False
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/4/16
Private Function ReadTMInput() As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ReadTMInput = True
   '商標處信件主檔
   strSql = "select *" & _
            " From TMinput" & _
            " where Ti01='" & m_II01 & "' and Ti02='" & m_II02 & "' and Ti03='" & m_II03 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Me.LblII01 = ChangeWStringToTDateString(m_II01)
      Me.LblII02 = Format(m_II02, "##:##:##")
      Me.LblII03 = "" & rsTmp.Fields("Ti03") & IIf("" & rsTmp.Fields("Ti22") <> "", " (" & "" & rsTmp.Fields("Ti22") & ")", "") 'Add By Sindy 2017/12/25
      Me.txtII17 = "" & rsTmp.Fields("Ti17")
      Me.txtII28 = "" 'Add By Sindy 2022/8/12
      Me.txtII11 = "" & rsTmp.Fields("Ti11")
      Me.LblII04 = rsTmp.Fields("Ti04") & " " & GetPrjSalesNM(rsTmp.Fields("Ti04"))
      
      'Add By Sindy 2020/6/15
'      Me.LblNote.Visible = False
''      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         LblNote = "系統記錄:" & rsTmp.Fields("Ti15")
'         LblNote.Visible = True
'      End If
      '2020/6/15 END
      
      If "" & rsTmp.Fields("Ti05") = "" Then
         Me.LblII05 = ""
      Else
         Select Case rsTmp.Fields("Ti05")
            Case "1"
               LblII05 = "MCTF"
            Case "2"
               LblII05 = "大陸案"
            Case "3"
               LblII05 = "個人"
            Case "4"
               LblII05 = "非大陸案"
            Case "5"
               LblII05 = "其他"
         End Select
      End If
      'Modify By Sindy 2019/7/1
      If "" & rsTmp.Fields("Ti07") = "Y" Then '刪除
         Me.LblII10_T = "刪除人員:"
         Me.LblII08_T = "未轉寄刪除日期:"
      Else
         Me.LblII10_T = "轉寄人員:"
         Me.LblII08_T = "轉寄日期:"
      End If
      '2019/7/1 END
         If Val("" & rsTmp.Fields("Ti08")) = 0 Then
            Me.LblII08 = ""
         Else
            Me.LblII08 = ChangeWStringToTDateString(rsTmp.Fields("Ti08"))
         End If
         If "" & rsTmp.Fields("Ti09") = "" Then
            Me.LblII09 = ""
         Else
            Me.LblII09 = Format(rsTmp.Fields("Ti09"), "##:##:##")
         End If
         If "" & rsTmp.Fields("Ti10") = "" Then
            Me.LblII10 = ""
         Else
            Me.LblII10 = rsTmp.Fields("Ti10") & " " & GetPrjSalesNM(rsTmp.Fields("Ti10"))
         End If
      
      If Val("" & rsTmp.Fields("Ti12")) = 0 Then
         Me.LblII12 = ""
      Else
         Me.LblII12 = ChangeWStringToTDateString(rsTmp.Fields("Ti12"))
      End If
      If "" & rsTmp.Fields("Ti13") = "" Then
         Me.LblII13 = ""
      Else
         Me.LblII13 = Format(rsTmp.Fields("Ti13"), "##:##:##")
      End If
      If "" & rsTmp.Fields("Ti14") = "" Then
         cmdOpenF.Enabled = False
      Else
         m_II14 = rsTmp.Fields("Ti14") 'Add By Sindy 2016/10/4
         cmdOpenF.Enabled = True
      End If
      'Msg檔可刪除日期
      Me.LblII16.Visible = False
      'If Pub_StrUserSt03 = "M51" Then
         Me.LblII16.Visible = True
         Me.LblII16 = ""
         If Val("" & rsTmp.Fields("Ti16")) > 0 Then
            Me.LblII16 = "可刪除日期:" & ChangeWStringToTDateString(rsTmp.Fields("Ti16"))
         End If
      'End If
      If "" & rsTmp.Fields("Ti18") = "" Then
         Me.LblPI18 = ""
      Else
         Me.LblPI18 = rsTmp.Fields("Ti18") & "-" & rsTmp.Fields("Ti19") & "-" & rsTmp.Fields("Ti20") & "-" & rsTmp.Fields("Ti21")
      End If
   Else
      ReadTMInput = False
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim bolState As Boolean
Dim bolHaveData As Boolean
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   LblNote.Caption = "系統記錄:"
   'Modify By Sindy 2016/9/22
   If Len(m_II03) = 5 And Left(m_II03, 1) = "P" Then
      'Add By Sindy 2023/5/23
      Me.FraII27.Visible = False
      Me.Lblii28.Visible = False
      Me.txtII28.Visible = False
      '2023/5/23 END
      bolHaveData = ReadPatentInput
   'Add By Sindy 2019/4/16
   ElseIf Len(m_II03) = 5 And Left(m_II03, 1) = "T" Then
      'Add By Sindy 2023/5/23
      Me.FraII27.Visible = False
      Me.Lblii28.Visible = False
      Me.txtII28.Visible = False
      '2023/5/23 END
      bolHaveData = ReadTMInput
   Else
      'Add By Sindy 2023/5/23
      Me.FraII27.Visible = True
      Me.Lblii28.Visible = True
      Me.txtII28.Visible = True
      '2023/5/23 END
      bolHaveData = ReadIPDeptInput
   End If
   If bolHaveData = False Then
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   '2016/9/22 END
   
   '明細檔
   GRD1.Clear
   'Modify By Sindy 2016/9/22
   bolState = False
   
   LblReceiver.Caption = PUB_GetMailInputData(m_II01, m_II03)
   If LblReceiver.Caption <> "" Then
      Label5.Visible = True
      LblReceiver.Visible = True
   Else
      Label5.Visible = False
      LblReceiver.Visible = False
   End If
   '專利處信箱
   'Modify By Sindy 2019/4/16 商標處信箱
   'Modify By Sindy 2022/6/22 + 外專信件沖銷啟用日
   If Len(m_II03) = 5 And (Left(m_II03, 1) = "P" Or Left(m_II03, 1) = "T" Or strSrvDate(1) >= 外專信件沖銷啟用日) Then
      '檢查有無處理狀態
      'Modify By Sindy 2021/1/25 +  or ir20 is not null) 有輸入處理原因
      strSql = "select ir16" & _
               " From inputrecord" & _
               " where ir01='" & m_II01 & "' and ir02='" & m_II02 & "' and ir03='" & ChgSQL(m_II03) & "'" & _
               " and (ir16 is not null or ir20 is not null)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         bolState = True
      End If
      rsTmp.Close
      
      Call SetGrd2(bolState)
      'Modify By Sindy 2019/7/5 + ,decode(ir24,'Y','副','') 副
      'Modify By Sindy 2022/8/5 + ,s7.st02 二次確認人員,ir21 歸卷文號
      strSql = "select decode(s1.st02,null,ir04,s1.st02) 收受者,decode(ir24,'Y','副','') 副,sqldatet(ir05)||' '||sqltime6(ir06) 讀取日期時間,s2.st02 讀取人員" & _
               ",decode(ir16," & 信件處理狀態 & ",ir16) 處理狀態" & _
               ",sqldatet(ir17)||' '||sqltime6(ir18) 處理日期時間,s6.st02 處理人員,decode(FO02,null,ir20,FO02) 處理原因" & _
               ",sqldatet(ir08)||' '||sqltime6(ir09) 沖銷日期時間,s3.st02 沖銷人員,sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
               ",decode(ir15,'Y','" & IIf(Left(m_II03, 1) = "P", "Patent", IIf(Left(m_II03, 1) = "T", "TM", "IPDept")) & "',s4.st02) 轉寄者,s5.st02 原收受者" & _
               ",s7.st02 二次確認人員,ir21 歸卷文號" & _
               " From inputrecord,staff s1,staff s2,staff s3,staff s4,staff s5,staff s6,staff s7,form" & _
               " where ir01='" & m_II01 & "' and ir02='" & m_II02 & "' and ir03='" & ChgSQL(m_II03) & "'" & _
               " and ir04=s1.st01(+) and ir07=s2.st01(+) and ir10=s3.st01(+) and ir13=s4.st01(+) and ir14=s5.st01(+)" & _
               " and ir19=s6.st01(+) and ir22=s7.st01(+) and ir20=FO01(+)" & _
               " order by ir11 desc,ir12 desc,decode(s1.st02,null,ir04,s1.st02) asc"
   '2019/4/16 END
   Else
      Call SetGrd
      'Modify By Sindy 2019/7/5 + ,decode(ir24,'Y','副','') 副
      strSql = "select decode(s1.st02,null,ir04,s1.st02) 收受者,decode(ir24,'Y','副','') 副,sqldatet(ir05)||' '||sqltime6(ir06) 讀取日期時間,s2.st02 讀取人員" & _
               ",sqldatet(ir08)||' '||sqltime6(ir09) 沖銷日期時間,s3.st02 沖銷人員,sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
               ",decode(ir15,'Y','IPDept',s4.st02) 轉寄者,s5.st02 原收受者" & _
               " From inputrecord,staff s1,staff s2,staff s3,staff s4,staff s5" & _
               " where ir01='" & m_II01 & "' and ir02='" & m_II02 & "' and ir03='" & ChgSQL(m_II03) & "'" & _
               " and ir04=s1.st01(+) and ir07=s2.st01(+) and ir10=s3.st01(+) and ir13=s4.st01(+) and ir14=s5.st01(+)" & _
               " order by ir11 desc,ir12 desc,decode(s1.st02,null,ir04,s1.st02) asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      If bolState = True Then
         GRD1.TextMatrix(0, 8) = "確認/沖銷日期時間"
         GRD1.TextMatrix(0, 9) = "確認/沖銷人員"
      End If
   End If
   rsTmp.Close
   
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   GRD1.Visible = True
   dblPrevRow = 0
   
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdNext_Click()
   If TypeName(m_PrevForm) <> "Nothing" Then
      If UCase(TypeName(m_PrevForm)) = UCase("frm100106_9") Then
         m_PrevForm.Show
      ElseIf m_PrevForm.PubShowNextData_2 = True Then
         Exit Sub
      Else
         m_PrevForm.Show
      End If
   End If
   Unload Me
End Sub

Private Sub cmdOpenF_Click()
Dim strFileName As String
   
On Error GoTo ErrHand
   
   '讀取檔案
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2016/10/4
   'strFileName = m_II03
   strFileName = Mid(m_II14, InStrRev(m_II14, "/") + 1)
   '2016/10/4 END
   Call PUB_ChkFileTypeOpenExE(strFileName) 'Add By Sindy 2017/9/13
   If GetAttachFile(m_II01, m_II02, m_II03, strFileName, m_II19, m_AttachPath & "\" & strFileName) = True Then
      ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 讀取失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Call QueryData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DestroyToolTip '清除物件
   Set m_PrevForm = Nothing
   Set frm06010613_1 = Nothing
End Sub

'國外部信件讀取記錄
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2019/7/5 + 副
   '                        0         1     2               3           4               5           6               7         8
   arrGridHeadText = Array("收受者", "副", "讀取日期時間", "讀取人員", "沖銷日期時間", "沖銷人員", "轉寄日期時間", "轉寄者", "原收受者")
   arrGridHeadWidth = Array(600, 400, 1600, 800, 1600, 800, 1600, 600, 800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next iRow
   GRD1.Visible = True
End Sub

Private Sub SetGrd2(bolState As Boolean)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2019/7/5 + 副
   '                        0         1     2               3           4           5           6           7           8               9           10              11        12          13              14
   arrGridHeadText = Array("收受者", "副", "讀取日期時間", "讀取人員", "處理狀態", "處理日期時間", "處理人員", "處理原因", "沖銷日期時間", "沖銷人員", "轉寄日期時間", "轉寄者", "原收受者", "二次確認人員", "歸卷文號")
   If bolState = True Then
      arrGridHeadWidth = Array(600, 400, 1600, 800, 800, 1600, 800, 1200, 1600, 800, 1600, 600, 800, 800, 1000)
   Else
      arrGridHeadWidth = Array(600, 400, 1600, 800, 0, 0, 0, 0, 1600, 800, 1600, 600, 800, 800, 1000)
   End If
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next iRow
   GRD1.Visible = True
End Sub

Private Sub Grd1_Click()
GRD1.Visible = False
GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col
If nRow = 0 Then
'   If GRD1.Text <> "V" Then
'      If GRD1.Text = "無" Then
'         If m_blnColOrderAsc = True Then
'            GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
'         If m_blnColOrderAsc = True Then
'            GRD1.Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            GRD1.Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      End If
'   End If
Else
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      GRD1.col = 0
'      GRD1.row = dblPrevRow
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = QBColor(15)
'      Next i
'      Call SetColor(dblPrevRow)
'   End If
'   '目前資料列反白
'   GRD1.row = nRow
'   dblPrevRow = GRD1.row
'
'   If GRD1.TextMatrix(GRD1.row, 14) <> "" Then
'      GRD1.col = 0
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   
   GRD1.row = nRow 'GRD1.MouseRow
   'dblPrevRow = GRD1.row '記錄目前筆數
   GRD1.col = 0
   If GRD1.TextMatrix(GRD1.row, 0) <> "" Then
      '清除反白
      'If GRD1.TextMatrix(GRD1.row, 0) = "V" Then
      If dblPrevRow <> GRD1.row Then
         GRD1.col = 0
         GRD1.row = dblPrevRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      End If
      If GRD1.CellBackColor = &HFFC0C0 Then
         GRD1.col = 0
         GRD1.row = nRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         '將點選資料列反白
         'GRD1.TextMatrix(GRD1.row, 0) = "V"
         GRD1.col = 0
         GRD1.row = nRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         dblPrevRow = GRD1.row
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, ByVal strPkey3 As String, _
                               ByRef pFileName As String, ByVal strCP09 As String, _
                               Optional pSavePath As String) As Boolean
Dim stAttPath As String
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   'Modify By Sindy 2019/5/2
   GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   'Modify By Sindy 2016/9/22
'   If Len(m_II03) = 5 And Left(m_II03, 1) = "P" Then
'      GetAttachFile = PUB_GetAttachFile_PImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   'Add By Sindy 2019/4/16
'   ElseIf Len(m_II03) = 5 And Left(m_II03, 1) = "T" Then
'      GetAttachFile = PUB_GetAttachFile_TImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   Else
'      If strCP09 <> "" Then '個案
'         GetAttachFile = PUB_GetAttachFile_CPP(strCP09, pFileName, stAttPath, True)
'         'ADD BY SONIA 2016/4/8 因之前放入個案,故個案讀不到加入下面語法
'         If GetAttachFile = False Then
'            GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'         End If
'         'END 2016/4/8
'      Else
'         GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'      End If
'   End If
   
   Exit Function
   
ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 Then
      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
            'GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
            CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
         End If
         iRow = GRD1.MouseRow
         iCol = GRD1.MouseCol
      End If
   End If
End Sub
