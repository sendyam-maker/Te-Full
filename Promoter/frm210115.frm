VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210115 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案件彙整表"
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9312
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   1530
      TabIndex        =   24
      Top             =   1440
      Width           =   2500
   End
   Begin VB.CheckBox chk 
      Caption         =   "含閉卷"
      Height          =   225
      Index           =   1
      Left            =   6390
      TabIndex        =   8
      Top             =   420
      Value           =   1  '核取
      Width           =   885
   End
   Begin VB.CheckBox chk 
      Caption         =   "含核駁"
      Height          =   225
      Index           =   0
      Left            =   5490
      TabIndex        =   7
      Top             =   420
      Value           =   1  '核取
      Width           =   885
   End
   Begin VB.OptionButton opt1 
      Caption         =   "客戶名稱："
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   20
      Top             =   810
      Width           =   1215
   End
   Begin VB.OptionButton opt1 
      Caption         =   "本所案號："
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   19
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "顯示代表圖(&I)"
      Height          =   375
      Index           =   4
      Left            =   6825
      TabIndex        =   12
      Top             =   30
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3945
      Left            =   30
      TabIndex        =   16
      Top             =   1800
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   6964
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "電子檔(&E)"
      Height          =   375
      Index           =   3
      Left            =   4935
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   2
      Left            =   8340
      TabIndex        =   13
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      ToolTipText     =   "列印下內表格內容"
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3990
      TabIndex        =   9
      Top             =   30
      Width           =   885
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1530
      TabIndex        =   6
      Top             =   1110
      Width           =   1905
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "3360;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   3210
      TabIndex        =   3
      Top             =   420
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   420
      Width           =   240
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "423;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2070
      TabIndex        =   1
      Top             =   420
      Width           =   735
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1296;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   420
      Width           =   465
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "820;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   6390
      TabIndex        =   5
      Top             =   750
      Width           =   1005
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1773;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   6
      Left            =   1530
      TabIndex        =   4
      Top             =   735
      Width           =   2925
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "5159;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   7440
      TabIndex        =   26
      Top             =   750
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "預設印表機："
      Height          =   255
      Left            =   300
      TabIndex        =   25
      Top             =   1463
      Width           =   1155
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "一次最多只可勾選10組列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   135
      TabIndex        =   23
      Top             =   90
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "列印時條件："
      Height          =   180
      Left            =   4380
      TabIndex        =   22
      Top             =   450
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(字首比對)"
      Height          =   180
      Index           =   2
      Left            =   4530
      TabIndex        =   21
      Top             =   810
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(不可為唯一條件)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4320
      TabIndex        =   18
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(字首比對)"
      Height          =   180
      Index           =   0
      Left            =   3450
      TabIndex        =   17
      Top             =   1140
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1650
      X2              =   3420
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號："
      Height          =   180
      Left            =   300
      TabIndex        =   15
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   1
      Left            =   5490
      TabIndex        =   14
      Top             =   810
      Width           =   900
   End
End
Attribute VB_Name = "frm210115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/07/04 改成Form 2.0 列印：逐字檢查Unicode文字改以圖片方式列印
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1改字型=新細明體-ExtB、txt1(index)、lbl1 ; Printer列印未改
'Memo by Lydia 2019/07/01 表單名稱:客戶專利案件整理表=> 專利案件彙整表
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'create by nickc 2007/08/24
Option Explicit

Dim NowKey1 As String
Dim NowKey2 As String
Dim NowKey3 As String
Dim i As Integer, j As Integer
Dim IsChk As Boolean
Dim iPage As Integer, AllPage As Integer
Dim iLine As Integer
Dim iMaxLine As Integer
Dim PLeft(1 To 12) As Integer
Dim strTemp(1 To 12) As String
Dim ChgToTmpPic As String
Dim SeekFirstPic As Integer
'若是有勾到該群組的案件，一筆就可以，要整個群組找圖
Dim SeekStd As Integer    '群組第一筆
Dim SeekEnd As Integer   '群組最後一筆
Dim SeekShowPic As String    '準備要顯示的本所案號，若都沒有圖，則"顯示代表圖建置中"  000-000000-0-00
Dim SeekPoint As Long
Dim SeekMo As Integer
Dim CalAllPages As Integer
Public ChkPwd As String
Dim IsCan2Page As Boolean   'add by nickc 2007/10/04   控制一筆時將兩組印再一張
Dim IfPrint As Boolean    '單組
Dim IfPrintAnyOne As Boolean   '任一筆
Dim SeekPrintGroup As String
Dim CalPrintCount As Integer
Dim tmpArr As Variant
Dim m_iSelCount As Integer 'Add by Morgan 2011/1/14 勾選組數
'Add by Amy 2014/05/22
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim stOrgPrinter As String 'Add by Amy 2018/07/19
Dim m_Device 'Add by Amy 2018/09/14
Dim Xo As Long, Yo As Long 'Added by Lydia 2022/07/04
Public bolRunPWD As Boolean  'Added by Lydia 2024/01/12

Public Sub cmdOK_Click(Index As Integer)
Dim k As Long
Select Case Index
Case 0
         Screen.MousePointer = vbHourglass
         grd1.MousePointer = flexArrowHourGlass
         If (opt1(0).Value = True And Trim(Txt1(0)) = "" And Trim(Txt1(1)) = "") Or (opt1(1).Value = True And Trim(Txt1(6)) = "" And Trim(Txt1(4)) = "" And Trim(Txt1(5)) = "") Then
             MsgBox "最少輸入一種條件！", vbExclamation, "輸入錯誤！"
             Txt1(0).SetFocus
             grd1.MousePointer = flexDefault
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
         If opt1(1).Value = True And Trim(Txt1(4)) = "" And Trim(Txt1(5)) <> "" And Trim(Txt1(6)) = "" Then
             MsgBox "使用客戶案件案號查詢時，請補充  客戶名稱  或 客戶編號！", vbExclamation, "輸入錯誤！"
             Txt1(6).SetFocus
             grd1.MousePointer = flexDefault
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
         'Added by Lydia 2022/07/07 本所案號條件要完整，不然insert into r020115的資料太大會發生"超出目前的範圍"的錯誤；
         If opt1(0).Value = True Then
             If Trim(Txt1(0)) = "" Then
                 MsgBox "請輸入系統別！", vbCritical
                 Txt1(0).SetFocus
                 txt1_GotFocus 0
                 grd1.MousePointer = flexDefault
                 Screen.MousePointer = vbDefault
                 Exit Sub
             End If
             If Trim(Txt1(1)) = "" Then
                 MsgBox "請輸入本所案號！", vbCritical
                 Txt1(1).SetFocus
                 txt1_GotFocus 1
                 grd1.MousePointer = flexDefault
                 Screen.MousePointer = vbDefault
                 Exit Sub
             End If
         End If
         'end 2022/07/07
         If opt1(1).Value = True And Trim(Txt1(4)) = "" Then
            Load frm210115_1
            frm210115_1.Hide
            frm210115_1.StrMenu
            If opt1(1).Value = True And Trim(Txt1(4)) = "" Then
                frm210115_1.Show vbModal
            Else
                Unload frm210115_1
            End If
         End If
         txt1_Validate 4, False
         'Debug.Print Timer
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
         StrMenu
         'Debug.Print Timer
         grd1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
Case 1
         '測試  將所有併為一組
         'For j = frm210115.Grd1.Rows - 2 To 1 Step -1
         '   If frm210115.Grd1.TextMatrix(j, 1) = "" Then frm210115.Grd1.RemoveItem j
         'Next j
         'add by nickc 2007/10/04
         IsCan2Page = False
         IfPrint = False
         IfPrintAnyOne = False
         '檢查是否可以列印
         
         If grd1.Rows <= 2 Then
            grd1.row = grd1.Rows - 1
            If grd1.row = 0 Then
                MsgBox "沒有任何資料可以列印！", vbExclamation, "發生錯誤！"
                Exit Sub
            Else
                grd1.col = 1
                If Trim(grd1.Text) = "" Then
                    MsgBox "沒有任何資料可以列印！", vbExclamation, "發生錯誤！"
                    Exit Sub
                End If
            End If
         End If
         Screen.MousePointer = vbHourglass
         'Add by Amy 2018/09/14 印表機
         PUB_RestorePrinter cboPrinter
         Set m_Device = Printer
         m_Device.PaperSize = 9
         m_Device.EndDoc
         m_Device.Orientation = 1
         'end 2018/09/13
         grd1.MousePointer = flexArrowHourGlass
         AllPage = 0    '所有的頁數
        '若是都沒有勾選，則代第一個有圖的案子
        IsChk = False
        SeekFirstPic = 0
        CalAllPages = 0
        CalPrintCount = 0
        SeekPrintGroup = ""
        For i = 1 To grd1.Rows - 1
            grd1.row = i
            grd1.col = 0
            If grd1.TextMatrix(i, 13) <> "" And SeekFirstPic = 0 Then SeekFirstPic = i
            If grd1.TextMatrix(i, 1) = "" Then CalAllPages = CalAllPages + 1
            If grd1.Text = "V" Then
                IsChk = True
                'Exit For
                If grd1.TextMatrix(i, 14) = "" Then
                        If SeekPrintGroup <> "" Then
                            SeekPrintGroup = SeekPrintGroup & ","
                        End If
                        SeekPrintGroup = SeekPrintGroup & grd1.TextMatrix(i, 1)
                Else
                    If InStr(1, SeekPrintGroup, grd1.TextMatrix(i, 14)) = 0 Then
                        If SeekPrintGroup <> "" Then
                            SeekPrintGroup = SeekPrintGroup & ","
                        End If
                        SeekPrintGroup = SeekPrintGroup & grd1.TextMatrix(i, 14)
                    End If
                End If
            End If
        Next
        tmpArr = Split(SeekPrintGroup, ",")
        CalPrintCount = UBound(tmpArr) + 1
     If SeekFirstPic = 0 Then SeekFirstPic = 1
     If IsChk = False And Pub_StrUserSt03 <> "M51" Then
         MsgBox "沒有選擇要列印的資料，請選擇 10 組以內，請重新選擇！", vbExclamation, "發生錯誤！"
         grd1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
         Exit Sub
     ElseIf IsChk = False And (Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0) Then 'edit by nickc 2007/10/16 中所修改說全部只能電腦中心印
                grd1.MousePointer = flexDefault
                Screen.MousePointer = vbDefault
            If ChkCanPrint Then
                Screen.MousePointer = vbHourglass
                grd1.MousePointer = flexArrowHourGlass
                    If CalAllPages <> 1 Then
                        If MsgBox("共約 " & CalAllPages & " 個家族，至少 " & CalAllPages & " 張，確定列印？", vbYesNo, "列印！") = vbNo Then
                            grd1.MousePointer = flexDefault
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    SeekStd = 1
                    SeekEnd = 0
                    '算最後一筆
                    For j = 1 To grd1.Rows - 1
                        grd1.col = 1
                        grd1.row = j
                        'SeekShowPic = Grd1.TextMatrix(SeekStd, 1)
                        SeekShowPic = "000-000000-0-00"
                        If grd1.Text = "" Then
                            SeekEnd = j - 1
                            For i = SeekStd To SeekEnd
                                If grd1.TextMatrix(i, 13) <> "" Then
                                    SeekShowPic = grd1.TextMatrix(i, 1)
                                    Exit For
                                End If
                            Next i
                            PrintData SeekStd, SeekEnd
                            If SeekEnd <> grd1.Rows - 2 Then
                                If IsCan2Page = False And IfPrint = True Then
                                    Printer.NewPage
                                End If
                            End If
                            SeekStd = SeekEnd + 2
                        End If
                    Next j
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
     'add by nickc 2007/10/16 中所新增的
     'Remove by Morgn 2011/1/14 改點選時控制
     'ElseIf IsChk = True And CalPrintCount > 5 And Pub_StrUserSt03 <> "M51" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
     '    MsgBox "超過5組，請重新選擇！", vbExclamation, "數量太多！"
     '    grd1.MousePointer = flexDefault
     '    Screen.MousePointer = vbDefault
     '    Exit Sub
     
     'Modified by Morgan 2011/12/7 不再限制 5 組
     'ElseIf IsChk = True And (CalPrintCount <= 5 Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0) Then
     ElseIf IsChk = True Then
                grd1.MousePointer = flexDefault
                Screen.MousePointer = vbDefault
            If ChkCanPrint Then
                Screen.MousePointer = vbHourglass
                grd1.MousePointer = flexArrowHourGlass
                    SeekStd = 0
                    SeekEnd = 0
                    SeekShowPic = "000-000000-0-00"
                    Me.Enabled = False
                    IfPrint = False
                    For i = 1 To grd1.Rows - 1
                        grd1.col = 0
                        grd1.row = i
                        If Trim(grd1.Text) = "V" Then
                             'SeekShowPic = grd1.TextMatrix(i, 1)
                             SeekShowPic = "000-000000-0-00"
                            '算第一筆
                            For j = i - 1 To 1 Step -1
                                grd1.col = 1
                                grd1.row = j
                                If grd1.Text = "" Then
                                    SeekStd = j + 1
                                    Exit For
                                End If
                            Next j
                            If SeekStd = 0 Then SeekStd = i
                            '算最後一筆
                            For j = i + 1 To grd1.Rows - 1
                                grd1.col = 1
                                grd1.row = j
                                If grd1.Text = "" Then
                                    SeekEnd = j - 1
                                    Exit For
                                End If
                            Next j
                            If SeekStd <> 0 And SeekEnd <> 0 Then
                                For j = SeekStd To SeekEnd
                                    If grd1.TextMatrix(j, 13) <> "" Then
                                        SeekShowPic = grd1.TextMatrix(j, 1)
                                        Exit For
                                    End If
                                Next j
                            End If
                            '拿掉V
                            grd1.Visible = False
                            For k = SeekStd To SeekEnd
                                grd1.row = k
                                grd1.col = 0
                                grd1.Text = ""
                                m_iSelCount = m_iSelCount - 1 'Add by Morgan 2011/3/9
                                For j = 4 To grd1.Cols - 1
                                    grd1.col = j
                                    grd1.CellBackColor = QBColor(15)
                                Next j
                            Next k
                            grd1.Visible = True
                            If IfPrint = True Then
                                If IsCan2Page = False Then
                                    Printer.NewPage
                                End If
                            End If
                            PrintData SeekStd, SeekEnd
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                        End If
                    Next i
            End If
        End If
        If IfPrintAnyOne = True Then
            AllPage = Printer.Page
            Printer.EndDoc
            '紀錄列印
    '        If Pub_StrUserSt03 <> "M51" Then
                Dim tmpMemo As String
                tmpMemo = ""
                If NowKey1 <> "" Then
                    tmpMemo = tmpMemo & "本所案號：" & NowKey1
                End If
                If NowKey2 <> "" Then
                    If tmpMemo <> "" Then tmpMemo = tmpMemo & vbCrLf
                    tmpMemo = tmpMemo & "客戶編號：" & NowKey2
                End If
                If NowKey3 <> "" Then
                    If tmpMemo <> "" Then tmpMemo = tmpMemo & vbCrLf
                    tmpMemo = tmpMemo & "客戶案件案號：" & NowKey3
                End If
                If chk(0).Value = vbUnchecked Then
                    If tmpMemo <> "" Then tmpMemo = tmpMemo & vbCrLf
                    tmpMemo = tmpMemo & "不含核駁"
                End If
                If chk(1).Value = vbUnchecked Then
                    If tmpMemo <> "" Then tmpMemo = tmpMemo & vbCrLf
                    tmpMemo = tmpMemo & "不含閉卷"
                End If
                If IsChk = False Then tmpMemo = tmpMemo & "全部列印"
                If tmpMemo <> "" Then tmpMemo = tmpMemo & vbCrLf & "共列印 " & AllPage & " 頁，列印到 " & Printer.DeviceName & " 印表機 "
                Pub_SaveLog IIf(Me.Tag = "", strUserNum, Me.Tag), tmpMemo, , , , , , Me.Caption & "(" & Me.Name & ")", "3"
    '        End If
            MsgBox "共列印 " & AllPage & " 頁", , "列印完成！"
         Else
            If bolRunPWD = True Then 'Added by Lydia 2024/01/12 避免在輸入密碼同時彈訊息
               MsgBox "沒有任何資料可以列印！", vbExclamation, "發生錯誤！"
               Me.Show 'Added by Lydia 2024/01/12
            End If
         End If
         grd1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
Case 2
        Unload Me
Case 3
        Dim MyPrt As Integer
        Dim SMyPrtN As String
        Dim SMyPrtI  As Integer
        Dim SMyPrtO As Integer
        Dim SMyPDF As Integer
        SMyPrtN = Printer.DeviceName
        SMyPrtO = Printer.Orientation
        For MyPrt = 0 To Printers.Count - 1
            Set Printer = Printers(MyPrt)
            If InStr(1, UCase(Printer.DeviceName), "PDF") <> 0 Or InStr(1, UCase(Printer.DeviceName), "ADOBE") <> 0 Or InStr(1, UCase(Printer.DeviceName), "ACROBAT") <> 0 Then
                SMyPDF = MyPrt
            End If
            If Printer.DeviceName = SMyPrtN Then
                SMyPrtI = MyPrt
            End If
        Next MyPrt
        Set Printer = Printers(SMyPDF)
        cmdOK_Click 1
        Set Printer = Printers(SMyPrtI)
Case 4
        '若是都沒有勾選，則代第一個有圖的案子
        IsChk = False
        SeekFirstPic = 0
        For i = 1 To grd1.Rows - 1
            grd1.row = i
            grd1.col = 0
            If grd1.TextMatrix(i, 13) <> "" And SeekFirstPic = 0 Then SeekFirstPic = i
            If grd1.Text = "V" Then
                IsChk = True
                Exit For
            End If
        Next
     If SeekFirstPic = 0 Then SeekFirstPic = 1
     If IsChk = False Then
            SeekStd = 1
            SeekEnd = 0
            SeekShowPic = grd1.TextMatrix(1, 1)
            '算最後一筆
            For j = 1 To grd1.Rows - 1
                grd1.col = 1
                grd1.row = j
                If grd1.Text = "" Then
                    SeekEnd = j - 1
                    Exit For
                End If
            Next j
            If SeekStd <> 0 And SeekEnd <> 0 Then
                For j = SeekStd To SeekEnd
                    If grd1.TextMatrix(j, 13) <> "" Then
                        SeekShowPic = grd1.TextMatrix(j, 1)
                        Exit For
                    End If
                Next j
            End If
            Me.Enabled = False
            grd1.row = SeekFirstPic
            grd1.col = 1
            Me.Hide
            Screen.MousePointer = vbHourglass
            frmPic001.oCP01 = SystemNumber(Pub_RplStr(SeekShowPic), 1)
            frmPic001.oCP02 = SystemNumber(Pub_RplStr(SeekShowPic), 2)
            frmPic001.oCP03 = SystemNumber(Pub_RplStr(SeekShowPic), 3)
            frmPic001.oCP04 = SystemNumber(Pub_RplStr(SeekShowPic), 4)
            frmPic001.StrMenu
            frmPic001.cmdok(0).Visible = False
            frmPic001.cmdok(1).Visible = False
            frmPic001.cmdok(2).Visible = False
            frmPic001.cmdok(4).Visible = False
            frmPic001.cmdok(5).Visible = False
            frmPic001.cmdok(6).Visible = False
            frmPic001.Label12.Visible = False
            frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
            frmPic001.Show vbModal
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Me.Show
     Else
            SeekStd = 0
            SeekEnd = 0
            SeekShowPic = "000-000000-0-00"
            Me.Enabled = False
            For i = 1 To grd1.Rows - 1
                grd1.col = 0
                grd1.row = i
                If Trim(grd1.Text) = "V" Then
                     SeekShowPic = grd1.TextMatrix(i, 1)
                    '算第一筆
                    For j = i - 1 To 1 Step -1
                        grd1.col = 1
                        grd1.row = j
                        If grd1.Text = "" Then
                            SeekStd = j + 1
                            Exit For
                        End If
                    Next j
                    If SeekStd = 0 Then SeekStd = i
                    '算最後一筆
                    For j = i + 1 To grd1.Rows - 1
                        grd1.col = 1
                        grd1.row = j
                        If grd1.Text = "" Then
                            SeekEnd = j - 1
                            Exit For
                        End If
                    Next j
                    If SeekStd <> 0 And SeekEnd <> 0 Then
                        For j = SeekStd To SeekEnd
                            If grd1.TextMatrix(j, 13) <> "" Then
                                SeekShowPic = grd1.TextMatrix(j, 1)
                                Exit For
                            End If
                        Next j
                    End If
                    '拿掉V
                    grd1.Visible = False
                    For k = SeekStd To SeekEnd
                        grd1.row = k
                        grd1.col = 0
                        grd1.Text = ""
                        m_iSelCount = m_iSelCount - 1 'Add by Morgan 2011/3/9
                        For j = 4 To grd1.Cols - 1
                            grd1.col = j
                            grd1.CellBackColor = QBColor(15)
                        Next j
                    Next k
                    grd1.Visible = True
                    Me.Hide
                    Screen.MousePointer = vbHourglass
                    frmPic001.oCP01 = SystemNumber(Pub_RplStr(SeekShowPic), 1)
                    frmPic001.oCP02 = SystemNumber(Pub_RplStr(SeekShowPic), 2)
                    frmPic001.oCP03 = SystemNumber(Pub_RplStr(SeekShowPic), 3)
                    frmPic001.oCP04 = SystemNumber(Pub_RplStr(SeekShowPic), 4)
                    frmPic001.StrMenu
                    frmPic001.cmdok(0).Visible = False
                    frmPic001.cmdok(1).Visible = False
                    frmPic001.cmdok(2).Visible = False
                    frmPic001.cmdok(4).Visible = False
                    frmPic001.cmdok(5).Visible = False
                    frmPic001.cmdok(6).Visible = False
                    frmPic001.Label12.Visible = False
                    frmPic001.Show vbModal
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                    Me.Show
                End If
            Next i
        End If
Case Else
End Select
If m_iSelCount < 0 Then m_iSelCount = 0
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Add by Amy 2018/07/19 +選印表機
stOrgPrinter = Printer.DeviceName
PUB_SetPrinter Me.Name, cboPrinter, cboPrinter.Tag, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數
'end 2018/07/19
'紀錄現在條件
NowKey1 = ""
NowKey2 = ""
NowKey3 = ""
If UCase(GetStaffDepartment(strUserNum)) = "M51" Then
    cmdok(3).Visible = True
    cmdok(3).Enabled = True
Else
    cmdok(3).Visible = False
    cmdok(3).Enabled = False
End If
'Add by Amy 2015/02/04 +總經理業務工作代理人員
If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
    bolSpecMan = True
    strSpecCode = "總經理業務工作代理人員"
'Modify  by Amy 2014/05/22 開放專利處部份智權同仁資料給彥葶代為處理
ElseIf CheckLevel(strUserNum, "A8") = True Then
    bolSpecMan = True
    strSpecCode = "A8"
End If
'end 2014/05/22
SeekMo = 0
SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add by Amy 2018/07/19 +還原印表機
PUB_RestorePrinter stOrgPrinter
If Me.cboPrinter.Enabled = True And Me.cboPrinter.Text <> Me.cboPrinter.Tag Then
    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
End If
'end 2018/07/19
Set frm210115 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim i As Integer
Dim SeekClkRow As Long
Dim iAdd As Integer 'Add by Morgan 2011/1/14

SeekStd = 0
SeekEnd = 0
grd1.Visible = False
SeekClkRow = grd1.MouseRow
grd1.row = SeekClkRow
grd1.col = 1
If grd1.Text <> "" Then
    '算第一筆
    For j = SeekClkRow - 1 To 1 Step -1
        grd1.col = 1
        grd1.row = j
        If grd1.Text = "" Then
            SeekStd = j + 1
            Exit For
        End If
    Next j
    If SeekStd = 0 Then SeekStd = 1
    '算最後一筆
    For j = SeekClkRow + 1 To grd1.Rows - 1
        grd1.col = 1
        grd1.row = j
        If grd1.Text = "" Then
            SeekEnd = j - 1
            Exit For
        End If
    Next j
    If SeekStd <> 0 And SeekEnd <> 0 Then
        grd1.row = SeekStd
        grd1.col = 0
        If grd1.Text = "V" Then
            iAdd = -1 'Add by Morgan 2011/1/14
            For j = SeekStd To SeekEnd
                 grd1.row = j
                 grd1.col = 0
                 grd1.Text = ""
                 For i = 4 To grd1.Cols - 1
                      grd1.col = i
                      grd1.CellBackColor = QBColor(15)
                Next i
            Next j
        Else
            'Add by Morgan 2011/1/14
            If Pub_StrUserSt03 <> "M51" And m_iSelCount + 1 > 10 Then
               grd1.Visible = True
               MsgBox lblMemo, vbExclamation
               Exit Sub
            Else
               iAdd = 1
            End If
            'end 2011/1/14
            
            For j = SeekStd To SeekEnd
                grd1.row = j
                grd1.col = 0
                grd1.Text = "V"
                
                For i = 4 To grd1.Cols - 1
                    grd1.col = i
                    grd1.CellBackColor = &HFFC0C0
                Next i
            Next j
        End If
    End If
End If

m_iSelCount = m_iSelCount + iAdd 'Add by Morgan 2011/1/14

grd1.Visible = True
End Sub

Private Sub Opt1_Click(Index As Integer)
Select Case Index
Case 0
        If Txt1(0).Enabled = True Then
            If SeekMo <> 1 Then
                Txt1(0).SetFocus
            End If
        End If
        CloseIme
Case 1
        If Txt1(6).Enabled = True Then
            If SeekMo <> 2 Then
                Txt1(6).SetFocus
            End If
        End If
        OpenIme
Case Else
End Select
End Sub

Private Sub txt1_Change(Index As Integer)
If Index = 6 Then
    If Trim(Txt1(6)) <> "" Then
        Txt1(4) = ""
        lbl1.Caption = ""
    End If
End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
If Index <> 6 And Index <> 4 And Index <> 5 Then
    SeekMo = 1
    CloseIme
    opt1(0).Value = True
Else
    SeekMo = 2
    OpenIme
    opt1(1).Value = True
End If
End Sub

'Modified by Lydia 2022/02/07 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If Index = 0 Or Index = 4 Then
    KeyAscii = UpperCase(KeyAscii)
End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
If Index = 4 Then
    lbl1.Caption = GetCustomerName(Txt1(Index), 0)
End If
End Sub

Sub StrMenu()
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexArrowHourGlass
grd1.Clear
grd1.Rows = 2
SetGrd
DoEvents
Dim i As Integer
'add by nickc 2005/05/27 將所有關係都抓出來
Dim TmpRecCount As Long  '回傳比數
Dim TmpRecCount1 As Long  '回傳比數
Dim TmpRecCount2 As Long  '回傳比數
Dim TmpRecCount3 As Long  '回傳比數
Dim TmpRecCount4 As Long  '回傳比數
Dim TmpRecCount5 As Long  '回傳比數
Dim TmpRecCount6 As Long  '回傳比數
Dim TmpRecCount7 As Long  '回傳比數
Dim TmpRecCount8 As Long  '回傳比數
Dim TmpRecCount9 As Long  '回傳比數
Dim TmpRecCount10 As Long  '回傳比數
Dim GroupCount As Integer

strSql = ""
If opt1(0).Value = True Then
    pub_QL05 = pub_QL05 & ";" & opt1(0).Caption 'Add By Sindy 2010/12/23
    If Trim(Txt1(0)) <> "" Then
        strSql = strSql & " and pa01='" & Txt1(0) & "' "
        pub_QL05 = pub_QL05 & Txt1(0) 'Add By Sindy 2010/12/23
    End If
    If Trim(Txt1(1)) <> "" Then
        strSql = strSql & " and pa02='" & Txt1(1) & "' "
        pub_QL05 = pub_QL05 & "-" & Txt1(1) 'Add By Sindy 2010/12/23
    End If
    If Trim(Txt1(0)) & Trim(Txt1(1)) <> "" Then
        If Trim(Txt1(2)) <> "" Then
            strSql = strSql & " and pa03='" & Txt1(2) & "' "
            pub_QL05 = pub_QL05 & "-" & Txt1(2) 'Add By Sindy 2010/12/23
        Else
            strSql = strSql & " and pa03='0' "
        End If
        If Trim(Txt1(3)) <> "" Then
            strSql = strSql & " and pa04='" & Txt1(3) & "' "
            pub_QL05 = pub_QL05 & "-" & Txt1(3) 'Add By Sindy 2010/12/23
        Else
            strSql = strSql & " and pa04='00' "
        End If
    End If
End If
If opt1(1).Value = True Then
    pub_QL05 = pub_QL05 & ";" & opt1(1).Caption & Txt1(6) & Label1(2)  'Add By Sindy 2010/12/23
    If Trim(Txt1(4)) <> "" Then
        strSql = strSql & " and (pa26='" & ChangeCustomerL(Txt1(4)) & "' or pa27='" & ChangeCustomerL(Txt1(4)) & "' or pa28='" & ChangeCustomerL(Txt1(4)) & "' or pa29='" & ChangeCustomerL(Txt1(4)) & "' or pa30='" & ChangeCustomerL(Txt1(4)) & "' ) "
        pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(4) 'Add By Sindy 2010/12/23
    End If
    If Trim(Txt1(5)) <> "" Then
        strSql = strSql + " and upper(PA48) LIKE '" + UCase(Txt1(5)) + "%'  "
        pub_QL05 = pub_QL05 & ";" & Label3 & Txt1(5) & Label1(0) 'Add By Sindy 2010/12/23
    End If
End If
If chk(0).Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & chk(0).Caption 'Add By Sindy 2010/12/23
End If
If chk(1).Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & chk(1).Caption 'Add By Sindy 2010/12/23
End If
If strSql = "" Then
    Exit Sub
End If
Dim tmpCount As Integer '迴圈次
'Debug.Print Timer
cnnConnection.Execute "delete from r020115 where id='" & strUserNum & "' ", intI
'加入基礎資料
cnnConnection.Execute "insert into r020115 select pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where 1=1 " & strSql, intI
'Debug.Print Timer
'cnnConnection.Execute "insert into r020115 select X.pa01,X.pa02,X.pa03,X.pa04,rownum * 10,'1','" & strUserNum & "' from (select pa01,pa02,pa03,pa04,decode(pa01,'P',1,2) from patent where 1=1 " & strSQL & " group by decode(pa01,'P',1,2),pa01,pa02,pa03,pa04) X "

'cnnConnection.Execute "update r020115 set r001005=0 where id='" & strUserNum & "' "
'Dim IsDataOK As Boolean
'
'Dim ChkDataRs As New ADODB.Recordset
'IsDataOK = False
'GroupCount = 10
'Do While IsDataOK = False
'      strSQL = "select * from r020115 where id='" & strUserNum & "' and r001005=0  "
'      CheckOC
'      With adoRecordset
'          .CursorLocation = adUseClient
'          .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'          If .RecordCount <> 0 Then
'            .MoveFirst
'               cnnConnection.Execute "delete from r0201152 where id='" & strUserNum & "' "
'               cnnConnection.Execute "insert into r0201152 select '" & CheckStr(.Fields("R001001")) & "','" & CheckStr(.Fields("R001002")) & "','" & CheckStr(.Fields("R001003")) & "','" & CheckStr(.Fields("R001004")) & "',0,'1','" & strUserNum & "' from dual "
'               TmpRecCount = 1
'               tmpCount = 1
'                     cnnConnection.Execute "insert into r0201152 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount1
'                     cnnConnection.Execute "insert into r0201152 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount2
'                     cnnConnection.Execute "insert into r0201152 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount3
'                     cnnConnection.Execute "insert into r0201152 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount4
'                     cnnConnection.Execute "insert into r0201152 select cr01,cr02,cr03,cr04," & tmpCount & ",'1','" & strUserNum & "' from caserelation where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount5
'                     cnnConnection.Execute "insert into r0201152 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount6
'                     cnnConnection.Execute "insert into r0201152 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r0201152 where  id='" & strUserNum & "') ", TmpRecCount7
'                     'add by nickc 2005/10/26 刪除沒有相關案的
'                     If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7) = 0 Then
'                         cnnConnection.Execute "delete r020115 where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "' and r001006='1') and id='" & strUserNum & "' and r001006='1' and R001005='0' "
'                     Else
'                        cnnConnection.Execute "update r020115 set r001005=(select nvl(min(r001005),0) from r020115 " & _
'                                                            " where r001001||r001002||r001003||r001004 in (select cm01||cm02||cm03||cm04 from casemap where cm05||cm06||cm07||cm08 in ( " & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'                                                            " union select cm05||cm06||cm07||cm08 from casemap where cm01||cm02||cm03||cm04 in ( " & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'                                                            " union select cr01||cr02||cr03||cr04 from caserelation where cr05||cr06||cr07||cr08 in (" & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "')) and r001005<>0 and id='" & strUserNum & "') " & _
'                                                            " where r001001||r001002||r001003||r001004 in (select cm01||cm02||cm03||cm04 from casemap where cm05||cm06||cm07||cm08 in ( " & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'                                                            " union select cm05||cm06||cm07||cm08 from casemap where cm01||cm02||cm03||cm04 in (" & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'                                                            " union select cr01||cr02||cr03||cr04 from caserelation where cr05||cr06||cr07||cr08 in (" & _
'                                                            " select r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "') ) and id='" & strUserNum & "' "
'                        TmpRecCount7 = 0
'                        cnnConnection.Execute "update r020115 set r001005=" & GroupCount & " where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r0201152 where id='" & strUserNum & "' ) and id='" & strUserNum & "' and r001005='0' ", TmpRecCount7
'                        If TmpRecCount7 <> 0 Then
'                           cnnConnection.Execute "insert into r020115 (r001005,id) values (" & GroupCount + 1 & ",'" & strUserNum & "') "
'                        End If
'                        GroupCount = GroupCount + 10
'                     End If
'                     tmpCount = tmpCount + 1
'                     cnnConnection.Execute "delete r020115 where id='" & strUserNum & "' and r001005=0  and r001001='" & CheckStr(.Fields("r001001").Value) & "' and r001002='" & CheckStr(.Fields("R001002")) & "' and r001003='" & CheckStr(.Fields("R001003")) & "' and r001004='" & CheckStr(.Fields("R001004")) & "'  "
'          Else
'               IsDataOK = True
'          End If
'      End With
'Loop
 'Debug.Print "proc S" & Timer
 'Memo by Lydia 2019/04/16 預儲程序(db_r020115) 抓子案資料
 cnnConnection.Execute "begin   db_r020115('" & strUserNum & "'); end;"
 'Debug.Print "proc E" & Timer

'Modified by Lydia 2019/04/16 本所案號後面加閉卷符號＊ =>DECODE(PA57,'Y','＊','')
'strSql = "select distinct '',replace(r001001||'-'||r001002||'-'||r001003||'-'||r001004,'---','')||DECODE(PA57,'Y','＊',''),substr(na03,1,3)||substr(decode(pa09,'020',ptm04,'013',ptm04,ptm03),1,3),nvl(pa05,nvl(pa06,pa07)),pa48," & SQLDate("pa10") & ",pa11,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & ",NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90))||decode(nvl(pa27,'')||nvl(pa28,'')||nvl(pa29,'')||nvl(pa30,''),'','','(多申請人)'),decode(pa16,null,decode(pa57,null,'',decode(pa16,'1','准','2','駁','')||'/'||decode(pa57,'Y','閉卷','')),decode(pa16,'1','准','2','駁','')||'/'||decode(pa57,'Y','閉卷','')) ,GetNextYear(pa01,pa02,pa03,pa04),pa47,decode(ibf13,null,'','Y'),r001005 "
'strSql = strSql & " from patent,nation,r020115,customer,patenttrademarkmap,caseprogress A,imgbytefile where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "' and pa09=na01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26||'000',9,1)=cu02(+) and r001001=A.cp01(+) and r001002=A.cp02(+) and r001003=A.cp03(+) and r001004=A.cp04(+) and '1'=ptm01(+) and pa08=ptm02(+) AND (A.CP09=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04) or A.cp09 is null) "
'strSql = strSql & " and r001001=ibf01(+) and r001002=ibf02(+) and r001003=ibf03(+) and r001004=ibf04(+) and '1'=ibf05(+) "
'strSql = strSql & " order by r001005, decode(substr(replace(r001001||'-'||r001002||'-'||r001003||'-'||r001004,'---',''),1,1),'P',1,'CFP',2,3),replace(r001001||'-'||r001002||'-'||r001003||'-'||r001004,'---','') "
strSql = "select distinct '' as v,replace(r001001||'-'||r001002||'-'||r001003||'-'||r001004,'---','')||DECODE(PA57,'Y','＊','') as caseno," & _
            "substr(na03,1,3)||substr(decode(pa09,'020',ptm04,'013',ptm04,ptm03),1,3) casetype,nvl(pa05,nvl(pa06,pa07)) as casename,pa48" & _
            ",DECODE(pa10,'','','0','',SUBSTR(pa10,1,4)-1911||'/'||SUBSTR(pa10,5,2)||'/'||SUBSTR(pa10,7,2)) pa10,pa11,pa22," & _
            "DECODE(pa24,'','','0','',SUBSTR(pa24,1,4)||'/'||SUBSTR(pa24,5,2)||'/'||SUBSTR(pa24,7,2))||decode(pa24,null,decode(pa25,null,'','-'),'-')||DECODE(pa25,'','','0','',SUBSTR(pa25,1,4)||'/'||SUBSTR(pa25,5,2)||'/'||SUBSTR(pa25,7,2)) pa24pa25," & _
            "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90))||decode(nvl(pa27,'')||nvl(pa28,'')||nvl(pa29,'')||nvl(pa30,''),'','','(多申請人)') as custname," & _
            "decode(pa16,null,decode(pa57,null,'',decode(pa16,'1','准','2','駁','')||'/'||decode(pa57,'Y','閉卷','')),decode(pa16,'1','准','2','駁','')||'/'||decode(pa57,'Y','閉卷','')) as ctype," & _
            "GetNextYear(pa01,pa02,pa03,pa04) as npyear ,pa47 ,decode(ibf13,null,'','Y') as pic01,r001005"
strSql = strSql & " from patent,nation,r020115,customer,patenttrademarkmap,caseprogress A,imgbytefile where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "' and pa09=na01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26||'000',9,1)=cu02(+) and r001001=A.cp01(+) and r001002=A.cp02(+) and r001003=A.cp03(+) and r001004=A.cp04(+) and '1'=ptm01(+) and pa08=ptm02(+) AND (A.CP09=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04) or A.cp09 is null) "
strSql = strSql & " and r001001=ibf01(+) and r001002=ibf02(+) and r001003=ibf03(+) and r001004=ibf04(+) and '1'=ibf05(+) "
strSql = "select * from (" & strSql & ") order by r001005, decode(substr(caseno,1,1),'P',1,'CFP',2,3),caseno"
'end 2019/04/16
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
     'Debug.Print Timer
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
     'Debug.Print Timer
    If .RecordCount <> 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/23
        grd1.FixedCols = 0
        Set grd1.Recordset = adoRecordset
        m_iSelCount = 0 'Add by Morgan 2011/1/14
        
        grd1.Visible = False
'        For i = 1 To grd1.Rows - 1
'            grd1.Row = i
'            grd1.col = 1
'            If Trim(grd1.Text) <> "" Then
'                grd1.TextMatrix(i, 11) = GetNextYear(SystemNumber(grd1.Text, 1), SystemNumber(grd1.Text, 2), SystemNumber(grd1.Text, 3), SystemNumber(grd1.Text, 4))
'            End If
'        Next i
        '紀錄條件
        NowKey1 = "": NowKey2 = "": NowKey3 = ""
        If opt1(0).Value = True Then
            NowKey1 = Txt1(0) & "-" & Txt1(1) & "-" & IIf(Trim(Txt1(2)) = "", "0", Txt1(2)) & "-" & IIf(Trim(Txt1(3)) = "", "00", Txt1(3))
        ElseIf opt1(1).Value = True Then
            NowKey2 = Txt1(4)
            NowKey3 = Txt1(5)
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/23
        grd1.Visible = False
        grd1.Clear
        grd1.Rows = 2
        grd1.Visible = True
        ShowNoData
    End If
End With
SetGrd
grd1.Visible = True
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub SetGrd()
grd1.Cols = 16
grd1.row = 0
grd1.col = 0: grd1.Text = "V"
grd1.ColWidth(0) = 200
grd1.col = 1: grd1.Text = "本所案號"
'Modified by Lydia 2019/04/16
'GRD1.ColWidth(1) = 1400
grd1.ColWidth(1) = 1520
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 2: grd1.Text = "國家種類"
grd1.ColWidth(2) = 1200
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 3: grd1.Text = "專利名稱"
grd1.ColWidth(3) = 1500
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 4: grd1.Text = "客戶案號"
grd1.ColWidth(4) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 5: grd1.Text = "申請日"
grd1.ColWidth(5) = 800
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 6: grd1.Text = "申請號"
grd1.ColWidth(6) = 1250
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 7: grd1.Text = "證書號"
grd1.ColWidth(7) = 1250
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 8: grd1.Text = "專利期限"
grd1.ColWidth(8) = 1800
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 9: grd1.Text = "申請人"
grd1.ColWidth(9) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 10: grd1.Text = "准駁及閉卷"
grd1.ColWidth(10) = 1250
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 11: grd1.Text = "下次繳費日期"
grd1.ColWidth(11) = 1200
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 12: grd1.Text = "分所號"
grd1.ColWidth(12) = 1200
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 13: grd1.Text = ""
grd1.ColWidth(13) = 0
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 14: grd1.Text = ""
grd1.ColWidth(14) = 0
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 15: grd1.Text = ""
grd1.ColWidth(15) = 0
grd1.CellAlignment = flexAlignCenterCenter
grd1.FixedCols = 4
End Sub

Function GetNextYear(oStr01 As String, oStr02 As String, oStr03 As String, oStr04 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrYearPay
Dim arrYearPaySet
Dim ii As Integer
Dim GetNextYear1 As String
Dim GetNextYear2 As String
GetNextYear1 = ""
GetNextYear2 = ""
StrSQLa = "Select PA24, PA25, PA57, PA72, PA08, NA21, NA23, NA25,pa09,pa21 From Patent, Nation Where PA09=NA01 And " & ChgPatent(oStr01 & oStr02 & oStr03 & oStr04)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If ("" & rsA("PA08").Value = "1" And "" & rsA("NA21").Value <> "") Or ("" & rsA("pa08").Value = "2" And "" & rsA("NA23").Value <> "") Or ("" & rsA("pa08").Value = "3" And "" & rsA("NA25").Value <> "") Then
        '若有繳年費記錄
        If "" & rsA("PA72").Value <> "" Then
            arrYearPay = Split("" & rsA("PA72").Value, ",")
            If "" & rsA("PA08").Value = "1" Then
                arrYearPaySet = Split("" & rsA("NA21").Value, ",")
            ElseIf rsA("PA08").Value = "2" Then
                arrYearPaySet = Split("" & rsA("NA23").Value, ",")
               If rsA("PA09").Value = "000" And rsA("PA21").Value <> "" Then
                  If rsA("PA21").Value < "20040701" Then
                     arrYearPaySet = Split("1,2,3,4,5,6,7,8,9,10,11,12", ",")
                  End If
               End If
            Else
                arrYearPaySet = Split("" & rsA("NA25").Value, ",")
            End If
            For ii = LBound(arrYearPaySet) To UBound(arrYearPaySet)
                If Val(arrYearPaySet(ii)) > Val(arrYearPay(UBound(arrYearPay))) Then
                    GetNextYear2 = "(" & arrYearPaySet(ii) & ")"
                    Exit For
                Else
                    GetNextYear2 = ""
                End If
            Next ii
        '若尚未繳年費
        Else
            If "" & rsA("PA08").Value = "1" And "" & rsA("NA21").Value <> "" Then
                arrYearPaySet = Split("" & rsA("NA21").Value, ",")
                GetNextYear2 = "(" & arrYearPaySet(LBound(arrYearPaySet)) & ")"
            ElseIf rsA("PA08").Value = "2" And "" & rsA("NA23").Value <> "" Then
                arrYearPaySet = Split("" & rsA("NA23").Value, ",")
                GetNextYear2 = "(" & arrYearPaySet(LBound(arrYearPaySet)) & ")"
            ElseIf rsA("PA08").Value = "3" And "" & rsA("NA25").Value <> "" Then
                arrYearPaySet = Split("" & rsA("NA25").Value, ",")
                GetNextYear2 = "(" & arrYearPaySet(LBound(arrYearPaySet)) & ")"
            End If
        End If
        StrSqlB = "Select Min(NP08) From NextProgress Where " & ChgNextProgress(oStr01 & oStr02 & oStr03 & oStr04) & " And NP07 In (605,606,607) And NP06 Is Null And NP08>=" & strSrvDate(1)
        rsB.CursorLocation = adUseClient
        rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
        If rsB.RecordCount > 0 Then
            GetNextYear1 = ChangeTStringToTDateString(ChangeWStringToTString("" & rsB.Fields(0).Value))
        End If
        If rsB.State <> adStateClosed Then rsB.Close
        Set rsB = Nothing
        If GetNextYear1 = "" Then GetNextYear2 = ""
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
GetNextYear = GetNextYear1 & GetNextYear2
End Function

'不同群組要分不同張
Sub PrintData(ByVal oStd As Long, ByVal oEnd As Long)
Dim oKk As Long
Dim oLl As Long
Dim DoPrint As Boolean
iPage = 1   '每個家族都會歸零
IfPrint = False

iMaxLine = 52
'add by nickc 2007/10/04 將兩頁單筆合併 IsCan2Page 觸發順序有影響，不可以亂搬移
If IsCan2Page = False Then
    If oEnd - oStd = 0 Then
        If chk(0).Value = vbUnchecked Then
            If InStr(1, grd1.TextMatrix(oStd, 10), "駁") <> 0 Then
                Exit Sub
            End If
        End If
        If chk(1).Value = vbUnchecked Then
            If InStr(1, grd1.TextMatrix(oStd, 10), "閉") <> 0 Then
                Exit Sub
            End If
        End If
        iLine = 1
        PrintTitle SeekShowPic
        IsCan2Page = True
    Else
        iLine = 1
        PrintTitle SeekShowPic
        IsCan2Page = False
    End If
Else
    If oEnd - oStd = 0 Then
        If chk(0).Value = vbUnchecked Then
            If InStr(1, grd1.TextMatrix(oStd, 10), "駁") <> 0 Then
                Exit Sub
            End If
        End If
        If chk(1).Value = vbUnchecked Then
            If InStr(1, grd1.TextMatrix(oStd, 10), "閉") <> 0 Then
                Exit Sub
            End If
        End If
        iLine = 26
        PrintTitle SeekShowPic
        IsCan2Page = False
    Else
        Printer.NewPage
        iLine = 1
        IsCan2Page = False
        PrintTitle SeekShowPic
    End If
End If

For oKk = oStd To oEnd
    strTemp(1) = Replace(Replace(grd1.TextMatrix(oKk, 1), "-0-00", ""), "-00", "")
    strTemp(2) = StrToStr(grd1.TextMatrix(oKk, 2), 6)
    strTemp(3) = StrToStr(grd1.TextMatrix(oKk, 3), 33)
    strTemp(4) = StrToStr(grd1.TextMatrix(oKk, 4), 5)
    strTemp(5) = grd1.TextMatrix(oKk, 5)
    strTemp(6) = StrToStr(grd1.TextMatrix(oKk, 6), 7.5)
    strTemp(7) = StrToStr(grd1.TextMatrix(oKk, 7), 7.5)
    strTemp(8) = grd1.TextMatrix(oKk, 8)
    strTemp(9) = StrToStr(grd1.TextMatrix(oKk, 9), 34)
    strTemp(10) = grd1.TextMatrix(oKk, 10)
    strTemp(11) = grd1.TextMatrix(oKk, 11)
    strTemp(12) = grd1.TextMatrix(oKk, 12)
    'add by nickc 2007/10/04 加入核駁或是閉卷控制
    DoPrint = True
    If chk(0).Value = vbUnchecked And DoPrint = True Then
        If InStr(1, strTemp(10), "駁") <> 0 Then
            DoPrint = False
        End If
    End If
    If chk(1).Value = vbUnchecked And DoPrint = True Then
        If InStr(1, strTemp(10), "閉") <> 0 Then
            DoPrint = False
        End If
    End If
    If DoPrint = True Then
        'Modified by Lydia 2019/04/16 CFP分案因為會壓到後面的國家,所以調整字體
        'For oLl = 1 To 12
        If InStrRev(strTemp(1), "-") < 9 Then
            Printer.Font.Size = 12
        Else
            Printer.Font.Size = 10
        End If
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iLine * 300
        Printer.Print strTemp(1)
        Printer.Font.Size = 12
        For oLl = 2 To 12
        'end 2019/04/116
            If oLl = 4 Or oLl = 9 Then
                iLine = iLine + 1
            End If
            Printer.CurrentX = PLeft(oLl)
            Printer.CurrentY = iLine * 300
            'Modified by Lydia 2022/07/04 逐字檢查Unicode文字改以圖片方式列印
            'Printer.Print strTemp(oLl)
            Xo = Printer.CurrentX
            Yo = Printer.CurrentY
            PUB_PrintUnicodeText strTemp(oLl), Xo, Yo, 0
            'end 2022/07/04
        Next oLl
        iLine = iLine + 1
        PrintEnd
        If iLine >= iMaxLine Then
            iPage = iPage + 1
            Printer.NewPage
            PrintTitle SeekShowPic
        End If
    End If
Next oKk
IfPrint = True
IfPrintAnyOne = True
End Sub

Sub PrintTitle(oSeekShowPic As String)
GetPleft
Dim oCUID As String
Dim oCUAddr As String
Dim oCUTel As String
Dim oCUFax As String
Dim oCUMail As String
Dim oCUMan As String
Dim oRS As New ADODB.Recordset

If iPage = 1 Then
    If NowKey1 <> "" Then
         'Modify by Morgan 2008/7/24 接洽人改先用聯絡人編號抓聯絡人檔
         'strSQL = "select cu01||cu02||' '||nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),cu31,cu16,cu18,cu20,cu08 from patent,customer where substr(pa26,1,8)=cu01(+) and substr(pa26||'000000',9,1)=cu02(+) and pa01='" & SystemNumber(NowKey1, 1) & "' and pa02='" & SystemNumber(NowKey1, 2) & "' and pa03='" & SystemNumber(NowKey1, 3) & "' and pa04='" & SystemNumber(NowKey1, 4) & "' "
         strSql = "select cu01||cu02||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu31,cu16,cu18,cu20,nvl(pcc05,cu08) from patent,customer,potcustcont where substr(pa26,1,8)=cu01(+) and substr(pa26||'000000',9,1)=cu02(+) and pa01='" & SystemNumber(NowKey1, 1) & "' and pa02='" & SystemNumber(NowKey1, 2) & "' and pa03='" & SystemNumber(NowKey1, 3) & "' and pa04='" & SystemNumber(NowKey1, 4) & "' and pcc01(+)=cu01 and pcc02(+)=cu127"
    Else
         'Modify by Morgan 2008/7/24 接洽人改先用聯絡人編號抓聯絡人檔
         'strSQL = "select cu01||cu02||' '||nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),cu31,cu16,cu18,cu20,nvl(pcc05,cu08) from customer,potcustcont where substr('" & NowKey2 & "'||'000000',1,8)=cu01(+) and substr('" & NowKey2 & "'||'000000',9,1)=cu02(+) "
         strSql = "select cu01||cu02||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu31,cu16,cu18,cu20,nvl(pcc05,cu08) from customer,potcustcont where substr('" & NowKey2 & "'||'000000',1,8)=cu01(+) and substr('" & NowKey2 & "'||'000000',9,1)=cu02(+) and pcc01(+)=cu01 and pcc02(+)=cu127"
    End If
    Set oRS = New ADODB.Recordset
    If oRS.State = 1 Then oRS.Close
    oRS.CursorLocation = adUseClient
    oRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If oRS.RecordCount <> 0 Then
        oCUID = CheckStr(oRS.Fields(0))
        oCUAddr = CheckStr(oRS.Fields(1))
        oCUTel = CheckStr(oRS.Fields(2))
        oCUFax = CheckStr(oRS.Fields(3))
        oCUMail = CheckStr(oRS.Fields(4))
        oCUMan = CheckStr(oRS.Fields(5))
    Else
        oCUID = ""
        oCUAddr = ""
        oCUTel = "'"
        oCUFax = "'"
        oCUMail = ""
        oCUMan = ""
    End If
    oRS.Close
    Set oRS = Nothing
    'add by nickc 2007/10/04 控制兩頁合併
    If IsCan2Page = False Then
        Printer.Font.Name = "標楷體" 'Add by Amy 2018/05/09 字型影響文字是否黏在一起
        Printer.Font.Size = 18
        Printer.Font.Underline = True
        Printer.FontBold = True
        Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("客戶專利案件整理表") / 2)
        Printer.CurrentY = 300
        Printer.Print "客戶專利案件整理表"
        Printer.Font.Size = 12
        Printer.Font.Underline = False
        Printer.FontBold = False
    End If
    Printer.CurrentX = 200
    Printer.CurrentY = 800 + IIf(IsCan2Page = False, 0, 7800)
    'Modified by Lydia 2022/07/04 逐字檢查Unicode文字改以圖片方式列印
    'Printer.Print "收件人：" & oCUID
    Xo = Printer.CurrentX
    Yo = Printer.CurrentY
    PUB_PrintUnicodeText "收件人：" & oCUID, Xo, Yo, 0
    'end 2022/07/04
    Printer.CurrentX = 200
    Printer.CurrentY = 1400 + IIf(IsCan2Page = False, 0, 7800)
    'Modified by Lydia 2022/07/04 逐字檢查Unicode文字改以圖片方式列印
    'Printer.Print "地　址：" & oCUAddr
    Xo = Printer.CurrentX
    Yo = Printer.CurrentY
    PUB_PrintUnicodeText "地　址：" & oCUAddr, Xo, Yo, 0
    'end 2022/07/04
    Printer.CurrentX = 200
    Printer.CurrentY = 2000 + IIf(IsCan2Page = False, 0, 7800)
    Printer.Print "電　話：" & oCUTel
    Printer.CurrentX = 200
    Printer.CurrentY = 2600 + IIf(IsCan2Page = False, 0, 7800)
    Printer.Print "傳　真：" & oCUFax
    Printer.CurrentX = 200
    Printer.CurrentY = 3200 + IIf(IsCan2Page = False, 0, 7800)
    Printer.Print "E-Mail ：" & oCUMail
    Printer.CurrentX = 200
    Printer.CurrentY = 3800 + IIf(IsCan2Page = False, 0, 7800)
    'Modified by Lydia 2022/07/04 逐字檢查Unicode文字改以圖片方式列印
    'Printer.Print "接洽人：" & oCUMan
    Xo = Printer.CurrentX
    Yo = Printer.CurrentY
    PUB_PrintUnicodeText "接洽人：" & oCUMan, Xo, Yo, 0
    'end 2022/07/04
    Printer.CurrentX = 200
    Printer.CurrentY = 4400 + IIf(IsCan2Page = False, 0, 7800)
    Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
    Dim tObj As New StdPicture
    Dim bytes() As Byte
    Dim file_num As Integer
    Dim rsPic As New ADODB.Recordset
    Dim pTop As Integer
    Dim pWidth As Integer
    Dim pWidthNew As Integer
    Dim Pleft2 As Integer
    Dim pHeight As Integer
    Dim pHeightNew As Integer
    Dim IsWmf As Boolean
    Dim bSuccess As Boolean
    Set rsPic = New ADODB.Recordset
    rsPic.CursorLocation = adUseClient
    rsPic.Open "select * from ImgByteFile where ibf01='" & SystemNumber(oSeekShowPic, 1) & "' and ibf02='" & SystemNumber(oSeekShowPic, 2) & "' and ibf03='" & SystemNumber(oSeekShowPic, 3) & "' and ibf04='" & SystemNumber(oSeekShowPic, 4) & "' and ibf05='1' ", cnnConnection, adOpenStatic, adLockOptimistic
    If rsPic.RecordCount <> 0 Then
      If CheckStr(rsPic.Fields("ibf06")) = "1" Or CheckStr(rsPic.Fields("ibf06")) = "2" Then
         IsWmf = False
      Else
         IsWmf = True
      End If
      'Add By Sindy 2017/8/10
'      If "" & rsPic.Fields("IBF15") <> "" Then
         Call PUB_GetFtpFile(rsPic.Fields("IBF15"), App.path & "\tmp." & IIf(IsWmf, "wmf", "jpg"), UCase("ImgByteFile"))
'      Else
'      '2017/8/10 END
'         ReDim bytes(Val(rsPic.Fields("ibf13").Value))
'         bytes() = rsPic.Fields("ibf14").GetChunk(Val(rsPic.Fields("ibf13").Value))
'         file_num = FreeFile
'         Open App.path & "\tmp." & IIf(IsWmf, "wmf", "jpg") For Binary Access Write As #file_num
'            Put #file_num, , bytes()
'         Close #file_num
'      End If
        Set tObj = pvGetStdPicture(App.path & "\tmp." & IIf(IsWmf, "wmf", "jpg"))
        pTop = 800 + IIf(IsCan2Page = False, 0, 7800)
        Pleft2 = 7500
        pWidth = 3500
        pHeight = 4000
        If tObj.Height >= tObj.Width Then  '以高為準
            pWidthNew = tObj.Width / (tObj.Height / pHeight)
            Pleft2 = Pleft2 + ((pWidth - pWidthNew) / 2)
            pWidth = pWidthNew
        Else   '已寬為準
            pHeightNew = tObj.Height / (tObj.Width / pWidth)
            pTop = pTop + ((pHeight - pHeightNew) / 2)
            pHeight = pHeightNew
        End If
        '因為標準是 35 行 所以總行數 - 35 在乘上行高
        Printer.PaintPicture tObj, Pleft2, pTop, pWidth, pHeight
        Printer.Line (Pleft2 - 1, pTop - 1)-(Pleft2 + pWidth + 1, pTop + pHeight + 1), , B
        If Dir(App.path & "\tmp." & IIf(IsWmf, "wmf", "jpg")) <> "" Then
           Kill App.path & "\tmp." & IIf(IsWmf, "wmf", "jpg")
        End If
    End If
    iLine = 17 + IIf(IsCan2Page = False, 0, (iMaxLine / 2))
'add by nickc 2008/02/21 因為跳頁沒控制到
Else
    iLine = 1
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "國家種類"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "專利名稱"
iLine = iLine + 1
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "客戶案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "申請日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "申請號"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "證書號"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "專利期限"
iLine = iLine + 1
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "申請人"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "准駁及閉卷"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "下次繳費日期"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "分所號"
iLine = iLine + 1
PrintEnd
'Printer.EndDoc
End Sub

Sub GetPleft()
PLeft(1) = 0 + 200
PLeft(2) = 1200 + 400 + 200 + 100
PLeft(3) = 3000 + 400 + 200 + 100
PLeft(4) = 0 + 200
PLeft(5) = 1200 + 400 + 200 + 100
PLeft(6) = 3000 + 400 + 200 + 100
PLeft(7) = 5500 + 400 + 200 + 100
PLeft(8) = 8000 + 400 + 200 + 100
PLeft(9) = 0 + 200
PLeft(10) = 5500 + 400 + 200 + 100
PLeft(11) = 8000 ' + 400 + 200 + 100 'Modify by Amy 2018/05/09 字會黏在一起Ex:CFP-019617
PLeft(12) = 9500 + 400 + 200 + 100
End Sub

Sub PrintEnd()
Printer.CurrentX = 200
Printer.CurrentY = iLine * 300
Printer.Print String(144, "-")
iLine = iLine + 1
End Sub

'若是離職或是虛建智權人員將傳回主管
Function ChkUserId(oOldUserID) As String
ChkUserId = ""
Dim oStrSQL As String
Dim UserRS As New ADODB.Recordset
'先檢查是否為員工
oStrSQL = "select * from staff where st01='" & oOldUserID & "' "
Set UserRS = New ADODB.Recordset
If UserRS.State = 1 Then UserRS.Close
UserRS.CursorLocation = adUseClient
UserRS.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
If UserRS.RecordCount <> 0 Then
    '檢查在不在職
    oStrSQL = "select * from staff where st01='" & oOldUserID & "' and st04='1' and st01>'63001' "
    Set UserRS = New ADODB.Recordset
    If UserRS.State = 1 Then UserRS.Close
    UserRS.CursorLocation = adUseClient
    UserRS.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If UserRS.RecordCount <> 0 Then
        ChkUserId = oOldUserID
    Else
        '2009/4/30 modify by sonia
        'oStrSQL = "select a0908 from acc090 where a0901=(select st15 from staff where st01='" & oOldUserID & "' ) and st04='1'  "
        'Added by Lydia 2023/12/28
        If strSrvDate(1) >= 新部門啟用日 Then
           oStrSQL = "select a0924,'1' as ord1 from acc090new,staff where a0921=(select st93 from staff where st01='" & oOldUserID & "') and a0924=st01(+) and st04='1' " & _
                     "union all select a0908,'2' as ord1 from acc090,staff where a0901=(select st15 from staff where st01='" & oOldUserID & "') and a0908=st01(+) and st04='1' " & _
                     "order by ord1"
        Else
        'end 2023/12/28
           oStrSQL = "select a0908 from acc090,staff where a0901=(select st15 from staff where st01='" & oOldUserID & "' ) and a0908=st01 and st04='1' "
        End If
        Set UserRS = New ADODB.Recordset
        If UserRS.State = 1 Then UserRS.Close
        UserRS.CursorLocation = adUseClient
        UserRS.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
        If UserRS.EOF And UserRS.BOF Then
            '完全沒有
            ChkUserId = "83002"
        Else
            ChkUserId = CheckStr(UserRS.Fields(0))
        End If
    End If
Else
    ChkUserId = oOldUserID
End If
End Function

'add by nickc 2007/10/17 檢查獨立
Function ChkCanPrint() As Boolean
Dim tmpUserId As String
If Pub_StrUserSt03 <> "M51" Then
   '檢查密碼
   'Me.Tag = ""
   'Modified by Lydia 2023/12/28
   'frm210106_1.setCaller Me
   'frm210106_1.Caption = "客戶專利案件整理表-智權人員登入"
   'frm210106_1.Show vbModal
    Me.Hide
    If Me.Tag = "" Then
       bolRunPWD = False 'Added by Lydia 2024/01/12
       Call frm210106_1.setCaller(frm210115, Me)
       frm210106_1.Show
    'Added by Lydia 2024/01/12
    Else
       ChkCanPrint = True
       Me.Show
    'end 2024/01/12
    End If
   'end 2023/12/28
   
   'Mark by Lydia 2023/12/28 以下改成模組PubShowNextData
'   If Trim(Me.Tag) = "" Then Exit Function
'
'   '檢查智權人員是否與條件相同
'   If NowKey1 <> "" Then
'       strSql = "select cu13 from patent,customer where substr(pa26,1,8)=cu01(+) and substr(pa26||'000000',9,1)=cu02(+) and pa01='" & SystemNumber(NowKey1, 1) & "' and pa02='" & SystemNumber(NowKey1, 2) & "' and pa03='" & SystemNumber(NowKey1, 3) & "' and pa04='" & SystemNumber(NowKey1, 4) & "' "
'   Else
'       strSql = "select cu13 from customer where substr('" & NowKey2 & "'||'000000',1,8)=cu01(+) and substr('" & NowKey2 & "'||'000000',9,1)=cu02(+) "
'   End If
'
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'       tmpUserId = ChkUserId(CheckStr(RsTemp.Fields(0)))
'       'Modified by Morgan 2013/8/22 改以共用函數抓有權限的員工來判斷
'       'If Me.Tag <> tmpUserId Then
'       'Modify by Amy 2014/05/22
'       If bolSpecMan = True Then
'            'Add by Amy 2015/02/24 +總經理業務工作代理人員
'            If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'                strExc(1) = "'" & Replace(Pub_GetSpecMan("總經理員工編號"), ";", "','") & "','" & Me.Tag & "'"
'            '開放專利處部份智權同仁資料給彥葶代為處理
'            ElseIf InStr(strSpecCode, "A8") > 0 Then
'                strExc(1) = "'" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "','" & Me.Tag & "'"
'            End If
'       Else
'            strExc(1) = PUB_GetSalesList(Me.Tag)
'       End If
'       'end 2014/05/22
'       If InStr(strExc(1), "'" & tmpUserId & "'") = 0 Then
'       'end 2013/8/22
'           If tmpUserId <> CheckStr(RsTemp.Fields(0)) Then
'               MsgBox "權限不足！" & vbCrLf & "請主管輸入驗證資料！", vbExclamation, "嚴重錯誤！"
'               Exit Function
'           Else
'               MsgBox "權限不足！" & vbCrLf & "不允許列印非屬於自己的客戶資料！", vbExclamation, "嚴重錯誤！"
'               Exit Function
'           End If
'       Else
'            ChkCanPrint = True
'       End If
'   Else
'       MsgBox "發生錯誤，異常資料！" & vbCrLf & "沒有智權人員！" & vbCrLf & "請將相關資訊提供電腦中心檢查！", vbCritical, "嚴重錯誤！"
'       Exit Function
'   End If
Else
    ChkCanPrint = True
End If
End Function

'Added by Lydia 2023/12/28
Public Sub PubShowNextData()
Dim tmpUserId As String

   If Trim(Me.Tag) = "" Then Exit Sub
   
   '檢查智權人員是否與條件相同
   If NowKey1 <> "" Then
       strSql = "select cu13 from patent,customer where substr(pa26,1,8)=cu01(+) and substr(pa26||'000000',9,1)=cu02(+) and pa01='" & SystemNumber(NowKey1, 1) & "' and pa02='" & SystemNumber(NowKey1, 2) & "' and pa03='" & SystemNumber(NowKey1, 3) & "' and pa04='" & SystemNumber(NowKey1, 4) & "' "
   Else
       strSql = "select cu13 from customer where substr('" & NowKey2 & "'||'000000',1,8)=cu01(+) and substr('" & NowKey2 & "'||'000000',9,1)=cu02(+) "
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       tmpUserId = ChkUserId(CheckStr(RsTemp.Fields(0)))
       'Modified by Morgan 2013/8/22 改以共用函數抓有權限的員工來判斷
       'If Me.Tag <> tmpUserId Then
       'Modify by Amy 2014/05/22
       If bolSpecMan = True Then
            'Add by Amy 2015/02/24 +總經理業務工作代理人員
            If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
                strExc(1) = "'" & Replace(Pub_GetSpecMan("總經理員工編號"), ";", "','") & "','" & Me.Tag & "'"
            '開放專利處部份智權同仁資料給彥葶代為處理
            ElseIf InStr(strSpecCode, "A8") > 0 Then
                strExc(1) = "'" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "','" & Me.Tag & "'"
            End If
       Else
            strExc(1) = PUB_GetSalesList(Me.Tag)
       End If
       'end 2014/05/22
       If InStr(strExc(1), "'" & tmpUserId & "'") = 0 Then
       'end 2013/8/22
           If tmpUserId <> CheckStr(RsTemp.Fields(0)) Then
               MsgBox "權限不足！" & vbCrLf & "請主管輸入驗證資料！", vbExclamation, "嚴重錯誤！"
               Exit Sub
           Else
               MsgBox "權限不足！" & vbCrLf & "不允許列印非屬於自己的客戶資料！", vbExclamation, "嚴重錯誤！"
               Exit Sub
           End If
       Else
            'ChkCanPrint = True
       End If
   Else
       MsgBox "發生錯誤，異常資料！" & vbCrLf & "沒有智權人員！" & vbCrLf & "請將相關資訊提供電腦中心檢查！", vbCritical, "嚴重錯誤！"
       Exit Sub
   End If
End Sub
