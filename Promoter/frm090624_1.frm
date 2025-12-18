VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090624_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利處每週速度考核表"
   ClientHeight    =   5724
   ClientLeft      =   1656
   ClientTop       =   1512
   ClientWidth     =   9312
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9312
   Visible         =   0   'False
   Begin VB.CommandButton cmdok 
      Caption         =   "放大"
      Height          =   400
      Index           =   3
      Left            =   4395
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "重算(&R)"
      Height          =   400
      Index           =   2
      Left            =   5625
      TabIndex        =   8
      Top             =   0
      Width           =   1230
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel(&E)"
      Height          =   400
      Index           =   1
      Left            =   6870
      TabIndex        =   7
      Top             =   0
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2565
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   8100
      TabIndex        =   0
      Top             =   0
      Width           =   1140
   End
   Begin TabDlg.SSTab stb 
      Height          =   5265
      Left            =   30
      TabIndex        =   4
      Top             =   450
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   9292
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "承辦人"
      TabPicture(0)   =   "frm090624_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "繪圖人員"
      TabPicture(1)   =   "frm090624_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd(1)"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd 
         Height          =   4800
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   9135
         _ExtentX        =   16108
         _ExtentY        =   8467
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedRows       =   3
         FixedCols       =   2
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd 
         Height          =   4800
         Index           =   1
         Left            =   -74940
         TabIndex        =   6
         Top             =   360
         Width           =   9135
         _ExtentX        =   16108
         _ExtentY        =   8467
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedRows       =   3
         FixedCols       =   2
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Label lblMonth 
      Caption         =   "lblMonth"
      Height          =   180
      Left            =   1140
      TabIndex        =   3
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "考核月份： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frm090624_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grd(0)改字型=新細明體-ExtB、grd(1)改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify by Morgan 2010/12/30 新舊制選項代碼對調,件數->基數
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim m_dblRow As Double
Dim m_dblCol As Double
Dim m_IsRun As Boolean
'Modify by Morgan 2010/7/2 Union 會 distinct 資料而造成少算(如不特定案件的支援統計會重複),故改為 Union all
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
Dim m_bol108Rule As Boolean 'Added by Morgan 2019/3/22 專利處108考核

Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer
    
    Select Case Index
    Case 0 '回前畫面
        'edit by nickc 2005/03/01  加入結束功能
'        If cmdok(Index).Caption = "結束(&X)" Then
'            Unload frm090624
'            Unload Me
'        Else
            frm090624.Show
            Unload Me
'        End If
    Case 1 '產生Excel
        Screen.MousePointer = vbHourglass
'不用 edit by nickc 2005/03/02
'        With Me.grd(0)
'            If .TextMatrix(0, 2) <> "" Then
'                '更新資料
'                For ii = 2 To .Cols - 1
'                    UpdateMonthAssess "1", .TextMatrix(34, ii), Val(frm090624.txt1(0).Text) + 191100, Val(.TextMatrix(33, ii))
'                Next ii
'            End If
'        End With
'        With Me.grd(1)
'            If .TextMatrix(0, 2) <> "" Then
'                '更新資料
'                For ii = 2 To .Cols - 1
'                    UpdateMonthAssess "2", .TextMatrix(34, ii), Val(frm090624.txt1(0).Text) + 191100, Val(.TextMatrix(33, ii))
'                Next ii
'            End If
'        End With
        ExcelSave
        Screen.MousePointer = vbDefault
    'add by nickc 2005/03/01 加入重算
    Case 2
         Screen.MousePointer = vbHourglass
         grd(0).MousePointer = flexHourglass
         grd(1).MousePointer = flexHourglass
         DoEvents
         Me.Enabled = False
         grd(0).Clear
         grd(1).Clear
         SetGrd
         ProcessNew
         Me.Enabled = True
         grd(1).MousePointer = flexDefault
         grd(0).MousePointer = flexDefault
         Screen.MousePointer = vbDefault
    'add by nickc 2005/08/03 放大
    Case 3
         frm090624_1.Top = 0
         frm090624_1.Left = 0
         frm090624_1.Width = mdiMain.ScaleWidth
         frm090624_1.Height = mdiMain.ScaleHeight
         frm090624_1.stb.Width = frm090624_1.ScaleWidth - 100
         frm090624_1.stb.Height = frm090624_1.ScaleHeight - 400
         'frm090624_1.stb.Tab = 1
         frm090624_1.grd(0).Width = frm090624_1.stb.Width - 150
         frm090624_1.grd(0).Height = frm090624_1.stb.Height - 350
         frm090624_1.grd(1).Width = frm090624_1.stb.Width - 150
         frm090624_1.grd(1).Height = frm090624_1.stb.Height - 350
         cmdok(3).Enabled = False
    Case Else
    End Select
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
If m_IsRun = False Then
   m_IsRun = True
         grd(0).Clear
         grd(1).Clear
          SetGrd
          If frm090624.txt1(9) = "1" Then
            cmdok(1).Visible = False
            
            pub_QL05 = pub_QL05 & ";" & Left(frm090624.Label2, 5) & "1：新制" 'Add By Sindy 2010/12/17
            Select Case ProState
            Case "2"
                  '只有協理和郭的等級可以重算
                  'Modifed by Lydia 2023/04/23 修改王副總退休之相關控制
                  'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "73" Or PUB_GetST05(strUserNum) = "00" Then
                  'Modified by Morgan 2025/2/4 +P10部門
                  'Modified by Morgan 2025/6/26 +79075
                  If InStr("71,73,00,", Pub_strUserST05 & ",") > 0 Or (strSrvDate(1) >= "20230501" And Pub_strUserST05 = "72") Or Pub_StrUserSt03 = "P10" Or strUserNum = "79075" Then
                        cmdok(2).Visible = True
                        cmdok(1).Visible = True 'Added by Morgan 2019/11/1
                  Else
                        cmdok(2).Visible = False
                  End If
              'add by nickc 2005/08/03 繪圖主管可以放大
              If PUB_GetST05(strUserNum) = "81" Or PUB_GetST05(strUserNum) = "82" Or PUB_GetST05(strUserNum) = "00" Then
                  cmdok(3).Visible = True
              Else
                  cmdok(3).Visible = False
              End If
          '新制不用重算，重算用重算鍵
      'edit by nick
      '      ProcessNew  '新制
            Case Else
                  cmdok(0).Caption = "結束(&X)"
                  cmdok(2).Visible = False
            End Select
            
            If StrMenu = False Then
               GoTo GoExit
            End If
          Else
            pub_QL05 = pub_QL05 & ";" & Left(frm090624.Label2, 5) & "2：舊制" 'Add By Sindy 2010/12/17
            Screen.MousePointer = vbHourglass
            DoEvents
            Process         '舊制
            Screen.MousePointer = vbDefault
            cmdok(1).Visible = True
            cmdok(2).Visible = False
          End If
          ChgGrdColor
End If
          Exit Sub
GoExit:
         cmdok_Click (0)
End Sub

Private Sub Form_Load()
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   m_IsRun = False
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
   Me.lblMonth.Caption = frm090624.txt1(0).Text
   
   'Added by Morgan 2019/3/25
   If Val(frm090624.txt1(0)) + 191100 >= Val(Left(PUB_108RuleDate, 6)) Then
      m_bol108Rule = True
   Else
      m_bol108Rule = False
   End If
   'end 2019/3/25
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090624_1 = Nothing
End Sub

Private Sub SetGrd()
Dim ii As Integer, jj As Integer

    For ii = 0 To Me.grd.Count - 1
        With Me.grd(ii)
            .Visible = False
            'Modify by Morgan 2009/7/14 +累計會稿量
            '.Rows = 35
            .Rows = 36
            .Cols = 3
            .TextMatrix(0, 0) = Val(Right(Me.lblMonth.Caption, 2))
            .TextMatrix(0, 1) = "月份"
            .TextMatrix(1, 0) = Val(Right(Me.lblMonth.Caption, 2))
            .TextMatrix(1, 1) = "月份"
            .TextMatrix(2, 0) = "目標": .TextMatrix(2, 1) = "基數"
            .TextMatrix(3, 0) = "本月": .TextMatrix(3, 1) = "工作天數"
            
            .TextMatrix(4, 0) = "第一週": .TextMatrix(4, 1) = "本週工作天數"
            .TextMatrix(5, 0) = "第一週": .TextMatrix(5, 1) = "本週目標"
            .TextMatrix(6, 0) = "第一週": .TextMatrix(6, 1) = "本週完成"
            .TextMatrix(7, 0) = "第一週": .TextMatrix(7, 1) = "本週得分"
            .TextMatrix(8, 0) = "第一週": .TextMatrix(8, 1) = "達成比例"
            .TextMatrix(9, 0) = "第二週": .TextMatrix(9, 1) = "本週工作天數"
            .TextMatrix(10, 0) = "第二週": .TextMatrix(10, 1) = "本週目標"
            .TextMatrix(11, 0) = "第二週": .TextMatrix(11, 1) = "本週完成"
            .TextMatrix(12, 0) = "第二週": .TextMatrix(12, 1) = "本週達成比例"
            .TextMatrix(13, 0) = "第二週": .TextMatrix(13, 1) = "本週得分"
            .TextMatrix(14, 0) = "第二週": .TextMatrix(14, 1) = "累計目標"
            .TextMatrix(15, 0) = "第二週": .TextMatrix(15, 1) = "累計完成"
            .TextMatrix(16, 0) = "第二週": .TextMatrix(16, 1) = "累計達成比例"
            
            .TextMatrix(17, 0) = "第三週": .TextMatrix(17, 1) = "本週工作天數"
            .TextMatrix(18, 0) = "第三週": .TextMatrix(18, 1) = "本週目標"
            .TextMatrix(19, 0) = "第三週": .TextMatrix(19, 1) = "本週完成"
            .TextMatrix(20, 0) = "第三週": .TextMatrix(20, 1) = "本週達成比例"
            .TextMatrix(21, 0) = "第三週": .TextMatrix(21, 1) = "本週得分"
            .TextMatrix(22, 0) = "第三週": .TextMatrix(22, 1) = "累計目標"
            .TextMatrix(23, 0) = "第三週": .TextMatrix(23, 1) = "累計完成"
            .TextMatrix(24, 0) = "第三週": .TextMatrix(24, 1) = "累計達成比例"
            
            .TextMatrix(25, 0) = "第四週": .TextMatrix(25, 1) = "本週工作天數"
            .TextMatrix(26, 0) = "第四週": .TextMatrix(26, 1) = "本週目標"
            .TextMatrix(27, 0) = "第四週": .TextMatrix(27, 1) = "本週完成"
            .TextMatrix(28, 0) = "第四週": .TextMatrix(28, 1) = "本週達成比例"
            .TextMatrix(29, 0) = "第四週": .TextMatrix(29, 1) = "本週得分"
            .TextMatrix(30, 0) = "第四週": .TextMatrix(30, 1) = "累計目標"
            'Added by Morgan 2019/3/18 108考核(工程師本月得分改以整月達成率計算)
            If ii = 0 And m_bol108Rule Then
               For jj = 4 To 30
                  grd(ii).RowHeight(jj) = 0
               Next
               .TextMatrix(31, 0) = "本月": .TextMatrix(31, 1) = "完成"
               .TextMatrix(32, 0) = "本月": .TextMatrix(32, 1) = "達成比例"
               .TextMatrix(33, 0) = "本月": .TextMatrix(33, 1) = "得分"
            Else
            'end 2019/3/18
               .TextMatrix(31, 0) = "第四週": .TextMatrix(31, 1) = "累計完成"
               .TextMatrix(32, 0) = "第四週": .TextMatrix(32, 1) = "累計達成比例"
               .TextMatrix(33, 0) = "本月": .TextMatrix(33, 1) = "得分平均"
               
            End If 'Added by Morgan 2019/3/18
            
            .TextMatrix(34, 0) = "員工": .TextMatrix(34, 1) = "編號"
            .RowHeight(34) = 0
            
            'Add by Morgan 2009/7/14 +累計會稿量
            .TextMatrix(35, 0) = "本月": .TextMatrix(35, 1) = "累計會稿量"
            .RowHeight(35) = 0
            
            .MergeCells = flexMergeRestrictRows
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeRow(1) = True
            .MergeCol(1) = True
            'add by nickc 2005/03/02
            If frm090624.txt1(9) = "1" And ii = 1 Then
               .MergeRow(33) = True
            End If
            .ColWidth(0) = 800
            .ColWidth(1) = 1200
            .Visible = True
        End With
    Next ii
    
    '預設目前在第一筆的位置
    With Me.grd(0)
        .row = 0
        .col = 2
    End With
End Sub

'舊制  2005/02/25 nickc 加註解
Private Sub Process()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer
Dim intColCount As Integer
Dim strST01 As String, strST03 As String, strST06 As String, strST13 As String, strST16 As String, strNCFPGoal As String, strCFPGoal As String, strGoal As String
    
    If frm090624.txt1(0) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(frm090624.Label1(2), 5) & frm090624.txt1(0) & "(以" & frm090624.Label1(5) & "計算)" 'Add By Sindy 2010/12/17
    End If
    If frm090624.txt1(1) <> "" Or frm090624.txt1(2) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090624.Label1(3) & frm090624.txt1(1) & "-" & frm090624.txt1(2) & "(" & frm090624.Label1(6) & ")" 'Add By Sindy 2010/12/17
    End If
    If frm090624.txt1(3) <> "" Or frm090624.txt1(4) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090624.Label1(0) & frm090624.txt1(3) & "-" & frm090624.txt1(4) & "(" & frm090624.Label1(7) & ")" 'Add By Sindy 2010/12/17
    End If
    If frm090624.txt1(5) <> "" Or frm090624.txt1(6) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090624.Label1(1) & frm090624.txt1(5) & "-" & frm090624.txt1(6) & "(" & frm090624.Label1(8) & ")" 'Add By Sindy 2010/12/17
    End If
    If frm090624.txt1(7) <> "" Or frm090624.txt1(7) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090624.Label1(4) & frm090624.txt1(7) & "-" & frm090624.txt1(8) & "(" & frm090624.Label1(9) & ")" 'Add By Sindy 2010/12/17
    End If
    InsertQueryLog ("") 'Add By Sindy 2010/12/17
    
    '承辦人速度考核
    With Me.grd(0)
        '員工編號
        '排除IAIN(88024)的資料
        '93.12.13 MODIFY BY SONIA 排除外翻人員
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/04/12
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/08/22
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2006/05/01
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'modify by sonia 2014/4/9 加入94007林景郁總經理
        'Modified by Morgan 2022/10/31 +99050
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050') And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        '93.12.13 END
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        intColCount = 0
        While Not rsA.EOF
            .Cols = .Cols + 1
            .TextMatrix(0, ii) = "" & rsA("ST01").Value
            .TextMatrix(34, ii) = "" & rsA("ST01").Value
            .TextMatrix(1, ii) = "稿"
            intColCount = intColCount + 1
            rsA.MoveNext
            ii = ii + 1
        Wend
        .Cols = .Cols - 1
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 2 To intColCount + 2 - 1
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strNCFPGoal = "0": strCFPGoal = "0"
            '非CFP目標件數
            'edit by nickc 2005/04/12
            'strSQLA = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            'MODIFY BY SONIA 2014/4/11 PE02<>'CFP' 改為 PE02='P', 因為杜燕文有T的目標
            'Modified by Morgan 2022/10/31 +99050
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050')) S Where S.ST01=PE01 And PE02='P' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strST01 = "" & rsA.Fields(0).Value
                strST03 = "" & rsA.Fields(1).Value
                strST06 = "" & rsA.Fields(2).Value
                strST13 = "" & rsA.Fields(3).Value
                strST16 = "" & rsA.Fields(4).Value
                strNCFPGoal = "" & rsA.Fields(5).Value
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'CFP目標件數
            'edit by nickc 2005/04/12
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            'Modified by Morgan 2022/10/31 +99050
            StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050')) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
                strST01 = "" & rsB.Fields(0).Value
                strST03 = "" & rsB.Fields(1).Value
                strST06 = "" & rsB.Fields(2).Value
                strST13 = "" & rsB.Fields(3).Value
                strST16 = "" & rsB.Fields(4).Value
                strCFPGoal = "" & rsB.Fields(5).Value
            End If
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            .TextMatrix(2, ii) = CalGoal(strST01, strST03, strST06, strST13, strST16, strNCFPGoal, strCFPGoal)
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(3, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(4, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(5, ii) = Format(Val(.TextMatrix(4, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(5, ii) = "0.00"
               End If
            End If
        Next ii
        '第一週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(6, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(2).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第一週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(5, ii)) <> 0 Then
                  .TextMatrix(8, ii) = Format(Val(.TextMatrix(6, ii)) / Val(.TextMatrix(5, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(8, ii) = "0.00"
               End If
            End If
        Next ii
        '第一週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(7, ii) = CalPoints(Val(Replace(.TextMatrix(8, ii), "%", "")) / 100)
'            End If
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(9, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(10, ii) = Format(Val(.TextMatrix(9, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(10, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(11, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(3).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第二週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(10, ii)) <> 0 Then
                  .TextMatrix(12, ii) = Format(Val(.TextMatrix(11, ii)) / Val(.TextMatrix(10, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(12, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(13, ii) = CalPoints(Val(Replace(.TextMatrix(12, ii), "%", "")) / 100)
'            End If
        Next ii
        '第二週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(14, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))), "0.00")
               Else
                  .TextMatrix(14, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(15, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第二週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(14, ii)) <> 0 Then
                  .TextMatrix(16, ii) = Format(Val(.TextMatrix(15, ii)) / Val(.TextMatrix(14, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(16, ii) = "0.00"
               End If
            End If
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(17, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(18, ii) = Format(Val(.TextMatrix(17, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(18, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(19, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(5).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第三週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(18, ii)) <> 0 Then
                  .TextMatrix(20, ii) = Format(Val(.TextMatrix(19, ii)) / Val(.TextMatrix(18, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(20, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(21, ii) = CalPoints(Val(Replace(.TextMatrix(20, ii), "%", "")) / 100)
'            End If
        Next ii
        '第三週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))) <> 0 Then
                  .TextMatrix(22, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))), "0.00")
               Else
                  .TextMatrix(22, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(23, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第三週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(22, ii)) <> 0 Then
                  .TextMatrix(24, ii) = Format(Val(.TextMatrix(23, ii)) / Val(.TextMatrix(22, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(24, ii) = "0.00"
               End If
            End If
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(25, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(26, ii) = Format(Val(.TextMatrix(25, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(26, ii) = "0.00"
               End If
            End If
        Next ii
        '第四週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(27, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(7).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第四週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(26, ii)) <> 0 Then
                  .TextMatrix(28, ii) = Format(.TextMatrix(27, ii) / Val(.TextMatrix(26, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(28, ii) = "0.00"
               End If
            End If
        Next ii
        '第四週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(29, ii) = CalPoints(Val(Replace(.TextMatrix(28, ii), "%", "")) / 100)
'            End If
        Next ii
        '第四週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(30, ii) = Format(.TextMatrix(2, ii), "0.00")
            End If
        Next ii
        '第四週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(31, ii) = Format(CalFinish(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第四週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(30, ii)) <> 0 Then
                  .TextMatrix(32, ii) = Format(Val(.TextMatrix(31, ii)) / Val(.TextMatrix(30, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(32, ii) = "0.00"
               End If
            End If
        Next ii
        
        '本月得分平均
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(33, ii) = Format((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4, "0.00")
'            End If
        Next ii
'        '更新資料
'        For ii = 2 To intColCount + 2 - 1
'            UpdateMonthAssess "1", .TextMatrix(34, ii), Val(frm090624.Txt1(0).Text) + 191100, Val(.TextMatrix(33, ii))
'        Next ii
                    
        '員工編號-->員工姓名
        '排除IAIN(88024)的資料
        '93.12.13 MODIFY BY SONIA 排除外翻人員
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/04/12
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/08/22
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2006/05/01
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'modify by sonia 2014/4/9 加入94007林景郁總經理
        'Modified by Morgan 2022/10/31 +99050
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050') And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        '93.12.13 END
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        While Not rsA.EOF
            .TextMatrix(0, ii) = "" & rsA("ST02").Value
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End With
    
'********************************************************
    '繪圖人員速度考核
    With Me.grd(1)
        '員工編號
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='82' Or ST05='AC') Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        intColCount = 0
        While Not rsA.EOF
            .Cols = .Cols + 1
            .TextMatrix(0, ii) = "" & rsA("ST01").Value
            .TextMatrix(34, ii) = "" & rsA("ST01").Value
            .TextMatrix(1, ii) = "草墨合計"
            intColCount = intColCount + 1
            rsA.MoveNext
            ii = ii + 1
        Wend
        .Cols = .Cols - 1
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 2 To intColCount + 2 - 1
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strGoal = "0"
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE09,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='82' Or ST05='AC')) S Where S.ST01=PE01 And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strST01 = "" & rsA.Fields(0).Value
                strST03 = "" & rsA.Fields(1).Value
                strST06 = "" & rsA.Fields(2).Value
                strST13 = "" & rsA.Fields(3).Value
                strST16 = "" & rsA.Fields(4).Value
                strGoal = Val("" & rsA.Fields(5).Value) * 2
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            .TextMatrix(2, ii) = IIf(strGoal <> "0", Format(strGoal, "0.00"), "0")
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(3, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(4, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(5, ii) = Format(Val(.TextMatrix(4, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(5, ii) = "0.00"
               End If
            End If
        Next ii
        '第一週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(6, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(2).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第一週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(5, ii)) <> 0 Then
                  .TextMatrix(8, ii) = Format(Val(.TextMatrix(6, ii)) / Val(.TextMatrix(5, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(8, ii) = "0.00"
               End If
            End If
        Next ii
        '第一週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(7, ii) = CalPoints(Val(Replace(.TextMatrix(8, ii), "%", "")) / 100)
'            End If
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(9, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(10, ii) = Format(Val(.TextMatrix(9, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(10, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(11, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(3).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第二週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(10, ii)) <> 0 Then
                  .TextMatrix(12, ii) = Format(Val(.TextMatrix(11, ii)) / Val(.TextMatrix(10, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(12, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(13, ii) = CalPoints(Val(Replace(.TextMatrix(12, ii), "%", "")) / 100)
'            End If
        Next ii
        '第二週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))) <> 0 Then
                  .TextMatrix(14, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))), "0.00")
               Else
                  .TextMatrix(14, ii) = "0.00"
               End If
            End If
        Next ii
        '第二週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(15, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第二週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(14, ii)) <> 0 Then
                  .TextMatrix(16, ii) = Format(Val(.TextMatrix(15, ii)) / Val(.TextMatrix(14, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(16, ii) = "0.00"
               End If
            End If
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(17, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(18, ii) = Format(Val(.TextMatrix(17, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(18, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(19, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(5).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第三週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(18, ii)) <> 0 Then
                  .TextMatrix(20, ii) = Format(Val(.TextMatrix(19, ii)) / Val(.TextMatrix(18, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(20, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(21, ii) = CalPoints(Val(Replace(.TextMatrix(20, ii), "%", "")) / 100)
'            End If
        Next ii
        '第三週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))) <> 0 Then
                  .TextMatrix(22, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))), "0.00")
               Else
                  .TextMatrix(22, ii) = "0.00"
               End If
            End If
        Next ii
        '第三週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(23, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第三週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(22, ii)) <> 0 Then
                  .TextMatrix(24, ii) = Format(Val(.TextMatrix(23, ii)) / Val(.TextMatrix(22, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(24, ii) = "0.00"
               End If
            End If
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(25, ii) = "" & rsA.Fields(0).Value
'            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(26, ii) = Format(Val(.TextMatrix(25, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(26, ii) = "0.00"
               End If
            End If
        Next ii
        '第四週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(27, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(7).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第四週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(26, ii)) <> 0 Then
                  .TextMatrix(28, ii) = Format(Val(.TextMatrix(27, ii)) / Val(.TextMatrix(26, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(28, ii) = "0.00"
               End If
            End If
        Next ii
        '第四週得分
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(29, ii) = CalPoints(Val(Replace(.TextMatrix(28, ii), "%", "")) / 100)
'            End If
        Next ii
        '第四週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(30, ii) = Format(.TextMatrix(2, ii), "0.00")
            End If
        Next ii
        '第四週累計完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(31, ii) = Format(CalFinish1(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000), "0.00")
'            End If
        Next ii
        '第四週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               'add by nick 20050215 加入判斷
               If Val(.TextMatrix(30, ii)) <> 0 Then
                  .TextMatrix(32, ii) = Format(Val(.TextMatrix(31, ii)) / Val(.TextMatrix(30, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(32, ii) = "0.00"
               End If
            End If
        Next ii
        
        '本月得分平均
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(33, ii) = Format((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4, "0.00")
'            End If
        Next ii
'        '更新資料
'        For ii = 2 To intColCount + 2 - 1
'            UpdateMonthAssess "2", .TextMatrix(34, ii), Val(frm090624.Txt1(0).Text) + 191100, Val(.TextMatrix(33, ii))
'        Next ii
                    
        '員工編號-->員工姓名
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='82' Or ST05='AC') Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        While Not rsA.EOF
            .TextMatrix(0, ii) = "" & rsA("ST02").Value
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End With

End Sub

'新制  add by nickc 2005/02/25 copy process 改
Private Sub ProcessNew()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer
Dim intColCount As Integer
Dim strST01 As String, strST03 As String, strST06 As String, strST13 As String, strST16 As String, strNCFPGoal As String, strCFPGoal As String, strGoal As String
Dim strMA54 As String

         grd(0).Visible = False
         grd(1).Visible = False
    '承辦人速度考核
    With Me.grd(0)
        '員工編號
        'modify by sonia 2014/4/9 加入94007林景郁總經理
        'modified by Morgan 2022/10/31 +99050
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050') And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        intColCount = 0
        While Not rsA.EOF
            .Cols = .Cols + 1
            .TextMatrix(0, ii) = "" & rsA("ST01").Value
            .row = 0
            .col = ii
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(34, ii) = "" & rsA("ST01").Value
            .TextMatrix(1, ii) = "稿"
            intColCount = intColCount + 1
            rsA.MoveNext
            ii = ii + 1
        Wend
        .Cols = .Cols - 1
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 2 To intColCount + 2 - 1
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strNCFPGoal = "0": strCFPGoal = "0"
            '非CFP目標件數
            'edit by nickc 2005/04/12
            'strSQLA = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            'MODIFY BY SONIA 2014/4/11 PE02<>'CFP' 改為 PE02='P', 因為杜燕文有T的目標
            'Modified by Morgan 2022/10/31 +99050
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050')) S Where S.ST01=PE01 And PE02='P' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strST01 = "" & rsA.Fields(0).Value
                strST03 = "" & rsA.Fields(1).Value
                strST06 = "" & rsA.Fields(2).Value
                strST13 = "" & rsA.Fields(3).Value
                strST16 = "" & rsA.Fields(4).Value
                strNCFPGoal = "" & rsA.Fields(5).Value
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'CFP目標件數
            'edit by nickc 2005/04/12
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum((Nvl(PE05,0)+Nvl(PE07,0)) * 2) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum((Nvl(PE05,0)+Nvl(PE07,0)) * 2) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum((Nvl(PE05,0)+Nvl(PE07,0)) * 2) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            'Modified by Morgan 2022/10/31 +99050
            StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum((Nvl(PE05,0)+Nvl(PE07,0)) * 2) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007' Or ST01='99050')) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
                strST01 = "" & rsB.Fields(0).Value
                strST03 = "" & rsB.Fields(1).Value
                strST06 = "" & rsB.Fields(2).Value
                strST13 = "" & rsB.Fields(3).Value
                strST16 = "" & rsB.Fields(4).Value
                strCFPGoal = "" & rsB.Fields(5).Value
            End If
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            .TextMatrix(2, ii) = CalGoal(strST01, strST03, strST06, strST13, strST16, strNCFPGoal, strCFPGoal)
            DoEvents
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(3, ii) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
                .TextMatrix(4, ii) = "" & rsA.Fields(0).Value
                DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(5, ii) = Format(Val(.TextMatrix(4, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(5, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第一週完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
                .TextMatrix(6, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(2).Text, "00")) + 19110000, False), "0.00")
                DoEvents
        Next ii
        '第一週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(5, ii)) <> 0 Then
                  .TextMatrix(8, ii) = Format(Val(.TextMatrix(6, ii)) / Val(.TextMatrix(5, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(8, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第一週得分
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
           .TextMatrix(7, ii) = CalPoints(Val(Replace(.TextMatrix(8, ii), "%", "")) / 100)
           DoEvents
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(9, ii) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(10, ii) = Format(Val(.TextMatrix(9, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(10, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第二週完成
        For ii = 2 To intColCount + 2 - 1
'            '若當月有目標
'            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(11, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(3).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, False), "0.00")
                DoEvents
'            End If
        Next ii
        '第二週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(10, ii)) <> 0 Then
                  .TextMatrix(12, ii) = Format(Val(.TextMatrix(11, ii)) / Val(.TextMatrix(10, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(12, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第二週得分
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
                .TextMatrix(13, ii) = CalPoints(Val(Replace(.TextMatrix(12, ii), "%", "")) / 100)
                DoEvents
        Next ii
        '第二週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(14, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))), "0.00")
               Else
                  .TextMatrix(14, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第二週累計完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(15, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, False), "0.00")
            DoEvents
        Next ii
        '第二週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(14, ii)) <> 0 Then
                  .TextMatrix(16, ii) = Format(Val(.TextMatrix(15, ii)) / Val(.TextMatrix(14, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(16, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(17, ii) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(18, ii) = Format(Val(.TextMatrix(17, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(18, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(19, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(5).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, False), "0.00")
            DoEvents
        Next ii
        '第三週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(18, ii)) <> 0 Then
                  .TextMatrix(20, ii) = Format(Val(.TextMatrix(19, ii)) / Val(.TextMatrix(18, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(20, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週得分
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(21, ii) = CalPoints(Val(Replace(.TextMatrix(20, ii), "%", "")) / 100)
            DoEvents
        Next ii
        '第三週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))) <> 0 Then
                  .TextMatrix(22, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))), "0.00")
               Else
                  .TextMatrix(22, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週累計完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(23, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, False), "0.00")
            DoEvents
        Next ii
        '第三週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(22, ii)) <> 0 Then
                  .TextMatrix(24, ii) = Format(Val(.TextMatrix(23, ii)) / Val(.TextMatrix(22, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(24, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
           .TextMatrix(25, ii) = "" & rsA.Fields(0).Value
           DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(26, ii) = Format(Val(.TextMatrix(25, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(26, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第四週完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            .TextMatrix(27, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(7).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, False), "0.00")
            DoEvents
        Next ii
        '第四週達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(26, ii)) <> 0 Then
                  .TextMatrix(28, ii) = Format(.TextMatrix(27, ii) / Val(.TextMatrix(26, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(28, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第四週得分
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
           .TextMatrix(29, ii) = CalPoints(Val(Replace(.TextMatrix(28, ii), "%", "")) / 100)
           DoEvents
        Next ii
        '第四週累計目標
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(30, ii) = Format(.TextMatrix(2, ii), "0.00")
            End If
            DoEvents
        Next ii
        '第四週累計完成
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
           .TextMatrix(31, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, False, , strMA54), "0.00")
           .TextMatrix(35, ii) = strMA54 'Add by Morgan 2009/7/14
           DoEvents
        Next ii
        '第四週累計達成比例
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(30, ii)) <> 0 Then
                  .TextMatrix(32, ii) = Format(Val(.TextMatrix(31, ii)) / Val(.TextMatrix(30, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(32, ii) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '本月得分平均
        For ii = 2 To intColCount + 2 - 1
            '若當月有目標
            'Added by Morgan 2019/3/18 108考核(工程師本月得分改以整月達成率計算)
            If m_bol108Rule Then
               .TextMatrix(33, ii) = CalPoints(Val(Replace(.TextMatrix(32, ii), "%", "")) / 100)
            Else
            'end 2019/3/18
               .TextMatrix(33, ii) = Format((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4, "0.00")
            End If
            .row = 33
            .col = ii
            .CellAlignment = flexAlignCenterCenter
             DoEvents
        Next ii
'edit by nickc 2005/03/02  最後做
'        '員工編號-->員工姓名
'        strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
'        rsA.CursorLocation = adUseClient
'        rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'        ii = 2
'        While Not rsA.EOF
'            .TextMatrix(0, ii) = "" & rsA("ST02").Value
'            rsA.MoveNext
'            ii = ii + 1
'        Wend
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
    End With
    
'********************************************************
    '繪圖人員速度考核
    With Me.grd(1)
        '員工編號
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='82' Or ST05='AC') Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        intColCount = 0
        While Not rsA.EOF
            .Cols = .Cols + 1
            .TextMatrix(0, ii) = "" & rsA("ST01").Value
            .TextMatrix(34, ii) = "" & rsA("ST01").Value
            .row = 0
            .col = ii
            .CellAlignment = flexAlignCenterCenter
            .row = 1
            .col = ii
            .TextMatrix(1, ii) = "草"
            .CellAlignment = flexAlignCenterCenter
            .row = 1
            .col = ii + 1
            .CellAlignment = flexAlignCenterCenter
            .Cols = .Cols + 1
            .TextMatrix(0, ii + 1) = "" & rsA("ST01").Value
            .TextMatrix(34, ii + 1) = "" & rsA("ST01").Value
            .TextMatrix(1, ii + 1) = "墨"
            .row = 1
            .col = ii + 2
            .CellAlignment = flexAlignCenterCenter
            .Cols = .Cols + 1
            .TextMatrix(0, ii + 2) = "" & rsA("ST01").Value
            .TextMatrix(34, ii + 2) = ""
            .TextMatrix(1, ii + 2) = ""
            .ColWidth(ii + 2) = 0
            ii = ii + 3
            intColCount = intColCount + 3
            rsA.MoveNext
            DoEvents
        Wend
        .Cols = .Cols - 1
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 2 To intColCount + 2 - 1 Step 3
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strGoal = "0"
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE09,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='82' Or ST05='AC')) S Where S.ST01=PE01 And PE03=" & (Val(frm090624.txt1(0).Text) + 191100) & " And PE01='" & .TextMatrix(34, ii) & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strST01 = "" & rsA.Fields(0).Value
                strST03 = "" & rsA.Fields(1).Value
                strST06 = "" & rsA.Fields(2).Value
                strST13 = "" & rsA.Fields(3).Value
                strST16 = "" & rsA.Fields(4).Value
                strGoal = Val("" & rsA.Fields(5).Value)
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            .TextMatrix(2, ii) = IIf(strGoal <> "0", Format(strGoal, "0.00"), "0")
            .TextMatrix(2, ii + 1) = IIf(strGoal <> "0", Format(strGoal, "0.00"), "0")
            DoEvents
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1 Step 3
'            '若當月有目標
            .TextMatrix(3, ii) = "" & rsA.Fields(0).Value
            .TextMatrix(3, ii + 1) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(4, ii) = "" & rsA.Fields(0).Value
            .TextMatrix(4, ii + 1) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(5, ii) = Format(Val(.TextMatrix(4, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(5, ii) = "0.00"
               End If
            End If
            If Val(.TextMatrix(3, ii + 1)) <> 0 Then
               If Val(.TextMatrix(3, ii + 1)) <> 0 Then
                  .TextMatrix(5, ii + 1) = Format(Val(.TextMatrix(4, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * Val(.TextMatrix(2, ii + 1)), "0.00")
               Else
                  .TextMatrix(5, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第一週完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
           .TextMatrix(6, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(2).Text, "00")) + 19110000, True), "0.00")
           .TextMatrix(6, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(2).Text, "00")) + 19110000, True, False), "0.00")
           DoEvents
        Next ii
        '第一週達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(5, ii)) <> 0 Then
                  .TextMatrix(8, ii) = Format(Val(.TextMatrix(6, ii)) / Val(.TextMatrix(5, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(8, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(5, ii + 1)) <> 0 Then
                  .TextMatrix(8, ii + 1) = Format(Val(.TextMatrix(6, ii + 1)) / Val(.TextMatrix(5, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(8, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第一週得分
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(7, ii) = CalPoints(Val(Replace(.TextMatrix(8, ii), "%", "")) / 100)
            .TextMatrix(7, ii + 1) = CalPoints(Val(Replace(.TextMatrix(8, ii + 1), "%", "")) / 100)
            DoEvents
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
           .TextMatrix(9, ii) = "" & rsA.Fields(0).Value
           .TextMatrix(9, ii + 1) = "" & rsA.Fields(0).Value
           DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(10, ii) = Format(Val(.TextMatrix(9, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(10, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(3, ii + 1)) <> 0 Then
                  .TextMatrix(10, ii + 1) = Format(Val(.TextMatrix(9, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * Val(.TextMatrix(2, ii + 1)), "0.00")
               Else
                  .TextMatrix(10, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第二週完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(11, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(3).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, True), "0.00")
            .TextMatrix(11, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(3).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, True, False), "0.00")
            DoEvents
        Next ii
        '第二週達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(10, ii)) <> 0 Then
                  .TextMatrix(12, ii) = Format(Val(.TextMatrix(11, ii)) / Val(.TextMatrix(10, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(12, ii) = "0.00%"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(10, ii + 1)) <> 0 Then
                  .TextMatrix(12, ii + 1) = Format(Val(.TextMatrix(11, ii + 1)) / Val(.TextMatrix(10, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(12, ii + 1) = "0.00%"
               End If
            End If
            DoEvents
        Next ii
        '第二週得分
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(13, ii) = CalPoints(Val(Replace(.TextMatrix(12, ii), "%", "")) / 100)
            .TextMatrix(13, ii + 1) = CalPoints(Val(Replace(.TextMatrix(12, ii + 1), "%", "")) / 100)
            DoEvents
        Next ii
        '第二週累計目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))) <> 0 Then
                  .TextMatrix(14, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii))), "0.00")
               Else
                  .TextMatrix(14, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(3, ii + 1)) * (Val(.TextMatrix(4, ii + 1)) + Val(.TextMatrix(9, ii + 1))) <> 0 Then
                  .TextMatrix(14, ii + 1) = Format(Val(.TextMatrix(2, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * (Val(.TextMatrix(4, ii + 1)) + Val(.TextMatrix(9, ii + 1))), "0.00")
               Else
                  .TextMatrix(14, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第二週累計完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
           .TextMatrix(15, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, True), "0.00")
           .TextMatrix(15, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(4).Text, "00")) + 19110000, True, False), "0.00")
           DoEvents
        Next ii
        '第二週累計達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(14, ii)) <> 0 Then
                  .TextMatrix(16, ii) = Format(Val(.TextMatrix(15, ii)) / Val(.TextMatrix(14, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(16, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(14, ii + 1)) <> 0 Then
                  .TextMatrix(16, ii + 1) = Format(Val(.TextMatrix(15, ii + 1)) / Val(.TextMatrix(14, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(16, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(17, ii) = "" & rsA.Fields(0).Value
            .TextMatrix(17, ii + 1) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(18, ii) = Format(Val(.TextMatrix(17, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(18, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(3, ii + 1)) <> 0 Then
                  .TextMatrix(18, ii + 1) = Format(Val(.TextMatrix(17, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * Val(.TextMatrix(2, ii + 1)), "0.00")
               Else
                  .TextMatrix(18, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(19, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(5).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, True), "0.00")
            .TextMatrix(19, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(5).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, True, False), "0.00")
            DoEvents
        Next ii
        '第三週達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(18, ii)) <> 0 Then
                  .TextMatrix(20, ii) = Format(Val(.TextMatrix(19, ii)) / Val(.TextMatrix(18, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(20, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(18, ii + 1)) <> 0 Then
                  .TextMatrix(20, ii + 1) = Format(Val(.TextMatrix(19, ii + 1)) / Val(.TextMatrix(18, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(20, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週得分
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(21, ii) = CalPoints(Val(Replace(.TextMatrix(20, ii), "%", "")) / 100)
            .TextMatrix(21, ii + 1) = CalPoints(Val(Replace(.TextMatrix(20, ii + 1), "%", "")) / 100)
            DoEvents
        Next ii
        '第三週累計目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))) <> 0 Then
                  .TextMatrix(22, ii) = Format(Val(.TextMatrix(2, ii)) / Val(.TextMatrix(3, ii)) * (Val(.TextMatrix(4, ii)) + Val(.TextMatrix(9, ii)) + Val(.TextMatrix(17, ii))), "0.00")
               Else
                  .TextMatrix(22, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(3, ii + 1)) * (Val(.TextMatrix(4, ii + 1)) + Val(.TextMatrix(9, ii + 1)) + Val(.TextMatrix(17, ii + 1))) <> 0 Then
                  .TextMatrix(22, ii + 1) = Format(Val(.TextMatrix(2, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * (Val(.TextMatrix(4, ii + 1)) + Val(.TextMatrix(9, ii + 1)) + Val(.TextMatrix(17, ii + 1))), "0.00")
               Else
                  .TextMatrix(22, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第三週累計完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(23, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, True), "0.00")
            .TextMatrix(23, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(6).Text, "00")) + 19110000, True, False), "0.00")
            DoEvents
        Next ii
        '第三週累計達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(22, ii)) <> 0 Then
                  .TextMatrix(24, ii) = Format(Val(.TextMatrix(23, ii)) / Val(.TextMatrix(22, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(24, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(22, ii + 1)) <> 0 Then
                  .TextMatrix(24, ii + 1) = Format(Val(.TextMatrix(23, ii + 1)) / Val(.TextMatrix(22, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(24, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((frm090624.txt1(0).Text) + 191100) & Format(frm090624.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 2
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(25, ii) = "" & rsA.Fields(0).Value
            .TextMatrix(25, ii + 1) = "" & rsA.Fields(0).Value
            DoEvents
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(3, ii)) <> 0 Then
                  .TextMatrix(26, ii) = Format(Val(.TextMatrix(25, ii)) / Val(.TextMatrix(3, ii)) * Val(.TextMatrix(2, ii)), "0.00")
               Else
                  .TextMatrix(26, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(3, ii + 1)) <> 0 Then
                  .TextMatrix(26, ii + 1) = Format(Val(.TextMatrix(25, ii + 1)) / Val(.TextMatrix(3, ii + 1)) * Val(.TextMatrix(2, ii + 1)), "0.00")
               Else
                  .TextMatrix(26, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第四週完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(27, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(7).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, True), "0.00")
            .TextMatrix(27, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(7).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, True, False), "0.00")
        Next ii
        '第四週達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(26, ii)) <> 0 Then
                  .TextMatrix(28, ii) = Format(Val(.TextMatrix(27, ii)) / Val(.TextMatrix(26, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(28, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(26, ii + 1)) <> 0 Then
                  .TextMatrix(28, ii + 1) = Format(Val(.TextMatrix(27, ii + 1)) / Val(.TextMatrix(26, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(28, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        '第四週得分
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(29, ii) = CalPoints(Val(Replace(.TextMatrix(28, ii), "%", "")) / 100)
            .TextMatrix(29, ii + 1) = CalPoints(Val(Replace(.TextMatrix(28, ii + 1), "%", "")) / 100)
            DoEvents
        Next ii
        '第四週累計目標
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
                .TextMatrix(30, ii) = Format(.TextMatrix(2, ii), "0.00")
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
                .TextMatrix(30, ii + 1) = Format(.TextMatrix(2, ii + 1), "0.00")
            End If
            DoEvents
        Next ii
        '第四週累計完成
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            .TextMatrix(31, ii) = Format(CalFinishNew(.TextMatrix(34, ii), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, True), "0.00")
            .TextMatrix(31, ii + 1) = Format(CalFinishNew(.TextMatrix(34, ii + 1), Val(frm090624.txt1(0).Text & Format(frm090624.txt1(1).Text, "00")) + 19110000, Val(frm090624.txt1(0).Text & Format(frm090624.txt1(8).Text, "00")) + 19110000, True, False), "0.00")
            DoEvents
        Next ii
        '第四週累計達成比例
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標
            If .TextMatrix(2, ii) <> "0" Then
               If Val(.TextMatrix(30, ii)) <> 0 Then
                  .TextMatrix(32, ii) = Format(Val(.TextMatrix(31, ii)) / Val(.TextMatrix(30, ii)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(32, ii) = "0.00"
               End If
            End If
            If .TextMatrix(2, ii + 1) <> "0" Then
               If Val(.TextMatrix(30, ii + 1)) <> 0 Then
                  .TextMatrix(32, ii + 1) = Format(Val(.TextMatrix(31, ii + 1)) / Val(.TextMatrix(30, ii + 1)) * 100, "0.00") & "%"
               Else
                  .TextMatrix(32, ii + 1) = "0.00"
               End If
            End If
            DoEvents
        Next ii
        
        '本月得分平均
        For ii = 2 To intColCount + 2 - 1 Step 3
            '若當月有目標  繪圖平均為草 65 % 墨 35 %
            .TextMatrix(33, ii) = Format(((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4 * 0.65) + ((Val(.TextMatrix(7, ii + 1)) + Val(.TextMatrix(13, ii + 1)) + Val(.TextMatrix(21, ii + 1)) + Val(.TextMatrix(29, ii + 1))) / 4 * 0.35), "0.00")
            .TextMatrix(33, ii + 1) = Format(((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4 * 0.65) + ((Val(.TextMatrix(7, ii + 1)) + Val(.TextMatrix(13, ii + 1)) + Val(.TextMatrix(21, ii + 1)) + Val(.TextMatrix(29, ii + 1))) / 4 * 0.35), "0.00")
            .MergeRow(33) = True
            .MergeCol(ii) = True
            .MergeCol(ii + 1) = True
            .row = 33
            .col = ii
            .CellAlignment = flexAlignCenterCenter
            DoEvents
        Next ii
    End With
ChgGrdColor

'加入存檔
Dim GrdIndex As Integer
Dim GrdCol As Integer
Dim GrdRow As Integer
Dim IsUpdate As Boolean
Dim IsRoung As Boolean
For GrdIndex = 0 To 1
   IsRoung = False
    For GrdCol = 2 To grd(GrdIndex).Cols - 1
         DoEvents
         grd(GrdIndex).col = GrdCol
         grd(GrdIndex).row = 34
         IsUpdate = False
         If Trim(grd(GrdIndex).Text) <> "" Then
            strSql = "select * from monthassess where ma01='" & grd(GrdIndex).Text & "' and ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma03='" & Trim(str(GrdIndex + 1)) & "' and ma36='"
            If GrdIndex = 0 Then
                strSql = strSql & "0' "
            Else
                grd(GrdIndex).row = 1
                If grd(GrdIndex).Text = "草" Then
                   strSql = strSql & "1' "
                   IsRoung = True
                Else
                   strSql = strSql & "2' "
                   IsRoung = False
                End If
            End If
            grd(GrdIndex).row = 34
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
               IsUpdate = True
            End If
            CheckOC3
            If IsUpdate = True Then
                 '速度資料
                 strSql = "update monthassess set "
                 'Modify by Morgan 2009/7/14
                 'For GrdRow = 2 To grd(GrdIndex).Rows - 2
                 For GrdRow = 2 To 33
                     grd(GrdIndex).row = GrdRow
                     DoEvents
                     strSql = strSql & " ma" & Format(GrdRow + 2, "00") & "=0" & Replace(grd(GrdIndex).Text, "%", "") & ","
                 Next GrdRow
                 '其餘資料
                 grd(GrdIndex).row = 34
                 If GrdIndex = 0 Then '承辦人
                     'Modify by Morgan 2009/7/14 +MA54
                     strSql = strSql & " ma54=" & Val(grd(GrdIndex).TextMatrix(35, GrdCol)) & ","
                     'Added by Morgan 2019/3/21 ma55發文實績點數(oPoint3)
                     StrSQLa = "select sum(nvl(oCount,0)),sum(nvl(oPoint,0)),sum(nvl(oCount2,0)),sum(nvl(oPoint2,0)),sum(nvl(oCount3,0)),sum(nvl(oCount4,0)),sum(nvl(oCount5,0)),sum(nvl(oPoint3,0))  from ("
                     'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                     'StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,sum(decode(CP26,'N',0,nvl(a0n03/1000,cp18-nvl(a1u07/1000,0)))) as oPoint3 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     'Modify by Morgan 2008/10/28 承辦件數也要加支援
                     'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
                     'StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2) as oCount,0 as oPoint,Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                     'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
                     StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) as oCount,0 as oPoint,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from supporthour,staff where st01(+)=sh02 and sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                     'end 2014/3/19
                     
                     If Not m_bol108Rule Then 'Added by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除)
                     
                        'Added by Morgan 2012/4/19 補收文點數
                        'Modified by Morgan 2014/3/19 2014/4/1 起非智權收文改每點折算0.04基數
                        'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount,0 as oPoint,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(" & Pt2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Pt2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                        'end 2014/3/19
                        'end 2012/4/19
                        'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
                        StrSQLa = StrSQLa & " Union All select mh06,mh07,mh12,Round(Nvl(MH05,0)*0.2 ,2) as oCount,0 as oPoint,Round(Nvl(MH05,0)*0.2 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ModifyHour where mh02='" & grd(GrdIndex) & "' and mh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and mh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and mh11='V' "
                        StrSQLa = StrSQLa & " Union All select eh06,eh07,eh12,Round(Nvl(EH05,0)*0.25 ,2) as oCount,0 as oPoint,Round(Nvl(EH05,0)*0.25 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ExtendHour where eh02='" & grd(GrdIndex) & "' and eh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and eh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and eh11='V' "
                        'end 2011/8/1
                        
                     End If 'Added by Morgan 2019/3/18
                     
                     'edit by nickc 2005/04/26
                     'strSQLA = strSQLA & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,cp18 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) "
                     'edit by nickc 2005/08/03 加快
                     'strSQLA = strSQLA & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),round(cp97 * cp98,2) "
                     'edit by nickc 2006/02/22 加入會稿加乘註記
                     'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),round(cp97 * cp98,2) "
                     'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                     'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount2,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,count(*) as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,engineerprogress where ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and ep09>cp48  group by cp01,cp02,cp09 "
                     StrSQLa = StrSQLa & " ) AA "
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     'Modified by Morgan 2019/3/21 +ma55發文實績點數
                     If AdoRecordSet3.RecordCount <> 0 Then
                           strSql = strSql & " ma37=0" & CheckStr(AdoRecordSet3.Fields(0).Value) & ", "
                           strSql = strSql & " ma40=0" & CheckStr(AdoRecordSet3.Fields(1).Value) & ", "
                           strSql = strSql & " ma43=0" & CheckStr(AdoRecordSet3.Fields(2).Value) & ", "
                           strSql = strSql & " ma50=0" & CheckStr(AdoRecordSet3.Fields(3).Value) & ", "
                           strSql = strSql & " ma51=0" & CheckStr(AdoRecordSet3.Fields(4).Value) & ", "
                           strSql = strSql & " ma52=0" & CheckStr(AdoRecordSet3.Fields(5).Value) & ", "
                           strSql = strSql & " ma47=0" & CheckStr(AdoRecordSet3.Fields(6).Value) & ", "
                           strSql = strSql & " ma55=0" & CheckStr(AdoRecordSet3.Fields(7).Value) & "  "
                     Else
                           strSql = strSql & " ma37=0,ma40=0,ma43=0,ma50=0,ma51=0,ma52=0,ma47=0,ma55=0 "
                     End If
                 Else   ' 繪圖
                     StrSQLa = "select sum(nvl(oCount,0)),sum(nvl(oPoint,0)),sum(nvl(oCount2,0)),sum(nvl(oPoint2,0)),0,sum(nvl(oCount4,0)),sum(nvl(oCount5,0)),sum(nvl(oPoint3,0))  from ("
                     If IsRoung = True Then
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep16,0) / 2 as oCount5 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep20 is null "
                        StrSQLa = StrSQLa & "select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep20,null,nvl(ep16,0) / 2,0) as oCount5,0 as oPoint3 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  "
                     Else
                        'edit by nickc 2005/04/26
                        'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,cp18 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null "
                        'edit by nickc 2005/08/03 加快
                        'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep29,null,nvl(ep19,0)/2,0) as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0)"
                        StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0) as oCount,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep29,null,nvl(ep19,0)/2,0) as oCount5,sum(decode(CP26,'N',0,nvl(a0n03/1000,cp18-nvl(a1u07/1000,0)))) as oPoint3 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0)"
                     End If
                     If IsRoung = True Then
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,nvl(ep16,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep15>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep15<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep20 is null "
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,decode(ep20,null,nvl(ep16,0)/2,0) as oCount4,0 as oCount5,0 as oPoint3 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep15>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep15<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  "
                     Else
                        '繪圖的支援只算在墨圖
                        'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
                        'StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(Nvl(SH05, 0)/4 ,2) as oCount,0 as oPoint,Round(Nvl(SH05, 0)/4 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                        StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(" & Sh2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Sh2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                        'end 2014/3/19
                        
                        If Not m_bol108Rule Then 'Added by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除)
                        
                           'Added by Morgan 2012/4/19 補收文點數
                           'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
                           'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount,0 as oPoint,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                           StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(" & Pt2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Pt2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                           'end 2014/3/19
                           'end 2012/4/19
                           
                           'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
                           StrSQLa = StrSQLa & " Union All select mh06,mh07,mh12,Round(Nvl(MH05,0)*0.2 ,2) as oCount,0 as oPoint,Round(Nvl(MH05,0)*0.2 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ModifyHour where mh02='" & grd(GrdIndex) & "' and mh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and mh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and mh11='V' "
                           StrSQLa = StrSQLa & " Union All select eh06,eh07,eh12,Round(Nvl(EH05,0)*0.25 ,2) as oCount,0 as oPoint,Round(Nvl(EH05,0)*0.25 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ExtendHour where eh02='" & grd(GrdIndex) & "' and eh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and eh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and eh11='V' "
                           'end 2011/8/1
                           
                        End If 'Added by Morgan 2019/3/18
                        
                        'edit by nickc 2005/04/26
                        'strSQLA = strSQLA & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,cp18 as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null "
                        'edit by nickc 2005/08/03 加快
                        'strSQLA = strSQLA & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2 "
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2 "
                        'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                        'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,decode(ep29,null,round(cp103 * cp104,2),0) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,decode(ep29,null,nvl(ep19,0)/2,0) as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0) "
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,decode(ep29,null,round(cp103 * cp104,2),0) as oCount2,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint2,0 as oCount3,decode(ep29,null,nvl(ep19,0)/2,0) as oCount4,0 as oCount5,0 as oPoint3 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0) "
                     End If
                     StrSQLa = StrSQLa & " ) AA "
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If AdoRecordSet3.RecordCount <> 0 Then
                           strSql = strSql & " ma37=0" & CheckStr(AdoRecordSet3.Fields(0).Value) & ", "
                           strSql = strSql & " ma40=0" & CheckStr(AdoRecordSet3.Fields(1).Value) & ", "
                           strSql = strSql & " ma43=0" & CheckStr(AdoRecordSet3.Fields(2).Value) & ", "
                           strSql = strSql & " ma50=0" & CheckStr(AdoRecordSet3.Fields(3).Value) & ",ma51=" & GetDelay(grd(GrdIndex), IsRoung) & ", "
                           strSql = strSql & " ma52=0" & CheckStr(AdoRecordSet3.Fields(5).Value) & ",  "
                           strSql = strSql & " ma47=0" & CheckStr(AdoRecordSet3.Fields(6).Value) & "  "
                     Else
                           strSql = strSql & " ma37=0,ma40=0,ma43=0,ma50=0,ma51=" & GetDelay(grd(GrdIndex), IsRoung) & ",ma52=0,ma47=0 "
                     End If
                 End If
                 strSql = strSql & " where ma01='" & grd(GrdIndex) & "' and ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma03='" & Trim(str(GrdIndex + 1)) & "' and ma36='"
                 If GrdIndex = 0 Then
                     strSql = strSql & "0' "
                 Else
                     grd(GrdIndex).row = 1
                     If grd(GrdIndex).Text = "草" Then
                        strSql = strSql & "1' "
                     Else
                        strSql = strSql & "2' "
                     End If
                 End If
            Else
                 strSql = "insert into monthassess (ma01,ma02,ma03,ma04,ma05,ma06,ma07,ma08,ma09,ma10,ma11,ma12,ma13,ma14,ma15,ma16,ma17,ma18,ma19,ma20,ma21,ma22,ma23,ma24,ma25,ma26,ma27,ma28,ma29,ma30,ma31,ma32,ma33,ma34,ma35,ma36,ma37,ma40,ma43,ma50,ma51,ma52,ma47,ma55) values ('" & grd(GrdIndex) & "'," & Trim(str(Val(lblMonth) + 191100)) & ",'" & Trim(str(GrdIndex + 1)) & "', "
                 'Modified by Morgan 2012/4/19
                 'For GrdRow = 2 To grd(GrdIndex).Rows - 2
                 For GrdRow = 2 To 33
                     grd(GrdIndex).row = GrdRow
                     DoEvents
                     strSql = strSql & " 0" & Replace(grd(GrdIndex).Text, "%", "") & ","
                 Next GrdRow
                 '其餘資料
                 grd(GrdIndex).row = 34
                 If GrdIndex = 0 Then '承辦人
                     'ma36
                     strSql = strSql & " '0',"
                     StrSQLa = "select sum(nvl(oCount,0)),sum(nvl(oPoint,0)),sum(nvl(oCount2,0)),sum(nvl(oPoint2,0)),sum(nvl(oCount3,0)),sum(nvl(oCount4,0)),sum(nvl(oCount5,0)),sum(nvl(oPoint3,0))  from ("
                     'edit by nickc 2005/04/26
                     'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp97 * cp98,2) as oCount,cp18 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 "
                     'edit by nickc 2005/08/03 加快
                     'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp97 * cp98,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,round(cp97 * cp98,2) "
                     'edit by nickc 2006/02/22 加入會稿加乘註記
                     'StrSQLa = StrSQLa & "select cp01,cp02,cp09,round(cp97 * cp98,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,round(cp97 * cp98,2) "
                     'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                     'StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,sum(decode(CP26,'N',0,nvl(a0n03/1000,cp18-nvl(a1u07/1000,0)))) as oPoint3 from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp09=a1u03(+) group by cp01,cp02,cp09,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     'Modify by Morgan 2008/10/28 承辦件數也要加支援
                     'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
                     'StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2) as oCount,0 as oPoint,Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                     'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
                     StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) as oCount,0 as oPoint,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from supporthour,staff where st01(+)=sh02 and sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                     'end 2014/3/19
                     
                     If Not m_bol108Rule Then 'Added by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除)
                        
                        'Added by Morgan 2012/4/19 補收文點數
                        'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
                        'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount,0 as oPoint,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(" & Pt2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Pt2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                        'end 2014/3/19
                        'end 2012/4/19
                        
                        'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
                        StrSQLa = StrSQLa & " Union All select mh06,mh07,mh12,Round(Nvl(MH05,0)*0.2 ,2) as oCount,0 as oPoint,Round(Nvl(MH05,0)*0.2 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ModifyHour where mh02='" & grd(GrdIndex) & "' and mh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and mh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and mh11='V' "
                        StrSQLa = StrSQLa & " Union All select eh06,eh07,eh12,Round(Nvl(EH05,0)*0.25 ,2) as oCount,0 as oPoint,Round(Nvl(EH05,0)*0.25 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ExtendHour where eh02='" & grd(GrdIndex) & "' and eh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and eh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and eh11='V' "
                        'end 2011/8/1
                        
                     End If 'Added by Morgan 2019/3/18
                     
                     'edit by nickc 2005/04/26
                     'strSQLA = strSQLA & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,cp18 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) "
                     'edit by nickc 2005/08/03 加快
                     'strSQLA = strSQLA & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),round(cp97 * cp98,2)"
                     'edit by nickc 2006/02/22 加入會稿加乘註記
                     'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,round(cp97 * cp98,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),round(cp97 * cp98,2)"
                     'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                     'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,round(Nvl(SCR02,0),2) as oCount,0 as oPoint,decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) as oCount2,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from specialcaserecord,engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep02=scr01(+) and ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and 'V'=scr03(+) and ep02=cp09(+) and ep02=a1u03(+) group by cp01,cp02,cp09,round(Nvl(SCR02,0),2),decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2)) "
                     StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,count(*) as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,engineerprogress where ep05='" & grd(GrdIndex) & "' and ep09>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep09<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and ep09>cp48  group by cp01,cp02,cp09 "
                     StrSQLa = StrSQLa & " ) AA "
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     'Modified by Morgan 2019/3/21 +ma55發文實績點數
                     If AdoRecordSet3.RecordCount <> 0 Then
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(0).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(1).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(2).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(3).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(4).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(5).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(6).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(7).Value) & " "
                     Else
                           strSql = strSql & " 0,0,0,0,0,0,0,0 "
                     End If
                 Else   ' 繪圖
                     grd(GrdIndex).row = 1
                     'ma36
                     strSql = strSql & " " & IIf(grd(GrdIndex).Text = "草", "'1'", "'2'") & ","
                     grd(GrdIndex).row = 34
                     StrSQLa = "select sum(nvl(oCount,0)),sum(nvl(oPoint,0)),sum(nvl(oCount2,0)),sum(nvl(oPoint2,0)),0,sum(nvl(oCount4,0)),sum(nvl(oCount5,0)),sum(nvl(oPoint3,0))  from ("
                     If IsRoung = True Then
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep16,0)/2 as oCount5 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep20 is null "
                        StrSQLa = StrSQLa & "select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep20,null,nvl(ep16,0)/2,0) as oCount5,0 as oPoint3 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)   "
                     Else
                        'edit by nickc 2005/04/26
                        'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,cp18 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null "
                        'edit by nickc 2005/08/03 加快
                        'strSQLA = strSQLA & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,round(cp103 * cp104,2) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,nvl(ep19,0)/2 as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                        'StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0) as oCount,sum(cp18-nvl(a1u07/1000,0)) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep29,null,nvl(ep19,0)/2,0) as oCount5 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0) "
                        StrSQLa = StrSQLa & "select cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0) as oCount,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,decode(ep29,null,nvl(ep19,0)/2,0) as oCount5,sum(decode(CP26,'N',0,nvl(a0n03/1000,cp18-nvl(a1u07/1000,0)))) as oPoint3 from caseprogress,engineerprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep13='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0) "
                     End If
                     If IsRoung = True Then
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,nvl(ep16,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep15>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep15<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  "
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,0 as oCount2,0 as oPoint2,0 as oCount3,decode(ep20,null,nvl(ep16,0)/2,0) as oCount4,0 as oCount5,0 as oPoint3 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep15>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep15<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  "
                     Else
                        'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
                        'StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(Round(Nvl(SH05, 0)/4 ,2) ,2) as oCount,0 as oPoint,Round(Nvl(SH05, 0)/4 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                        StrSQLa = StrSQLa & " Union All select sh06,sh07,sh12,Round(" & Sh2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Sh2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from supporthour where sh02='" & grd(GrdIndex) & "' and sh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and sh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and sh11='V' "
                        'end 2014/3/19
                        
                        If Not m_bol108Rule Then 'Added by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除)
                        
                           'Added by Morgan 2012/4/19 補收文點數
                           'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
                           'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount,0 as oPoint,Round(nvl(a0n03/1000,cp18)*0.05 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                           StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,Round(" & Pt2EPtCode & " ,2) as oCount,0 as oPoint,Round(" & Pt2EPtCode & " ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from caseprogress,acc0n0 where a0n02(+)=cp09 and cp13='" & grd(GrdIndex) & "' and cp05>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp05<=" & Trim(str(Val(lblMonth) + 191100)) & "31 And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' "
                           'end 2014/3/19
                           'end 2012/4/19
                        
                           'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
                           StrSQLa = StrSQLa & " Union All select mh06,mh07,mh12,Round(Nvl(MH05,0)*0.2 ,2) as oCount,0 as oPoint,Round(Nvl(MH05,0)*0.2 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ModifyHour where mh02='" & grd(GrdIndex) & "' and mh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and mh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and mh11='V' "
                           StrSQLa = StrSQLa & " Union All select eh06,eh07,eh12,Round(Nvl(EH05,0)*0.25 ,2) as oCount,0 as oPoint,Round(Nvl(EH05,0)*0.25 ,2) as oCount2,0 as oPoint2,0 as oCount3,0 as oCount4,0 as oCount5,0 as oPoint3 from ExtendHour where eh02='" & grd(GrdIndex) & "' and eh01>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and eh01<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and eh11='V' "
                           'end 2011/8/1
                           
                        End If 'Added by Morgan 2019/3/18
                        
                        'edit by nickc 2005/04/26
                        'strSQLA = strSQLA & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,cp18 as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null "
                        'edit by nickc 2005/08/03 加快
                        'strSQLA = strSQLA & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'edit by nickc 2006/01/19 點數不管計不計件
                        'StrSQLa = StrSQLa & " union select cp01,cp02,cp09,0 as oCount,0 as oPoint,round(cp103 * cp104,2) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,nvl(ep19,0)/2 as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+)  and ep29 is null and ep02=a1u03(+) group by cp01,cp02,cp09,round(cp103 * cp104,2),nvl(ep19,0)/2"
                        'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
                        'StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,decode(ep29,null,round(cp103 * cp104,2),0) as oCount2,sum(cp18-nvl(a1u07/1000,0)) as oPoint2,0 as oCount3,decode(ep29,null,nvl(ep19,0)/2,0) as oCount4,0 as oCount5 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE where ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910')  and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0)"
                        StrSQLa = StrSQLa & " Union All select cp01,cp02,cp09,0 as oCount,0 as oPoint,decode(ep29,null,round(cp103 * cp104,2),0) as oCount2,sum(nvl(a0n03/1000,cp18-nvl(a1u07/1000,0))) as oPoint2,0 as oCount3,decode(ep29,null,nvl(ep19,0)/2,0) as oCount4,0 as oCount5,0 as oPoint3 from engineerprogress,caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp14='" & grd(GrdIndex) & "' and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31) group by a1u03) ABCDE ,acc0n0 where a0n02(+)=cp09 and ep13='" & grd(GrdIndex) & "' and ep18>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and ep18<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and ep02=cp09(+) and (ep29 is null or ep20||ep29||cp10='NN910')  and ep02=a1u03(+) group by cp01,cp02,cp09,decode(ep29,null,round(cp103 * cp104,2),0),decode(ep29,null,nvl(ep19,0)/2,0)"
                     End If
                     StrSQLa = StrSQLa & " ) AA "
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     'Modified by Morgan 2019/3/21 +ma55發文實績點數
                     If AdoRecordSet3.RecordCount <> 0 Then
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(0).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(1).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(2).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(3).Value) & "," & GetDelay(grd(GrdIndex), IsRoung) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(5).Value) & ", "
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(6).Value) & ","
                           strSql = strSql & " 0" & CheckStr(AdoRecordSet3.Fields(7).Value) & " "
                     Else
                           strSql = strSql & " 0,0,0,0," & GetDelay(grd(GrdIndex), IsRoung) & ",0,0,0 "
                     End If
                  End If
                 strSql = strSql & " ) "
            End If
            cnnConnection.Execute strSql
            strSql = "'"
         End If
    Next GrdCol
   '本月達成率
   cnnConnection.Execute "update monthassess set ma38=decode(ma04,0,0,round((ma37/ma04) ,2)),ma44=decode(ma04,0,0,round((ma43/ma04),2)) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " "
    For GrdCol = 2 To grd(GrdIndex).Cols - 1
         grd(GrdIndex).col = GrdCol
         grd(GrdIndex).row = 0
         'MODIFY BY SONIA 2014/4/11 加入 pe02 in ('P','CFP') 杜燕文有T的目標
         'Modified by Morgan 2019/3/21 點數目標抓錯
         'StrSQLa = "select st02,sum(Nvl(decode(pe02,'CFP',pe05 * 2,pe05),0)+Nvl(decode(pe02,'CFP',pe07 * 2 ,PE07),0)) as T1,sum(Nvl(PE11,0)) as T2,sum(Nvl(PE10,0)) as T3  from staff,Performance where st01='" & grd(GrdIndex).Text & "' and " & Trim(str(Val(lblMonth) + 191100)) & "=pe03(+) and pe02 in ('P','CFP') AND st01=pe01(+) group by st02 "
         StrSQLa = "select st02,sum(nvl(pe06,0) + nvl(pe08,0)) as T1,sum(Nvl(PE11,0)) as T2,sum(Nvl(PE10,0)) as T3  from staff,Performance where st01='" & grd(GrdIndex).Text & "' and " & Trim(str(Val(lblMonth) + 191100)) & "=pe03(+) and pe02(+) in ('P','CFP') AND st01=pe01(+) group by st02 "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount <> 0 Then
               'Modified by Morgan 2019/3/21 +ma56發文實績點數達成率,修正原語法少了ma01所以過去都沒有資料(最後一筆是99999不繪圖沒有目標)
               cnnConnection.Execute "update monthassess set ma41=decode(" & Val(CheckStr(AdoRecordSet3.Fields("T1").Value)) & ",0,0,round((ma40/" & Val(CheckStr(AdoRecordSet3.Fields("T1").Value)) & ")*100,2)),ma56=decode(" & Val(CheckStr(AdoRecordSet3.Fields("T1").Value)) & ",0,0,round((ma55/" & Val(CheckStr(AdoRecordSet3.Fields("T1").Value)) & ")*100,2))  where ma01='" & grd(GrdIndex).Text & "' and ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma03='1' "
               cnnConnection.Execute "update monthassess set ma41=decode(" & Val(CheckStr(AdoRecordSet3.Fields("T2").Value)) & ",0,0,round((ma40/" & Val(CheckStr(AdoRecordSet3.Fields("T2").Value)) & ")*100,2)),ma48=decode(" & Val(CheckStr(AdoRecordSet3.Fields("T3").Value)) & ",0,0,round((ma47/" & Val(CheckStr(AdoRecordSet3.Fields("T3").Value)) & ")*100,2)),ma56=decode(" & Val(CheckStr(AdoRecordSet3.Fields("T2").Value)) & ",0,0,round((ma55/" & Val(CheckStr(AdoRecordSet3.Fields("T2").Value)) & ")*100,2))  where ma01='" & grd(GrdIndex).Text & "' and ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma03='2' "
              grd(GrdIndex).Text = CheckStr(AdoRecordSet3.Fields(0).Value)
         End If
   Next GrdCol
Next GrdIndex
'更新繪圖之修改及複雜時數
cnnConnection.Execute "update monthassess set ma53=(select sum(nvl(ep21,0)+nvl(ep22,0)+nvl(ep23,0)+nvl(ep24,0)+nvl(ep25,0)) from caseprogress,engineerprogress where cp09=ep02(+) and ep13=ma01 and cp27>=" & Trim(str(Val(lblMonth) + 191100)) & "01 and cp27<=" & Trim(str(Val(lblMonth) + 191100)) & "31 and cp01 in ('P','CFP','FCP') ) where ma03='2' and ma36='1' and ma02=" & Trim(str(Val(lblMonth) + 191100)) & " "

'算得分
SetPoint '程序太大改寫函數

CheckOC3

grd(0).Visible = True
grd(1).Visible = True
End Sub
'Added by Morgan 2019/3/18 程序太大改寫函數(超過64K)
'算得分
Private Sub SetPoint()
   Dim StrSQLa As String
   
   StrSQLa = "select  * from  assessrate where ar01 in (select max(ar01) from assessrate where ar01<=" & Val(Trim(str(Val(lblMonth) + 191100)) & "01") & ") "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2019/3/21 修正達成率少除以100問題(實際上月考核或季考核都會重新計算)
   If AdoRecordSet3.RecordCount <> 0 Then
      'Added by Morgan 2019/3/18 108考核(得分取消(達成率)^2的計算方式)
      If m_bol108Rule Then
         cnnConnection.Execute "update monthassess set ma39=round(ma38 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar09").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100))
         cnnConnection.Execute "update monthassess set ma42=round(ma41 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar10").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100))
         cnnConnection.Execute "update monthassess set ma45=round(ma44 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar11").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100))
         'Added by Morgan 2019/3/21 ma57發文實績點數得分
         cnnConnection.Execute "update monthassess set ma57=round(ma56 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar27").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100))
         '高於所佔比重的 150% 以 150 % 為上限只有點數
         cnnConnection.Execute "update monthassess set ma57=" & str(Val(CheckStr(AdoRecordSet3.Fields("ar27").Value)) * 1.5) & " where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma57>" & str(Val(CheckStr(AdoRecordSet3.Fields("ar27").Value)) * 1.5) & " "
      Else
      'end 2019/3/18
      
         '高於 100%
         cnnConnection.Execute "update monthassess set ma39=round(ma38 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar09").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma38>1 "
         cnnConnection.Execute "update monthassess set ma42=round(ma41 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar10").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma41>1 "
         cnnConnection.Execute "update monthassess set ma45=round(ma44 /100 * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar11").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma44>1 "
         '低於 100%
         cnnConnection.Execute "update monthassess set ma39=round(power(ma38 /100,2) * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar09").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma38<=1 "
         cnnConnection.Execute "update monthassess set ma42=round(power(ma41 /100,2) * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar10").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma41<=1 "
         
         cnnConnection.Execute "update monthassess set ma45=round(power(ma44 /100,2) * 0.8 * " & Val(CheckStr(AdoRecordSet3.Fields("ar11").Value)) & ",2) where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma44<=1 "
      End If 'Added by Morgan 2019/3/18
      
      '高於所佔比重的 150% 以 150 % 為上限只有點數
      cnnConnection.Execute "update monthassess set ma42=" & str(Val(CheckStr(AdoRecordSet3.Fields("ar10").Value)) * 1.5) & " where ma02=" & Trim(str(Val(lblMonth) + 191100)) & " and ma42>" & str(Val(CheckStr(AdoRecordSet3.Fields("ar10").Value)) * 1.5) & " "
   End If
End Sub

Private Sub DesignTitle(wksfrm090624 As Worksheet)
Dim ii As Integer

    With wksfrm090624
        .Range("A1") = Val(Right(frm090624.txt1(0).Text, 2)) & "月份"
        .Range("A1:B2").Select
        With .Range("A1:B2")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("A1:B2").Merge
        With .Range("A1:B2")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("A3") = "目標基數"
        .Range("A3:B3").Select
        With .Range("A3:B3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("A3:B3").Merge
        With .Range("A3:B3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("A4") = "本月工作天數"
        .Range("A4:B4").Select
        With .Range("A4:B4")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("A4:B4").Merge
        With .Range("A4:B4")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("A5") = "第一週"
        .Range("A5:A9").Select
        With .Range("A5:A9")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("A5:A9").Merge
        With .Range("A5:A9")
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
        
        For ii = 1 To 3
            .Range("A" & (10 + (ii - 1) * 8)) = IIf(ii = 1, "第二週", (IIf(ii = 2, "第三週", "第四週")))
            .Range("A" & (10 + (ii - 1) * 8) & ":A" & (10 + ii * 8 - 1)).Select
            With .Range("A" & (10 + (ii - 1) * 8) & ":A" & (10 + ii * 8 - 1))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("A" & (10 + (ii - 1) * 8) & ":A" & (10 + ii * 8 - 1)).Merge
            With .Range("A" & (10 + (ii - 1) * 8) & ":A" & (10 + ii * 8 - 1))
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
        Next ii

        .Range("B5") = "本週工作天數"
        .Range("B5:B5").Select
        With .Range("B5:B5")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("B5:B5").Merge
        With .Range("B5:B5")
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("B6") = "本週目標"
        .Range("B6:B6").Select
        With .Range("B6:B6")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("B6:B6").Merge
        With .Range("B6:B6")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("B7") = "本週完成"
        .Range("B7:B7").Select
        With .Range("B7:B7")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("B7:B7").Merge
        With .Range("B7:B7")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("B8") = "本週得分"
        .Range("B8:B8").Select
        With .Range("B8:B8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("B8:B8").Merge
        With .Range("B8:B8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
        .Range("B9") = "達成比例"
        .Range("B9:B9").Select
        With .Range("B9:B9")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("B9:B9").Merge
        With .Range("B9:B9")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
            
        For ii = 1 To 3
            .Range("B" & (10 + (ii - 1) * 8)) = "本週工作天數"
            .Range("B" & (10 + (ii - 1) * 8) & ":" & "B" & (10 + (ii - 1) * 8)).Select
            With .Range("B" & (10 + (ii - 1) * 8) & ":" & "B" & (10 + (ii - 1) * 8))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8) & ":" & "B" & (10 + (ii - 1) * 8)).Merge
            With .Range("B" & (10 + (ii - 1) * 8) & ":" & "B" & (10 + (ii - 1) * 8))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
        
            .Range("B" & (10 + (ii - 1) * 8 + 1)) = "本週目標"
            .Range("B" & (10 + (ii - 1) * 8 + 1) & ":" & "B" & (10 + (ii - 1) * 8) + 1).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 1) & ":" & "B" & (10 + (ii - 1) * 8 + 1))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 1) & ":" & "B" & (10 + (ii - 1) * 8) + 1).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 1) & ":" & "B" & (10 + (ii - 1) * 8 + 1))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
    
            .Range("B" & (10 + (ii - 1) * 8 + 2)) = "本週完成"
            .Range("B" & (10 + (ii - 1) * 8 + 2) & ":" & "B" & (10 + (ii - 1) * 8) + 2).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 2) & ":" & "B" & (10 + (ii - 1) * 8 + 2))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 2) & ":" & "B" & (10 + (ii - 1) * 8) + 2).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 2) & ":" & "B" & (10 + (ii - 1) * 8 + 2))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
    
            .Range("B" & (10 + (ii - 1) * 8 + 3)) = "本週達成比例"
            .Range("B" & (10 + (ii - 1) * 8 + 3) & ":" & "B" & (10 + (ii - 1) * 8) + 3).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 3) & ":" & "B" & (10 + (ii - 1) * 8 + 3))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 3) & ":" & "B" & (10 + (ii - 1) * 8) + 3).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 3) & ":" & "B" & (10 + (ii - 1) * 8 + 3))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With

            .Range("B" & (10 + (ii - 1) * 8 + 4)) = "本週得分"
            .Range("B" & (10 + (ii - 1) * 8 + 4) & ":" & "B" & (10 + (ii - 1) * 8) + 4).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 4) & ":" & "B" & (10 + (ii - 1) * 8 + 4))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 4) & ":" & "B" & (10 + (ii - 1) * 8) + 4).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 4) & ":" & "B" & (10 + (ii - 1) * 8 + 4))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With

            .Range("B" & (10 + (ii - 1) * 8 + 5)) = "累計目標"
            .Range("B" & (10 + (ii - 1) * 8 + 5) & ":" & "B" & (10 + (ii - 1) * 8) + 5).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 5) & ":" & "B" & (10 + (ii - 1) * 8 + 5))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 5) & ":" & "B" & (10 + (ii - 1) * 8) + 5).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 5) & ":" & "B" & (10 + (ii - 1) * 8 + 5))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With

            .Range("B" & (10 + (ii - 1) * 8 + 6)) = "累計完成"
            .Range("B" & (10 + (ii - 1) * 8 + 6) & ":" & "B" & (10 + (ii - 1) * 8) + 6).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 6) & ":" & "B" & (10 + (ii - 1) * 8 + 6))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 6) & ":" & "B" & (10 + (ii - 1) * 8) + 6).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 6) & ":" & "B" & (10 + (ii - 1) * 8 + 6))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With

            .Range("B" & (10 + (ii - 1) * 8 + 7)) = "累計達成比例"
            .Range("B" & (10 + (ii - 1) * 8 + 7) & ":" & "B" & (10 + (ii - 1) * 8) + 7).Select
            With .Range("B" & (10 + (ii - 1) * 8 + 7) & ":" & "B" & (10 + (ii - 1) * 8 + 7))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
            .Range("B" & (10 + (ii - 1) * 8 + 7) & ":" & "B" & (10 + (ii - 1) * 8) + 7).Merge
            With .Range("B" & (10 + (ii - 1) * 8 + 7) & ":" & "B" & (10 + (ii - 1) * 8 + 7))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
        Next ii
        
        .Range("A34") = "本月得分平均"
        .Range("A34").Select
        With .Range("A34")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Range("A34:B34").Merge
        With .Range("A34:B34")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
    
    End With
End Sub

'取得欄位座標名
Private Function GetCol(intCol As Integer) As String
Dim jj As Double
Dim kk As Integer

GetCol = ""
jj = intCol / 26
kk = intCol Mod 26
If jj <= 1 Then
    GetCol = ""
ElseIf jj <= 2 Then
    GetCol = "A"
ElseIf jj <= 3 Then
    GetCol = "B"
ElseIf jj <= 4 Then
    GetCol = "C"
ElseIf jj <= 5 Then
    GetCol = "D"
ElseIf jj <= 6 Then
    GetCol = "E"
ElseIf jj <= 7 Then
    GetCol = "F"
ElseIf jj <= 8 Then
    GetCol = "G"
ElseIf jj <= 9 Then
    GetCol = "H"
ElseIf jj <= 10 Then
    GetCol = "I"
ElseIf jj <= 11 Then
    GetCol = "J"
ElseIf jj <= 12 Then
    GetCol = "K"
ElseIf jj <= 13 Then
    GetCol = "L"
ElseIf jj <= 14 Then
    GetCol = "M"
ElseIf jj <= 15 Then
    GetCol = "N"
ElseIf jj <= 16 Then
    GetCol = "O"
ElseIf jj <= 17 Then
    GetCol = "P"
ElseIf jj <= 18 Then
    GetCol = "Q"
ElseIf jj <= 19 Then
    GetCol = "R"
ElseIf jj <= 20 Then
    GetCol = "S"
ElseIf jj <= 21 Then
    GetCol = "T"
ElseIf jj <= 22 Then
    GetCol = "U"
ElseIf jj <= 23 Then
    GetCol = "V"
ElseIf jj <= 24 Then
    GetCol = "W"
ElseIf jj <= 25 Then
    GetCol = "X"
ElseIf jj <= 26 Then
    GetCol = "Y"
ElseIf jj <= 27 Then
    GetCol = "Z"
End If
Select Case kk
Case "0": GetCol = GetCol & "Z"
Case "1": GetCol = GetCol & "A"
Case "2": GetCol = GetCol & "B"
Case "3": GetCol = GetCol & "C"
Case "4": GetCol = GetCol & "D"
Case "5": GetCol = GetCol & "E"
Case "6": GetCol = GetCol & "F"
Case "7": GetCol = GetCol & "G"
Case "8": GetCol = GetCol & "H"
Case "9": GetCol = GetCol & "I"
Case "10": GetCol = GetCol & "J"
Case "11": GetCol = GetCol & "K"
Case "12": GetCol = GetCol & "L"
Case "13": GetCol = GetCol & "M"
Case "14": GetCol = GetCol & "N"
Case "15": GetCol = GetCol & "O"
Case "16": GetCol = GetCol & "P"
Case "17": GetCol = GetCol & "Q"
Case "18": GetCol = GetCol & "R"
Case "19": GetCol = GetCol & "S"
Case "20": GetCol = GetCol & "T"
Case "21": GetCol = GetCol & "U"
Case "22": GetCol = GetCol & "V"
Case "23": GetCol = GetCol & "W"
Case "24": GetCol = GetCol & "X"
Case "25": GetCol = GetCol & "Y"
Case Else
End Select

End Function

Private Function CalGoal(strST01 As String, strST03 As String, strST06 As String, strST13 As String, strST16 As String, strNCFPGoal As String, strCFPGoal As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim dblMonth As Double '月份
If frm090624.txt1(9).Text = "1" Then
   CalGoal = Format(Val(strNCFPGoal) + Val(strCFPGoal), "0.00")
Else
   CalGoal = "0"
   '若無到職日, 則視做新人
   If strST13 = "" Then
       Select Case strST16
       Case "CFP"
           CalGoal = "0"
       Case Else
           CalGoal = "0"
       End Select
   '若有到職日
   Else
      
       'Modify by Morgan 2004/5/10
       '只有化學案的新工程師才要算特定比重；工作月份的基準日期用該月5號計算
       
   '    '若為林育輝
   '    If strST01 = "91013" Then
   '        dblMonth = 25
   '    Else
   '        dblMonth = DateDiff("m", ChangeWStringToWDateString(strST13), ChangeWStringToWDateString(Val(frm090624.txt1(0).Text & "01") + 19110000))
   '    End If
       
      If strST01 = "91021" Then '91021(畢君慧)
         'dt1=到職日；dt2=統計基準日
         Dim dt1 As Date, dt2 As Date
         dt1 = CDate(ChangeWStringToWDateString(strST13))
         dt2 = CDate(ChangeWStringToWDateString(Val(frm090624.txt1(0).Text & "05") + 19110000))
         dblMonth = DateDiff("m", dt1, dt2)
         If DateDiff("d", DateAdd("m", dblMonth, dt1), dt2) < 0 Then
            dblMonth = dblMonth - 1
         End If
      Else
         dblMonth = 25
      End If
       
       Select Case strST16
       Case "CFP"
           If dblMonth >= 4 And dblMonth <= 6 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           ElseIf dblMonth >= 7 And dblMonth <= 9 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           ElseIf dblMonth >= 10 And dblMonth <= 12 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           ElseIf dblMonth >= 13 And dblMonth <= 18 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           ElseIf dblMonth >= 19 And dblMonth <= 24 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           ElseIf dblMonth >= 25 Then
               CalGoal = Val(strNCFPGoal) / GetWeights(dblMonth) + Val(strCFPGoal)
           Else
               CalGoal = "0"
           End If
       Case Else
           
           If dblMonth >= 4 And dblMonth <= 6 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           ElseIf dblMonth >= 7 And dblMonth <= 9 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           ElseIf dblMonth >= 10 And dblMonth <= 12 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           ElseIf dblMonth >= 13 And dblMonth <= 18 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           ElseIf dblMonth >= 19 And dblMonth <= 24 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           ElseIf dblMonth >= 25 Then
               CalGoal = Val(strNCFPGoal) + Val(strCFPGoal) * GetWeights(dblMonth)
           Else
               CalGoal = "0"
           End If
       End Select
   End If
   CalGoal = IIf(CalGoal <> "0", Format(CalGoal, "0.00"), "0")
End If
End Function

'計算完成數(承辦人)
Private Function CalFinish(strST01 As String, strDateFrom As String, strDateTo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim dblNCFP As Double '非CFP件數
Dim dblCFP As Double 'CFP件數
Dim dblMonth As Double '在職月份

CalFinish = "0"
dblNCFP = 0: dblCFP = 0
StrSQLa = "Select * From EngineerProgress, CaseProgress Where EP02=CP09 And CP26 Is Null And EP05='" & strST01 & "' And EP09>=" & strDateFrom & " And EP09<=" & strDateTo
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
While Not rsA.EOF
    If rsA("CP01").Value = "CFP" Then
        dblCFP = dblCFP + 1
    Else
        dblNCFP = dblNCFP + 1
    End If
    rsA.MoveNext
Wend
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

    'Modify by Morgan 2004/5/10
    '只有化學案的新工程師才要算特定比重；工作月份的基準日期用該月5號計算
    
''若為林育輝
'If strST01 = "91013" Then
'    dblMonth = 25
'Else
'    dblMonth = DateDiff("m", ChangeWStringToWDateString(GetST13(strST01)), ChangeWStringToWDateString(Val(frm090624.txt1(0).Text & "01") + 19110000))
'End If

   If strST01 = "91021" Then '91021(畢君慧)
      'dt1=到職日；dt2=統計基準日
      Dim dt1 As Date, dt2 As Date
      dt1 = CDate(ChangeWStringToWDateString(GetST13(strST01)))
      dt2 = CDate(ChangeWStringToWDateString(Val(frm090624.txt1(0).Text & "05") + 19110000))
      dblMonth = DateDiff("m", dt1, dt2)
      If DateDiff("d", DateAdd("m", dblMonth, dt1), dt2) < 0 Then
         dblMonth = dblMonth - 1
      End If
   Else
      dblMonth = 25
   End If

'2012/1/12 modify by sonia 改公用模組
'If GetST16(strST01) = "CFP" Then
If PUB_GetStaffST16(strST01) = "CFP" Then
    CalFinish = dblNCFP / GetWeights(dblMonth) + dblCFP
Else
    CalFinish = dblNCFP + dblCFP * GetWeights(dblMonth)
End If
'Modify By Cheng 2004/03/01
'strSQLA = "Select Sum(Round(Nvl(SH05,0)/4,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
StrSQLa = "Select Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4) ,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' "
'End
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    CalFinish = Val(CalFinish) + Val("" & rsA.Fields(0).Value)
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
CalFinish = Format(CalFinish, "0.00")
End Function

'計算得分
Private Function CalPoints(strPercent As String) As String
    
strPercent = Format(Val("" & strPercent) * 100, "##0")
If Val(strPercent) >= 200 Then
    CalPoints = "40"
ElseIf Val(strPercent) >= 190 Then
    CalPoints = "39"
ElseIf Val(strPercent) >= 180 Then
    CalPoints = "38"
ElseIf Val(strPercent) >= 170 Then
    CalPoints = "37"
ElseIf Val(strPercent) >= 160 Then
    CalPoints = "36"
ElseIf Val(strPercent) >= 150 Then
    CalPoints = "35"
ElseIf Val(strPercent) >= 140 Then
    CalPoints = "34"
ElseIf Val(strPercent) >= 130 Then
    CalPoints = "33"
ElseIf Val(strPercent) >= 120 Then
    CalPoints = "32"
ElseIf Val(strPercent) >= 110 Then
    CalPoints = "31"
ElseIf Val(strPercent) >= 100 Then
    CalPoints = "30"
ElseIf Val(strPercent) >= 90 Then
    CalPoints = "26"
ElseIf Val(strPercent) >= 80 Then
    CalPoints = "22"
ElseIf Val(strPercent) >= 70 Then
    CalPoints = "18"
ElseIf Val(strPercent) >= 60 Then
    CalPoints = "14"
ElseIf Val(strPercent) >= 50 Then
    CalPoints = "10"
ElseIf Val(strPercent) >= 40 Then
    CalPoints = "6"
ElseIf Val(strPercent) >= 30 Then
    CalPoints = "2"
Else
    CalPoints = "0"
End If
    
End Function

'更新資料
Private Sub UpdateMonthAssess(strKind As String, strMA01 As String, strMonth As String, strPoints As String)
'strKind : 1為承辦人 , 2為繪圖人員
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    StrSQLa = "Select * From MonthAssess Where MA01='" & strMA01 & "' And MA02=" & Val(strMonth)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若有資料, 更新
    If rsA.RecordCount > 0 Then
        If strKind = "1" Then
            StrSQLa = "Update MonthAssess Set MA03=" & Val(strPoints) & " Where MA01='" & strMA01 & "' And MA02=" & Val(strMonth)
        Else
            StrSQLa = "Update MonthAssess Set MA04=" & Val(strPoints) & " Where MA01='" & strMA01 & "' And MA02=" & Val(strMonth)
        End If
        cnnConnection.Execute StrSQLa
    '若無資料, 新增
    Else
        If strKind = "1" Then
            StrSQLa = "Insert Into MonthAssess(MA01, MA02, MA03) Values('" & strMA01 & "'," & Val(strMonth) & "," & Val(strPoints) & " )"
        Else
            StrSQLa = "Insert Into MonthAssess(MA01, MA02, MA04) Values('" & strMA01 & "'," & Val(strMonth) & "," & Val(strPoints) & " )"
        End If
        cnnConnection.Execute StrSQLa
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing

End Sub

'計算完成數(繪圖)
Private Function CalFinish1(strST01 As String, strDateFrom As String, strDateTo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'If strST01 = "87025" Then
'    Debug.Print
'End If
CalFinish1 = "0"
'計算草圖件數
StrSQLa = "Select Count(*) From EngineerProgress, CaseProgress Where EP02=CP09 And EP20 Is Null And EP13='" & strST01 & "' And EP15>=" & strDateFrom & " And EP15<=" & strDateTo
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    CalFinish1 = Val(CalFinish1) + rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'計算墨圖件數
StrSQLa = "Select Count(*) From EngineerProgress, CaseProgress Where EP02=CP09 And EP13='" & strST01 & "' And EP18>=" & strDateFrom & " And EP18<=" & strDateTo
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    CalFinish1 = Val(CalFinish1) + rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'Modify By Cheng 2004/03/01
'edit by nick 2004/07/05 邱小姐說 2小時算一件，因為之前草墨圖一起算了
'Modify by Morgan 2004/9/2 等級=79 的不算繪圖點數
'strSQLA = "Select Sum(Round(Nvl(SH05,0)/2,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' " 'End
StrSQLa = "Select Sum(Round(Nvl(SH05,0)/2,2)) From SupportHour,staff Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' AND ST01(+)=SH02 AND ST05<>'79' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify By Cheng 2004/03/25
    'edit by nick 2004/07/05
    'CalFinish1 = Val(CalFinish1) + Val("" & rsA.Fields(0).Value) * 2
    CalFinish1 = Val(CalFinish1) + Val("" & rsA.Fields(0).Value)
    'End
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
CalFinish1 = Format(CalFinish1, "0.00")
End Function

'設定版面
Private Sub DesignWS_Format(wksfrm090624 As Worksheet, intLastRow As Integer, intLastCol As Integer)
Dim ii As Integer

    With wksfrm090624
    
        .Range("A1:" & GetCol(intLastCol) & intLastRow).Select
        .Range("A1:" & GetCol(intLastCol) & intLastRow).Columns.AutoFit
        With .Range("A1:" & GetCol(intLastCol) & intLastRow)
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.Name = "標楷體"
            .Font.Name = "TimesNewRoman"
        End With
        '畫線
        .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A1:" & GetCol(intLastCol) & intLastRow).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        '設定顏色
        '第一週目標(橘)
        .Range("C6:" & GetCol(intLastCol) & "6").Select
        .Range("C6:" & GetCol(intLastCol) & "6").Font.ColorIndex = 46
        '第二週目標(橘)
        .Range("C11:" & GetCol(intLastCol) & "11").Select
        .Range("C11:" & GetCol(intLastCol) & "11").Font.ColorIndex = 46
        '第三週目標(橘)
        .Range("C19:" & GetCol(intLastCol) & "19").Select
        .Range("C19:" & GetCol(intLastCol) & "19").Font.ColorIndex = 46
        '第四週目標(橘)
        .Range("C27:" & GetCol(intLastCol) & "27").Select
        .Range("C27:" & GetCol(intLastCol) & "27").Font.ColorIndex = 46
        '第一週本週完成底色(藍)
        .Range("C7:" & GetCol(intLastCol) & "7").Select
        With .Range("C7:" & GetCol(intLastCol) & "7").Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With
        '第一週得分底色(藍)
        .Range("C8:" & GetCol(intLastCol) & "8").Select
        With .Range("C8:" & GetCol(intLastCol) & "8").Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With

        '第二週累計達成比例(藍)
        .Range("C16:" & GetCol(intLastCol) & "16").Select
        .Range("C16:" & GetCol(intLastCol) & "16").Font.ColorIndex = 5
        '第二週累計完成底色(藍)
        .Range("C16:" & GetCol(intLastCol) & "16").Select
        With .Range("C16:" & GetCol(intLastCol) & "16").Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With
        '第三週累計達成比例(藍)
        .Range("C24:" & GetCol(intLastCol) & "24").Select
        .Range("C24:" & GetCol(intLastCol) & "24").Font.ColorIndex = 5
        '第三週累計完成底色(藍)
        .Range("C24:" & GetCol(intLastCol) & "24").Select
        With .Range("C24:" & GetCol(intLastCol) & "24").Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With
        '第四週累計達成比例(藍)
        .Range("C32:" & GetCol(intLastCol) & "32").Select
        .Range("C32:" & GetCol(intLastCol) & "32").Font.ColorIndex = 5
        '第四週累計完成底色(藍)
        .Range("C32:" & GetCol(intLastCol) & "32").Select
        With .Range("C32:" & GetCol(intLastCol) & "32").Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With
        
        '第一週達成比例底色(黃)
        .Range("C9:" & GetCol(intLastCol) & "9").Select
        With .Range("C9:" & GetCol(intLastCol) & "9").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第二週達成比例底色(黃)
        .Range("C13:" & GetCol(intLastCol) & "13").Select
        With .Range("C13:" & GetCol(intLastCol) & "13").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第二週得分底色(黃)
        .Range("C14:" & GetCol(intLastCol) & "14").Select
        With .Range("C14:" & GetCol(intLastCol) & "14").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第三週達成比例底色(黃)
        .Range("C21:" & GetCol(intLastCol) & "21").Select
        With .Range("C21:" & GetCol(intLastCol) & "21").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第三週得分底色(黃)
        .Range("C22:" & GetCol(intLastCol) & "22").Select
        With .Range("C22:" & GetCol(intLastCol) & "22").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第四週達成比例底色(黃)
        .Range("C29:" & GetCol(intLastCol) & "29").Select
        With .Range("C29:" & GetCol(intLastCol) & "29").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        '第四週得分底色(黃)
        .Range("C30:" & GetCol(intLastCol) & "30").Select
        With .Range("C30:" & GetCol(intLastCol) & "30").Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With

        '第二週累計達成比例(藍)
        .Range("C17:" & GetCol(intLastCol) & "17").Select
        .Range("C17:" & GetCol(intLastCol) & "17").Font.ColorIndex = 5
        '第二週累計達成比例底色(橘)
        .Range("C17:" & GetCol(intLastCol) & "17").Select
        With .Range("C17:" & GetCol(intLastCol) & "17").Interior
            .ColorIndex = 45
            .Pattern = xlSolid
        End With
        '第三週累計達成比例(藍)
        .Range("C25:" & GetCol(intLastCol) & "25").Select
        .Range("C25:" & GetCol(intLastCol) & "25").Font.ColorIndex = 5
        '第三週累計達成比例底色(橘)
        .Range("C25:" & GetCol(intLastCol) & "25").Select
        With .Range("C25:" & GetCol(intLastCol) & "25").Interior
            .ColorIndex = 45
            .Pattern = xlSolid
        End With
        '第四週累計達成比例(藍)
        .Range("C33:" & GetCol(intLastCol) & "33").Select
        .Range("C33:" & GetCol(intLastCol) & "33").Font.ColorIndex = 5
        '第四週累計達成比例底色(橘)
        .Range("C33:" & GetCol(intLastCol) & "33").Select
        With .Range("C33:" & GetCol(intLastCol) & "33").Interior
            .ColorIndex = 45
            .Pattern = xlSolid
        End With

    End With
End Sub

'取得在職月份的比重
Private Function GetWeights(dblMonth As Double) As Double

'Modify by Morgan 2004/5/10
'If dblMonth >= 4 And dblMonth <= 6 Then
If dblMonth <= 6 Then
    GetWeights = 4
ElseIf dblMonth >= 7 And dblMonth <= 9 Then
    GetWeights = 3.5
ElseIf dblMonth >= 10 And dblMonth <= 12 Then
    GetWeights = 3
ElseIf dblMonth >= 13 And dblMonth <= 18 Then
    GetWeights = 2.5
ElseIf dblMonth >= 19 And dblMonth <= 24 Then
    GetWeights = 2.25
ElseIf dblMonth >= 25 Then
    GetWeights = 2
'避免程式錯誤
Else
    GetWeights = 1
End If

End Function

Private Function GetST13(strST01 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select ST13 From Staff Where ST01='" & strST01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetST13 = "" & rsA.Fields(0).Value
Else
    GetST13 = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function GetST16(strST01 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select ST16 From Staff Where ST01='" & strST01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetST16 = "" & rsA.Fields(0).Value
Else
    GetST16 = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub grd_DblClick(Index As Integer)
    '只可修改本週完成
    With Me.grd(Index)
        If .TextMatrix(0, 2) = "" Then Exit Sub
        Select Case .col
        Case 0, 1
            Exit Sub
        Case Else
        End Select
        Select Case .row
        Case 0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31, 32, 33
            Exit Sub
        Case Else
        End Select
        m_dblRow = .row
        m_dblCol = .col
        Me.Text1.Move .CellLeft + 60, .CellTop + Me.stb.Top + .Top - 30, .CellWidth, .CellHeight
        Me.Text1.Text = .TextMatrix(.row, .col)
        Me.Text1.Visible = True
        Me.Text1.SetFocus
    End With
End Sub

Private Sub grd_Scroll(Index As Integer)
    Me.Text1.Visible = False
End Sub

Private Sub Text1_GotFocus()
    TextInverse Me.Text1
End Sub

Private Sub Text1_LostFocus()
    If IsNumeric(Me.Text1.Text) = True Then
        Me.grd(Me.stb.Tab).TextMatrix(m_dblRow, m_dblCol) = Format(Val(Me.Text1.Text), "0.00")
        '重新計算
        ReCompute m_dblRow, m_dblCol
    End If
    Me.Text1.Visible = False
End Sub

Private Sub ReCompute(dblRow As Double, dblCol As Double)
Dim ii As Integer

    With Me.grd(Me.stb.Tab)
        ii = dblCol
        '第一週達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(8, ii) = Format(Val(.TextMatrix(6, ii)) / Val(.TextMatrix(5, ii)) * 100, "0.00") & "%"
        End If
        '第一週得分
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(7, ii) = CalPoints(Val(Replace(.TextMatrix(8, ii), "%", "")) / 100)
        End If
        '第二週達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(12, ii) = Format(Val(.TextMatrix(11, ii)) / Val(.TextMatrix(10, ii)) * 100, "0.00") & "%"
        End If
        '第二週得分
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(13, ii) = CalPoints(Val(Replace(.TextMatrix(12, ii), "%", "")) / 100)
        End If
        '第二週累計完成
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(15, ii) = Format(Val(.TextMatrix(6, ii)) + Val(.TextMatrix(11, ii)), "0.00")
        End If
        '第二週累計達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(16, ii) = Format(Val(.TextMatrix(15, ii)) / Val(.TextMatrix(14, ii)) * 100, "0.00") & "%"
        End If
        '第三週達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(20, ii) = Format(Val(.TextMatrix(19, ii)) / Val(.TextMatrix(18, ii)) * 100, "0.00") & "%"
        End If
        '第三週得分
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(21, ii) = CalPoints(Val(Replace(.TextMatrix(20, ii), "%", "")) / 100)
        End If
        '第三週累計完成
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(23, ii) = Format(Val(.TextMatrix(6, ii)) + Val(.TextMatrix(11, ii)) + Val(.TextMatrix(19, ii)), "0.00")
        End If
        '第三週累計達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(24, ii) = Format(Val(.TextMatrix(23, ii)) / Val(.TextMatrix(22, ii)) * 100, "0.00") & "%"
        End If
        '第四週達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(28, ii) = Format(.TextMatrix(27, ii) / .TextMatrix(26, ii) * 100, "0.00") & "%"
        End If
        '第四週得分
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(29, ii) = CalPoints(Val(Replace(.TextMatrix(28, ii), "%", "")) / 100)
        End If
        '第四週累計完成
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(31, ii) = Format(Val(.TextMatrix(6, ii)) + Val(.TextMatrix(11, ii)) + Val(.TextMatrix(19, ii)) + Val(.TextMatrix(27, ii)), "0.00")
        End If
        '第四週累計達成比例
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(32, ii) = Format(Val(.TextMatrix(31, ii)) / Val(.TextMatrix(30, ii)) * 100, "0.00") & "%"
        End If
        '本月得分平均
        '若當月有目標
        If .TextMatrix(2, ii) <> "0" Then
            .TextMatrix(33, ii) = Format((Val(.TextMatrix(7, ii)) + Val(.TextMatrix(13, ii)) + Val(.TextMatrix(21, ii)) + Val(.TextMatrix(29, ii))) / 4, "0.00")
        End If
    End With

End Sub


'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
Dim xlsSalesPoint As New Excel.Application
Dim wksfrm090624_1 As New Worksheet
Dim wksfrm090624_2 As New Worksheet
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer
Dim jj As Integer
Dim intColCount As Integer
Dim strST01 As String, strST03 As String, strST06 As String, strST13 As String, strST16 As String, strNCFPGoal As String, strCFPGoal As String, strGoal As String
Dim strFileName As String 'Added by Lydia 2018/02/12
On Error GoTo ErrorHandler

    'Modified by Lydia 2018/02/12 改放在桌面
'    If Dir("D:\專利速度考核" & Trim(frm090624.txt1(0).Text) & ".xls") <> MsgText(601) Then
'        Kill "D:\專利速度考核" & Trim(frm090624.txt1(0).Text) & ".xls"
'    End If
    strFileName = PUB_Getdesktop & "\專利速度考核" & Trim(frm090624.txt1(0).Text) & ".xls"
    If Dir(strFileName) <> MsgText(601) Then
         Kill strFileName
    End If
    'end 2018/02/12
    '承辦人速度考核
    xlsSalesPoint.SheetsInNewWorkbook = 2 'Added by Lydia 2019/03/13 預設工作表數量
    xlsSalesPoint.Workbooks.add
    'Modified by Lydia 2018/02/12
    'Set wksfrm090624_1 = xlsSalesPoint.Sheets("Sheet1")
    Set wksfrm090624_1 = xlsSalesPoint.Sheets(1)
    With wksfrm090624_1
        .Activate
        xlsSalesPoint.ActiveWindow.Zoom = 75
        .Name = "承辦人"
        DesignTitle wksfrm090624_1
        For ii = 2 To Me.grd(0).Cols - 1
            For jj = 0 To Me.grd(0).Rows - 1 - 1
                .Range(GetCol(ii + 1) & jj + 1) = Me.grd(0).TextMatrix(jj, ii)
            Next jj
        Next ii
        .Range("A1").Select
        DesignWS_Format wksfrm090624_1, 34, ii
        .Range("A1").Select
    End With
    
    'Remove by Lydia 2018/02/12
   ' xlsSalesPoint.Workbooks(1).SaveAs FileName:="D:\專利速度考核" & frm090624.txt1(0).Text & ".xls"
   ' xlsSalesPoint.Workbooks.Open "D:\專利速度考核" & Trim(frm090624.txt1(0).Text) & ".xls"
   'end 2018/02/12
'********************************************************
    '繪圖人員速度考核
    'Modified by Lydia 2018/02/12
    'Set wksfrm090624_2 = xlsSalesPoint.Sheets("Sheet2")
    Set wksfrm090624_2 = xlsSalesPoint.Sheets(2)
    With wksfrm090624_2
        .Activate
        xlsSalesPoint.ActiveWindow.Zoom = 75
        .Name = "繪圖人員"
        DesignTitle wksfrm090624_2
        For ii = 2 To Me.grd(1).Cols - 1
            For jj = 0 To Me.grd(1).Rows - 1 - 1
                .Range(GetCol(ii + 1) & jj + 1) = Me.grd(1).TextMatrix(jj, ii)
            Next jj
        Next ii
        .Range("A1").Select
        DesignWS_Format wksfrm090624_2, 34, ii
        .Range("A1").Select
    End With
    
    Set wksfrm090624_1 = xlsSalesPoint.Sheets("承辦人")
    wksfrm090624_1.Activate
    'Modified by Lydia 2018/02/12
    'xlsSalesPoint.Workbooks(1).Save: DoEvents
   '判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   'end 2018/02/12
    xlsSalesPoint.Workbooks.Close: DoEvents
    xlsSalesPoint.Quit: DoEvents
    Set xlsSalesPoint = Nothing: DoEvents
    'Modifie by Lydia 2018/02/12
    'MsgBox "Excel檔案產生完成!!!", vbExclamation + vbOKOnly
    MsgBox "Excel檔案已產生在使用者桌面!!!" & vbCrLf & strFileName, vbExclamation + vbOKOnly
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub ChgGrdColor()
Dim ii As Integer
Dim jj As Integer
    
    '若有產生資料
    If Me.grd(Me.stb.Tab).TextMatrix(0, 2) <> "" Then
        For ii = 0 To Me.grd.Count - 1
            For jj = 2 To Me.grd(ii).Cols - 1
                Me.grd(ii).col = jj
                Me.grd(ii).row = 6
                Me.grd(ii).CellBackColor = &H80FF&
                Me.grd(ii).col = jj
                Me.grd(ii).row = 11
                Me.grd(ii).CellBackColor = &H80FF&
                Me.grd(ii).col = jj
                Me.grd(ii).row = 19
                Me.grd(ii).CellBackColor = &H80FF&
                Me.grd(ii).col = jj
                Me.grd(ii).row = 27
                Me.grd(ii).CellBackColor = &H80FF&
            Next jj
        Next ii
    End If
End Sub

'add by nickc 2005/03/01  新制算件數
'oIsDraw 是否繪圖的
'Modify by Morgan 2009/7/14 +strMA54:累計會稿量
Private Function CalFinishNew(strST01 As String, strDateFrom As String, strDateTo As String, oIsDraw As Boolean, Optional oIsRough As Boolean = True, Optional strMA54 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim dblNCFP As Double '非CFP件數
Dim dblCFP As Double 'CFP件數
Dim dblMonth As Double '在職月份
Dim dblDeMA54 As Double '累計未會稿件數

'Add by Morgan 2010/3/30 累計會稿量
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset

dblDeMA54 = 0

DoEvents
CalFinishNew = "0"
dblNCFP = 0: dblCFP = 0
StrSQLa = "Select * From EngineerProgress, CaseProgress Where EP02=CP09 "
If oIsDraw = True Then
   If oIsRough = True Then
      StrSQLa = StrSQLa & " and ep20 is null and EP13='" & strST01 & "' And EP15>=" & strDateFrom & " And EP15<=" & strDateTo
   Else
      StrSQLa = StrSQLa & " and ep29 is null and EP13='" & strST01 & "' And EP18>=" & strDateFrom & " And EP18<=" & strDateTo
   End If
Else
   StrSQLa = StrSQLa & " and EP05='" & strST01 & "' And EP09>=" & strDateFrom & " And EP09<=" & strDateTo
   
   'Add by Morgan 2010/3/30 累計會稿量統計條件改只考慮會稿日區間,不管完稿日區間
   'Modify by Morgan 2010/9/27 +智權人員為工程師的以收文點數*0.05換算為基數
   'Modify by Morgan 2010/11/4 +智權人員收文不算(74018杜燕文)
   'If bolNewPromoterRule Then 'Removed by Morgan 2014/3/20 早已實施不用再判斷
      StrSqlB = "select nvl(sum(pp),0) from (select sum(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)) pp from engineerprogress,caseprogress where EP05='" & strST01 & "' And EP07>=" & strDateFrom & " And EP07<=" & strDateTo & " and cp09(+)=EP02 "
      
      If Not m_bol108Rule Then 'Added by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除)
      
         'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
         'StrSqlB = StrSqlB & " union all Select Sum(cp18*0.05) pp From caseprogress Where cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And cp18>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S'"
         'Modified by Morgan 2014/3/19 2014/4/1起非智權收文改每點折算0.04基數
         'StrSqlB = StrSqlB & " union all Select Sum(nvl(a0n03/1000,cp18)*0.05) pp From caseprogress ,acc0n0 where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S'"
         StrSqlB = StrSqlB & " union all Select Sum(" & Pt2EPtCode & ") pp From caseprogress ,acc0n0 where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S'"
         'end 2014/3/19
         'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
         StrSqlB = StrSqlB & " Union All Select Sum(Round(Nvl(MH05,0)*0.2 ,2)) pp From ModifyHour Where MH02='" & strST01 & "' And MH01>=" & strDateFrom & " And MH01<=" & strDateTo & " And MH11='V'"
         StrSqlB = StrSqlB & " Union All Select Sum(Round(Nvl(EH05,0)*0.25 ,2)) pp From ExtendHour Where EH02='" & strST01 & "' And EH01>=" & strDateFrom & " And EH01<=" & strDateTo & " And EH11='V'"
         'end 2011/8/1
         
      End If 'Added by Morgan 2019/3/18
      
      'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
      'StrSqlB = StrSqlB & " Union All Select Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2)) pp From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' ) X"
      'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
      StrSqlB = StrSqlB & " Union All Select Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)) pp From SupportHour,staff Where st01(+)=sh02 and SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' ) X"
      'end 2014/3/19
      
   'Removed by Morgan 2014/3/20 早已實施不用再判斷
   'Else
   '   StrSqlB = "select nvl(sum(pp),0) from (select sum(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)) pp from engineerprogress,caseprogress where EP05='" & strST01 & "' And EP07>=" & strDateFrom & " And EP07<=" & strDateTo & " and cp09(+)=EP02 " & _
   '      " Union All Select Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2)) pp From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' ) X"
   'End If
   'end 2014/3/20
   
   If rsB.State <> adStateClosed Then rsB.Close
   rsB.CursorLocation = adUseClient
   rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If rsB.RecordCount > 0 Then
      strMA54 = rsB(0)
   End If
   'end 2010/3/30
   
End If

If rsA.State <> adStateClosed Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
While Not rsA.EOF
   If oIsDraw = False Then
      'edit by nickc 2006/02/22 加入會稿加乘註記
      'CalFinishNew = str((Val("" & rsA.Fields("cp97").Value) * Val("" & rsA.Fields("cp98").Value)) + Val(CalFinishNew))
      'Modify by Morgan 2010/10/26
      'CalFinishNew = str((Val("" & rsA.Fields("cp97").Value) * Val("" & rsA.Fields("cp98").Value) * IIf(CheckStr(rsA.Fields("cp112")) = "Y", Val(CheckStr(rsA.Fields("cp111"))), 1)) + Val(CalFinishNew))
      CalFinishNew = str((Val("" & rsA.Fields("cp97").Value) * Val("" & rsA.Fields("cp98").Value) * IIf(CheckStr(rsA.Fields("cp112")) = "Y", Val(CheckStr(IIf(IsNull(rsA.Fields("cp111")), 1, rsA.Fields("cp111")))), 1)) + Val(CalFinishNew))
      
      'Remove by Morgan 2010/3/30 累計會稿量統計條件改只考慮會稿日區間,不管完稿日區間
      ''Add by Morgan 2009/7/14
      'If IsNull(rsA.Fields("ep07")) Then
      '   dblDeMA54 = dblDeMA54 + Val("" & rsA.Fields("cp97").Value) * Val("" & rsA.Fields("cp98").Value) * IIf(CheckStr(rsA.Fields("cp112")) = "Y", Val(CheckStr(rsA.Fields("cp111"))), 1)
      'End If
      
   Else
      If oIsRough = True Then  '草圖
         CalFinishNew = str((Val("" & rsA.Fields("cp100").Value) * Val("" & rsA.Fields("cp101").Value)) + Val(CalFinishNew))
      Else
         CalFinishNew = str((Val("" & rsA.Fields("cp103").Value) * Val("" & rsA.Fields("cp104").Value)) + Val(CalFinishNew))
      End If
   End If
    rsA.MoveNext
Wend

If rsA.State <> adStateClosed Then rsA.Close
If oIsDraw = False Then
      'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
      'StrSQLa = "Select Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2)) pp From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' "
      'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
      StrSQLa = "Select Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)) pp From SupportHour,staff Where st01(+)=sh02 and SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' "
      'end 2014/3/19
      'Add by Morgan 2010/11/8
      'Modified by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除,另新制早上線不必再判斷)
      'If bolNewPromoterRule Then
      If Not m_bol108Rule Then
      'end 2019/3/818
      
         'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
         StrSQLa = StrSQLa & " Union All Select Sum(Round(Nvl(MH05,0)*0.2 ,2)) pp From ModifyHour Where MH02='" & strST01 & "' And MH01>=" & strDateFrom & " And MH01<=" & strDateTo & " And MH11='V'"
         StrSQLa = StrSQLa & " Union All Select Sum(Round(Nvl(EH05,0)*0.25 ,2)) pp From ExtendHour Where EH02='" & strST01 & "' And EH01>=" & strDateFrom & " And EH01<=" & strDateTo & " And EH11='V'"
         'end 2011/8/1
         '收文點數
         'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
         'StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(cp18*0.05) pp From caseprogress Where cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And cp18>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S') X"
         'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
         'StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(nvl(a0n03/1000,cp18)*0.05) pp From caseprogress ,acc0n0 where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S') X"
         StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(" & Pt2EPtCode & ") pp From caseprogress ,acc0n0 where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S') X"
         'end 2014/3/20
      End If
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          CalFinishNew = Val(CalFinishNew) + Val("" & rsA.Fields(0).Value)
      End If
Else
      '與工程師算法相同，4 hr= 一件，但是因為草墨各一件，所以除4  因為草墨都會各來抓一次
      'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數(原來沒有和工程師同步,本次一併更正)
      'StrSQLa = "Select Sum(Round(Nvl(SH05,0)/4,2)) pp From SupportHour,staff Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' AND ST01(+)=SH02 AND ST05<>'79' "
      StrSQLa = "Select Sum(Round(" & Sh2EPtCode & ",2)) pp From SupportHour,staff Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo & " And SH11='V' AND ST01(+)=SH02 AND ST05<>'79' "
      'end 2014/3/19
      'Add by Morgan 2010/11/8
      'Modified by Morgan 2019/3/18 108考核(取消收文點數轉換,另原修改紀錄及衍生工作紀錄103年就取消,一併排除,另新制早上線不必再判斷)
      'If bolNewPromoterRule Then
      If Not m_bol108Rule Then
      'end 2019/3/818
      
         'Add by Morgan 2011/8/1 + 修改紀錄,衍生工作紀錄
         StrSQLa = StrSQLa & " Union All Select Sum(Round(Nvl(MH05,0)*0.2 ,2)) pp From ModifyHour,staff Where MH02='" & strST01 & "' And MH01>=" & strDateFrom & " And MH01<=" & strDateTo & " And MH11='V' AND ST01(+)=MH02 AND ST05<>'79' "
         StrSQLa = StrSQLa & " Union All Select Sum(Round(Nvl(EH05,0)*0.25 ,2)) pp From ExtendHour,staff Where EH02='" & strST01 & "' And EH01>=" & strDateFrom & " And EH01<=" & strDateTo & " And EH11='V' AND ST01(+)=EH02 AND ST05<>'79' "
         'end 2011/8/1
         '收文點數
         'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
         'StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(cp18*0.05) pp From caseprogress Where cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And cp18>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S') X"
         'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
         'StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(nvl(a0n03/1000,cp18)*0.05) pp From caseprogress ,acc0n0,staff where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' AND ST01(+)=CP13 AND ST05<>'79' ) X"
         StrSQLa = "select sum(pp) from (" & StrSQLa & " union all Select Sum(" & Pt2EPtCode & ") pp From caseprogress ,acc0n0,staff where a0n02(+)=cp09 and cp13='" & strST01 & "' And cp05>=" & strDateFrom & " And cp05<=" & strDateTo & " And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp57 is null and substr(cp12,1,1)<>'S' AND ST01(+)=CP13 AND ST05<>'79' ) X"
         'end 2014/3/19
      End If
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If oIsRough = True Then  '草圖
            'edit by nickc 2005/08/03
            'CalFinishNew = Val(CalFinishNew) + (Val("" & rsA.Fields(0).Value) * 0.65)
            CalFinishNew = Val(CalFinishNew) + (Val("" & rsA.Fields(0).Value) * 0.65 * 2)
         Else
            'edit by nickc 2005/08/03
            'CalFinishNew = Val(CalFinishNew) + (Val("" & rsA.Fields(0).Value) * 0.35)
            CalFinishNew = Val(CalFinishNew) + (Val("" & rsA.Fields(0).Value) * 0.35 * 2)
         End If
      End If
End If

'Remove by Morgan 2010/3/30 累計會稿量統計條件改只考慮會稿日區間,不管完稿日區間
'strMA54 = Format(Val(CalFinishNew) - dblDeMA54, "0.00") 'Add by Morgan 2009/7/14

CalFinishNew = Format(CalFinishNew, "0.00")

Set rsA = Nothing
Set rsB = Nothing

End Function

'add by nickc 2005/03/01 直接抓資料庫資料
Function StrMenu() As Boolean
Dim iii As Integer
Dim BeginDay As String
Dim EndDay As String
Dim StrSQLa As String
Dim bolMove2User As Boolean, iColUser As Integer 'Added by Morgan 2019/3/25

If frm090624.txt1(0) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Left(frm090624.Label1(2), 5) & frm090624.txt1(0) & "(以" & frm090624.Label1(5) & "計算)" 'Add By Sindy 2010/12/17
End If
If frm090624.txt1(1) <> "" Or frm090624.txt1(2) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090624.Label1(3) & frm090624.txt1(1) & "-" & frm090624.txt1(2) & "(" & frm090624.Label1(6) & ")" 'Add By Sindy 2010/12/17
End If
If frm090624.txt1(3) <> "" Or frm090624.txt1(4) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090624.Label1(0) & frm090624.txt1(3) & "-" & frm090624.txt1(4) & "(" & frm090624.Label1(7) & ")" 'Add By Sindy 2010/12/17
End If
If frm090624.txt1(5) <> "" Or frm090624.txt1(6) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090624.Label1(1) & frm090624.txt1(5) & "-" & frm090624.txt1(6) & "(" & frm090624.Label1(8) & ")" 'Add By Sindy 2010/12/17
End If
If frm090624.txt1(7) <> "" Or frm090624.txt1(7) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090624.Label1(4) & frm090624.txt1(7) & "-" & frm090624.txt1(8) & "(" & frm090624.Label1(9) & ")" 'Add By Sindy 2010/12/17
End If
InsertQueryLog ("") 'Add By Sindy 2010/12/17

StrMenu = True
StrSQLa = "select distinct S1.st02 as st02,MA.*,s1.st06 from monthassess MA,staff S1,staff S2 where ma01=S1.st01(+)  and S1.st03=S2.st03(+) "
Select Case ProState
Case "2"   '管理 可以讀全部
      StrSQLa = StrSQLa & " and ma02=" & Val(frm090624.txt1(0)) + 191100
      'add by nickc 2005/04/12 依等級區分權限
      Select Case PUB_GetST05(strUserNum)
      Case "77"
               StrSQLa = StrSQLa & " and s2.st03 in ('P10','P11') "
      Case "81", "82"
              StrSQLa = StrSQLa & " and s2.st03 in ('P13')  "
      Case Else
      End Select
      StrSQLa = StrSQLa & " and s1.st04='1'  "
      
Case Else   '個人只能讀同部門同所
      'edit by nickc 2005/03/27 前 5 個工作天 ，可以查同區，且為上個月
      'edit by nickc 2005/04/12  依照前畫面所輸入之月份
      'StrSQLA = StrSQLA & " and to_char(ma02)=(select max(substr(to_char(wd01),1,6)) from workday where wd01<to_number(substr(to_char(sysdate, 'YYYYMMDD'),1,6)||'01') ) "
      StrSQLa = StrSQLa & " and ma02=" & Val(frm090624.txt1(0)) + 191100
      BeginDay = Trim(Val(frm090624.txt1(0).Text & Format("01", "00")) + 19110000)
      EndDay = CompWorkDay(5, BeginDay)
      
      'Added by Morgan 2019/3/25 108考核,開放工程師成員(P11)可以查詢最近兩個月(當月及前一個月)的所有人員的速度考核及月考核成績的功能
      If Pub_StrUserSt03 = "P11" And ProSysState = "1" And strSrvDate(1) >= PUB_108RuleDate And Left(BeginDay, 6) >= Left(CompDate(1, -1, strSrvDate(1)), 6) Then
         StrSQLa = StrSQLa & " and S2.st01='" & strUserNum & "'"
         bolMove2User = True
      Else
      'end 2019/3/25
      
         If strSrvDate(1) >= BeginDay And strSrvDate(1) <= EndDay Then
            StrSQLa = StrSQLa & " and S2.st01='" & strUserNum & "' "
         Else
            'Modified by Morgan 2019/3/25
            'StrSQLa = StrSQLa & " and S1.st01='" & strUserNum & "'"
            StrSQLa = StrSQLa & " and S1.st01='" & strUserNum & "' and s2.st01(+)=s1.st01"
            'end 2019/3/25
         End If
         
      End If 'Adde by Morgan 2019/3/25
      
      
      StrSQLa = StrSQLa & " and s1.st04='1'  "
End Select
 'edit by nickc 2005/08/03
 'strSQLA = strSQLA & " order by ma03 desc ,ma01,ma36 "
 StrSQLa = StrSQLa & " order by ma03 desc,s1.st06 ,ma01,ma36 "
 CheckOC3
 With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         stb.TabEnabled(0) = False
         stb.TabEnabled(1) = False
'         grd(0).ColWidth(2) = 0
'         grd(1).ColWidth(2) = 0
         .MoveFirst
         grd(0).Visible = False
         grd(1).Visible = False
         Do While Not .EOF
            stb.TabEnabled(Val(.Fields("ma03").Value) - 1) = True
            stb.Tab = .Fields("ma03").Value - 1
            '設定變色 flag
            If Trim(grd(Val(.Fields("ma03").Value) - 1).TextMatrix(0, 2)) <> "" Then
               grd(Val(.Fields("ma03").Value) - 1).Cols = grd(Val(.Fields("ma03").Value) - 1).Cols + 1
            End If
            grd(Val(.Fields("ma03").Value) - 1).col = grd(Val(.Fields("ma03").Value) - 1).Cols - 1
            grd(Val(.Fields("ma03").Value) - 1).row = 0
            grd(Val(.Fields("ma03").Value) - 1).Text = .Fields("st02").Value
            grd(Val(.Fields("ma03").Value) - 1).CellAlignment = flexAlignCenterCenter
            If .Fields("ma03").Value = "1" Then
               grd(Val(.Fields("ma03").Value) - 1).row = 1
               grd(Val(.Fields("ma03").Value) - 1).Text = "稿"
            Else
               If .Fields("ma36").Value = "1" Then
                     grd(Val(.Fields("ma03").Value) - 1).row = 1
                     grd(Val(.Fields("ma03").Value) - 1).Text = "草"
                     grd(Val(.Fields("ma03").Value) - 1).CellAlignment = flexAlignCenterCenter
               Else
                     grd(Val(.Fields("ma03").Value) - 1).row = 1
                     grd(Val(.Fields("ma03").Value) - 1).Text = "墨"
                     grd(Val(.Fields("ma03").Value) - 1).CellAlignment = flexAlignCenterCenter
               End If
            End If
            For iii = 2 To grd(Val(.Fields("ma03").Value) - 1).Rows - 2
                     grd(Val(.Fields("ma03").Value) - 1).row = iii
                     grd(Val(.Fields("ma03").Value) - 1).Text = .Fields("ma" & Format(iii + 2, "00")).Value
            Next iii
            grd(Val(.Fields("ma03").Value) - 1).CellAlignment = flexAlignCenterCenter
            If .Fields("ma36").Value = "2" Then '若是墨圖，要加一 col
               grd(Val(.Fields("ma03").Value) - 1).Cols = grd(Val(.Fields("ma03").Value) - 1).Cols + 1
               grd(Val(.Fields("ma03").Value) - 1).col = grd(Val(.Fields("ma03").Value) - 1).Cols - 1
               grd(Val(.Fields("ma03").Value) - 1).ColWidth(grd(Val(.Fields("ma03").Value) - 1).Cols - 1) = 0
            End If
            
            'Added by Morgan 2019/3/25
            If .Fields("ma01") = strUserNum Then
               iColUser = grd(Val(.Fields("ma03").Value) - 1).col
            End If
            'end 2019/3/25
            
            .MoveNext
         Loop
         grd(0).Visible = True
         grd(1).Visible = True
         
         'Added by Morgan 2019/3/25
         If bolMove2User And iColUser > 0 Then
            grd(0).LeftCol = iColUser
         End If
         'end 2019/3/25
         
      Else
         ShowNoData
         If ProState = "2" Then
            StrMenu = True
         Else
            StrMenu = False
         End If
      End If
 End With
 CheckOC3
End Function

Function GetDelay(oUser As String, oIsRoung As Boolean) As Integer
Dim strMonthLastDate As String '某月份最後一天
Dim strBeginDate As String
Dim strEndDate As String
Dim strSQLc As String
Dim NickRS As New ADODB.Recordset
Dim strSQL1 As String
Dim StrSQL6 As String
    GetDelay = 0
'    Screen.MousePointer = vbHourglass
    Set NickRS = New ADODB.Recordset
      strBeginDate = Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString((Val(frm090624.txt1(0).Text) + 191100) & "01"))), 6) & "01"
      strEndDate = Left((Val(frm090624.txt1(0).Text) + 191100), 6) & PUB_GetMonthDays(Left((Val(frm090624.txt1(0).Text) + 191100), 4), Mid((Val(frm090624.txt1(0).Text) + 191100), 5, 2))
      strSQL1 = " AND CP05<=" & Val(frm090624.txt1(0).Text) + 191100 & "31 "
      StrSQL6 = " AND CP05<=" & Val(frm090624.txt1(0).Text) + 191100 & "31 "
      strSQL1 = " AND EP13='" & Trim(oUser) & "' " & strSQL1
      strSQL1 = strSQL1 & " And cp05>=19980101 "
      StrSQL6 = " AND EP13='" & Trim(oUser) & "'  " & StrSQL6
      StrSQL6 = StrSQL6 & " And cp05>=19980101 "
      strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
      
    strSQLc = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(oUser) & "' And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL6
    strSQLc = strSQLc & " Union All SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57, PA08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(oUser) & "' And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & " ) " & StrSQL6
    
    NickRS.CursorLocation = adUseClient
    NickRS.Open strSQLc, cnnConnection, adOpenStatic, adLockReadOnly
    If NickRS.RecordCount <> 0 Then
        NickRS.MoveFirst
        Do While NickRS.EOF = False
            Select Case CheckStr(NickRS.Fields("PA08").Value)
               Case "3"
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                     If oIsRoung = True Then
                          '若有草齊日及草完日
                          If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                              '草完日必須為當月
                              If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                                  If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 5 Then
                                      GetDelay = GetDelay + 1
                                  End If
                              End If
      '                    '若無發文日有草齊日無草完日無取消收文日
                          '若有草齊日無草完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                          ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                              If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 5 Then
'                                 GetDelay = GetDelay + 1
'                              End If
                          End If
                     End If
                Else '墨圖
                    '若有墨齊日及墨完日
                    If oIsRoung = False Then
                          If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                              '墨完日必須為當月
                              If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                                  If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 3 Then
                                      GetDelay = GetDelay + 1
                                  End If
                              End If
      '                    '若無發文日有墨齊日無墨完日無取消收文日
                          '若有墨齊日無墨完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                          ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                              If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                                  GetDelay = GetDelay + 1
'                              End If
                          End If
                     End If
                End If
            Case Else
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                     If oIsRoung = True Then
                             '若有草齊日及草完日
                             If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                                 '草完日必須為當月
                                 If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                                     If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 4 Then
                                        GetDelay = GetDelay + 1
                                     End If
                                 End If
         '                    '若無發文日有草齊日無草完日無取消收文日
                             '若有草齊日無草完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                             ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                                 If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 4 Then
'                                    GetDelay = GetDelay + 1
'                                 End If
                             End If
                        End If
                Else '墨圖
                     If oIsRoung = False Then
                             '若有墨齊日及墨完日
                             If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                                 '墨完日必須為當月
                                 If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.Text1.Text) + 191100), Left(strSrvDate(1), 6)) Then
                                     If GetWorkDay((CheckStr(NickRS.Fields(4))), "" & NickRS.Fields(2).Value) > 3 Then
                                         GetDelay = GetDelay + 1
                                     End If
                                 End If
         '                    '若無發文日有墨齊日無墨完日無取消收文日
                             '若有墨齊日無墨完日
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                             ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                                 If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                                    GetDelay = GetDelay + 1
'                                 End If
                             End If
                     End If
                End If
            End Select
            NickRS.MoveNext
        Loop
    End If
    If NickRS.State = 1 Then NickRS.Close
'Screen.MousePointer = vbDefault
End Function
