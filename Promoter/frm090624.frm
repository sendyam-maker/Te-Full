VERSION 5.00
Begin VB.Form frm090624 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "專利處每週速度考核表"
   ClientHeight    =   2940
   ClientLeft      =   432
   ClientTop       =   408
   ClientWidth     =   4896
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4896
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2505
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2160
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1770
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1770
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1380
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3450
      TabIndex        =   11
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2670
      TabIndex        =   10
      Top             =   20
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1020
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   1020
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1380
      MaxLength       =   5
      TabIndex        =   0
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "計算方法：             (1：新制 2：舊制 )"
      Height          =   180
      Left            =   165
      TabIndex        =   22
      Top             =   2535
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   9
      Left            =   3900
      TabIndex        =   21
      Top             =   2190
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   8
      Left            =   3900
      TabIndex        =   20
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   7
      Left            =   3900
      TabIndex        =   19
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   6
      Left            =   3900
      TabIndex        =   18
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   5
      Left            =   3900
      TabIndex        =   17
      Top             =   690
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "第四週日期："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   16
      Top             =   2205
      Width           =   1185
   End
   Begin VB.Line Line4 
      X1              =   1815
      X2              =   2970
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "第三週日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Line Line2 
      X1              =   1815
      X2              =   2970
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "第二週日期："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   1425
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   1815
      X2              =   2970
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Line Line3 
      X1              =   1815
      X2              =   2970
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "第一週日期："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   13
      Top             =   1065
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "考核月份：                             (Ex : 9206)"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   690
      Width           =   3465
   End
End
Attribute VB_Name = "frm090624"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/03 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify by Morgan 2010/12/30 新舊制選項代碼對調
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, L(1 To 5, 1 To 12) As Single
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, IngL(1 To 3) As Single, tmpnickG As Integer, k As Integer
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, StrTemp99(0 To 7) As String, StrTemp7(0 To 9) As String
Dim allG As Integer, Hightmp As Single, Middletmp As Single, Lowtmp As Single, Avgtmp As Single, Sumtmp As Single, SumCount As Single
Dim RsTmpNick As New ADODB.Recordset
'add by nick 2005/03/01
Dim IsRun As Boolean

Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer

Select Case Index
Case 0
    If Len(txt1(0)) = 0 Then
        s = MsgBox("考核月份不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1_GotFocus 0
        Exit Sub
    End If
    
    'Removed by Morga 2019/3/18 舊制早已取消
    'If Len(txt1(9)) = 0 Then
    '    s = MsgBox("計算方法不可空白!!", , "USER 輸入錯誤")
    '    txt1(9).SetFocus
    '    txt1_GotFocus 9
    '    Exit Sub
    'End If
    'end 2019/3/18
    
    If PUB_CheckKeyInYYMM(Me.txt1(0)) = -1 Then
       Me.txt1(0).SetFocus
       txt1_GotFocus 0
       Exit Sub
    End If
    For ii = 1 To 8
        If Me.txt1(ii).Text = "" Then
            MsgBox "請輸入日期!!!", vbExclamation + vbOKOnly
            Me.txt1(ii).SetFocus
            txt1_GotFocus ii
            Exit Sub
        End If
    Next ii
'    Process
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
    Me.Hide: DoEvents
    frm090624_1.Show
Case 1 '回前畫面
    Unload Me
Case Else
End Select
End Sub

Sub Process()
    ExcelSave
End Sub

Private Sub Form_Activate()
'add by nickc 2005/03/01
If IsRun = False Then
      Select Case ProState
      Case "2"     '管理
            'Modifed by Lydia 2023/04/23 修改王副總退休之相關控制
            'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "73" Or PUB_GetST05(strUserNum) = "00" Then
            'Modified by Morgan 2025/2/4 +P10部門
            'Modified by Morgan 2025/6/26 +79075
            If InStr("71,73,00,", Pub_strUserST05 & ",") > 0 Or (strSrvDate(1) >= "20230501" And Pub_strUserST05 = "72") Or Pub_StrUserSt03 = "P10" Or strUserNum = "79075" Then
                  'add by nickc 2005/03/01 加入新舊制預設值
                  'Modified by Morga 2019/3/18 舊制早已取消
                  'If CheckCanUse Then
                  '   'edit by nickc 2005/05/04
                  '   If Mid(Pub_StrUserSt03, 1, 1) = "P" Then
                  '      txt1(9).Text = "1"
                  '   Else
                  '      txt1(9).Text = "2"
                  '   End If
                  '   txt1(9).Visible = True
                  '   Label2.Visible = True
                  'Else
                  '   txt1(9).Text = "2"
                  '   txt1(9).Visible = False
                  '   Label2.Visible = False
                  'End If
                  txt1(9).Text = "1"
                  txt1(9).Visible = False
                  Label2.Visible = False
                  'end 2019/3/18
             Else
                  Me.Caption = "每週速度查詢"
                  DoEvents
                  
                  txt1(1).Locked = True
                  txt1(2).Locked = True
                  txt1(3).Locked = True
                  txt1(4).Locked = True
                  txt1(5).Locked = True
                  txt1(6).Locked = True
                  txt1(7).Locked = True
                  txt1(8).Locked = True
                  txt1(9).Locked = True
                  
                  'edit by nickc 2005/05/04
                  'Modified by Morga 2019/3/18 舊制早已取消
                  'If Mid(Pub_StrUserSt03, 1, 1) = "P" Then
                  '   txt1(9).Text = "1"
                  'Else
                  '   txt1(9).Text = "2"
                  'End If
                  txt1(9).Text = "1"
                  txt1(9).Visible = False
                  Label2.Visible = False
                  'end 2019/3/19
             End If
             
      Case Else
            Me.Caption = "每週速度查詢"
            DoEvents
            
            txt1(1).Locked = True
            txt1(2).Locked = True
            txt1(3).Locked = True
            txt1(4).Locked = True
            txt1(5).Locked = True
            txt1(6).Locked = True
            txt1(7).Locked = True
            txt1(8).Locked = True
            txt1(9).Locked = True
            
            'edit by nickc 2005/05/04
            'Modified by Morga 2019/3/18 舊制早已取消
            'If Mid(Pub_StrUserSt03, 1, 1) = "P" Then
            '   txt1(9).Text = "1"
            'Else
            '   txt1(9).Text = "2"
            'End If
            txt1(9).Text = "1"
            txt1(9).Visible = False
            Label2.Visible = False
            'end 2019/3/18
      End Select
      IsRun = True
End If
End Sub

Private Sub Form_Load()
'add by nickc 2005/03/01
IsRun = False
MoveFormToCenter Me
Me.txt1(0).Text = Left(strSrvDate(1), 6) - 191100
Me.txt1(8).Text = PUB_GetMonthDays(Left(Me.txt1(0).Text + 191100, 4), Val(Mid(Me.txt1(0).Text + 191100, 5, 2)))
SetDate

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090624 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '考核月份
   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
    Me.txt1(8).Text = PUB_GetMonthDays(Left(Me.txt1(0).Text + 191100, 4), Val(Mid(Me.txt1(0).Text + 191100, 5, 2)))
    SetDate
Case 2, 4, 6, 8 '日
    If Me.txt1(Index).Text <> "" Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
    End If
    If Me.txt1(Index - 1).Text <> "" And Me.txt1(Index).Text <> "" Then
        Me.Label1(5 + Index / 2).Caption = GetWorkDay(DBDATE(Val(Me.txt1(0).Text) & Format(Me.txt1(Index).Text, "00")), DBDATE(Val(Me.txt1(0).Text) & Format(Me.txt1(Index - 1).Text, "00")))
    End If
'add by nickc 2005/03/01
Case 9
   Select Case txt1(Index)
   Case "1", "2"
   Case Else
          MsgBox "請輸入 1 或 2 ！", , "選擇新舊制！"
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
   End Select
Case Else
End Select
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
Dim intColCount As Integer
Dim strST01 As String, strST03 As String, strST06 As String, strST13 As String, strST16 As String, strNCFPGoal As String, strCFPGoal As String, strGoal As String
    
On Error GoTo ErrorHandler
    If Dir("D:\專利速度考核" & Trim(Me.txt1(0).Text) & ".xls") = MsgText(601) Then
'        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'        End If
    Else
        Kill "D:\專利速度考核" & Trim(Me.txt1(0).Text) & ".xls"
    End If
    '承辦人速度考核
    xlsSalesPoint.Workbooks.add
    Set wksfrm090624_1 = xlsSalesPoint.Sheets("Sheet1")
    With wksfrm090624_1
        .Activate
        xlsSalesPoint.ActiveWindow.Zoom = 75
        .Name = "承辦人"
        DesignTitle wksfrm090624_1
        '員工編號
        'Modify By Cheng 2003/06/23
        '排除IAIN(88024)的資料
'        strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) Order By ST06, ST03, ST01 "
        '93.12.13 MODIFY BY SONIA 排除外翻人員
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/04/12 加一個等級
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/08/22
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'modify by sonia 2014/4/9 加入94007林景郁總經理
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74'  Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' Or ST01='94007') And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        '93.12.13 END
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        intColCount = 0
        While Not rsA.EOF
            .Range(GetCol(ii) & "1").Value = "" & rsA("ST01").Value
            .Range(GetCol(ii) & "2").Value = "稿"
            intColCount = intColCount + 1
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 3 To intColCount + 3 - 1
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strNCFPGoal = "0": strCFPGoal = "0"
            '非CFP目標件數
            'edit by nickc 2005/04/12 加一個等級
            'strSQLA = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' )) S Where S.ST01=PE01 And PE02<>'CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007')) S Where S.ST01=PE01 And PE02='P' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
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
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2005/08/22
            'strSQLB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'edit by nickc 2006/05/01
            'StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' )) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            'modify by sonia 2014/4/9 加入94007林景郁總經理
            StrSqlB = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE05,0)+Nvl(PE07,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007')) S Where S.ST01=PE01 And PE02='CFP' And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
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
            .Range(GetCol(ii) & "3").Value = CalGoal(strST01, strST03, strST06, strST13, strST16, strNCFPGoal, strCFPGoal)
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "4").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "5").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "6").Formula = "=" & GetCol(ii) & "5/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "6").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第一週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "7").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(2).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "7").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第一週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "9").Formula = "=" & GetCol(ii) & "7/" & GetCol(ii) & "6"
                .Range(GetCol(ii) & "9").Style = "Percent"
                .Range(GetCol(ii) & "9").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第一週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "8").Value = CalPoints(.Range(GetCol(ii) & "9").Value)
            End If
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "10").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "11").Formula = "=" & GetCol(ii) & "10/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "11").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "12").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(3).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "12").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "13").Formula = "=" & GetCol(ii) & "12/" & GetCol(ii) & "11"
                .Range(GetCol(ii) & "13").Style = "Percent"
                .Range(GetCol(ii) & "13").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第二週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "14").Value = CalPoints(.Range(GetCol(ii) & "13").Value)
            End If
        Next ii
        '第二週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "15").Formula = "=" & GetCol(ii) & "3/" & GetCol(ii) & "4*(" & GetCol(ii) & "5+" & GetCol(ii) & "10)"
                .Range(GetCol(ii) & "15").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "16").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "16").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "17").Formula = "=" & GetCol(ii) & "16/" & GetCol(ii) & "15"
                .Range(GetCol(ii) & "17").Style = "Percent"
                .Range(GetCol(ii) & "17").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "18").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "19").Formula = "=" & GetCol(ii) & "18/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "19").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "20").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(5).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "20").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "21").Formula = "=" & GetCol(ii) & "20/" & GetCol(ii) & "19"
                .Range(GetCol(ii) & "21").Style = "Percent"
                .Range(GetCol(ii) & "21").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第三週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "22").Value = CalPoints(.Range(GetCol(ii) & "21").Value)
            End If
        Next ii
        '第三週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "23").Formula = "=" & GetCol(ii) & "3/" & GetCol(ii) & "4*(" & GetCol(ii) & "5+" & GetCol(ii) & "10+" & GetCol(ii) & "18)"
                .Range(GetCol(ii) & "23").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "24").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "24").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "25").Formula = "=" & GetCol(ii) & "24/" & GetCol(ii) & "23"
                .Range(GetCol(ii) & "25").Style = "Percent"
                .Range(GetCol(ii) & "25").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "26").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "27").Formula = "=" & GetCol(ii) & "26/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "27").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "28").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(7).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(8).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "28").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "29").Formula = "=" & GetCol(ii) & "28/" & GetCol(ii) & "27"
                .Range(GetCol(ii) & "29").Style = "Percent"
                .Range(GetCol(ii) & "29").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第四週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "30").Value = CalPoints(.Range(GetCol(ii) & "29").Value)
            End If
        Next ii
        '第四週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "31").Value = .Range(GetCol(ii) & "3").Value
                .Range(GetCol(ii) & "31").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "32").Value = CalFinish(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(8).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "32").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "33").Formula = "=" & GetCol(ii) & "32/" & GetCol(ii) & "31"
                .Range(GetCol(ii) & "33").Style = "Percent"
                .Range(GetCol(ii) & "33").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '本月得分平均
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "34").Formula = "=(" & GetCol(ii) & "8+" & GetCol(ii) & "14+" & GetCol(ii) & "22+" & GetCol(ii) & "30)/4"
                .Range(GetCol(ii) & "34").NumberFormatLocal = "0.00"
            End If
        Next ii
        '更新資料
        For ii = 3 To intColCount + 3 - 1
            UpdateMonthAssess "1", .Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text) + 191100, .Range(GetCol(ii) & "34").Value
        Next ii
                    
        '員工編號-->員工姓名
        'Modify By Cheng 2003/06/23
        '排除IAIN(88024)的資料
'        strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) Order By ST06, ST03, ST01 "
        '93.12.13 MODIFY BY SONIA 排除外翻人員
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/04/12
        'strSQLA = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2005/08/22
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'edit by nickc 2006/05/01
        'StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' ) And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        'modify by sonia 2014/4/9 加入94007林景郁總經理
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87' Or ST01='94007') And ST01<>'88024' And ST01<'F' Order By ST06, ST03, ST01 "
        '93.12.13 END
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        While Not rsA.EOF
            .Range(GetCol(ii) & "1").Value = "" & rsA("ST02").Value
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        .Range("A1").Select
        DesignWS_Format wksfrm090624_1, 34, intColCount + 3 - 1
        .Range("A1").Select
    End With
    
    xlsSalesPoint.Workbooks(1).SaveAs FileName:="D:\專利速度考核" & Me.txt1(0).Text & ".xls"
    xlsSalesPoint.Workbooks.Open "D:\專利速度考核" & Trim(Me.txt1(0).Text) & ".xls"
'********************************************************
    '繪圖人員速度考核
    Set wksfrm090624_2 = xlsSalesPoint.Sheets("Sheet2")
    With wksfrm090624_2
        .Activate
        xlsSalesPoint.ActiveWindow.Zoom = 75
        .Name = "繪圖人員"
        DesignTitle wksfrm090624_2
        '員工編號
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='AC') Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        intColCount = 0
        While Not rsA.EOF
            .Range(GetCol(ii) & "1").Value = "" & rsA("ST01").Value
            .Range(GetCol(ii) & "2").Value = "草墨合計"
            intColCount = intColCount + 1
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '當月目標件數
        For ii = 3 To intColCount + 3 - 1
            strST01 = "": strST03 = "": strST06 = "": strST13 = "": strST16 = "": strGoal = "0"
            'Modify by Morgan 2010/10/29 改只要抓P,CFP的目標(杜燕文有T的目標)
            'StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE09,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='AC')) S Where S.ST01=PE01 And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
            StrSQLa = " Select S.ST01, S.ST03, S.ST06, S.ST13, S.ST16, Sum(Nvl(PE09,0)) From Performance, (Select ST01, ST03, ST06, ST13, ST16 From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='AC')) S Where S.ST01=PE01 And PE03=" & (Val(Me.txt1(0).Text) + 191100) & " And PE01='" & .Range(GetCol(ii) & "1").Value & "' AND PE02 IN ('P','CFP') Group By S.ST06, S.ST03, S.ST01, S.ST13, S.ST16 Order By S.ST06, S.ST03, S.ST01 "
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
            .Range(GetCol(ii) & "3").Value = IIf(strGoal <> "0", Format(strGoal, "0.00"), "0")
        Next ii
        '本月工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "4").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(1).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(2).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "5").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第一週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "6").Formula = "=" & GetCol(ii) & "5/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "6").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第一週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "7").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(2).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "7").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第一週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "9").Formula = "=" & GetCol(ii) & "7/" & GetCol(ii) & "6"
                .Range(GetCol(ii) & "9").Style = "Percent"
                .Range(GetCol(ii) & "9").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第一週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "8").Value = CalPoints(.Range(GetCol(ii) & "9").Value)
            End If
        Next ii
        
        '第二週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(3).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(4).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "10").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第二週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "11").Formula = "=" & GetCol(ii) & "10/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "11").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "12").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(3).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "12").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "13").Formula = "=" & GetCol(ii) & "12/" & GetCol(ii) & "11"
                .Range(GetCol(ii) & "13").Style = "Percent"
                .Range(GetCol(ii) & "13").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第二週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "14").Value = CalPoints(.Range(GetCol(ii) & "13").Value)
            End If
        Next ii
        '第二週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "15").Formula = "=" & GetCol(ii) & "3/" & GetCol(ii) & "4*(" & GetCol(ii) & "5+" & GetCol(ii) & "10)"
                .Range(GetCol(ii) & "15").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "16").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "16").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第二週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "17").Formula = "=" & GetCol(ii) & "16/" & GetCol(ii) & "15"
                .Range(GetCol(ii) & "17").Style = "Percent"
                .Range(GetCol(ii) & "17").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '第三週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(5).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(6).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "18").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第三週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "19").Formula = "=" & GetCol(ii) & "18/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "19").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "20").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(5).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "20").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "21").Formula = "=" & GetCol(ii) & "20/" & GetCol(ii) & "19"
                .Range(GetCol(ii) & "21").Style = "Percent"
                .Range(GetCol(ii) & "21").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第三週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "22").Value = CalPoints(.Range(GetCol(ii) & "21").Value)
            End If
        Next ii
        '第三週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "23").Formula = "=" & GetCol(ii) & "3/" & GetCol(ii) & "4*(" & GetCol(ii) & "5+" & GetCol(ii) & "10+" & GetCol(ii) & "18)"
                .Range(GetCol(ii) & "23").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "24").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "24").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第三週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "25").Formula = "=" & GetCol(ii) & "24/" & GetCol(ii) & "23"
                .Range(GetCol(ii) & "25").Style = "Percent"
                .Range(GetCol(ii) & "25").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '第四週工作天數
        StrSQLa = "Select Count(*) From WorkDay Where WD01>=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(7).Text, "00")) & " And WD01<=" & Val(Val((Me.txt1(0).Text) + 191100) & Format(Me.txt1(8).Text, "00"))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "26").Value = "" & rsA.Fields(0).Value
            End If
        Next ii
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '第四週目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "27").Formula = "=" & GetCol(ii) & "26/" & GetCol(ii) & "4*" & GetCol(ii) & "3"
                .Range(GetCol(ii) & "27").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "28").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(7).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(8).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "28").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "29").Formula = "=" & GetCol(ii) & "28/" & GetCol(ii) & "27"
                .Range(GetCol(ii) & "29").Style = "Percent"
                .Range(GetCol(ii) & "29").NumberFormatLocal = "0.00%"
            End If
        Next ii
        '第四週得分
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "30").Value = CalPoints(.Range(GetCol(ii) & "29").Value)
            End If
        Next ii
        '第四週累計目標
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "31").Value = .Range(GetCol(ii) & "3").Value
                .Range(GetCol(ii) & "31").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週累計完成
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "32").Value = CalFinish1(.Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")) + 19110000, Val(Me.txt1(0).Text & Format(Me.txt1(8).Text, "00")) + 19110000)
                .Range(GetCol(ii) & "32").NumberFormatLocal = "0.00"
            End If
        Next ii
        '第四週累計達成比例
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "33").Formula = "=" & GetCol(ii) & "32/" & GetCol(ii) & "31"
                .Range(GetCol(ii) & "33").Style = "Percent"
                .Range(GetCol(ii) & "33").NumberFormatLocal = "0.00%"
            End If
        Next ii
        
        '本月得分平均
        For ii = 3 To intColCount + 3 - 1
            '若當月有目標
            If .Range(GetCol(ii) & "3").Value <> "0" Then
                .Range(GetCol(ii) & "34").Formula = "=(" & GetCol(ii) & "8+" & GetCol(ii) & "14+" & GetCol(ii) & "22+" & GetCol(ii) & "30)/4"
                .Range(GetCol(ii) & "34").NumberFormatLocal = "0.00"
            End If
        Next ii
        '更新資料
        For ii = 3 To intColCount + 3 - 1
            UpdateMonthAssess "2", .Range(GetCol(ii) & "1").Value, Val(Me.txt1(0).Text) + 191100, .Range(GetCol(ii) & "34").Value
        Next ii
                    
        '員工編號-->員工姓名
        StrSQLa = "Select * From Staff Where ST04='1' And (ST05='79' Or ST05='81' Or ST05='AC') Order By ST06, ST03, ST01 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        ii = 3
        While Not rsA.EOF
            .Range(GetCol(ii) & "1").Value = "" & rsA("ST02").Value
            rsA.MoveNext
            ii = ii + 1
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        .Range("A1").Select
        DesignWS_Format wksfrm090624_2, 34, intColCount + 3 - 1
        .Range("A1").Select
    End With
    
    Set wksfrm090624_1 = xlsSalesPoint.Sheets("承辦人")
    wksfrm090624_1.Activate
    xlsSalesPoint.Workbooks(1).Save: DoEvents
    xlsSalesPoint.Workbooks.Close: DoEvents
    xlsSalesPoint.Quit: DoEvents
    Set xlsSalesPoint = Nothing: DoEvents
    MsgBox "Excel檔案產生完成!!!", vbExclamation + vbOKOnly
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub DesignTitle(wksfrm090624 As Worksheet)
Dim ii As Integer

    With wksfrm090624
        .Range("A1") = Val(Right(Me.txt1(0).Text, 2)) & "月份"
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
Dim jj As Integer
Dim kk As Integer

GetCol = ""
jj = Fix(intCol / 26)
kk = intCol Mod 26
Select Case jj
Case "0": GetCol = ""
Case "1": GetCol = "A"
Case "2": GetCol = "B"
Case "3": GetCol = "C"
Case "4": GetCol = "D"
Case "5": GetCol = "E"
Case "6": GetCol = "F"
Case "7": GetCol = "G"
Case "8": GetCol = "H"
Case "9": GetCol = "I"
Case "10": GetCol = "J"
Case "11": GetCol = "K"
Case "12": GetCol = "L"
Case "13": GetCol = "M"
Case "14": GetCol = "N"
Case "15": GetCol = "O"
Case "16": GetCol = "P"
Case "17": GetCol = "Q"
Case "18": GetCol = "R"
Case "19": GetCol = "S"
Case "20": GetCol = "T"
Case "21": GetCol = "U"
Case "22": GetCol = "V"
Case "23": GetCol = "W"
Case "24": GetCol = "X"
Case "25": GetCol = "Y"
Case "26": GetCol = "Z"
Case Else
End Select

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
    '若為林育輝
    If strST01 = "91013" Then
        dblMonth = 25
    Else
        dblMonth = DateDiff("m", ChangeWStringToWDateString(strST13), ChangeWStringToWDateString(Val(Me.txt1(0).Text & "01") + 19110000))
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
'若為林育輝
If strST01 = "91013" Then
    dblMonth = 25
Else
    dblMonth = DateDiff("m", ChangeWStringToWDateString(GetST13(strST01)), ChangeWStringToWDateString(Val(Me.txt1(0).Text & "01") + 19110000))
End If
'2012/1/12 modify by sonia 改公用模組
'If GetST16(strST01) = "CFP" Then
If PUB_GetStaffST16(strST01) = "CFP" Then
    CalFinish = dblNCFP / GetWeights(dblMonth) + dblCFP
Else
    CalFinish = dblNCFP + dblCFP * GetWeights(dblMonth)
End If
'strSQLA = "Select Sum(Nvl(SH04,0)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
'strSQLA = "Select Sum(Nvl(SH05,0)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
'Modify By Cheng 2004/03/01
'strSQLA = "Select Sum(Round(Nvl(SH05,0)/4,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
StrSQLa = "Select Sum(Round(Decode(SH06,'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4 ),2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
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
'strSQLA = "Select Sum(Nvl(SH04,0)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
'strSQLA = "Select Sum(Nvl(SH05,0)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
'Modify By Cheng 2004/03/01
'strSQLA = "Select Sum(Round(Nvl(SH05,0)/4,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
StrSQLa = "Select Sum(Round(Decode(SH06,'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4) ,2)) From SupportHour Where SH02='" & strST01 & "' And SH01>=" & strDateFrom & " And SH01<=" & strDateTo
'End
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    CalFinish1 = Val(CalFinish1) + Val("" & rsA.Fields(0).Value)
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

If dblMonth >= 4 And dblMonth <= 6 Then
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

'預設期間
Private Sub SetDate()
Dim intWeekDay As Integer

intWeekDay = Weekday(ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")))
'若1號小於星期三
If intWeekDay < 4 Then
    Me.txt1(2).Text = Day(DateAdd("d", ((7 - intWeekDay)), ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00"))))
'若1號大於等於星期三
Else
    Me.txt1(2).Text = Day(DateAdd("d", ((7 - intWeekDay) + 7), ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00"))))
End If
Me.txt1(3).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(2).Text, "00"))))
Me.txt1(4).Text = Day(DateAdd("d", 6, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(3).Text, "00"))))
Me.txt1(5).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00"))))
Me.txt1(6).Text = Day(DateAdd("d", 6, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(5).Text, "00"))))
Me.txt1(7).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00"))))
txt1_LostFocus 2
txt1_LostFocus 4
txt1_LostFocus 6
txt1_LostFocus 8
End Sub
