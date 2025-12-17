VERSION 5.00
Begin VB.Form frm160202 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "对痁る参璸"
   ClientHeight    =   3230
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3230
   ScaleWidth      =   5320
   Begin VB.Frame Frame1 
      Caption         =   "砞﹚"
      Height          =   600
      Left            =   30
      TabIndex        =   3
      Top             =   2550
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '虫┰Α
         TabIndex        =   4
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "诀"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2490
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "瞒秨(&X)"
      Height          =   435
      Index           =   1
      Left            =   4335
      TabIndex        =   2
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3390
      TabIndex        =   1
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "る"
      Height          =   180
      Left            =   1560
      TabIndex        =   6
      Top             =   1050
      Width           =   900
   End
End
Attribute VB_Name = "frm160202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 醇舦逆э
'Memo By Sindy 2011/2/17 SQLDate浪琩
'Memo By Sindy 2010/11/25 絪腹逆э
'Modify By Sindy 2010/7/21 ら戳逆э
'Create by SINDY 2009/01/08
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim strTempA(1 To 22) As String
Dim strTempA2(1 To 22) As String
Dim strTempB(1 To 22) As String
Dim strTempB2(1 To 22) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt As Double, dblAmt2 As Double, dblTotAmt As Double, dblTotAmt2 As Double
Dim strSDate As String, strEDate As String
Dim strPSDate As String, strPEDate As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "るぃフ", vbInformation, "巨岿粇"
            txt1(0).SetFocus
            Exit Sub
        End If
        If Len(txt1(0)) <= 3 Then
            MsgBox "る块岿粇", vbInformation, "巨岿粇"
            txt1(0).SetFocus
            Exit Sub
        End If
        If Val(Left(ChangeTStringToWString(Trim(txt1(0)) & "01"), 6)) > Val(Left(Format(Date, "YYYYMMDD"), 6)) Then
            MsgBox "礚才戈!!!", vbExclamation + vbOKOnly
            txt1(0).SetFocus
            Exit Sub
        End If
        
        'セる
        strSDate = Left(ChangeTStringToWString(Trim(txt1(0)) & "01"), 6) & "01"
        strEDate = Left(ChangeTStringToWString(Trim(txt1(0)) & "01"), 6) & "31"
        'る
        If Mid(strSDate, 5, 2) = "01" Then
            strPSDate = Val(strSDate) - 10100 + 1200
            strPEDate = Val(strEDate) - 10100 + 1200
        Else
            strPSDate = Val(strSDate) - 100
            strPEDate = Val(strEDate) - 100
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & "AND SO02 Between '" & strPSDate & "' AND '" & strEDate & "' "
        End If
        StrMenu1 '对参璸
        StrMenu2 '痁计参璸
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()
'Dim dblHour(18) As Double, dblCnt(18) As Double
'Dim dblHour(22) As Double, dblCnt(22) As Double...簿basPersonノ跑计跋

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1. 2.绢
'Printer.PaperSize = 9  'PDF

For m_i = 1 To 22 '21 '20
    strTempA(m_i) = ""   'セる计
    strTempA2(m_i) = "" 'セる计
    strTempB(m_i) = ""   'る计
    strTempB2(m_i) = "" 'る计
Next m_i

'眔安计の计-セる
If PUB_GetAbsenceHour("", strSDate, strEDate, dblHour(), dblCnt()) = True Then
   '10.玻安=玻安+17.Ι沧玻安
   dblHour(10) = dblHour(10) + dblHour(17)
   '11.瑈玻安=瑈玻安+18.Ι沧瑈玻安
   dblHour(11) = dblHour(11) + dblHour(18)
   '恶セる计
   strTempA(1) = dblHour(1)
   strTempA(2) = dblHour(2)
   strTempA(3) = dblHour(3)
   strTempA(4) = dblHour(4)  '04.畉
   strTempA(5) = dblHour(5)
   strTempA(6) = dblHour(22) 'Add By Sindy 2014/12/9 22.產畑酚臮安
   strTempA(7) = dblHour(24) 'Add By Sindy 2020/2/5 24.ň酚臮安
   strTempA(8) = dblHour(6)
   strTempA(9) = dblHour(20) 'Add By Sindy 2014/12/9 20.ネ瞶安
   strTempA(10) = dblHour(23) 'Add By Sindy 2015/1/5 23.胺浪安
   strTempA(11) = dblHour(7)  'そ安
   strTempA(12) = dblHour(8)
   strTempA(13) = dblHour(9)
   strTempA(14) = dblHour(21) 'Add By Sindy 2014/12/9 21.玻浪安
   strTempA(15) = dblHour(10) '10.玻安
   strTempA(16) = dblHour(11) '11.瑈玻安
   strTempA(17) = dblHour(19) 'Add By Sindy 2012/1/4 19.抄玻安
   strTempA(18) = dblHour(12)
   strTempA(19) = dblHour(13)
   strTempA(20) = dblHour(14)
   strTempA(21) = dblHour(25) 'Add By Sindy 2025/11/17 25.ぱ╝ぃ倒羱
   strTempA(22) = dblHour(15) '15.ㄤ
   '恶セる计
   strTempA2(1) = dblCnt(1)
   strTempA2(2) = dblCnt(2)
   strTempA2(3) = dblCnt(3)
   strTempA2(4) = dblCnt(4)  '04.畉
   strTempA2(5) = dblCnt(5)
   strTempA2(6) = dblCnt(22) 'Add By Sindy 2014/12/9 22.產畑酚臮安
   strTempA2(7) = dblCnt(24) 'Add By Sindy 2020/2/5 24.ň酚臮安
   strTempA2(8) = dblCnt(6)
   strTempA2(9) = dblCnt(20) 'Add By Sindy 2014/12/9 20.ネ瞶安
   strTempA2(10) = dblCnt(23) 'Add By Sindy 2015/1/5 23.胺浪安
   strTempA2(11) = dblCnt(7)  'そ安
   strTempA2(12) = dblCnt(8)
   strTempA2(13) = dblCnt(9)
   strTempA2(14) = dblCnt(21) 'Add By Sindy 2014/12/9 21.玻浪安
   strTempA2(15) = dblCnt(10) '10.玻安
   strTempA2(16) = dblCnt(11) '11.瑈玻安
   strTempA2(17) = dblCnt(19) 'Add By Sindy 2012/1/4 19.抄玻安
   strTempA2(18) = dblCnt(12)
   strTempA2(19) = dblCnt(13)
   strTempA2(20) = dblCnt(14)
   strTempA2(21) = dblCnt(25) 'Add By Sindy 2025/11/17 25.ぱ╝ぃ倒羱
   strTempA2(22) = dblCnt(15) '15.ㄤ
End If

'眔安计の计-る
If PUB_GetAbsenceHour("", strPSDate, strPEDate, dblHour(), dblCnt()) = True Then
'10.玻安=玻安+17.Ι沧玻安
   dblHour(10) = dblHour(10) + dblHour(17)
   '11.瑈玻安=瑈玻安+18.Ι沧瑈玻安
   dblHour(11) = dblHour(11) + dblHour(18)
   '恶る计
   strTempB(1) = dblHour(1)
   strTempB(2) = dblHour(2)
   strTempB(3) = dblHour(3)
   strTempB(4) = dblHour(4)  '04.畉
   strTempB(5) = dblHour(5)
   strTempB(6) = dblHour(22) 'Add By Sindy 2014/12/9 22.產畑酚臮安
   strTempB(7) = dblHour(24) 'Add By Sindy 2020/2/5 24.ň酚臮安
   strTempB(8) = dblHour(6)
   strTempB(9) = dblHour(20) 'Add By Sindy 2014/12/9 20.ネ瞶安
   strTempB(10) = dblHour(23) 'Add By Sindy 2015/1/5 23.胺浪安
   strTempB(11) = dblHour(7)  'そ安
   strTempB(12) = dblHour(8)
   strTempB(13) = dblHour(9)
   strTempB(14) = dblHour(21) 'Add By Sindy 2014/12/9 21.玻浪安
   strTempB(15) = dblHour(10) '10.玻安
   strTempB(16) = dblHour(11) '11.瑈玻安
   strTempB(17) = dblHour(19) 'Add By Sindy 2012/1/4 19.抄玻安
   strTempB(18) = dblHour(12)
   strTempB(19) = dblHour(13)
   strTempB(20) = dblHour(14)
   strTempB(21) = dblHour(25) 'Add By Sindy 2025/11/17 25.ぱ╝ぃ倒羱
   strTempB(22) = dblHour(15) '15.ㄤ
   '恶る计
   strTempB2(1) = dblCnt(1)
   strTempB2(2) = dblCnt(2)
   strTempB2(3) = dblCnt(3)
   strTempB2(4) = dblCnt(4)  '04.畉
   strTempB2(5) = dblCnt(5)
   strTempB2(6) = dblCnt(22) 'Add By Sindy 2014/12/9 22.產畑酚臮安
   strTempB2(7) = dblCnt(24) 'Add By Sindy 2020/2/5 24.ň酚臮安
   strTempB2(8) = dblCnt(6)
   strTempB2(9) = dblCnt(20) 'Add By Sindy 2014/12/9 20.ネ瞶安
   strTempB2(10) = dblCnt(23) 'Add By Sindy 2015/1/5 23.胺浪安
   strTempB2(11) = dblCnt(7)  'そ安
   strTempB2(12) = dblCnt(8)
   strTempB2(13) = dblCnt(9)
   strTempB2(14) = dblCnt(21) 'Add By Sindy 2014/12/9 21.玻浪安
   strTempB2(15) = dblCnt(10) '10.玻安
   strTempB2(16) = dblCnt(11) '11.瑈玻安
   strTempB2(17) = dblCnt(19) 'Add By Sindy 2012/1/4 19.抄玻安
   strTempB2(18) = dblCnt(12)
   strTempB2(19) = dblCnt(13)
   strTempB2(20) = dblCnt(14)
   strTempB2(21) = dblCnt(25) 'Add By Sindy 2025/11/17 25.ぱ╝ぃ倒羱
   strTempB2(22) = dblCnt(15) '15.ㄤ
End If

iLine = 1
PrintTitle '繷
PrintDetail 'いЮ

Printer.EndDoc
'ShowPrintOk
End Sub

Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("对痁る参璸") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "对痁る参璸"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "ら戳" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000    00  る") / 2)
Printer.CurrentY = iLine * 300
Printer.Print Left(Right("0" & Trim(txt1(0)), 5), 3) & "    " & Right("00000" & Trim(txt1(0)), 2) & "  る"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "Ω" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "对参璸"

iLine = iLine + 1
Printer.CurrentX = 5000 - Printer.TextWidth("セる")
Printer.CurrentY = iLine * 300
Printer.Print "セる"
Printer.CurrentX = 9000 - Printer.TextWidth("る")
Printer.CurrentY = iLine * 300
Printer.Print "る"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "兜ヘ"
'Modify By Sindy 2025/11/17 "羆计 (Ω计)" => "羆计"
Printer.CurrentX = PLeft(2) - Printer.TextWidth("羆计")
Printer.CurrentY = iLine * 300
Printer.Print "羆计"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("羆计")
Printer.CurrentY = iLine * 300
Printer.Print "羆计"
'Modify By Sindy 2025/11/17 "羆计 (Ω计)" => "羆计"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("羆计")
Printer.CurrentY = iLine * 300
Printer.Print "羆计"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("羆计")
Printer.CurrentY = iLine * 300
Printer.Print "羆计"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 1000
PLeft(2) = 4000
PLeft(3) = 5500
PLeft(4) = 8000
PLeft(5) = 9500
End Sub

Sub PrintDetail()
Dim m_j As Integer
   
   'For m_j = 1 To 15
   For m_j = 1 To 22 '21 '20 '19 '16
       Printer.CurrentX = PLeft(1)
       Printer.CurrentY = iLine * 300
       If m_j = 1 Then
            Printer.Print "а ゴ "
       ElseIf m_j = 2 Then
            Printer.Print "筐      "
       ElseIf m_j = 3 Then
            Printer.Print "胢      戮"
       ElseIf m_j = 4 Then
            Printer.Print "      畉"
       ElseIf m_j = 5 Then
            Printer.Print "ㄆ      安"
       'Add By Sindy 2014/12/9 22.產畑酚臮安
       ElseIf m_j = 6 Then
            Printer.Print "產畑酚臮安"
       'Add By Sindy 2020/2/5 24.ň酚臮安
       ElseIf m_j = 7 Then
            Printer.Print "ň酚臮安"
       ElseIf m_j = 8 Then
            Printer.Print "痜      安"
       'Add By Sindy 2014/12/9 20.ネ瞶安
       ElseIf m_j = 9 Then
            Printer.Print "ネ 瞶 安"
       ElseIf m_j = 10 Then
            GoTo goStep
            Printer.Print "胺 浪 安"
       ElseIf m_j = 11 Then
            Printer.Print "そ      安"
       ElseIf m_j = 12 Then
            Printer.Print "疭  安"
       ElseIf m_j = 13 Then
            Printer.Print "盉      安"
       'Add By Sindy 2014/12/9 21.玻浪安
       ElseIf m_j = 14 Then
            Printer.Print "玻 浪 安"
       ElseIf m_j = 15 Then
            Printer.Print "玻      安"
       ElseIf m_j = 16 Then
            Printer.Print "瑈 玻 安"
       'Add By Sindy 2012/1/4 +19.抄玻安
       ElseIf m_j = 17 Then
            Printer.Print "抄 玻 安"
       ElseIf m_j = 18 Then
            Printer.Print "赤      安"
       ElseIf m_j = 19 Then
            Printer.Print "そ 端 安"
       ElseIf m_j = 20 Then
            Printer.Print "干      ヰ"
       'Add By Sindy 2025/11/17 25.ぱ╝ぃ倒羱
       ElseIf m_j = 21 Then
            Printer.Print "ぱ╝ぃ倒羱"
       ElseIf m_j = 22 Then
            Printer.Print "ㄤ      "
       End If
       
       'Modify By Sindy 2025/11/17
       If m_j = 1 Or m_j = 2 Or m_j = 3 Then
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTempA(m_j), "##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(strTempA(m_j), "##0") & IIf(m_j = 1 Or m_j = 2, "Ω", IIf(m_j = 3, "だ", ""))
       Else
       '2025/11/17 END
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTempA(m_j), "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(strTempA(m_j), "##0.0")
       End If
       Printer.CurrentX = PLeft(3) - Printer.TextWidth(strTempA2(m_j))
       Printer.CurrentY = iLine * 300
       Printer.Print strTempA2(m_j)
       'Modify By Sindy 2025/11/17
       If m_j = 1 Or m_j = 2 Or m_j = 3 Then
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTempB(m_j), "##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(strTempB(m_j), "##0") & IIf(m_j = 1 Or m_j = 2, "Ω", IIf(m_j = 3, "だ", ""))
       Else
       '2025/11/17 END
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTempB(m_j), "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(strTempB(m_j), "##0.0")
       End If
       Printer.CurrentX = PLeft(5) - Printer.TextWidth(strTempB2(m_j))
       Printer.CurrentY = iLine * 300
       Printer.Print strTempB2(m_j)
       
       iLine = iLine + 1
goStep:
   Next m_j
End Sub

Sub StrMenu2()
Dim strSql As String
'Dim dblHour(18) As Double, dblCnt(18) As Double
'Dim dblHour(22) As Double, dblCnt(22) As Double...簿basPersonノ跑计跋

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1. 2.绢
'Printer.PaperSize = 9  'PDF

'琩高るセるΤ痁计场
'Modify By Sindy 2023/12/27 场秸俱эъST93
If Val(txt1(0)) >= 11301 Then
   m_str = "SELECT substr(ST93,1,1),ST93 ST03,a0923 a0902,sum(nvl(SO05,0))+sum(nvl(SO06,0)) as T5 " & _
            "From Staff, Staff_Overtime, acc090NEW " & _
            "Where ST01 = SO01 " & m_StrSQL & _
            "AND ST93=a0921(+) " & _
            "Group by ST93,a0923 " & _
            "Order by ST93 "
Else
'2023/12/27 END
   m_str = "SELECT substr(ST03,1,1),ST03,a0902,sum(nvl(SO05,0))+sum(nvl(SO06,0)) as T5 " & _
            "From Staff, Staff_Overtime, acc090 " & _
            "Where ST01 = SO01 " & m_StrSQL & _
            "AND ST03=a0901(+) " & _
            "Group by ST03,a0902 " & _
            "Order by ST03 "
End If
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        '箇砞
        iLine = 1
        strType = "" 'ち兵ン
        dblAmt = 0
        dblAmt2 = 0
        dblTotAmt = 0
        dblTotAmt2 = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 3
                strTemp(m_i) = ""
            Next m_i
            
            'Modify By Sindy 2023/12/27 场秸俱эъST93
            If Val(txt1(0)) >= 11301 Then
               strSql = " and ST93='" & CheckStr(m_rs.Fields("ST03")) & "' "
            Else
            '2023/12/27 END
               strSql = " and ST03='" & CheckStr(m_rs.Fields("ST03")) & "' "
            End If
            
            strTemp(1) = CheckStr(m_rs.Fields("a0902"))
            'セる计 (16.痁)
            If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
               strTemp(2) = dblHour(16)
            Else
               strTemp(2) = 0
            End If
            'る计 (16.痁)
            If PUB_GetAbsenceHour(strSql, strPSDate, strPEDate, dblHour(), dblCnt()) = True Then
               strTemp(3) = dblHour(16)
            Else
               strTemp(3) = 0
            End If
            
            If iLine > 48 Or iLine = 1 Then
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle2 '繷
                'End If
            End If
            
            If (strType <> "" And strType <> CheckStr(m_rs.Fields(0))) Then
               PrintEnd2 '璸
               
               Printer.CurrentX = 500
               Printer.CurrentY = iLine * 300
               Printer.Print String(140, "-")
               iLine = iLine + 1
            End If
            
            PrintDetail2 'い
            
            strType = CheckStr(m_rs.Fields(0))
            dblAmt = dblAmt + strTemp(2)
            dblAmt2 = dblAmt2 + strTemp(3)
            dblTotAmt = dblTotAmt + strTemp(2)
            dblTotAmt2 = dblTotAmt2 + strTemp(3)
            m_rs.MoveNext
        Loop
        
         'Ю
         PrintEnd2 '璸
         
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = 2000
         Printer.CurrentY = iLine * 300
         Printer.Print "羆璸"
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblTotAmt, "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt, "##0.0")
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTotAmt2, "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "##0.0")
    End With
'Else
'    MsgBox "礚才戈!!!", vbExclamation + vbOKOnly
'    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd2()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = 2000
   Printer.CurrentY = iLine * 300
   Printer.Print "璸"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblAmt, "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt, "##0.0")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt2, "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt2, "##0.0")
   
   iLine = iLine + 1
   dblAmt = 0
   dblAmt2 = 0
End Sub

Sub PrintTitle2()
GetPleft2

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("对痁る参璸") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "对痁る参璸"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "ら戳" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000    00  る") / 2)
Printer.CurrentY = iLine * 300
Printer.Print Left(Right("0" & Trim(txt1(0)), 5), 3) & "    " & Right("00000" & Trim(txt1(0)), 2) & "  る"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "Ω" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "痁计参璸"

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "场"
Printer.CurrentX = PLeft(2) - Printer.TextWidth("セる计")
Printer.CurrentY = iLine * 300
Printer.Print "セる计"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("る计")
Printer.CurrentY = iLine * 300
Printer.Print "る计"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "弧"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft2()
PLeft(1) = 1000
PLeft(2) = 5500
PLeft(3) = 7500
PLeft(4) = 8000
End Sub

Sub PrintDetail2()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTemp(2), "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(2), "##0.0")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(3), "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(3), "##0.0")
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160202 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
