VERSION 5.00
Begin VB.Form frm170227 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "ゼヰ安灿"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5070
   Begin VB.Frame Frame1 
      Caption         =   "砞﹚"
      Height          =   600
      Left            =   20
      TabIndex        =   6
      Top             =   2160
      Width           =   5000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '虫┰Α
         TabIndex        =   3
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label2 
         Caption         =   "诀"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   0
      Top             =   930
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1380
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2610
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1380
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "挡(&X)"
      Height          =   375
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ゼヰ安"
      Height          =   180
      Left            =   580
      TabIndex        =   9
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "絪腹"
      Height          =   180
      Index           =   3
      Left            =   780
      TabIndex        =   8
      Top             =   1410
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2880
      Y1              =   1500
      Y2              =   1500
   End
End
Attribute VB_Name = "frm170227"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 醇舦逆э
'Memo by Morgan 2010/12/2 絪腹逆э
'Memo by Morgan 2010/7/27 ら戳逆э
'2009/1/21 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmtS As Double, dblAmtT As Double
Dim dblCntS As Double, dblCntT As Double
'Dim douHour(18) As Double   '对皚
'Dim douCnt(18) As Double    '对计皚
'Modify By Sindy 2012/1/4
Dim douHour(25) As Double   '对皚
Dim douCnt(25) As Double    '对计皚

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "ゼヰ安ぃフ", vbInformation, "巨岿粇"
            txt1(0).SetFocus
            Exit Sub
         End If
         If RunNick(txt1(1), txt1(2)) Then
            txt1(1).SetFocus
            Exit Sub
         End If
           
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim strSQL As String

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1. 2.绢
   
   m_StrSQL = ""
   If txt1(0) <> "" Then 'ゼヰ安
      m_StrSQL = m_StrSQL & " and yb01=" & Val(txt1(0)) + 1911
   End If
   If txt1(1) <> "" Then '絪腹癬
      m_StrSQL = m_StrSQL & " and YB02 >='" & Trim(txt1(1)) & "' "
   End If
   If txt1(2) <> "" Then '絪腹ù
      m_StrSQL = m_StrSQL & " and YB02 <='" & Trim(txt1(2)) & "' "
   End If
   
   m_str = "SELECT YB24 そ, YB03 场,YB02 絪腹,ST02 ﹎,NVL(YV04,0) ヰ疭安," & _
           "NVL(YB07,0) ゼヰ安计,NVL(YB08,0) ゼヰ安 FROM YEARBONUS,STAFF,YEARVACATION " & _
           "WHERE YB07>0 AND YB02=ST01(+) AND YB01=YV01(+) AND YB02=YV02(+) " & m_StrSQL & _
           " ORDER BY そ,场,絪腹"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '箇砞
         iLine = 1
         strType = "" 'ち兵ン
         dblAmtS = 0: dblCntS = 0
         dblAmtT = 0: dblCntT = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 12
                strTemp(m_i) = ""
            Next m_i
             
            strTemp(1) = CheckStr(m_rs.Fields(0))   'そ
            strTemp(2) = CheckStr(m_rs.Fields(2))   '絪腹
            strTemp(3) = CheckStr(m_rs.Fields(3))   '﹎
            strTemp(4) = CheckStr(m_rs.Fields(4))   'ヰ疭安
            strTemp(5) = CheckStr(m_rs.Fields(5))   'ゼヰぱ计
            strTemp(6) = CheckStr(m_rs.Fields(6))   'ゼヰ安
            
            If iLine > 50 Or iLine = 1 Or strType <> strTemp(1) Then

               If strType <> "" And strType <> strTemp(1) Then
                  PrintEnd '璸
               End If

               If iLine <> 1 Then Printer.NewPage
               iLine = 1
               PrintTitle '繷
            End If
            
            'ъ俱对皚
            strSQL = " and ST01='" & Trim(strTemp(2)) & "' "
            If PUB_GetAbsenceHour(strSQL, (Val(txt1(0)) + 1911) * 10000 + "0101", (Val(txt1(0)) + 1911) * 10000 + "1231", douHour(), douCnt()) = True Then
               strTemp(7) = Round(Val(douHour(8)) / 8, 1)   'ヰ疭安検
            End If
            
            PrintDetail 'い
            
            strType = strTemp(1) 'ㄌそ铬
            
            '璸
            dblCntS = dblCntS + 1
            dblAmtS = dblAmtS + strTemp(6)
            '璸
            dblCntT = dblCntT + 1
            dblAmtT = dblAmtT + strTemp(6)
            
            m_rs.MoveNext
         Loop
          
         'Ю
         PrintEnd    '璸
         'PrintTotal  '璸   2009/2/19 CANCEL BY SONIA 匝薇弧程щ戈┮ぃ璶
           
      End With
   Else
      MsgBox "礚才戈!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Printer.EndDoc
   ShowPrintOk

End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")

   iLine = iLine + 1
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "璸"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(dblCntS & "")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntS & ""
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmtS, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS, "###,###,###")

   dblAmtS = 0
   dblCntS = 0

End Sub

Sub PrintTotal()
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")

   iLine = iLine + 1
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "璸"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(dblCntT & "")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntT & ""
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmtT, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT, "###,###,###")
   
   dblAmtT = 0
   dblCntT = 0

End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Val(txt1(0)) & " 疭安ゼヰ灿") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print Val(txt1(0)) & " 疭安ゼヰ灿"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "ら戳" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ら戳" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "Ω" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "そ" & strTemp(1) & "" & CompNameQuery(strTemp(1))

   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "絪腹"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "﹎"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("戈")
   Printer.CurrentY = iLine * 300
   Printer.Print "戈"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("ヰ疭安")
   Printer.CurrentY = iLine * 300
   Printer.Print "ヰ疭安"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("ヰ疭安")
   Printer.CurrentY = iLine * 300
   Printer.Print "ヰ疭安"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("ゼヰぱ计")
   Printer.CurrentY = iLine * 300
   Printer.Print "ゼヰぱ计"
   Printer.CurrentX = PLeft(7) - Printer.TextWidth("疭安ゼヰ")
   Printer.CurrentY = iLine * 300
   Printer.Print "疭安ゼヰ"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 700
   PLeft(2) = 1600
   PLeft(3) = 4100
   PLeft(4) = 5600
   PLeft(5) = 7100
   PLeft(6) = 8300
   PLeft(7) = 10300
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(PUB_ChangeNianZi(CalYear(CheckStr(strTemp(2)), (Val(txt1(0)) - 1 + 1911) * 10000 + "1231")))
   Printer.CurrentY = iLine * 300
   '戈衡玡┏
   Printer.Print PUB_ChangeNianZi(CalYear(CheckStr(strTemp(2)), (Val(txt1(0)) - 1 + 1911) * 10000 + "1231"))
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(4), "##.0")) - 300
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##.0")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(7), "##.0")) - 300
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##.0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(Round(Val(strTemp(5)) / 8, 1), "#0.0")) - 100
   Printer.CurrentY = iLine * 300
   Printer.Print Format(Round(Val(strTemp(5)) / 8, 1), "#0.0")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(6), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,###,###")
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSQL As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSQL = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSQL Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170227 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
               Cancel = True
            End If
         End If
      Case 1, 2
         ' 耞腹斗 6~9 ┪ F 秨繷
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Cancel = True
            End If
         End If
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub
