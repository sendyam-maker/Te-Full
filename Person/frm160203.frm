VERSION 5.00
Begin VB.Form frm160203 
   BorderStyle     =   1  '單線固定
   Caption         =   "部門加班時數統計"
   ClientHeight    =   3290
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3290
   ScaleWidth      =   5440
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3510
      TabIndex        =   3
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4455
      TabIndex        =   4
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2430
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1290
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1290
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2430
      MaxLength       =   5
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   5
      Top             =   2610
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1500
      TabIndex        =   9
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   3390
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "列印年月："
      Height          =   180
      Left            =   1500
      TabIndex        =   8
      Top             =   930
      Width           =   900
   End
End
Attribute VB_Name = "frm160203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2009/01/07
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt4 As Double, dblTotAmt4 As Double
Dim dblAmt5 As Double, dblTotAmt5 As Double
Dim dblAmt6 As Double, dblTotAmt6 As Double
Dim dblCnt As Double, dblTotCnt As Double
Dim strTName As String


Private Sub cmdok_Click(Index As Integer)
Dim strSDate As String, strEDate As String
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "列印年月不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If Len(txt1(0)) <= 3 Then
            MsgBox "列印年月輸入錯誤！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            strSDate = Left(ChangeTStringToWString(Trim(txt1(0)) & "01"), 6) & "01"
            strEDate = Left(ChangeTStringToWString(Trim(txt1(0)) & "01"), 6) & "31"
            m_StrSQL = m_StrSQL & "AND SO02 Between '" & strSDate & "' AND '" & strEDate & "' "
        End If
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
'            If Val(txt1(0)) >= 11301 Then
               m_StrSQL = m_StrSQL & " AND ST93 >= '" & txt1(1) & "' "
'            Else
'            '2023/12/27 END
'               m_StrSQL = m_StrSQL & " AND ST03 >= '" & txt1(1) & "' "
'            End If
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
'            If Val(txt1(0)) >= 11301 Then
               m_StrSQL = m_StrSQL & " AND ST93 <= '" & txt1(2) & "' "
'            Else
'            '2023/12/27 END
'               m_StrSQL = m_StrSQL & " AND ST03 <= '" & txt1(2) & "' "
'            End If
        End If
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()
Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'Modify By Sindy 2023/12/27 部門調整改抓ST93
If Val(txt1(0)) >= 11301 Then
   m_str = "SELECT a0921,a0922 a0902,ST01,ST02,sum(nvl(SO05,0)) T4,sum(nvl(SO06,0)) T5,sum(nvl(SO05,0))+sum(nvl(SO06,0)) T6 " & _
            "From Staff, Staff_Overtime, acc090NEW " & _
            "Where ST01 = SO01 " & m_StrSQL & _
            "AND ST93=a0921(+) " & _
            "Group by a0921,a0922,ST01,ST02 " & _
            "Order by a0921,ST01 "
Else
'2023/12/27 END
   m_str = "SELECT a0901,a0902,ST01,ST02,sum(nvl(SO05,0)) T4,sum(nvl(SO06,0)) T5,sum(nvl(SO05,0))+sum(nvl(SO06,0)) T6 " & _
            "From Staff, Staff_Overtime, acc090 " & _
            "Where ST01 = SO01 " & m_StrSQL & _
            "AND ST03=a0901(+) " & _
            "Group by a0901,a0902,ST01,ST02 " & _
            "Order by a0901,ST01 "
End If
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        dblAmt4 = 0
        dblAmt5 = 0
        dblCnt = 0
        dblTotAmt4 = 0
        dblTotAmt5 = 0
        dblTotCnt = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 5 '4
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("a0902"))
            strTemp(2) = CheckStr(m_rs.Fields("ST01"))
            strTemp(3) = CheckStr(m_rs.Fields("ST02"))
            strTemp(4) = CheckStr(m_rs.Fields("T4"))
            strTemp(5) = CheckStr(m_rs.Fields("T5")) 'Add By Sindy 2016/7/26
            strTemp(6) = CheckStr(m_rs.Fields("T6")) 'Add By Sindy 2016/7/26
            
            If iLine > 48 Or iLine = 1 Then
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                'End If
            End If
            
            If (strType <> "" And strType <> strTemp(1)) Then
               PrintEnd '小計
               
               Printer.CurrentX = 500
               Printer.CurrentY = iLine * 300
               Printer.Print String(140, "-")
               iLine = iLine + 1
            End If
            If strType <> strTemp(1) Then
               strTName = strTemp(1)
            Else
               strTName = ""
            End If
            
            PrintDetail '列印表中
            
            strType = strTemp(1)
            dblAmt4 = dblAmt4 + strTemp(4)
            dblAmt5 = dblAmt5 + strTemp(5) 'Add By Sindy 2016/7/26
            dblAmt6 = dblAmt6 + strTemp(6) 'Add By Sindy 2016/7/26
            dblCnt = dblCnt + 1
            dblTotAmt4 = dblTotAmt4 + strTemp(4)
            dblTotAmt5 = dblTotAmt5 + strTemp(5) 'Add By Sindy 2016/7/26
            dblTotAmt6 = dblTotAmt6 + strTemp(6) 'Add By Sindy 2016/7/26
            dblTotCnt = dblTotCnt + 1
            m_rs.MoveNext
        Loop
        
         '列印表尾
         PrintEnd '小計
         
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iLine * 300
         Printer.Print "總　計："
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(dblTotCnt & "人")
         Printer.CurrentY = iLine * 300
         Printer.Print dblTotCnt & "人"
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTotAmt4, "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt4, "##0.0")
         'Add By Sindy 2016/7/26
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt5, "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt5, "##0.0")
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTotAmt6, "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt6, "##0.0")
         '2016/7/26 END
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(dblCnt & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCnt & "人"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmt4, "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt4, "##0.0")
   'Add By Sindy 2016/7/26
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt5, "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt5, "##0.0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmt6, "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt6, "##0.0")
   '2016/7/26 END
   
   iLine = iLine + 1
   dblAmt4 = 0
   dblAmt5 = 0 'Add By Sindy 2016/7/26
   dblAmt6 = 0 'Add By Sindy 2016/7/26
   dblCnt = 0
End Sub


Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("部門加班時數統計表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "部門加班時數統計表"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000  年  00  月") / 2)
Printer.CurrentY = iLine * 300
Printer.Print Left(Right("0" & Trim(txt1(0)), 5), 3) & "  年  " & Right("00000" & Trim(txt1(0)), 2) & "  月"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部　　門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工編號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("平日加班時數")
Printer.CurrentY = iLine * 300
Printer.Print "平日加班時數"
'Add By Sindy 2016/7/26
Printer.CurrentX = PLeft(5) - Printer.TextWidth("假日加班時數")
Printer.CurrentY = iLine * 300
Printer.Print "假日加班時數"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("合計")
Printer.CurrentY = iLine * 300
Printer.Print "小計"
'2016/7/26 END

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 7000
PLeft(5) = 9000 'Add By Sindy 2016/7/26
PLeft(6) = 10500 'Add By Sindy 2016/7/26
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTName 'strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(4), "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##0.0")
   'Add By Sindy 2016/7/26
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(5), "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##0.0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(6), "##0.0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##0.0")
   '2016/7/26 END
   
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
   Set frm160203 = Nothing
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
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case 1, 2
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
