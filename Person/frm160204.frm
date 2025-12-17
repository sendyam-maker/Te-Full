VERSION 5.00
Begin VB.Form frm160204 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡明細資料"
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
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2430
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1170
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   3060
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1170
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2430
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1530
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   3210
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1530
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2430
      MaxLength       =   5
      TabIndex        =   0
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3510
      TabIndex        =   5
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   7
      Top             =   2610
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1500
      TabIndex        =   12
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1500
      TabIndex        =   11
      Top             =   1560
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   3390
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "列印年月："
      Height          =   180
      Left            =   1500
      TabIndex        =   10
      Top             =   810
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   3150
      X2              =   3390
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frm160204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by SINDY 2013/8/5
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
Dim dblAmt As Double, dblTotAmt As Double
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
        If Len(txt1(0)) < 5 Then
            MsgBox "列印年月輸入錯誤！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            strSDate = Val((txt1(0) & "01")) + 19110000
            strEDate = Val((txt1(0) & "31")) + 19110000
            m_StrSQL = m_StrSQL & "AND pr01 Between '" & strSDate & "' AND '" & strEDate & "' "
        End If
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
'            If Val(txt1(0)) >= 11301 Then
               m_StrSQL = m_StrSQL & " AND ST93>='" & txt1(1) & "'"
'            Else
'            '2023/12/27 END
'               m_StrSQL = m_StrSQL & " AND ST03>='" & txt1(1) & "'"
'            End If
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
'            If Val(txt1(0)) >= 11301 Then
               m_StrSQL = m_StrSQL & " AND ST93<='" & txt1(2) & "'"
'            Else
'            '2023/12/27 END
'               m_StrSQL = m_StrSQL & " AND ST03<='" & txt1(2) & "'"
'            End If
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " AND ST01>='" & txt1(3) & "'"
        End If
        If txt1(4) <> "" Then
            m_StrSQL = m_StrSQL & " AND ST01<='" & txt1(4) & "'"
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

m_str = "select scd01,st02,sqldatet(pr01) as pr01,sqltime(min(pr02)) as minpr02,sqltime(max(pr02)) as maxpr02,count(*) as cnt" & _
        " From pollrecord,staffcarddata,staff" & _
        " where scd01=st01(+) and scd02=pr03(+)" & m_StrSQL & _
        " group by scd01,st02,pr01" & _
        " order by scd01,pr01"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 4
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("pr01"))
            strTemp(2) = CheckStr(m_rs.Fields("minpr02"))
            If m_rs.Fields("cnt") > 1 Then
               strTemp(3) = CheckStr(m_rs.Fields("maxpr02"))
            Else
               strTemp(3) = IIf(DBDATE(m_rs.Fields("pr01")) = strSrvDate(1), "", CheckStr(m_rs.Fields("maxpr02")))
            End If
            strTemp(4) = ""
            If m_rs.Fields("cnt") > 2 Then
               strSql = "select sqltime(pr02)" & _
                        " From pollrecord,staffcarddata,staff" & _
                        " where scd01=st01(+) and scd02=pr03(+)" & m_StrSQL & _
                        " and st01='" & m_rs.Fields("scd01") & "' and pr01=" & DBDATE(m_rs.Fields("pr01")) & _
                        " order by pr02"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     strTemp(4) = strTemp(4) & RsTemp.Fields(0) & "，"
                     RsTemp.MoveNext
                  Loop
                  If strTemp(4) <> "" Then
                     strTemp(4) = Left(strTemp(4), Len(strTemp(4)) - 1)
                  End If
               End If
            End If
            
            If iLine > 48 Or iLine = 1 Or strType <> m_rs.Fields("st02") Then
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    Call PrintTitle(m_rs.Fields("scd01") & " " & m_rs.Fields("st02"))  '列印表頭
                'End If
            End If
            
            PrintDetail '列印表中
            
            strType = m_rs.Fields("st02")
            m_rs.MoveNext
        Loop
        
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
                  
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

'Sub PrintEnd()
'   Printer.CurrentX = 500
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(140, "-")
'
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(3)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "合　計："
'   Printer.CurrentX = 8500 - Printer.TextWidth(dblCnt & "人")
'   Printer.CurrentY = iLine * 300
'   Printer.Print dblCnt & "人"
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmt, "##0.0"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmt, "##0.0")
'
'   iLine = iLine + 1
'   dblAmt = 0
'   dblCnt = 0
'End Sub


Sub PrintTitle(strUser As String)
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("打卡明細資料") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "打卡明細資料"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000  年  00  月") / 2)
Printer.CurrentY = iLine * 300
Printer.Print Left(Right("0" & Trim(txt1(0)), 5), 3) & "  年  " & Right("00000" & Trim(txt1(0)), 2) & "  月"

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名：" & strUser

Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "日期"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "上班打卡"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "下班打卡"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "備註"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 800
PLeft(2) = 2000
PLeft(3) = 3250
PLeft(4) = 4500
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   
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
   Set frm160204 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2, 3, 4
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
      Case 3, 4
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
