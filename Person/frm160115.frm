VERSION 5.00
Begin VB.Form frm160115 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人獎懲資料明細表"
   ClientHeight    =   3260
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   12
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2370
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2370
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3585
      TabIndex        =   10
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "部門代號：                 －"
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "獎懲日期：                 －"
      Height          =   180
      Left            =   450
      TabIndex        =   7
      Top             =   1710
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號：                 －"
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   1290
      Width           =   1845
   End
End
Attribute VB_Name = "frm160115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2014/9/19
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim LongPrintCurCnt As Long
Dim StrMenu2Cnt As Long ' Add By Sindy 98/03/06


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0

        If txt1(0) = "" And txt1(1) = "" And _
            txt1(2) = "" And txt1(3) = "" And _
            txt1(4) = "" And txt1(5) = "" Then
            MsgBox "部門代號或員工代號或獎懲日期至少輸入一項！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & " and st03>='" & txt1(0) & "' "
            m_StrSQL = m_StrSQL & " and st93>='" & txt1(0) & "' "
        End If
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & " and st03<='" & txt1(1) & "' "
            m_StrSQL = m_StrSQL & " and st93<='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "' "
        End If
        If txt1(4) <> "" Then
            m_StrSQL = m_StrSQL & " and SR02>='" & ChangeTStringToWString(txt1(4)) & "' "
        End If
        If txt1(5) <> "" Then
            m_StrSQL = m_StrSQL & " and SR02<='" & ChangeTStringToWString(txt1(5)) & "' "
        End If
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()
Dim strSaSQL As String, strSbSQL As String, strSa2SQL As String

Set Printer = Printers(Combo1.ListIndex)

'XP自定紙張需手動設定並將印表機預設為該紙張
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
Printer.PaperSize = 9  'PDF
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "select ('(舊)'||A0901||' '||A0902) as A1,(ST01||' '||ST02) as A2,SR02,AC03,SR11,SR04,ST01 " & _
        " From Staff_Reward, acc090, staff, allcode " & _
        " where SR01=ST01 and SR02<20240101 and ST03=A0901 and SR03=AC02 AND AC01='08'" & m_StrSQL
m_str = m_str & " union "
m_str = m_str & "select (A0921||' '||A0922) as A1,(ST01||' '||ST02) as A2,SR02,AC03,SR11,SR04,ST01 " & _
        " From Staff_Reward, acc090NEW, staff, allcode " & _
        " where SR01=ST01 and SR02>=20240101 and ST93=A0921 and SR03=AC02 AND AC01='08'" & m_StrSQL
m_str = m_str & " order by A1,st01,sr02 "
        
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
LongPrintCurCnt = 0

If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
       m_rs.MoveFirst
        
       If LongPrintCurCnt > 0 Then
          Printer.NewPage
        End If
       iLine = 1
       
       PrintTitle '列印表頭
        
       Do While Not m_rs.EOF
            LongPrintCurCnt = LongPrintCurCnt + 1
            
            strTemp(5) = CheckStr(m_rs.Fields("A1")) '部門
            strTemp(6) = CheckStr(m_rs.Fields("A2")) '姓名
            strTemp(7) = CheckStr(m_rs.Fields("SR02")) '獎懲日期
            strTemp(8) = CheckStr(m_rs.Fields("AC03")) '類別
            strTemp(9) = CheckStr(m_rs.Fields("SR11")) '次數
            strTemp(10) = CheckStr(m_rs.Fields("SR04")) '備註

            PrintDetail '列印表中
            
            If iLine >= 35 Then
                If .AbsolutePosition <> .RecordCount Then
                    Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                End If
            End If
            m_rs.MoveNext
        Loop

    End With

   '列印表尾
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(210, "-")
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If

Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("個人獎懲資料明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "個人獎懲資料明細表"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "獎懲日期：" & ChangeTStringToTDateString(txt1(4)) & " -- " & ChangeTStringToTDateString(txt1(5))


Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部  門"

Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓  名"

Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "獎懲日期"

Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "類  別"

Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "次  數"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "備  註"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(210, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 3500
PLeft(3) = 5500
PLeft(4) = 7000
PLeft(5) = 8000
'PLeft(6) = 8000
Exit Sub
'明細抬頭
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4500
PLeft(4) = 6000
PLeft(5) = 7000
PLeft(6) = 8000
'明細內文
PLeft(7) = 500
PLeft(8) = 2500
PLeft(9) = 4500
PLeft(10) = 6000
PLeft(11) = 7000
PLeft(12) = 8000
End Sub

Sub PrintDetail()
   '部門
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   '姓名
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(6)
   '獎懲日期
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   If Val(Left(strTemp(7), 4)) < 2011 Then  '---民國100年前退後位置
   Printer.Print "  " + ChangeWStringToTDateString(strTemp(7))
   Else
   Printer.Print ChangeWStringToTDateString(strTemp(7))
   End If
   '類別
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(8)
   '次數
   Printer.CurrentX = PLeft(5) '+ (PLeft(12) - PLeft(11)) / 2  '---置中
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(9)
   '備註
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(10)
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
   
'   InitialData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160115 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 4, 5
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         If Index = 4 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 5 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
