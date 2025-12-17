VERSION 5.00
Begin VB.Form frm160308 
   BorderStyle     =   1  '單線固定
   Caption         =   "應繳健檢報告清單"
   ClientHeight    =   3200
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3200
   ScaleWidth      =   5040
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2190
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1110
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2820
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1110
      Width           =   495
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   5
      Top             =   2490
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "下次應繳年度："
      Height          =   180
      Left            =   900
      TabIndex        =   7
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   2610
      X2              =   3000
      Y1              =   1230
      Y2              =   1230
   End
End
Attribute VB_Name = "frm160308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by SINDY 2015/8/12
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "起始年度不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(1) = "" Then
            MsgBox "截止年度不可以空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         m_StrSQL = " and (ST68 between " & Val(txt1(0)) + 1911 & " and " & Val(txt1(1)) + 1911 & " or ST68 is null)"
         Call StrMenu
         Screen.MousePointer = vbDefault
      Case 1
           Unload Me
   End Select
   Printer.Font.Size = 12
End Sub

'明細表
Sub StrMenu()
Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
Printer.PaperSize = 9  'PDF

'Modify By Sindy 2023/12/27 部門調整改抓ST93
'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_str = "select a0922,st01,st02,decode(nvl(st23,''),'','',to_char(sysdate,'YYYY')-substr(st23,1,4)),sqldatet(sh02),decode(nvl(st68,0),'','',st68-1911)" & _
           " FROM staff,(select sh01,max(sh02) sh02 from staff_health group by sh01),acc090NEW" & _
           " where substr(st01,1,1) in(" & ST01CodeNum1 & ")" & _
           " and st04='1'" & _
           " and substr(st01,4,1)<>'9' and st01 not in('60000','96029','96030')" & _
           " and st01=sh01(+)" & m_StrSQL & _
           " and a0921=st93(+) and not(substr(st01,5,1)>='A')" & _
           " order by st93,st01 asc"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        
        Do While Not .EOF
            For m_i = 1 To 6
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(.Fields(0))
            strTemp(2) = CheckStr(.Fields(1))
            strTemp(3) = CheckStr(.Fields(2))
            strTemp(4) = CheckStr(.Fields(3))
            strTemp(5) = CheckStr(.Fields(4))
            strTemp(6) = CheckStr(.Fields(5))
            
            If iLine > 53 Or iLine = 1 Then
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            PrintDetail
            
            strType = CheckStr(m_rs.Fields(0))
            .MoveNext
        Loop
    End With
Else
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()

GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

iLine = iLine + 2
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & " ~ " & txt1(1) & "年應繳健檢報告清單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print txt1(0) & " ~ " & txt1(1) & "年應繳健檢報告清單"

iLine = iLine + 2
Printer.Font.Size = 12
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
Printer.Print "部門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工編號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "姓名"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("目前年齡")
Printer.CurrentY = iLine * 300
Printer.Print "目前年齡"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("上次健檢日期")
Printer.CurrentY = iLine * 300
Printer.Print "上次健檢日期"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("下次應繳年度")
Printer.CurrentY = iLine * 300
Printer.Print "下次應繳年度"

iLine = iLine + 1
Printer.CurrentX = 300
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 6500
PLeft(5) = 8500
PLeft(6) = 10500
End Sub

Sub PrintDetail()
Dim i As Integer
   
   For i = 1 To 3
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
   For i = 4 To 6
      Printer.CurrentX = PLeft(i) - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
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
   Set frm160308 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
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
      Case Else
   End Select
End Sub
