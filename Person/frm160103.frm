VERSION 5.00
Begin VB.Form frm160103 
   BorderStyle     =   1  '單線固定
   Caption         =   "父、母親節名條列印"
   ClientHeight    =   3090
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4950
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   11
      Top             =   2460
      Width           =   4875
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
      Height          =   255
      Index           =   4
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   0
      Top             =   900
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2730
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2580
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1260
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1260
      Width           =   465
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3930
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2970
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：列印紙張為單排名條"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1980
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   2700
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   3270
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.父親節 2.母親節)"
      Height          =   180
      Left            =   2310
      TabIndex        =   10
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "名條類別："
      Height          =   180
      Left            =   1110
      TabIndex        =   9
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   8
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1110
      TabIndex        =   7
      Top             =   1290
      Width           =   900
   End
End
Attribute VB_Name = "frm160103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer, i As Integer, j As Integer
Dim m_StrSQL As String
Dim m_str  As String
Dim m_rs As New ADODB.Recordset
Dim m_i As Integer
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim strTempS(1 To 7) As String
Dim iPgae As Integer, iLine As Integer
Dim strPrinter As String 'Add By Sindy 2022/3/14


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
'        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then
'            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
'            txt1(0).SetFocus
'            Exit Sub
'        End If
        If txt1(4) = "" Then
            MsgBox "名條類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(4).SetFocus
            Exit Sub
        End If
        
        '設定印表機
'        Set Printer = Printers(Combo1.ListIndex)
'        Printer.PaperSize = PUB_GetPaperSize(2)
        PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2022/3/14 切換Word/Excel印表機
        PUB_RestorePrinter Combo1 'Add By Sindy 2022/3/14
        
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
        
        Select Case Val(txt1(4))
        Case 1
            m_StrSQL = m_StrSQL & " and sr03='1' and (sr11 is not null) "
        Case 2
            m_StrSQL = m_StrSQL & " and sr03='2' and (sr11 is not null) "
        End Select
        m_StrSQL = m_StrSQL & " and st04='1' "
        StrMenu
        Screen.MousePointer = vbDefault
        
        PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2022/3/14 切換Word/Excel印表機
        PUB_RestorePrinter strPrinter 'Add By Sindy 2022/3/14
Case 1
        Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Dim m_PageNo As Integer
Dim strTempAddressList As String
'Modify By Sindy 2010/5/6 已有刪除日期的資料不顯示
'm_str = "select st01,sr10||' '||sr11,sr04,sr05 from staff,staff_relation,SalaryData where ST01=SD01 and (SD02 not in('P','F') or SD02 is null) and st01=sr01(+) " & m_StrSQL & " order by st01"
'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_str = "select st01,sr10||' '||sr11,sr04,sr05" & _
        " from staff,staff_relation,SalaryData" & _
        " where ST01=SD01 and (SD02 not in('P','F') or SD02 is null)" & _
        " and st01=sr01(+) and (sr12 is null or sr12=0) and sr13 is null " & m_StrSQL & _
        " and not(substr(st01,5,1)>='A')" & _
        " order by st01"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
m_PageNo = 0: strTempAddressList = ""
If Not m_rs.EOF And Not m_rs.BOF Then
'   frm170205.Hide
   With m_rs
      .MoveFirst
      Do While Not .EOF
         For m_i = 1 To 1
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = CheckStr(.Fields("st01"))
         
         'Modify By Sindy 2022/3/14 改寫用Excel
'         frm170205.FormReset
'         frm170205.Text1(0) = strTemp(1)
'         frm170205.Text1(1) = strTemp(1)
'         frm170205.Text1(4) = Trim(txt1(4))
'         frm170205.cmbPrinter = Me.Combo1.Text
'         frm170205.m_bolBeCalled = True
'         frm170205.PrintSheet
         
         '父親、母親(排除已歿)
         'Modify by Morgan 2009/7/7 sr12 改放 'Y'
         'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
         strExc(0) = "select st01,sr04 name,sr10 zipc,sr11 addr" & _
                     " from staff,Staff_Relation" & _
                     " where sr01(+)=st01 and sr03='" & Trim(txt1(4)) & "'" & _
                     " and sr13 is null and st01>='" & strTemp(1) & "' and st01<='" & strTemp(1) & "'" & _
                     " and not(substr(st01,5,1)>='A')"
         '依照部門排序
         strExc(0) = strExc(0) & " order by 1,2,3"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               m_PageNo = m_PageNo + 1
               If strTempAddressList <> "" Then strTempAddressList = strTempAddressList & "|"
               '郵遞區號
               strExc(1) = "" & .Fields("zipc")
               If strExc(1) <> "" Then
                  strTempAddressList = strTempAddressList & strExc(1) & vbCrLf
               End If
               '地址
               strExc(1) = Trim("" & .Fields("addr"))
               If strExc(1) <> "" Then
                  strTempAddressList = strTempAddressList & strExc(1) & "$"
               End If
               '收件人
               Select Case Trim(txt1(4))
                  Case "1"
                     strExc(1) = "" & .Fields("name") & "　　　　　先生　鈞啟"
                  Case "2"
                     strExc(1) = "" & .Fields("name") & "　　　　　女士　鈞啟"
                  Case Else
                     strExc(1) = "" & .Fields("name") & "　　　　　君　鈞啟"
               End Select
               If strExc(1) <> "" Then
                  strTempAddressList = strTempAddressList & strExc(1) & vbCrLf & "~"
               End If
               '員工編號+頁次
               strExc(1) = .Fields("st01")
               strExc(1) = strExc(1) & String(10, "　") & Format(m_PageNo, String(6, "0"))
               strTempAddressList = strTempAddressList & strExc(1)
               
               .MoveNext
            Loop
            End With
         End If
         .MoveNext
      Loop
      'Add By Sindy 2022/3/14 改用Execl列印地址條
      If strTempAddressList <> "" Then
         If PUB_XlsAccAddress(strTempAddressList, 46, False) = False Then
            MsgBox "列印失敗！", vbCritical
            Exit Sub
         End If
      End If
      '2022/03/14 END
   End With
'   Unload frm170205
Else
   ShowNoData
   Exit Sub
End If
'Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

'strSql = Printer.DeviceName
'SeekPrintL = Printer.Orientation
'j = 0
'For i = 0 To Printers.Count - 1
'    Set Printer = Printers(i)
'    Combo1.AddItem Printer.DeviceName, j
'    j = j + 1
'    If Printer.DeviceName = strSql Then
'        SeekPrint = i
'    End If
'Next i
'Set Printer = Printers(SeekPrint)
'Combo1.Text = Combo1.List(0)
PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2022/3/14
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set Printer = Printers(SeekPrint)
'Printer.Orientation = SeekPrintL
Set frm160103 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
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
      Case 4
        If txt1(4) <> "" Then
            Select Case txt1(4)
            Case "1", "2"
            Case Else
                MsgBox "報表類別只可以輸入 1 或 2！", vbInformation, "輸入錯誤！"
                Cancel = True
            End Select
        End If
      Case Else
   End Select
End Sub
