VERSION 5.00
Begin VB.Form frm170223 
   BorderStyle     =   1  '單線固定
   Caption         =   "各類所得申報明細"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4752
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   4752
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3690
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2670
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   8
      Top             =   2310
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
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
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   0
      Top             =   900
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1740
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   3120
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1230
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "給付年度："
      Height          =   180
      Left            =   810
      TabIndex        =   7
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "所得人代號："
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   6
      Top             =   1260
      Width           =   1080
   End
   Begin VB.Line Line2 
      X1              =   2790
      X2              =   3450
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frm170223"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by SINDY 2009/01/10
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt As Double, dblAmt2 As Double, dblAmt3 As Double
Dim dblTotAmt As Double, dblTotAmt2 As Double, dblTotAmt3 As Double


Private Sub cmdok_Click(Index As Integer)
Dim strYear As String
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "給付年度不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        strYear = Val(txt1(0)) + 1911
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " AND ID14 = '" & strYear & "' "
        End If
        If txt1(1) <> "" Then
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'm_StrSQL = m_StrSQL & " AND replace(ID25,'A','0') >= '" & txt1(1) & "' "
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_StrSQL = m_StrSQL & " AND substr(id25,1,2)||replace(substr(id25,3,1),'A','0')||substr(id25,4) >= '" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'm_StrSQL = m_StrSQL & " AND replace(ID25,'A','0') <= '" & txt1(2) & "' "
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_StrSQL = m_StrSQL & " AND substr(id25,1,2)||replace(substr(id25,3,1),'A','0')||substr(id25,4) <= '" & txt1(2) & "' "
        End If
        If StrMenu1 = True Then
            StrMenu2
        End If
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub


Function StrMenu1() As Boolean

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
'modify by sonia 2016/1/20 婧瑄說股利資料不帶id17勞退自提
'm_str = "SELECT A0801,A0802,ID25,nvl(OI04,ST02),ID05,ID22||'~'||ID23,nvl(ID08,0),nvl(ID09,0),nvl(substr(rtrim(ID17),0,10),0) " & _
                   "From incomedata, Staff, acc080, otherincomer " & _
             "WHERE ID03=A0807(+) " & _
             "AND ID25=OI01(+) " & _
             "AND substr(id25,1,1)||replace(substr(id25,2),'A','0')=ST01(+) " & m_StrSQL & _
             "Order By A0801,ID25,ID05 "
'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
m_str = "SELECT A0801,A0802,ID25,nvl(OI04,ST02),ID05,ID22||'~'||ID23,nvl(ID08,0),nvl(ID09,0),DECODE(ID05,'54',0,nvl(substr(rtrim(ID17),0,10),0)) " & _
                   "From incomedata, Staff, acc080, otherincomer " & _
             "WHERE ID03=A0807(+) " & _
             "AND ID25=OI01(+) " & _
             "AND substr(id25,1,2)||replace(substr(id25,3,1),'A','0')||substr(id25,4)=ST01(+) " & m_StrSQL & _
             "Order By A0801,ID25,ID05 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        iLine = 1
        strType = "" '切頁條件
        dblAmt = 0
        dblAmt2 = 0
        dblAmt3 = 0
        dblTotAmt = 0
        dblTotAmt2 = 0
        dblTotAmt3 = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 10
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(2))
            strTemp(2) = Left(CheckStr(m_rs.Fields(3)), 9)
            strTemp(3) = CheckStr(m_rs.Fields(4))
            strTemp(4) = CheckStr(m_rs.Fields(5))
            strTemp(5) = CheckStr(m_rs.Fields(6))
            strTemp(6) = CheckStr(m_rs.Fields(7))
            strTemp(7) = CheckStr(m_rs.Fields(8))
            If strTemp(7) = "" Then strTemp(7) = 0
            
            If iLine > 50 Or iLine = 1 Or _
                  (strType <> CheckStr(m_rs.Fields(1))) Then
                
                If (strType <> "" And strType <> CheckStr(m_rs.Fields(1))) Then
                   PrintEnd '小計
                End If
                
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                'End If
            End If
            
            PrintDetail '列印表中
            
            strType = CheckStr(m_rs.Fields(1)) '依公司別跳頁
            dblAmt = dblAmt + strTemp(5)
            dblAmt2 = dblAmt2 + strTemp(6)
            dblAmt3 = dblAmt3 + strTemp(7)
            dblTotAmt = dblTotAmt + strTemp(5)
            dblTotAmt2 = dblTotAmt2 + strTemp(6)
            dblTotAmt3 = dblTotAmt3 + strTemp(7)
            m_rs.MoveNext
        Loop
        
        PrintEnd '小計
        
         '合計
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iLine * 300
         Printer.Print "合　計："
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt, "##,##0")
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTotAmt2, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "##,##0")
         Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblTotAmt3, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt3, "##,##0")
    End With
Else
   StrMenu1 = False
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Function
End If
StrMenu1 = True
Printer.EndDoc
'ShowPrintOk
End Function

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt, "##,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmt2, "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt2, "##,##0")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmt3, "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt3, "##,##0")
   
   dblAmt = 0
   dblAmt2 = 0
   dblAmt3 = 0
   iLine = iLine + 1
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("各 類 所 得 申 報 明 細") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "各 類 所 得 申 報 明 細"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("給付年度：" & txt1(0) & "  年") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "給付年度：" & txt1(0) & "  年"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "公司別：" & CheckStr(m_rs.Fields(0)) & "　" & CheckStr(m_rs.Fields(1))

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "所得人代號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "名　　稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "格式"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "月份"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("所得總額")
Printer.CurrentY = iLine * 300
Printer.Print "所得總額"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("扣繳總額")
Printer.CurrentY = iLine * 300
Printer.Print "扣繳總額"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("勞退自提")
Printer.CurrentY = iLine * 300
Printer.Print "勞退自提"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1800
PLeft(3) = 4000
PLeft(4) = 5000
PLeft(5) = 7500
PLeft(6) = 9000
PLeft(7) = 10500
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
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(5), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(6), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,##0")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(7), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,##0")
   
   iLine = iLine + 1
End Sub


Function StrMenu2() As Boolean

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'modify by sonia 2016/1/20 婧瑄說股利資料不帶id17勞退自提
'm_str = "SELECT A0801,A0802,ID05,sum(nvl(ID08,0)),sum(nvl(ID09,0)),sum(nvl(substr(rtrim(ID17),0,10),0)) " & _
             "From incomedata, acc080 " & _
             "WHERE ID03=A0807(+) " & m_StrSQL & _
             "Group By A0801,A0802,ID05 " & _
             "Order By A0801,ID05 "
m_str = "SELECT A0801,A0802,ID05,sum(nvl(ID08,0)),sum(nvl(ID09,0)),sum(DECODE(ID05,'54',0,nvl(substr(rtrim(ID17),0,10),0))) " & _
             "From incomedata, acc080 " & _
             "WHERE ID03=A0807(+) " & m_StrSQL & _
             "Group By A0801,A0802,ID05 " & _
             "Order By A0801,ID05 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        iLine = 1
        strType = "" '切頁條件
'        dblAmt = 0
'        dblAmt2 = 0
'        dblAmt3 = 0
        dblTotAmt = 0
        dblTotAmt2 = 0
        dblTotAmt3 = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 10
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0)) & "　" & CheckStr(m_rs.Fields(1))
            strTemp(2) = CheckStr(m_rs.Fields(2))
            strTemp(3) = CheckStr(m_rs.Fields(3))
            strTemp(4) = CheckStr(m_rs.Fields(4))
            strTemp(5) = CheckStr(m_rs.Fields(5))
            If strTemp(5) = "" Then strTemp(5) = 0
            
            If iLine > 50 Or iLine = 1 Then
'                If (strType <> "" And strType <> CheckStr(m_rs.Fields(1))) Then
'                   PrintEnd2 '小計
'                End If
                
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle2 '列印表頭
                'End If
            End If
            
            PrintDetail2 '列印表中
            
            strType = CheckStr(m_rs.Fields(1))
'            dblAmt = dblAmt + strTemp(3)
'            dblAmt2 = dblAmt2 + strTemp(4)
'            dblAmt3 = dblAmt3 + strTemp(5)
            dblTotAmt = dblTotAmt + strTemp(3)
            dblTotAmt2 = dblTotAmt2 + strTemp(4)
            dblTotAmt3 = dblTotAmt3 + strTemp(5)
            m_rs.MoveNext
        Loop
                
         '合計
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iLine * 300
         Printer.Print "合　計："
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTotAmt, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt, "##,##0")
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTotAmt2, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "##,##0")
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt3, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt3, "##,##0")
    End With
Else
   StrMenu2 = False
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Function
End If
StrMenu2 = True
Printer.EndDoc
ShowPrintOk
End Function

Sub PrintTitle2()
GetPleft2

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("各 類 所 得 申 報 明 細") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "各 類 所 得 申 報 明 細"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("給付年度：" & txt1(0) & "  年") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "給付年度：" & txt1(0) & "  年"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "公司別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "格式"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("所得總額")
Printer.CurrentY = iLine * 300
Printer.Print "所得總額"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("扣繳總額")
Printer.CurrentY = iLine * 300
Printer.Print "扣繳總額"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("勞退自提")
Printer.CurrentY = iLine * 300
Printer.Print "勞退自提"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft2()
PLeft(1) = 500
PLeft(2) = 5000
PLeft(3) = 7500
PLeft(4) = 9000
PLeft(5) = 10500
End Sub

Sub PrintDetail2()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(3), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(3), "##,##0")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(4), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##,##0")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(5), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,##0")
   
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
   Set frm170223 = Nothing
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
