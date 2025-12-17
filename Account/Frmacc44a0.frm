VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44a0 
   AutoRedraw      =   -1  'True
   Caption         =   "部門費用統計表"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   5160
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   612
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1200
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別              (1.台一 2.智權 空白.全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   4245
   End
End
Attribute VB_Name = "Frmacc44a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccrpt411 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Dim dllaccrpt411 As Object
'Added by Lydia 2016/01/30 列印使用
Dim strFieldN(), intWidth()
Dim strTemp(0 To 11) As String
Dim PLeft(0 To 12) As Integer
Dim PTitle(0 To 11) As String
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 10
Private Const ciStartX = 0
Private Const ciStartY = 500
Private Const ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPrint As Integer, iPage As Integer
Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt411Delete
   ProduceData
   '2014/2/20 modify by sonia
   'dllaccrpt411.Acc44a0 ReportTitle(411), Text6, Text7, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'Modified by Lydia 2016/02/15 改成Printer
   'dllaccrpt411.Acc44a0 ReportTitle(411), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
    adoaccrpt411.Open "select * from accrpt411 Where r41101='" & strUserNum & "' order by r41102 ", adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt411.RecordCount <> 0 Then
        PrintData
    End If
    If adoaccrpt411.State = adStateOpen Then
        adoaccrpt411.Close
    End If
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 2200
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt411 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt411 = Nothing
   Set Frmacc44a0 = Nothing
End Sub

Private Sub Text6_Change()
   '2014/2/20 modify by sonia
   'If Text6 = MsgText(601) Then
   '   Exit Sub
   'End If
   'Text7 = A0802Query(Text6)
   Select Case Text6
      Case "1"
         Text7 = A0802Query(Text6)
      Case "2"
         Text7 = A0802Query("J")
      Case ""
         Text7 = "台一　專利商標/智權"
   End Select
   '2014/2/20 end
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim douDebit As Double
Dim intCounter As Integer
Dim strSql As String

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt411.CursorLocation = adUseClient
   adoaccrpt411.Open "select * from accrpt411", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   If adoacc010.RecordCount = 0 Then
      adoacc010.Close
      adoaccrpt411.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc010.EOF = False
      adoaccrpt411.AddNew
      adoaccrpt411.Fields("r41101").Value = strUserNum
      adoaccrpt411.Fields("r41102").Value = adoacc010.Fields("a0101").Value
      If IsNull(adoacc010.Fields("a0102").Value) Then
         adoaccrpt411.Fields("r41103").Value = Null
      Else
         adoaccrpt411.Fields("r41103").Value = adoacc010.Fields("a0102").Value
      End If
      For intCounter = 3 To 11
         adoaccrpt411.Fields(intCounter).Value = 0
      Next intCounter
      adoacc090.CursorLocation = adUseClient
      adoacc090.Open "select * from acc090 where a0904 = 'Y' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc090.EOF = False
         adoacc040.CursorLocation = adUseClient
         strSql = MsgText(601)
         If Text6 <> MsgText(601) Then
            '2014/2/20 modify by sonia
            'strSql = " and a0403 = '" & Text6 & "'"
            strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
            '2014/2/20 end
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
            strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402)  >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
            strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
         End If
         'Added by Lydia 2016/02/15 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
         If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
            If adoacc090.Fields("a0901").Value = "L" Then
               strSql = strSql & " and a0404 in ('" & adoacc090.Fields("a0901").Value & "','FCL','CFL')"
            Else
               strSql = strSql & " and a0404 = '" & adoacc090.Fields("a0901").Value & "'"
            End If
         'end 2016/02/15
         'MODIFY BY SONIA 2013/11/7
         'strSql = strSql & " and a0404 = '" & adoacc090.Fields("a0901").Value & "'"
         ElseIf adoacc090.Fields("a0901").Value = "FCL" Then
            strSql = strSql & " and a0404 in ('" & adoacc090.Fields("a0901").Value & "','CFL')"
         Else
            strSql = strSql & " and a0404 = '" & adoacc090.Fields("a0901").Value & "'"
         End If
         '2013/11/7 end
         
         strSql = strSql & " and a0405 = '" & adoacc010.Fields("a0101").Value & "'"
         If strSql <> MsgText(601) Then
            strSql = Mid(strSql, 5, Len(strSql) - 4)
         End If
         adoacc040.Open "select sum(a0408) from acc040 where" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoacc040.RecordCount <> 0 Then
            If IsNull(adoacc040.Fields(0).Value) Then
               douDebit = 0
            Else
               douDebit = Val(Format(adoacc040.Fields(0).Value, FAmount))
            End If
            Select Case adoacc090.Fields("a0901").Value
               Case "P"
                  adoaccrpt411.Fields(3).Value = douDebit
               Case "T"
                  adoaccrpt411.Fields(4).Value = douDebit
               Case "L"
                  'Modified by Lydia 2016/02/15 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
                  'adoaccrpt411.Fields(5).Value = douDebit
                   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
                      adoaccrpt411.Fields(10).Value = douDebit
                   Else
                      adoaccrpt411.Fields(5).Value = douDebit
                   End If
               Case "CFP"
                  adoaccrpt411.Fields(6).Value = douDebit
               Case "CFT"
                  adoaccrpt411.Fields(7).Value = douDebit
               Case "FCP"
                  adoaccrpt411.Fields(8).Value = douDebit
               Case "FCT"
                  adoaccrpt411.Fields(9).Value = douDebit
               Case "FCL"
                  adoaccrpt411.Fields(10).Value = douDebit
               Case "SAL"
                  adoaccrpt411.Fields(11).Value = douDebit
               Case "TOT"
                  adoaccrpt411.Fields(12).Value = douDebit
               Case "M"
                  adoaccrpt411.Fields(13).Value = douDebit
            End Select
         End If
         adoacc040.Close
         adoacc090.MoveNext
      Loop
      adoacc090.Close
      Select Case Mid(adoacc010.Fields("a0101").Value, 1, 1)
         Case "6"
         Case "9"
            adoaccrpt411.Fields(13).Value = 0
         Case Else
            adoaccrpt411.Fields(13).Value = 0
            For intCounter = 3 To 11
               adoaccrpt411.Fields(13).Value = Val(Format(adoaccrpt411.Fields(13).Value, FAmount)) - Val(Format(adoaccrpt411.Fields(intCounter).Value, FAmount))
            Next intCounter
            adoaccrpt411.Fields(13).Value = Format(Val(Format(adoaccrpt411.Fields(13).Value, FAmount)) + Val(Format(adoaccrpt411.Fields(12).Value, FAmount)), FAmount)
      End Select
'      adoaccrpt411.Fields("r41113").Value = 0
'      For intCounter = 3 To 11
'         If IsNull(adoaccrpt411.Fields(intCounter).Value) = False Then
'            adoaccrpt411.Fields("r41113").Value = Val(adoaccrpt411.Fields("r41113").Value) + Val(adoaccrpt411.Fields(intCounter).Value)
'         End If
'      Next intCounter
      adoaccrpt411.UpdateBatch
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt411.Close
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt411Delete()
   adoTaie.Execute "delete from accrpt411"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text6 = ""
   Text7 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   Text6.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'2014/2/20 add by sonia
Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub
'2014/2/20 end

'Added by Lydia 2016/02/15 從AccReport改成Printer
Private Sub PrintData()
Dim strTot(0 To 11) As String '合計
Dim ii As Integer
    Printer.EndDoc
    Printer.Orientation = 1 '1.直印 2.橫印
    Printer.PaperSize = PUB_GetPaperSize(15) '美國標準
       
    lngPageHeight = Printer.ScaleHeight
    lngPageWidth = Printer.ScaleWidth
    lngLineHeight = 300
           
    iPage = 0
    GetPleft
    Erase strTot
    
    PrintHeader '列印表頭
    With adoaccrpt411
        Do While Not .EOF
        '列印明細
           iPage = iPage + 1
           strTemp(0) = "" & .Fields("R41103") '會計科目
           strTemp(1) = "" & .Fields("R41104") '專利
           strTemp(2) = "" & .Fields("R41105") '商標
           strTemp(3) = "" & .Fields("R41106") '法務->105年以後併入"法務部"
           strTemp(4) = "" & .Fields("R41107") 'CFP
           strTemp(5) = "" & .Fields("R41108") 'CFT
           strTemp(6) = "" & .Fields("R41109") 'FCP
           strTemp(7) = "" & .Fields("R41110") 'FCT
           strTemp(8) = "" & .Fields("R41111") '投法->105年以後"法務部"
           strTemp(9) = "" & .Fields("R41112") '智權部
           strTemp(10) = "" & .Fields("R41114") '總所/管理
           strTemp(11) = "" & .Fields("R41113") '全所
           For ii = 0 To UBound(strFieldN)
              If intWidth(ii) > 0 Then
                '靠左
                If ii = 0 Then
                    Printer.CurrentX = PLeft(ii) + 50
                    Printer.CurrentY = iPrint
                    Printer.Print strTemp(ii)
                '靠右
                Else
                    Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(Format(Val(strTemp(ii)), "###,##0.00")) - ciColGap
                    Printer.CurrentY = iPrint
                    Printer.Print Format(Val(strTemp(ii)), "###,##0.00")
                    strTot(ii) = Val(strTot(ii)) + Val(strTemp(ii))
                End If
              End If
           Next
           PrintNewLine
           
           .MoveNext
        Loop
    End With
    
    For ii = 0 To UBound(strFieldN)
      If intWidth(ii) > 0 Then
        If ii = 0 Then
           PrintLine 1, 1
           Printer.CurrentX = PLeft(1) - 500
           Printer.CurrentY = iPrint
           Printer.Print "合計:"
        Else
           Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(Format(Val(strTot(ii)), "###,##0.00")) - ciColGap
           Printer.CurrentY = iPrint
           Printer.Print Format(Val(strTot(ii)), "###,##0.00")
        End If
      End If
    Next
    PrintNewLine
    PrintLine 1, 2
           
Printer.EndDoc
ShowPrintOk

End Sub


Private Sub GetPleft() '明細表邊界
Dim inX As Integer

'105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "法務部", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 0, 9, 9, 9, 9, 9, 9, 9, 10)
Else
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 10)
End If
   
Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLeft
Erase PTitle
  
   PLeft(0) = ciStartX
   For inX = 1 To UBound(strFieldN)
       If intWidth(inX - 1) = 0 Then
           PLeft(inX) = PLeft(inX - 1)
       Else
           PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(intWidth(inX - 1), "A")) + ciColGap
       End If
   Next
   PLeft(12) = PLeft(11) + Printer.TextWidth(String(10, "A")) + ciColGap
End Sub
Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader

   End If
End Sub

Private Sub PrintHeader()
Dim strPTmp As String
Dim pa1 As Integer
Dim ii As Integer
iPrint = ciStartY
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False

'報表抬頭
strPTmp = ReportTitle(411)
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strPTmp

PrintNewLine
PrintNewLine

Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

strPTmp = "公司別：" & IIf(Text6 = "2", "J", Text6) & " " & IIf(Text7 = "", "台一　專利商標/智權", Text7)
pa1 = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print strPTmp

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & CFDate(strSrvDate(2))

PrintNewLine

Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print "年　月：" & MaskEdBox1.Text & " ～ " & MaskEdBox2.Text

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

PrintNewLine

Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName

PrintNewLine

For ii = 0 To UBound(strFieldN)
   '顯示／欄位
   If intWidth(ii) > 0 And strFieldN(ii) <> "" Then
       strPTmp = strFieldN(ii)
       Printer.CurrentX = PLeft(ii) + ((PLeft(ii + 1) - PLeft(ii) - Printer.TextWidth(strPTmp)) / 2) - ciColGap
       Printer.CurrentY = iPrint
       Printer.Print strPTmp
   End If
Next

PrintNewLine

PrintLine

End Sub

Private Sub PrintLine(Optional aX1 As Integer, Optional aL1 As Integer)
   If aX1 = 0 Then
      Printer.Line (PLeft(0) - 50, iPrint)-(PLeft(12) + 50, iPrint)
   ElseIf aL1 = 2 Then
        Printer.Line (PLeft(1), iPrint)-(PLeft(12) + 50, iPrint)
        iPrint = iPrint + 50
        Printer.Line (PLeft(1), iPrint)-(PLeft(12) + 50, iPrint)
   Else
        Printer.Line (PLeft(1), iPrint)-(PLeft(12) + 50, iPrint)
   End If
   iPrint = iPrint + 150
   
End Sub
'end 2016/02/15

'Added by Lydia 2016/02/17
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox1.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox1.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox1.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox1.SetFocus
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/02/17
Private Sub MaskEdBox2_LostFocus()
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox2.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox2.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox2.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox2.SetFocus
         End If
      End If
   End If
End Sub
