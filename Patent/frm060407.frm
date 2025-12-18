VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060407 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯瑕疵案件統計表"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1335
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   285
   End
   Begin VB.TextBox txtNo 
      Height          =   285
      Left            =   1335
      MaxLength       =   6
      TabIndex        =   2
      Top             =   930
      Width           =   750
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1335
      MaxLength       =   5
      TabIndex        =   0
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   1
      Top             =   570
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3675
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   90
      Width           =   756
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   1395
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2010
      X2              =   2270
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label4 
      Caption         =   "(1.統計 2.明細)"
      Height          =   180
      Left            =   1665
      TabIndex        =   9
      Top             =   1365
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "報表內容："
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   8
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "翻譯人員："
      Height          =   180
      Index           =   4
      Left            =   315
      TabIndex        =   7
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "發文年月："
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   615
      Width           =   900
   End
End
Attribute VB_Name = "frm060407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/15 Form2.0已修改
'Created by Lydia 2019/10/28 翻譯瑕疵案件統計表
Option Explicit

Dim rsA1 As New ADODB.Recordset
'列印用
Dim mPrtOrt As Integer, mPrtPage As Integer  '原本預設印表機的列印方向/紙張
Private Const ciTitleFontSize = 16, cInX = 5
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 150
Dim strTitle As String, strTitle2 As String '欄位抬頭/起始位置
Dim ciFontSize As Integer '報表內容字型大小
Dim PLeft(0 To cInX) As Integer '欄位起始位置陣列
Dim PTitle(0 To cInX) As String '欄位抬頭陣列
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPageLine As Integer '頁面資料列
Dim strAcnt As String '人員小計-案件數
Dim strTCnt As String  '總計-案件數
Dim strTmp As String
Dim intLeft As Integer


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
   Case 0
      If TxtValidate() = False Then Exit Sub
      ClearQueryLog (Me.Name) 'Added By Lydia 2023/04/20 清除查詢印表記錄檔欄位
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Process
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Set rsA1 = Nothing
      Printer.Orientation = mPrtOrt
      Printer.PaperSize = mPrtPage
   Case 1
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ''預設系統日之前一月
   txt1(0) = Format(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")), "YYYYMM") - 191100
   txt1(1) = txt1(0)
   mPrtOrt = Printer.Orientation
   mPrtPage = Printer.PaperSize
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060407 = Nothing
End Sub

Private Sub Process()
   Dim strDate1 As String, strDate2 As String, stCon As String
   Dim strVTB As String
      
   strDate1 = (txt1(0) + 191100) & "01"
   strDate2 = (txt1(1) + 191100) & "31"

   stCon = stCon & " and cp27>=" & strDate1 & " and cp27<=" & strDate2
   
   'Added by Lydia 2023/04/20
   pub_QL05 = pub_QL05 & ";" & Me.Caption
   pub_QL05 = pub_QL05 & ";" & Label2(0) & strDate1 & "-" & strDate2
   'end 2023/04/20
   If txtNo <> "" Then
      stCon = stCon & " and cp14='" & txtNo & "'"
      pub_QL05 = pub_QL05 & ";" & Label2(4) & txtNo 'Added by Lydia 2023/04/20
   End If
   
   strVTB = "select fncaseno(cp01,cp02,cp03,cp04) caseno,sqldatet(cp27) cp27t,cp14,st02 as cp14n,tf01,tf37 " & _
               "From caseprogress, transfee, staff " & _
               "where cp01 in ('FCP','P') and cp10='201' and cp158>0 and cp159=0 " & _
               "and cp09=tf01 and nvl(tf37,'N')<>'N' and cp14=st01(+) " & stCon
   
   If txt1(2) = "1" Then '統計
       strSql = "select cp14,cp14n,count(tf01) cnt from (" & strVTB & ") group by cp14,cp14n order by cp14"
   Else   '明細
       strSql = strVTB & " order by cp14, cp27 "
   End If
   Set rsA1 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      InsertQueryLog (rsA1.RecordCount) 'Added by Lydia 2023/04/20
      If txt1(2) = "1" Then
         PrintRpt1 '統計
      Else
         PrintRpt2 '明細
      End If
   Else
      InsertQueryLog (0) 'Added by Lydia 2023/04/20
      MsgBox "無資料！"
   End If
End Sub

Private Function TxtValidate() As Boolean
   
   If Trim(txt1(0)) = "" Then
      MsgBox "請輸起始年月！", vbExclamation
      txt1(0).SetFocus
      Exit Function
   End If
   
   If Trim(txt1(1)) = "" Then
      MsgBox "請輸迄止年月！", vbExclamation
      txt1(1).SetFocus
      Exit Function
   End If
  
   If txtNo <> "" Then
      If Lbl1 = "" Then
         MsgBox "員工編號錯誤！", vbExclamation
         txtNo.SetFocus
         Exit Function
      End If
   End If
   
   If Trim(txt1(2)) = "" Then
      If txt1(2).Enabled Then
         MsgBox "請輸入報表內容！", vbExclamation
         txt1(2).SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      If Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
      
      Select Case Index
         Case 0, 2
            If Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Beep
            End If
         Case 2 '報表內容
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "2" Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

Private Sub txtNo_Change()
   If txtNo <> "" Then
      txt1(2) = "2"
      'Txt1(2).Enabled = False
   Else
      txt1(2) = "1"
      'Txt1(2).Enabled = True
   End If
End Sub

Private Sub txtNo_GotFocus()
   TextInverse txtNo
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_Validate(Cancel As Boolean)
   Lbl1 = GetStaffName(txtNo)
End Sub

'設定印表機
Private Sub SettingPrtSet()
Dim inX As Integer
Dim tmpArr As Variant, tmpArr2 As Variant

    '設定印表機
     Printer.EndDoc
     Printer.PaperSize = 9  'A4
     Printer.Orientation = 1 '1.直印
     
     lngPageHeight = Printer.ScaleHeight
     lngPageWidth = Printer.ScaleWidth
     lngLineHeight = 300
     Printer.Font.Name = "新細明體"
     Printer.Font.Size = ciFontSize
     Erase PLeft
     Erase PTitle
     tmpArr = Empty: tmpArr2 = Empty
     
     '設定欄位抬頭和位置
     If strTitle <> "" And strTitle2 <> "" Then
        tmpArr = Split(strTitle, ",")
        tmpArr2 = Split(strTitle2, ",")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" And Trim(tmpArr2(inX)) <> "" Then
                If Trim(tmpArr(inX)) <> "結束" Then PTitle(inX) = Trim(tmpArr(inX))
                
                If inX < 1 Then
                   PLeft(inX) = ciStartX
                Else
                   PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(Val(tmpArr2(inX)), "　")) + ciColGap
                End If
                
                If Trim(tmpArr(inX)) = "結束" Then
                   intLeft = inX
                   Exit For
                End If
            End If
        Next
     End If
     
     iPage = 0
     
End Sub

'列印-統計表
Private Sub PrintRpt1()
Dim inP As Integer

    strTitle = "翻譯人員,名稱,瑕疵案件數,結束"
    strTitle2 = "0,5,6,6"
    ciFontSize = 12
    
    SettingPrtSet '設定印表機
    With rsA1
       .MoveFirst
       strTCnt = "0"
       strAcnt = "0"
       iPage = iPage + 1
       PrintHeader
       Printer.Font.Size = ciFontSize
       Printer.FontBold = False
       Do While Not .EOF
         strTCnt = Val(strTCnt) + Val("" & .Fields("cnt"))
         '列印內容
         For inP = 0 To cInX
            If PTitle(inP) = "" Then Exit For
            
            If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
               Printer.CurrentX = PLeft(inP)
               Printer.CurrentY = iPrint
               
               Select Case inP
                   Case 0 '翻譯人員
                        Printer.Print "" & .Fields("cp14")
                        
                   Case 1 '翻譯人員名稱
                        Printer.Print PUB_StrToStr("" & .Fields("cp14n"), 12)
                        
                   Case 2 '瑕疵案件數
                        strTmp = Format(Val("" & .Fields("cnt")), "##,##0")
                        Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                        Printer.CurrentY = iPrint
                        Printer.Print strTmp
               End Select
            End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
         Next 'For inP = 1 To cInX

         PrintNewLine
JumpPrint:
          .MoveNext
       Loop

    PrintTotal
    
    End With
    
    Printer.EndDoc
    ShowPrintOk

End Sub

'列印-明細表
Private Sub PrintRpt2()
Dim inP As Integer
Dim strGrp As String

    strTitle = "翻譯人員,名稱,本所案號,發文日,翻譯瑕疵備註,結束"
    strTitle2 = "0,5,6,7,5,20"
    ciFontSize = 12
    
    SettingPrtSet '設定印表機
    With rsA1
       .MoveFirst
       strTCnt = "0"
       strAcnt = "0"
       strGrp = "" & .Fields("cp14")
       iPage = iPage + 1
       PrintHeader
       Printer.Font.Size = ciFontSize
       Printer.FontBold = False
       Do While Not .EOF
         If strGrp <> "" & .Fields("cp14") Then
             PrintSubTotal
             PrintNewLine
             strAcnt = "0"
             strGrp = "" & .Fields("cp14")
         End If
         strAcnt = Val(strAcnt) + 1
         strTCnt = Val(strTCnt) + 1
         '列印內容
         For inP = 0 To cInX
            If PTitle(inP) = "" Then Exit For
            
            If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
               Printer.CurrentX = PLeft(inP)
               Printer.CurrentY = iPrint
               
               Select Case inP
                   Case 0 '翻譯人員
                        Printer.Print "" & .Fields("cp14")
                        
                   Case 1 '翻譯人員名稱
                        Printer.Print PUB_StrToStr("" & .Fields("cp14n"), 12)
                        
                   Case 2 '本所案號
                        Printer.Print PUB_StrToStr("" & .Fields("caseno"), 15)
                        
                   Case 3 '發文日
                        Printer.Print PUB_StrToStr("" & .Fields("cp27t"), 10)

                   Case 4 '翻譯瑕疵備註
                        Printer.Print PUB_StrToStr("" & .Fields("tf37"), 40)
               End Select
            End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
         Next 'For inP = 1 To cInX

         PrintNewLine
JumpPrint:
          .MoveNext
       Loop

    PrintTotal
    
    End With
    
    Printer.EndDoc
    ShowPrintOk

End Sub
'換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub

'列印表頭
Private Sub PrintHeader()
Dim x1 As Integer
Dim x2 As Integer
Dim iPos As Integer

iPrint = ciStartY
iPageLine = 0

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
If txt1(2) = "1" Then
    strTmp = "翻譯瑕疵案件明細表"
Else
    strTmp = "翻譯瑕疵案件統計表"
End If
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strTmp

Printer.Font.Size = ciFontSize
PrintNewLine
PrintNewLine

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName
x1 = Printer.ScaleWidth - Printer.TextWidth(String(12, "　"))
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文年月：" & txt1(0) & " ~ " & txt1(1)
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

PrintNewLine
PrintNewLine

'列印欄位抬頭
For iPos = 0 To cInX
    If PTitle(iPos) <> "" And PTitle(iPos) <> "結束" Then
       If InStr(PTitle(iPos), "數") > 0 Then '置中
          x2 = PLeft(iPos) + (PLeft(iPos + 1) - PLeft(iPos) - Printer.TextWidth(PTitle(iPos))) / 2
       Else
          x2 = PLeft(iPos)
       End If
       Printer.CurrentX = x2 'PLeft(iPos)
       Printer.CurrentY = iPrint
       Printer.Print PTitle(iPos)
    ElseIf iPos > 1 Then
        x1 = iPos '結束
        Exit For
    End If
Next
Printer.Font.Bold = False

PrintNewLine
Printer.Line (PLeft(0), iPrint)-(PLeft(x1), iPrint)
iPrint = iPrint + 150

End Sub

'列印-小計
Private Sub PrintSubTotal()

    Printer.Line (PLeft(2), iPrint)-(PLeft(intLeft), iPrint)
    Printer.Font.Bold = True
    iPrint = iPrint + 150
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "小計:"

    strTmp = Format(Val(strAcnt), "##,##0")
    Printer.CurrentX = PLeft(2) + Printer.TextWidth(String(8, "A")) - Printer.TextWidth(strTmp)
    Printer.CurrentY = iPrint
    Printer.Print strTmp
    Printer.Font.Bold = False
    
End Sub

'列印-總計
Private Sub PrintTotal()

    Printer.Line (PLeft(0), iPrint)-(PLeft(intLeft), iPrint)
    Printer.Font.Bold = True
    iPrint = iPrint + 150
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "總計:"

    strTmp = Format(Val(strTCnt), "##,##0")
    Printer.CurrentX = PLeft(2) + Printer.TextWidth(String(8, "A")) - Printer.TextWidth(strTmp)
    Printer.CurrentY = iPrint
    Printer.Print strTmp
    
    PrintNewLine
    strTmp = "*** 結束 ***"
    Printer.Font.Bold = True
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strTmp
    Printer.Font.Bold = False
End Sub
