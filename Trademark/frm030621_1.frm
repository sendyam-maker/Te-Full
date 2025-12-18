VERSION 5.00
Begin VB.Form frm030621_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "本所案件申請人國籍統計表"
   ClientHeight    =   2565
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4650
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "2"
      Top             =   1650
      Width           =   525
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   2460
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3750
      TabIndex        =   4
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2940
      TabIndex        =   3
      Top             =   90
      Width           =   756
   End
   Begin VB.Label Label4 
      Caption         =   "(1.查詢 2.報表)"
      Height          =   210
      Left            =   2100
      TabIndex        =   7
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "報表種類："
      Height          =   210
      Left            =   570
      TabIndex        =   6
      Top             =   1710
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2100
      X2              =   2790
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Left            =   570
      TabIndex        =   5
      Top             =   1170
      Width           =   900
   End
End
Attribute VB_Name = "frm030621_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2014/2/21
'Modified by Lydia 2015/12/18 更名為frm030621_1
Option Explicit

Dim m_Device
Dim m_rs As New ADODB.Recordset
Dim PLeft(1 To 5) As Integer
Dim strTemp(1 To 5) As String
Dim iLine As Integer
Dim PrinterPages As Integer


Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0
         If Trim(Text1) = "" Then
            MsgBox "起始公報年月不可空白！", vbInformation, "輸入錯誤！"
            Text1.SetFocus
            Exit Sub
         End If
         If Trim(Text2) = "" Then
            MsgBox "截止公報年月不可空白！", vbInformation, "輸入錯誤！"
            Text2.SetFocus
            Exit Sub
         End If
         If Trim(Text7) = "" Then
            MsgBox "報表種類不可空白！", vbInformation, "輸入錯誤！"
            Text7.SetFocus
            Exit Sub
         End If
         Text1_Validate Cancel
         If Cancel = True Then
            Text1.SetFocus
            Exit Sub
         End If
         Text2_Validate Cancel
         If Cancel = True Then
            Text2.SetFocus
            Exit Sub
         End If
         If Val(Text2) < Val(Text1) Then
            MsgBox "截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
            Text2.SetFocus
            Exit Sub
         End If
         
         PrinterPages = 0
         If Text7 = "1" Then '螢幕
            Set m_Device = Picture1
            m_Device.AutoRedraw = True
            m_Device.Width = 11899 '16836
            m_Device.Height = 16838 '11904
            DelPic
         Else '印表機
            Set m_Device = Printer
            '故意設定紙張屬性以便清除印表機狀態(相同印表機驅動程式會沿用原設定值,Ex.進紙槽)
            m_Device.PaperSize = 9
            m_Device.EndDoc
            m_Device.Orientation = 1 '1.直印 2.橫印
         End If
         
         If StrMenu = True Then
            If Text7 = "2" Then '印表機
               m_Device.EndDoc
               MsgBox "列印完畢！", vbExclamation + vbOKOnly
            Else '螢幕
               Screen.MousePointer = vbDefault
               SetPic PrinterPages
               frm090801_3.m_ImageW = m_Device.Width
               frm090801_3.m_ImageH = m_Device.Height
               frm090801_3.m_iPages = PrinterPages
               frm090801_3.cmdOK(0).Visible = False
               frm090801_3.txtPCnt.Visible = False
               frm090801_3.Caption = "本所案件申請人國籍統計表(預覽)"
               frm090801_3.Show vbModal
               Unload frm090801_3
            End If
         End If
         
      Case 1
         Unload Me
   End Select
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   
   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

Private Function StrMenu() As Boolean
Dim i As Integer
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim dblTot1 As Double, dblTot2 As Double, dblTot3 As Double, dblTot4 As Double
   
   StrMenu = False
   
   '報表樣式:
   '申請人國籍           FCT件數    FCT類別數  T件數      T類別數
   '-------------------- ---------- ---------- ---------- ----------
   '日本                         17         40          1          1
   '韓國                          1          1          1          1
   '香港                          0          0         10         10
   '.
   '.
   '.
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & Text1 & "-" & Text2
   pub_QL05 = pub_QL05 & ";" & Label2.Caption & Text7.Text
  'end 2022/01/13
  
   Call ChgDateToTMBM07(Text1, strVol1_S, strVol2_S)
   Call ChgDateToTMBM07(Text2, strVol1_E, strVol2_E)
   'Modified by Lydia 2016/01/19 類別依公報為準(tm09->tmbm08)
   'strSql = "select decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05)," & _
            "sum(nvl(FCTcnt,0)),sum(nvl(FCTclass,0)),sum(nvl(Tcnt,0)),sum(nvl(Tclass,0))" & _
            " from nation,(" & _
            " select TMBM05,count(*) FCTcnt,sum(counting(tm09)) FCTclass,0 Tcnt,0 Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='FCT' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            " Union select TMBM05,0 FCTcnt,0 FCTclass,count(*) Tcnt,sum(counting(tm09)) Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='T' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            ") where TMBM05=na03(+)" & _
            " group by decode(substr(na01,1,1),'A','000','B','002',na01),decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05)" & _
            " order by decode(substr(na01,1,1),'A','000','B','002',na01) asc"
   strSql = "select decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05)," & _
            "sum(nvl(FCTcnt,0)),sum(nvl(FCTclass,0)),sum(nvl(Tcnt,0)),sum(nvl(Tclass,0))" & _
            " from nation,(" & _
            " select TMBM05,count(*) FCTcnt,sum(counting(tmbm08)) FCTclass,0 Tcnt,0 Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='FCT' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            " Union select TMBM05,0 FCTcnt,0 FCTclass,count(*) Tcnt,sum(counting(tmbm08)) Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='T' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            ") where TMBM05=na03(+)" & _
            " group by decode(substr(na01,1,1),'A','000','B','002',na01),decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05)" & _
            " order by decode(substr(na01,1,1),'A','000','B','002',na01) asc"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      Screen.MousePointer = vbHourglass
      With m_rs
         m_rs.MoveFirst
         iLine = 0
         PrinterPages = 1
         Do While Not m_rs.EOF
            For i = 1 To 5
               strTemp(i) = ""
            Next i
            strTemp(1) = Left(CheckStr(m_rs.Fields(0)) & "          ", 10)
            strTemp(2) = m_rs.Fields(1)
            strTemp(3) = m_rs.Fields(2)
            strTemp(4) = m_rs.Fields(3)
            strTemp(5) = m_rs.Fields(4)
            If iLine > 50 Or iLine = 0 Then
               If iLine <> 0 Then
                  PrinterPages = PrinterPages + 1
                  If Text7 = "1" Then '螢幕
                     If PrinterPages > 1 Then
                        SetPic PrinterPages - 1
                     End If
                  Else
                     m_Device.NewPage
                  End If
               End If
               iLine = 1
               PrintTitle
            End If
            PrintDetail '列印表中
            dblTot1 = dblTot1 + Val(strTemp(2))
            dblTot2 = dblTot2 + Val(strTemp(3))
            dblTot3 = dblTot3 + Val(strTemp(4))
            dblTot4 = dblTot4 + Val(strTemp(5))
            m_rs.MoveNext
         Loop
         m_Device.CurrentX = 500
         m_Device.CurrentY = iLine * 300
         m_Device.Print String(140, "-")
         iLine = iLine + 1
         m_Device.CurrentX = PLeft(1)
         m_Device.CurrentY = iLine * 300
         m_Device.Print "        合計"
         m_Device.CurrentX = PLeft(2) - m_Device.TextWidth(dblTot1)
         m_Device.CurrentY = iLine * 300
         m_Device.Print dblTot1
         m_Device.CurrentX = PLeft(3) - m_Device.TextWidth(dblTot2)
         m_Device.CurrentY = iLine * 300
         m_Device.Print dblTot2
         m_Device.CurrentX = PLeft(4) - m_Device.TextWidth(dblTot3)
         m_Device.CurrentY = iLine * 300
         m_Device.Print dblTot3
         m_Device.CurrentX = PLeft(5) - m_Device.TextWidth(dblTot4)
         m_Device.CurrentY = iLine * 300
         m_Device.Print dblTot4
      End With
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      InsertQueryLog (0) 'Added by Lydia 2022/01/13
      Exit Function
   End If
   Screen.MousePointer = vbDefault
End Function

Sub PrintTitle()
GetPleft

m_Device.Font.Size = 16
m_Device.Font.Underline = False
m_Device.FontBold = False

m_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("商標公報 " & Format(Text1, "###/##") & "月 ~ " & Format(Text2, "###/##") & "月 本所案件統計表") / 2)
m_Device.CurrentY = iLine * 300
m_Device.Print "商標公報 " & Format(Text1, "###/##") & "月 ~ " & Format(Text2, "###/##") & "月 本所案件統計表"

m_Device.Font.Size = 12
iLine = iLine + 2
m_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
m_Device.CurrentY = iLine * 300
m_Device.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine = iLine + 1
m_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
m_Device.CurrentY = iLine * 300
m_Device.Print "頁　　次：" & PrinterPages

iLine = iLine + 1
m_Device.CurrentX = PLeft(1)
m_Device.CurrentY = iLine * 300
m_Device.Print "申請人國籍"
m_Device.CurrentX = PLeft(2) - m_Device.TextWidth("FCT 件數")
m_Device.CurrentY = iLine * 300
m_Device.Print "FCT 件數"
m_Device.CurrentX = PLeft(3) - m_Device.TextWidth("FCT 類別數")
m_Device.CurrentY = iLine * 300
m_Device.Print "FCT 類別數"
m_Device.CurrentX = PLeft(4) - m_Device.TextWidth("T 件數")
m_Device.CurrentY = iLine * 300
m_Device.Print "T 件數"
m_Device.CurrentX = PLeft(5) - m_Device.TextWidth("T 類別數")
m_Device.CurrentY = iLine * 300
m_Device.Print "T 類別數"

iLine = iLine + 1
m_Device.CurrentX = 500
m_Device.CurrentY = iLine * 300
m_Device.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 3500
PLeft(3) = 5500
PLeft(4) = 7500
PLeft(5) = 9500
End Sub

Sub PrintDetail()
Dim i As Integer
   For i = 1 To 5
      If i = 1 Then
         m_Device.CurrentX = PLeft(i)
      Else
         m_Device.CurrentX = PLeft(i) - m_Device.TextWidth(strTemp(i))
      End If
      m_Device.CurrentY = iLine * 300
      m_Device.Print strTemp(i)
   Next i
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Text1 = Left(strSrvDate(2), 5)
   Text2 = Left(strSrvDate(2), 5)
End Sub

'將公報年月轉換為卷期 ex.920101 : 3001, 920116 : 3002 ...(每年會有24期)
Private Function ChgDateToTMBM07(strData As String, ByRef strVol1 As String, ByRef strVol2 As String) As String
   ChgDateToTMBM07 = ""
   strData = Right("00000" & Trim(strData), 5)
   If strData = "00000" Then Exit Function
   
   ChgDateToTMBM07 = CStr(Val(Left(strData, 3)) - 62)
   
   Select Case Right(strData, 2)
      Case "01"
         strVol1 = ChgDateToTMBM07 & "01"
         strVol2 = ChgDateToTMBM07 & "02"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "02"
      Case "02"
         strVol1 = ChgDateToTMBM07 & "03"
         strVol2 = ChgDateToTMBM07 & "04"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "04"
      Case "03"
         strVol1 = ChgDateToTMBM07 & "05"
         strVol2 = ChgDateToTMBM07 & "06"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "06"
      Case "04"
         strVol1 = ChgDateToTMBM07 & "07"
         strVol2 = ChgDateToTMBM07 & "08"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "08"
      Case "05"
         strVol1 = ChgDateToTMBM07 & "09"
         strVol2 = ChgDateToTMBM07 & "10"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "10"
      Case "06"
         strVol1 = ChgDateToTMBM07 & "11"
         strVol2 = ChgDateToTMBM07 & "12"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "12"
      Case "07"
         strVol1 = ChgDateToTMBM07 & "13"
         strVol2 = ChgDateToTMBM07 & "14"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "14"
      Case "08"
         strVol1 = ChgDateToTMBM07 & "15"
         strVol2 = ChgDateToTMBM07 & "16"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "16"
      Case "09"
         strVol1 = ChgDateToTMBM07 & "17"
         strVol2 = ChgDateToTMBM07 & "18"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "18"
      Case "10"
         strVol1 = ChgDateToTMBM07 & "19"
         strVol2 = ChgDateToTMBM07 & "20"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "20"
      Case "11"
         strVol1 = ChgDateToTMBM07 & "21"
         strVol2 = ChgDateToTMBM07 & "22"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "22"
      Case "12"
         strVol1 = ChgDateToTMBM07 & "23"
         strVol2 = ChgDateToTMBM07 & "24"
         ChgDateToTMBM07 = ChgDateToTMBM07 & "24"
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Modified by Lydia 2015/12/18
   'Set frm030621 = Nothing
   Set frm030621_1 = Nothing
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer
   
   If Text1 <> "" Then
      If ChkDate(Text1 & "01") = False Then
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call ChgDateToTMBM07(Text1, strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      rsQuery.Close
      If intCnt = 0 Then
         MsgBox Text1 & "此月份尚無公報資料!!"
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      ElseIf intCnt = 0 Then
         MsgBox Text1 & "此月份公報資料尚不足!!"
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   
   Set rsQuery = Nothing
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer
   
   If Text2 <> "" Then
      If ChkDate(Text2 & "01") = False Then
          Call Text2_GotFocus
          Cancel = True
          Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call ChgDateToTMBM07(Text2, strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      rsQuery.Close
      If intCnt = 0 Then
         MsgBox Text2 & "此月份尚無公報資料!!"
         Call Text2_GotFocus
         Cancel = True
         Exit Sub
      ElseIf intCnt = 0 Then
         MsgBox Text2 & "此月份公報資料尚不足!!"
         Call Text2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   
   Set rsQuery = Nothing
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 And KeyAscii <> 23 Then
      KeyAscii = 0
   End If
End Sub
