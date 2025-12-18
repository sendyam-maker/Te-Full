VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利申請案保密同意書"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9468
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   9468
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   5220
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   30
      Width           =   920
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_8.frx":0000
      Left            =   6840
      List            =   "frm210114_8.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   2490
      Width           =   2475
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋客戶名稱(&Q)"
      Height          =   330
      Left            =   7320
      TabIndex        =   1
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5544
      TabIndex        =   14
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4632
      TabIndex        =   13
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   30
      TabIndex        =   19
      Top             =   0
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   30
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6456
      TabIndex        =   15
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8460
      TabIndex        =   18
      Top             =   30
      Width           =   920
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7890
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9870
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印       份"
      Height          =   330
      Index           =   0
      Left            =   7368
      TabIndex        =   17
      Top             =   30
      Width           =   1100
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊為必填欄位"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   7320
      TabIndex        =   30
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "＊"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   540
      TabIndex        =   29
      Top             =   653
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "＊"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   720
      TabIndex        =   28
      Top             =   2078
      Width           =   195
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1770
      TabIndex        =   6
      Top             =   2490
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1244;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1770
      TabIndex        =   5
      Top             =   2040
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1770
      TabIndex        =   4
      Top             =   1635
      Width           =   5490
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "9684;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1770
      TabIndex        =   3
      Top             =   1290
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1770
      TabIndex        =   2
      Top             =   960
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1770
      TabIndex        =   0
      Top             =   615
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2820
      TabIndex        =   7
      Top             =   2490
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1244;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   3900
      TabIndex        =   8
      Top             =   2490
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1244;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6090
      TabIndex        =   26
      Top             =   2550
      Width           =   720
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱："
      Height          =   180
      Left            =   780
      TabIndex        =   25
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   960
      TabIndex        =   24
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   960
      TabIndex        =   23
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "ID NO.："
      Height          =   180
      Left            =   990
      TabIndex        =   22
      Top             =   1350
      Width           =   690
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人："
      Height          =   180
      Left            =   960
      TabIndex        =   21
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　　月　　　　　日"
      Height          =   180
      Left            =   390
      TabIndex        =   27
      Top             =   2550
      Width           =   4500
   End
End
Attribute VB_Name = "frm210114_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/04/26 專利申請案保密同意書
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iCount As Integer
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim strNowCustNo As String '客戶編號
Dim iPrintC As Integer '目前列印第幾份
Dim bolAddSeal As Boolean '是否用印
Dim strPrinter As String
Dim strDetail As String '記錄內容
Dim strCompSeal As String '記錄"公司名稱|用印編號",用,區隔
'加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
Dim m_TempPDF As String
Dim oControl As Control
'Word套印: Word需要顯示,不然公司章會偏移
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean

Private Sub cmdFind_Click()
Dim strCmpName As String, strMsg As String

   If Me.txt1(0).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(0).SetFocus
      Exit Sub
   End If
   
   Set frm090801_1.m_frm0908A = Me
   frm090801_1.m_DouChk = False
   
   frm090801_1.m_strCustChnName = Me.txt1(0).Text
   frm090801_1.lblName.Caption = Me.txt1(0).Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   Combo2.Tag = "": strNowCustNo = ""
   If m_blnOneRec = True And m_strCustCode <> "" Then
     '記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "LA", "000", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> frm210114_1.GetComp(Combo2) Then
        strMsg = "您輸入之收據公司別「" & Combo2 & "」與客戶檔設定值「" & strCmpName & "」不同" & vbCrLf & _
                     "是否依客戶檔設定覆蓋您的輸入值？"
        If MsgBox(strMsg, vbYesNo + vbCritical) = vbYes Then
            'Modified by Lydia 2024/08/06
            'Combo2 = strCmpName
            Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
        End If
      ElseIf strCmpName = MsgText(601) Then
        Combo2.ListIndex = 0
      Else
        'Modified by Lydia 2024/08/06
        'Combo2 = strCmpName
         Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
      End If
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
   End If
End Sub


Private Sub cmdOK_Click(Index As Integer)
Dim tb As Control
Dim fN As Integer
Dim strBuffer As String
Dim AllObj(0 To 9) As String
Dim AllObjV As Variant
Dim strNowCmp As String '目前收據公司別
Dim iErr As Integer
   
   iErr = -1
   
   Select Case Index
      Case 0  '列印
      
          If txt1(0) = "" Then
              MsgBox "客戶名稱不可空白！", vbInformation, "錯誤！"
              iErr = 0
              GoTo EXITSUB
          End If
                    
          If Trim(txt1(4)) = "" Then
              MsgBox "聯絡人不可空白！", vbInformation, "錯誤！"
              iErr = 4
              GoTo EXITSUB
          End If
                
         '檢查四縣市地址
         If txt1(3) <> "" Then
           If CheckTaiwanAddr(txt1(3), "000", "地址") = False Then
              iErr = 3
              GoTo EXITSUB
           End If
         End If
         
         If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
         End If
          If Trim(txt1(5)) = "" Or Trim(txt1(6)) = "" Or Trim(txt1(7)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
                iErr = 5
                GoTo EXITSUB
             End If
          End If
          
          If ChkSeal.Value = 1 Then
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If

          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
         
          '畫面與客戶檔收據公司別不同更新客戶檔
          'Mark by Lydia 2022/08/30 此為新功能，請刪除更新客戶檔的程式碼。
          'strNowCmp = frm210114_1.GetComp(Combo2)
          'If Combo2.Tag <> strNowCmp Then
          '   Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          'End If
          'end 2022/08/30
          
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) ' 刪除暫存檔
          If m_TempPDF <> "" Then ShowPrintOk
          
      Case 1  '回前畫面
          frm210114.Show
          Unload Me
      Case 2 '清空資料
          For Each tb In txt1
              tb.Text = Empty
          Next

      Case 3  '儲存文檔
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "專利申請案保密同意書"
              For iCount = 1 To 8
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              AllObj(9) = Combo2.Text '受任人
              
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          '畫面與客戶檔收據公司別不同更新客戶檔
          'Mark by Lydia 2022/08/30 此為新功能，請刪除更新客戶檔的程式碼。
          'strNowCmp = frm210114_1.GetComp(Combo2)
          'If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
          '   Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          'End If
          'end 2022/08/30
          
      Case 4  '讀取文檔
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowOpen
          If cd1.FileName <> "" Then
              fN = FreeFile
              Open cd1.FileName For Input As fN
              Input #fN, strBuffer
              Close #fN
              strBuffer = StrDecrypt(strBuffer)
              AllObjV = Split(strBuffer, Chr(30))
              If AllObjV(0) = "專利申請案保密同意書" Then
                  cmdOK_Click 2
                  'TextBox
                  For iCount = 1 To 8
                    txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  '受任人
                  If AllObjV(9) = MsgText(601) Then
                      Combo2.ListIndex = 0
                  Else
                      Combo2.Text = AllObjV(9)
                  End If
                  
                  '委任人地址=>檢查地址欄
                  If txt1(0).Text <> "" And txt1(3).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(0).Text), Trim(txt1(3).Text), "委任人", True) = False Then
                         iErr = 3
                         GoTo EXITSUB
                     End If
                  End If
                  '讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 專利申請案保密同意書 格式！", vbExclamation
              End If
          End If

      Case 5   '空白委任書
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          '文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If

            bolAddSeal = True
          End If

          Call cmdOK_Click(2) '清空資料
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) ' 刪除暫存檔
          If m_TempPDF <> "" Then ShowPrintOk
      Case Else
   End Select
   Exit Sub
DialogCancel:

EXITSUB:
   If iErr >= 0 Then
       txt1(iErr).SetFocus
       txt1_GotFocus iErr
   End If
End Sub

'只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me '委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True
   
   '設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '還原預設印表機
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Call RunEndProc(False) '刪除暫存檔
   Set frm210114_8 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If

   End If
   '限數字
   If InStr("05,06,07,", Format(Index, "00") & ",") > 0 Then
       If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
           KeyAscii = 0
       End If
   End If
   
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) <> "" Then
       txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
       Cancel = False
       If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
           txt1(Index).SetFocus
           txt1_GotFocus Index
           Cancel = True
           Exit Sub
       End If

       If Index = 6 Then
           If Val(txt1(Index)) > 12 Or Val(txt1(Index)) < 1 Then
               MsgBox "月份輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       ElseIf Index = 7 Then
           If Val(txt1(Index)) > 31 Or Val(txt1(Index)) < 1 Then
               MsgBox "日輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub txtPCnt_GotFocus()
   txtPCnt.SelStart = 0
   txtPCnt.SelLength = Len(txtPCnt)
End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
       KeyAscii = 0
   End If
End Sub

Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(0).Text = "" & rsA("CU04").Value
      '代表人1中文
      Me.txt1(1).Text = "" & rsA("CU07").Value
      'ID. No
      Me.txt1(2).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(3).Text = "" & rsA("CU23").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Mark by Lydia 2022/08/30 此為新功能，請刪除更新客戶檔的程式碼。
'Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
'    Dim strUpd As String
'
'    '同業務區或為MCTF同組人員才可回寫收據公司別
'    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
'
'    strUpd = "Update Customer Set CU164='" & stNowCmp & "' " & _
'                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
'    Pub_SeekTbLog strUpd
'    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
'End Sub
'end 2022/08/30

'下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 13) As String    '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape
Dim oWord

On Error GoTo ErrHand

   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-08 智權部委任契約書_保密同意書.docx", "M51", "000300", "0", "08", "4", "1")

   m_DefPath = App.path & "\" & strUserNum
   Call Pub_ChkExcelPath(m_DefPath)
   
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   
    strDetail = ""
    
   '下載範本檔: M51-000300-0-08 智權部委任契約書_保密同意書.docx
   m_FileName = Pub_RepFileName(IIf(pSpace = False, IIf(Trim(txt1(0)) <> "", Mid(Trim(txt1(0)), 1, 4), Mid(Trim(txt1(1)), 1, 4)), "空白"))
   m_FileName = "$$" & strUserNum & "_保密同意書_" & m_FileName & ".docx"
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-08", , m_DefPath) = False Then
        Exit Sub
   End If
   
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      'g_WordAp.Visible = False
   End If

   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 13
         strName = "PS" & Format(intA, "000")
         strText = ""
         If intA = 1 Then
              '客戶名稱(底線)
              If pSpace = True Then
                 strText = String(40, " ")
              Else
                 strText = "  " & PUB_StrToStr(Trim(txt1(0)), 60) & "  "
              End If
         ElseIf intA = 2 Then
              '受任人(底線)
              strText = "  " & Combo2.Text & "  "
         ElseIf intA = 3 Then
              '甲方：客戶名稱
              strText = PUB_StrToStr(txt1(0) & " ", 60)
         ElseIf intA = 4 Then
              '代表人
              strText = PUB_StrToStr(txt1(1) & " ", 60)
         ElseIf intA = 5 Then
              'ID.NO 統一編號
              strText = PUB_StrToStr(txt1(2) & " ", 60)
         ElseIf intA = 6 Then
              '(著作人)地址
              strText = PUB_StrToStr(txt1(3), 80)
         ElseIf intA = 7 Then
              '乙方：受任人
              strText = Combo2.Text
         ElseIf intA = 8 Then
             strSql = "select a0801,a0802,st02,a0807 from acc080,staff where a0806=st01(+) and a0802='" & Trim(Combo2.Text) & "' "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                 iStr(8) = "" & RsTemp.Fields("st02") '代表人
                 iStr(9) = "" & RsTemp.Fields("a0807") '統一編號
             Else
                 iStr(8) = ""
                 iStr(9) = ""
             End If
             strText = iStr(8)
         ElseIf intA = 9 Then
             strText = iStr(9)
         ElseIf intA = 10 Then
              '地　址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 11 Then
              '聯絡人
              strText = PUB_StrToStr(txt1(4), 60)
         ElseIf intA = 12 Then
              strText = "     中    華    民    國 " & Pub_StrToCenter(txt1(5), 8) & "年" & Pub_StrToCenter(txt1(6), 8) & "月" & Pub_StrToCenter(txt1(7), 8) & "日"
         ElseIf intA = 13 Then  '用印
              strText = ""
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            If intA = 1 Or intA = 2 Then
                '底線
                .Selection.Font.Underline = True
            End If
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If intA = 1 Or intA = 3 Or intA = 4 Or intA = 6 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            
            If intA <> 13 Then  '用印記錄
               iStr(intA) = strText
            End If
            If intA = 13 And bolAddSeal = True Then  '公司章: 放在受任人的儲存格
                strExc(9) = Mid(strCompSeal, InStr(strCompSeal, Combo2))
                If InStr(strExc(9), ",") > 0 Then
                    strExc(9) = Right(Mid(strExc(9), 1, InStr(strExc(9), ",") - 1), 2)
                Else
                    strExc(9) = Right(strExc(9), 2)
                End If
                If PUB_ReadDB2File(m_DefPath & "\$$" & Me.Name & "TempFile", Val(strExc(9))) Then
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=m_DefPath & "\$$" & Me.Name & "TempFile", LinkToFile:=False, SaveWithDocument:=True)
                    '--------設定圖片=文蓋圖(文字在前)
                        oShape.Fill.Visible = msoFalse
                        oShape.Fill.Solid
                        oShape.Fill.Transparency = 0#
                        oShape.Line.Weight = 0.75
                        oShape.Line.DashStyle = msoLineSolid
                        oShape.Line.Style = msoLineSingle
                        oShape.Line.Transparency = 0#
                        oShape.Line.Visible = msoFalse
                        oShape.LockAspectRatio = msoTrue
                        oShape.Rotation = 0#
                        oShape.PictureFormat.Brightness = 0.5
                        oShape.PictureFormat.Contrast = 0.5
                        oShape.PictureFormat.ColorType = msoPictureAutomatic
                        oShape.PictureFormat.CropLeft = 0#
                        oShape.PictureFormat.CropRight = 0#
                        oShape.PictureFormat.CropTop = 0#
                        oShape.PictureFormat.CropBottom = 0#

                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(10.25)
                        'oShape.Top = .CentimetersToPoints(0.3) '不要調Top
                        oShape.LockAnchor = False
                        oShape.LayoutInCell = True
                        oShape.WrapFormat.AllowOverlap = True
                        oShape.WrapFormat.Side = wdWrapBoth
                        oShape.WrapFormat.DistanceTop = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.Type = 3
                        oShape.ZOrder 5 '文蓋圖(文字在前)
                        '---------------------------
                End If
          
            End If
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText

            If intA = 1 Or intA = 2 Then
                '底線
                .Selection.Font.Underline = False
            End If
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If intA = 1 Or intA = 3 Or intA = 4 Or intA = 6 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "標楷體"
            End If
         End If
      Next intA
      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop '還原Word位置
   
   '因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1-2", Collate:=True
   Next intI
   
   '保留: 存檔
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName

   
If bolAddSeal = True Then  '用印記錄
                   strDetail = "保密同意書" & vbCrLf
   strDetail = strDetail & "茲" & iStr(1) & "(以下稱甲方)委請" & iStr(2) & "(以下稱乙方)為專利申請案(包括檢索)，甲方擬將其持有機密資訊告知或交付乙方，乙方瞭解甲方所告知或交付之機密資訊，內含其所擁有之研發成果或技術秘密重要智慧財產權之法定權利或期待利益。為保持所知悉或交付資訊之機密性，乙方同意恪遵本同意書下列各項規定：" & vbCrLf
   strDetail = strDetail & "第一條  所謂「研發成果」係包括專利、著作權、積體電路佈局、營業秘密、電腦軟體、專門技" & vbCrLf & _
                                    "　　　　術(know-how)及其他技術資料等智慧財產權。" & vbCrLf
   strDetail = strDetail & "第二條  所謂「技術秘密」係指與甲方相關無論是否標示「機密」、「限閱」或其他同義字之一" & vbCrLf & _
                                    "　　　　切商業上、技術上或生產上尚未公開之秘密，及依一般商業及法律觀念，應視為機密之" & vbCrLf & _
                                    "　　　　物品、文件及資料等。" & vbCrLf
   strDetail = strDetail & "第三條  甲方所交付之機密資訊，包括但不限於書面、圖樣、電腦或磁碟片檔案、錄音、錄影帶" & vbCrLf & _
                                    "　　　　或光碟片資料檔案等，凡乙方自甲方取得或知悉或接觸的一切資訊均屬之。但下列情形" & vbCrLf & _
                                    "不在此限：" & vbCrLf
   strDetail = strDetail & "　　　　　一、已有書面證據證明甲方所交付或告知之資訊為乙方所已知；" & vbCrLf
   strDetail = strDetail & "　　　　　二、已見於公開發行之刊物或出版品等欠缺機密性質之資訊；" & vbCrLf
   strDetail = strDetail & "　　　　　三、經甲方事先書面同意乙方公開或揭露給第三人之資訊；" & vbCrLf
   strDetail = strDetail & "　　　　　四、乙方從不須承擔任何保密義務及責任的第三人處合法取得者。" & vbCrLf
   strDetail = strDetail & "　　　　　甲方交付予乙方之機密資訊 , 其所有權仍歸甲方所有, 乙方不得任意處分之, 而本同" & vbCrLf & _
                                    "　　　　　意書之簽訂並不構成任何關於機密資訊上所有之專利權、著作權、商標權、電路佈局" & vbCrLf & _
                                    "　　　　　權、營業秘密或其他智慧財產權之明示或暗示之授權或移轉。" & vbCrLf
   strDetail = strDetail & "第四條  乙方保證嚴守保密之義務，非經甲方書面同意，關於甲方為就委託乙方辦理之專利申請" & vbCrLf & _
                                    "　　　　案(包括檢索)所交付之機密資訊之內容負保密義務，絕不以任何方式使其他第三人知悉" & vbCrLf & _
                                    "　　　　或持有任何甲方之機密資訊，更不得自行利用或以任何方式使第三人利用甲方之機密資" & vbCrLf & _
                                    "　　　　訊或取得任何權利，至專利案公開為止。但自交付時起逾3年者，保密義務即自動解除。" & vbCrLf & _
                                    "　　　　如乙方依政府機關合法之命令或法院之判決或命令而須揭露機密資訊時，乙方應於合理" & vbCrLf & _
                                    "　　　　之期間前，以書面通知甲方，並應協助甲方為一切必要之防禦行為及採取任何降低或減" & vbCrLf & _
                                    "　　　　少損害或較為有利之其他保密措施，且僅可依判決或命令所要求最小的範圍內揭露機密" & vbCrLf & _
                                    "　　　　資訊。" & vbCrLf
   strDetail = strDetail & "第五條  乙方同意於本同意書簽署時，完成與其職務作業必須知悉甲方機密資訊的員工及相關人" & vbCrLf & _
                                    "　　　　員簽署保密合約，要求其負擔與乙方相同之保密義務。" & vbCrLf
   strDetail = strDetail & "第六條  乙方如有故意違反本同意書之約定或有因可歸責乙方之事由，致使甲方之機密資訊被洩" & vbCrLf & _
                                    "　　　　露者，乙方除應負擔一切法律責任，並應賠償甲方因此所受之損害。" & vbCrLf
   strDetail = strDetail & "第七條  若非因乙方因素造成，當該等機密資訊對外公開或解除其機密性時，乙方亦同時解除對" & vbCrLf & _
                                    "　　　　該等機密資訊之保密責任。" & vbCrLf
   strDetail = strDetail & "第八條  本同意書之條款，如部份無效或無法執行，不影響其他條款之效力。" & vbCrLf
   strDetail = strDetail & "第九條  因本同意書所生之通知催告，除另有規定外，應以書面對各該當事人於簽立本同意書時" & vbCrLf & _
                                    "　　　　所書寫之地址為通知。地址變更時，應以書面通知變更事宜，否則不得對發通知人主張" & vbCrLf & _
                                    "　　　　通知無效。" & vbCrLf
   strDetail = strDetail & "第十條  本同意書構成甲乙雙方完整之合意，任何未載於本同意書之事項，對雙方皆無拘束力，" & vbCrLf & _
                                    "　　　　同時取代先前所為之書面或口頭討論、通訊、說明。" & vbCrLf
   strDetail = strDetail & "第十一條  凡因本同意書而生之爭議，雙方同意先本誠信原則磋商，磋商不成時，同意以智慧財" & vbCrLf & _
                                    "　　　　產及商業法院為第一審管轄法院。" & vbCrLf
   strDetail = strDetail & "第十二條  本同意書一式二份，甲、乙方各執存一份。" & vbCrLf & vbCrLf
   strDetail = strDetail & "立同意書人:" & vbCrLf
   strDetail = strDetail & "甲方：" & iStr(3) & vbCrLf
   strDetail = strDetail & "代表人：" & iStr(4) & vbCrLf
   strDetail = strDetail & "統一編號：" & iStr(5) & vbCrLf
   strDetail = strDetail & "地址：" & iStr(6) & vbCrLf & vbCrLf
   strDetail = strDetail & "乙方：" & iStr(7) & vbCrLf
   strDetail = strDetail & "代表人：" & iStr(8) & vbCrLf
   strDetail = strDetail & "統一編號：" & iStr(9) & vbCrLf
   strDetail = strDetail & "地址：" & iStr(10) & vbCrLf
   strDetail = strDetail & "聯絡人：" & iStr(11) & vbCrLf
   strDetail = strDetail & "　　" & iStr(12)
   If PUB_AddRecSeal("8", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
   End If
End If
          
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_保密同意書*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub

