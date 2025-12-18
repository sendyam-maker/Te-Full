VERSION 5.00
Begin VB.Form frm04060204 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸公報開拓函列印"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   3270
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4635
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   660
      Left            =   336
      TabIndex        =   4
      Top             =   1032
      Width           =   3732
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   240
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   264
         Width           =   540
      End
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2868
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3696
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox text01 
      Height          =   264
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公告日："
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "frm04060204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim Prn As Printer

Private Sub buttonOK_Click()
Dim strTmp(1 To 2) As String, strTxt(1 To 5) As String, rsTemp1 As New ADODB.Recordset
Dim i As Integer
   
   If CheckDataValid = True Then
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
      pub_QL05 = pub_QL05 & ";" & Label1 & text01 'Add By Sindy 2010/12/2
      strExc(0) = "SELECT CPB01 FROM CPBULLETIN WHERE CPB03='" & TransDate(text01, 2) & "' AND CPB09 IS NOT NULL ORDER BY CPB01"
      intI = 0
      Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         InsertQueryLog (rsTemp1.RecordCount) 'Add By Sindy 2010/12/2
         Screen.MousePointer = vbHourglass
         Do While Not rsTemp1.EOF
            Select Case Mid(rsTemp1.Fields(0), 3, 1)
               Case "1"
                  strTmp(1) = "發明"
                  strTmp(2) = "公開"
               Case "2"
                  strTmp(1) = "實用新型"
                  strTmp(2) = "公告"
               Case "3"
                  strTmp(1) = "外觀設計"
                  strTmp(2) = "公告"
            End Select

            EndLetter "17", "@", "00", strUserNum
            strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('17','@','00','" & strUserNum & _
               "','大陸公報專利種類','" & strTmp(1) & "')"
            strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('17','@','00','" & strUserNum & _
               "','大陸公告型式','" & strTmp(2) & "')"
            'edit by nickc 2007/02/05 不用 dll 了
            'If Not objLawDll.ExecSQL(2, strTxt) Then
            If Not ClsLawExecSQL(2, strTxt) Then
               MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
            End If
            NowPrint "@", "17", "00", False, strUserNum, rsTemp1.Fields(0)
            rsTemp1.MoveNext
         Loop
         Screen.MousePointer = vbDefault
         
         MsgBox "準備列印地址條 !", vbInformation
         
         Screen.MousePointer = vbHourglass
         For Each Prn In Printers
            If Prn.DeviceName = Combo2.Text Then
               Set Printer = Prn
               Exit For
            End If
         Next
         
         'Modify by Morgan 2008/4/9 9x才能自訂
         '9x
         If pub_OS = "1" Then
            Printer.Height = 2200
            Printer.Width = 4500
         'NT
         Else
            Printer.Orientation = 1
            Printer.EndDoc
         End If
         'end 2008/4/9
         
         Printer.Orientation = 1
         Printer.Font.Size = 12
         
         strExc(0) = "SELECT CPB08,CPB09 FROM CPBULLETIN WHERE CPB03='" & TransDate(text01, 2) & "' AND CPB09 IS NOT NULL ORDER BY CPB01"
         intI = 0
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Screen.MousePointer = vbHourglass
            i = 1
            Do While Not rsTemp1.EOF
               If Not IsNull(rsTemp1.Fields("CPB08")) And Not IsNull(rsTemp1.Fields("CPB09")) Then
                  PrtAddress rsTemp1.Fields("CPB08"), rsTemp1.Fields("CPB09"), i
                  i = i + 1
               End If
               rsTemp1.MoveNext
            Loop
         End If
         
         MsgBox "地址條列印結束 !", vbInformation
         
         Screen.MousePointer = vbDefault
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/2
      End If
   End If
End Sub

Private Sub PrtAddress(strName As String, strAddress As String, iCount As Integer)
 Dim i As Integer, j As Integer, k As Integer, lPos As Integer
 Dim iLeft(1 To 3) As Integer
 Dim iTxtWidth As Integer
 Dim iTop As Integer
   iLeft(1) = 220
   iLeft(2) = 2600
   iTop = 350
   iTxtWidth = 4200
   
   If Printer.TextWidth(strAddress) > iTxtWidth Then
      strExc(0) = ""
      j = 0
      lPos = 0
      For i = 1 To Len(strAddress)
         strExc(0) = strExc(0) & Mid(strAddress, i, 1)
         If Printer.TextWidth(strExc(0)) > iTxtWidth Then
            Printer.CurrentX = iLeft(1)
            Printer.CurrentY = iTop + j * 300
            Printer.Print strExc(0)
            strExc(0) = ""
            lPos = i
            j = j + 1
         End If
      Next
      
      strExc(0) = ""
      For i = lPos + 1 To Len(strAddress)
         strExc(0) = strExc(0) & Mid(strAddress, i, 1)
      Next
      Printer.CurrentX = iLeft(1)
      Printer.CurrentY = iTop + j * 300
      Printer.Print strExc(0)
   
   Else
      Printer.CurrentX = iLeft(1)
      Printer.CurrentY = iTop
      Printer.Print strAddress
   End If
   
   Printer.CurrentX = iLeft(1)
   Printer.CurrentY = 1300
   Printer.Print strName
   
   Printer.CurrentX = 3400
   Printer.CurrentY = 1600
   Printer.Print "君　啟　" & iCount
   
   Printer.EndDoc
End Sub

Private Sub Form_Load()
 Dim i As Integer, j As Integer
   MoveFormToCenter Me
   strExc(0) = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   j = 0
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      If Printer.DeviceName = strExc(0) Then
         SeekPrint = i
      Else
         Combo2.AddItem Printer.DeviceName, j
         j = j + 1
      End If
   Next i
   If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
   
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   Set frm04060204 = Nothing
End Sub

Private Sub text01_GotFocus()
  TextInverse text01
End Sub

Private Sub text01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text01) = False Then
      If CheckIsTaiwanDate(text01, False) = False Then
         Cancel = True
         strTit = "檢核輸入"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 判斷資料是否為空的
Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   CheckDataValid = True
   
   CheckDataValid = False
   If IsEmpty(text01) = True Then
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   'Add By Cheng 2002/03/19
   Else
      If PUB_CheckKeyInDate(Me.text01) = -1 Then
         Me.text01.SetFocus
         text01_GotFocus
         GoTo EXITSUB
      End If
   End If
   CheckDataValid = True
EXITSUB:
End Function

