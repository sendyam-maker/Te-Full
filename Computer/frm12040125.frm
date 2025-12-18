VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040125 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人改號作業"
   ClientHeight    =   4920
   ClientLeft      =   792
   ClientTop       =   1392
   ClientWidth     =   6828
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6828
   Begin VB.CommandButton CmdNoSaveExit 
      Cancel          =   -1  'True
      Caption         =   "不更新離開"
      Height          =   500
      Left            =   5310
      TabIndex        =   23
      Top             =   120
      Width           =   1300
   End
   Begin VB.TextBox TextMerge 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   5
      Top             =   2580
      Width           =   495
   End
   Begin VB.TextBox TextDelete 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   4
      Top             =   2190
      Width           =   495
   End
   Begin VB.TextBox TextUpdate2 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1860
      Width           =   495
   End
   Begin VB.TextBox textUpdate1 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1500
      Width           =   495
   End
   Begin VB.TextBox textStatus 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3288
      Width           =   6372
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   276
      ItemData        =   "frm12040125.frx":0000
      Left            =   1290
      List            =   "frm12040125.frx":0002
      TabIndex        =   6
      Top             =   2940
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   500
      Left            =   2952
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "更新並結束(&X)"
      Height          =   500
      Left            =   3735
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox textNewNum 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   1
      Top             =   1152
      Width           =   1215
   End
   Begin VB.TextBox textOldNum 
      Height          =   264
      Left            =   2310
      MaxLength       =   9
      TabIndex        =   0
      Top             =   792
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "PS : 客戶改號時,Trigger會清除個案接洽人!!"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   216
      TabIndex        =   24
      Top             =   4680
      Width           =   6372
   End
   Begin MSForms.TextBox textOldNum_2 
      Height          =   300
      Left            =   3600
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   792
      Width           =   3012
      VariousPropertyBits=   671105055
      Size            =   "5313;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNewNum_2 
      Height          =   300
      Left            =   3600
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1152
      Width           =   3012
      VariousPropertyBits=   671105055
      Size            =   "5313;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      Caption         =   "客戶併號後，請檢查聯絡人檔。"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   540
      TabIndex        =   20
      Top             =   4428
      Width           =   2688
   End
   Begin VB.Label Label9 
      Caption         =   "接洽人是否併入新編號:                       (Y:併入)"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2580
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "PS : 代理人併號錯誤時，可能是彼所案號異動紀錄資料在案件基本檔之Trigger所致 !"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   240
      TabIndex        =   18
      Top             =   4152
      Width           =   6552
   End
   Begin VB.Label Label8 
      Caption         =   "PS : 改號時, 先檢查二編號聯絡人、發明人是否重覆？再加潛在客戶之改編號"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   3912
      Width           =   6372
   End
   Begin VB.Label Label7 
      Caption         =   "原編號基本資料是否刪除:                   (Y: 刪除)"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2250
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "更新國外未收未付資料:                       (Y: 更新)"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "歷史資料是否更新:                               (Y: 更新)"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "PS : 結束時才列印改號記錄表"
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   3660
      Width           =   3732
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   2940
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "客戶或代理人新編號 :"
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   1152
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "客戶或代理人原編號 :"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   792
      Width           =   1932
   End
End
Attribute VB_Name = "frm12040125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/10/24 接洽單申請人,代理人資料是獨立的資料檔,不更名
'                               發明人資料是獨立的資料檔,不更名
'Memo By Sonia 2021/12/10 Form2.0已修改(textOldNum_2,textNewNum_2)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
' 記錄修改的資料
Private Type MODIFIEDITEM
   OldNum As String
   NewNum As String
    'Add By Cheng 2004/03/16
    ModHistory As String '歷史資料是否更新
    ModNeverClose As String '更新國外未收未付資料
    DelOldNum As String '原編號基本資料是否刪除
    ModMerge As String 'Added by Lydia 2020/06/10 接洽人是否併入新編號
    'End
   Name As String
   ItemList() As String
   ItemListCount As Long
End Type
Dim m_ModifiedList() As MODIFIEDITEM
Dim m_ModifiedListCount As Integer
Const m_CharWidth = 120
Const m_CharHeight = 240
Const m_PaperSize = "A4"
' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(8) As REPORTFIELD
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
' 報表頁首的高度
Dim m_HeaderHeight As Integer
' 報表文件的寬度
Dim m_ReportWidth As Integer
' 報表文件中可容納的資料列數
Dim m_ReportDataRows As Integer
' 預設印表機
Dim m_DefaultPrinter As String
' 記錄列印的狀態
Dim m_PageNo As Integer
Dim m_CurrRow As Integer
Dim strNewFagent As Boolean   '2010/10/13 ADD BY SONIA若是改號時CP139也要更新
Dim m_bolReNo As Boolean 'Added by Lydia 2025/01/13 是否為更名前後的編號(前8碼相同)
Dim strDelOrgNo As String, strOldNumXYSData(2) As String, strNewNumXYSData(2) As String, intType As Integer, bolData As Boolean 'Add by Amy 2025/07/03 從ChkHasXYSData搬過來

' 加入一個本所案號
'2008/3/26 modify by sonia
'Private Sub InsertItem(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strItem As String)
Private Sub InsertItem(ByVal strOldNum As String, ByVal strNewNum As String, ByVal CaseNo1 As String, ByVal CaseNo2 As String, ByVal CaseNo3 As String, ByVal CaseNo4 As String)
   Dim nIndex As Integer
   Dim nPos As Integer
   Dim bFind As Boolean
   '2008/3/26 ADD BY SONIA 北所已銷卷者加印●
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strItem As String
   
   Select Case CaseNo1
      Case "P", "CFP", "FCP"
         strSql = "SELECT PA108 FROM PATENT " & _
                  "WHERE PA01 = '" & CaseNo1 & "' AND " & _
                        "PA02 = '" & CaseNo2 & "' AND " & _
                        "PA03 = '" & CaseNo3 & "' AND " & _
                        "PA04 = '" & CaseNo4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         strItem = CaseNo1 & "-" & CaseNo2 & "-" & CaseNo3 & "-" & CaseNo4
         If rsTmp.RecordCount > 0 Then
            If rsTmp.Fields("PA108") <> "" Then strItem = strItem & "●"
         End If
      Case "T", "CFT", "FCT", "TF"
         strSql = "SELECT TM57 FROM TRADEMARK " & _
                  "WHERE TM01 = '" & CaseNo1 & "' AND " & _
                        "TM02 = '" & CaseNo2 & "' AND " & _
                        "TM03 = '" & CaseNo3 & "' AND " & _
                        "TM04 = '" & CaseNo4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If CaseNo1 = "TF" Then
            strItem = CaseNo1 & "-" & Mid(CaseNo2, 1, 5) & "-" & Mid(CaseNo2, 6, 1) & "-" & CaseNo3 & "-" & CaseNo4
         Else
            strItem = CaseNo1 & "-" & CaseNo2 & "-" & CaseNo3 & "-" & CaseNo4
         End If
         If rsTmp.RecordCount > 0 Then
            If rsTmp.Fields("TM57") <> "" Then strItem = strItem & "●"
         End If
      Case "L", "CFL", "FCL"
         strSql = "SELECT LC34 FROM LAWCASE " & _
                  "WHERE LC01 = '" & CaseNo1 & "' AND " & _
                        "LC02 = '" & CaseNo2 & "' AND " & _
                        "LC03 = '" & CaseNo3 & "' AND " & _
                        "LC04 = '" & CaseNo4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         strItem = CaseNo1 & "-" & CaseNo2 & "-" & CaseNo3 & "-" & CaseNo4
         If rsTmp.RecordCount > 0 Then
            If rsTmp.Fields("LC34") <> "" Then strItem = strItem & "●"
         End If
      Case "LA"
         strSql = "SELECT HC19 FROM HIRECASE " & _
                  "WHERE HC01 = '" & CaseNo1 & "' AND " & _
                        "HC02 = '" & CaseNo2 & "' AND " & _
                        "HC03 = '" & CaseNo3 & "' AND " & _
                        "HC04 = '" & CaseNo4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         strItem = CaseNo1 & "-" & CaseNo2 & "-" & CaseNo3 & "-" & CaseNo4
         If rsTmp.RecordCount > 0 Then
            If rsTmp.Fields("HC19") <> "" Then strItem = strItem & "●"
         End If
      Case Else
         strSql = "SELECT SP61 FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & CaseNo1 & "' AND " & _
                        "SP02 = '" & CaseNo2 & "' AND " & _
                        "SP03 = '" & CaseNo3 & "' AND " & _
                        "SP04 = '" & CaseNo4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         strItem = CaseNo1 & "-" & CaseNo2 & "-" & CaseNo3 & "-" & CaseNo4
         If rsTmp.RecordCount > 0 Then
            If rsTmp.Fields("SP61") <> "" Then strItem = strItem & "●"
         End If
   End Select
   '2008/3/26 END
   For nIndex = 0 To m_ModifiedListCount - 1
      If m_ModifiedList(nIndex).OldNum = strOldNum And m_ModifiedList(nIndex).NewNum = strNewNum Then
         bFind = False
         For nPos = 0 To m_ModifiedList(nIndex).ItemListCount - 1
            If m_ModifiedList(nIndex).ItemList(nPos) = strItem Then
               bFind = True
               Exit For
            End If
         Next nPos
         If bFind = False Then
            ReDim Preserve m_ModifiedList(nIndex).ItemList(m_ModifiedList(nIndex).ItemListCount + 1)
            m_ModifiedList(nIndex).ItemList(m_ModifiedList(nIndex).ItemListCount) = strItem
            m_ModifiedList(nIndex).ItemListCount = m_ModifiedList(nIndex).ItemListCount + 1
         End If
         Exit For
      End If
   Next nIndex
End Sub

'不更新離開
Private Sub CmdNoSaveExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
   
   textOldNum_2.BackColor = &H8000000F
   textNewNum_2.BackColor = &H8000000F
   textStatus.BackColor = &H8000000F
   TextMerge.Text = "Y" 'Added by Lydia 2019/05/28 預設接洽人要合併
   
   m_ModifiedListCount = 0
   
   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
   nSel = 0
   nIndex = 0
   For Each Prn In Printers
      cboPrinter.AddItem Prn.DeviceName
      If Prn.DeviceName = m_DefaultPrinter Then
         nSel = nIndex
      End If
      nIndex = nIndex + 1
   Next
   cboPrinter.ListIndex = nSel
   
   '2024/10/14 add by sonia
   MsgBox "併號要先檢查發明人資料 !"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim nIndex As Integer
   Dim Prn As Printer
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   ' 清除修改明細的暫存區
   If m_ModifiedListCount > 0 Then
      For nIndex = 0 To m_ModifiedListCount - 1
         If m_ModifiedList(nIndex).ItemListCount > 0 Then
            Erase m_ModifiedList(nIndex).ItemList
         End If
         m_ModifiedList(nIndex).ItemListCount = 0
      Next nIndex
      Erase m_ModifiedList
   End If
   m_ModifiedListCount = 0
   'Add By Cheng 2002/07/18
   Set frm12040125 = Nothing
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   '911030 nick
   Dim nick911030rs As New ADODB.Recordset
   Dim nickstrsql As String
   Dim strCName As String 'Added by Lydia 2019/05/28
   Dim strNo As String, strMail As String, strArrMail 'Add by Amy 2022/10/05
   
   strDelOrgNo = "" 'Add by Amy 2025/07/03
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      'Add by Amy 2022/10/04 商申承辦人責任業務區分配人員確認
      strMsg = ""
      strNo = textOldNum & String(9 - Len(textOldNum), "0")
      'Modify by Amy 2022/10/17 +Mid(strNo, 1, 1) = "X" 客戶編號才檢查
      If Mid(strNo, 1, 1) = "X" And Right(strNo, 1) = "0" Then
        If ChkDutyZoneAssign(Me.Name, Left(strNo, 8), False, False, , strMail) = True Then
            If strMail <> MsgText(601) Then
                '詢問是否發Mail
                strArrMail = Split(strMail, ";")
                strMsg = "編號 " & textOldNum & " 已設「商申承辦人責任業務區分配」" & vbCrLf & _
                            "無法改號，需待" & GetPrjSalesNM("" & strArrMail(0)) & "修改後回覆才可繼續執行" & vbCrLf & _
                            "要發Mail通知「" & GetPrjSalesNM("" & strArrMail(0)) & "」修改？"
                If MsgBox(strMsg, vbYesNo + vbCritical + vbQuestion, "確認") = vbYes Then
                    PUB_SendMail strUserNum, strArrMail(0), "", strArrMail(1), strArrMail(2), , , , , , , , , , True, , , , False
                End If
            End If
            textOldNum.SetFocus
            Exit Sub
        End If
      End If
      'end 2022/10/05
   
      ReDim Preserve m_ModifiedList(m_ModifiedListCount + 1)
      m_ModifiedList(m_ModifiedListCount).OldNum = textOldNum & String(9 - Len(textOldNum), "0")
      m_ModifiedList(m_ModifiedListCount).NewNum = textNewNum & String(9 - Len(textNewNum), "0")
        'Add By Cheng 2004/03/16
      m_ModifiedList(m_ModifiedListCount).ModHistory = Me.textUpdate1.Text
      m_ModifiedList(m_ModifiedListCount).ModNeverClose = Me.TextUpdate2.Text
      m_ModifiedList(m_ModifiedListCount).DelOldNum = Me.textDelete.Text
        'End
      m_ModifiedList(m_ModifiedListCount).ModMerge = Me.TextMerge.Text 'Added by Lydia 2020/06/10
      m_ModifiedList(m_ModifiedListCount).Name = Empty
      m_ModifiedList(m_ModifiedListCount).ItemListCount = 0
      If Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 1, 1) = "X" Then
         '911030 nick 邱小姐檢查舊的編號若存在已中英日順序擇一，若不存在已新的編號，中英日順序擇一
         '***** start
         'm_ModifiedList(m_ModifiedListCount).Name = GetCustomerName(m_ModifiedList(m_ModifiedListCount).OldNum, 0)
         Set nick911030rs = New ADODB.Recordset
         nickstrsql = "select * from customer where cu01='" & Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 1, 8) & "' and cu02='" & Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 9, 1) & "' "
         nick911030rs.CursorLocation = adUseClient
         nick911030rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
         If nick911030rs.RecordCount <> 0 Then
            m_ModifiedList(m_ModifiedListCount).Name = GetPrjPeople1(m_ModifiedList(m_ModifiedListCount).OldNum)
         Else
            m_ModifiedList(m_ModifiedListCount).Name = GetPrjPeople1(m_ModifiedList(m_ModifiedListCount).NewNum)
         End If
         '***** end
      'Modified by Lydia 2019/05/28
      'Else
      ElseIf Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 1, 1) = "Y" Then '代理人
         '911030 nick 邱小姐檢查舊的編號若存在已中英日順序擇一，若不存在已新的編號，中英日順序擇一
         '***** start
         'm_ModifiedList(m_ModifiedListCount).Name = GetFAgentName(m_ModifiedList(m_ModifiedListCount).OldNum)
         Set nick911030rs = New ADODB.Recordset
         nickstrsql = "select * from fagent where fa01='" & Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 1, 8) & "' and fa02='" & Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 9, 1) & "' "
         nick911030rs.CursorLocation = adUseClient
         nick911030rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
         If nick911030rs.RecordCount <> 0 Then
            m_ModifiedList(m_ModifiedListCount).Name = GetPrjName1(m_ModifiedList(m_ModifiedListCount).OldNum)
         Else
            m_ModifiedList(m_ModifiedListCount).Name = GetPrjName1(m_ModifiedList(m_ModifiedListCount).NewNum)
         End If
         '***** end
      'Added by Lydia 2019/05/28 潛在客戶 (R編號 轉 R編號)
      ElseIf Mid(m_ModifiedList(m_ModifiedListCount).OldNum, 1, 1) = "R" Then
         strCName = GetPotCustName(m_ModifiedList(m_ModifiedListCount).OldNum, strTit)
         If strCName <> "" Then
            m_ModifiedList(m_ModifiedListCount).Name = strCName
         Else
             strCName = GetPotCustName(m_ModifiedList(m_ModifiedListCount).NewNum, strTit)
             m_ModifiedList(m_ModifiedListCount).Name = strCName
         End If
      End If
      m_ModifiedListCount = m_ModifiedListCount + 1
      
      strTit = "資料處理"
      strMsg = "資料檢查完畢"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)

      ' 清除欄位
      textOldNum = Empty
      textOldNum_2 = Empty
      textNewNum = Empty
      textNewNum_2 = Empty
        Me.textUpdate1.Text = ""
        Me.TextUpdate2.Text = ""
        Me.textDelete.Text = ""
      ' 輸入欄位
      textOldNum.SetFocus
      
      TextMerge.Text = "Y" 'Added by Lydia 2019/05/28 預設接洽人要合併
      strDelOrgNo = textDelete 'Add by Amy 2025/07/03 記錄畫面「原編號基本資料是否刪除」,因按「確定」鈕資料會刪
   End If
End Sub

'更新並結束
Private Sub cmdExit_Click()
   Dim Prn As Printer
   
   '有要需執行改號的資料
   If m_ModifiedListCount > 0 Then
      'Add by Amy 2022/10/05
      If MsgBox("要執行改號作業？", vbYesNo + vbCritical + vbQuestion, "確認") = vbNo Then
        Exit Sub
      End If
      
      '搜尋 Printer, 設定 Printer
      For Each Prn In Printers
         If Prn.DeviceName = cboPrinter.Text Then
            Set Printer = Prn
            Exit For
         End If
      Next
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      ' 設定欄位規格
      BuildField
      ' 產生報表
      '92.3.5 modify by sonia
      'GenerateData
      If GenerateData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
   'Add by Amy 2022/10/05
   Else
        MsgBox "無資料要改號!"
   End If
      
   Unload Me
End Sub

' 列印分隔線
Public Sub PrintSplitLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "-"
   Next nCount
End Sub

' 列印分隔線
Public Sub PrintTerminateLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "="
   Next nCount
End Sub

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   Dim nLeft As Integer
   
   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 1
         m_TopMargin = 2
         'm_ReportWidth = 130
         m_ReportWidth = 92
         'm_ReportDataRows = 30
         m_ReportDataRows = 50
         nFieldWidth = 7
   End Select
   
   nLeft = m_LeftMargin
   'For nIndex = 0 To 7
   For nIndex = 0 To 5
      m_Field(nIndex).Left = nLeft
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Name = "原編號"
            m_Field(nIndex).Width = 10
         Case 1:
            m_Field(nIndex).Name = "新編號"
            m_Field(nIndex).Width = 10
         Case 2:
            m_Field(nIndex).Name = "客戶/代理人名稱"
            m_Field(nIndex).Width = 24
         Case 3:
            m_Field(nIndex).Name = Empty
            m_Field(nIndex).Width = 17   '2008/3/26 MODIFY BY SONIA 原16改為17因TF銷案有17碼
         Case 4:
            m_Field(nIndex).Name = "案件記錄"
            m_Field(nIndex).Width = 17   '2008/3/26 MODIFY BY SONIA 原16改為17因TF銷案有17碼
         Case 5:
            m_Field(nIndex).Name = Empty
            m_Field(nIndex).Width = 17   '2008/3/26 MODIFY BY SONIA 原16改為17因TF銷案有17碼
         'Case 6:
         '   m_Field(nIndex).Name = Empty
         '   m_Field(nIndex).Width = 16
         'Case 7:
         '   m_Field(nIndex).Name = Empty
         '   m_Field(nIndex).Width = 16
      End Select
      nLeft = nLeft + m_Field(nIndex).Width
   Next nIndex
End Sub

' 列印表頭
Private Sub PrintPageHeader(ByVal nPage As Integer)
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim nY As Long
   Dim nCenter As Long
   Dim strTemp As String

   ' 表頭
   nRow = 1
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   nX = m_LeftMargin + m_ReportWidth / 2 - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "客戶/代理人改號記錄"
   
   Printer.Font.Underline = False
   ' 下二列
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
      
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(Date, "EE/MM/DD")
   ' 下一列
   nRow = nRow + 1
   '2008/3/26 ADD BY SONIA
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "PS : 案號後有●表示已銷卷"
   '2008/3/26 END
' 頁次
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   
   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
   
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow
   
   nRow = nRow + 1
   'For nIndex = 0 To 7
   For nIndex = 0 To 5
      nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strTemp
   Next nIndex
     
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow
   
   m_HeaderHeight = nRow
End Sub

' 變更專利基本檔的申請人或代理人編號
'Modified by Lydia 2024/12/12 + ByVal strModMerge As String
Private Sub ModifyPatent(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   'Modify By Sindy 2022/10/24
   Dim strOldCU01 As String
   Dim strOldCU02 As String
   Dim strNewCU01 As String
   Dim strNewCU02 As String
   
   If Len(strOldNum) > 8 Then
      strOldCU01 = Mid(strOldNum, 1, 8)
      strOldCU02 = Mid(strOldNum, 9, 1)
   Else
      strOldCU01 = strOldNum & String(8 - Len(strOldNum), "0")
      strOldCU02 = "0"
   End If
   
   If Len(strNewNum) > 8 Then
      strNewCU01 = Mid(strNewNum, 1, 8)
      strNewCU02 = Mid(strNewNum, 9, 1)
   Else
      strNewCU01 = strNewNum & String(8 - Len(strNewNum), "0")
      strNewCU02 = "0"
   End If
   '2022/10/24 END
   
   ShowStatus "更新專利基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
      ' 變更專利檔的客戶編號
      strSql = "SELECT PA01,PA02,PA03,PA04,PA26,PA27,PA28,PA29,PA30 FROM PATENT " & _
               "WHERE PA26 = '" & strOldNum & "' OR " & _
                     "PA27 = '" & strOldNum & "' OR " & _
                     "PA28 = '" & strOldNum & "' OR " & _
                     "PA29 = '" & strOldNum & "' OR " & _
                     "PA30 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            '2009/3/31 modify by sonia EPC子案不列印
            'InsertItem strOldNum, strNewNum, rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04")
            If rsTmp.Fields("PA04") = "00" Then
               InsertItem strOldNum, strNewNum, rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04")
            End If
            '2009/3/31 end
            '2008/3/26 end
            'GoTo NEXTRECORD1
            ' 更新申請人成新值
            If IsNull(rsTmp.Fields("PA26")) = False Then
               If rsTmp.Fields("PA26") = strOldNum Then
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                     strSql = "UPDATE PATENT SET PA26 = '" & strNewNum & "' " & _
                              "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                                    "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                                    "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                                    "PA04 = '" & rsTmp.Fields("PA04") & "' "
                     cnnConnection.Execute strSql
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                  '2011/10/14 ADD BY SONIA
                  
                  '2011/10/14 END
               End If
            End If
            If IsNull(rsTmp.Fields("PA27")) = False Then
               If rsTmp.Fields("PA27") = strOldNum Then
                  strSql = "UPDATE PATENT SET PA27 = '" & strNewNum & "' " & _
                           "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                                 "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                                 "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                                 "PA04 = '" & rsTmp.Fields("PA04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("PA28")) = False Then
               If rsTmp.Fields("PA28") = strOldNum Then
                  strSql = "UPDATE PATENT SET PA28 = '" & strNewNum & "' " & _
                           "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                                 "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                                 "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                                 "PA04 = '" & rsTmp.Fields("PA04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("PA29")) = False Then
               If rsTmp.Fields("PA29") = strOldNum Then
                  strSql = "UPDATE PATENT SET PA29 = '" & strNewNum & "' " & _
                           "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                                 "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                                 "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                                 "PA04 = '" & rsTmp.Fields("PA04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("PA30")) = False Then
               If rsTmp.Fields("PA30") = strOldNum Then
                  strSql = "UPDATE PATENT SET PA30 = '" & strNewNum & "' " & _
                           "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                                 "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                                 "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                                 "PA04 = '" & rsTmp.Fields("PA04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
NEXTRECORD1:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   Else
      ' 變更專利檔的代理人編號
      strSql = "SELECT PA01,PA02,PA03,PA04,PA75 FROM PATENT " & _
               "WHERE PA75 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            '2009/3/31 modify by sonia EPC子案不列印
            'InsertItem strOldNum, strNewNum, rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04")
            If rsTmp.Fields("PA04") = "00" Then
               InsertItem strOldNum, strNewNum, rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04")
            End If
            '2009/3/31 end
            '2008/3/26 end
            'GoTo NEXTRECORD2
            ' 更新代理人成新值
            strSql = "UPDATE PATENT SET PA75 = '" & strNewNum & "' " & _
                     "WHERE PA01 = '" & rsTmp.Fields("PA01") & "' AND " & _
                           "PA02 = '" & rsTmp.Fields("PA02") & "' AND " & _
                           "PA03 = '" & rsTmp.Fields("PA03") & "' AND " & _
                           "PA04 = '" & rsTmp.Fields("PA04") & "' "
            cnnConnection.Execute strSql
NEXTRECORD2:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '93.2.19 ADD BY SONIA
   strSql = "UPDATE PATENT SET PA76 = '" & strNewNum & "' " & _
                  "WHERE PA76 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA86 = '" & strNewNum & "' " & _
                  "WHERE PA86 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA88 = '" & strNewNum & "' " & _
                  "WHERE PA88 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA101 = '" & strNewNum & "' " & _
                  "WHERE PA101 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA105 = '" & strNewNum & "' " & _
                  "WHERE PA105 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA133 = '" & strNewNum & "' " & _
                  "WHERE PA133 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PATENT SET PA134 = '" & strNewNum & "' " & _
                  "WHERE PA134 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   
   'Added by Morgan 2025/9/3 副本收件人
   strSql = "UPDATE PATENT SET PA168 = '" & strNewNum & "' " & _
                  "WHERE PA168 = '" & strOldNum & "' "
   cnnConnection.Execute strSql, intI
   'end 2025/9/3
   
'   'Add By Sindy 2022/10/24
'   '接洽單申請人
'   strSql = "UPDATE ConsultRecApp SET CRA05 = '" & strNewCU01 & "', " & _
'                                     "CRA06 = '" & strNewCU02 & "' " & _
'            "WHERE CRA05 = '" & strOldCU01 & "' AND " & _
'                  "CRA06 = '" & strOldCU02 & "' "
'   Pub_SeekTbLog strSql
'   cnnConnection.Execute strSql
'   '接洽單主檔-代理人
'   strSql = "UPDATE ConsultRecordList SET CRL60 = '" & strNewCU01 & "', " & _
'                                         "CRL61 = '" & strNewCU02 & "' " & _
'            "WHERE CRL60 = '" & strOldCU01 & "' AND " & _
'                  "CRL61 = '" & strOldCU02 & "' "
'   Pub_SeekTbLog strSql
'   cnnConnection.Execute strSql
'   '2022/10/24 END
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更商標基本檔的申請人或代理人編號
'Modified by Lydia 2024/12/12 + ByVal strModMerge As String
Private Sub ModifyTradeMark(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ShowStatus "更新商標基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
      ' 變更商標基本檔的客戶編號
      'edit by nickc 2007/01/12
      'strSQL = "SELECT TM01,TM02,TM03,TM04,TM23 FROM TRADEMARK " & _
               "WHERE TM23 = '" & strOldNum & "' "
      strSql = "SELECT TM01,TM02,TM03,TM04,TM23,TM78,TM79,TM80,TM81 FROM TRADEMARK " & _
               "WHERE TM23 = '" & strOldNum & "' or " & _
                     "TM78 = '" & strOldNum & "' or " & _
                     "TM79 = '" & strOldNum & "' or " & _
                     "TM80 = '" & strOldNum & "' or " & _
                     "TM81 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If rsTmp.Fields("TM01") = "TF" Then
               strTemp = rsTmp.Fields("TM01") & "-" & Mid(rsTmp.Fields("TM02"), 1, 5) & "-" & Mid(rsTmp.Fields("TM02"), 6, 1) & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
            Else
               strTemp = rsTmp.Fields("TM01") & "-" & rsTmp.Fields("TM02") & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
            End If
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            '2009/3/31 modify by sonia TF子案不列印
            'InsertItem strOldNum, strNewNum, rsTmp.Fields("TM01"), rsTmp.Fields("TM02"), rsTmp.Fields("TM03"), rsTmp.Fields("TM04")
            If rsTmp.Fields("TM04") = "00" Then
               InsertItem strOldNum, strNewNum, rsTmp.Fields("TM01"), rsTmp.Fields("TM02"), rsTmp.Fields("TM03"), rsTmp.Fields("TM04")
            End If
            '2009/3/31 end
            '2008/3/26 end
            'GoTo NEXTRECORD1
            ' 更新申請人成新值
            'add by nickc 2007/01/12 多個時，必須相同才可以變更
            If IsNull(rsTmp.Fields("TM23")) = False Then
               If rsTmp.Fields("TM23") = strOldNum Then
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                     strSql = "UPDATE TRADEMARK SET TM23 = '" & strNewNum & "' " & _
                              "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                                    "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                                    "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                                    "TM04 = '" & rsTmp.Fields("TM04") & "' "
                     cnnConnection.Execute strSql
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                'add by nickc 2007/01/12 搭配上面 if
               End If
            End If
            'add by nickc 2007/01/12 加申請人
            If IsNull(rsTmp.Fields("TM78")) = False Then
               If rsTmp.Fields("TM78") = strOldNum Then
                    strSql = "UPDATE TRADEMARK SET TM78 = '" & strNewNum & "' " & _
                             "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                                   "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                                   "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                                   "TM04 = '" & rsTmp.Fields("TM04") & "' "
                    cnnConnection.Execute strSql
                End If
            End If
            If IsNull(rsTmp.Fields("TM79")) = False Then
               If rsTmp.Fields("TM79") = strOldNum Then
                    strSql = "UPDATE TRADEMARK SET TM79 = '" & strNewNum & "' " & _
                             "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                                   "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                                   "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                                   "TM04 = '" & rsTmp.Fields("TM04") & "' "
                    cnnConnection.Execute strSql
                End If
            End If
            If IsNull(rsTmp.Fields("TM80")) = False Then
               If rsTmp.Fields("TM80") = strOldNum Then
                    strSql = "UPDATE TRADEMARK SET TM80 = '" & strNewNum & "' " & _
                             "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                                   "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                                   "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                                   "TM04 = '" & rsTmp.Fields("TM04") & "' "
                    cnnConnection.Execute strSql
                End If
            End If
            If IsNull(rsTmp.Fields("TM81")) = False Then
               If rsTmp.Fields("TM81") = strOldNum Then
                    strSql = "UPDATE TRADEMARK SET TM81 = '" & strNewNum & "' " & _
                             "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                                   "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                                   "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                                   "TM04 = '" & rsTmp.Fields("TM04") & "' "
                    cnnConnection.Execute strSql
                End If
            End If
NEXTRECORD1:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   Else
      ' 變更商標基本檔的代理人編號
      strSql = "SELECT TM01,TM02,TM03,TM04,TM44 FROM TRADEMARK " & _
               "WHERE TM44 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If rsTmp.Fields("TM01") = "TF" Then
               strTemp = rsTmp.Fields("TM01") & "-" & Mid(rsTmp.Fields("TM02"), 1, 5) & "-" & Mid(rsTmp.Fields("TM02"), 6, 1) & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
            Else
               strTemp = rsTmp.Fields("TM01") & "-" & rsTmp.Fields("TM02") & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
            End If
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            '2009/3/31 modify by sonia TF子案不列印
            'InsertItem strOldNum, strNewNum, rsTmp.Fields("TM01"), rsTmp.Fields("TM02"), rsTmp.Fields("TM03"), rsTmp.Fields("TM04")
            If rsTmp.Fields("TM04") = "00" Then
               InsertItem strOldNum, strNewNum, rsTmp.Fields("TM01"), rsTmp.Fields("TM02"), rsTmp.Fields("TM03"), rsTmp.Fields("TM04")
            End If
            '2009/3/31 end
            '2008/3/26 end
            'GoTo NEXTRECORD2
            ' 更新代理人成新值
            strSql = "UPDATE TRADEMARK SET TM44 = '" & strNewNum & "' " & _
                     "WHERE TM01 = '" & rsTmp.Fields("TM01") & "' AND " & _
                           "TM02 = '" & rsTmp.Fields("TM02") & "' AND " & _
                           "TM03 = '" & rsTmp.Fields("TM03") & "' AND " & _
                           "TM04 = '" & rsTmp.Fields("TM04") & "' "
            cnnConnection.Execute strSql
NEXTRECORD2:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '93.2.19 ADD BY SONIA
   strSql = "UPDATE TRADEMARK SET TM33 = '" & strNewNum & "' " & _
                  "WHERE TM33 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE TRADEMARK SET TM66 = '" & strNewNum & "' " & _
                  "WHERE TM66 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE TRADEMARK SET TM54 = '" & strNewNum & "' " & _
                  "WHERE TM54 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE TRADEMARK SET TM56 = '" & strNewNum & "' " & _
                  "WHERE TM56 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE TRADEMARK SET TM69 = '" & strNewNum & "' " & _
                  "WHERE TM69 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE TRADEMARK SET TM70 = '" & strNewNum & "' " & _
                  "WHERE TM70 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   
   'Added by Morgan 2025/9/3 副本收件人
   strSql = "UPDATE TRADEMARK SET TM132 = '" & strNewNum & "' " & _
                  "WHERE TM132 = '" & strOldNum & "' "
   cnnConnection.Execute strSql, intI
   'end 2025/9/3
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更法務基本檔的申請人或代理人編號
'Modified by Lydia 2024/12/12 + ByVal strModMerge As String
Private Sub ModifyLawCase(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ShowStatus "更新法務基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
      ' 變更法務基本檔的客戶編號
      'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46
'      strSql = "SELECT LC01,LC02,LC03,LC04,LC11 FROM LAWCASE " & _
'               "WHERE LC11 = '" & strOldNum & "' "
      strSql = "SELECT LC01,LC02,LC03,LC04,LC11,LC43,LC44,LC45,LC46 FROM LAWCASE " & _
               "WHERE LC11 = '" & strOldNum & "' OR " & _
                     "LC43 = '" & strOldNum & "' OR " & _
                     "LC44 = '" & strOldNum & "' OR " & _
                     "LC45 = '" & strOldNum & "' OR " & _
                     "LC46 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("LC01") & "-" & rsTmp.Fields("LC02") & "-" & rsTmp.Fields("LC03") & "-" & rsTmp.Fields("LC04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            InsertItem strOldNum, strNewNum, rsTmp.Fields("LC01"), rsTmp.Fields("LC02"), rsTmp.Fields("LC03"), rsTmp.Fields("LC04")
            '2008/3/26 end
            'GoTo NEXTRECORD1
            ' 更新申請人成新值
            If IsNull(rsTmp.Fields("LC11")) = False Then
               If rsTmp.Fields("LC11") = strOldNum Then
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                     strSql = "UPDATE LAWCASE SET LC11 = '" & strNewNum & "' " & _
                              "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                                    "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                                    "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                                    "LC04 = '" & rsTmp.Fields("LC04") & "' "
                     cnnConnection.Execute strSql
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
               End If
            End If
            'Add By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46
            If IsNull(rsTmp.Fields("LC43")) = False Then
               If rsTmp.Fields("LC43") = strOldNum Then
                  strSql = "UPDATE LAWCASE SET LC43 = '" & strNewNum & "' " & _
                           "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                                 "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                                 "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                                 "LC04 = '" & rsTmp.Fields("LC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("LC44")) = False Then
               If rsTmp.Fields("LC44") = strOldNum Then
                  strSql = "UPDATE LAWCASE SET LC44 = '" & strNewNum & "' " & _
                           "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                                 "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                                 "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                                 "LC04 = '" & rsTmp.Fields("LC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("LC45")) = False Then
               If rsTmp.Fields("LC45") = strOldNum Then
                  strSql = "UPDATE LAWCASE SET LC45 = '" & strNewNum & "' " & _
                           "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                                 "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                                 "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                                 "LC04 = '" & rsTmp.Fields("LC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("LC46")) = False Then
               If rsTmp.Fields("LC46") = strOldNum Then
                  strSql = "UPDATE LAWCASE SET LC46 = '" & strNewNum & "' " & _
                           "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                                 "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                                 "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                                 "LC04 = '" & rsTmp.Fields("LC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            '2011/2/24 End
NEXTRECORD1:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   Else
      ' 變更法務基本檔的代理人編號
      strSql = "SELECT LC01,LC02,LC03,LC04,LC22 FROM LAWCASE " & _
               "WHERE LC22 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("LC01") & "-" & rsTmp.Fields("LC02") & "-" & rsTmp.Fields("LC03") & "-" & rsTmp.Fields("LC04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            InsertItem strOldNum, strNewNum, rsTmp.Fields("LC01"), rsTmp.Fields("LC02"), rsTmp.Fields("LC03"), rsTmp.Fields("LC04")
            '2008/3/26 end
            'GoTo NEXTRECORD2
            ' 更新代理人成新值
            strSql = "UPDATE LAWCASE SET LC22 = '" & strNewNum & "' " & _
                     "WHERE LC01 = '" & rsTmp.Fields("LC01") & "' AND " & _
                           "LC02 = '" & rsTmp.Fields("LC02") & "' AND " & _
                           "LC03 = '" & rsTmp.Fields("LC03") & "' AND " & _
                           "LC04 = '" & rsTmp.Fields("LC04") & "' "
            cnnConnection.Execute strSql
NEXTRECORD2:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '93.2.19 ADD BY SONIA
   strSql = "UPDATE LAWCASE SET LC12 = '" & strNewNum & "' " & _
                  "WHERE LC12 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE LAWCASE SET LC26 = '" & strNewNum & "' " & _
                  "WHERE LC26 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE LAWCASE SET LC35 = '" & strNewNum & "' " & _
                  "WHERE LC35 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub
'93.2.14 add by sonia
' 變更國外帳款未收未付歷史資料檔的申請人或代理人編號
Private Sub ModifyNeverClose(ByVal strOldNum As String, ByVal strNewNum As String)
   ' 修改案件進度檔 2005/4/13 ADD BY SONIA
   ' 當修改的是代理人編號時
   If Mid(strOldNum, 1, 1) = "Y" Then
      ModifyCaseProgress strOldNum, strNewNum
   End If
   '2005/4/13 END
   ' 更新國外請款資料(主檔)
   ShowStatus "變更國外請款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC1K0 SET A1K03 = '" & strNewNum & "' " & _
                  "WHERE A1K03 = '" & strOldNum & "' AND A1K29 IS NULL AND (A1K30=0 or a1k30 is null)"
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC1K0 SET A1K27 = '" & strNewNum & "' " & _
                  "WHERE A1K27 = '" & strOldNum & "' AND A1K29 IS NULL AND (A1K30=0 or a1k30 is null)"
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC1K0 SET A1K28 = '" & strNewNum & "' " & _
                  "WHERE A1K28 = '" & strOldNum & "' AND A1K29 IS NULL AND (A1K30=0 or a1k30 is null)"
   cnnConnection.Execute strSql
   
   ' 更新國外帳單資料(主檔)
   ShowStatus "變更國外帳單資料(主檔)中, 請稍候 . . ."
   strSql = "UPDATE ACC150 SET A1503 = '" & strNewNum & "' " & _
                  "WHERE A1503 = '" & strOldNum & "' AND (A1520 IS NULL OR A1520=0)"
   cnnConnection.Execute strSql
   
End Sub
' 變更歷史資料檔的申請人或代理人編號
Private Sub ModifyHistoryData(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strDelOldNum As String)
'Add By Sindy 2014/11/6
Dim stSQL As String, intQ As Integer
Dim rsQuery As ADODB.Recordset
Dim strOldIN01 As String, strOldIN02 As String
Dim strOldIN03 As String, strOldIN04 As String
Dim strOldIN05 As String, strOldIN06 As String
Dim strUpdID As String
Dim strMaxIN02 As String
'2014/11/6 END
     
   ' 修改案件進度檔
   ModifyCaseProgress strOldNum, strNewNum
   
   ' 更新資料刪除記錄檔
   ShowStatus "變更資料刪除記錄檔中, 請稍候 . . ."
   strSql = "UPDATE DATADELETERECORD SET DD06 = '" & strNewNum & "' " & _
                  "WHERE DD06 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE DATADELETERECORD SET DD12= '" & strNewNum & "' " & _
                  "WHERE DD12 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE DATADELETERECORD SET DD13 = '" & strNewNum & "' " & _
                  "WHERE DD13 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 ADD BY SONIA
   ' 更新客戶發明人檔
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      ShowStatus "變更客戶發明人檔中, 請稍候 . . ."
      'edit by nick 2004/11/04
      'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
      stSQL = "select * from Inventor where IN01='" & Left(strOldNum, 8) & "'"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         rsQuery.MoveFirst
         Do While Not rsQuery.EOF
            strOldIN01 = "" & rsQuery.Fields("IN01")
            strOldIN02 = "" & rsQuery.Fields("IN02")
            strOldIN03 = "" & rsQuery.Fields("IN03")
            strOldIN04 = "" & rsQuery.Fields("IN04")
            strOldIN05 = "" & rsQuery.Fields("IN05")
            strOldIN06 = "" & rsQuery.Fields("IN06")
            strUpdID = "" '預設值
            '用ID檢查
            If strUpdID = "" And strOldIN03 <> "" Then
               strExc(0) = "select * from Inventor where IN01='" & Left(strNewNum, 8) & "' and IN03='" & strOldIN03 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strUpdID = RsTemp.Fields("IN01") & RsTemp.Fields("IN02")
               End If
            End If
            '用中文名稱檢查
            If strUpdID = "" And strOldIN04 <> "" Then
               strExc(0) = "select * from Inventor where IN01='" & Left(strNewNum, 8) & "' and IN04='" & strOldIN04 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strUpdID = RsTemp.Fields("IN01") & RsTemp.Fields("IN02")
               End If
            End If
            '用英文名稱檢查
            If strUpdID = "" And strOldIN05 <> "" Then
               strExc(0) = "select * from Inventor where IN01='" & Left(strNewNum, 8) & "' and IN05='" & strOldIN05 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strUpdID = RsTemp.Fields("IN01") & RsTemp.Fields("IN02")
               End If
            End If
            '用日文名稱檢查
            If strUpdID = "" And strOldIN06 <> "" Then
               strExc(0) = "select * from Inventor where IN01='" & Left(strNewNum, 8) & "' and IN06='" & strOldIN06 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strUpdID = RsTemp.Fields("IN01") & RsTemp.Fields("IN02")
               End If
            End If
            '發明人已存在
            If strUpdID <> "" Then
               '更新專利發明人檔
               strSql = "update patentInventor set pi06='" & strUpdID & "'" & _
                        " where substr(pi06,1,8)='" & strOldIN01 & "' and substr(pi06,9,2)='" & strOldIN02 & "'"
               Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
               cnnConnection.Execute strSql, intI
   '            'Add By Sindy 2022/10/24
   '            '接洽單發明人資料
   '            strSql = "update ConsultRecInv set CRi03='" & substr(strUpdID, 1, 8) & "', CRi04='" & substr(strUpdID, 9, 2) & "'" & _
   '                     " where CRi03='" & strOldIN01 & "' and CRi04='" & strOldIN02 & "'"
   '            cnnConnection.Execute strSql, intI
   '            '2022/10/24 END
               '刪除客戶發明人檔
               strSql = "delete from Inventor" & _
                        " where IN01='" & strOldIN01 & "' and IN02='" & strOldIN02 & "'"
               Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
               cnnConnection.Execute strSql, intI
            '不存在
            Else
               '更新客戶發明人檔
               strMaxIN02 = PUB_GetNewIN02(Left(strNewNum, 8)) '*****
               strSql = "UPDATE Inventor SET IN01='" & Left(strNewNum, 8) & "',IN02='" & strMaxIN02 & "'" & _
                        " WHERE IN01='" & strOldIN01 & "' and IN02='" & strOldIN02 & "'"
               Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
               cnnConnection.Execute strSql, intI
               '更新專利發明人檔
               strSql = "update patentInventor set pi06='" & Left(strNewNum, 8) & strMaxIN02 & "'" & _
                        " where substr(pi06,1,8)='" & strOldIN01 & "' and substr(pi06,9,2)='" & strOldIN02 & "'"
               Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
               cnnConnection.Execute strSql, intI
   '            'Add By Sindy 2022/10/24
   '            '接洽單發明人資料
   '            strSql = "update ConsultRecInv set CRi03='" & Left(strNewNum, 8) & "', CRi04='" & Format(intMaxIN02, "00") & "'" & _
   '                     " where CRi03='" & strOldIN01 & "' and CRi04='" & strOldIN02 & "'"
   '            cnnConnection.Execute strSql, intI
   '            '2022/10/24 END
            End If
            rsQuery.MoveNext
         Loop
      End If
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   Set rsQuery = Nothing
   '2014/11/6 END
   'end 2010/5/25
   
   'add by nick 2004/11/04
   'modify by sonia 2018/4/3  非更名前的舊名稱編號才可刪除
   'If strDelOldNum = "Y" Then
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      If strDelOldNum = "Y" And Mid(strOldNum, 9, 1) = "0" Then
         strSql = "delete from Inventor where IN01 = '" & Left(strOldNum, 8) & "' "
         cnnConnection.Execute strSql
      End If
      ' 更新申請人國外ID對照檔
      ShowStatus "變更申請人國外ID對照檔中, 請稍候 . . ."
      'edit by nick 2004/11/04
      'strSQL = "UPDATE ApplicantForeignID SET AFID01 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE AFID01 = '" & Left(strOldNum, 8) & "' "
      strSql = "UPDATE ApplicantForeignID SET AFID01 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE AFID01 = '" & Left(strOldNum, 8) & "' and not exists(select * from ApplicantForeignID where AFID01='" & Left(strNewNum, 8) & "') "
      cnnConnection.Execute strSql
      'add by nick 2004/11/04
      'modify by sonia 2018/4/3  非更名前的舊名稱編號才可刪除
      'If strDelOldNum = "Y" Then
      If strDelOldNum = "Y" And Mid(strOldNum, 9, 1) = "0" Then
         strSql = "delete from ApplicantForeignID where AFID01 = '" & Left(strOldNum, 8) & "' "
         cnnConnection.Execute strSql
      End If
   End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   ' 更新變更事項檔
   ShowStatus "變更變更事項檔中, 請稍候 . . ."
   strSql = "UPDATE ChangeEvent SET CE04 = '" & strNewNum & "' " & _
                  "WHERE CE04 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ChangeEvent SET CE05= '" & strNewNum & "' " & _
                  "WHERE CE05 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ChangeEvent SET CE06 = '" & strNewNum & "' " & _
                  "WHERE CE06 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ChangeEvent SET CE07= '" & strNewNum & "' " & _
                  "WHERE CE07 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ChangeEvent SET CE08= '" & strNewNum & "' " & _
                  "WHERE CE08 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   
   ' 更新轉傳票分錄資料
   ShowStatus "變更轉傳票分錄資料中, 請稍候 . . ."
   strSql = "UPDATE ACC1P0 SET A1P15 = '" & strNewNum & "' " & _
                  "WHERE A1P15 = '" & strOldNum & "' "
   'Modified by Lydia 2017/09/05 避免觸發Trigger
   'cnnConnection.Execute strSql
   cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & "; end;"
   
   ' 更新傳票資料(交易檔-財務)
   ShowStatus "變更傳票資料中, 請稍候 . . ."
   strSql = "UPDATE ACC021 SET AX208 = '" & strNewNum & "' " & _
                  "WHERE AX208 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新傳票資料(交易檔-帳務)
   ShowStatus "變更傳票資料中, 請稍候 . . ."
   strSql = "UPDATE ACC031 SET AX308 = '" & strNewNum & "' " & _
                  "WHERE AX308 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之票據資料。
   ShowStatus "變更票據資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0E0 SET A0E06 = '" & strNewNum & "' " & _
                  "WHERE A0E06 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
            
   ' 更新國內未開收據案件資料(暫存檔)
   ShowStatus "變更國內未開收據案件資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0J0 SET A0J11 = '" & strNewNum & "' " & _
                  "WHERE A0J11 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之國內收據資料。
   ShowStatus "變更國內收據資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0K0 SET A0K03 = '" & strNewNum & "' " & _
                  "WHERE A0K03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之國內暫收款資料（主檔）。
   ShowStatus "變更國內暫收款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0T0 SET A0T06 = '" & strNewNum & "' " & _
                  "WHERE A0T06 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
            
   ' 修改財務之國內應付款資料（主檔）。
   ShowStatus "變更國內應付款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0O0 SET A0O03 = '" & strNewNum & "' " & _
                  "WHERE A0O03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
            
   ' 修改財務之國內付款資料（主檔）。
   ShowStatus "變更國內付款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0Q0 SET A0Q03 = '" & strNewNum & "' " & _
                  "WHERE A0Q03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外請款資料(主檔)
   ShowStatus "變更國外請款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC1K0 SET A1K03 = '" & strNewNum & "' " & _
                  "WHERE A1K03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC1K0 SET A1K27 = '" & strNewNum & "' " & _
                  "WHERE A1K27 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC1K0 SET A1K28 = '" & strNewNum & "' " & _
                  "WHERE A1K28 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之國外收款資料（主檔）。
   ShowStatus "變更國外收款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0Y0 SET A0Y07 = '" & strNewNum & "' " & _
                  "WHERE A0Y07 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC0Y0 SET A0Y08 = '" & strNewNum & "' " & _
                  "WHERE A0Y08 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE ACC0Y0 SET A0Y09 = '" & strNewNum & "' " & _
                  "WHERE A0Y09 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外帳單資料(主檔)
   ShowStatus "變更國外帳單資料(主檔)中, 請稍候 . . ."
   strSql = "UPDATE ACC150 SET A1503 = '" & strNewNum & "' " & _
                  "WHERE A1503 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外帳單資料(交易檔)
   ShowStatus "變更國外帳單資料(交易檔)中, 請稍候 . . ."
   strSql = "UPDATE ACC151 SET AXF05 = '" & strNewNum & "' " & _
                  "WHERE AXF05 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外抵帳單資料(主檔)
   ShowStatus "變更國外抵帳單資料(主檔)中, 請稍候 . . ."
   strSql = "UPDATE ACC160 SET A1603 = '" & strNewNum & "' " & _
                  "WHERE A1603 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外抵帳單資料(交易檔)
   ShowStatus "變更國外抵帳單資料(交易檔)中, 請稍候 . . ."
   strSql = "UPDATE ACC161 SET AXG05 = '" & strNewNum & "' " & _
                  "WHERE AXG05 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之國外付款資料（主檔）。
    ShowStatus "變更國外付款資料（主檔）中, 請稍候 . . ."
   strSql = "UPDATE ACC180 SET A1803 = '" & strNewNum & "' " & _
                  "WHERE A1803 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
  
   ' 更新國外結匯資料
   ShowStatus "變更國外結匯資料中, 請稍候 . . ."
   strSql = "UPDATE ACC170 SET A1705 = '" & strNewNum & "' " & _
                  "WHERE A1705 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
            
   ' 修改財務之國外匯票資料（主檔）。
   ShowStatus "變更國外匯票資料（主檔）中, 請稍候 . . ."
   strSql = "UPDATE ACC1B0 SET A1B02 = '" & strNewNum & "' " & _
                  "WHERE A1B02 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
            
   ' 修改財務之國外匯票資料（交易檔）。
   ShowStatus "變更國外匯票資料（交易檔）中, 請稍候 . . ."
   strSql = "UPDATE ACC1C0 SET A1C02 = '" & strNewNum & "' " & _
                  "WHERE A1C02 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外暫收款資料
   ShowStatus "變更國外暫收款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC120 SET A1203 = '" & strNewNum & "' " & _
                  "WHERE A1203 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ' 修改財務之國外暫收款退費資料（主檔）
   ShowStatus "變更國外暫收款退費資料（主檔）中, 請稍候 . . ."
   strSql = "UPDATE ACC130 SET A1304 = '" & strNewNum & "' " & _
                  "WHERE A1304 = '" & strOldNum & "' "
   cnnConnection.Execute strSql

   'add by sonia 2015/3/24 修改財務之客戶/代理人匯款銀行資料維護
   ShowStatus "變更客戶/代理人匯款銀行資料維護中, 請稍候 . . ."
   strSql = "UPDATE ACC220 SET A2201 = '" & strNewNum & "' " & _
                  "WHERE A2201 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   'Add by Morgan 2016/5/30
   ShowStatus "變更繳款簽收資料維護中, 請稍候 . . ."
   strSql = "UPDATE ACC230 SET A2304 = '" & strNewNum & "' WHERE A2304 = '" & strOldNum & "'"
   cnnConnection.Execute strSql
   'end 2016/5/30

   '2009/3/30 ADD BY SONIA
   ShowStatus "變更往來記錄資料中, 請稍候 . . ."
   'Modify By Sindy 2019/7/4 strNewNum => Left(strNewNum, 8) & "0"
   strSql = "UPDATE ContactRecord SET CR03 = '" & Left(strNewNum, 8) & "0" & "' " & _
                  "WHERE CR03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ShowStatus "變更國內往來記錄資料中, 請稍候 . . ."
   strSql = "UPDATE ContactRecord1 SET COR03 = '" & strNewNum & "' " & _
                  "WHERE COR03 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      ShowStatus "變更客戶減免身分資料中, 請稍候 . . ."
      strSql = "UPDATE ApplicantDiscount SET AD01 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE AD01 = '" & Left(strOldNum, 8) & "' and not exists(select * from ApplicantDiscount where AD01='" & Left(strNewNum, 8) & "') "
      Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
      cnnConnection.Execute strSql
      'modify by sonia 2018/4/3  非更名前的舊名稱編號才可刪除
      'If strDelOldNum = "Y" Then
      If strDelOldNum = "Y" And Mid(strOldNum, 9, 1) = "0" Then
         strSql = "delete from ApplicantDiscount where AD01 = '" & Left(strOldNum, 8) & "' "
         Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
         cnnConnection.Execute strSql
      End If
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   ShowStatus "變更特殊客戶記錄異動資料中, 請稍候 . . ."
   strSql = "UPDATE CustSpecialLog SET CL01 = '" & strNewNum & "' " & _
                  "WHERE CL01 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      ShowStatus "變更代理人互惠設定資料中, 請稍候 . . ."
      strSql = "UPDATE FagentConfig SET FC01 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE FC01 = '" & Left(strOldNum, 8) & "' "
      Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
      cnnConnection.Execute strSql
      
      ShowStatus "變更代理人目標給案量資料中, 請稍候 . . ."
      strSql = "UPDATE FagentTarget SET FT01 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE FT01 = '" & Left(strOldNum, 8) & "' "
      Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
      cnnConnection.Execute strSql
      '2009/3/30 END
   End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   '2010/8/6 add by sonia
   ShowStatus "變更重新委任客戶資料中, 請稍候 . . ."
   strSql = "UPDATE LINREASIGNREC SET LR01 = '" & strNewNum & "' " & _
                  "WHERE LR01 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2010/8/6 end
   
   '2010/8/9 ADD BY SONIA 案件進度檔的原收文FC代理人
   ' 案件進度檔的代理人編號
   '2010/10/13 MODIFY BY SONIA 改號CP139也要更新
   'If strDelOldNum = "Y" Then
   If strDelOldNum = "Y" Or strNewFagent = True Then
      strSql = "UPDATE CASEPROGRESS SET CP139 = '" & strNewNum & "' " & _
               "WHERE CP139 = '" & strOldNum & "' "
      cnnConnection.Execute strSql
   End If
   '2010/8/9 END
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      '2011/9/19 ADD BY SONIA ACC260客製化請款項目資料
      ShowStatus "變更客製化請款項目資料中, 請稍候 . . ."
      strSql = "UPDATE ACC260 SET A2601 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE A2601 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
   '2011/9/19 END
   End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   '2015/7/29 ADD BY SONIA 彼所案號異動紀錄 FCCaseNoLog
   ShowStatus "變更彼所案號異動紀錄資料中, 請稍候 . . ."
   strSql = "UPDATE FCCaseNoLog SET FL05 = '" & strNewNum & "' " & _
                  "WHERE FL05 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2015/7/29 END
End Sub
' 變更顧問基本檔的申請人或代理人編號
'Modified by Lydia 2024/12/12 + ByVal strModMerge As String
Private Sub ModifyHireCase(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ShowStatus "更新顧問基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
      ' 變更顧問基本檔的客戶編號
      'Modify By Sindy 2011/2/24 增加HC24,HC25,HC26,HC27
'      strSql = "SELECT HC01,HC02,HC03,HC04,HC05 FROM HIRECASE " & _
'               "WHERE HC05 = '" & strOldNum & "' "
      strSql = "SELECT HC01,HC02,HC03,HC04,HC05,HC24,HC25,HC26,HC27 FROM HIRECASE " & _
               "WHERE HC05 = '" & strOldNum & "' OR " & _
                     "HC24 = '" & strOldNum & "' OR " & _
                     "HC25 = '" & strOldNum & "' OR " & _
                     "HC26 = '" & strOldNum & "' OR " & _
                     "HC27 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("HC01") & "-" & rsTmp.Fields("HC02") & "-" & rsTmp.Fields("HC03") & "-" & rsTmp.Fields("HC04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            InsertItem strOldNum, strNewNum, rsTmp.Fields("HC01"), rsTmp.Fields("HC02"), rsTmp.Fields("HC03"), rsTmp.Fields("HC04")
            '2008/3/26 end
            'GoTo NEXTRECORD1
            ' 更新申請人成新值
            If IsNull(rsTmp.Fields("HC05")) = False Then
               If rsTmp.Fields("HC05") = strOldNum Then
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                     strSql = "UPDATE HIRECASE SET HC05 = '" & strNewNum & "' " & _
                              "WHERE HC01 = '" & rsTmp.Fields("HC01") & "' AND " & _
                                    "HC02 = '" & rsTmp.Fields("HC02") & "' AND " & _
                                    "HC03 = '" & rsTmp.Fields("HC03") & "' AND " & _
                                    "HC04 = '" & rsTmp.Fields("HC04") & "' "
                     cnnConnection.Execute strSql
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                  
               End If
            End If
            'Add By Sindy 2011/2/24 增加HC24,HC25,HC26,HC27
            If IsNull(rsTmp.Fields("HC24")) = False Then
               If rsTmp.Fields("HC24") = strOldNum Then
                  strSql = "UPDATE HIRECASE SET HC24 = '" & strNewNum & "' " & _
                           "WHERE HC01 = '" & rsTmp.Fields("HC01") & "' AND " & _
                                 "HC02 = '" & rsTmp.Fields("HC02") & "' AND " & _
                                 "HC03 = '" & rsTmp.Fields("HC03") & "' AND " & _
                                 "HC04 = '" & rsTmp.Fields("HC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("HC25")) = False Then
               If rsTmp.Fields("HC25") = strOldNum Then
                  strSql = "UPDATE HIRECASE SET HC25 = '" & strNewNum & "' " & _
                           "WHERE HC01 = '" & rsTmp.Fields("HC01") & "' AND " & _
                                 "HC02 = '" & rsTmp.Fields("HC02") & "' AND " & _
                                 "HC03 = '" & rsTmp.Fields("HC03") & "' AND " & _
                                 "HC04 = '" & rsTmp.Fields("HC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("HC26")) = False Then
               If rsTmp.Fields("HC26") = strOldNum Then
                  strSql = "UPDATE HIRECASE SET HC26 = '" & strNewNum & "' " & _
                           "WHERE HC01 = '" & rsTmp.Fields("HC01") & "' AND " & _
                                 "HC02 = '" & rsTmp.Fields("HC02") & "' AND " & _
                                 "HC03 = '" & rsTmp.Fields("HC03") & "' AND " & _
                                 "HC04 = '" & rsTmp.Fields("HC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("HC27")) = False Then
               If rsTmp.Fields("HC27") = strOldNum Then
                  strSql = "UPDATE HIRECASE SET HC27 = '" & strNewNum & "' " & _
                           "WHERE HC01 = '" & rsTmp.Fields("HC01") & "' AND " & _
                                 "HC02 = '" & rsTmp.Fields("HC02") & "' AND " & _
                                 "HC03 = '" & rsTmp.Fields("HC03") & "' AND " & _
                                 "HC04 = '" & rsTmp.Fields("HC04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            '2011/2/24 End
NEXTRECORD1:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更服務業務基本檔的申請人或代理人編號
'Modified by Lydia 2024/12/12 + ByVal strModMerge As String
Private Sub ModifyServicePractice(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ShowStatus "更新服務業務基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
      ' 變更服務業務基本檔的客戶編號
'edit by nickc 2007/01/12 加秀申請人 4 & 5
'      strSQL = "SELECT SP01,SP02,SP03,SP04,SP08,SP58,SP59 FROM SERVICEPRACTICE " &
      strSql = "SELECT SP01,SP02,SP03,SP04,SP08,SP58,SP59,SP65,SP66 FROM SERVICEPRACTICE " & _
               "WHERE SP08 = '" & strOldNum & "' OR " & _
                     "SP58 = '" & strOldNum & "' OR " & _
                     "SP59 = '" & strOldNum & "' OR " & _
                     "SP65 = '" & strOldNum & "' OR " & _
                     "SP66 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("SP01") & "-" & rsTmp.Fields("SP02") & "-" & rsTmp.Fields("SP03") & "-" & rsTmp.Fields("SP04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            InsertItem strOldNum, strNewNum, rsTmp.Fields("SP01"), rsTmp.Fields("SP02"), rsTmp.Fields("SP03"), rsTmp.Fields("SP04")
            '2008/3/26 end
            'GoTo NEXTRECORD1
            ' 更新申請人成新值
            If IsNull(rsTmp.Fields("SP08")) = False Then
               If rsTmp.Fields("SP08") = strOldNum Then
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
                     strSql = "UPDATE SERVICEPRACTICE SET SP08 = '" & strNewNum & "' " & _
                              "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                                    "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                                    "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                                    "SP04 = '" & rsTmp.Fields("SP04") & "' "
                     cnnConnection.Execute strSql
                  If strModMerge = "Y" Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/12/12 +控制Trigger 不被觸發
               End If
            End If
            If IsNull(rsTmp.Fields("SP58")) = False Then
               If rsTmp.Fields("SP58") = strOldNum Then
                  strSql = "UPDATE SERVICEPRACTICE SET SP58 = '" & strNewNum & "' " & _
                           "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                                 "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                                 "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                                 "SP04 = '" & rsTmp.Fields("SP04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("SP59")) = False Then
               If rsTmp.Fields("SP59") = strOldNum Then
                  strSql = "UPDATE SERVICEPRACTICE SET SP59 = '" & strNewNum & "' " & _
                           "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                                 "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                                 "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                                 "SP04 = '" & rsTmp.Fields("SP04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            'add by nickc 2007/01/12 加申請人
            If IsNull(rsTmp.Fields("SP65")) = False Then
               If rsTmp.Fields("SP65") = strOldNum Then
                  strSql = "UPDATE SERVICEPRACTICE SET SP65 = '" & strNewNum & "' " & _
                           "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                                 "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                                 "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                                 "SP04 = '" & rsTmp.Fields("SP04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
            If IsNull(rsTmp.Fields("SP66")) = False Then
               If rsTmp.Fields("SP66") = strOldNum Then
                  strSql = "UPDATE SERVICEPRACTICE SET SP66 = '" & strNewNum & "' " & _
                           "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                                 "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                                 "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                                 "SP04 = '" & rsTmp.Fields("SP04") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
NEXTRECORD1:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   Else
      ' 變更服務業務基本檔的代理人編號
      strSql = "SELECT SP01,SP02,SP03,SP04,SP26 FROM SERVICEPRACTICE " & _
               "WHERE SP26 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            strTemp = rsTmp.Fields("SP01") & "-" & rsTmp.Fields("SP02") & "-" & rsTmp.Fields("SP03") & "-" & rsTmp.Fields("SP04")
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            InsertItem strOldNum, strNewNum, rsTmp.Fields("SP01"), rsTmp.Fields("SP02"), rsTmp.Fields("SP03"), rsTmp.Fields("SP04")
            '2008/3/26 end
            'GoTo NEXTRECORD2
            ' 更新代理人成新值
            strSql = "UPDATE SERVICEPRACTICE SET SP26 = '" & strNewNum & "' " & _
                     "WHERE SP01 = '" & rsTmp.Fields("SP01") & "' AND " & _
                           "SP02 = '" & rsTmp.Fields("SP02") & "' AND " & _
                           "SP03 = '" & rsTmp.Fields("SP03") & "' AND " & _
                           "SP04 = '" & rsTmp.Fields("SP04") & "' "
            cnnConnection.Execute strSql
NEXTRECORD2:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '93.2.19 ADD BY SONIA
   strSql = "UPDATE SERVICEPRACTICE SET SP35 = '" & strNewNum & "' " & _
                  "WHERE SP35 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE SERVICEPRACTICE SET SP37 = '" & strNewNum & "' " & _
                  "WHERE SP37 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE SERVICEPRACTICE SET SP67 = '" & strNewNum & "' " & _
                  "WHERE SP67 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END

   'Added by Morgan 2025/9/3 副本收件人
   strSql = "UPDATE SERVICEPRACTICE SET SP86 = '" & strNewNum & "' " & _
                  "WHERE SP86 = '" & strOldNum & "' "
   cnnConnection.Execute strSql, intI
   'end 2025/9/3
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更案件進度檔的申請人或代理人編號
Private Sub ModifyCaseProgress(ByVal strOldNum As String, ByVal strNewNum As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ShowStatus "更新案件進度檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   

   ' 當修改的是客戶編號時
   If Mid(strOldNum, 1, 1) = "X" Then
   '911206 nick 邱小姐新增加的規則
   '***** start
        ' 案件進度檔的客戶編號
        'edit by nickc 2007/01/12 加申請人
        'strSQL = "SELECT CP01,CP02,CP03,CP04,cp55,cp56,cp72 FROM CASEPROGRESS " & _
                 "WHERE cp55 = '" & strOldNum & "' or cp56='" & strOldNum & "' OR cp72='" & strOldNum & "' "
        strSql = "SELECT CP01,CP02,CP03,CP04,cp55,cp56,cp72,cp93,cp94,cp95,cp96,cp89,cp90,cp91,cp92 FROM CASEPROGRESS " & _
                 "WHERE cp55 = '" & strOldNum & "' or " & _
                       "cp56 = '" & strOldNum & "' OR " & _
                       "cp72 = '" & strOldNum & "' or " & _
                       "cp93 = '" & strOldNum & "' or " & _
                       "cp94 = '" & strOldNum & "' or " & _
                       "cp95 = '" & strOldNum & "' or " & _
                       "cp96 = '" & strOldNum & "' or " & _
                       "cp89 = '" & strOldNum & "' or " & _
                       "cp90 = '" & strOldNum & "' or " & _
                       "cp91 = '" & strOldNum & "' or " & _
                       "cp92 = '" & strOldNum & "' "
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
           Do While rsTmp.EOF = False
              If rsTmp.Fields("CP01") = "TF" Then
                 strTemp = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
              Else
                 strTemp = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
              End If
              '2008/3/26 modif by sonia
              'InsertItem strOldNum, strNewNum, strTemp
              '2009/3/31 modify by sonia 子案不列印
              'InsertItem strOldNum, strNewNum, rsTmp.Fields("CP01"), rsTmp.Fields("CP02"), rsTmp.Fields("CP03"), rsTmp.Fields("CP04")
              If rsTmp.Fields("CP04") = "00" Then
                 InsertItem strOldNum, strNewNum, rsTmp.Fields("CP01"), rsTmp.Fields("CP02"), rsTmp.Fields("CP03"), rsTmp.Fields("CP04")
              End If
              '2009/3/31 end
              '2008/3/26 end
                If IsNull(rsTmp.Fields("cp55")) = False Then
                   If rsTmp.Fields("cp55") = strOldNum Then
                      '93.8.10 modify by sonia
                      'strSQL = "UPDATE CASEPROGRESS SET cp55 = '" & strNewNum & "' " & _
                      '         "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                      '               "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                      '               "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                      '               "CP04 = '" & rsTmp.Fields("CP04") & "' "
                      strSql = "UPDATE CASEPROGRESS SET cp55 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP55 = '" & strOldNum & "' "
                      '93.8.10 END
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp56")) = False Then
                   If rsTmp.Fields("cp56") = strOldNum Then
                      '93.8.10 modify by sonia
                      'strSQL = "UPDATE CASEPROGRESS SET cp56 = '" & strNewNum & "' " & _
                      '         "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                      '               "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                      '               "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                      '               "CP04 = '" & rsTmp.Fields("CP04") & "' "
                      strSql = "UPDATE CASEPROGRESS SET cp56 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP56 = '" & strOldNum & "' "
                      '93.8.10 END
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp72")) = False Then
                   If rsTmp.Fields("cp72") = strOldNum Then
                      '93.8.10 modify by sonia
                      'strSQL = "UPDATE CASEPROGRESS SET cp72 = '" & strNewNum & "' " & _
                      '         "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                      '               "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                      '               "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                      '               "CP04 = '" & rsTmp.Fields("CP04") & "' "
                      strSql = "UPDATE CASEPROGRESS SET cp72 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP72 = '" & strOldNum & "' "
                      '93.8.10 END
                      cnnConnection.Execute strSql
                   End If
                End If
                'add by nickc 2007/01/12 加申請人
                If IsNull(rsTmp.Fields("cp93")) = False Then
                   If rsTmp.Fields("cp93") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp93 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP93 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp94")) = False Then
                   If rsTmp.Fields("cp94") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp94 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP94 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp95")) = False Then
                   If rsTmp.Fields("cp95") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp95 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP95 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp96")) = False Then
                   If rsTmp.Fields("cp96") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp96 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP96 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp89")) = False Then
                   If rsTmp.Fields("cp89") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp89 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP89 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp90")) = False Then
                   If rsTmp.Fields("cp90") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp90 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP90 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp91")) = False Then
                   If rsTmp.Fields("cp91") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp91 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP91 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
                If IsNull(rsTmp.Fields("cp92")) = False Then
                   If rsTmp.Fields("cp92") = strOldNum Then
                      strSql = "UPDATE CASEPROGRESS SET cp92 = '" & strNewNum & "' " & _
                               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP92 = '" & strOldNum & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
              rsTmp.MoveNext
           Loop
        End If
        rsTmp.Close
    '***** end
   Else
      ' 案件進度檔的代理人編號
      strSql = "SELECT CP01,CP02,CP03,CP04,CP44 FROM CASEPROGRESS " & _
               "WHERE CP44 = '" & strOldNum & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If rsTmp.Fields("CP01") = "TF" Then
               strTemp = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
            Else
               strTemp = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
            End If
            '2008/3/26 modif by sonia
            'InsertItem strOldNum, strNewNum, strTemp
            '2009/3/31 modify by sonia 子案不列印
            'InsertItem strOldNum, strNewNum, rsTmp.Fields("CP01"), rsTmp.Fields("CP02"), rsTmp.Fields("CP03"), rsTmp.Fields("CP04")
            If rsTmp.Fields("CP04") = "00" Then
               InsertItem strOldNum, strNewNum, rsTmp.Fields("CP01"), rsTmp.Fields("CP02"), rsTmp.Fields("CP03"), rsTmp.Fields("CP04")
            End If
            '2009/3/31 END
            '2008/3/26 end
            'GoTo NEXTRECORD2
            ' 更新代理人成新值
            '93.8.10 modify by sonia
            'strSQL = "UPDATE CASEPROGRESS SET CP44 = '" & strNewNum & "' " & _
            '         "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
            '               "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
            '               "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
            '               "CP04 = '" & rsTmp.Fields("CP04") & "' "
            strSql = "UPDATE CASEPROGRESS SET CP44 = '" & strNewNum & "' " & _
                     "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                           "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                           "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                           "CP04 = '" & rsTmp.Fields("CP04") & "' AND CP44 = '" & strOldNum & "' "
            '93.8.10 END
            cnnConnection.Execute strSql
NEXTRECORD2:
            ' 下一筆
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If

   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更客戶基本檔的申請人編號
'Private Sub ModifyCustomer(ByVal strOldNum As String, ByVal strNewNum As String)
'Modified by Lydia 2020/06/10 + ByVal strModMerge As String
Private Sub ModifyCustomer(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strDelOldNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   Dim strOldCU01 As String
   Dim strOldCU02 As String
   Dim strNewCU01 As String
   Dim strNewCU02 As String
   
   If Len(strOldNum) > 8 Then
      strOldCU01 = Mid(strOldNum, 1, 8)
      strOldCU02 = Mid(strOldNum, 9, 1)
   Else
      strOldCU01 = strOldNum & String(8 - Len(strOldNum), "0")
      strOldCU02 = "0"
   End If
   
   If Len(strNewNum) > 8 Then
      strNewCU01 = Mid(strNewNum, 1, 8)
      strNewCU02 = Mid(strNewNum, 9, 1)
   Else
      strNewCU01 = strNewNum & String(8 - Len(strNewNum), "0")
      strNewCU02 = "0"
   End If
   
   ShowStatus "更新客戶基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
    
   'Added by Lydia 2022/06/24 待活化客戶
   'Added by Lydia 2023/12/28 (112/12/1)待活化客戶系統重新設定規則：判斷是否為關係企業
   'Move by Lydia 2024/05/23 要在更新客戶檔之前，從下方移上來
   If Left(strNewCU01, 1) = "X" Or Left(strOldCU01, 1) = "X" Then
      If Left(strNewCU01, 6) <> Left(strOldCU01, 6) Then
         '非關係企業：判斷新編號是否為已存在的待活化客戶(包含關係企業)，若是則直接變更原編號的待活化客戶記錄為新編號，若不存在則刪除原編號的待活化客戶記錄。
         'Modifiec by Lydia 2024/05/28 (調整)改成以客戶檔+待活化客戶判斷
         'strExc(0) = "SELECT * FROM OLDCUSTOMER WHERE OCU03 IS NULL AND SUBSTR(OCU01,1,6)=" & CNULL(Left(strNewCU01, 6))
         'Modified by Lydia 2025/06/09 只變更代號,不用變更狀態; 拿掉 OCU03 IS NULL AND
         strExc(0) = "SELECT CU01,CU02,OCU01 FROM CUSTOMER, OLDCUSTOMER WHERE CU01=OCU01(+) AND SUBSTR(CU01,1,6)=" & CNULL(Left(strNewCU01, 6))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Added by Lydia 2024/05/28
            If "" & RsTemp.Fields("OCU01") <> "" Then
               If GetCustomerName(strNewCU01 & strNewCU02) = "" Then
                  strSql = "UPDATE OLDCUSTOMER SET OCU01 = '" & strNewCU01 & "' " & _
                           "WHERE OCU01 = '" & strOldCU01 & "' "
                  cnnConnection.Execute strSql
               Else
                  strSql = "DELETE FROM OLDCUSTOMER WHERE OCU01=" & CNULL(strOldCU01)
                  cnnConnection.Execute strSql
               End If
            Else
            'end 2024/05/28
               strSql = "DELETE FROM OLDCUSTOMER WHERE OCU01=" & CNULL(strOldCU01)
               cnnConnection.Execute strSql
            End If
         Else
            strSql = "UPDATE OLDCUSTOMER SET OCU01 = '" & strNewCU01 & "' " & _
                     "WHERE OCU01 = '" & strOldCU01 & "' "
            cnnConnection.Execute strSql
         End If
      Else   '是關係企業編號：直接變更為待活化客戶新編號
      'end 2023/12/28
         'Added by Lydia 2024/02/15 判斷是否存在待活化客戶新編號; ex.X28373050改為X28373010
         'Modifiec by Lydia 2024/05/28 (調整)改成以客戶檔+待活化客戶判斷
         'strExc(0) = "SELECT * FROM OLDCUSTOMER WHERE OCU03 IS NULL AND OCU01=" & CNULL(strNewCU01)
         'Modified by Lydia 2025/06/09 只變更代號,不用變更狀態; 拿掉 OCU03 IS NULL AND
         strExc(0) = "SELECT CU01,CU02,OCU01 FROM CUSTOMER, OLDCUSTOMER WHERE CU01=OCU01(+) AND CU01=" & CNULL(strNewCU01)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "DELETE FROM OLDCUSTOMER WHERE OCU01=" & CNULL(strOldCU01)
            cnnConnection.Execute strSql
         Else
            strSql = "UPDATE OLDCUSTOMER SET OCU01 = '" & strNewCU01 & "' " & _
                     "WHERE OCU01 = '" & strOldCU01 & "' "
            cnnConnection.Execute strSql
         End If

   'Added by Lydia 2023/12/28
      End If
   End If
   'end 2023/12/28
   
   If Mid(strNewNum, 1, 1) = "X" Then
      strSql = "SELECT * FROM CUSTOMER " & _
               "WHERE CU01 = '" & strNewCU01 & "' AND " & _
                     "CU02 = '" & strNewCU02 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If m_bolReNo = False Then   'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更; EX. X25878040>>X25878041
            'rsTmp.MoveFirst 'Remove by Lydia 2019/05/28
            ' 91.01.28 modify by louis (已存在則不管)
            'InsertItem strOldNum, strNewNum, strOldNum
            'strSQL = "UPDATE CUSTOMER SET CU01 = '" & strNewCU01 & "', " & _
            '                             "CU02 = '" & strNewCU02 & "' " & _
            '         "WHERE CU01 = '" & strOldCU01 & "' AND " & _
            '               "CU02 = '" & strOldCU02 & "' "
            'cnnConnection.Execute strSQL
            
            'Added by Lydia 2019/05/28 接洽人資料合併
            'Modified by Lydia 2020/06/10
            'If Me.TextMerge = "Y" Then
            If strModMerge = "Y" Then
JumpToMergeCu:    'Added by Lydia 2024/05/28
                strTemp = GetMaxNo(strNewCU01)
               'Added by Lydia 2024/05/10 聯絡人相片
                strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') as pcc02 from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   RsTemp.MoveFirst
                   Do While Not RsTemp.EOF
                      strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "2") & "' " & _
                               ",IBF03='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                               "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                      cnnConnection.Execute strSql
                      RsTemp.MoveNext
                   Loop
                End If
                'end 2024/05/10
                strSql = "UPDATE POTCUSTCONT SET PCC01 = '" & strNewCU01 & "' " & _
                                         ", PCC02=LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') " & _
                                         "WHERE PCC01 = '" & strOldCU01 & "' "
                Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
                cnnConnection.Execute strSql
                Call UpdCR04(strOldCU01, strNewCU01, strTemp) 'Added by Lydia 2019/05/30 更新往來記錄的連絡人
                'Added by Lydia 2024/12/12 變更個案聯絡人; 申請人留到基本檔再變更
                strSql = "UPDATE PATENT SET PA149=LPAD(PA149+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE PA26='" & strOldCU01 & strOldCU02 & "' AND PA149 IS NOT NULL "
                cnnConnection.Execute strSql
                strSql = "UPDATE TRADEMARK SET TM123=LPAD(TM123+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE TM23='" & strOldCU01 & strOldCU02 & "' AND TM123 IS NOT NULL "
                cnnConnection.Execute strSql
                strSql = "UPDATE SERVICEPRACTICE SET SP78=LPAD(SP78+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE SP08='" & strOldCU01 & strOldCU02 & "' AND SP78 IS NOT NULL "
                cnnConnection.Execute strSql
                strSql = "UPDATE LAWCASE SET LC42=LPAD(LC42+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE LC11='" & strOldCU01 & strOldCU02 & "' AND LC42 IS NOT NULL "
                cnnConnection.Execute strSql
                strSql = "UPDATE HIRECASE SET HC23=LPAD(HC23+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE HC05='" & strOldCU01 & strOldCU02 & "' AND HC23 IS NOT NULL "
                cnnConnection.Execute strSql
                'end 2024/12/12
                
                'Added by Morgan 2025/9/3 副本聯絡人
                strSql = "UPDATE CUSTOMER SET CU167=LPAD(CU167+" & Val(strTemp) & ", 2 ,'0') " & _
                              "WHERE CU166 = '" & strOldNum & "' and cu167 is not null"
                cnnConnection.Execute strSql, intI
                strSql = "UPDATE PATENT SET PA169=LPAD(PA169+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE PA168='" & strOldCU01 & strOldCU02 & "' AND PA169 IS NOT NULL "
                cnnConnection.Execute strSql, intI
                strSql = "UPDATE TRADEMARK SET TM133=LPAD(TM133+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE TM132='" & strOldCU01 & strOldCU02 & "' AND TM133 IS NOT NULL "
                cnnConnection.Execute strSql, intI
                strSql = "UPDATE SERVICEPRACTICE SET SP87=LPAD(SP87+" & Val(strTemp) & ", 2 ,'0') " & _
                         "WHERE SP86='" & strOldCU01 & strOldCU02 & "' AND SP87 IS NOT NULL "
                cnnConnection.Execute strSql, intI
                'end 2025/9/3
                
            End If
         End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      Else
            strSql = "UPDATE CUSTOMER SET CU01 = '" & strNewCU01 & "', " & _
                                         "CU02 = '" & strNewCU02 & "' " & _
                     "WHERE CU01 = '" & strOldCU01 & "' AND " & _
                           "CU02 = '" & strOldCU02 & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            If m_bolReNo = False Then 'Added by Lydia 2025/07/23 更名前後的編號(前8碼相同)不用變更
               'Added by Lydia 2024/05/28 檢查是否存在聯絡人資料
               If GetMaxNo(strNewCU01) <> "" Then
                  GoTo JumpToMergeCu
               Else
               'end 2024/05/28
                     'Added by Lydia 2024/05/10 聯絡人相片
                     strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,pcc02 from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        RsTemp.MoveFirst
                        Do While Not RsTemp.EOF
                           strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "2") & "' " & _
                                    ",IBF03='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                                    "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                           cnnConnection.Execute strSql
                           RsTemp.MoveNext
                        Loop
                     End If
                     'end 2024/05/10
                     '2009/3/31 ADD BY SONIA 同時更新聯絡人資料
                     strSql = "UPDATE PotCustCont SET PCC01 = '" & strNewCU01 & "' " & _
                                               "WHERE PCC01 = '" & strOldCU01 & "' "
                     Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
                     cnnConnection.Execute strSql
                     '2009/3/31 END
               End If
            End If  'Added by Lydia 2025/07/23
      End If
      rsTmp.Close
   End If
   
   '93.2.14 add by sonia
'   If TextDelete = "Y" Then
   If strDelOldNum = "Y" Then
      strSql = "DELETE CUSTOMER WHERE CU01 = '" & strOldCU01 & "' AND " & _
                     "CU02 = '" & strOldCU02 & "' "
      Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
      cnnConnection.Execute strSql
      If m_bolReNo = False Then   'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
         '2011/10/12 ADD BY SONIA 同時刪除接洽人
         If strOldCU02 = "0" Then  'add by sonia 2018/4/3  非更名前的舊名稱編號才可刪除
            'Added by Lydia 2024/05/10 刪除聯絡人相片
            strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
                  strSql = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
                  cnnConnection.Execute strSql
                  RsTemp.MoveNext
               Loop
            End If
            'end 2024/05/10
            strSql = "DELETE POTCUSTCONT WHERE PCC01 = '" & strOldCU01 & "' "
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         '2011/10/12 END
      End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   End If
   '93.2.14 end
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   '93.2.19 ADD BY SONIA
      strSql = "UPDATE CUSTOMER SET CU03 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE CU03 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   '專利固定請款對象
   strSql = "UPDATE CUSTOMER SET CU57 = '" & strNewNum & "' " & _
                  "WHERE CU57 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU71 = '" & strNewNum & "' " & _
                  "WHERE CU71 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU94 = '" & strNewNum & "' " & _
                  "WHERE CU94 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU96 = '" & strNewNum & "' " & _
                  "WHERE CU96 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU97 = '" & strNewNum & "' " & _
                  "WHERE CU97 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU98 = '" & strNewNum & "' " & _
                  "WHERE CU98 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CUSTOMER SET CU99 = '" & strNewNum & "' " & _
                  "WHERE CU99 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '專利固定列印對象
   strSql = "UPDATE CUSTOMER SET CU105 = '" & strNewNum & "' " & _
                  "WHERE CU105 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '年費D/N列印對象
   strSql = "UPDATE CUSTOMER SET CU106 = '" & strNewNum & "' " & _
                  "WHERE CU106 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   'Add By Sindy 2011/3/3
   '商標固定請款對象
   strSql = "UPDATE CUSTOMER SET CU147 = '" & strNewNum & "' " & _
                  "WHERE CU147 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '商標固定列印對象
   strSql = "UPDATE CUSTOMER SET CU151 = '" & strNewNum & "' " & _
                  "WHERE CU151 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '延展D/N列印對象
   strSql = "UPDATE CUSTOMER SET CU152 = '" & strNewNum & "' " & _
                  "WHERE CU152 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2011/3/3 End
   
   'Added by Morgan 2017/3/14 國內副本收件人
   strSql = "UPDATE CUSTOMER SET CU166 = '" & strNewNum & "' " & _
                  "WHERE CU166 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   'end 2017/3/14
   
   'Added by Lydia 2021/11/25 法律所案源資料:介紹客戶編號
   strSql = "UPDATE LAWOFFICESOURCE SET LOS05 = '" & strNewNum & "' " & _
               "WHERE LOS05 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   'end 2021/11/25
   
   '2009/6/30 ADD BY SONIA
   strSql = "UPDATE PotCustomer SET PCU47 = '" & strNewNum & "' " & _
                  "WHERE PCU47 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE PotCustomer1 SET POC16 = '" & strNewNum & "' " & _
                  "WHERE POC16 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2009/6/30 END
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      'Added by Lydia 2016/10/28 申請人指定國外代理人檔
      strSql = "UPDATE CustAssignAgent SET CAA02 = '" & Left(strNewNum, 8) & "' " & _
               "WHERE CAA02 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
      'end 2016/10/28
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   If Left(strOldNum, 8) <> Left(strNewNum, 8) Then 'Added by Lydia 2021/10/27 增加判斷是否為更名後的號碼;  ex.刪除代理人Y34013002一併刪除各項指示
        'Added by Lydia 2016/11/30 各項指示檔
        strSql = "UPDATE INSTRUCTIONS SET ITS01 = '" & Pub_GetITS01Type(Left(strNewNum, 8)) & "', ITS02='" & Left(strNewNum, 8) & "' " & _
                 "WHERE ITS01 = '" & Pub_GetITS01Type(Left(strOldNum, 8)) & "' AND ITS02='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'Added by Lydia 2022/11/25 各項指示檔Instructions.ITS13複製對象編號
        strSql = "UPDATE INSTRUCTIONS SET ITS13 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE ITS13 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2022/11/25
        'Added by Lydia 2016/11/30 國外部關聯企業資料檔
        strSql = "UPDATE FRELATION SET FR01 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE FR01 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        strSql = "UPDATE FRELATION SET FR02 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE FR02 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2016/11/30
        'Added by Lydia 2019/12/23 更新-申請人代理人特殊權限檔CUFA_Right
        strSql = "UPDATE CUFA_RIGHT SET CFR01 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE CFR01 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2019/12/23
   End If 'Added by Lydia 2021/10/27
   
   'Added by Lydia 2018/11/28 更新備註檔
   '------分成6碼和8碼
   '下一程序固定備註(NpMemo)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Memo byb Lydia 2025/01/13 --- 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE NPMEMO SET NM05='" & Left(strNewNum, 6) & "' WHERE NM05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
      strSql = "UPDATE NPMEMO SET NM05='" & Left(strNewNum, 8) & "' WHERE NM05='" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   '核准函輸入備註(ApprovalMemo2)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
    '    strSql = "UPDATE APPROVALMEMO2 SET AM05='" & Left(strNewNum, 6) & "' WHERE AM05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
       strSql = "UPDATE APPROVALMEMO2 SET AM05='" & Left(strNewNum, 8) & "' WHERE AM05='" & Left(strOldNum, 8) & "' "
       cnnConnection.Execute strSql
   End If   'Added by Lydia 2021/10/27
   '核駁及審查意見通知函備註(IncomMemo)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE INCOMMEMO SET IM05='" & Left(strNewNum, 6) & "' WHERE IM05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
       strSql = "UPDATE INCOMMEMO SET IM05='" & Left(strNewNum, 8) & "' WHERE IM05='" & Left(strOldNum, 8) & "' "
       cnnConnection.Execute strSql
   End If   'Added by Lydia 2021/10/27
   '請款函預設備註維護檔(DebitNotePS)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE DEBITNOTEPS SET DNPS05='" & Left(strNewNum, 6) & "' WHERE DNPS05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE DEBITNOTEPS SET DNPS05='" & Left(strNewNum, 8) & "' WHERE DNPS05='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If   'Added by Lydia 2021/10/27
   'end 2018/11/28
   'Added by Lydia 2019/03/11 FCP承辦單設定維護(FcpEMPbill)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE FCPEMPBILL SET FEB05='" & Left(strNewNum, 6) & "' WHERE FEB05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
       strSql = "UPDATE FCPEMPBILL SET FEB05='" & Left(strNewNum, 8) & "' WHERE FEB05='" & Left(strOldNum, 8) & "' "
       cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   'Added by Lydia 2019/03/11 通知告准加註(ApprovalPS)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE APPROVALPS SET APS05='" & Left(strNewNum, 6) & "' WHERE APS05='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
       strSql = "UPDATE APPROVALPS SET APS05='" & Left(strNewNum, 8) & "' WHERE APS05='" & Left(strOldNum, 8) & "' "
       cnnConnection.Execute strSql
   End If  'Added by Lydia 2021/10/27
   
   'Add By Sindy 2025/6/18
   '定稿特殊請款文字維護檔(LetterSetText)
   If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
      strSql = "UPDATE LetterSetText SET LST02='" & Left(strNewNum, 6) & "' WHERE LST02='" & Left(strOldNum, 6) & "' "
      Pub_SeekTbLog strSql '新增維護記錄檔
      cnnConnection.Execute strSql
   End If
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then
      strSql = "UPDATE LetterSetText SET LST02='" & Left(strNewNum, 8) & "' WHERE LST02='" & Left(strOldNum, 8) & "' "
      Pub_SeekTbLog strSql '新增維護記錄檔
      cnnConnection.Execute strSql
   End If
   '2025/6/18 END
   
   'Added by Lydia 2020/02/04 集團客戶應收帳款收文檢查上限CustRecAmtLmt
   If Left(strNewNum, 6) <> Left(strOldNum, 6) And Mid(strOldNum, 7, 3) = "000" And Mid(strNewNum, 7, 3) = "000" Then
       strSql = "UPDATE CUSTRECAMTLMT SET CRA01='" & Left(strNewNum, 6) & "' WHERE CRA01='" & Left(strOldNum, 6) & "' "
       Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
       cnnConnection.Execute strSql
   End If
   'Added by Lydia 2022/03/28 DHL輸入資料
    strSql = "UPDATE  DHL_INPUT_DATA SET DID01 = '" & strNewCU01 & "', DID02 = '" & strNewCU02 & "' " & _
             "WHERE DID01 = '" & strOldCU01 & "' AND DID02 = '" & strOldCU02 & "' "
    cnnConnection.Execute strSql
     
   'Added by Morgan 2022/10/19
   'LEDES設定
   strSql = "UPDATE ledes SET ld01 = '" & strNewNum & "' " & _
             "WHERE ld01 = '" & strOldNum & "' "
   cnnConnection.Execute strSql, intI
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      '客製化請款項目
      strSql = "UPDATE acc260 SET a2601 = '" & strNewCU01 & "' " & _
               "WHERE a2601 = '" & strOldCU01 & "' "
      cnnConnection.Execute strSql, intI
      'end 2022/10/19
      'Added by Morgan 2023/5/3
      '客戶承辦工程師對照檔
      strSql = "UPDATE CustEngMap SET CEM01 = '" & strNewCU01 & "' " & _
               "WHERE CEM01 = '" & strOldCU01 & "' "
      cnnConnection.Execute strSql, intI
      'end 2023/5/3
      'Added by Lydia 2025/07/24 不得代理案件之客戶或代理人NOTAGENT：管制對象(可多個)X/Y編號8碼。
      strSql = "UPDATE NOTAGENT SET NT35=REPLACE(UPPER(NT35),UPPER('" & strOldCU01 & "'), UPPER('" & strNewCU01 & "')) WHERE INSTR(UPPER(NT35),UPPER('" & strOldCU01 & "')) > 0"
      cnnConnection.Execute strSql, intI
      'end 2025/07/24
   End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   'Added by Lydia 2023/06/18 名條清單AddressA4List:兼具特殊控制
   strSql = "UPDATE ADDRESSA4LIST SET AAL04 = '" & strNewCU01 & strNewCU02 & "' " & _
            "WHERE AAL04 = '" & strOldCU01 & strOldCU02 & "' "
   cnnConnection.Execute strSql, intI
   '客戶平台CustWeb: CW04為串接的客戶編號
   strSql = "UPDATE CUSTWEB SET CW04 = REPLACE(UPPER(CW04), UPPER('" & strOldCU01 & strOldCU02 & "'), UPPER('" & strNewCU01 & strNewCU02 & "')) " & _
            "WHERE INSTR(UPPER(CW04), UPPER('" & strOldCU01 & strOldCU02 & "')) > 0 "
   cnnConnection.Execute strSql, intI
   '客戶平台帳號CustWebID
   strSql = "UPDATE CUSTWEBID SET CD02 = '" & strNewCU01 & strNewCU02 & "' " & _
            "WHERE CD02 = '" & strOldCU01 & strOldCU02 & "' "
   cnnConnection.Execute strSql, intI
   'end 2023/06/18
     
   'Added by Lydia 2024/02/05 外專案件清單Excel
   strSql = "UPDATE FCPELISTREC SET FER06 = REPLACE(FER06,'" & strOldCU01 & strOldCU02 & "','" & strNewCU01 & strNewCU02 & "') " & _
            "WHERE INSTR(FER06,'" & strOldCU01 & strOldCU02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   strSql = "UPDATE FCPELISTREC SET FER07 = REPLACE(FER07,'" & strOldCU01 & strOldCU02 & "','" & strNewCU01 & strNewCU02 & "') " & _
            "WHERE INSTR(FER07,'" & strOldCU01 & strOldCU02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   strSql = "UPDATE FCPELISTREC SET FER10 = REPLACE(FER10,'" & strOldCU01 & strOldCU02 & "','" & strNewCU01 & strNewCU02 & "') " & _
            "WHERE INSTR(FER10,'" & strOldCU01 & strOldCU02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   'end 2024/02/05
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更國外代理人基本檔的代理人編號
'Modified by Lydia 2020/06/10 + ByVal strModMerge As String
Private Sub ModifyFAgent(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strDelOldNum As String, ByVal strModMerge As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strTemp As String
Dim strOldFA01 As String
Dim strOldFA02 As String
Dim strNewFA01 As String
Dim strNewFA02 As String
   
   If Len(strOldNum) > 8 Then
      strOldFA01 = Mid(strOldNum, 1, 8)
      strOldFA02 = Mid(strOldNum, 9, 1)
   Else
      strOldFA01 = strOldNum & String(8 - Len(strOldNum), "0")
      strOldFA02 = "0"
   End If
   
   If Len(strNewNum) > 8 Then
      strNewFA01 = Mid(strNewNum, 1, 8)
      strNewFA02 = Mid(strNewNum, 9, 1)
   Else
      strNewFA01 = strNewNum & String(8 - Len(strNewNum), "0")
      strNewFA02 = "0"
   End If
   
   ShowStatus "更新國外代理人基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
   If Mid(strNewNum, 1, 1) = "Y" Then
      strSql = "SELECT * FROM FAGENT " & _
               "WHERE FA01 = '" & strNewFA01 & "' AND " & _
                     "FA02 = '" & strNewFA02 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
         If rsTmp.RecordCount > 0 Then
         'rsTmp.MoveFirst 'Remove by Lydia 2019/05/28
         ' 91.01.28 modify by louis (已存在則不管)
         'InsertItem strOldNum, strNewNum, strOldNum
         'strSQL = "UPDATE FAGENT SET FA01 = '" & strNewFA01 & "', " & _
         '                           "FA02 = '" & strNewFA02 & "' " & _
         '         "WHERE FA01 = '" & strOldFA01 & "' AND " & _
         '               "FA02 = '" & strOldFA02 & "' "
         'cnnConnection.Execute strSQL
         
         'Added by Lydia 2019/05/28 接洽人資料合併
         'Modified by Lydia 2020/06/10
         'If Me.TextMerge = "Y" Then
            If strModMerge = "Y" Then
JumpToMergeFA:      'Added by Lydia 2024/05/28
                strTemp = GetMaxNo(strNewFA01)
               'Added by Lydia 2024/05/10 聯絡人相片
                strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') as pcc02 from potcustcont,imgbytefile where pcc01='" & strOldFA01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   RsTemp.MoveFirst
                   Do While Not RsTemp.EOF
                      strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "2") & "' " & _
                               ",IBF03='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                               "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                      cnnConnection.Execute strSql
                      RsTemp.MoveNext
                   Loop
                End If
                'end 2024/05/10
                strSql = "UPDATE POTCUSTCONT SET PCC01 = '" & strNewFA01 & "' " & _
                                         ", PCC02=LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') " & _
                                         "WHERE PCC01 = '" & strOldFA01 & "' "
                Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
                cnnConnection.Execute strSql
                Call UpdCR04(strOldFA01, strNewFA01, strTemp) 'Added by Lydia 2019/05/30 更新往來記錄的連絡人
            End If
         Else
            strSql = "UPDATE FAGENT SET FA01 = '" & strNewFA01 & "', " & _
                                       "FA02 = '" & strNewFA02 & "' " & _
                     "WHERE FA01 = '" & strOldFA01 & "' AND " & _
                           "FA02 = '" & strOldFA02 & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
   
            'Added by Lydia 2024/05/28 檢查是否存在聯絡人資料
            If GetMaxNo(strNewFA01) <> "" Then
               strNewFagent = True
               GoTo JumpToMergeFA
            Else
            'end 2024/05/28
               'Added by Lydia 2024/05/10 聯絡人相片
               strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,pcc02 from potcustcont,imgbytefile where pcc01='" & strOldFA01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "2") & "' " & _
                              ",IBF03='" & Pub_GetPCCtoIBF(strNewFA01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                              "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                     cnnConnection.Execute strSql
                     RsTemp.MoveNext
                  Loop
               End If
               'end 2024/05/10
               '2009/3/31 ADD BY SONIA 同時更新聯絡人資料
               strSql = "UPDATE PotCustCont SET PCC01 = '" & strNewFA01 & "' " & _
                                         "WHERE PCC01 = '" & strOldFA01 & "' "
               Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
               cnnConnection.Execute strSql
            '2009/3/31 END
            End If
            strNewFagent = True   '2010/10/13 ADD BY SONIA 同時更新CP139
         End If
      End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      rsTmp.Close
   End If
   
   '93.2.14 add by sonia
'   If TextDelete = "Y" Then
   If strDelOldNum = "Y" Then
      strSql = "DELETE FAGENT WHERE FA01 = '" & strOldFA01 & "' AND " & _
                     "FA02 = '" & strOldFA02 & "' "
      Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
      cnnConnection.Execute strSql
      If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
         'Added by Lydia 2024/05/10 同時刪除接洽人
         If strOldFA02 = "0" Then  '非更名前的舊名稱編號才可刪除
            '刪除聯絡人相片
            strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & strOldFA01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
                  strSql = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
                  cnnConnection.Execute strSql
                  RsTemp.MoveNext
               Loop
            End If
            strSql = "DELETE POTCUSTCONT WHERE PCC01 = '" & strOldFA01 & "' "
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         'end 2024/05/10
      End If  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   End If
   '93.2.14 end
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      '93.2.19 ADD BY SONIA
      strSql = "UPDATE FAGENT SET FA03 = '" & Left(strNewNum, 8) & "' " & _
                     "WHERE FA03 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   '專利固定請款對象
   strSql = "UPDATE FAGENT SET FA30 = '" & strNewNum & "' " & _
                  "WHERE FA30 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA38 = '" & strNewNum & "' " & _
                  "WHERE FA38 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA59 = '" & strNewNum & "' " & _
                  "WHERE FA59 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA61 = '" & strNewNum & "' " & _
                  "WHERE FA61 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA62 = '" & strNewNum & "' " & _
                  "WHERE FA62 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA66 = '" & strNewNum & "' " & _
                  "WHERE FA66 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE FAGENT SET FA67 = '" & strNewNum & "' " & _
                  "WHERE FA67 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '專利固定列印對象
   strSql = "UPDATE FAGENT SET FA71 = '" & strNewNum & "' " & _
                  "WHERE FA71 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '年費D/N列印對象
   strSql = "UPDATE FAGENT SET FA72 = '" & strNewNum & "' " & _
                  "WHERE FA72 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '93.2.19 END
   'Add By Sindy 2011/3/3
   '商標固定請款對象
   strSql = "UPDATE FAGENT SET FA107 = '" & strNewNum & "' " & _
                  "WHERE FA107 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '商標固定列印對象
   strSql = "UPDATE FAGENT SET FA111 = '" & strNewNum & "' " & _
                  "WHERE FA111 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '延展D/N列印對象
   strSql = "UPDATE FAGENT SET FA112 = '" & strNewNum & "' " & _
                  "WHERE FA112 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2011/3/3 End
   '2009/6/30 ADD BY SONIA
   strSql = "UPDATE PotCustomer SET PCU47 = '" & strNewNum & "' " & _
                  "WHERE PCU47 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   '2009/6/30 END
   
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      'Added by Lydia 2016/10/28 申請人指定國外代理人檔
      strSql = "UPDATE CustAssignAgent SET CAA04 = '" & Left(strNewNum, 8) & "' " & _
               "WHERE CAA04 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
      'end 2016/10/28
      
      'Added by Lydia 2016/11/22 國外固定寄催款單代理人檔
      strSql = "UPDATE ACC225 SET A2251 = '" & Left(strNewNum, 8) & "' " & _
               "WHERE A2251 = '" & Left(strOldNum, 8) & "' "
      cnnConnection.Execute strSql
      'end 2016/11/22
      'Added by Lydia 2025/07/24 不得代理案件之客戶或代理人NOTAGENT：管制對象(可多個)X/Y編號8碼。
      strSql = "UPDATE NOTAGENT SET NT35=REPLACE(UPPER(NT35),UPPER('" & strOldFA01 & "'), UPPER('" & strNewFA01 & "')) WHERE INSTR(UPPER(NT35),UPPER('" & strOldFA01 & "')) > 0"
      cnnConnection.Execute strSql, intI
      'end 2025/07/24
   End If 'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
   
   If Left(strOldNum, 8) <> Left(strNewNum, 8) Then 'Added by Lydia 2021/10/27 增加判斷是否為更名後的號碼;  ex.刪除代理人Y34013002一併刪除各項指示
        'Added by Lydia 2016/11/30 各項指示檔
        strSql = "UPDATE INSTRUCTIONS SET ITS01 = '" & Pub_GetITS01Type(Left(strNewNum, 8)) & "', ITS02='" & Left(strNewNum, 8) & "' " & _
                 "WHERE ITS01 = '" & Pub_GetITS01Type(Left(strOldNum, 8)) & "' AND ITS02='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'Added by Lydia 2022/11/25 各項指示檔Instructions.ITS13複製對象編號
        strSql = "UPDATE INSTRUCTIONS SET ITS13 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE ITS13 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2022/11/25
        'Added by Lydia 2016/11/30 國外部關聯企業資料檔
        strSql = "UPDATE FRELATION SET FR01 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE FR01 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        strSql = "UPDATE FRELATION SET FR02 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE FR02 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2016/11/30
        'Added by Lydia 2019/12/23 更新-申請人代理人特殊權限檔CUFA_Right
        strSql = "UPDATE CUFA_RIGHT SET CFR01 = '" & Left(strNewNum, 8) & "' " & _
                 "WHERE CFR01 = '" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
        'end 2019/12/23
   End If 'Added by Lydia 2021/10/27
   
   'Added by Lydia 2018/11/28 更新備註檔
   '------分成6碼和8碼
   '下一程序固定備註(NpMemo)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Memo by Lydia 2025/01/13 --- 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE NPMEMO SET NM04='" & Left(strNewNum, 6) & "' WHERE NM04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE NPMEMO SET NM04='" & Left(strNewNum, 8) & "' WHERE NM04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   '核准函輸入備註(ApprovalMemo2)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE APPROVALMEMO2 SET AM04='" & Left(strNewNum, 6) & "' WHERE AM04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE APPROVALMEMO2 SET AM04='" & Left(strNewNum, 8) & "' WHERE AM04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   '核駁及審查意見通知函備註(IncomMemo)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE INCOMMEMO SET IM04='" & Left(strNewNum, 6) & "' WHERE IM04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE INCOMMEMO SET IM04='" & Left(strNewNum, 8) & "' WHERE IM04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   '請款函預設備註維護檔(DebitNotePS)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE DEBITNOTEPS SET DNPS04='" & Left(strNewNum, 6) & "' WHERE DNPS04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE DEBITNOTEPS SET DNPS04='" & Left(strNewNum, 8) & "' WHERE DNPS04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   'end 2018/11/28
   'Added by Lydia 2019/03/11 FCP承辦單設定維護(FcpEMPbill)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
   '     strSql = "UPDATE FCPEMPBILL SET FEB04='" & Left(strNewNum, 6) & "' WHERE FEB04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE FCPEMPBILL SET FEB04='" & Left(strNewNum, 8) & "' WHERE FEB04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   'Added by Lydia 2019/03/11 通知告准加註(ApprovalPS)
   'Modified by Lydia 2022/09/08 判斷同為母號
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) Then
   'Mark by Lydia 2025/01/13
   'If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
    '    strSql = "UPDATE APPROVALPS SET APS04='" & Left(strNewNum, 6) & "' WHERE APS04='" & Left(strOldNum, 6) & "' "
   '     cnnConnection.Execute strSql
   'End If
   'end 2025/01/13
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then 'Added by Lydia 2021/10/27
        strSql = "UPDATE APPROVALPS SET APS04='" & Left(strNewNum, 8) & "' WHERE APS04='" & Left(strOldNum, 8) & "' "
        cnnConnection.Execute strSql
   End If 'Added by Lydia 2021/10/27
   
   'Add By Sindy 2025/6/18
   '定稿特殊請款文字維護檔(LetterSetText)
   If Left(strNewNum, 6) <> Left(strOldNum, 6) And Len(ChangeCustomerS(strNewNum)) = 6 And Len(ChangeCustomerS(strOldNum)) = 6 Then
      strSql = "UPDATE LetterSetText SET LST01='" & Left(strNewNum, 6) & "' WHERE LST01='" & Left(strOldNum, 6) & "' "
      Pub_SeekTbLog strSql '新增維護記錄檔
      cnnConnection.Execute strSql
   End If
   If Left(strNewNum, 8) <> Left(strOldNum, 8) Then
      strSql = "UPDATE LetterSetText SET LST01='" & Left(strNewNum, 8) & "' WHERE LST01='" & Left(strOldNum, 8) & "' "
      Pub_SeekTbLog strSql '新增維護記錄檔
      cnnConnection.Execute strSql
   End If
   '2025/6/18 END

   'Added by Lydia 2022/03/28 DHL輸入資料
    strSql = "UPDATE  DHL_INPUT_DATA SET DID01 = '" & strNewFA01 & "', DID02 = '" & strNewFA02 & "' " & _
             "WHERE DID01 = '" & strOldFA01 & "' AND DID02 = '" & strOldFA02 & "' "
    cnnConnection.Execute strSql
    
   'Memo by Lydia 2024/05/23 刪除程式碼---待活化客戶OldCustomer
      
   'Added by Morgan 2022/10/19
   'LEDES設定
   strSql = "UPDATE ledes SET ld01 = '" & strNewNum & "' " & _
             "WHERE ld01 = '" & strOldNum & "' "
   cnnConnection.Execute strSql, intI
   If m_bolReNo = False Then  'Added by Lydia 2025/01/13 更名前後的編號(前8碼相同)不用變更
      '客製化請款項目
      strSql = "UPDATE acc260 SET a2601 = '" & strNewFA01 & "' " & _
               "WHERE a2601 = '" & strOldFA01 & "' "
      cnnConnection.Execute strSql, intI
      'end 2022/10/19
   End If 'Added by Lydia 2025/01/13
    
   'Added by Lydia 2023/06/18 名條清單AddressA4List:兼具特殊控制
   strSql = "UPDATE ADDRESSA4LIST SET AAL04 = '" & strNewFA01 & strNewFA02 & "' " & _
            "WHERE AAL04 = '" & strOldFA01 & strOldFA02 & "' "
   cnnConnection.Execute strSql, intI
   '客戶平台CustWeb: CW04為串接的客戶編號
   strSql = "UPDATE CUSTWEB SET CW04 = REPLACE(UPPER(CW04), UPPER('" & strOldFA01 & strOldFA02 & "'), UPPER('" & strNewFA01 & strNewFA02 & "')) " & _
            "WHERE INSTR(UPPER(CW04), UPPER('" & strOldFA01 & strOldFA02 & "')) > 0 "
   cnnConnection.Execute strSql, intI
   '客戶平台帳號CustWebID
   strSql = "UPDATE CUSTWEBID SET CD02 = '" & strNewFA01 & strNewFA02 & "' " & _
            "WHERE CD02 = '" & strOldFA01 & strOldFA02 & "' "
   cnnConnection.Execute strSql, intI
   'end 2023/06/18
   'Added by Lydia 2024/02/05 外專案件清單Excel
   strSql = "UPDATE FCPELISTREC SET FER06 = REPLACE(FER06,'" & strOldFA01 & strOldFA02 & "','" & strNewFA01 & strNewFA02 & "') " & _
            "WHERE INSTR(FER06,'" & strOldFA01 & strOldFA02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   strSql = "UPDATE FCPELISTREC SET FER07 = REPLACE(FER07,'" & strOldFA01 & strOldFA02 & "','" & strNewFA01 & strNewFA02 & "') " & _
            "WHERE INSTR(FER07,'" & strOldFA01 & strOldFA02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   strSql = "UPDATE FCPELISTREC SET FER10 = REPLACE(FER10,'" & strOldFA01 & strOldFA02 & "','" & strNewFA01 & strNewFA02 & "') " & _
            "WHERE INSTR(FER10,'" & strOldFA01 & strOldFA02 & "') > 0 "
   cnnConnection.Execute strSql, intI
   'end 2024/02/05
   
   Set rsTmp = Nothing
End Sub

' 產生資料
Private Function GenerateData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim strOldNum As String
   Dim strNewNum As String
   Dim nX As Integer
   Dim nY As Integer
   Dim strTmp As String
   Dim strMod As String
    'Add By Cheng 2004/03/16
   Dim strModHistory As String
   Dim strModNeverClose As String
   Dim strDelOldNum As String
    'End
   Dim strModMerge As String 'Added by Lydia 2020/06/10
   Dim stMsg As String 'Add by Amy 2024/11/29
   
   GenerateData = False
   '92.3.5 Add By sonia
   On Error GoTo ErrorHandler
   cnnConnection.BeginTrans
   
   For nIndex = 0 To m_ModifiedListCount - 1
      strOldNum = m_ModifiedList(nIndex).OldNum
      strNewNum = m_ModifiedList(nIndex).NewNum
        'Add By Cheng 2004/03/16
      strModHistory = m_ModifiedList(nIndex).ModHistory
      strModNeverClose = m_ModifiedList(nIndex).ModNeverClose
      strDelOldNum = m_ModifiedList(nIndex).DelOldNum
        'End
      strModMerge = m_ModifiedList(nIndex).ModMerge 'Added by Lydia 2020/06/10
      strNewFagent = False    '2010/10/13 ADD BY SONIA
      
      'Added by Lydia 2025/01/13 是否為更名前後的編號(前8碼相同); EX. X25878040>>X25878041
      If Left(ChangeCustomerL(strOldNum), 8) = Left(ChangeCustomerL(strNewNum), 8) Then
         m_bolReNo = True
      Else
         m_bolReNo = False
      End If
      'end 2025/01/13
      
      ' 當修改的是客戶編號時
      If Mid(strOldNum, 1, 1) = "X" Then
         'Add By Sindy 2018/1/11
         ' 修改客戶會計師資料檔
         ShowStatus "變更客戶會計師資料檔中, 請稍候 . . ."
         strSql = "UPDATE ACC490 SET A4901 = '" & strNewNum & "' " & _
                        "WHERE A4901 = '" & strOldNum & "' "
         cnnConnection.Execute strSql
         '2018/1/11 END
         
         ' 修改客戶基本資料檔
         'Modified by Lydia 2020/06/10 +strModMerge
         'Move by Lydia 2024/12/12 為了先更新個案聯絡人，從服務業務基本檔下方移上來
         ModifyCustomer strOldNum, strNewNum, strDelOldNum, strModMerge
         
         ' 修改專利檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyPatent strOldNum, strNewNum, strModMerge
         ' 修改商標基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyTradeMark strOldNum, strNewNum, strModMerge
         ' 修改法務基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyLawCase strOldNum, strNewNum, strModMerge
         ' 修改顧問基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyHireCase strOldNum, strNewNum, strModMerge
         ' 修改服務業務基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyServicePractice strOldNum, strNewNum, strModMerge

         ' 修改國外代理人基本資料檔
         'Modified by Lydia 2020/06/10 +strModMerge
         ModifyFAgent strOldNum, strNewNum, strDelOldNum, strModMerge
         If strModHistory = "Y" Then
            ' 修改其他歷史記錄檔
            ModifyHistoryData strOldNum, strNewNum, strDelOldNum
         End If
         If strModNeverClose = "Y" Then
            ' 修改更新國外未收未付記錄檔
            ModifyNeverClose strOldNum, strNewNum
         End If
         
      ' 當修改的是代理人編號時
      'Modified by Lydia 2019/05/28
      'Else
      ElseIf Mid(strOldNum, 1, 1) = "Y" Then
         ' 修改客戶基本資料檔
         'Modified by Lydia 2020/06/10 +strModMerge
         'Move by Lydia 2024/12/12 為了先更新個案聯絡人，從國外代理人基本檔下方移上來
         ModifyCustomer strOldNum, strNewNum, strDelOldNum, strModMerge
         ' 修改專利檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyPatent strOldNum, strNewNum, strModMerge
         ' 修改商標基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyTradeMark strOldNum, strNewNum, strModMerge
         ' 修改法務基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyLawCase strOldNum, strNewNum, strModMerge
         ' 修改顧問基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyHireCase strOldNum, strNewNum, strModMerge
         ' 修改服務業務基本檔
         'Modified by Lydia 2024/12/12 +strModMerge
         ModifyServicePractice strOldNum, strNewNum, strModMerge
         ' 修改國外代理人基本資料檔
         'Modified by Lydia 2020/06/10 +strModMerge
         ModifyFAgent strOldNum, strNewNum, strDelOldNum, strModMerge
         If strModHistory = "Y" Then
             ' 修改其他歷史記錄檔
            ModifyHistoryData strOldNum, strNewNum, strDelOldNum
         End If
'         If TextUpdate2 = "Y" Then
         If strModNeverClose = "Y" Then
            ' 修改更新國外未收未付記錄檔
            ModifyNeverClose strOldNum, strNewNum
         End If
         
      'Added by Lydia 2019/05/28 潛在客戶
      ElseIf Mid(strOldNum, 1, 1) = "R" Then
         strExc(1) = GetPotCustName(strOldNum, strExc(2))
         If strExc(2) = "1" Then
             ' 修改國外潛在客戶基本資料檔
             'Modified by Lydia 2020/06/10 + strModMerge
             ModifyPotCustomer strOldNum, strNewNum, strDelOldNum, strModMerge
         ElseIf strExc(2) = "2" Then
             ' 修改國內潛在客戶基本資料檔
             'Modified by Lydia 2020/06/10 + strModMerge
             ModifyPotCustomer1 strOldNum, strNewNum, strDelOldNum, strModMerge
         End If
         If strModHistory = "Y" Then
             ' 修改其他歷史記錄檔
            ModifyHistoryData strOldNum, strNewNum, strDelOldNum
         End If
      End If
      'Add by Amy 2024/11/29 +客戶代理人來源資料檔 (先更新XYS02資料,再確認XYNoSource 是否已有資料,再依狀況修改其資料)
      ShowStatus "變更客戶代理人來源資料檔中, 請稍候 . . ."
      'Modify by Amy 2025/07/03 Y00043 要併到 Y56151 會錯,改畫面「原編號基本資料是否刪除] =Y,則刪除XYNoSource,不刪XYNoSource則原資料不需更動-秀玲
      '     更新XYS02資料時,Update XYNoSource Set XYS01=新編號 Where XYS01=舊編號,但新編號已有資料會出現違反唯一限制條件
      If (strOldNumXYSData(0) <> "" And strOldNumXYSData(0) <> "NoData" And strOldNumXYSData(0) = "11" _
        And strOldNumXYSData(0) = strNewNumXYSData(0) And strOldNumXYSData(2) = strNewNumXYSData(2)) Or strDelOrgNo = "Y" Then
         '刪除舊資料
         strExc(8) = SaveXYNoSource(3, Me.Name, Left(strNewNum, 8), , , , , Left(strOldNum, 8))
         If Len(strExc(8)) > 1 Then
            stMsg = strExc(8)
         Else
            '更新XYS02
            strExc(8) = SaveXYNoSource(2, Me.Name, Left(strNewNum, 8), , , , , Left(strOldNum, 8))
            If Len(strExc(8)) > 1 Then stMsg = strExc(8)
         End If
      Else
         stMsg = SaveXYNoSource(5, Me.Name, Left(strNewNum, 8), , , , , Left(strOldNum, 8))
      End If
      If Len(stMsg) > 1 Then
         GoTo ErrorHandler
      End If
      stMsg = ""
      'end 2024/11/29
      
      'Modify By Cheng 2002/12/17
      '不管是否有案件都要印
'      If m_ModifiedList(nIndex).ItemListCount > 0 Then
         GenerateData = True
'      End If
      
      ' 排序
      For nX = 0 To m_ModifiedList(nIndex).ItemListCount - 1
         For nY = nX To m_ModifiedList(nIndex).ItemListCount - 1
            If m_ModifiedList(nIndex).ItemList(nX) > m_ModifiedList(nIndex).ItemList(nY) Then
               strTmp = m_ModifiedList(nIndex).ItemList(nX)
               m_ModifiedList(nIndex).ItemList(nX) = m_ModifiedList(nIndex).ItemList(nY)
               m_ModifiedList(nIndex).ItemList(nY) = strTmp
            End If
         Next nY
      Next nX
   Next nIndex
   
   ' 若有搜尋到資料才印表
   If GenerateData = True Then
      ShowStatus "產生列印資料中,請稍候 . . ."
      PrintData
      ShowStatus Empty
   End If

'92.3.5 Add By sonia
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    'Modify by Amy 2024/11/29
    If Err.Number <> 0 Then
      If stMsg = MsgText(601) Then
         stMsg = Err.Description 'Added by Lydia 2019/05/30
      End If
    End If
    MsgBox stMsg
    'end 2024/11/29
    cnnConnection.RollbackTrans
    GenerateData = False
   
End Function

Private Sub PrintData()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(8) As String
   Dim nType As Integer
   Dim nIndex As Integer
   Dim nCenter As Long
   Dim nLeft As Long
   Dim nRight As Long
   Dim nPos As Long
   Dim nField As Integer
   Dim nSubLimit As Integer
   
   ' 案件記錄欄位只可放三筆
   nSubLimit = 3
      
   ' 紙張大小, 方向
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         'Printer.Orientation = vbPRORLandscape
         Printer.Orientation = vbPRORPortrait
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else:
         Printer.PaperSize = vbPRPSA4
         'Printer.Orientation = vbPRORLandscape
         Printer.Orientation = vbPRORPortrait
   End Select
   
   Printer.Font.Name = "新細明體"
      
   ' 印表頭
   nPage = 1
   PrintPageHeader nPage
   nRow = 1
   
   For nIndex = 0 To m_ModifiedListCount - 1
      ' 判斷該修改編號的案件是否有資料
      If m_ModifiedList(nIndex).ItemListCount <= 0 Then
            'Add By Cheng 2002/12/17
            '若無案件資料, 仍要列印
            ' 若列數超過頁面的高度限制時則換頁
            If nRow > m_ReportDataRows Then
               Printer.NewPage
               nPage = nPage + 1
               PrintPageHeader nPage
               nRow = 1
            End If
            ' 清除欄位
            For nField = 0 To 7: fld(nField) = Empty: Next nField
            ' 新列或換頁時
            'If nPos < 5 Or nRow = 1 Then
            If nPos < nSubLimit Or nRow = 1 Then
               ' 原編號
               fld(0) = m_ModifiedList(nIndex).OldNum
               ' 新編號
               fld(1) = m_ModifiedList(nIndex).NewNum
               ' 客戶代理人名稱
               fld(2) = m_ModifiedList(nIndex).Name
            Else
               fld(0) = Empty
               fld(1) = Empty
               fld(2) = Empty
            End If
            ' 輸出
            For nField = 0 To 5
               Select Case nField
                  'Case 3, 4, 5, 6, 7:
                  Case 3, 4, 5:
                     Printer.CurrentX = m_Field(nField).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
                  Case 2:
                     Printer.CurrentX = (m_Field(nField).Left * m_CharWidth) + 100
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print StrToStr(LeftStr(fld(nField), m_Field(nField).Width), 10)
                  Case Else:
                     nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
               End Select
            Next nField
            ' 列數加一
            nRow = nRow + 1
         GoTo NextRecord
      End If
      
      ' 若列數超過頁面的高度限制時則換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader nPage
         nRow = 1
      End If
      
      ' 清除欄位
      For nField = 0 To 7: fld(nField) = Empty: Next nField
      
      For nPos = 0 To m_ModifiedList(nIndex).ItemListCount - 1
         ' 清除欄位內容
         'If nPos Mod 5 = 0 Then
         If nPos Mod nSubLimit = 0 Then
            ' 清除欄位
            For nField = 0 To 7: fld(nField) = Empty: Next nField
         End If
         
         ' 放入本所案號
         'fld((nPos Mod 5) + 3) = m_ModifiedList(nIndex).ItemList(nPos)
         fld((nPos Mod nSubLimit) + 3) = m_ModifiedList(nIndex).ItemList(nPos)
         
         'If (nPos Mod 5 = 4) Or (nPos = m_ModifiedList(nIndex).ItemListCount - 1) Then
         If (nPos Mod nSubLimit = nSubLimit - 1) Or (nPos = m_ModifiedList(nIndex).ItemListCount - 1) Then
            ' 若列數超過頁面的高度限制時則換頁
            If nRow > m_ReportDataRows Then
               Printer.NewPage
               nPage = nPage + 1
               PrintPageHeader nPage
               nRow = 1
            End If
         
            ' 新列或換頁時
            'If nPos < 5 Or nRow = 1 Then
            If nPos < nSubLimit Or nRow = 1 Then
               ' 原編號
               fld(0) = m_ModifiedList(nIndex).OldNum
               ' 新編號
               fld(1) = m_ModifiedList(nIndex).NewNum
               ' 客戶代理人名稱
               fld(2) = m_ModifiedList(nIndex).Name
            Else
               fld(0) = Empty
               fld(1) = Empty
               fld(2) = Empty
            End If
            
            ' 輸出
            For nField = 0 To 5
               Select Case nField
                  'Case 3, 4, 5, 6, 7:
                  Case 3, 4, 5:
                     Printer.CurrentX = m_Field(nField).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
                  Case 2:
                     
                     Printer.CurrentX = (m_Field(nField).Left * m_CharWidth) + 100
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print StrToStr(LeftStr(fld(nField), m_Field(nField).Width), 10)
                  Case Else:
                     nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
               End Select
            Next nField
            ' 列數加一
            nRow = nRow + 1
         End If
      Next nPos
NextRecord:
      nPos = 0 'Add by Amy 2014/11/11 前一筆有案件資料印完nPos若未設0再跑沒案件資料會空列
   Next nIndex
   
   Printer.EndDoc
End Sub

Private Sub TextDelete_GotFocus()
    TextInverse Me.textDelete
End Sub

Private Sub textOldNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textOldNum_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   textOldNum_2 = Empty
   If IsEmptyText(textOldNum) = False Then
      Select Case Mid(textOldNum, 1, 1)
         Case "X":
            textOldNum_2 = GetCustomerName(textOldNum, 0)
            If textOldNum_2 = "" Then
                Cancel = True
                strTit = "檢核資料"
                strMsg = "客戶/代理人編號錯誤"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textOldNum_GotFocus
            End If
         Case "Y":
            textOldNum_2 = GetFAgentName(textOldNum)
            If textOldNum_2 = "" Then
                Cancel = True
                strTit = "檢核資料"
                strMsg = "客戶/代理人編號錯誤"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textOldNum_GotFocus
            End If
         'Added by Lydia 2019/05/28 潛在客戶
         Case "R":
            textOldNum_2 = GetPotCustName(textOldNum, strTit)
            textOldNum.Tag = ""
            If textOldNum_2 = "" Then
                Cancel = True
                strTit = "檢核資料"
                strMsg = "潛在客戶編號錯誤"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textOldNum_GotFocus
            Else
                textOldNum.Tag = strTit
            End If
         'end 2019/05/28
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入正確的客戶/代理人編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textOldNum_GotFocus
      End Select
      
      Call ChkHasXYSData 'Add by Amy 2022/12/07
   
      '2009/1/17 add by sonia 若有聯絡人提醒人工處理
      If TextMerge <> "Y" Then 'Added by Lydia 2024/12/12
         strExc(0) = "select * from PotCustCont where pcc01='" & Left(textOldNum & "00", 8) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTit = "檢核資料"
            'Modified by Morgan 2023/12/15 下面增加檢查案件的接洽人
            'strMsg = "此編號有聯絡人資料, 請人工核對是否重覆並自行更新或刪除 !! 案件的接洽人也要注意 !"
            strMsg = "此編號有聯絡人資料, 請人工核對是否重覆並自行更新或刪除 !!"
            'end 2023/12/15
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            
            'Added by Morgan 2023/12/15
            If Left(textOldNum, 1) = "X" Then
               strExc(0) = "select pa26,pa149 from patent where pa26='" & textOldNum & "' and pa149 is not null" & _
                           " union select tm23,tm123 from trademark where tm23='" & textOldNum & "' and tm123 is not null" & _
                           " union select lc11,lc42 from lawcase where lc11='" & textOldNum & "' and lc42 is not null" & _
                           " union select hc05,hc23 from hirecase where hc05='" & textOldNum & "' and hc23 is not null" & _
                           " union select sp08,sp78 from servicepractice where sp08='" & textOldNum & "' and sp78 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "此客戶編號個案有設定接洽人，請檢查是否要重新設定!!!", vbExclamation
               End If
            End If
            'end 2023/12/15
            
         End If
         '2009/1/17 end
      End If 'Added by Lydia 2024/12/12
      'Add by Morgan 2010/5/25 有更名過要提醒
      If Mid(textOldNum, 9) = "0" Then
         strExc(0) = "select * from customer where cu01='" & Left(textOldNum, 8) & "' and cu02<>'0'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTit = "檢核資料"
            strMsg = "此編號有更名過, 請確認 !!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
      End If
      'end 2010/5/25
   End If
End Sub

Private Sub textNewNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textNewNum_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textNewNum_2 = Empty
   If IsEmptyText(textNewNum) = False Then
      If Mid(textNewNum, 1, 1) <> Mid(textOldNum, 1, 1) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "客戶/代理人編號新舊編號輸入錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNewNum_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case Mid(textNewNum, 1, 1)
         Case "X":
            textNewNum_2 = GetCustomerName(textNewNum, 0)
         Case "Y":
            textNewNum_2 = GetFAgentName(textNewNum)
         'Added by Lydia 2019/05/28 潛在客戶
         Case "R":
            textNewNum_2 = GetPotCustName(textNewNum, strTit)
            textNewNum.Tag = ""
            If textNewNum_2 = "" Then
                'Modified by Lydia 2019/07/04 新編號只要檢查流水號
                'Cancel = True
                'strTit = "檢核資料"
                'strMsg = "潛在客戶編號錯誤"
                'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                'textNewNum_GotFocus
                If IsOverAutoNumber("R", Empty, Mid(textNewNum, 2, 5)) = True Then
                   Cancel = True
                   MsgBox "客戶代碼超過自動編號! ", vbCritical + vbOKOnly, "檢核資料"
                   textNewNum_GotFocus
                   GoTo EXITSUB
                End If
                'end 2019/07/04
            Else
                textNewNum.Tag = strTit
                If textOldNum.Tag <> textNewNum.Tag Then
                    Cancel = True
                    strTit = "檢核資料"
                    strMsg = "潛在客戶編號新舊編號不可同時輸入國內或國外"
                    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                    textNewNum_GotFocus
                    GoTo EXITSUB
                End If
            End If
         'end 2019/05/28
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入正確的客戶/代理人編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNewNum_GotFocus
      End Select
      '2009/3/30 add by sonia 若新編號之客戶減免身份與舊編號不同則不可執行
      strExc(0) = "select a1.ad03,a2.ad03,a1.ad10,a2.ad10 from ApplicantDiscount a1,ApplicantDiscount a2 where a1.AD01='" & Left(textNewNum & "00", 8) & "' " & _
                  "and '" & Left(textOldNum & "00", 8) & "'=a2.AD01 and a1.ad02=a2.ad02 and " & _
                  "(a1.ad03<>a2.ad03 or a1.ad10<>a2.ad10 or (a1.ad10 is not null and a2.ad10 is null) or (a1.ad10 is null and a2.ad10 is not null)) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "新編號之客戶減免身份與舊編號不同, 請專業部確認 !!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNewNum_GotFocus
      End If
      '2009/3/30 end
   End If
EXITSUB:
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 原客戶/代理人編號
   If IsEmptyText(textOldNum) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入原客戶/代理人編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 新客戶/代理人編號
   If IsEmptyText(textNewNum) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入新客戶/代理人編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 客戶或代理人編號輸入錯誤
   If Mid(textOldNum, 1, 1) <> Mid(textNewNum, 1, 1) Then
      strTit = "檢核資料"
      strMsg = "原客戶/代理人編號與新客戶/代理人編號輸入錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 客戶或代理人新編號不可為原編號
   If (textOldNum & String(9 - Len(textOldNum), "0")) = (textNewNum & String(9 - Len(textNewNum), "0")) Then
      strTit = "檢核資料"
      strMsg = "原客戶/代理人編號與新客戶/代理人編號不可相同"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   
   'Add by Amy 2022/12/07 若存在XYS02介紹來源編號,則不可改
   If ChkHasXYSData = True Then
        GoTo EXITSUB
   End If
   
   'Modify Sindy 2022/5/17 如果原編號沒有會計師資料，就不必檢查新編號的會計師資料。
   strExc(0) = "select A4901 from ACC490 where A4901='" & Left(textOldNum & "000", 9) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   '2022/5/17 END
      'Add By Sindy 2018/1/11
      strExc(0) = "select A4901 from ACC490 where A4901='" & Left(textNewNum & "000", 9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTit = "檢核資料"
         strMsg = "「客戶會計師資料檔」新編號已有資料存在，請先確認資料!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   '2018/1/11 END
   
   '93.2.14 add by sonia
   If textUpdate1 = "Y" And TextUpdate2 = "Y" Then
      strTit = "檢核資料"
      strMsg = "歷史資料是否更新 與 更新國外未收未付資料 不可同時輸入 Y !!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textUpdate1.SetFocus
      GoTo EXITSUB
   End If
   '93.2.14 end
   
   'Add by Morgan 2007/4/26
   If textUpdate1 <> "Y" Then
      strExc(0) = "select cu01 from customer where cu01||cu02='" & Left(textNewNum & "000", 9) & "'"
      strExc(0) = strExc(0) & "union select fa01 from fagent where fa01||fa02='" & Left(textNewNum & "000", 9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         strTit = "檢核資料"
         strMsg = "新編號不存在，歷史資料是否更新必須輸入 Y !!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUpdate1.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2007/4/26
   
   'Added by Lydia 2018/11/28 檢查備註檔在更新代號後，是否會造成重複
   If CheckMemoDual(ChangeCustomerL(textOldNum), ChangeCustomerL(textNewNum)) = True Then
        textOldNum.SetFocus
        GoTo EXITSUB
   End If
   'end 2018/11/28
   'Added by Lydia 2020/05/07 檢查各項指示檔在更新代號後，是否會造成重複
   If CheckInstructionsDual(Left(ChangeCustomerL(textOldNum), 8), Left(ChangeCustomerL(textNewNum), 8)) = True Then
        textOldNum.SetFocus
        GoTo EXITSUB
   End If
   'end 2020/05/07
   
   'Added by Morgan 2023/5/3
   If Left(textNewNum, 1) = "X" Then
      If CheckCustEngMap(Left(ChangeCustomerL(textNewNum), 8)) = True Then
         textOldNum.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2023/5/3
   
   'Added by Lydia 2023/06/18 其他Table檢查: 在更新代號後，是否會造成重複
   If CheckOtherDual(ChangeCustomerL(textOldNum), ChangeCustomerL(textNewNum)) = True Then
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   'end 2023/06/18
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textOldNum_GotFocus()
   InverseTextBox textOldNum
End Sub

Private Sub textNewNum_GotFocus()
   InverseTextBox textNewNum
End Sub

Private Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = strConV(MidB(strConV(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

Private Sub ShowStatus(ByVal strData As String)
   If IsEmptyText(strData) = False Then
      textStatus = "執行 : " & strData
   Else
      textStatus = strData
   End If
   textStatus.Refresh
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textNewNum.Enabled = True Then
   Cancel = False
   textNewNum_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOldNum.Enabled = True Then
   Cancel = False
   textOldNum_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2022/10/24
textOldNum = ChangeCustomerL(textOldNum)
textNewNum = ChangeCustomerL(textNewNum)
If Mid(textOldNum, 1, 1) = "X" Then
   strExc(0) = "SELECT * FROM ConsultRecApp,flow003 WHERE CRA05='" & Left(textOldNum, 8) & "' AND CRA01=f0301 AND f0309 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "此客戶編號尚有接洽單未收文，請確認 !", vbCritical
      Cancel = True
      Exit Function
   End If
   strExc(0) = "SELECT * FROM ConsultRecInv,flow003 WHERE CRi03='" & Left(textOldNum, 8) & "' AND CRi01=f0301 AND f0309 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "此客戶編號(發明人資料檔)尚有接洽單未收文，請確認 !", vbCritical
      Cancel = True
      Exit Function
   End If
   
ElseIf Mid(textOldNum, 1, 1) = "Y" Then
   strExc(0) = "SELECT * FROM ConsultRecordList,flow003 WHERE CRL60='" & Left(textOldNum, 8) & "' AND CRL01=f0301 AND f0309 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "此代理人編號尚有接洽單未收文，請確認 !", vbCritical
      Cancel = True
      Exit Function
   End If
End If
'2022/10/24 END

TxtValidate = True
End Function

Private Sub textUpdate1_GotFocus()
    TextInverse Me.textUpdate1
End Sub

Private Sub textupdate1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If

End Sub

Private Sub TextUpdate2_GotFocus()
    TextInverse Me.TextUpdate2
End Sub

Private Sub textupdate2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If

End Sub

Private Sub textDelete_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If

End Sub

'Memo by Lydia 2022/09/08 日常工作109項: 抓外專特殊設定備註檔\\LINUX\PolyCOM\TaieNew\Script\外專特殊備註-XY編號合併清單.txt
'Added by Lydia 2018/11/28 檢查備註檔在更新代號後，是否會造成重複
Public Function CheckMemoDual(ByVal pOldNo As String, ByVal pNewNo As String) As Boolean
Dim intJ As Integer
Dim rsR1 As New ADODB.Recordset
Dim strCon1 As String, strCon2 As String, strCon3 As String
Dim strCon4 As String 'Add By Sindy 2025/6/18
Dim strMsg As String

     Select Case Left(pOldNo, 1)
        Case "X", "Y"
              '依代理人編號,6碼和8碼分開找
              strCon3 = "代理人編號: "
              'Added by Lydia 2022/09/06
              pOldNo = ChangeCustomerS(pOldNo)
              pNewNo = ChangeCustomerS(pNewNo)
              '下一程序固定備註(NpMemo)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                  strCon1 = "union select '下一程序固定備註(NpMemo)' as pTitle, " & CNULL(pNewNo) & "||nm05||nm06 as nowno,1 as cnt, nm01 as sno,nm04 as oldno from npmemo where nm04='" & pOldNo & "' " & _
                                 "union select '下一程序固定備註(NpMemo)' as pTitle,nm04||nm05||nm06 as nowno,1 as cnt, nm01 as sno,nm04 as oldno from npmemo where nm04='" & pNewNo & "' "
              
              Else  'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                  strCon1 = strCon1 & "union select '下一程序固定備註(NpMemo)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||nm05||nm06 as nowno,1 as cnt, nm01 as sno,nm04 as oldno from npmemo where nm04='" & Left(pOldNo, 8) & "' " & _
                                "union select '下一程序固定備註(NpMemo)' as pTitle, nm04||nm05||nm06 as nowno,1 as cnt, nm01 as sno,nm04 as oldno from npmemo where nm04='" & Left(pNewNo, 8) & "' "
              End If   'Added by Lydia 2022/09/08

              '核准函輸入備註(ApprovalMemo2)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                 strCon1 = strCon1 & "union select '核准函輸入備註(ApprovalMemo2)' as pTitle, " & CNULL(pNewNo) & "||am05||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am04 as oldno from ApprovalMemo2 where am04='" & pOldNo & "' " & _
                                "union select '核准函輸入備註(ApprovalMemo2)' as pTitle, am04||am05||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am04 as oldno from ApprovalMemo2 where am04='" & pNewNo & "' "
                                
              Else  'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                 strCon1 = strCon1 & "union select '核准函輸入備註(ApprovalMemo2)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||am05||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am04 as oldno from ApprovalMemo2 where am04='" & Left(pOldNo, 8) & "' " & _
                                "union select '核准函輸入備註(ApprovalMemo2)' as pTitle, am04||am05||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am04 as oldno from ApprovalMemo2 where am04='" & Left(pNewNo, 8) & "' "
              End If

              '核駁及審查意見通知函備註(IncomMemo)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                  strCon1 = strCon1 & "union select '核駁及審查意見通知函備註(IncomMemo)' as pTitle, " & CNULL(Left(pNewNo, 6)) & "||im05||im06 as nowno,1 as cnt, im01 as sno,im04 as oldno from IncomMemo where im04='" & Left(pOldNo, 6) & "' " & _
                                "union select '核駁及審查意見通知函備註(IncomMemo)' as pTitle,im04||im05||im06 as nowno,1 as cnt, im01 as sno,im04 as oldno from IncomMemo where im04='" & Left(pNewNo, 6) & "' "
              Else 'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                  strCon1 = strCon1 & "union select '核駁及審查意見通知函備註(IncomMemo)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||im05||im06 as nowno,1 as cnt, im01 as sno,im04 as oldno from IncomMemo where im04='" & Left(pOldNo, 8) & "' " & _
                                "union select '核駁及審查意見通知函備註(IncomMemo)' as pTitle,im04||im05||im06 as nowno,1 as cnt, im01 as sno,im04 as oldno from IncomMemo where im04='" & Left(pNewNo, 8) & "' "
              End If 'Added by Lydia 2022/09/08
              
              '請款函預設備註維護檔(DebitNotePS)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                 strCon1 = strCon1 & "union select '請款函預設備註(DebitNotePS)' as pTitle, " & CNULL(Left(pNewNo, 6)) & "||dnps05 as nowno,1 as cnt, dnps01 as sno,dnps04 as oldno from DebitNotePS where dnps04='" & Left(pOldNo, 6) & "' " & _
                               "union select '請款函預設備註(DebitNotePS)' as pTitle,dnps04||dnps05 as nowno,1 as cnt, dnps01 as sno,dnps04 as oldno from DebitNotePS where dnps04='" & Left(pNewNo, 6) & "' "
              Else 'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                 strCon1 = strCon1 & "union select '請款函預設備註(DebitNotePS)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||dnps05 as nowno,1 as cnt, dnps01 as sno,dnps04 as oldno from DebitNotePS where dnps04='" & Left(pOldNo, 8) & "' " & _
                                "union select '請款函預設備註(DebitNotePS)' as pTitle,dnps04||dnps05 as nowno,1 as cnt, dnps01 as sno,dnps04 as oldno from DebitNotePS where dnps04='" & Left(pNewNo, 8) & "' "
              End If 'Added by Lydia 2022/09/08
              
              'Added by Lydia 2019/03/11 FCP承辦單設定維護(FcpEMPbill)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                 strCon1 = strCon1 & "union select '承辦單設定維護(FcpEMPbill)' as pTitle, " & CNULL(Left(pNewNo, 6)) & "||feb05||feb06 as nowno,1 as cnt, feb01 as sno,feb04 as oldno from FcpEMPbill where feb04='" & Left(pOldNo, 6) & "' " & _
                                "union select '承辦單設定維護(FcpEMPbill)' as pTitle,feb04||feb05||feb06 as nowno,1 as cnt, feb01 as sno,feb04 as oldno from FcpEMPbill where feb04='" & Left(pNewNo, 6) & "' "
              Else 'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                 strCon1 = strCon1 & "union select '承辦單設定維護(FcpEMPbill)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||feb05||feb06 as nowno,1 as cnt, feb01 as sno,feb04 as oldno from FcpEMPbill where feb04='" & Left(pOldNo, 8) & "' " & _
                                "union select '承辦單設定維護(FcpEMPbill)' as pTitle,feb04||feb05||feb06 as nowno,1 as cnt, feb01 as sno,feb04 as oldno from FcpEMPbill where feb04='" & Left(pNewNo, 8) & "' "
              End If 'Added by Lydia 2022/09/08
              
              'Added by Lydia 2019/03/11 通知告准加註(ApprovalPS)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then 'Added by Lydia 2022/09/08 判斷同為母號才抓前6碼
                 strCon1 = strCon1 & "union select '通知告准加註(ApprovalPS)' as pTitle, " & CNULL(Left(pNewNo, 6)) & "||aps05 as nowno,1 as cnt, aps01 as sno,aps04 as oldno from ApprovalPS where aps04='" & Left(pOldNo, 6) & "' " & _
                                "union select '通知告准加註(ApprovalPS)' as pTitle,aps04||aps05 as nowno,1 as cnt, aps01 as sno,aps04 as oldno from ApprovalPS where aps04='" & Left(pNewNo, 6) & "' "
              Else 'Added by Lydia 2022/09/08 依號碼長度抓前X碼
                 strCon1 = strCon1 & "union select '通知告准加註(ApprovalPS)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||aps05 as nowno,1 as cnt, aps01 as sno,aps04 as oldno from ApprovalPS where aps04='" & Left(pOldNo, 8) & "' " & _
                                "union select '通知告准加註(ApprovalPS)' as pTitle,aps04||aps05 as nowno,1 as cnt, aps01 as sno,aps04 as oldno from ApprovalPS where aps04='" & Left(pNewNo, 8) & "' "
              End If 'Added by Lydia 2022/09/08
 '-------------------------------------------------------
              If Left(pOldNo, 1) = "X" Then '客戶
                  strCon3 = "客戶編號: "
                  strCon1 = Replace(strCon1, "04", "AA")
                  strCon1 = Replace(strCon1, "05", "04")
                  strCon1 = Replace(strCon1, "AA", "05")
              End If
'-------------------------------------------------------
              
              'Add By Sindy 2025/6/18 定稿特殊請款文字維護檔(LetterSetText)
              If Len(pOldNo) = 6 And Len(pNewNo) = 6 Then '判斷同為母號才抓前6碼
                  strCon4 = "union select '定稿特殊請款文字維護檔(LetterSetText)' as pTitle, " & CNULL(pNewNo) & "||LST02||LST10 as nowno,1 as cnt, 0 as sno,LST01 as oldno from LetterSetText where LST01='" & pOldNo & "' " & _
                                 "union select '定稿特殊請款文字維護檔(LetterSetText)' as pTitle,LST01||LST02||LST10 as nowno,1 as cnt, 0 as sno,LST01 as oldno from LetterSetText where LST01='" & pNewNo & "' "
              
              Else '依號碼長度抓前X碼
                  strCon4 = "union select '定稿特殊請款文字維護檔(LetterSetText)' as pTitle," & CNULL(Left(pNewNo, 8)) & "||LST02||LST10 as nowno,1 as cnt, 0 as sno,LST01 as oldno from LetterSetText where LST01='" & Left(pOldNo, 8) & "' " & _
                                "union select '定稿特殊請款文字維護檔(LetterSetText)' as pTitle, LST01||LST02||LST10 as nowno,1 as cnt, 0 as sno,LST01 as oldno from LetterSetText where LST01='" & Left(pNewNo, 8) & "' "
              End If
'-------------------------------------------------------
              If Left(pOldNo, 1) = "X" Then '客戶
                  strCon4 = Replace(strCon4, "01", "AA")
                  strCon4 = Replace(strCon4, "02", "01")
                  strCon4 = Replace(strCon4, "AA", "02")
              End If
'-------------------------------------------------------
              '2025/6/18 END
              
        Case Else '本所案號
              strCon3 = "本所案號: "
              '下一程序固定備註(NpMemo)
              strCon1 = "union select '11' as ord1,'下一程序固定備註(NpMemo)' as pTitle, " & CNULL(pNewNo) & "||nm06 as nowno,1 as cnt, nm01 as sno,nm03 as oldno from npmemo where nm03='" & pOldNo & "' " & _
                             "union select '12' as ord1,'下一程序固定備註(NpMemo)' as pTitle,nm03||nm06 as nowno,1 as cnt, nm01 as sno,nm03 as oldno from npmemo where nm03='" & pNewNo & "' "
              '核准函輸入備註(ApprovalMemo2)
              strCon1 = strCon1 & "union select '21' as ord1,'核准函輸入備註(ApprovalMemo2)' as pTitle, " & CNULL(pNewNo) & "||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am03 as oldno from ApprovalMemo2 where am03='" & pOldNo & "' " & _
                             "union select '22' as ord1,'核准函輸入備註(ApprovalMemo2)' as pTitle, am03||am06||decode(am07,'1','3','2','3',am07) as nowno,1 as cnt, am01 as sno,am03 as oldno from ApprovalMemo2 where am03='" & pNewNo & "' "
              '核駁及審查意見通知函備註(IncomMemo)
              strCon1 = strCon1 & "union select '31' as ord1,'核駁及審查意見通知函備註(IncomMemo)' as pTitle, " & CNULL(pNewNo) & "||im06 as nowno,1 as cnt, im01 as sno,im03 as oldno from IncomMemo where im03='" & pOldNo & "' " & _
                             "union select '32' as ord1,'核駁及審查意見通知函備註(IncomMemo)' as pTitle,im03||im06 as nowno,1 as cnt, im01 as sno,im03 as oldno from IncomMemo where im03='" & pNewNo & "' "
              '請款函預設備註維護檔(DebitNotePS)
              strCon1 = strCon1 & "union select '41' as ord1,'請款函預設備註(DebitNotePS)' as pTitle, " & CNULL(pNewNo) & " as nowno,1 as cnt, dnps01 as sno,dnps03 as oldno from DebitNotePS where dnps03='" & pOldNo & "' " & _
                             "union select '42' as ord1,'請款函預設備註(DebitNotePS)' as pTitle,dnps03 as nowno,1 as cnt, dnps01 as sno,dnps03 as oldno from DebitNotePS where dnps03='" & pNewNo & "' "
              'Added by Lydia 2019/03/11 FCP承辦單設定維護(FcpEMPbill)
              strCon1 = strCon1 & "union select '51' as ord1,'承辦單設定維護(FcpEMPbill)' as pTitle, " & CNULL(pNewNo) & "||feb06 as nowno,1 as cnt, feb01 as sno,feb03 as oldno from FcpEMPbill where feb03='" & pOldNo & "' " & _
                             "union select '52' as ord1,'承辦單設定維護(FcpEMPbill)' as pTitle,feb03||feb06 as nowno,1 as cnt, feb01 as sno,feb03 as oldno from FcpEMPbill where feb03='" & pNewNo & "' "
              'Added by Lydia 2019/03/11 通知告准加註(ApprovalPS)
              strCon1 = strCon1 & "union select '61' as ord1,'通知告准加註(ApprovalPS)' as pTitle, " & CNULL(pNewNo) & " as nowno,1 as cnt, aps01 as sno,aps03 as oldno from ApprovalPS where aps03='" & pOldNo & "' " & _
                             "union select '62' as ord1,'通知告准加註(ApprovalPS)' as pTitle,aps03 as nowno,1 as cnt, aps01 as sno,aps03 as oldno from ApprovalPS where aps03='" & pNewNo & "' "
     End Select
     
     'Modify By Sindy 2025/6/18 + & strCon4
     strCon2 = "select pTitle,nowno ,sum(cnt) tot from (" & Mid(strCon1, 7) & strCon4 & ") group by pTitle,nowno having sum(cnt) > 1 order by 1,nowno "
     intJ = 1
     Set rsR1 = ClsLawReadRstMsg(intJ, strCon2)
     If intJ = 1 Then
         rsR1.MoveFirst
         Do While Not rsR1.EOF
              If InStr(strMsg, "" & rsR1.Fields("pTitle")) = 0 Then
                  strMsg = strMsg & IIf(InStr(strMsg, "" & rsR1.Fields("pTitle")) = 0, "" & rsR1.Fields("pTitle") & vbCrLf, "")
                  'Modify By Sindy 2025/6/18 + & strCon4
                  strCon2 = "select * from (" & Mid(strCon1, 7) & strCon4 & ") where pTitle= '" & rsR1.Fields("pTitle") & "' order by sno "
                  intJ = 1
                  Set RsTemp = ClsLawReadRstMsg(intJ, strCon2)
                  If intJ = 1 Then
                      RsTemp.MoveFirst
                      Do While Not RsTemp.EOF
                        'Add By Sindy 2025/6/18
                        If Val("" & RsTemp.Fields("sno")) = 0 Then
                           strMsg = "    " & strMsg & strCon3 & RsTemp.Fields("oldno") & vbCrLf
                        Else
                        '2025/6/18 END
                           strMsg = strMsg & "    流水號: " & RsTemp.Fields("sno") & vbTab & strCon3 & RsTemp.Fields("oldno") & vbCrLf
                        End If
                        RsTemp.MoveNext
                      Loop
                  End If
              End If
              rsR1.MoveNext
         Loop
         If strMsg <> "" Then
             MsgBox "以下備註檔在更新代號會有重複記錄，請先整合記錄！" & vbCrLf & strMsg, vbCritical, "檢核資料"
             CheckMemoDual = True
         End If
     End If
     
     Set rsR1 = Nothing
     
      'Add By Sindy 2025/6/18 定稿特殊請款文字維護檔(LetterSetText)
      '   先檢查是否有此狀況,再看要怎麼調整
      strCon4 = "select * from LetterSetText where instr(LST10,'" & pOldNo & "')>0 "
      intJ = 1
      Set RsTemp = ClsLawReadRstMsg(intJ, strCon4)
      If intJ = 1 Then
         MsgBox "定稿特殊請款文字維護檔(LetterSetText)" & vbCrLf & _
                "有排除的申請人編號(" & RsTemp.Fields("LST10") & ")記錄，尚未遇過要怎麼調整呢？", vbCritical, "檢核資料"
         CheckMemoDual = True
      End If
      '2025/6/18 END
End Function

'Added by Lydia 2019/05/28 取得潛在客戶的名稱
Private Function GetPotCustName(ByVal pNo As String, Optional ByRef iType As String) As String
Dim intR As Integer
Dim rsR1 As New ADODB.Recordset
Dim strB1 As String

     iType = ""
     GetPotCustName = ""
     '國外潛在
     strB1 = "select '1' as pType, nvl(pcu03,nvl(pcu08,pcu07)) pName from potcustomer where pcu01||pcu02='" & ChangeCustomerL(pNo) & "' "
     '國內潛在
     strB1 = strB1 & "union all select '2' as pType,nvl(poc03,nvl(poc23,poc27)) pName from potcustomer1 where poc01||poc02='" & ChangeCustomerL(pNo) & "' "
     
     strB1 = strB1 & " order by 1 "
     intR = 1
     Set rsR1 = ClsLawReadRstMsg(intR, strB1)
     If intR = 1 Then
         GetPotCustName = "" & rsR1.Fields("pname")
         iType = "" & rsR1.Fields("ptype")
     End If
     Set rsR1 = Nothing
End Function

'取得接洽人檔的最大序號
Private Function GetMaxNo(ByVal pKeyNo As String) As String
Dim intB As Integer
Dim rsB As New ADODB.Recordset
Dim strB1 As String

    If Trim(pKeyNo) = "" Then Exit Function
         
    strB1 = "select nvl(max(pcc02),0) mno from potcustcont where pcc01=" & CNULL(Left(ChangeCustomerL(pKeyNo), 8))
    intB = 1
    Set rsB = ClsLawReadRstMsg(intB, strB1)
    If intB = 1 Then
        GetMaxNo = "" & rsB.Fields("mno")
    End If
    
    Set rsB = Nothing
    
End Function

' 變更潛在客戶基本檔的編號
'Modified by Lydia 2020/06/10 + ByVal strModMerge As String
Private Sub ModifyPotCustomer(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strDelOldNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   Dim strOldCU01 As String
   Dim strOldCU02 As String
   Dim strNewCU01 As String
   Dim strNewCU02 As String
   
   If Len(strOldNum) > 8 Then
      strOldCU01 = Mid(strOldNum, 1, 8)
      strOldCU02 = Mid(strOldNum, 9, 1)
   Else
      strOldCU01 = strOldNum & String(8 - Len(strOldNum), "0")
      strOldCU02 = "0"
   End If
   
   If Len(strNewNum) > 8 Then
      strNewCU01 = Mid(strNewNum, 1, 8)
      strNewCU02 = Mid(strNewNum, 9, 1)
   Else
      strNewCU01 = strNewNum & String(8 - Len(strNewNum), "0")
      strNewCU02 = "0"
   End If
   
   ShowStatus "更新潛在客戶基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"

    strSql = "SELECT * FROM POTCUSTOMER " & _
             "WHERE PCU01 = '" & strNewCU01 & "' AND " & _
                   "PCU02 = '" & strNewCU02 & "' "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
         '接洽人資料合併
         'Modified by Lydia 2020/06/10
         'If Me.TextMerge = "Y" Then
         If strModMerge = "Y" Then
JumpToMergePCU: 'Added by Lydia 2024/05/28
             strTemp = GetMaxNo(strNewCU01)
            'Added by Lydia 2024/05/10 聯絡人相片
             strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') as pcc02 from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                RsTemp.MoveFirst
                Do While Not RsTemp.EOF
                   strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "2") & "' " & _
                            ",IBF03='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                            "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                   cnnConnection.Execute strSql
                   RsTemp.MoveNext
                Loop
             End If
             'end 2024/05/10
             strSql = "UPDATE POTCUSTCONT SET PCC01 = '" & strNewCU01 & "' " & _
                                      ", PCC02=LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') " & _
                                      "WHERE PCC01 = '" & strOldCU01 & "' "
             Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
             cnnConnection.Execute strSql
             Call UpdCR04(strOldCU01, strNewCU01, strTemp) 'Added by Lydia 2019/05/30 更新往來記錄的連絡人
         End If
    Else
       strSql = "UPDATE POTCUSTOMER SET PCU01 = '" & strNewCU01 & "', " & _
                                    "PCU02 = '" & strNewCU02 & "' " & _
                "WHERE PCU01 = '" & strOldCU01 & "' AND " & _
                      "PCU02 = '" & strOldCU02 & "' "
       Pub_SeekTbLog strSql   '新增維護記錄檔
       cnnConnection.Execute strSql

       'Added by Lydia 2024/05/28 檢查是否存在聯絡人資料
       If GetMaxNo(strNewCU01) <> "" Then
          GoTo JumpToMergePCU
       Else
       'end 2024/05/28
          'Added by Lydia 2024/05/10 聯絡人相片
          strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,pcc02 from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "2") & "' " & _
                         ",IBF03='" & Pub_GetPCCtoIBF(strNewCU01, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                         "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
                cnnConnection.Execute strSql
                RsTemp.MoveNext
             Loop
          End If
          'end 2024/05/10
          '同時更新聯絡人資料
          strSql = "UPDATE PotCustCont SET PCC01 = '" & strNewCU01 & "' " & _
                                    "WHERE PCC01 = '" & strOldCU01 & "' "
          Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
          cnnConnection.Execute strSql
       End If
    End If
    rsTmp.Close
    
   If strDelOldNum = "Y" Then
      strSql = "DELETE POTCUSTOMER WHERE PCU01 = '" & strOldCU01 & "' AND " & _
                     "PCU02 = '" & strOldCU02 & "' "
      Pub_SeekTbLog strSql   '新增維護記錄檔
      cnnConnection.Execute strSql
      If strOldCU02 = "0" Then
         'Added by Lydia 2024/05/10 刪除聯絡人相片
         strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & strOldCU01 & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
               strSql = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
               cnnConnection.Execute strSql
               RsTemp.MoveNext
            Loop
         End If
         'end 2024/05/10
         strSql = "DELETE POTCUSTCONT WHERE PCC01 = '" & strOldCU01 & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
 
   '潛在客戶-關係企業
   strSql = "UPDATE PotCustomer SET PCU47 = '" & strNewNum & "' " & _
                  "WHERE PCU47 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   'Added by Lydia 2022/03/28 DHL輸入資料
    strSql = "UPDATE  DHL_INPUT_DATA SET DID01 = '" & strNewCU01 & "', DID02 = '" & strNewCU02 & "' " & _
             "WHERE DID01 = '" & strOldCU01 & "' AND DID02 = '" & strOldCU02 & "' "
    cnnConnection.Execute strSql
    
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

' 變更潛在客戶基本檔的編號
'Modified by Lydia 2020/06/10 + ByVal strModMerge As String
Private Sub ModifyPotCustomer1(ByVal strOldNum As String, ByVal strNewNum As String, ByVal strDelOldNum As String, ByVal strModMerge As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   Dim strOldCU01 As String
   Dim strOldCU02 As String
   Dim strNewCU01 As String
   Dim strNewCU02 As String
   
   If Len(strOldNum) > 8 Then
      strOldCU01 = Mid(strOldNum, 1, 8)
      strOldCU02 = Mid(strOldNum, 9, 1)
   Else
      strOldCU01 = strOldNum & String(8 - Len(strOldNum), "0")
      strOldCU02 = "0"
   End If
   
   If Len(strNewNum) > 8 Then
      strNewCU01 = Mid(strNewNum, 1, 8)
      strNewCU02 = Mid(strNewNum, 9, 1)
   Else
      strNewCU01 = strNewNum & String(8 - Len(strNewNum), "0")
      strNewCU02 = "0"
   End If
   
   ShowStatus "更新國內潛在客戶基本檔 原編號:<" & strOldNum & ">為新編號:<" & strNewNum & ">"
   
    strSql = "SELECT * FROM POTCUSTOMER1 " & _
             "WHERE POC01 = '" & strNewCU01 & "' AND " & _
                   "POC02 = '" & strNewCU02 & "' "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
         '接洽人資料合併
         'Modified by Lydia 2020/06/10
         'If Me.TextMerge = "Y" Then
         If strModMerge = "Y" Then
             strTemp = GetMaxNo(strNewCU01)
             strSql = "UPDATE POTCUSTCONT SET PCC01 = '" & strNewCU01 & "' " & _
                                      ", PCC02=LPAD(PCC02+" & Val(strTemp) & ", 2 ,'0') " & _
                                      "WHERE PCC01 = '" & strOldCU01 & "' "
             Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
             cnnConnection.Execute strSql
             Call UpdCR04(strOldCU01, strNewCU01, strTemp) 'Added by Lydia 2019/05/30 更新往來記錄的連絡人
         End If
    Else
       strSql = "UPDATE POTCUSTOMER1 SET POC01 = '" & strNewCU01 & "', " & _
                                    "POC02 = '" & strNewCU02 & "' " & _
                "WHERE POC01 = '" & strOldCU01 & "' AND " & _
                      "POC02 = '" & strOldCU02 & "' "
       Pub_SeekTbLog strSql   '新增維護記錄檔
       cnnConnection.Execute strSql
        '同時更新聯絡人資料
        strSql = "UPDATE PotCustCont SET PCC01 = '" & strNewCU01 & "' " & _
                                  "WHERE PCC01 = '" & strOldCU01 & "' "
        Pub_SeekTbLog strSql 'Added by Lydia 2025/07/24 新增維護記錄檔
        cnnConnection.Execute strSql
    End If
    rsTmp.Close
   
   If strDelOldNum = "Y" Then
      strSql = "DELETE POTCUSTOMER1 WHERE POC01 = '" & strOldCU01 & "' AND " & _
                     "POC02 = '" & strOldCU02 & "' "
      Pub_SeekTbLog strSql   '新增維護記錄檔
      cnnConnection.Execute strSql
      If strOldCU02 = "0" Then
         strSql = "DELETE POTCUSTCONT WHERE PCC01 = '" & strOldCU01 & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
 
   '國內潛在客戶-關係企業
   strSql = "UPDATE PotCustomer1 SET POC16 = '" & strNewNum & "' " & _
                  "WHERE POC16 = '" & strOldNum & "' "
   cnnConnection.Execute strSql
   
   ShowStatus Empty
   
   Set rsTmp = Nothing
End Sub

Private Sub TextMerge_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextMerge_GotFocus()
   TextInverse Me.TextMerge
End Sub

'Added by Lydia 2019/05/30 更新往來記錄的聯絡人
Private Sub UpdCR04(ByVal pOldNo As String, ByVal pNewNo As String, ByVal pAddVal As String)
Dim tmpArr As Variant
Dim mCR04 As String
Dim intR As Integer
Dim rsRd As New ADODB.Recordset
Dim strMid As String

     '有合併接洽人,並且有更新聯絡人編號
     If Trim(Me.TextMerge.Text) <> "Y" Or Val(pAddVal) = 0 Then Exit Sub
   
     strMid = "select CR01,CR02,CR03,CR04 from ContactRecord where CR03='" & ChangeCustomerL(pOldNo) & "' and cr04 is not null order by 1 "
     intR = 1
     Set rsRd = ClsLawReadRstMsg(intR, strMid)
     If intR = 1 Then
         rsRd.MoveFirst
         Do While Not rsRd.EOF
             tmpArr = Empty
             tmpArr = Split("" & rsRd.Fields("CR04"), ",")
             mCR04 = ""
             For intI = 0 To UBound(tmpArr)
                If Trim(tmpArr(intI)) <> "" Then
                   mCR04 = mCR04 & "," & Format(Val(tmpArr(intI)) + Val(pAddVal), "00")
                End If
             Next intI
             
             strSql = "Update ContactRecord set CR04='" & Mid(mCR04, 2) & "' Where CR01='" & rsRd.Fields("CR01") & "' "
             cnnConnection.Execute strSql
             rsRd.MoveNext
         Loop
     End If
     Set rsRd = Nothing
End Sub

'Added by Lydia 2020/05/07 檢查各項指示檔在更新代號後，是否會造成重複
Public Function CheckInstructionsDual(ByVal pOldNo As String, ByVal pNewNo As String) As Boolean
Dim intJ As Integer
Dim rsR1 As New ADODB.Recordset
Dim strCon1 As String
     
     strCon1 = "SELECT ITS01,ITS02,ITS03,ITS04 FROM INSTRUCTIONS WHERE ITS02='" & pNewNo & "' " & _
                    "AND (ITS01,ITS02,ITS03,ITS04) IN (SELECT ITS01,'" & pNewNo & "',ITS03,ITS04 FROM INSTRUCTIONS WHERE ITS02='" & pOldNo & "') "
     intJ = 1
     Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
     If intJ = 1 Then
          MsgBox "各項指示在更新代號會有重複記錄，請先整合記錄！", vbCritical, "檢核資料"
          CheckInstructionsDual = True
     End If
     '檢查國外部關聯企業資料
     If Left(pOldNo, 1) = "X" Or Left(pOldNo, 1) = "Y" Then
        strCon1 = "SELECT FR01,FR02 FROM FRELATION WHERE FR01='" & pNewNo & "' AND (FR01,FR02) IN (SELECT '" & pNewNo & "',FR02 FROM FRELATION WHERE FR01='" & pOldNo & "') " & _
                       "UNION SELECT FR01,FR02 FROM FRELATION WHERE FR02='" & pNewNo & "' AND (FR01,FR02) IN (SELECT FR01,'" & pNewNo & "' FROM FRELATION WHERE FR02='" & pOldNo & "') "
        intJ = 1
        Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
        If intJ = 1 Then
             MsgBox "國外部關聯企業資料在更新代號會有重複記錄，請先整合記錄！", vbCritical, "檢核資料"
             CheckInstructionsDual = True
        End If
     End If
     
     Set rsR1 = Nothing
End Function

'Add by Amy 2022/12/07 是否存在XYS02介紹來源編號
Private Function ChkHasXYSData() As Boolean
    'Add by Amy 2024/11/29
    'Modify by Amy 2025/07/03  strOldNumXYSData(2) As String, strNewNumXYSData(2) As String, intType As Integer,bolData As Boolean
    Dim strTp As String
    
    strOldNumXYSData(0) = "": strOldNumXYSData(1) = "": strOldNumXYSData(2) = ""
    strNewNumXYSData(0) = "": strNewNumXYSData(1) = "": strNewNumXYSData(2) = ""
    intType = Empty: bolData = False
    'end 2025/07/03
    ChkHasXYSData = False
    'Mark by Amy 2024/11/29 潛在客戶也可轉,也加來所原因
    'If Left(textOldNum, 1) <> 代理人編號 And Left(textOldNum, 1) <> 客戶編號 Then Exit Function
    If Len(textOldNum) = 9 And Right(textOldNum, 1) = "1" Then Exit Function 'Add by Amy 2022/12/14
    
    'Modify by Amy 2022/12/16 輸9碼會錯 原:String(8 - Len(textOldNum), "0")
    'Modify by Amy 2024/11/29 +顯示編號(可能多筆)及改訊息至共用/Y編號併號來所原因檢查;增加併號檢查
    ChkHasXYSData = Pub_GetXYSource(2, Left(textOldNum & String(9 - Len(textOldNum), "0"), 8), , , , Me.Name, strTp)
    If ChkHasXYSData = True Then
        'Modify by Amy 2022/12/14 +訊息
        MsgBox strTp, vbOKOnly, "注意"
        textOldNum_GotFocus
        Exit Function
    End If
    
'*** 原編號 textOldNum [併]入新編號 textNewNum編號(保留textNewNum編號資料)時,檢查來所原因 ***
    strTp = ""
    '舊/新編號欄都有輸 再檢查
    'Modify by Amy 2025/01/15 會有改號後要還原,故改用變數判斷(原直接使用textOldNum/textNewNum)-秀玲
    If Trim(textOldNum) <> MsgText(601) And Trim(textNewNum) <> MsgText(601) Then
      strExc(2) = Left(GetNewFagent(textOldNum), 8)
      strExc(3) = Left(GetNewFagent(textNewNum), 8)
      '[原]編號 textOldNum 的介紹者編號 (XYS02)/其他說明 (XYS03)
      Call Pub_GetXYSource(3, strExc(2), strOldNumXYSData(1), , strOldNumXYSData(2), Me.Name, , , strOldNumXYSData(0))
      '抓[新] 編號 來所原因/XYS02/XYS03
      bolData = Pub_GetXYSource(3, strExc(3), strNewNumXYSData(1), , strNewNumXYSData(2), Me.Name, , intType, strNewNumXYSData(0))
      
      '轉入 代理人 Or 國外潛在客戶 [新]編號 已有資料(非轉入新號=併號)
      If bolData = True And (intType = "1" Or intType = "3") And strNewNumXYSData(0) <> "NoData" Then
         strTp = "　原　代理人 " & strExc(2) & " 來所原因 為[原因OLD]" & vbCrLf & _
                        "併入之代理人 " & strExc(3) & " 來所原因 為[原因NEW]" & vbCrLf & _
                        "請確認 來所原因 應如何修改！"
         If intType = 3 Then strTp = Replace(strTp, "代理人", "國外潛在客戶")
         
         '來所原因[相同]
         If strOldNumXYSData(0) = strNewNumXYSData(0) Then
            '介紹者 or 其他說明 資料不同
            If strOldNumXYSData(1) <> strNewNumXYSData(1) Or strOldNumXYSData(2) <> strNewNumXYSData(2) Then
               ChkHasXYSData = True
               If strOldNumXYSData(1) <> strNewNumXYSData(1) Then
                  strExc(8) = IIf(strOldNumXYSData(1) = "", "空", strOldNumXYSData(1))
                  strExc(9) = IIf(strNewNumXYSData(1) = "", "空", strNewNumXYSData(1))
                  strTp = Replace(Replace(strTp, "原因OLD", strExc(8)), "原因NEW", strExc(9))
                  strTp = "[介紹者編號]不同如下:" & vbCrLf & Replace(strTp, "來所原因", "介紹者編號")
               Else
                  strExc(8) = IIf(strOldNumXYSData(2) = "", "空", strOldNumXYSData(2))
                  strExc(9) = IIf(strNewNumXYSData(2) = "", "空", strNewNumXYSData(2))
                  strTp = Replace(Replace(strTp, "原因OLD", strExc(8)), "原因NEW", strExc(9))
                  strTp = "[其他說明]不同如下:" & vbCrLf & Replace(strTp, "來所原因", "其他說明")
               End If
               If ChkHasXYSData = True Then
                  strTp = "來所原因 相同, " & strTp
               End If
            End If
         '來所原因[不同]
         Else
            strExc(8) = IIf(strOldNumXYSData(0) = "", "空", strOldNumXYSData(0))
            strExc(9) = IIf(strNewNumXYSData(0) = "", "空", strNewNumXYSData(0))
            strTp = Replace(Replace(strTp, "原因OLD", strExc(8)), "原因NEW", strExc(9))
            '[原]編號 來所原因 有值,[併]入[新]編號 來所原因 為空,彈訊息確認
            If strOldNumXYSData(0) <> MsgText(601) And strNewNumXYSData(0) = MsgText(601) Then
               ChkHasXYSData = True
            '來所原因 [不同]且都有值,彈訊息確認
            ElseIf strOldNumXYSData(0) <> MsgText(601) And strNewNumXYSData(0) <> MsgText(601) Then
               ChkHasXYSData = True
            End If
            If ChkHasXYSData = True Then
               '其中一個 來所原因 為04/05/11
               If InStr("04;05;11", strNewNumXYSData(0)) > 0 Or InStr("04;05;11", strOldNumXYSData(0)) > 0 Then
                  strTp = strTp & vbCrLf & vbCrLf & _
                              "PS.[來所原因] 為 [04/05/11] " & vbCrLf & _
                              "　 會產生 XYNoSource 資料檔,請一併確認！"
               End If
            End If
         End If
      '國內潛在客戶 / 客戶檔 (無 來所原因 欄位,但XYNoSource可能有資料)
      ElseIf intType = 2 Or intType = 4 Then
         'Memo by Amy 國內潛在客戶[無] 來所原因-秀玲:Widen 使用外國潛在客戶         '
         '[原]編號的 介紹者編號 與[新]編號 相同
         If strOldNumXYSData(1) = strExc(3) Then
            ChkHasXYSData = True
            strTp = "併入[新]客戶編號的【介紹者編號】為 " & strOldNumXYSData(1) & " 與[原]客戶編號相同" & vbCrLf & _
                           "(若直接新更新資料,會變成自己介給自己)" & vbCrLf & _
                           "請確認應如何修改！"
            
         End If
      End If
      If ChkHasXYSData = True Then
         MsgBox strTp, vbOKOnly, "注意"
         textOldNum_GotFocus
         Exit Function
      End If
    End If
    'end 2025/01/15
'*** End textOldNum代理人[併]入舊代理textNewNum編號(保留textNewNum編號資料)時,檢查來所原因 ***
    'end 2024/11/29
End Function

'Added by Morgan 2023/5/3
Public Function CheckCustEngMap(ByVal pNewNo As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * From CustEngMap where CEM01='" & pNewNo & "'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      CheckCustEngMap = True
      MsgBox "客戶新編號已有〔客戶承辦工程師對照檔 CustEngMap〕資料,請修改資料後,再執行改號作業！", vbOKOnly, "檢核資料"
      textOldNum_GotFocus
   End If
   Set rsQuery = Nothing
End Function

'Added by Lydia 2023/06/18 其他Table檢查: 在更新代號後，是否會造成重複
Public Function CheckOtherDual(ByVal pOldNo As String, ByVal pNewNo As String) As Boolean
Dim intJ As Integer
Dim rsR1 As New ADODB.Recordset
Dim strCon1 As String, strMsg As String
     
   CheckOtherDual = False
   '名條清單AddressA4List:兼具特殊控制
   strCon1 = "select aal01,aal04, count(*) cnt from (" & _
             "select aal01,aal04,'1' as ord1 from addressa4list where aal04='" & pNewNo & "' " & _
             "union select aal01,'" & pNewNo & "' aal04,'2' as ord1 from addressa4list where aal04='" & pOldNo & "' " & _
             ") group by aal01,aal04 having count(*) > 1 "
   intJ = 1
   Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
   If intJ = 1 Then
      rsR1.MoveFirst
      Do While Not rsR1.EOF
         strMsg = strMsg & vbCrLf & "AddressA4List特殊設定(" & rsR1.Fields("AAL01") & ") 更新後會有重複AAL04編號;"
         rsR1.MoveNext
      Loop
   End If
   '客戶平台CustWeb: CW04為串接的客戶編號
   strCon1 = "select CW01,CW04 from custweb where cw04 is not null and instr(upper(cw04),upper('" & pOldNo & "')) > 0 and instr(upper(cw04),upper('" & pNewNo & "')) > 0 "
   intJ = 1
   Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
   If intJ = 1 Then
      rsR1.MoveFirst
      Do While Not rsR1.EOF
         strMsg = strMsg & vbCrLf & "客戶平台CustWeb ：" & vbCrLf & _
               "　平台編號: " & rsR1.Fields("CW01") & "，客戶編號=>" & rsR1.Fields("CW04") & ";"
         rsR1.MoveNext
      Loop
   End If
   '客戶平台帳號CustWebID:CD02依客戶編號分別建立帳號資料
   strCon1 = "select cd01,cd02, count(*) cnt from (" & _
            "select cd01,cd02,'1' as ord1 from custwebid where cd02='" & pNewNo & "' " & _
            "union select cd01,'" & pNewNo & "' cd02,'2' as ord1 from custwebid where cd02='" & pOldNo & "' " & _
            ") group by cd01,cd02 having count(*) > 1 "
   intJ = 1
   Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
   If intJ = 1 Then
      rsR1.MoveFirst
      Do While Not rsR1.EOF
         If InStr(strMsg, "客戶平台帳號CustWebID") = 0 Then
            strMsg = strMsg & vbCrLf & "客戶平台帳號CustWebID：" & vbCrLf & _
                  "　平台編號: " & rsR1.Fields("CD01") & "，客戶編號=>" & rsR1.Fields("CD02") & ";"
         Else
            strMsg = strMsg & vbCrLf & "　平台編號: " & rsR1.Fields("CD01") & "，客戶編號=>" & rsR1.Fields("CD02") & ";"
         End If
         rsR1.MoveNext
      Loop
   End If
   
   'Add by Amy 2025/07/04 合併資料會錯 ex:Y00043 併為Y56151 更新會錯
   If textUpdate1 = "Y" Then
      '客戶/代理人匯款銀行資料
       strCon1 = "Select A2202,stNo, COUNT(*) CNT From (" & _
                     "Select A2202,A2201 as STNO,'1' as ord1 From ACC220 Where a2201='" & pNewNo & "' " & _
        "Union Select A2202,'" & pNewNo & "' as stNo,'2' as ord1 From ACC220 Where a2201='" & pOldNo & "' " & _
               ") Group by a2202,stNo Having count(*) > 1 "
      intJ = 1
      Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
      If intJ = 1 Then
         rsR1.MoveFirst
         Do While Not rsR1.EOF
            strMsg = strMsg & vbCrLf & "客戶/代理人匯款銀行資料ACC220：" & vbCrLf & _
                                    "　客戶/代理人編號: " & pOldNo & "，編號=>" & rsR1.Fields("stNo") & ";"
            rsR1.MoveNext
         Loop
      End If
      '客戶減免身分資料
      strCon1 = "Select AD02,stNo, COUNT(*) CNT From (" & _
                     "Select AD02,AD01 as STNO,'1' as ord1 From ApplicantDiscount Where ad01='" & pNewNo & "' " & _
        "Union Select AD02,'" & pNewNo & "' as stNo,'2' as ord1 From ApplicantDiscount Where ad01='" & pOldNo & "' " & _
               ") Group by AD02,stNo Having count(*) > 1 "
      intJ = 1
      Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
      If intJ = 1 Then
         rsR1.MoveFirst
         Do While Not rsR1.EOF
            strMsg = strMsg & vbCrLf & "客戶減免身分資料ApplicantDiscount：" & vbCrLf & _
                  "　客戶編號: " & pOldNo & "，編號=>" & rsR1.Fields("stNo") & ";"
            rsR1.MoveNext
         Loop
      End If
      '重新委任客戶資料
      strCon1 = "Select stNo, COUNT(*) CNT From (" & _
                     "Select LR01 as STNO,'1' as ord1 From LINREASIGNREC Where lr01='" & pNewNo & "' " & _
        "Union Select '" & pNewNo & "' as stNo,'2' as ord1 From LINREASIGNREC Where lr01='" & pOldNo & "' " & _
               ") Group by stNo Having count(*) > 1 "
      intJ = 1
      Set rsR1 = ClsLawReadRstMsg(intJ, strCon1)
      If intJ = 1 Then
         rsR1.MoveFirst
         Do While Not rsR1.EOF
            strMsg = strMsg & vbCrLf & "重新委任客戶資料LINREASIGNREC：" & vbCrLf & _
                                    "　客戶編號: " & pOldNo & "，編號=>" & rsR1.Fields("stNo") & ";"
            rsR1.MoveNext
         Loop
      End If
   End If
            
   If strMsg <> "" Then
      MsgBox "以下設定在更新代號會有重複記錄，請先整合記錄！" & strMsg, vbCritical, "檢核資料"
      CheckOtherDual = True
   End If
     
   Set rsR1 = Nothing
End Function
