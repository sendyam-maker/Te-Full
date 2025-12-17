VERSION 5.00
Begin VB.Form Frmacc14g0 
   AutoRedraw      =   -1  'True
   Caption         =   "名條列印"
   ClientHeight    =   4008
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5964
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4008
   ScaleWidth      =   5964
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2040
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   3060
      Width           =   3450
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   0
      Top             =   900
      Width           =   5220
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3420
      Width           =   5295
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   6
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "※只可列印曾付過款的客戶或廠商!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   315
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "※列印前請先確認標籤地址條機器是否開啟!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   30
      Width           =   4995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "編　　號：(下多筆時, 請以逗號分開)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   630
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc14g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/03/14 Form2.0已修改 (改為標籤地址條套印)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Dim adoacc0e0 As New ADODB.Recordset
'Public adoaccrpt111 As New ADODB.Recordset
'Public adoquery As New ADODB.Recordset
Dim strSql As String
'Dim strAmount As String, intLength As Integer, intCounter As Integer,intPage As Integer,intRecord As Integer,dblLin As Double
'預設印表機
Dim m_DefaultPrinter As String
Dim ii As Integer, bolGODEXPrinter As Boolean, strPrinter As String 'Add by Amy 2022/03/14

Private Sub Command1_Click()
   Dim stAddrData As String 'Add by Amy 2022/03/14
   
   If FormCheck = False Then
      Exit Sub
   End If
'   strSQL = ""
'   If Text1 <> MsgText(601) Then
'      strSQL = strSQL & " and a0q03 >= '" & Text1 & "'"
'   End If
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2022/03/14 從原程式搬出來,改標籤地址條套印
   'Call PrintAddress
   If GetPrintData(stAddrData) = False Then
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
   End If
   PUB_SetOsDefaultPrinter Combo1
   PUB_RestorePrinter Combo1
   If PUB_XlsAccAddress(stAddrData) = False Then
        MsgBox "列印失敗！", vbCritical
   End If
   PUB_SetOsDefaultPrinter strPrinter
   PUB_RestorePrinter strPrinter
   
   Screen.MousePointer = vbDefault
   FormClear
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   'end 2022/03/14
End Sub

'Add by Amy 2022/03/14 從原程式搬過來修改
Private Function GetPrintData(ByRef stData As String) As Boolean
    Dim arrCustID() As String, strCustID As String, StrSQLa As String

    GetPrintData = False
    strSql = "": strCustID = ""
    If Text1 <> MsgText(601) Then
        arrCustID = Split(Me.Text1.Text, ",")
        For ii = LBound(arrCustID) To UBound(arrCustID)
            strCustID = strCustID & "'" & Trim(arrCustID(ii)) & "',"
        Next ii
        strCustID = Left(strCustID, Len(strCustID) - 1)
        strSql = strSql & " And A0Q03 In (" & strCustID & ")"
    End If
   
   'Modify by Morgan 2006/1/20 抓最大付款日的資料
'    StrSQLa = "Select A0I03 As A0Q16, A0Q05, A0Q04, A0Q03 From ACC0Q0, ACC0I0 Where A0Q03=A0I01 " & strSQL & " "
'    StrSQLa = StrSQLa & " Union Select CU30||CU31 As A0Q16, A0Q05, A0Q04, A0Q03 From ACC0Q0, Customer Where substr(A0Q03,1,8)=CU01 And substr(A0Q03,9,1)=CU02 " & strSQL & " Order By 3, 4 "
    'Modify by Morgan 2007/1/22 客戶地址若客戶狀態有資料時優先抓
    'Modify by Morgan 2009/1/21 廠商郵遞區號欄位自地址欄拆開a0i04
    StrSQLa = "Select NVL(a0i04||A0I03,NVL(CU80,CU30||CU31)) As A0Q16, A0Q05, A0Q04, A0Q03" & _
      " From (SELECT A0Q03,MAX(A0Q04) A0Q04,SUBSTR(MAX(A0Q01||A0Q05),7) A0Q05 FROM ACC0Q0" & _
      " WHERE 1=1 " & strSql & " GROUP BY A0Q03) X, ACC0I0, Customer" & _
      " Where A0I01(+)=A0Q03 and CU01(+)=substr(A0Q03,1,8) And CU02(+)=substr(A0Q03,9,1)" & _
      " Order By 3, 4"
      
    If adoacc0e0.State <> adStateClosed Then adoacc0e0.Close
    adoacc0e0.CursorLocation = adUseClient
    adoacc0e0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    With adoacc0e0
        If .RecordCount <> 0 Then
            Do While .EOF = False
                '                                                                       地  址                                      票據抬頭
                stData = stData & Trim(.Fields("A0Q16")) & "$" & Trim(.Fields("A0Q05")) & "|"
                .MoveNext
            Loop
            GetPrintData = True
        End If
    End With
End Function

'確認是否有印表機
Private Function ChkPrinter() As Boolean
    ChkPrinter = False
    PUB_SetPrinter Me.Name, Combo1, strPrinter
    If Combo1.ListCount > 0 Then
        For ii = 0 To Combo1.ListCount - 1
            If InStr(UCase(Combo1.List(ii)), "GODEX") > 0 Then
                ChkPrinter = True
                Exit For
            End If
        Next ii
    End If
End Function
'end 2022/03/14

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 6090, 4452 'Modify by Amy 2023/08/16 原:4245
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   'Add by Amy 2022/03/14 判斷是否有標籤地址條印表機
   bolGODEXPrinter = ChkPrinter
   m_DefaultPrinter = strPrinter
   If bolGODEXPrinter = False Then
        MsgBox "請洽電腦中心安裝「標籤地址條印表機(Godex EZ530)」"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/03/14若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc14g0 = Nothing
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
    Me.Text1.SetFocus
End Sub


'*************************************************
'  列印付款簽收簿 (抬頭及報表格式)
'  Memo by Amy 2022/03/01 原:單排印出來的資料為直的(因以前信封為直的),且「名條清單」已不使用-瑞婷
'*************************************************
Private Sub PrintHead()
'   Printer.FontSize = 20
'    Printer.Font.Underline = True
'   Printer.CurrentX = 4200
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "名　條　清　單"
'    Printer.Font.Underline = False
'   Printer.FontSize = 12
'   intCounter = intCounter + 2
'   Printer.CurrentX = 0
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "列印人員 : " & StaffQuery(strUserNum)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "列印日期 : " & CFDate(ACDate(ServerDate))
'   intCounter = intCounter + 1
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "頁　　次: " & intPage
'   intCounter = intCounter + 1
'   Printer.CurrentX = 0
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "編號"
'   Printer.CurrentX = 1500
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "抬　　頭"
'   intCounter = intCounter + 1
'   Printer.CurrentX = 1500
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "地　　址"
'   Printer.Line (0, 300 + intCounter * 300 + 350)-(13500 - 1000, 300 + intCounter * 300 + 350)
'   intCounter = intCounter + 2
End Sub

'*************************************************
'  列印地址條
'  Memo by Amy 2022/03/01 原:單排印出來的資料為直的(因以前信封為直的),且「名條清單」已不使用-瑞婷
'*************************************************
Private Sub PrintAddress()
'   Dim intCounter As Integer
'   Dim strName As String
'   Dim StrSQLa As String
'   Dim arrCustID() As String
'   Dim strCustID As String
'   Dim ii As Integer
'
'   strSql = ""
'   If Text1 <> MsgText(601) Then
'        arrCustID = Split(Me.Text1.Text, ",")
'        strCustID = ""
'        For ii = LBound(arrCustID) To UBound(arrCustID)
'            strCustID = strCustID & "'" & Trim(arrCustID(ii)) & "',"
'        Next ii
'        strCustID = Left(strCustID, Len(strCustID) - 1)
'        strSql = strSql & " And A0Q03 In (" & strCustID & ")"
'   End If
'   intCounter = 0
'   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
'   '9x
'   If pub_OS = "1" Then
'      Printer.Height = 2880
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(2)
'   End If
'   'end 2008/3/25
'
'   Printer.Font = "@新細明體"
'   Printer.FontSize = 12
'   adoacc0e0.CursorLocation = adUseClient
'   'Modify by Morgan 2006/1/20 抓最大付款日的資料
''    StrSQLa = "Select A0I03 As A0Q16, A0Q05, A0Q04, A0Q03 From ACC0Q0, ACC0I0 Where A0Q03=A0I01 " & strSQL & " "
''    StrSQLa = StrSQLa & " Union Select CU30||CU31 As A0Q16, A0Q05, A0Q04, A0Q03 From ACC0Q0, Customer Where substr(A0Q03,1,8)=CU01 And substr(A0Q03,9,1)=CU02 " & strSQL & " Order By 3, 4 "
'
'    'Modify by Morgan 2007/1/22 客戶地址若客戶狀態有資料時優先抓
'    'Modify by Morgan 2009/1/21 廠商郵遞區號欄位自地址欄拆開a0i04
'    StrSQLa = "Select NVL(a0i04||A0I03,NVL(CU80,CU30||CU31)) As A0Q16, A0Q05, A0Q04, A0Q03" & _
'      " From (SELECT A0Q03,MAX(A0Q04) A0Q04,SUBSTR(MAX(A0Q01||A0Q05),7) A0Q05 FROM ACC0Q0" & _
'      " WHERE 1=1 " & strSql & " GROUP BY A0Q03) X, ACC0I0, Customer" & _
'      " Where A0I01(+)=A0Q03 and CU01(+)=substr(A0Q03,1,8) And CU02(+)=substr(A0Q03,9,1)" & _
'      " Order By 3, 4"
'
'    adoacc0e0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0e0.RecordCount <> 0 Then
'      Screen.MousePointer = vbHourglass
'   Else
'      adoacc0e0.Close
'      ShowNoData
'      Printer.Font = "新細明體"
'      Exit Sub
'   End If
   
'   Do While adoacc0e0.EOF = False
'        Printer.CurrentX = 100
'        Printer.CurrentY = 300 + 2200 * intCounter
'        If IsNull(adoacc0e0.Fields("a0q16").Value) Then
'           Printer.Print ""
'        Else
'            'Modify by Morgan 2006/1/20 控制折行
'            'Printer.Print adoacc0e0.Fields("a0q16").Value
'            PUB_PrintAddress adoacc0e0.Fields("a0q16").Value, intCounter, 0
'        End If
'        Printer.CurrentX = 100
'        Printer.CurrentY = 1000 + 2200 * intCounter
'        If IsNull(adoacc0e0.Fields("a0q05").Value) Then
'           Printer.Print ""
'        Else
'           Printer.Print "　　　" & adoacc0e0.Fields("a0q05").Value & MsgText(104)
'        End If
'
'        'Modify by Morgan 2006/1/18 改用大張地址條
''        intCounter = intCounter + 1
''        If intCounter = 3 Then
''           intCounter = 0
''           Printer.NewPage
''        End If
'         Printer.NewPage
'
'      adoacc0e0.MoveNext
'   Loop
'
'   Printer.Font = "新細明體"
'   Printer.EndDoc
'   '列印名條清單
'   MsgBox "準備列印名條清單, 請更換紙張!!!", vbExclamation + vbOKOnly
'   PrintAddressList
   
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   'Add by Amy 2022/03/14 Form_Load判斷沒有,列印時再判斷是否有標籤地址條印表機
   If bolGODEXPrinter = False Then
        Combo1.Clear
        bolGODEXPrinter = ChkPrinter
        If bolGODEXPrinter = False Then
             MsgBox "請洽電腦中心安裝「標籤地址條印表機(Godex EZ530)」"
             Exit Function
        End If
   ElseIf InStr(UCase(Combo1), "GODEX") = 0 Then
        MsgBox "請選擇「標籤地址條印表機(Godex EZ530)」"
        Exit Function
   End If
   'end 2022/03/14
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   Else
      MsgBox MsgText(181), , MsgText(5)
   End If
End Function

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2022/03/01 原:單排印出來的資料為直的(因以前信封為直的),且「名條清單」已不使用-瑞婷
'Modify by Morgan 2006/1/20 用原查詢結果不必重抓
Private Sub PrintAddressList()
'   intCounter = 3
'   intRecord = 1
'   intPage = 0
'   dblLin = 0
'   intPage = intPage + 1
'   PrintHead
'   With adoacc0e0
'      .MoveFirst
'      While Not .EOF
'         Printer.CurrentX = 0
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print "" & .Fields("a0q03").Value
'         Printer.CurrentX = 1500
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print "" & .Fields("a0q05").Value
'         intCounter = intCounter + 1
'         Printer.CurrentX = 1500
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print "" & .Fields("a0q16").Value
'         intCounter = intCounter + 1
'         .MoveNext
'         If .EOF = False Then
'            If intCounter > 46 Then
'               Printer.NewPage
'               intCounter = 3
'               intRecord = 1
'               intPage = 0
'               dblLin = 0
'               intPage = intPage + 1
'               PrintHead
'            End If
'         End If
'      Wend
'      .Close
'   End With
'
'   Printer.EndDoc

End Sub
