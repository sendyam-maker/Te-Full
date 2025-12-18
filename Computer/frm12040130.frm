VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040130 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員調區作業"
   ClientHeight    =   2220
   ClientLeft      =   2040
   ClientTop       =   2616
   ClientWidth     =   5856
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5856
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3984
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4812
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textNewDep 
      Height          =   270
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox textSales 
      DataField       =   "ST01"
      Height          =   270
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
   Begin MSForms.TextBox textNewDep_2 
      Height          =   300
      Left            =   2400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2110
      VariousPropertyBits=   671105055
      Size            =   "3722;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSales_2 
      Height          =   300
      Left            =   2412
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2100
      VariousPropertyBits=   671105055
      Size            =   "3704;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "新業務區："
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frm12040130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0已修改(textSales_2,textNewDep_2)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Private Sub Form_Load()
   textSales_2.BackColor = &H8000000F
   textNewDep.BackColor = &H8000000F
   textNewDep_2.BackColor = &H8000000F
   EnableTextBox textNewDep, False
   MoveFormToCenter Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 執行更新的作業
      OnSaveData textSales, textNewDep
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      ' 清除欄位內容
      textSales = Empty
      textSales_2 = Empty
      textNewDep = Empty
      textNewDep_2 = Empty
      ' 設定輸入欄位
      textSales.SetFocus
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub OnSaveData(ByVal strSales As String, ByVal strDepart As String)
   Dim strSql As String
   Dim nCount As Integer
   
   ' 更新客戶基本資料檔
   strSql = "UPDATE CUSTOMER SET CU12 = '" & strDepart & "' " & _
            "WHERE CU13 = '" & strSales & "' "
   cnnConnection.Execute strSql, nCount
   
   If nCount <= 0 Then
      MsgBox "無此智權人員的客戶資料可更新", vbOKOnly + vbInformation, "智權人員調區作業"
   End If
   
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False

   ' 智權人員編號
   If IsEmptyText(textSales) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入智權人員編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSales.SetFocus
      GoTo EXITSUB
   End If
   
   ' 再檢查一次智權人員編號
   textSales_2 = GetStaffName(textSales, False)
   If IsEmptyText(textSales_2) = True Then
      strTit = "檢核資料"
      strMsg = "智權人員代號不存在或已離職!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSales.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040130 = Nothing
End Sub

Private Sub textSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員代號
Private Sub textSales_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSales_2 = Empty
   If IsEmptyText(textSales) = False Then
      textSales_2 = GetStaffName(textSales, False)
      If IsEmptyText(textSales_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "智權人員代號不存在或已離職!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSales_GotFocus
         GoTo EXITSUB
      End If
      '2010/3/23 MODIFY BY SONIA 改抓ST15(98008,98024)
      'textNewDep = GetStaffDepartment(textSales)
      textNewDep = GetST15(textSales)
      textNewDep_2 = GetDepartmentName(textNewDep)
   End If
EXITSUB:
End Sub

Private Sub textSales_GotFocus()
   InverseTextBox textSales
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textSales.Enabled = True Then
   Cancel = False
   textSales_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

