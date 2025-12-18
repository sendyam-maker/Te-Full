VERSION 5.00
Begin VB.Form frm880008 
   BorderStyle     =   1  '單線固定
   Caption         =   "領證及年費"
   ClientHeight    =   5430
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5415
   Begin VB.CommandButton Command1 
      Caption         =   "延伸國"
      Height          =   375
      Index           =   1
      Left            =   1305
      TabIndex        =   5
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "成員國"
      Height          =   375
      Index           =   0
      Left            =   405
      TabIndex        =   4
      Top             =   90
      Width           =   855
   End
   Begin VB.ListBox lstCountry 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  '項目包含核取方塊
      TabIndex        =   2
      Top             =   780
      Width           =   5172
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3216
      TabIndex        =   0
      Top             =   50
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   4152
      TabIndex        =   1
      Top             =   50
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "國家名稱(英)　　　　　國家名稱(中)　　　　　國家代號"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   3
      Top             =   540
      Width           =   4890
   End
End
Attribute VB_Name = "frm880008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'StrCountry存放指定國家  strLicenceCountry存放領證國家
Public strCountry As String, strLicenceCountry As String
'add by nickc 2006/06/15 若是商標時控制
Public IsByTM As Boolean
Public strPA10 As String
Public strCP10 As String 'Added by Morgan 2023/3/7

Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer
   Dim bolUp As Boolean, bolMember As Boolean
   
   If Index = 0 Then
      'Added by Morgan 2023/3/7
      If Not IsByTM Then
         'EPC領證/指定國註冊費: 不可同時勾選ＵＰ及其會員國
         If strCP10 = "601" Or strCP10 = "224" Then
            For i = 0 To lstCountry.ListCount - 1
               If lstCountry.Selected(i) Then
                  If Right(lstCountry.List(i), 3) = "224" Then
                     bolUp = True
                  ElseIf InStr(UPMember, Right(lstCountry.List(i), 3)) > 0 Then
                     bolMember = True
                  End If
               End If
            Next
            If bolUp And bolMember Then
               MsgBox "不可同時勾選ＵＰ及其會員國！", vbExclamation
               Exit Sub
            End If
         End If
      End If
      'end 2023/3/7
      
      strLicenceCountry = ""
      For i = 0 To lstCountry.ListCount - 1
             If lstCountry.Selected(i) Then
                'Modify by Morgan 2005/7/4 '國家代號移到最右邊
                'strLicenceCountry = strLicenceCountry + Left(lstCountry.List(i), 3) + ","
                strLicenceCountry = strLicenceCountry + Right(lstCountry.List(i), 3) + ","
             'Remove by Morgan 2008/1/21 沒選不做，否則判斷是否有選會錯
             'Else
             '   strLicenceCountry = strLicenceCountry + ","
             End If
      Next
      If Right(strLicenceCountry, 1) = "," Then
         strLicenceCountry = Mid(strLicenceCountry, 1, Len(strLicenceCountry) - 1)
      End If
   End If
   'add by nickc 2006/06/15
   If IsByTM = False Then
       Unload Me
   Else
       If Index = 1 Then IsByTM = False
       Me.Hide
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim varTmp As Variant, j As Integer
   
   If lstCountry.ListCount = 0 Then Exit Sub
   '會員國
   If Index = 0 Then
      strExc(0) = PUB_GetMemberCountry("1", strPA10)
   '延伸國
   Else
      strExc(0) = PUB_GetMemberCountry("2", strPA10)
   End If
   For j = 0 To lstCountry.ListCount - 1
      strExc(1) = Right(lstCountry.List(j), 3)
      If InStr(strExc(0), strExc(1)) > 0 Then
         lstCountry.Selected(j) = True
      End If
   Next
   lstCountry.ListIndex = 0
End Sub

'分析字串並存入ListBox
Private Sub Form_Load()
Dim i As Integer, varCountryTemp As Variant, strTemp As String

   'Modify by Morgan 2005/7/4
   '依照國家英文排序
   'Dim objPublicData As Object
   'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
   'strCountry = PUB_GetCountryData(strCountry)
   'varCountryTemp = Split(strCountry, ",")
   'For i = 0 To UBound(varCountryTemp)
   '       If objPublicData.GetNation(CStr(varCountryTemp(i)), strTemp) Then
   '          lstCountry.AddItem varCountryTemp(i) + vbTab + "     " + strTemp
   '       End If
   'Next
   strTemp = PUB_GetCountryData(strCountry)
   varCountryTemp = Split(strTemp, ",")
   For i = 0 To UBound(varCountryTemp)      '
      If Not (Not IsByTM And strCP10 = "224" And Right(varCountryTemp(i), 3) = "224") Then 'Added by Morgan 2023/3/9 CFP指定國註冊費不列出224ＵＰ
         lstCountry.AddItem varCountryTemp(i)
      End If
   Next
   '2005/7/4 end
   
   'edit by nickc 2006/10/16 set nothing  會造成選指定國家後失敗
   'Set objPublicData = Nothing
   
   If strLicenceCountry <> "" Then
      varCountryTemp = Split(strLicenceCountry, ",")
      For i = 0 To lstCountry.ListCount - 1
         'Modify by Morgan 2005/7/4 '國家代號移到最右邊
         'If Left(lstCountry.List(i), 3) = varCountryTemp(i) Then
         'Modified by Morgan 2017/12/27 修正第1次沒有全選時點第2次會當問題
         'If Right(lstCountry.List(i), 3) = varCountryTemp(i) Then
         If InStr(strLicenceCountry, Right(lstCountry.List(i), 3)) > 0 Then
            lstCountry.Selected(i) = True
         End If
      Next
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   'Set frm880008 = Nothing
End Sub

