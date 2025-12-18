VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm082005 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓客戶查詢"
   ClientHeight    =   2460
   ClientLeft      =   1830
   ClientTop       =   1200
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   760
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2592
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   2844
      TabIndex        =   4
      Top             =   1692
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1404
      TabIndex        =   3
      Top             =   1695
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1404
      TabIndex        =   2
      Top             =   1212
      Width           =   2535
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "4471;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   2844
      TabIndex        =   1
      Top             =   732
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1404
      TabIndex        =   0
      Top             =   732
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      X1              =   2604
      X2              =   2724
      Y1              =   1872
      Y2              =   1872
   End
   Begin VB.Line Line1 
      X1              =   2604
      X2              =   2724
      Y1              =   912
      Y2              =   912
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收件人名稱："
      Height          =   180
      Index           =   4
      Left            =   324
      TabIndex        =   9
      Top             =   1272
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國        籍："
      Height          =   180
      Index           =   3
      Left            =   324
      TabIndex        =   8
      Top             =   1752
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "屬性代號："
      Height          =   180
      Index           =   0
      Left            =   324
      TabIndex        =   7
      Top             =   792
      Width           =   900
   End
End
Attribute VB_Name = "frm082005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; Text1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdSure_Click()
 Dim Rules As String
   If Text1(0).Text = "" And Text1(1).Text = "" Then
      MsgBox "屬性代號不可同時為空值 !", vbCritical
      Exit Sub
   End If
   If Text1(0).Text <> "" And Text1(1).Text <> "" Then
      Rules = " ECD02 BETWEEN '" & Text1(0).Text & "' AND '" & Text1(1).Text & "' AND "
   ElseIf Text1(0).Text <> "" And Text1(1).Text = "" Then
      Rules = " ECD02 >= '" & Text1(0).Text & "' AND "
   ElseIf Text1(0).Text = "" And Text1(1).Text <> "" Then
      Rules = " ECD02 <= '" & Text1(1).Text & "' AND "
   End If
   If Text1(2).Text <> "" Then
      Rules = Rules & " ECD03='" & Text1(2).Text & "' AND "
   End If
   If Text1(3).Text <> "" And Text1(4).Text <> "" Then
      Rules = Rules & " ECD09 BETWEEN '" & Text1(3).Text & "' AND '" & Text1(4).Text & "'"
   ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
      Rules = Rules & " ECD09 >= '" & Text1(3).Text & "'"
   ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
      Rules = Rules & " ECD09 <= '" & Text1(4).Text & "'"
   End If
   If Rules <> "" Then Rules = " AND " & Trim(Rules)
   If Right(Rules, 3) = "AND" Then Rules = Left(Rules, Len(Rules) - 3)
   strExc(0) = "SELECT DECODE(ECD02,ECA01,ECA02),ECD01,ECD03,DECODE(ECD09,NA01,NA03),ECD04 " & _
      "FROM EXPANDCUSDETAIL,EXPANDCUSATTR,NATION " & _
      "WHERE ECD02=ECA01(+) AND ECD09=NA01(+)" & Rules & " ORDER BY ECD01,ECD02"
   intI = 0
   'edit by nickc 2007/02/27 不用 dll
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
   If intI <> 1 Then Exit Sub
   Screen.MousePointer = vbHourglass
   frm082006.Show
   Me.Hide
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm082005 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 0 Or Index = 1 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
         If Text1(0).Text <> "" And Text1(1).Text <> "" Then
            If Text1(0).Text > Text1(1).Text Then
               MsgBox "前一屬性代號大於後一屬性代號 !", vbCritical
               Cancel = True
            End If
         End If
      Case 3, 4
      'edit by nickc 2007/02/27 不用 dll
         'If objPublicData.GetNation(Text1(Index), strTempName) = False Then
         If ClsPDGetNation(Text1(Index), strTempName) = False Then
            Cancel = True
         End If
   End Select
End Sub
