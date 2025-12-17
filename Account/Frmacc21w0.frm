VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21w0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外固定寄催款單代理人維護"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8940.926
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   300
      Left            =   3645
      Picture         =   "Frmacc21w0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   350
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2070
      MaxLength       =   8
      TabIndex        =   0
      Top             =   270
      Width           =   1572
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   9
      Left            =   1800
      TabIndex        =   5
      Top             =   1748
      Width           =   1065
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1879;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   8
      Left            =   1320
      TabIndex        =   10
      Top             =   3653
      Width           =   5939
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "10476;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   2183
      Width           =   345
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "609;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Top             =   3053
      Width           =   5939
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "10476;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   5
      Left            =   1320
      TabIndex        =   8
      Top             =   2618
      Width           =   5939
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "10476;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      Top             =   2183
      Width           =   345
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "609;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   3
      Left            =   6000
      TabIndex        =   4
      Top             =   1313
      Width           =   585
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1032;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   345
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1313
      Width           =   345
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "609;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   735
      Width           =   6375
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "11245;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.每月 2.每季)"
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
      Index           =   12
      Left            =   1800
      TabIndex        =   29
      Top             =   1380
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(多個以;區隔)"
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
      Index           =   11
      Left            =   7320
      TabIndex        =   28
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(多個以;區隔)"
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
      Index           =   10
      Left            =   7320
      TabIndex        =   27
      Top             =   2685
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(Y:合併寄 空白:分別寄)"
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
      Index           =   9
      Left            =   6450
      TabIndex        =   26
      Top             =   2250
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(Y:附請款單)"
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
      Index           =   8
      Left            =   2237
      TabIndex        =   25
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "指定主旨："
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
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   3720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "副本信箱："
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
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "指定信箱："
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
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   2685
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "關係企業是否合併寄："
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
      Index           =   4
      Left            =   3720
      TabIndex        =   21
      Top             =   2250
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否附請款單："
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
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2250
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "下次寄發日期："
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
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   1815
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "每次固定日："
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
      Index           =   1
      Left            =   4620
      TabIndex        =   18
      Top             =   1380
      Width           =   1350
   End
   Begin VB.Label lblNa 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   17
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label lblNa 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   16
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "國籍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5295
      TabIndex        =   15
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS：若為過去的日期則表示以後不再固定寄催款單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3000
      TabIndex        =   14
      Top             =   1815
      Width           =   4965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "催款週期："
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
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1380
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "名　　稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "客戶/代理人編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   11
      Top             =   300
      Width           =   1755
   End
End
Attribute VB_Name = "Frmacc21w0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/02 改成Form2.0 ; Combo1、txtData(index)
'Create by Lydia 2016/11/04 國外固定寄催款單代理人維護
Option Explicit

Dim adoacc225 As New ADODB.Recordset
Dim oText As Object
Dim inX As Integer
Dim bolExist As Boolean

Public Sub FormClear()
   
   For Each oText In txtData
      oText = ""
   Next
   
   Combo1.Clear
   txtKey = ""
   lblNa(0) = "": lblNa(1) = ""
   
End Sub

Public Sub Command1_Click()
 
   If txtKey = "" Then
      MsgBox "請輸入欲查詢編號！"
   Else
        Acc225Refresh
        If adoacc225.RecordCount <> 0 Then
            FormShow
            RecordShow
        End If
   End If
End Sub

Private Sub Form_Activate()
    strFormName = Name
    Acc225Refresh
    If adoacc225.RecordCount <> 0 Then
        FormShow
        RecordShow
    End If
End Sub

'Added by Lydia 2021/12/02
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序
   
   If KeyCode = vbKeyF12 Then
      If Command1.Enabled = True Then
         Command1_Click
      End If
   Else
      KeyEnter KeyCode
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 8970, 4770, strBackPicPath1
   'tool6_enabled '要有新增功能
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       Cancel = 1
       Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
  
   strUserLevel = MsgText(601)
   Set Frmacc21w0 = Nothing
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

'Modified by Lydia 2021/12/02 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
       Case 4, 7
           KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   
   Cancel = False
   
   Select Case Index
       Case 2
            If txtData(Index) <> "1" And txtData(Index) <> "2" Then
               MsgBox "催款週期請輸入1或2！", vbExclamation, "檢核資料"
               Cancel = True
            End If
       Case 3
            If Val(txtData(Index)) < 1 Or Val(txtData(Index)) > 31 Then
               MsgBox "每次固定日請輸入1~31！", vbExclamation, "檢核資料"
               Cancel = True
            Else
               txtData(Index) = Format(txtData(Index), "00")
            End If
       Case 4
            If Trim(txtData(Index)) <> "" And Trim(txtData(Index)) <> "Y" Then
               MsgBox "是否附請款單請輸入空白或Y！", vbExclamation, "檢核資料"
               Cancel = True
            End If
       Case 7
            If Trim(txtData(Index)) <> "" And Trim(txtData(Index)) <> "Y" Then
               MsgBox "關係企業是否合併寄請輸入空白或Y！", vbExclamation, "檢核資料"
               Cancel = True
            End If
       Case 9
            If Trim(txtData(Index)) = "" Then
               MsgBox "下次寄發日期不可輸入空白！", vbExclamation, "檢核資料"
               Cancel = True
            Else
               If ChkDate(txtData(Index)) = False Then
                  Cancel = True
               Else
                  If txtData(Index) >= strSrvDate(2) Then
                     If Right(txtData(Index), 2) <> txtData(3) Then
                        MsgBox "請注意下次寄發日期是否符合固定日!", vbExclamation, "檢核資料"
                     End If
                  End If
               End If
            End If
   End Select
   
   If CheckLengthIsOK(txtData(Index), txtData(Index).MaxLength) = False Then
      Cancel = True
   End If
   
   If Cancel Then
      Txtdata_GotFocus Index
   Else
      CloseIme
   End If
   
End Sub

Private Function ReadData(ByVal p_Key As String) As Boolean
     
   FormClear
  
   If Left(p_Key, 1) = "X" Then
      strExc(0) = "select cu01||cu02 No,cu04 name1,cu05||cu88||cu89||cu90 name2,cu06 name3,cu10 na01,na03,a2251" & _
         " from customer,nation,acc225 where cu01 = '" & p_Key & "' and cu02='0' and cu10=na01(+) and cu01=a2251(+) "
   Else
      strExc(0) = "select fa01||fa02 No,fa04 name1,fa05||fa63||fa64||fa65 name2,fa06 name3,fa10 na01,na03,a2251" & _
         " from fagent,nation,acc225 where fa01 = '" & p_Key & "' and fa02='0' and fa10=na01(+) and fa01=a2251(+) "
   End If
   
   inX = 0
   Set RsTemp = ClsLawReadRstMsg(inX, strExc(0))
   If inX = 1 Then
      txtKey = p_Key
      lblNa(0).Caption = Mid("" & RsTemp.Fields("na01"), 1, 3)
      lblNa(1).Caption = Trim("" & RsTemp.Fields("na03"))
      If lblNa(0).Caption < "011" Or InStr("013,020,044", lblNa(0).Caption) > 0 Then
         Combo1.AddItem "中: " & RsTemp.Fields("name1")
         Combo1.AddItem "英: " & RsTemp.Fields("name2")
         Combo1.AddItem "日: " & RsTemp.Fields("name3")
      Else
         Combo1.AddItem "英: " & RsTemp.Fields("name2")
         Combo1.AddItem "中: " & RsTemp.Fields("name1")
         Combo1.AddItem "日: " & RsTemp.Fields("name3")
      End If
      Combo1.ListIndex = 0
      If "" & RsTemp.Fields("a2251") <> "" Then
         MsgBox "此筆資料已存在!", vbCritical
         bolExist = True
      End If
   Else
      Exit Function
   End If
   
   ReadData = True
   
End Function

Private Sub txtKey_GotFocus()
   TextInverse txtKey
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtKey_Validate(Cancel As Boolean)
    
    If strSaveConfirm <> MsgText(3) Then
        Exit Sub
    End If
    
    bolExist = False
    If txtKey = MsgText(601) Then
        MsgBox Label19 & MsgText(52), , MsgText(5)
        Cancel = True
    Else
       If Len(txtKey) < 9 Then txtKey = Mid(txtKey & String(8, "0"), 1, 8)
       If ReadData(txtKey) = False Then
          Cancel = True
       End If
    End If
    
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
    
On Error GoTo ErrorHandler
    If adoacc225.RecordCount = 0 Then
        Exit Sub
    End If
    CountShow adoacc225.AbsolutePosition, adoacc225.RecordCount
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    
End Sub

Private Sub Acc225Refresh()
On Error GoTo Checking
    If adoacc225.State = adStateOpen Then
        adoacc225.Close
    End If

    strSql = "Select * From ACC225 Order By 1 "
    adoacc225.CursorLocation = adUseClient
    adoacc225.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic

    With adoacc225
        If .RecordCount > 0 Then
            .Find "a2251 ='" & Left(Me.txtKey.Text & "00000000", 8) & "'", 0, adSearchForward, 1
            If .EOF = True Then
                If Trim(Me.txtKey.Text) <> "" Then MsgBox MsgText(28), , MsgText(5)
                .MoveFirst
            End If
        End If
    End With

Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub


'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()

   FormClear
   With adoacc225
      txtKey.Text = "" & .Fields("A2251")

      For Each oText In txtData
         inX = oText.Index - 1
         Select Case inX
             Case 2 '每次固定日
              oText.Text = Format("" & .Fields(inX), "00")
'             Case 8 '下次寄發日期
'              oText.Text = TransDate("" & .Fields(inX), 1)
             Case Else
              oText.Text = "" & .Fields(inX)
         End Select
         oText.Tag = oText.Text
      Next
      
       If Left(txtKey, 1) = "X" Then
          strExc(0) = "select cu01||cu02 No,cu04 name1,cu05||cu88||cu89||cu90 name2,cu06 name3,cu10 na01,na03" & _
             " from customer,nation where cu01 = '" & txtKey & "' and cu02='0' and cu10=na01(+) "
       Else
          strExc(0) = "select fa01||fa02 No,fa04 name1,fa05||fa63||fa64||fa65 name2,fa06 name3,fa10 na01,na03" & _
             " from fagent,nation where fa01 = '" & txtKey & "' and fa02='0' and fa10=na01(+) "
       End If
       
       inX = 0
       Set RsTemp = ClsLawReadRstMsg(inX, strExc(0))
       If inX = 1 Then
          lblNa(0).Caption = Mid("" & RsTemp.Fields("na01"), 1, 3)
          lblNa(1).Caption = Trim("" & RsTemp.Fields("na03"))
          If lblNa(0).Caption < "011" Or InStr("013,020,044", lblNa(0).Caption) > 0 Then
             Combo1.AddItem "中: " & RsTemp.Fields("name1")
             Combo1.AddItem "英: " & RsTemp.Fields("name2")
             Combo1.AddItem "日: " & RsTemp.Fields("name3")
          Else
             Combo1.AddItem "英: " & RsTemp.Fields("name2")
             Combo1.AddItem "中: " & RsTemp.Fields("name1")
             Combo1.AddItem "日: " & RsTemp.Fields("name3")
          End If
          Combo1.ListIndex = 0
       End If
   
      txtKey.Tag = txtKey
      
      Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
      RecordShow
   End With
End Sub

Public Sub MoveNext()
    If adoacc225.EOF = False Then
        adoacc225.MoveNext
        If adoacc225.EOF Then
            adoacc225.MoveLast
            MsgBox MsgText(8), , MsgText(5)
        End If
        FormShow
        RecordShow
    End If
End Sub

Public Sub MovePrevious()
    If adoacc225.BOF = False Then
        adoacc225.MovePrevious
        If adoacc225.BOF Then
            adoacc225.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
        End If
        FormShow
        RecordShow
    End If
End Sub

Public Sub MoveFirst()
    If adoacc225.RecordCount <> 0 Then
        adoacc225.MoveFirst
        FormShow
        RecordShow
    End If
End Sub

Public Sub MoveLast()
    If adoacc225.RecordCount <> 0 Then
        adoacc225.MoveLast
        FormShow
        RecordShow
    End If
End Sub

Public Function FormSave() As Boolean
Dim rsChk As New ADODB.Recordset
Dim bolTmp As Boolean
Dim strQ1 As String, intQ As Integer
Dim strUpd As String
    
    'Added by Lydia 2021/12/02 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        Exit Function
    End If
    'end 2021/12/02
    
   FormSave = False
   If lblNa(0).Caption = "" Or lblNa(1).Caption = "" Then
       MsgBox "客戶/代理人編號查無資料！", vbCritical
       Exit Function
   End If
   
   If bolExist = True Then
       MsgBox "此筆資料已存在!", vbCritical
       Exit Function
   End If
   
   For Each oText In txtData
       Txtdata_Validate oText.Index, bolTmp
       If bolTmp = True Then
          txtData(oText.Index).SetFocus
          Txtdata_GotFocus oText.Index
          Exit Function
       End If
   Next
   
   'Added by Lydia 2021/12/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   '若為過去的日期則存檔時提醒但仍可儲存，表示以後不再固定寄催款單；
   If txtData(9) < strSrvDate(2) Then
      MsgBox "下次寄發日期若為過去的日期，表示以後不再固定寄催款單！", vbCritical
   End If
      
   strQ1 = "select * from acc225 where a2251='" & txtKey & "' "
   intQ = 1
   Set rsChk = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
        '修改
       For Each oText In txtData
          If oText.Tag <> oText.Text Then
             strUpd = strUpd & " ,A225" & oText.Index & "=" & CNULL(PUB_StringFilter(Trim(oText.Text)))
          End If
       Next
       If strUpd <> "" Then
          strUpd = strUpd & " ,A2213='" & strUserNum & "', A2214=" & strSrvDate(2) & ", A2215=" & Format(ServerTime, "000000")
          strUpd = "UPDATE ACC225 SET " & Mid(strUpd, InStr(strUpd, ",") + 1) & " WHERE A2251='" & txtKey & "' "
       End If
   Else '新增
       strUpd = "INSERT INTO ACC225(A2251,A2252,A2253,A2254,A2255,A2256,A2257,A2258,A2259,A2210,A2211,A2212) "
       strUpd = strUpd & "VALUES ('" & txtKey & "','" & txtData(2) & "'," & Val(txtData(3)) & "," & CNULL(txtData(4))
       strUpd = strUpd & "," & CNULL(PUB_StringFilter(Trim(txtData(5)))) & "," & CNULL(PUB_StringFilter(Trim(txtData(6)))) & "," & CNULL(txtData(7))
       strUpd = strUpd & "," & CNULL(PUB_StringFilter(Trim(txtData(8)))) & "," & txtData(9)
       strUpd = strUpd & ",'" & strUserNum & "'," & strSrvDate(2) & "," & Format(ServerTime, "000000") & ")"
   End If
   
   If strUpd <> "" Then
       adoTaie.BeginTrans
          Pub_SeekTbLog strUpd
          adoTaie.Execute strUpd, intQ
       adoTaie.CommitTrans
   End If
   
   FormSave = True
   Acc225Refresh
   
   Exit Function
   
ErrorHand:
   If Err.Number <> 0 Then
      If strUpd <> "" Then adoTaie.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Public Sub FormDelete()
Dim strDel As String

On Error GoTo Checking
    
    strDel = "SELECT A2251 FROM ACC225 WHERE A2251 = '" & Me.txtKey.Text & "' "
    
    If DeleteCheck(strDel) = MsgText(603) Then
       Exit Sub
    End If
    
    strDel = "DELETE FROM ACC225 WHERE A2251 = '" & Me.txtKey.Text & "' "
    If strDel <> "" Then
        adoTaie.BeginTrans
          Pub_SeekTbLog strDel
          adoTaie.Execute strDel
        adoTaie.CommitTrans
    End If
   
    Me.txtKey.Text = ""
    
    Acc225Refresh
    If adoacc225.RecordCount <> 0 Then
       FormShow
       RecordShow
    Else
       FormClear
    End If
    
    Exit Sub
   
Checking:
   If Err.Number <> 0 Then
      If strDel <> "" Then adoTaie.RollbackTrans
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub
