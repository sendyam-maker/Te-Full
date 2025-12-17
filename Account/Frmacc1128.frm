VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1128 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "發票備註維護作業"
   ClientHeight    =   4044
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8328
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   8328
   Begin VB.CommandButton cmdFind 
      Height          =   300
      Left            =   2970
      Picture         =   "Frmacc1128.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   180
      Width           =   350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   150
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   0
      Top             =   150
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "label9(0)"
      ForeColor       =   &H00FF0000&
      Height          =   1250
      Index           =   0
      Left            =   4950
      TabIndex        =   30
      Top             =   2460
      Width           =   5005
   End
   Begin MSForms.TextBox txtA4326 
      Height          =   795
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   4700
      VariousPropertyBits=   -1466941413
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "8290;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "發票備註："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "目前使用發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   27
      Top             =   3780
      Width           =   2205
   End
   Begin VB.Label labA4112 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4112"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2190
      TabIndex        =   26
      Top             =   3780
      Width           =   945
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "零 稅 率："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label labA4323 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4323"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   24
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label labA4319 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4319"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6090
      TabIndex        =   23
      Top             =   3780
      Width           =   945
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "發票上傳日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   4380
      TabIndex        =   22
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "未收款沖帳傳票："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label labA4317 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4317"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1980
      TabIndex        =   20
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label labA4302 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4302"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   19
      Top             =   1350
      Width           =   945
   End
   Begin VB.Label labA4305 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4305"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   18
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label labA4304 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4304"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   17
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label labA4303 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4303"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   16
      Top             =   1350
      Width           =   945
   End
   Begin MSForms.Label labA0K04 
      Height          =   255
      Left            =   1290
      TabIndex        =   15
      Top             =   930
      Width           =   6000
      VariousPropertyBits=   19
      Caption         =   "labA0K04"
      Size            =   "10583;450"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label labA0K20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K20"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   14
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label labA0K03 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K03"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   13
      Top             =   600
      Width           =   990
   End
   Begin VB.Label labA0K02 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K02"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   12
      Top             =   150
      Width           =   990
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4380
      TabIndex        =   11
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "稅　　額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4380
      TabIndex        =   10
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "銷 售 額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "統一編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4380
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收據日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4380
      TabIndex        =   4
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   60
      Top             =   30
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc1128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2022/08/24 Form2.0 已修改 labA0k04/labA0K20
Option Explicit
 
Dim adoAcc As New ADODB.Recordset, strQ As String, intQ As Integer
Dim bolUpload As Boolean, strCaseNo As String, strA4326 As String

Private Sub cmdFind_Click()
    If Text1.Text = MsgText(601) Then
        MsgBox "請輸入發票號碼！"
        Text1.SetFocus
        Text1_GotFocus
    End If
    bolUpload = False: strCaseNo = ""
    Call QueryTable
End Sub

Private Sub CmdSave_Click()
    Dim strCmd As String
    
    If FormCheck = False Then
        Exit Sub
    End If
    
    txtA4326.Text = PUB_StringFilter(txtA4326.Text) '去除換行
    strCmd = "Update Acc430 Set A4326='" & txtA4326.Text & "' Where A4301='" & Text1.Text & "' "
    intQ = 0
    cnnConnection.Execute strCmd, intQ
    If intQ = 1 Then
        Call QueryTable
        MsgBox "資料已存檔！"
    Else
        MsgBox "存檔有誤，請洽電腦中心！"
    End If
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
    
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Height = 4455
    Me.Width = 8445
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath1)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    Call Form_Clear
    Label9(0).Caption = "說明：" & vbCrLf & _
                                    "1.(非ACS案)保留系統預設值，" & vbCrLf & _
                                    "    請輸「系統預設」四字；" & vbCrLf & _
                                    "2.不需備註欄請輸「空白」二字；" & vbCrLf & _
                                    "3.其他內容請自行輸入" & vbCrLf & _
                                    "4.不可使用換行"
    labA4112.Caption = GetA4112
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StatusClear
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Set Frmacc1127 = Nothing
End Sub

Private Sub Form_Clear()
    Label9(0).Caption = ""
    labA0K02.Caption = ""
    labA0K03.Caption = ""
    labA0K04.Caption = ""
    labA0K20.Caption = ""
    labA4302.Caption = ""
    labA4303.Caption = ""
    labA4304.Caption = ""
    labA4305.Caption = ""
    labA4112.Caption = ""
    labA4317.Caption = ""
    labA4319.Caption = ""
    labA4323.Caption = ""
End Sub

Private Sub QueryTable()
    
    strQ = "Select a4301,a4302,a4303,a4304,a4305,a4317,a4319,a4321,a4323,a4326,a0k01,a0k02,a0k03,a0k04,a0k20,a0j02,st02 " & _
                "From Acc430,Acc431,Acc0k0,Acc0j0,Staff " & _
                "Where A4301='" & Text1.Text & "' And A4301=AXC01(+) And AXC02=A0k01(+) And A0K01=A0J13(+) And A0K20=ST01(+) "
    intQ = 1
    If adoAcc.State = 1 Then adoAcc.Close
    Set adoAcc = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        labA0K02.Caption = Format("" & adoAcc.Fields("a0k02"), DFormat)  '收據日
        labA0K03.Caption = "" & adoAcc.Fields("a0k03") '客戶編號
        labA0K04.Caption = "" & adoAcc.Fields("a0k04") '收據抬頭
        labA0K20.Caption = "" & adoAcc.Fields("a0k20") & " " & adoAcc.Fields("st02")  '智權人員
        strCaseNo = "" & adoAcc.Fields("a0j02") '本所案號
        labA4302.Caption = Format("" & adoAcc.Fields("a4302"), DFormat)  '發票日
        labA4303.Caption = "" & adoAcc.Fields("a4303") '統一編號
        '銷售額
        labA4304.Caption = "" & adoAcc.Fields("a4304")
        If labA4304.Caption <> MsgText(601) Then
            labA4304.Caption = Format(labA4304, DDollar)
        End If
        '稅額
        labA4305.Caption = "" & adoAcc.Fields("a4305")
        If labA4305.Caption <> MsgText(601) Then
            labA4305.Caption = Format(labA4305, DDollar)
        End If
        labA4317.Caption = "" & adoAcc.Fields("a4317") '未收款沖帳傳票編號
        
        '電子發票有作廢且上傳,顯示作廢上傳日
        If "" & adoAcc.Fields("a4321") <> MsgText(601) Then
            labA4319.Caption = Format("" & adoAcc.Fields("a4321"), DFormat) & "(作廢)"
            bolUpload = True
        ElseIf "" & adoAcc.Fields("a4319") <> MsgText(601) Then
            labA4319.Caption = Format("" & adoAcc.Fields("a4319"), DFormat)
            bolUpload = True
        End If
        
        labA4323.Caption = "" & adoAcc.Fields("a4323") '零稅率
        '發票備註(Memo 上傳換行無作用,故不使用換行)
        txtA4326.Text = "" & adoAcc.Fields("a4326")
        
        '電子發票上傳日不為空(已上傳不可再修改),鎖住發票儲存鈕
        cmdSave.Enabled = True
        txtA4326.Locked = False
        If bolUpload = True Then
            cmdSave.Enabled = False
            txtA4326.Locked = True
            'Add by Amy 2025/07/15 彈已上傳無法更改
            MsgBox "此發票已上傳盟立,無法更改！" & vbCrLf & _
                              "只能作廢重開！", vbInformation
        End If
    '無資料
    Else
        ShowNoData
    End If
End Sub

Private Function FormCheck() As Boolean
    If PUB_ChkUniText(Me) = False Then
        FormCheck = False
        Exit Function
    End If
    
    strA4326 = ""
    If bolUpload = False Then
        strA4326 = Pub_GetInvRemark(Me.Name, strCaseNo, , bolUpload, strA4326)
        'ACS 案備註不為空白,彈訊息
        If strA4326 <> txtA4326.Text And Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "ACS" And txtA4326.Text <> MsgText(601) Then
            If MsgBox("此為ACS案確定要帶此備註？" & vbCrLf & "[註]ACS案不帶備註,需輸「空白」", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                FormCheck = False
                txtA4326.SetFocus
                txtA4326_GotFocus
                Exit Function
            End If
        End If
    End If
    
    FormCheck = True
End Function

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA4326_GotFocus()
    TextInverse txtA4326
End Sub


