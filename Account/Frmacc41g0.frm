VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc41g0 
   AutoRedraw      =   -1  'True
   Caption         =   "智權期末實績保留傳票產生"
   ClientHeight    =   3684
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6648
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3684
   ScaleWidth      =   6648
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ACS 當月收款,當月分潤"
      ForeColor       =   &H00FF0000&
      Height          =   700
      Left            =   84
      TabIndex        =   17
      Top             =   1560
      Width           =   6500
      Begin VB.CommandButton CmdACS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ACS 期末傳票更正"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   5160
         Style           =   1  '圖片外觀
         TabIndex        =   21
         Top             =   130
         Width           =   1200
      End
      Begin VB.CommandButton CmdACS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ACS 期末傳票產生"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   3900
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Lbl1 
         BackStyle       =   0  '透明
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   2220
         TabIndex        =   19
         Top             =   240
         Width           =   1992
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '透明
         Caption         =   "轉ACS期末傳票號碼："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3960
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "更正傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   5220
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   840
      Width           =   1200
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2280
      TabIndex        =   16
      Top             =   2736
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "label9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   632
      Left            =   120
      TabIndex        =   22
      Top             =   70
      Width           =   6420
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "隔月初轉回傳票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   15
      Top             =   2736
      Width           =   2100
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "PS.若有轉撥資料，產生傳票後會自動開啟該傳票，請自行修改轉撥摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   96
      TabIndex        =   14
      Top             =   3336
      Width           =   6420
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(4)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   3720
      TabIndex        =   11
      Top             =   3048
      Width           =   1152
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3480
      TabIndex        =   10
      Top             =   3048
      Width           =   252
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(3)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   3048
      Width           =   1152
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "隔月初轉回傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   8
      Top             =   3048
      Width           =   2100
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   2376
      Width           =   1992
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "　轉撥傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   6
      Top             =   2376
      Width           =   1680
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   1248
      Width           =   1152
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3480
      TabIndex        =   4
      Top             =   1248
      Width           =   252
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   1236
      Width           =   1152
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   2
      Top             =   1236
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   120
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1680
   End
End
Attribute VB_Name = "Frmacc41g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 (無需修改)
'Create by Amy 2016/04/18
Option Explicit
Dim adoSalesPoint As New ADODB.Recordset, adoQ As New ADODB.Recordset
Dim i As Integer
Dim strAxb(4 To 8) As String
Dim strA0b01 As String, strA0b05 As String '會計過帳日/業績輸入關閉年月
Dim bolAxbHasDt As Boolean, bolHasAx210 As Boolean 'Acc0b1之對應Axbxx是否有資料/是否已過帳
Dim bolTrans As Boolean '是否有產生轉撥資料 for 是否run 傳票畫面
'Add by Amy 2017/10/20
Dim bolFirst As Boolean, strDate(1) As String
Dim bol0b1HasIns As Boolean 'Acc0b1 是否有Insert 當月資料
Dim strAxb02 As String, strAxb03 As String  '當月期末保留傳票日/隔月初轉回傳票日
Dim stMsg As String
Dim strAxb17(0) As String 'Add by Amy 2023/04/17 期末ACS實績保留傳票號
Dim strMaxSP01 As String, strTp As String 'Add by Amy 2023/07/14 目前智權點數輸入年月/暫存變數

Private Sub CmdACS_Click(Index As Integer)
   If FormCheck(Index + 3) = False Then Exit Sub
   
   strTp = GetACSData(0, Me.Name, Left(FCDate(MaskEdBox1.Text), 5), ",Acc020")
   If adoQ.State = adStateOpen Then adoQ.Close
   adoQ.CursorLocation = adUseClient
   adoQ.Open strTp, adoTaie, adOpenStatic, adLockReadOnly
   If adoQ.RecordCount = 0 Then
      MsgBox "無ACS收款資料！"
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   If FormSave(Index + 3) = True Then
      Call ShowBt
      'Add by Amy 2024/11/05 +訊息
      If Index = 1 Then MsgBox "ACS 期末傳票已更正！"
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim strMsg As String

    If FormCheck(Index) = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Call ClearLabel
    
    '1.產生傳票/2.更正傳票
    If FormSave(Index) = True Then
        If bolTrans = True Then
            strMsg = "已產生轉撥傳票，請自行修改轉撥傳票摘要！"
        Else
            strMsg = "實績傳票已產生！"
        End If
        MsgBox strMsg
        If bolTrans = True Then
            '開啟傳票輸入畫面
            If Lbl1(2) <> MsgText(601) Then
                With Frmacc4120
                    .Tag = Me.Name
                    Me.Hide
                    .Text1 = "1"
                    .Text2 = Lbl1(2)
                    'Add by Amy 2024/08/05 整合檢查,避免彈訊息後又可以操作
                    .MaskEdBox1 = CFDate(Pub_GetField("Acc020", "a0201='1' And a0202='" & Lbl1(2) & "'", "a0205"))
                    'Modify by Amy 2022/05/16 +if '按修改->Insert->取消->修改->Insert->存檔會出現error
                    .bolF3 = True
                    .Command3_Click
                    .bolF3 = False
                    'end 2022/05/16
                    .Show
                End With
            End If
        End If
        Call ShowBt
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    tool3_enabled
End Sub

Private Sub Form_Load()
    
    strFormName = Name
    'Modify by Amy 2023/07/14 增加ACS 當月收款,當月分潤 可產生傳票
    Me.Width = 6864  '6675
    Me.Height = 4224 '2850 'Moidfy by Amy 2023/04/17 2790
    'end 2023/07/14
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    
    Call ClearLabel
    
    strA0b01 = GetA0b01(strA0b05)
    'Add by Amy 2023/07/14 增加可先產生ACS期末傳票
    strMaxSP01 = GetMaxSP01(False) '
    Label9 = "1.「產生傳票 」需先至「每月業績開放/關閉輸入」操作「關閉」鈕" & vbCrLf & _
                     "2.「ACS 當月收款,當月分潤」可先按「ACS 期末傳票產生」，" & vbCrLf & _
                     "     再至「ACS待分潤」操作(每月業績不需關閉)"
    
    'end 2023/07/14
    bolFirst = True
    Call ShowBt
    bolFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add by Amy 2021/09/23 +判斷未更正實績保留傳票不可離開
    'Modify by Amy 2023/04/17 +strAxb17(0)
    If HasActualP(1, Left(FCDate(MaskEdBox1.Text), 5), strAxb(4), strAxb(5), strAxb17(0)) = True Then
        MsgBox "有修改資料但尚未更正傳票,請按「更正傳票」鈕！", , MsgText(5)
        Cancel = True
        Exit Sub
    'Mark by Amy 2024/06/25  拿掉不彈-婉莘
'    ElseIf bolHasAx210 = False Then
'        MsgBox "請記得過帳！" & vbCrLf & _
'                    "在確認專業點數及業務點數相同後,再通知智權主管寫報告 ！"
    End If
                    
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Call PUB_GetLock("", "Frmacc41g0") 'Add by Amy 2017/10/20
    Set Frmacc41g0 = Nothing
End Sub

Private Function FormSave(ByVal intCmd As Integer) As Boolean
    Dim strSql As String, strCmd As String, strField As String
    Dim strDate As String, OldData As String
    Dim stra0202 As String, strax203 As String
    Dim strTmp(3) As String, strMsg As String 'Modify by Amy 2023/04/17 原:strTmp(1)
    Dim strDate2 As String 'Add by Amy 2017/10/20 隔月初轉回傳票日
    Dim strAxbTP(6 To 6) As String
    Dim strax204 As String 'Add by Amy 2021/02/09
    Dim strFixAcsWhere As String, strax208 As String, strax212 As String, strTo As String  'Add by Amy 2023/04/17
    Dim strAxbTP17(0) As String, strCmd2 As String 'Add by Amy 2023/04/17
    Dim intACSState As Integer 'Add by Amy 2023/07/14
    
On Error GoTo ErrHand
    
    strDate = Left(FCDate(MaskEdBox1.Text), 5)
    strDate2 = FCDate(MaskEdBox2.Text)
    
    adoTaie.BeginTrans
    adoTaie.Execute "Update Acc0b0 Set a0b10= '01' Where a0b04 = '1'"
    
    '====== 當月ACS 分潤 (不需產生轉回傳票) ======
    'Modify by Amy 2023/07/14 從下面搬上來(原於「回轉傳票 」後做)
    'Add by Amy 2023/04/17 ACS 分潤期末保留傳票
    intACSState = FormSave_ACS(strDate, stra0202)
    If intACSState = 2 Then
         strMsg = "ACS期末傳票產生有誤,請洽電腦中心！"
         MsgBox strMsg
         adoTaie.RollbackTrans
         Exit Function
    End If
    'end 2023/07/14
    
    If intCmd < 3 Then
       If intCmd = 2 Then
           stra0202 = strAxb(4)
       End If
       
   'Modify by Amy 2017/10/20 1.國外部F41xx期末若為大於0才需產生傳票,若期末小於等於0,不需產生傳票,故迄號不可直接設於strAxb(4)
   '                                                     2.隔月初轉回傳票日寫入Axb03
   '====== 期末傳票 ======
   '*** 國外部F41xx ***
       '國外部F41xx-借方
       strSql = GetPoint_SP(strDate, strDate, , , "F41XX", False, Me.Name, , True)
       If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
       adoSalesPoint.CursorLocation = adUseClient
       adoSalesPoint.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
       If adoSalesPoint.RecordCount <> 0 Then
           '按「產生傳票」鈕
           If intCmd = 1 Then
               stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
               If strAxb(4) = MsgText(601) Then strAxb(4) = stra0202
               strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                               "Values('1','" & stra0202 & "', " & Val(FCDate(MaskEdBox1.Text)) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
               adoTaie.Execute strCmd
               strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(FCDate(MaskEdBox1.Text), 1, 3), Mid(FCDate(MaskEdBox1.Text), 4, 2))
           '按「更正傳票」鈕
           Else
               strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
               adoTaie.Execute strCmd
           End If
            
           With adoSalesPoint
               Do While .EOF = False
                   strax203 = GetSeqNo("1", stra0202)
                   Select Case .Fields("SP02")
                       'Modify by Amy 2021/02/09 ax204拆開並改ax212名稱
                       Case "F4101"
                           strax204 = "FCL"
                           strTmp(0) = strax204
                       'modify by sonia 2021/1/22 +F4104,F4105
                       Case "F4102"
                           strax204 = "FCP"
                           strTmp(0) = strax204
                       Case "F4104"
                           strax204 = "FCP"
                           strTmp(0) = "專利國外部"
                       Case "F4105"
                           strax204 = "FCP"
                           strTmp(0) = "專利日本部"
                       'modify by sonia 2021/1/22 +F4106,F4107
                       Case "F4103"
                           strax204 = "FCT"
                           strTmp(0) = strax204
                       Case "F4106"
                           strax204 = "FCT"
                           strTmp(0) = "FCT英文組"
                       Case "F4107"
                           strax204 = "FCT"
                           strTmp(0) = "FCT日文組"
                   End Select
                   'end 2021/02/09
                   'Modify by Amy 2021/02/09 ax204改抓strAx204變數 原:strTmp(0)
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                               "Values('1','" & stra0202 & "', '" & strax203 & "','" & strax204 & "','4192'," & _
                                Val(.Fields("SP15")) * 1000 & ",0,'" & .Fields("SP02") & "','" & strTmp(0) & "/保留')"
                   adoTaie.Execute strCmd
                   .MoveNext
               Loop
           End With
           
           '國外部F41xx-貸方
           strSql = GetPoint_SP(strDate, strDate, , , "F41XX", False, Me.Name, True, True)
           If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
           adoSalesPoint.CursorLocation = adUseClient
           adoSalesPoint.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
           If adoSalesPoint.RecordCount <> 0 Then
               strax203 = GetSeqNo("1", stra0202)
               strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                               "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                                "0," & Val(adoSalesPoint.Fields("SP15")) * 1000 & ",'國外部/保留','國外部')"
                   adoTaie.Execute strCmd
           End If
           If intCmd = 2 Then
               stra0202 = Left(stra0202, 6) & Val(Right(stra0202, 4)) + 1
           End If
       End If
   '*** End 國外部F41xx ***
       
   '*** 特殊員編 ***
       'Add by Amy 2021/05/20 +特殊員編(W1001/W2001/P2005/P1005)
       OldData = ""
       strSql = GetPoint_SP(strDate, strDate, , , "SpecNo", False, Me.Name, False, True)
       If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
       adoSalesPoint.CursorLocation = adUseClient
       adoSalesPoint.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
       If adoSalesPoint.RecordCount <> 0 Then
           '按「產生傳票」鈕
           If intCmd = 1 Then
               stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
               If strAxb(4) = MsgText(601) Then
                   strAxb(4) = stra0202
               Else
                   strAxb(5) = stra0202
               End If
               strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                               "Values('1','" & stra0202 & "', " & Val(FCDate(MaskEdBox1.Text)) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
               adoTaie.Execute strCmd
               strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(FCDate(MaskEdBox1.Text), 1, 3), Mid(FCDate(MaskEdBox1.Text), 4, 2))
           '按「更正傳票」鈕
           Else
               strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
               adoTaie.Execute strCmd
           End If
           With adoSalesPoint
               .MoveFirst
               Do While .EOF = False
                   strax203 = GetSeqNo("1", stra0202)
                   Select Case .Fields("SP02")
                       'Modify by Amy 2022/05/13 +P1005
                       Case "W1001", "W2001", "P1005"
                           strax204 = "P"    '2021/9/9 因為會影響秘書專業點數報表,故由W部門改為P部門(1公司D110082572)
                       Case "P2005"
                           strax204 = "T"
                   End Select
                   
                   '借方
                   strax203 = GetSeqNo("1", stra0202)
                   strTmp(0) = Replace(GetStaffName(.Fields("SP02"), True), "商標部", "") & "/保留"
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                   "Values('1','" & stra0202 & "', '" & strax203 & "','" & strax204 & "','4191'," & _
                                   Val(.Fields("SP15")) * 1000 & ",0,'" & .Fields("SP02") & "','" & strTmp(0) & "')"
                   adoTaie.Execute strCmd
                   
                   '貸方
                   strax203 = GetSeqNo("1", stra0202)
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                                           "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                                           "0," & Val(.Fields("SP15")) * 1000 & ",'" & strTmp(0) & "','" & Replace(strTmp(0), "/保留", "") & "')"
                   adoTaie.Execute strCmd
                  
                   .MoveNext
               Loop
           End With
           If intCmd = 2 Then
               stra0202 = Left(stra0202, 6) & Val(Right(stra0202, 4)) + 1
           End If
       End If
       'end 2021/05/20
   '*** End 特殊員編 ***
       
   '*** 智權部門 ***
       OldData = ""
       strSql = GetPoint_SP(strDate, strDate, "S10", "S99", , False, Me.Name, False, True)
       If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
       adoSalesPoint.CursorLocation = adUseClient
       adoSalesPoint.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
       If adoSalesPoint.RecordCount <> 0 Then
           With adoSalesPoint
               Do While .EOF = False
                   If Left(OldData, 2) <> Left(.Fields("SP48"), 2) Then
                       If OldData <> MsgText(601) Then
                           '貸方
                           strSql = GetPoint_SP(strDate, strDate, Left(OldData, 2) & "0", Left(OldData, 2) & "9", , False, Me.Name, True, True)
                           If adoQ.State = adStateOpen Then adoQ.Close
                           adoQ.CursorLocation = adUseClient
                           adoQ.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
                           If adoQ.RecordCount <> 0 Then
                               strax203 = GetSeqNo("1", stra0202)
                               strTmp(0) = PUB_GetZone(OldData)
                               strTmp(1) = Val(adoQ.Fields("SP15")) * 1000
                               strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                                           "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                                           "0," & strTmp(1) & ",'" & strTmp(0) & "/保留','" & strTmp(0) & "')"
                               adoTaie.Execute strCmd
                               If intCmd = 2 Then
                                   If Left(stra0202, 6) & Val(Right(stra0202, 4)) + 1 <= strAxb(5) Then
                                       stra0202 = Left(stra0202, 6) & Val(Right(stra0202, 4)) + 1
                                   Else
                                       MsgBox "智權部門 更正傳票有誤，請確認！"
                                       adoTaie.RollbackTrans
                                       If adoQ.State = adStateOpen Then adoQ.Close
                                       If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
                                       Exit Function
                                   End If
                               End If
                           End If
                       End If
                       If intCmd = 1 Then
                           stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
                           If strAxb(4) = MsgText(601) Then
                               strAxb(4) = stra0202
                           Else
                               strAxb(5) = stra0202
                           End If
                           strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                       "Values('1','" & stra0202 & "', " & Val(FCDate(MaskEdBox1.Text)) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                           adoTaie.Execute strCmd
                           strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(FCDate(MaskEdBox1.Text), 1, 3), Mid(FCDate(MaskEdBox1.Text), 4, 2))
                       Else
                           strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
                           adoTaie.Execute strCmd
                       End If
                   End If
                   '借方
                   strax203 = GetSeqNo("1", stra0202)
                   strTmp(0) = GetStaffName(.Fields("SP02"), True) & "/保留"
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                   "Values('1','" & stra0202 & "', '" & strax203 & "','P','4191'," & _
                                   Val(.Fields("SP15")) * 1000 & ",0,'" & .Fields("SP02") & "','" & strTmp(0) & "')"
                   adoTaie.Execute strCmd
                       
                   OldData = .Fields("SP48")
                   strTmp(1) = Val(adoSalesPoint.Fields("SP15")) * 1000
                   .MoveNext
               Loop
           End With
           '智權部門-最後一筆貸方
           strSql = GetPoint_SP(strDate, strDate, Left(OldData, 2) & "0", Left(OldData, 2) & "9", , False, Me.Name, True, True)
           If adoQ.State = adStateOpen Then adoQ.Close
           adoQ.CursorLocation = adUseClient
           adoQ.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
           If adoQ.RecordCount <> 0 Then
               strax203 = GetSeqNo("1", stra0202)
               strTmp(0) = PUB_GetZone(OldData)
               strTmp(1) = Val(adoQ.Fields("SP15")) * 1000
               strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                               "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                               "0," & strTmp(1) & ",'" & strTmp(0) & "/保留','" & strTmp(0) & "')"
               adoTaie.Execute strCmd
           End If
           Lbl1(0) = strAxb(4): Lbl1(1) = strAxb(5)
       End If
   '*** End 智權部門 ***
   If strAxb(5) = MsgText(601) Then strAxb(5) = strAxb(4) '避免無迄號
           
   '====== 回轉傳票 ======
       OldData = "": stra0202 = "": strax203 = ""
   
       strSql = "Select * From Acc021 Where ax201='1' And ax202>='" & strAxb(4) & "' And ax202<='" & strAxb(5) & "' " & _
                   "Order by ax202,Decode(ax206,0,'000',ax203)"
       If adoQ.State = adStateOpen Then adoQ.Close
       adoQ.CursorLocation = adUseClient
       adoQ.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
       If adoQ.RecordCount <> 0 Then
           With adoQ
               Do While .EOF = False
                   If OldData <> .Fields("Ax202") Then
                       If intCmd = 1 Then
                           stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(strDate2, 1, 3)), Val(Mid(strDate2, 4, 2)))
                           strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                           "Values('1','" & stra0202 & "', " & Val(strDate2) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                           adoTaie.Execute strCmd
                           strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(strDate2, 1, 3), Mid(strDate2, 4, 2))
                           If OldData = MsgText(601) Then strAxb(7) = stra0202
                       ElseIf intCmd = 2 Then
                           If stra0202 = MsgText(601) Then
                               stra0202 = strAxb(7)
                           Else
                               stra0202 = Left(stra0202, 1) & Val(Mid(stra0202, 2)) + 1
                           End If
                           strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
                           adoTaie.Execute strCmd
                       End If
                   End If
                   
                   strField = "": strCmd = ""
                   strax203 = GetSeqNo("1", stra0202)
                   For i = 4 To .Fields.Count
                       If IsNull(.Fields(i - 1)) = False Then
                           strField = strField & "," & GetAxName(i)
                           If i = 6 Or i = 7 Or i = 10 Then
                               strTmp(0) = ""
                               Select Case i
                                   Case 6
                                       '原借變貸
                                       strTmp(0) = .Fields(6)
                                   Case 7
                                       '原借變貸
                                       strTmp(0) = .Fields(5)
                                   Case 10
                                       strTmp(0) = .Fields(i)
                               End Select
                               strCmd = strCmd & "," & Val(strTmp(0))
                           Else
                               strCmd = strCmd & ",'" & .Fields(i - 1) & "'"
                           End If
                       End If
                   Next i
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203" & strField & ") " & "Values('1','" & stra0202 & "','" & strax203 & "'" & strCmd & ")"
                   adoTaie.Execute strCmd
                       
                   OldData = .Fields("Ax202")
                   .MoveNext
               Loop
           End With
           strAxb(8) = stra0202: Lbl1(3) = strAxb(7): Lbl1(4) = strAxb(8)
       End If
           
   '====== 實績轉撥傳票 ======
       strDate = Left(GetPreMonLastDate(strSrvDate(1)), 5)
       strSql = GetPoint_SP(strDate, strDate, , , "SP19", False, Me.Name, , True)
       If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
       adoSalesPoint.CursorLocation = adUseClient
       adoSalesPoint.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
       If adoSalesPoint.RecordCount <> 0 Then
           bolTrans = True
           With adoSalesPoint
               '若產生傳票時並無實績轉撥資料,但做更正傳票前又加了實績轉撥資料則再多加一張傳票
               If intCmd = 1 Or (intCmd = 2 And strAxb(6) = MsgText(601)) Then
                   stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
                   strAxb(6) = stra0202
                   strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                    "Values('1','" & stra0202 & "', " & Val(FCDate(MaskEdBox1.Text)) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                   adoTaie.Execute strCmd
                   strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(FCDate(MaskEdBox1.Text), 1, 3), Mid(FCDate(MaskEdBox1.Text), 4, 2))
               Else
                   stra0202 = strAxb(6)
                   strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
                   adoTaie.Execute strCmd
               End If
               
               Do While .EOF = False
                   strax203 = GetSeqNo("1", stra0202)
                   strTmp(0) = "": strTmp(1) = ""
                   If Val(.Fields("SP19")) < 0 Then
                       strTmp(0) = Abs(.Fields("SP19"))
                   Else
                       strTmp(1) = .Fields("SP19")
                   End If
                   'Modify by Amy 2022/05/13 取代換行
                   strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                   "Values('1','" & stra0202 & "', '" & strax203 & "','P','411101'," & Val(strTmp(0)) * 1000 & "," & Val(strTmp(1)) * 1000 & _
                                   ",'" & .Fields("SP02") & "','" & GetStaffName(.Fields("SP02"), True) & "：轉撥" & Replace("" & .Fields("SP20"), vbCrLf, "") & "')"
                   adoTaie.Execute strCmd
                   
                   .MoveNext
               Loop
           End With
           Lbl1(2) = strAxb(6)
       End If
       'Modify by Amy 2025/09/03 從國外部更新傳票程式中搬下來,只要有按「更新傳票」就更新日期
       '     1140903 建慈過帳後,再取消過帳,修改1 公司 D114081462 對沖業務-85035(陳建宏) 之非金額欄位,不需修改點數資料,而一直彈有傳票資料未更正
       If intCmd = 2 Then
         'Add by Amy 2018/06/13 更新傳票修改日期時間
         strCmd = "Update Acc020 set A0209=" & Val(strSrvDate(2)) & ",A0210=" & ServerTime & ",A0211='" & strUserNum & "' " & _
                           "Where A0201='1' And (A0202>='" & strAxb(4) & "' And A0202<='" & strAxb(5) & "' " & _
                           "Or A0202>='" & strAxb(7) & "' And A0202<='" & strAxb(8) & "' " & IIf(strAxb(6) = "", "", " Or A0202='" & strAxb(6) & "' ") & ") "
         adoTaie.Execute strCmd
       End If
    End If 'End intCmd < 3
    
    '更新Acc01b對應傳票日期及號碼
    strExc(0) = "": strCmd = "": strField = ""
    
    'Modify by Amy 2023/07/14 ACS 期末保留傳票可先產生
    'Acc0b1[沒]當月資料
    If bol0b1HasIns = False Then
        If intCmd < 3 Then
            For i = LBound(strAxb) To UBound(strAxb)
                strField = strField & ",axb0" & i
                strCmd = strCmd & ",'" & strAxb(i) & "'"
            Next i
        End If
        'Add by Amy 2023/06/06+strAxb17(0) ACS期末保留
        If strAxb17(0) <> MsgText(601) Then
            strField = strField & ",axb17"
            strCmd = strCmd & ",'" & strAxb17(0) & "'"
        End If
        'ACS期末保留產生傳票
        If intCmd = 3 Then
            strCmd = "Insert Into Acc0b1 (axb01,axb02" & strField & ") " & _
                  "Values(" & strDate & "," & Val(FCDate(MaskEdBox1.Text)) & strCmd & ")"
         Else
            strCmd = "Insert Into Acc0b1 (axb01,axb02,axb03" & strField & ") " & _
                     "Values(" & strDate & "," & Val(FCDate(MaskEdBox1.Text)) & "," & Val(FCDate(MaskEdBox2.Text)) & strCmd & ")"
         End If
    '當月期末傳票日尚未產生(frmacc41j0-隱藏版可能先產生資料)
    ElseIf strAxb02 = MsgText(601) Or intCmd = 1 Then
        If intCmd < 3 Then
            For i = LBound(strAxb) To UBound(strAxb)
                strField = strField & ",axb0" & i & "=" & "'" & strAxb(i) & "'"
            Next i
         End If
         Call bolAcc0b1(9, Left(FCDate(MaskEdBox1.Text), 5), strAxbTP17())
         '原本 沒有 ACS,現在有
         If strAxb17(0) <> MsgText(601) And strAxbTP17(0) = MsgText(601) Then
            strField = strField & ",axb17" & "=" & "'" & strAxb17(0) & "'"
         End If
         '期末保留傳票日期
         If strAxb02 = MsgText(601) Then
            strField = strField & ",axb02=" & Val(FCDate(MaskEdBox1.Text))
         End If
         If intCmd < 3 Then
            strField = strField & ",axb03=" & Val(FCDate(MaskEdBox2.Text))
         End If
         strCmd = "Update Acc0b1 Set " & Mid(strField, 2) & " Where axb01=" & Val(Left(FCDate(MaskEdBox1.Text), 5))
    '更正後 有 轉撥資料 or 更正 ACS期末傳票
    ElseIf intCmd = 2 Or intCmd = 4 Then
       If intCmd = 2 And strAxb(6) <> MsgText(601) Then
            Call bolAcc0b1(6, Left(FCDate(MaskEdBox1.Text), 5), strAxbTP())
            '原本 沒有 轉撥
            If strAxbTP(6) = MsgText(601) Then
                strCmd = "Update Acc0b1 Set axb06='" & strAxb(6) & "' Where axb01=" & Val(Left(FCDate(MaskEdBox1.Text), 5))
            End If
        End If
        'Add by Amy 2023/04/14 '更正後 有 ACS資料
        Call bolAcc0b1(9, Left(FCDate(MaskEdBox1.Text), 5), strAxbTP17())
        If strAxb17(0) <> MsgText(601) And strAxbTP17(0) = MsgText(601) Then
            strCmd2 = "Update Acc0b1 Set axb17='" & strAxb17(0) & "' Where axb01=" & Val(Left(FCDate(MaskEdBox1.Text), 5))
        End If
    End If
    'end 2023/07/14
    'end 2017/10/20
    
    If strCmd <> MsgText(601) Then
        adoTaie.Execute strCmd
    End If
    'Add by Amy 2023/04/17
    If strCmd2 <> MsgText(601) Then
        adoTaie.Execute strCmd2
    End If
    
    'Modify by Amy 2023/07/14 +ACS 期末傳票,可能adoSalesPoint未使用
    If adoQ.State = adStateOpen Then adoQ.Close
    If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
    
    adoTaie.Execute "Update Acc0b0 Set a0b10= null Where a0b04 = '1'"
    adoTaie.CommitTrans
    
    FormSave = True
    Exit Function
    
ErrHand:
    adoTaie.RollbackTrans
    If adoQ.State = adStateOpen Then adoQ.Close
    If adoSalesPoint.State = adStateOpen Then adoSalesPoint.Close
    bolTrans = False
    MsgBox Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Function

'Add by Amy 2023/07/14 從FromSave把ACS 期末傳票拆出,回傳:0-執行 /1-執行完成 /2-錯誤
Private Function FormSave_ACS(ByVal strDate As String, ByRef stra0202 As String) As Integer
    Dim strCmd As String, strTmp(3) As String, strax203 As String, strax208 As String, strax212 As String, strTo As String

On Error GoTo ErrHand1
    
    FormSave_ACS = 0: stMsg = ""
    'Add by Amy 2023/04/17 ACS 分潤期末保留傳票
    '抓取ACS當月有收款之傳票,依案號產生分錄寫於一張傳票中
    strSql = GetACSData(0, Me.Name, strDate, ",Acc020")
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        With adoQ
            '按「產生傳票」/「ACS期末傳票產生」鈕 or 產生傳票時並無ACS實績資料,但做更正傳票前有系統-1個月(收入傳票可改日期)收入資料則再多加一張傳票
            If strAxb17(0) = MsgText(601) Then
                stra0202 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
                strAxb17(0) = stra0202
                strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                "Values('1','" & stra0202 & "', " & Val(FCDate(MaskEdBox1.Text)) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                adoTaie.Execute strCmd
                strCmd = AccSaveAutoNo(MsgText(801), Right(stra0202, 4), Mid(FCDate(MaskEdBox1.Text), 1, 3), Mid(FCDate(MaskEdBox1.Text), 4, 2))
            '按「更正傳票」鈕
            Else
                stra0202 = strAxb17(0)
                strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stra0202 & "' "
                adoTaie.Execute strCmd
            End If
            .MoveFirst
            Do While .EOF = False
                '抓取「當月」此案會計科目420101貸方且對沖-業務=M0101傳票號最小之收款傳票的客戶及對沖-其他(記錄收文智權人員),避免換智權不在外層抓
                strTmp(0) = "" & .Fields("ax214")
                strTmp(1) = Val(.Fields("ax207"))
                
                strTmp(3) = "Y"
                strTmp(2) = GetACSData("1", Me.Name, strDate, ",Acc020", "And ax214='" & strTmp(0) & "'", strTmp(3))
                strax208 = Mid(strTmp(3), 1, Val(InStr(strTmp(3), ";")) - 1) '對沖-客戶
                strTmp(3) = Replace(strTmp(3), strax208 & ";", "")
                strax212 = Mid(strTmp(3), 1, Val(InStr(strTmp(3), ";")) - 1) '摘要
                strTmp(3) = Replace(strTmp(3), strax212 & ";", "") '對沖-其他
                
                '資料不是只有一筆
                stMsg = ""
                If strTmp(2) <> "1" Then
                    '借方
                    strax203 = GetSeqNo("1", stra0202)
                    stMsg = stMsg & "本所案號：" & strTmp(0) & " (「當月」收款資料" & strTmp(2) & "筆，資料無法自動帶)" & vbCrLf
                    stMsg = stMsg & "　　項次：" & strax203 & "(借)" & vbCrLf & _
                                                  "　　　　　需補「對沖-客戶」、「對沖-其他(補收文智權人員)」及「摘要」" & vbCrLf
                                                  
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax214) " & _
                                    "Values('1','" & stra0202 & "', '" & strax203 & "','W','4191'," & _
                                    strTmp(1) & ",0,'M0101','" & strTmp(0) & "')"
                    adoTaie.Execute strCmd
                    
                    '貸方
                    strax203 = GetSeqNo("1", stra0202)
                    stMsg = stMsg & "　　項次：" & strax203 & "(貸)" & vbCrLf & _
                                                  "　　　　　需補「對沖-客戶」及「摘要」" & vbCrLf & vbCrLf & vbCrLf
                                                  
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax213,ax214) " & _
                                    "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                                    "0," & strTmp(1) & ",'M0101','顧服組','" & strTmp(0) & "')"
                    adoTaie.Execute strCmd
                    
                Else
                    '借方
                    strax203 = GetSeqNo("1", stra0202)
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax213,ax214) " & _
                                    "Values('1','" & stra0202 & "', '" & strax203 & "','W','4191'," & _
                                    strTmp(1) & ",0,'" & strax208 & "','M0101','" & strax212 & "','" & strTmp(3) & "','" & strTmp(0) & "')"
                    adoTaie.Execute strCmd
                    
                    '貸方
                    strax203 = GetSeqNo("1", stra0202)
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax213,ax214) " & _
                                    "Values('1','" & stra0202 & "', '" & strax203 & "','TOT','2492'," & _
                                    "0," & strTmp(1) & ",'" & strax208 & "','M0101','" & strax212 & "','顧服組','" & strTmp(0) & "')"
                    adoTaie.Execute strCmd
                End If
                
                .MoveNext
            Loop
        End With
        Lbl1(5) = strAxb17(0)
        
        If stMsg <> MsgText(601) Then
            strTo = "A2004"
            stMsg = "ACS 期末實績保留傳票資料有問題，需補１公司「" & stra0202 & "」傳票，其項次 / 案號資料如下：" & vbCrLf & vbCrLf _
                         & stMsg & "以上資料不齊，請補上！！"
            PUB_SendMail strUserNum, strTo, "", "ACS分潤期末實績保留資料有誤（需補資料）", stMsg
        End If
    End If
    FormSave_ACS = 1
    Exit Function

ErrHand1:
   FormSave_ACS = 2

End Function

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim strLabel As String
    
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
    End If
    
    strLabel = "轉期末傳票日期"
    If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox1.Text))) = False Then
        MsgBox strLabel & "輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    'Modify by Amy 2017/10/20 ＋需大於A0b05
    'Modify by Amy 2023/07/14 改抓目前點數輸入最大年月
    If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(Left(strSrvDate(2), 5)) _
     Or Val(Left(FCDate(MaskEdBox1.Text), 5)) < Val(strMaxSP01) Then
        MsgBox strLabel & "需小於系統月份且大於業績輸入關閉日！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If ChkWorkDay(Val(FCDate(MaskEdBox1.Text)) + 19110000) = False Then
        MsgBox strLabel & "需是工作日！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If MaskEdBox1.Enabled = True Then
        If ChkWorkData("1", DBDATE(MaskEdBox1.Text), stMsg) = False Then
            MsgBox strLabel & stMsg, , MsgText(5)
            Cancel = True
            MaskEdBox1.SetFocus
            Exit Sub
        End If
    End If
    
    Call ShowBt
End Sub

Private Sub ClearLabel()
    Dim objLbl As LABEL
    
    For Each objLbl In Lbl1
        objLbl = ""
    Next
End Sub

'intCmd:1-產生傳票/2-更正傳票/3-ACS期末傳票產生/4-ACS期未傳票更正
Private Function FormCheck(ByVal intCmd As Integer) As Boolean
    Dim strSql As String, strLabel As String, bolCancel As Boolean
        
    FormCheck = False: bolCancel = False
    strLabel = "轉期末傳票日期"
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox strLabel & "不可為空值！", , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
    Call MaskEdBox1_Validate(bolCancel)
    If bolCancel = True Then Exit Function
    
    'Mark by Amy 2019/09/03 改陳經理及王文安輸不需判斷
'    If ChkF41xxInput = False Then
'        MsgBox "國外部實績資料尚未輸入！", , MsgText(5)
'        Exit Function
'    End If
    'Modify by Amy 2023/07/14 +if
    If intCmd < 3 Then
         strLabel = "隔月初轉回傳票日期"
         If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
             MsgBox strLabel & "不可為空值！", , MsgText(5)
             MaskEdBox2.SetFocus
             Exit Function
         End If
         Call MaskEdBox2_Validate(bolCancel)
         If bolCancel = True Then Exit Function
    End If
    
    'Modify by Amy 2023/07/14 ACS期末傳票相關按鈕
    stMsg = ""
    '產生傳票/ACS期末傳票產生
    If intCmd = 1 Or intCmd = 3 Then
        If Val(Left(FCDate(MaskEdBox1.Text), 5)) < Val(業績自動轉傳票啟用年月) Then
            MsgBox "此程式於 " & 業績自動轉傳票啟用年月 & " 月開始使用！", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
        strTp = strAxb(4)
        If intCmd = 3 Then strTp = strAxb17(0): stMsg = "ACS" 'ACS期末保留傳票號
        If Val(Left(GetAxb0203(1), 5)) >= Val(Left(FCDate(MaskEdBox1.Text), 5)) And strTp <> MsgText(601) Then
            MsgBox "此年月已產生" & stMsg & "期末傳票！", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
    End If
    
    '判斷作業狀態(A0b10) 是否有值
    If Pub_GetAcc0b0("A0b10", "1") <> MsgText(601) Then
        MsgBox MsgText(197), , MsgText(5)
        Exit Function
    End If
    
    FormCheck = True
End Function

'判斷SalesPoint 員編F41XX是否已輸入(其中一個有輸就算有輸)
Private Function ChkF41xxInput() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    ChkF41xxInput = False
    
    'modify by sonia 2021/1/22 +F4104~F4107
    strQ = "Select sp02,Nvl(sp15||sp19||sp36||sp40,'N') as YData From SalesPoint Where sp02 in('F4101','F4102','F4103','F4104','F4105','F4106','F4107') And sp01=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100 & _
                " Order by Nvl(sp15||sp19||sp36||sp40,'N')"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If RsQ.Fields("YData") <> "N" Then
            ChkF41xxInput = True
        End If
    End If
    RsQ.Close
End Function

'取得Acc021傳票資料-交易檔欄位名
Private Function GetAxName(ByVal intSeq As Integer) As String
    If intSeq >= 10 Then
        Select Case intSeq
            Case 10
                GetAxName = "ax214"
            Case 11
                GetAxName = "ax211"
            Case 13
                GetAxName = "ax212"
            Case 14
                GetAxName = "ax213"
            Case 15
                GetAxName = "ax215"
        End Select
    Else
        GetAxName = "ax20" & intSeq
    End If
End Function

Private Sub ShowBt()
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    Frame1.Enabled = False 'Add by Amy 2023/07/14
    'Modify by Amy 2017/10/20
    If bolFirst = True Then
        strDate(0) = GetPreMonLastDate(strSrvDate(1))
        strDate(1) = Pub_GetMaxA0205("1", Left(strSrvDate(2), 5))
        '若當月尚無最大傳票日則預設系統日 Ex:1061002進入時為當月第一天工作日
        If strDate(1) = MsgText(601) Then strDate(1) = strSrvDate(2)
    Else
        strDate(0) = FCDate(MaskEdBox1.Text)
        strDate(1) = FCDate(MaskEdBox2.Text)
    End If
    strAxb02 = GetAxb0203(1, Left(strDate(0), 5))
    strAxb03 = GetAxb0203(2, Left(strDate(0), 5))
    
    '當月期末保留傳票日(strAxb02) 有值 則不可修改,沒值 預設上個月最後一個工作日
    If strAxb02 <> MsgText(601) Then
        strDate(0) = strAxb02
        MaskEdBox1.Enabled = False
    Else
        MaskEdBox1.Enabled = True
    End If
    '隔月初轉回傳票日(strAxb03) 有值 則不可修改,沒值 預設1公司最大傳票日期
    If strAxb03 <> MsgText(601) Then
        strDate(1) = strAxb03
        MaskEdBox2.Enabled = False
    Else
        MaskEdBox2.Enabled = True
    End If
 
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = CFDate(strDate(0))
    MaskEdBox1.Mask = DFormat
    
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = CFDate(strDate(1))
    MaskEdBox2.Mask = DFormat
    'end 2017/10/20
    
    '抓智權點數傳票起始值
    bolAxbHasDt = bolAcc0b1(1, Left(FCDate(MaskEdBox1.Text), 5), strAxb())
    'Add by Amy 2017/10/20 判斷Acc0b1是否已有當月資料
    If bolAxbHasDt = False Then
        bol0b1HasIns = ExistCheck("Acc0b1", "Axb01", Left(FCDate(MaskEdBox1.Text), 5), "", False)
    Else
        bol0b1HasIns = True
    End If
    Call bolAcc0b1(9, Left(FCDate(MaskEdBox1.Text), 5), strAxb17()) 'Add by Amy 2023/04/17 ACS期末
   
    '畫面日期等於業績輸入關閉年月
    If Val(Left(FCDate(MaskEdBox1.Text), 5)) = Val(strA0b05) Then
        If bolAxbHasDt = True Then
            bolHasAx210 = Pub_ChkAxbPost(strAxb(4), strAxb(5), strAxb(6))
            '未產生傳票才可按「產生傳票鈕」
            If strAxb(4) = MsgText(601) Then
                Command1(1).Enabled = True
            '已有傳票號尚未過帳只能按「更正傳票鈕」
            ElseIf bolHasAx210 = False Then
                Command1(2).Enabled = True
            End If
        Else
            Command1(1).Enabled = True
        End If
    'Add by Ａｍｙ2023/07/14 ACS 可當月收款,當月分潤(點數輸入已開放,畫面日期=目前最大點數輸入年月)
    Else
         CmdACS(0).Enabled = False: CmdACS(1).Enabled = False
         strTp = Val(Mid(DBDATE(DateAdd("m", -2, Format(strSrvDate(1), "####/##/##"))), 1, 6)) - 191100
         If strA0b05 = strTp And Val(Left(FCDate(MaskEdBox1.Text), 5)) = strMaxSP01 Then
            Frame1.Enabled = True
            If strAxb17(0) = MsgText(601) Then
               CmdACS(0).Enabled = True
            Else
               CmdACS(1).Enabled = True
            End If
         End If
    End If
    Lbl1(0) = strAxb(4): Lbl1(1) = strAxb(5): Lbl1(2) = strAxb(6): Lbl1(3) = strAxb(7): Lbl1(4) = strAxb(8)
    Lbl1(5) = strAxb17(0) 'Add by Amy 2023/04/17
End Sub

'Add by Amy 2019/09/20
Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    Dim strLabel As String
    
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        Exit Sub
    End If
    
    strLabel = "隔月初轉回傳票日期"
    If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox2.Text))) = False Then
        MsgBox strLabel & "輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    '不可大於系統日
     If Val(FCDate(MaskEdBox2.Text)) > Val(strSrvDate(2)) Then
        MsgBox strLabel & "不可大於系統日！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If Val(FCDate(MaskEdBox1.Text)) >= Val(FCDate(MaskEdBox2.Text)) Then
        MsgBox strLabel & "不可小於轉期末傳票日", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If MaskEdBox2.Enabled = True Then
        If ChkWorkData("1", DBDATE(MaskEdBox2.Text), stMsg) = False Then
            MsgBox strLabel & stMsg, , MsgText(5)
            Cancel = True
            MaskEdBox2.SetFocus
            Exit Sub
        End If
    End If
End Sub
'end 2017/10/20


