VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1480 
   AutoRedraw      =   -1  'True
   Caption         =   "國內帳齡分析表"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   5565
   Begin VB.TextBox Text4 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1740
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1500
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2448
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   3135
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text5 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1272
      Width           =   612
   End
   Begin VB.TextBox Text3 
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
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   2
      Top             =   564
      Width           =   945
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      MaxLength       =   9
      TabIndex        =   1
      Top             =   204
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   0
      Top             =   204
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   924
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS:分析內容選2時不考慮客戶代號條件，以免遺漏"
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
      Height          =   204
      Left            =   252
      TabIndex        =   17
      Top             =   2150
      Width           =   4728
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.已發文 2.未發文 空白:全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   2196
      TabIndex        =   16
      Top             =   1776
      Width           =   2988
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "是否發文"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   1776
      Width           =   1212
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   2436
      TabIndex        =   14
      Top             =   600
      Width           =   1356
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.客戶 2.智權人員   3.智權人員、客戶)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2196
      TabIndex        =   13
      Top             =   1260
      Width           =   2052
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "分析內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳款截止日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   9
      Top             =   204
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoaccrpt108 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
'Remove by Lydia 2016/12/09
'Dim dllaccrpt108 As Object
'Add By Sindy 2014/6/4
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim intPage As Integer
'2014/6/4 END

'列印
Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt108Delete
   ProduceData
   'Remove by Lydia 2016/12/09
'   If adoaccrpt108.State = adStateOpen Then
'      adoaccrpt108.Close
'   End If
'   adoaccrpt108.CursorLocation = adUseClient
'   adoaccrpt108.Open "select * from accrpt108", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt108.RecordCount <> 0 Then
'      Select Case Text5
'         Case "1"
'            dllaccrpt108.Acc1480 ReportTitle(1081), Text1, Text2, MaskEdBox1.Text, "", "", strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Case "2"
'            'Modify by Morgan 2007/10/2 智權人員範圍改成一個
'            'dllaccrpt108.Acc1480 ReportTitle(1082), Text3, Text4, MaskEdBox1.Text, "", "", strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'            dllaccrpt108.Acc1480 ReportTitle(1082), Text3, Text3, MaskEdBox1.Text, "", "", strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Case Else
'            If adoquery.State = adStateOpen Then
'               adoquery.Close
'            End If
'            adoquery.CursorLocation = adUseClient
'            'Modify by Morgan 2007/10/2 智權人員範圍改成一個
'            'adoquery.Open "select distinct st01, st02 from staff, accrpt108 where st01 = r10809 and st01 >= '" & Text3 & "' and st01 <= '" & Text4 & "' order by st01 asc", adoTaie, adOpenStatic, adLockReadOnly
'            adoquery.Open "select distinct st01, st02 from staff, accrpt108 where st01 = r10809 and st01 = '" & Text3 & "' order by st01 asc", adoTaie, adOpenStatic, adLockReadOnly
'            'end 2007/10/2
'            Do While adoquery.EOF = False
'               'Modify by Morgan 2007/10/2 智權人員範圍改成一個
'               'dllaccrpt108.Acc1480 ReportTitle(1083), Text3, Text4, MaskEdBox1.Text, adoquery.Fields("st01").Value, adoquery.Fields("st02").Value, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'               dllaccrpt108.Acc1480 ReportTitle(1083), Text3, Text3, MaskEdBox1.Text, adoquery.Fields("st01").Value, adoquery.Fields("st02").Value, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'               'end 2007/10/2
'               adoquery.MoveNext
'            Loop
'            adoquery.Close
'      End Select
'   End If
'   adoaccrpt108.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
End Sub

'Excel
Private Sub Command2_Click()
   'Add By Sindy 2014/6/4
   If Text5 <> "" And Text5 <> "1" And Text5 <> "2" Then
      If Text3 = MsgText(601) Then
         MsgBox "請輸入智權人員！", vbExclamation
         Text3.SetFocus
         Exit Sub
      End If
   End If
   '2014/6/4 END
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   Else
      If MaskEdBox1.Text = MsgText(29) Then
         MsgBox "請輸入帳款截止日！", vbExclamation
         MaskEdBox1.SetFocus
         Exit Sub
      End If
   End If
   Screen.MousePointer = vbHourglass
   Accrpt108Delete
   ProduceData
   If Text5 = "2" And Trim(Text3) = "" Then '依各區做合計
      PrintExcel2
   Else
      PrintExcel
   End If
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5655
   'Modify by Amy 2023/10/11 原3348
   Me.Height = 3570
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
   'Set dllaccrpt108 = CreateObject("AccReport.ReportSelect") 'Remove by Lydia 2016/12/09
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Set dllaccrpt108 = Nothing 'Remove by Lydia 2016/12/09
   Set Frmacc1480 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_Change()
   If Len(Text3) = 5 Then
      lblSalesName = StaffQuery(Text3)
   Else
      lblSalesName = MsgText(601)
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub
'Add By Sindy 2014/6/3
Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
   End If
End Sub
'2014/6/3 END

'Add By Sindy 2014/6/3
Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub
'2014/6/3 END

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim douCal1, douCal2, douCal3, douCal4 As Double
Dim LngDays As Long, douAmount As Double
Dim strSql As String
Dim strCPSql As String 'Add By Sindy 2014/6/4
   
On Error GoTo Checking
   'Add By Sindy 2014/6/4
   strSql = ""
   strCPSql = ""
   If Text4 = "1" Then '已發文
      strCPSql = strCPSql & " and cp27>0"
   ElseIf Text4 = "2" Then '未發文
      strCPSql = strCPSql & " and nvl(cp27,0)=0"
   End If
   '2014/6/4 END
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt108.CursorLocation = adUseClient
   'Modified by Lydia 2016/12/09 + where r10801='" & strUserNum & "'
   adoaccrpt108.Open "select * from accrpt108 where r10801='" & strUserNum & "' order by r10801 asc, r10802 asc, r10809 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text5
      Case "1"
         adoacc0k0.CursorLocation = adUseClient
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and a0k03 >= '" & Text1 & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and a0k03 <= '" & Text2 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         'Modify By Sindy 2014/6/4
         'strExc(0) = "select * from acc0k0 where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & " and (a0k09 is null or a0k09=0) order by a0k03 asc"
         '2015/7/9 MODIFY BY SONIA 要扣除銷帳97011張祐昇
         'strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress" & _
                     " where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k03 asc"
         strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress," & _
                     "(select a1u02,sum(a1u08+a1u10-a1u07-a1u09) a1u08 from acc0k0,acc1u0" & _
                     " Where (a0k06 + a0k07) > (nvl(a0k17, 0) + nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0) and a0k01=a1u02(+) group by a1u02) " & _
                     " where (a0k06+a0k07)+nvl(a1u08,0) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+) and a0k01=a1u02(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k03 asc"
         '2014/6/4 END
         adoacc0k0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
         If adoacc0k0.RecordCount = 0 Then
            adoacc0k0.Close
            adoaccrpt108.Close
            MsgBox MsgText(28), , MsgText(5)
            Exit Sub
         End If
         Do While adoacc0k0.EOF = False
            If adoaccrpt108.RecordCount <> 0 Then
               adoaccrpt108.MoveFirst
               adoaccrpt108.Find "r10801 = '" & strUserNum & "'", 0, adSearchForward, 1
               If adoaccrpt108.EOF Then
                  adoaccrpt108.AddNew
                  CustomerSave
               Else
                  adoaccrpt108.Find "r10802 = '" & adoacc0k0.Fields("a0k03").Value & "'", 0, adSearchForward, adoaccrpt108.Bookmark
                  If adoaccrpt108.EOF Then
                     adoaccrpt108.AddNew
                     CustomerSave
                  End If
               End If
            Else
               adoaccrpt108.AddNew
               CustomerSave
            End If
            Calculate
            adoaccrpt108.UpdateBatch
            adoacc0k0.MoveNext
         Loop
         adoacc0k0.Close
      Case "2"
         adoacc0k0.CursorLocation = adUseClient
         'Modify by Morgan 2007/10/2 智權人員範圍改成一個
         'If Text3 <> MsgText(601) Then
         '   strSQL = " and a0k20 >= '" & Text3 & "'"
         'End If
         'If Text4 <> MsgText(601) Then
         '   strSQL = strSQL & " and a0k20 <= '" & Text4 & "'"
         'End If
         If Text3 <> MsgText(601) Then
            strSql = strSql & " and a0k20 = '" & Text3 & "'"
         End If
         'end 2007/10/2
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         'Modify By Sindy 2014/6/4
         'strExc(0) = "select * from acc0k0 where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & " and (a0k09 is null or a0k09=0) order by a0k20 asc"
         '2015/7/9 MODIFY BY SONIA 要扣除銷帳97011張祐昇
         'strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress" & _
                     " where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k20 asc"
         strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress," & _
                     "(select a1u02,sum(a1u08+a1u10-a1u07-a1u09) a1u08 from acc0k0,acc1u0" & _
                     " Where (a0k06 + a0k07) > (nvl(a0k17, 0) + nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0) and a0k01=a1u02(+) group by a1u02) " & _
                     " where (a0k06+a0k07)+nvl(a1u08,0) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+) and a0k01=a1u02(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k20 asc"
         '2014/6/4 END
         adoacc0k0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
         If adoacc0k0.RecordCount = 0 Then
            adoacc0k0.Close
            adoaccrpt108.Close
            MsgBox MsgText(28), , MsgText(5)
            Exit Sub
         End If
         Do While adoacc0k0.EOF = False
            If adoaccrpt108.RecordCount <> 0 Then
               adoaccrpt108.MoveFirst
               adoaccrpt108.Find "r10801 = '" & strUserNum & "'", 0, adSearchForward, 1
               If adoaccrpt108.EOF Then
                  adoaccrpt108.AddNew
                  StaffSave
               Else
                  adoaccrpt108.Find "r10802 = '" & IIf(IsNull(adoacc0k0.Fields("a0k20").Value), "1", adoacc0k0.Fields("a0k20").Value) & "'", 0, adSearchForward, adoaccrpt108.Bookmark
                  If adoaccrpt108.EOF Then
                     adoaccrpt108.AddNew
                     StaffSave
                  End If
               End If
            Else
               adoaccrpt108.AddNew
               StaffSave
            End If
            Calculate
            adoaccrpt108.UpdateBatch
            adoacc0k0.MoveNext
         Loop
         adoacc0k0.Close
      Case Else
         adoacc0k0.CursorLocation = adUseClient
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and a0k03 >= '" & Text1 & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and a0k03 <= '" & Text2 & "'"
         End If
         'Modify by Morgan 2007/10/2 智權人員範圍改成一個
         'If Text3 <> MsgText(601) Then
         '   strSQL = strSQL & " and a0k20 >= '" & Text3 & "'"
         'End If
         'If Text4 <> MsgText(601) Then
         '   strSQL = strSQL & " and a0k20 <= '" & Text4 & "'"
         'End If
         If Text3 <> MsgText(601) Then
            strSql = strSql & " and a0k20 = '" & Text3 & "'"
         End If
         'end 2007/10/2
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         'Modify By Sindy 2014/6/4
         'strExc(0) = "select * from acc0k0 where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & " and (a0k09 is null or a0k09 = 0) order by a0k20 asc"
         '2015/7/9 MODIFY BY SONIA 要扣除銷帳97011張祐昇
         'strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress" & _
                     " where (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k20 asc"
         strExc(0) = "select distinct a0k01,a0k02,a0k03,a0k06,a0k07,a0k17,a0k18,a0k20" & _
                     " from acc0k0,acc0j0,caseprogress," & _
                     "(select a1u02,sum(a1u08+a1u10-a1u07-a1u09) a1u08 from acc0k0,acc1u0" & _
                     " Where (a0k06 + a0k07) > (nvl(a0k17, 0) + nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0) and a0k01=a1u02(+) group by a1u02) " & _
                     " where (a0k06+a0k07)+nvl(a1u08,0) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & strSql & _
                     " and (a0k09 is null or a0k09=0)" & _
                     " and a0k01=a0j13(+) and a0k01=a1u02(+)" & _
                     " and a0j01=cp09(+)" & strCPSql & _
                     " order by a0k20 asc"
         '2014/6/4 END
         adoacc0k0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
         If adoacc0k0.RecordCount = 0 Then
            adoacc0k0.Close
            MsgBox MsgText(28), , MsgText(5)
            Exit Sub
         End If
         Do While adoacc0k0.EOF = False
            If adoaccrpt108.RecordCount <> 0 Then
               adoaccrpt108.MoveFirst
               adoaccrpt108.Find "r10801 = '" & strUserNum & "'", 0, adSearchForward, 1
               If adoaccrpt108.EOF Then
                  adoaccrpt108.AddNew
                  CustomerSave
               Else
                  adoaccrpt108.Find "r10802 = '" & adoacc0k0.Fields("a0k03").Value & "'", 0, adSearchForward, adoaccrpt108.Bookmark
                  If adoaccrpt108.EOF Then
                     adoaccrpt108.AddNew
                     CustomerSave
                  Else
                     adoaccrpt108.Find "r10809 = '" & adoacc0k0.Fields("a0k20").Value & "'", 0, adSearchForward, adoaccrpt108.Bookmark
                     If adoaccrpt108.EOF Then
                        adoaccrpt108.AddNew
                        CustomerSave
                     End If
                  End If
               End If
            Else
               adoaccrpt108.AddNew
               CustomerSave
            End If
            Calculate
            adoaccrpt108.UpdateBatch
            adoacc0k0.MoveNext
         Loop
         adoacc0k0.Close
   End Select
   adoaccrpt108.Close
   'Modify by Morgan 2005/2/3 應收帳款>0才要
   'adoTaie.Execute "delete from accrpt108 where r10802 is null"
   'Modified by Lydia 2016/12/09 + and r10801=" & CNULL(strUserNum)
   adoTaie.Execute "delete from accrpt108 where r10802 is null or r10808=0 and r10801=" & CNULL(strUserNum)
   '2005/2/3 end
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt108Delete()
   'Modified by Lydia 2016/12/09 + where r10801=" & CNULL(strUserNum)
   adoTaie.Execute "delete from accrpt108 where r10801=" & CNULL(strUserNum)
End Sub

'*************************************************
'  客戶資料存檔
'
'*************************************************
Private Sub CustomerSave()
   adoaccrpt108.Fields("r10801").Value = strUserNum
   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
      adoaccrpt108.Fields("r10802").Value = Null
   Else
      adoaccrpt108.Fields("r10802").Value = adoacc0k0.Fields("a0k03").Value
      adoaccrpt108.Fields("r10803").Value = CustomerQuery(adoacc0k0.Fields("a0k03").Value, 1)
   End If
   adoaccrpt108.Fields("r10804").Value = 0
   adoaccrpt108.Fields("r10805").Value = 0
   adoaccrpt108.Fields("r10806").Value = 0
   adoaccrpt108.Fields("r10807").Value = 0
   adoaccrpt108.Fields("r10810").Value = 0 'Added by Lydia 2016/12/09
   
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt108.Fields("r10809").Value = ""
   Else
      Select Case Text5
         Case "3"
            adoaccrpt108.Fields("r10809").Value = adoacc0k0.Fields("a0k20").Value
         Case Else
'            adoaccrpt108.Fields("r10809").Value = ""
      End Select
   End If
End Sub

'*************************************************
'  智權人員資料存檔
'
'*************************************************
Private Sub StaffSave()
   adoaccrpt108.Fields("r10801").Value = strUserNum
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
      adoaccrpt108.Fields("r10802").Value = Null
   Else
      adoaccrpt108.Fields("r10802").Value = adoacc0k0.Fields("a0k20").Value
      adoaccrpt108.Fields("r10803").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
   End If
   adoaccrpt108.Fields("r10804").Value = 0
   adoaccrpt108.Fields("r10805").Value = 0
   adoaccrpt108.Fields("r10806").Value = 0
   adoaccrpt108.Fields("r10807").Value = 0
'   adoaccrpt108.Fields("r10809").Value = ""
   adoaccrpt108.Fields("r10810").Value = 0 'Added by Lydia 2016/12/09
End Sub

'*************************************************
'  計算應收帳款
'
'*************************************************
Private Sub Calculate()
Dim douCAmount As Double
'add by nickc 2007/02/08
Dim douAmount As Double, LngDays As Integer
   
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'E09313961
   adoquery.Open "select sum(a1u08+a1u10-a1u07-a1u09) as CAmount from acc1u0 where a1u02 = '" & adoacc0k0.Fields("a0k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("CAmount").Value) Then
         douCAmount = 0
      Else
         douCAmount = adoquery.Fields("CAmount").Value
      End If
   Else
      douCAmount = 0
   End If
   adoquery.Close
   douAmount = Val(adoacc0k0.Fields("a0k06").Value) + Val(adoacc0k0.Fields("a0k07").Value) - Val(IIf(IsNull(adoacc0k0.Fields("a0k17").Value), 0, adoacc0k0.Fields("a0k17").Value)) - Val(IIf(IsNull(adoacc0k0.Fields("a0k18").Value), 0, adoacc0k0.Fields("a0k18").Value))
   'E09313961
   douAmount = douAmount + douCAmount
   LngDays = CalculateDays(CADate(adoacc0k0.Fields("a0k02").Value), CADate(FCDate(MaskEdBox1.Text)))
   If LngDays <= 30 Then
      adoaccrpt108.Fields("r10804").Value = Val(adoaccrpt108.Fields("r10804").Value) + douAmount
   Else
      If LngDays >= 31 And LngDays <= 60 Then
         adoaccrpt108.Fields("r10805").Value = Val(adoaccrpt108.Fields("r10805").Value) + douAmount
      Else
         If LngDays >= 61 And LngDays <= 90 Then
            adoaccrpt108.Fields("r10806").Value = Val(adoaccrpt108.Fields("r10806").Value) + douAmount
         'Modified by Lydia 2016/12/09 增加1年以上應收款
         'Else
         ElseIf LngDays >= 91 And LngDays <= 365 Then
            adoaccrpt108.Fields("r10807").Value = Val(adoaccrpt108.Fields("r10807").Value) + douAmount
         'Added by Lydia 2016/12/09
         Else
            adoaccrpt108.Fields("r10810").Value = Val(adoaccrpt108.Fields("r10810").Value) + douAmount
         End If
      End If
   End If
   'Modified by Lydia 2016/12/09 + Val(adoaccrpt108.Fields("r10810").Value)
   adoaccrpt108.Fields("r10808").Value = Val(adoaccrpt108.Fields("r10804").Value) + Val(adoaccrpt108.Fields("r10805").Value) + Val(adoaccrpt108.Fields("r10806").Value) + Val(adoaccrpt108.Fields("r10807").Value) + Val(adoaccrpt108.Fields("r10810").Value)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = "X"
   Text2 = "X"
   Text3 = ""
   Text4 = "" 'Add By Sindy 2014/6/4
   lblSalesName = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text5 = ""
   Text1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'Modify by Morgan 2007/10/2 智權人員範圍改成一個
   'If Text4 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   'end 2007/10/2
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add By Sindy 2014/6/4
'*************************************************
' 產生Excel資料
'
'*************************************************
Public Sub PrintExcel()
Dim strFilePath As String
Dim dblSkipPageRow As Double
Dim strTemp As String
Dim dbl_30Tot As Double
Dim dbl_60Tot As Double
Dim dbl_90Tot As Double
Dim dbl_91Tot As Double
Dim dbl_Tot As Double
'Added by Lydia 2016/12/09
Dim strEnd As String
Dim dbl_Year As Double

On Error GoTo ErrHnd
   
   If adoaccrpt108.State = adStateOpen Then
      adoaccrpt108.Close
   End If
   adoaccrpt108.CursorLocation = adUseClient
   'Modified by Lydia 2016/12/09 +where r10801=" & CNULL(strUserNum)
   adoaccrpt108.Open "select * from accrpt108 where r10801=" & CNULL(strUserNum), adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt108.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccrpt108.Close
      Exit Sub
   End If
   
   intPage = 0
   intCounter = 0
   dblSkipPageRow = 0
   dbl_30Tot = 0
   dbl_60Tot = 0
   dbl_90Tot = 0
   dbl_91Tot = 0
   dbl_Year = 0 'Added by Lydia 2016/12/09
   dbl_Tot = 0
      
   'Excel檔案路徑
   strFilePath = strExcelPath & IIf(Text5 = "1", "客戶帳齡分析表", IIf(Text5 = "2", "智權人員帳齡分析表", "智權人員客戶帳齡分析表")) & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   Set xlsAnnuity = New Excel.Application
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   With wksAnnuity
      .PageSetup.Orientation = xlPortrait '橫印xlLandscape,直印xlPortrait
      .PageSetup.LeftMargin = 28.34
      .PageSetup.RightMargin = 28.34
      .PageSetup.TopMargin = 42.51
      .PageSetup.BottomMargin = 42.51
      .PageSetup.HeaderMargin = 28.34
      .PageSetup.FooterMargin = 28.34
      '設定各欄位長度
      .Columns("A:A").ColumnWidth = 10 '客戶編號/智權人員
      .Columns("B:B").ColumnWidth = 17 '名稱
      .Columns("C:C").ColumnWidth = 12 '0~30天
      .Columns("D:D").ColumnWidth = 12  '31~60天
      .Columns("E:E").ColumnWidth = 12 '61~90天
      .Columns("F:F").ColumnWidth = 12 '90天以上 'Memo by Lydia 2016/12/09 改成90天~1年
      .Columns("G:G").ColumnWidth = 12 '應收帳款 'Memo by Lydia 2016/12/09 改成1年以上
      .Columns("H:H").ColumnWidth = 12 'Added by Lydia 2016/12/09 應收帳款
      strEnd = "H"  'Added by Lydia 2016/12/09
      '逐筆填值
      adoaccrpt108.MoveFirst
      Do While adoaccrpt108.EOF = False
         intCounter = intCounter + 1
         '一開始或資料已填滿一頁時跳頁
         If intCounter = 1 Or dblSkipPageRow >= 41 Then
            If dblSkipPageRow >= 41 Then
               '換頁
               .Range("A" & intCounter).Select
               .HPageBreaks.add Before:=.Application.ActiveCell
            End If
            dblSkipPageRow = 0
            Call PrintExcelTitle
         End If
         If IsNull(adoaccrpt108.Fields("R10802").Value) = False Then
            .Range("A" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "@"
            .Range("A" & intCounter).Value = adoaccrpt108.Fields("R10802").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10803").Value) = False Then
            .Range("B" & intCounter).Value = adoaccrpt108.Fields("R10803").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10804").Value) = False Then
            .Range("C" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("C" & intCounter).Value = CStr(adoaccrpt108.Fields("R10804").Value)
            dbl_30Tot = dbl_30Tot + adoaccrpt108.Fields("R10804").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10805").Value) = False Then
            .Range("D" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("D" & intCounter).Value = CStr(adoaccrpt108.Fields("R10805").Value)
            dbl_60Tot = dbl_60Tot + adoaccrpt108.Fields("R10805").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10806").Value) = False Then
            .Range("E" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("E" & intCounter).Value = CStr(adoaccrpt108.Fields("R10806").Value)
            dbl_90Tot = dbl_90Tot + adoaccrpt108.Fields("R10806").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10807").Value) = False Then
            .Range("F" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("F" & intCounter).Value = CStr(adoaccrpt108.Fields("R10807").Value)
            dbl_91Tot = dbl_91Tot + adoaccrpt108.Fields("R10807").Value
         End If
         'Added by Lydia 2016/12/09 增加1年以上的欄位
         If IsNull(adoaccrpt108.Fields("R10810").Value) = False Then
            .Range("G" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("G" & intCounter).Value = CStr(adoaccrpt108.Fields("R10810").Value)
            dbl_Year = dbl_Year + adoaccrpt108.Fields("R10810").Value
         End If
         'end 2016/12/09
         
         If IsNull(adoaccrpt108.Fields("R10808").Value) = False Then
            'Modified by Lydia 2016/12/09 應收帳款總額
            '.Range("G" & intCounter).Select
            .Range(strEnd & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            'Modified by Lydia 2016/12/09
            .Range(strEnd & intCounter).Value = CStr(adoaccrpt108.Fields("R10808").Value)
            dbl_Tot = dbl_Tot + adoaccrpt108.Fields("R10808").Value
         End If
         dblSkipPageRow = dblSkipPageRow + 1
         adoaccrpt108.MoveNext
      Loop
      
      '單線
      'Modified by Lydia 2016/12/09
      'strTemp = "C" & intCounter & ":G" & intCounter
      strTemp = "C" & intCounter & ":" & strEnd & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      '合計
      intCounter = intCounter + 1
      .Range("B" & intCounter).Value = "合計"
      .Range("C" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("C" & intCounter).Value = dbl_30Tot
      .Range("D" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("D" & intCounter).Value = dbl_60Tot
      .Range("E" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("E" & intCounter).Value = dbl_90Tot
      .Range("F" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("F" & intCounter).Value = dbl_91Tot
      
      'Modified by Lydia 2016/12/09
'      .Range("G" & intCounter).Select
'      .Application.Selection.NumberFormatLocal = "#,##0_ "
'      .Range("G" & intCounter).Value = dbl_Tot
      '1年以上
      .Range("G" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("G" & intCounter).Value = dbl_Year
      '應收帳款總額
      .Range(strEnd & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range(strEnd & intCounter).Value = dbl_Tot
      'end 2016/12/09
      
      '雙線
      'Modified by Lydia 2016/12/09
      'strTemp = "C" & intCounter & ":G" & intCounter
      strTemp = "C" & intCounter & ":" & strEnd & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlDouble
         .Weight = xlThick
         .ColorIndex = xlAutomatic
      End With
      intCounter = intCounter + 1
      .Range("D" & intCounter).Value = "***結束***"
   End With
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt108.Close
   MsgBox "檔案已產生！" & vbCrLf & vbCrLf & "存放至 " & strFilePath
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt108.Close
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2014/6/4 依各區統計
Public Sub PrintExcel2()
Dim strFilePath As String
Dim dblSkipPageRow As Double
Dim strTemp As String
Dim strST15 As String, strA0902 As String
'各區小計
Dim dbl_sub30Tot As Double
Dim dbl_sub60Tot As Double
Dim dbl_sub90Tot As Double
Dim dbl_sub91Tot As Double
Dim dbl_subTot As Double
'台北,台中所小計
Dim dbl_S1030Tot As Double
Dim dbl_S1060Tot As Double
Dim dbl_S1090Tot As Double
Dim dbl_S1091Tot As Double
Dim dbl_S10Tot As Double
'合計
Dim dbl_30Tot As Double
Dim dbl_60Tot As Double
Dim dbl_90Tot As Double
Dim dbl_91Tot As Double
Dim dbl_Tot As Double
'Added by Lydia 2016/12/09
Dim strEnd As String
Dim dbl_subYear As Double
Dim dbl_S10Year As Double
Dim dbl_Year As Double

On Error GoTo ErrHnd
   
   If adoaccrpt108.State = adStateOpen Then
      adoaccrpt108.Close
   End If
   adoaccrpt108.CursorLocation = adUseClient
   'Modified by Lydia 2016/12/09
   'adoaccrpt108.Open "select st15,a0902,accrpt108.* from accrpt108,staff,acc090 where R10802=st01(+)" & _
                     " and st15=a0901(+) order by st15,st01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoaccrpt108.Open "select decode(substr(st15,1,1),'S','1','2') ord1,st15,a0902,accrpt108.* from accrpt108,staff,acc090 " & _
                     "where R10802=st01(+) and r10801=" & CNULL(strUserNum) & " and st15=a0901(+) order by ord1,st15,st01 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt108.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccrpt108.Close
      Exit Sub
   End If
   
   intPage = 0
   intCounter = 0
   dblSkipPageRow = 0
   strST15 = ""
   '各區小計
   dbl_sub30Tot = 0
   dbl_sub60Tot = 0
   dbl_sub90Tot = 0
   dbl_sub91Tot = 0
   dbl_subTot = 0
   '台北,台中所小計
   dbl_S1030Tot = 0
   dbl_S1060Tot = 0
   dbl_S1090Tot = 0
   dbl_S1091Tot = 0
   dbl_S10Tot = 0
   '合計
   dbl_30Tot = 0
   dbl_60Tot = 0
   dbl_90Tot = 0
   dbl_91Tot = 0
   dbl_Tot = 0
   'Added by Lydia 2016/12/09
   dbl_subYear = 0 '各區小計
   dbl_S10Year = 0 '台北,台中所小計
   dbl_Year = 0    '合計
   
   'Excel檔案路徑
   strFilePath = strExcelPath & IIf(Text5 = "1", "客戶帳齡分析表", IIf(Text5 = "2", "智權人員帳齡分析表", "智權人員客戶帳齡分析表")) & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   Set xlsAnnuity = New Excel.Application
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   With wksAnnuity
      .PageSetup.Orientation = xlPortrait '橫印xlLandscape,直印xlPortrait
      .PageSetup.LeftMargin = 28.34
      .PageSetup.RightMargin = 28.34
      .PageSetup.TopMargin = 42.51
      .PageSetup.BottomMargin = 42.51
      .PageSetup.HeaderMargin = 28.34
      .PageSetup.FooterMargin = 28.34
      '設定各欄位長度
      .Columns("A:A").ColumnWidth = 10 '客戶編號/智權人員
      .Columns("B:B").ColumnWidth = 17 '名稱
      .Columns("C:C").ColumnWidth = 12 '0~30天
      .Columns("D:D").ColumnWidth = 12  '31~60天
      .Columns("E:E").ColumnWidth = 12 '61~90天
      .Columns("F:F").ColumnWidth = 12 '90天以上 'Memo by Lydia 2016/12/09 改成90天~1年
      .Columns("G:G").ColumnWidth = 12 '應收帳款 'Memo by Lydia 2016/12/09 改成1年以上
      .Columns("H:H").ColumnWidth = 12 'Added by Lydia 2016/12/09 應收帳款
      strEnd = "H"  'Added by Lydia 2016/12/09
      '逐筆填值
      adoaccrpt108.MoveFirst
      Do While adoaccrpt108.EOF = False
         If strST15 <> "" And strST15 <> adoaccrpt108.Fields("ST15").Value Then
            '智權部的人員才需要小計
            If Left(strST15, 1) = "S" And Left(strST15, 2) >= "S1" Then
               '單線
   '            strTemp = "C" & intCounter & ":G" & intCounter
   '            .Range(strTemp).Select
   '            With .Application.Selection.Borders(xlEdgeBottom)
   '               .LineStyle = xlContinuous
   '               .Weight = xlThin
   '               .ColorIndex = xlAutomatic
   '            End With
               intCounter = intCounter + 1
               '已填滿一頁時跳頁
               If dblSkipPageRow >= 41 Then
                  '換頁
                  .Range("A" & intCounter).Select
                  .HPageBreaks.add Before:=.Application.ActiveCell
                  dblSkipPageRow = 0
                  Call PrintExcelTitle
               End If
               '小計
               dblSkipPageRow = dblSkipPageRow + 1
               .Range("A" & intCounter).Value = strA0902
               .Range("B" & intCounter).Value = "小計"
               .Range("C" & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range("C" & intCounter).Value = dbl_sub30Tot
               .Range("D" & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range("D" & intCounter).Value = dbl_sub60Tot
               .Range("E" & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range("E" & intCounter).Value = dbl_sub90Tot
               .Range("F" & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range("F" & intCounter).Value = dbl_sub91Tot
               'Modified by Lydia 2016/12/09
               '.Range("G" & intCounter).Select
               '.Application.Selection.NumberFormatLocal = "#,##0_ "
               '.Range("G" & intCounter).Value = dbl_subTot
               '1年以上
               .Range("G" & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range("G" & intCounter).Value = dbl_subYear
               '應收帳款總額
               .Range(strEnd & intCounter).Select
               .Application.Selection.NumberFormatLocal = "#,##0_ "
               .Range(strEnd & intCounter).Value = dbl_subTot
               'end 2016/12/09
               
               '單線
               'Modified by Lydia 2016/12/09
               'strTemp = "C" & intCounter & ":G" & intCounter
               strTemp = "C" & intCounter & ":" & strEnd & intCounter
               .Range(strTemp).Select
               With .Application.Selection.Borders(xlEdgeBottom)
                  .LineStyle = xlContinuous
                  .Weight = xlThin
                  .ColorIndex = xlAutomatic
               End With
               dbl_sub30Tot = 0
               dbl_sub60Tot = 0
               dbl_sub90Tot = 0
               dbl_sub91Tot = 0
               dbl_subTot = 0
               dbl_subYear = 0 'Added by Lydia 2016/12/09
               '台北所及台中所要再小計一次
               If Left(strST15, 2) <> Left(adoaccrpt108.Fields("ST15").Value, 2) Then
                  If Left(strST15, 2) = "S1" Or Left(strST15, 2) = "S2" Then
                     intCounter = intCounter + 1
                     '已填滿一頁時跳頁
                     If dblSkipPageRow >= 41 Then
                        '換頁
                        .Range("A" & intCounter).Select
                        .HPageBreaks.add Before:=.Application.ActiveCell
                        dblSkipPageRow = 0
                        Call PrintExcelTitle
                     End If
                     '小計
                     dblSkipPageRow = dblSkipPageRow + 1
                     .Range("A" & intCounter).Value = IIf(Left(strST15, 2) = "S1", "台北所", "台中所")
                     .Range("B" & intCounter).Value = "小計"
                     .Range("C" & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range("C" & intCounter).Value = dbl_S1030Tot
                     .Range("D" & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range("D" & intCounter).Value = dbl_S1060Tot
                     .Range("E" & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range("E" & intCounter).Value = dbl_S1090Tot
                     .Range("F" & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range("F" & intCounter).Value = dbl_S1091Tot
                     'Modified by Lydia 2016/12/09
'                     .Range("G" & intCounter).Select
'                     .Application.Selection.NumberFormatLocal = "#,##0_ "
'                     .Range("G" & intCounter).Value = dbl_S10Tot
                     '1年以上
                     .Range("G" & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range("G" & intCounter).Value = dbl_S10Year
                     '應收帳款總額
                     .Range(strEnd & intCounter).Select
                     .Application.Selection.NumberFormatLocal = "#,##0_ "
                     .Range(strEnd & intCounter).Value = dbl_S10Tot
                     'end 2016/12/09
                     
                     '單線
                     'Modified by Lydia 2016/12/09
                     'strTemp = "C" & intCounter & ":G" & intCounter
                     strTemp = "C" & intCounter & ":" & strEnd & intCounter
                     .Range(strTemp).Select
                     With .Application.Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                     End With
                     dbl_S1030Tot = 0
                     dbl_S1060Tot = 0
                     dbl_S1090Tot = 0
                     dbl_S1091Tot = 0
                     dbl_S10Tot = 0
                     dbl_S10Year = 0 'Added by Lydia 2016/12/09
                  End If
               End If
            Else
               '再Run下一筆是S1X資料前,先小計一次
               If Left(adoaccrpt108.Fields("ST15").Value, 2) = "S1" Then
                  intCounter = intCounter + 1
                  '已填滿一頁時跳頁
                  If dblSkipPageRow >= 41 Then
                     '換頁
                     .Range("A" & intCounter).Select
                     .HPageBreaks.add Before:=.Application.ActiveCell
                     dblSkipPageRow = 0
                     Call PrintExcelTitle
                  End If
                  '小計
                  dblSkipPageRow = dblSkipPageRow + 1
                  .Range("A" & intCounter).Value = "其他"
                  .Range("B" & intCounter).Value = "小計"
                  .Range("C" & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range("C" & intCounter).Value = dbl_sub30Tot
                  .Range("D" & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range("D" & intCounter).Value = dbl_sub60Tot
                  .Range("E" & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range("E" & intCounter).Value = dbl_sub90Tot
                  .Range("F" & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range("F" & intCounter).Value = dbl_sub91Tot
                  'Modified by Lydia 2016/12/09
                  '.Range("G" & intCounter).Select
                  '.Application.Selection.NumberFormatLocal = "#,##0_ "
                  '.Range("G" & intCounter).Value = dbl_subTot
                  '1年以上
                  .Range("G" & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range("G" & intCounter).Value = dbl_subYear
                  '應收帳款總額
                  .Range(strEnd & intCounter).Select
                  .Application.Selection.NumberFormatLocal = "#,##0_ "
                  .Range(strEnd & intCounter).Value = dbl_subTot
                  'end 2016/12/09
               
                  '單線
                  'Modified by Lydia 2016/12/09
                  'strTemp = "C" & intCounter & ":G" & intCounter
                  strTemp = "C" & intCounter & ":" & strEnd & intCounter
                  .Range(strTemp).Select
                  With .Application.Selection.Borders(xlEdgeBottom)
                     .LineStyle = xlContinuous
                     .Weight = xlThin
                     .ColorIndex = xlAutomatic
                  End With
                  dbl_sub30Tot = 0
                  dbl_sub60Tot = 0
                  dbl_sub90Tot = 0
                  dbl_sub91Tot = 0
                  dbl_subTot = 0
                  dbl_subYear = 0 'Added by Lydia 2016/12/09
               End If
            End If
         End If
         intCounter = intCounter + 1
         '一開始或資料已填滿一頁時跳頁
         If intCounter = 1 Or dblSkipPageRow >= 41 Then
            If dblSkipPageRow >= 41 Then
               '換頁
               .Range("A" & intCounter).Select
               .HPageBreaks.add Before:=.Application.ActiveCell
            End If
            dblSkipPageRow = 0
            Call PrintExcelTitle
         End If
         If IsNull(adoaccrpt108.Fields("R10802").Value) = False Then
            .Range("A" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "@"
            .Range("A" & intCounter).Value = adoaccrpt108.Fields("R10802").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10803").Value) = False Then
            .Range("B" & intCounter).Value = adoaccrpt108.Fields("R10803").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10804").Value) = False Then
            .Range("C" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("C" & intCounter).Value = CStr(adoaccrpt108.Fields("R10804").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S1030Tot = dbl_S1030Tot + adoaccrpt108.Fields("R10804").Value
            End If
            dbl_30Tot = dbl_30Tot + adoaccrpt108.Fields("R10804").Value
            dbl_sub30Tot = dbl_sub30Tot + adoaccrpt108.Fields("R10804").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10805").Value) = False Then
            .Range("D" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("D" & intCounter).Value = CStr(adoaccrpt108.Fields("R10805").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S1060Tot = dbl_S1060Tot + adoaccrpt108.Fields("R10805").Value
            End If
            dbl_60Tot = dbl_60Tot + adoaccrpt108.Fields("R10805").Value
            dbl_sub60Tot = dbl_sub60Tot + adoaccrpt108.Fields("R10805").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10806").Value) = False Then
            .Range("E" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("E" & intCounter).Value = CStr(adoaccrpt108.Fields("R10806").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S1090Tot = dbl_S1090Tot + adoaccrpt108.Fields("R10806").Value
            End If
            dbl_90Tot = dbl_90Tot + adoaccrpt108.Fields("R10806").Value
            dbl_sub90Tot = dbl_sub90Tot + adoaccrpt108.Fields("R10806").Value
         End If
         If IsNull(adoaccrpt108.Fields("R10807").Value) = False Then
            .Range("F" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("F" & intCounter).Value = CStr(adoaccrpt108.Fields("R10807").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S1091Tot = dbl_S1091Tot + adoaccrpt108.Fields("R10807").Value
            End If
            dbl_91Tot = dbl_91Tot + adoaccrpt108.Fields("R10807").Value
            dbl_sub91Tot = dbl_sub91Tot + adoaccrpt108.Fields("R10807").Value
         End If
         'Added by Lydia 2016/12/09 1年以上
         If IsNull(adoaccrpt108.Fields("R10810").Value) = False Then
            .Range("G" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("G" & intCounter).Value = CStr(adoaccrpt108.Fields("R10810").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S10Year = dbl_S10Year + adoaccrpt108.Fields("R10810").Value
            End If
            dbl_Year = dbl_Year + adoaccrpt108.Fields("R10810").Value
            dbl_subYear = dbl_subYear + adoaccrpt108.Fields("R10810").Value
         End If
         'end 2016/12/09
         
         If IsNull(adoaccrpt108.Fields("R10808").Value) = False Then
            'Modified by Lydia 2016/12/09
            '.Range("G" & intCounter).Select
            .Range(strEnd & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            'Modified by Lydia 2016/12/09
            '.Range("G" & intCounter).Value = CStr(adoaccrpt108.Fields("R10808").Value)
            .Range(strEnd & intCounter).Value = CStr(adoaccrpt108.Fields("R10808").Value)
            If Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S1" Or _
               Left(Trim(adoaccrpt108.Fields("ST15").Value), 2) = "S2" Then
               dbl_S10Tot = dbl_S10Tot + adoaccrpt108.Fields("R10808").Value
            End If
            dbl_Tot = dbl_Tot + adoaccrpt108.Fields("R10808").Value
            dbl_subTot = dbl_subTot + adoaccrpt108.Fields("R10808").Value
         End If
         strST15 = adoaccrpt108.Fields("ST15").Value
         'Added by Lydia 2019/10/08 非智權人員以"其他"部門小計
         If Left(adoaccrpt108.Fields("ST15").Value, 1) <> "S" Then
             strA0902 = "其他"
         Else
         'end 2019/10/08
             strA0902 = adoaccrpt108.Fields("A0902").Value
         End If 'end 2019/10/08
         
         dblSkipPageRow = dblSkipPageRow + 1
         adoaccrpt108.MoveNext
      Loop
      '單線
'      strTemp = "C" & intCounter & ":G" & intCounter
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      intCounter = intCounter + 1
      '已填滿一頁時跳頁
      If dblSkipPageRow >= 41 Then
         '換頁
         .Range("A" & intCounter).Select
         .HPageBreaks.add Before:=.Application.ActiveCell
         dblSkipPageRow = 0
         Call PrintExcelTitle
      End If
      '小計
      dblSkipPageRow = dblSkipPageRow + 1
      .Range("A" & intCounter).Value = strA0902
      .Range("B" & intCounter).Value = "小計"
      .Range("C" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("C" & intCounter).Value = dbl_sub30Tot
      .Range("D" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("D" & intCounter).Value = dbl_sub60Tot
      .Range("E" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("E" & intCounter).Value = dbl_sub90Tot
      .Range("F" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("F" & intCounter).Value = dbl_sub91Tot
      'Modified by Lydia 2016/12/09
'      .Range("G" & intCounter).Select
'      .Application.Selection.NumberFormatLocal = "#,##0_ "
'      .Range("G" & intCounter).Value = dbl_subTot
      '1年以上
      .Range("G" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("G" & intCounter).Value = dbl_subYear
      '應收帳款總額
      .Range(strEnd & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range(strEnd & intCounter).Value = dbl_subTot
      'end 2016/12/09
      
      '單線
      'Modified by Lydia 2016/12/09
      'strTemp = "C" & intCounter & ":G" & intCounter
      strTemp = "C" & intCounter & ":" & strEnd & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      intCounter = intCounter + 1
      '已填滿一頁時跳頁
      If dblSkipPageRow >= 41 Then
         '換頁
         .Range("A" & intCounter).Select
         .HPageBreaks.add Before:=.Application.ActiveCell
         dblSkipPageRow = 0
         Call PrintExcelTitle
      End If
      '合計
      .Range("B" & intCounter).Value = "合計"
      .Range("C" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("C" & intCounter).Value = dbl_30Tot
      .Range("D" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("D" & intCounter).Value = dbl_60Tot
      .Range("E" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("E" & intCounter).Value = dbl_90Tot
      .Range("F" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("F" & intCounter).Value = dbl_91Tot
      'Modified by Lydia 2016/12/09
'      .Range("G" & intCounter).Select
'      .Application.Selection.NumberFormatLocal = "#,##0_ "
'      .Range("G" & intCounter).Value = dbl_Tot
      '1年以上
      .Range("G" & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range("G" & intCounter).Value = dbl_Year
      '應收帳款總額
      .Range(strEnd & intCounter).Select
      .Application.Selection.NumberFormatLocal = "#,##0_ "
      .Range(strEnd & intCounter).Value = dbl_Tot
      'end 2016/12/09
      
      '雙線
      'Modified by Lydia 2016/12/09
      'strTemp = "C" & intCounter & ":G" & intCounter
      strTemp = "C" & intCounter & ":" & strEnd & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlDouble
         .Weight = xlThick
         .ColorIndex = xlAutomatic
      End With
      intCounter = intCounter + 1
      .Range("D" & intCounter).Value = "***結束***"
   End With
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt108.Close
   MsgBox "檔案已產生！" & vbCrLf & vbCrLf & "存放至 " & strFilePath
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt108.Close
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2014/6/4
Public Sub PrintExcelTitle()
Dim i As Integer, strTemp As String
   
   intPage = intPage + 1
   With wksAnnuity
      For i = 1 To 3
         If i = 1 Then
            .Range("D" & intCounter).Value = IIf(Text5 = "1", "客戶帳齡分析表", IIf(Text5 = "2", "智權人員帳齡分析表", "智權人員客戶帳齡分析表"))
         ElseIf i = 2 Then
            intCounter = intCounter + 1
            If Text5 = "1" Then
               .Range("C" & intCounter).Value = "編　　　號：" & Text1.Text & "~" & Text2.Text
            Else
               .Range("C" & intCounter).Value = "編　　　號：" & Text3.Text & "~" & Text3.Text
            End If
         ElseIf i = 3 Then
            intCounter = intCounter + 1
            .Range("C" & intCounter).Value = "帳款截止日：" & MaskEdBox1.Text & " 止"
         End If
         If i = 1 Then
            strTemp = "A" & intCounter & ":G" & intCounter
            .Range(strTemp).Select
            With .Application.Selection
               .HorizontalAlignment = xlCenter
               .Font.Size = 18
            End With
         End If
      Next i
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "列印人員：" & strUserName
      .Range("F" & intCounter).Value = "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "是否發文：" & IIf(Text4 = "1", "已發文", IIf(Text4 = "2", "未發文", "全部"))
      .Range("F" & intCounter).Value = "頁　　次：" & intPage
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = IIf(Text5 = "2", "智權人員", "客戶編號")
      .Range("B" & intCounter).Value = "名稱"
      .Range("C" & intCounter).Value = "0~30天"
      .Range("D" & intCounter).Value = "31~60天"
      .Range("E" & intCounter).Value = "61~90天"
      'Modified by Lydia 2016/12/09
'      .Range("F" & intCounter).Value = "90天以上"
'      .Range("G" & intCounter).Value = "應收帳款"
'      strTemp = "A" & intCounter & ":G" & intCounter
      .Range("F" & intCounter).Value = "90天~365天"
      .Range("G" & intCounter).Value = "1年以上"
      .Range("H" & intCounter).Value = "應收帳款"
      strTemp = "A" & intCounter & ":H" & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       intCounter = intCounter + 1
   End With
End Sub
