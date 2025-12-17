VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44h0 
   AutoRedraw      =   -1  'True
   Caption         =   "部門綜合損益表"
   ClientHeight    =   3440
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3440
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   30
      Width           =   3350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
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
      Index           =   1
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   810
      Width           =   4600
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
      Index           =   0
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2300
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   450
      Width           =   855
      _ExtentX        =   1517
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   450
      Width           =   855
      _ExtentX        =   1517
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "label2(3)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "label2(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "label2(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "label2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1230
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "年月"
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
      Left            =   360
      TabIndex        =   5
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   360
      TabIndex        =   4
      Top             =   90
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc44h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoaccrpt412 As New ADODB.Recordset
Dim lngCounter As Long
'Modify by Amy 2016/07/27 douTotalX 改名稱及型態 原Double
Dim stTotal1(13) As String, stTotal2(13) As String, stTotal3(13) As String, stTotal4(13) As String, stTotal5(13) As String
'end 2016/07/27
Dim dllaccrpt412 As Object
'Add by Amy 2015/03/04
Dim strFieldN, intWidth()  'Modify by Amy 2020/04/21 原:strFieldN()
'Added by Lydia 2016/01/30 列印使用
Dim strTemp(0 To 11) As String
Dim PLeft(0 To 12) As Integer
Dim PTitle(0 To 11) As String
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 10
Private Const ciStartX = 0
Private Const ciStartY = 500
Private Const ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPrint As Integer, iPage As Integer
Dim ii As Integer, intField As Integer, intCounter As Integer, intTitleRow As Integer '2017/02/15 從ExcelSave搬過來
'Add by Amy 2020/04/21
Dim strFieldTB() 'Table對應欄位
Dim strCmp As String, strCmpN As String
Dim jj As Integer 'Add by Amy 2021/03/11

'Add by Amy 2020/04/21
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If

    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/21

Private Sub Command1_Click(Index As Integer)
    'Modify by Amy 2020/04/21 +bolShowMsg,公司改下拉 原:Text5
    Dim strQ As String
    Dim bolShowMsg As Boolean
    
    If FormCheck(bolShowMsg) = False Then
        If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    
    'Add by Amy 2023/02/21 公司別下拉拿掉L公司,且預設1+J,避免下法律所成立前年月未將L加入,故彈訊息
    If Trim(CboCmp) <> MsgText(601) And Replace(MaskEdBox1.Text, "/", "") < Val(Left(智慧所更名日, 6)) - 191100 Then
        If MsgBox("年月「起始日」為法律所成立前" & vbCrLf & _
                            "需含L資料請將公司別設「空白」" & vbCrLf & _
                            "若修改公司別請按「否」" & vbCrLf & _
                            "繼續請按「是」", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    strCmp = "": strCmpN = ""
    If Trim(CboCmp) <> MsgText(601) Then
       strCmp = CboCmp
       If InStr(strCmp, "　") > 0 Then
             strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
       End If
    End If
    strCmpN = GetAccReportCmpN(CboCmp, , True)
   
    'add by sonia 2017/12/8 檢查是否有未過帳傳票
    If CheckAX210(strCmp, Replace(MaskEdBox1.Text, "/", ""), Replace(MaskEdBox2.Text, "/", "")) = True Then
    'end 2020/04/21
        Exit Sub
    End If
    'end 2017/12/8
    
    'Memo 暫存檔若需加 數字欄或加總欄 請加於最後,若文字欄或 非加總欄 請加於R001-R010
    Screen.MousePointer = vbHourglass
    Accrpt412Delete
    ProduceData
    If adoaccrpt412.State = adStateOpen Then
       adoaccrpt412.Close
    End If
    adoaccrpt412.CursorLocation = adUseClient
    'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤 +strUserNum
'   adoaccrpt412.Open "select * from accrpt412 Where r41201='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2016/01/30 改成Printer
'   If adoaccrpt412.RecordCount <> 0 Then
'      '2014/1/23 modify by sonia
'      'dllaccrpt412.Acc44b0 ReportTitle(416), Text5, Text6, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      dllaccrpt412.Acc44b0 ReportTitle(416), IIf(Text5 = "2", "J", Text5), IIf(Text6 = "", "台一　專利商標/智權", Text6), MaskEdBox1.Text, MaskEdBox2.Text, strUserNum & "-" & StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   End If
'   'end 2015/04/15
'   adoaccrpt412.Close
    'Modify by Amy 2020/08/12 改暫存檔,先設欄位
    'strQ = "Select * From accrpt412 Where r41201='" & strUserNum & "' order by r41202 "
    strQ = "Select R003 as AccName,R011,R012,R013,R014,R015,R016,R017,R018 ,R019,R020,R021,R022,R023," & _
                         " '',R002 as AccNo " & _
                         "From Accrpt44H0 Where ID='" & strUserNum & "' order by R001 "
    adoaccrpt412.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt412.RecordCount <> 0 Then
        'Modify by Amy 2015/03/04 產生Excel
        If Index = 1 Then
            ExcelSave
        Else
            'PrintData 'Mark by Amy 2020/04/21 不使用
        End If
    End If
    If adoaccrpt412.State = adStateOpen Then
        adoaccrpt412.Close
    End If
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102) 'Mark by Amy 2020/04/21
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 4000  'modify by sonia 2025/4/10 原3670 'Moidfy by Amy 2020/09/14 調整說明 原:2200
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/21 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   'Modify by Amy 2023/02/21 預設1+J,不顯示L公司-莘
   Call Pub_SetCboCmp(CboCmp, True, False, False, Mid(組合作帳公司, 2), 1, , "L")
   'end 2020/04/21
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   'Modify by Amy 2020/09/14 調整說明
   Label2(0).Caption = "．報表有誤確認之步驟：" & vbCrLf & _
                                 "  1.先確認是哪個公司別資料有誤" & vbCrLf & _
                                 "  2.若「分攤部門費用」有錯，請確認「部門」是否輸錯"
   Label2(1).Caption = "Ex：1+J公司10908後，不應有法務收入或創新業務收入" & vbCrLf & _
                                    "　   科目輸SAL" & vbCrLf & _
                                    "       L公司105年後，部門應只能輸L"
   Label2(2).Caption = " 3.若「費用科目」有「差額」，請確認是否未輸" & vbCrLf & _
                                    "　「分攤比例」"
   'Modify by Amy 2023/02/21 拿掉L公司,修改說明
'   Label2(3).Caption = "PS.10904月後法律所成立，當年度L部門各營業損益" & vbCrLf & _
'                                    "      請注意!!"
   Label2(3).Caption = "PS.年月條件若為法律所成立前（10904月前），若需" & vbCrLf & _
                                    "     包含 L，請注意公司別需改為「空白」!!"
                            
   'Mark by Amy 2020/04/21 不使用
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'Set dllaccrpt412 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44h0 = Nothing
   Set dllaccrpt412 = Nothing
End Sub

'Mark by Amy 2020/04/21
'Private Sub Text5_Change()
'   '2014/1/23 modify by sonia
'   'If Text5 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text6 = A0802Query(Text5)
'   Select Case Text5
'      Case "1"
'         Text6 = A0802Query(Text5)
'      Case "2"
'         Text6 = A0802Query("J")
'      Case ""
'         Text6 = "台一　專利商標/智權"
'   End Select
'   '2014/1/23 end
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'End Sub
'
''2014/1/23 add by sonia
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/23 end
'end 2020/04/21

'*************************************************
'  產生報表資料
'
'*************************************************
'Modify by Amy 2020/08/12 增加ACS ,因列印已不用整理程式
Private Sub ProduceData()
    Dim intCounter As Integer
    Dim strQ As String, strQ2 As String, strWherea0101 As String, str9997 As String

On Error GoTo Checking
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
    lngCounter = 0
    Call SetField(strQ2)
    
    'Memo 暫存檔若需加 數字欄或加總欄 請加於最後,若文字欄或 非加總欄 請加於R001-R010
    '財務處可能同時兩個人執行此報表,造成資料錯誤
    strQ = "Select " & strQ2 & " From Accrpt44H0 Where ID='" & strUserNum & "' Order by R001 "
    If adoaccrpt412.State <> adStateClosed Then adoaccrpt412.Close
    adoaccrpt412.CursorLocation = adUseClient
    adoaccrpt412.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    
    '公司別
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strWherea0101 = "and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "') )"
        Else
            strWherea0101 = "and (a0109 is null or a0109='" & strCmp & "')"
        End If
    End If
    
'-------------------------------------------------
' 實際營業收入
'-------------------------------------------------
    strQ = "Select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Accrpt412Save
        adoacc010.MoveNext
    Loop
    adoacc010.Close
   
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "4S"
    adoaccrpt412.Fields("R003").Value = ReportSum(14)
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "4－"
    adoaccrpt412.UpdateBatch
   
'-------------------------------------------------
' 實際營業支出
'-------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Accrpt412Save
        adoacc010.MoveNext
    Loop
    adoacc010.Close
       
    adoaccrpt412.UpdateBatch
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "6S"
    adoaccrpt412.Fields("R003").Value = ReportSum(15)
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "6－"
    adoaccrpt412.UpdateBatch
    
   'Modify by Amy 2016/07/27 避免報表與Excel 值不同合計可能誤差,故資料先四捨五入到小數2位
    '部門損益
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "DS"
    adoaccrpt412.Fields("R003").Value = ReportSum(16)
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "DS－"
    adoaccrpt412.UpdateBatch
   
'-------------------------------------------------
' 分攤管理、智權部門費用
' Memo 2016/07/27 婧瑄:分攤科目改為公式顯示(參閱 UpdAccrpt412)
'-------------------------------------------------
    '105年起才有9997分攤法務部門費用科目,智慧所更名日後不需顯示9997分攤法務部門費用科目
    str9997 = ""
    If Val(Mid(MaskEdBox1.Text, 1, 3)) < 105 Or (Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(智慧所更名日, 6) - 191100)) Then
       str9997 = " and a0101<>'9997' "
    End If
    strQ = "select * from acc010 where a0101 >= '999' and a0101 <= '999999' " & str9997 & strWherea0101 & "order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Accrpt412Save
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "9S"
    adoaccrpt412.Fields("R003").Value = ReportSum(17)
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "9S－"
    adoaccrpt412.UpdateBatch
    
'---------------------------
'  計算各部門營業損益/利潤率
'---------------------------
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "ICS"
    adoaccrpt412.Fields("R003").Value = "各部門營業損益"
    adoaccrpt412.UpdateBatch
    
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "PMS"
    adoaccrpt412.Fields("R003").Value = "利潤率"
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "PMS－"
    adoaccrpt412.UpdateBatch
'-------------------------------------------------
' 實際營業外收入
'-------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Accrpt412Save
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "71S"
    adoaccrpt412.Fields("R003").Value = ReportSum(5)
    adoaccrpt412.UpdateBatch
    
    '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "71S－"
    adoaccrpt412.UpdateBatch
    
'-------------------------------------------------
' 實際營業外支出
'-------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Accrpt412Save
        adoacc010.MoveNext
    Loop
    adoacc010.Close
 
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "72S"
    adoaccrpt412.Fields("R003").Value = ReportSum(6)
    adoaccrpt412.UpdateBatch

   '空行
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "72S－"
    adoaccrpt412.UpdateBatch
       
'-------------------------------------------------
' 計算損益
'-------------------------------------------------
    adoaccrpt412.AddNew
    adoaccrpt412.Fields("ID").Value = strUserNum
    adoaccrpt412.Fields("R001").Value = Counter
    adoaccrpt412.Fields("R002").Value = "ZZZZZZZZ"
    adoaccrpt412.Fields("R003").Value = ReportSum(20)
    adoaccrpt412.UpdateBatch

    StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData_Old()
'Dim intCounter As Integer
'Dim str9997 As String   'add by sonia 2016/1/27
'Dim strQ As String, strWherea0101 As String 'Add by Amy 2020/04/21
'
'On Error GoTo Checking
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   lngCounter = 0
'   If adoacc010.State = adStateOpen Then adoacc010.Close
'   adoaccrpt412.CursorLocation = adUseClient
'   'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤
'   adoaccrpt412.Open "select * from accrpt412 Where r41201='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
''-------------------------------------------------
'' 實際營業收入
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'Modify by Amy 2020/04/21 公司別改下拉 原:Text5/IIf(Text5 = "2", "J", Text5),取消 and instr(a0102,'不用')=0
'   If strCmp <> MsgText(601) Then
'        If InStr(strCmp, "+") > 0 Then
'            strWherea0101 = "and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "') )"
'        Else
'            strWherea0101 = "and (a0109 is null or a0109='" & strCmp & "')"
'        End If
'   End If
'   strQ = "Select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
'   adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'   'end 2020/04/21
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt412Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "4" 'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(14)
'   adoaccrpt412.Fields("r41215").Value = "4S" 'Add by Amy 2015/03/04
'   For intCounter = 3 To 13
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate "4", "499999"
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
'         stTotal1(intCounter) = "0"
'      Else
'         stTotal1(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
''         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'      End If
'   Next intCounter
'
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41215").Value = "4E" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
''-------------------------------------------------
'' 實際營業支出
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'Modify by Amy 2020/04/21 公司別改下拉 原:Text5/IIf(Text5 = "2", "J", Text5),取消 and instr(a0102,'不用')=0
'    'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text5 = "2", "J", Text5) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   strQ = "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
'   adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'   'end 2020/04/21
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt412Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "6"  'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(15)
'   adoaccrpt412.Fields("r41215").Value = "6S" 'Add by Amy 2015/03/04
'   For intCounter = 3 To 13
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate "6", "699999"
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
'         stTotal2(intCounter) = "0"
'      Else
'         stTotal2(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
''         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "6"  'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.Fields("r41215").Value = "6E" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'   'Modify by Amy 2016/07/27 避免報表與Excel 值不同合計可能誤差,故資料先四捨五入到小數2位
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(16) '部門損益
'   adoaccrpt412.Fields("r41215").Value = "DS" 'Add by Amy 2015/03/04
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'      If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) = 0 Then
'         adoaccrpt412.Fields(intCounter).Value = Null
'      Else
'         adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)), 2), FAmount)
''         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41215").Value = "DE" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'
''-------------------------------------------------
'' 分攤管理、智權部門費用
'' Memo 2016/07/27 婧瑄:分攤科目改為公式顯示(參閱 UpdAccrpt412)
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'modify by sonia 2016/1/27 105年起才有9997分攤法務部門費用科目
'   'If Text5 <> "" Then
'   '   adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text5 = "2", "J", Text5) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'Else
'   '   adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'End If
'   str9997 = ""
'   'Modify by Amy 2020/04/21 智慧所更名日後不需顯示9997分攤法務部門費用科目
'   If Val(Mid(MaskEdBox1.Text, 1, 3)) < 105 Or (Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(智慧所更名日, 6) - 191100)) Then
'      str9997 = " and a0101<>'9997' "
'   End If
'   'Modify by Amy 2020/04/21 公司別改下拉 原:Text5/IIf(Text5 = "2", "J", Text5),取消 and instr(a0102,'不用')=0
'   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text5 = "2", "J", Text5) & "')" & str9997 & "order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   strQ = "select * from acc010 where a0101 >= '999' and a0101 <= '999999' " & str9997 & strWherea0101 & "order by a0101 asc"
'   adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'   'end 2020/04/21
'   'end 2016/1/27
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt412Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "9"  'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(17)
'   adoaccrpt412.Fields("r41215").Value = "9S" 'Add by Amy 2015/03/04
'   For intCounter = 3 To 12
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate "9", "999999"
'   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 11
'      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
'         stTotal3(intCounter) = "0"
'      Else
'         stTotal3(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
'         adoaccrpt412.Fields(12).Value = Format(Round(Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value), 2), FAmount)
'      End If
'   Next intCounter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41215").Value = "9E" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'  'Added by Lydia 2016/02/17
''---------------------------
''  計算各部門營業損益
''---------------------------
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "9E" '分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(18)
'   adoaccrpt412.Fields("r41215").Value = "VS"
'   For intCounter = 3 To 13
'      If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) = 0 Then
'         adoaccrpt412.Fields(intCounter).Value = 0
'      Else
'         adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)), 2), FAmount)
'      End If
'   Next intCounter
'   adoaccrpt412.Fields(11).Value = 0
'   adoaccrpt412.Fields(13).Value = 0   '財務處說總所/管理之各部門營業損益印0
'   'add by sonia 2016/7/13 105年起之法務部也是0
'   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'      adoaccrpt412.Fields(10).Value = 0
'   End If
'   'end 2016/7/13
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41215").Value = "VE"
'   adoaccrpt412.UpdateBatch
'   'end 2016/02/17
''-------------------------------------------------
'' 實際營業外收入
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'Modify by Amy 2020/04/21 公司別改下拉 原:Text5/IIf(Text5 = "2", "J", Text5),取消 and instr(a0102,'不用')=0
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text5 = "2", "J", Text5) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   strQ = "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
'   adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'   'end 2020/04/21
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt412Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "71" 'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(5)
'   adoaccrpt412.Fields("r41215").Value = "71S" 'Add by Amy 2015/03/04
'   For intCounter = 3 To 13
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate "71", "719999"
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
'         stTotal4(intCounter) = "0"
'      Else
'         stTotal4(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
''         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4) '分隔線
'   adoaccrpt412.Fields("r41215").Value = "71E" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'
''-------------------------------------------------
'' 實際營業外支出
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'Modify by Amy 2020/04/21 公司別改下拉 原:Text5/IIf(Text5 = "2", "J", Text5),取消 and instr(a0102,'不用')=0
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text5 = "2", "J", Text5) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   strQ = "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' " & strWherea0101 & " order by a0101 asc"
'   adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'   'end 2020/04/21
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt412Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4), "72"   'Modify by Amy 2015/03/04 分隔線,+會計科目
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(6)
'   adoaccrpt412.Fields("r41215").Value = "72S" 'Add by Amy 2015/03/04
'   For intCounter = 3 To 13
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate "72", "729999"
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
'         stTotal5(intCounter) = "0"
'      Else
'         stTotal5(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
''         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
''   adoaccrpt412.Fields("r41201").Value = strUserNum
''   adoaccrpt412.Fields("r41202").Value = Counter
''   adoaccrpt412.UpdateBatch
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4)  '分隔線
'   adoaccrpt412.Fields("r41215").Value = "72E" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'
''-------------------------------------------------
'' 計算損益
''-------------------------------------------------
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(4)  '分隔線
'   adoaccrpt412.Fields("r41215").Value = "" 'Add by Amy 2015/03/04
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41203").Value = ReportSum(20)
'   adoaccrpt412.Fields("r41215").Value = "ZZZZZZZZ" 'Add by Amy 2015/03/04
''   adoaccrpt412.Fields(12).Value = 0
'   For intCounter = 3 To 13
'        'edit by nick 2004/08/06 判斷，當分攤費用等於 0 時只計算營業外收入
'        '2014/3/6 modify by sonia 改為判斷intCounter=11,13(智權部,總所/管理),否則J公司的所有部門分攤費都為0
'        'If douTotal3(intCounter) = 0 And intCounter <> 12 Then
'        If intCounter = 11 Or intCounter = 13 Then
'            '智權部11,總所/管理13
'            If Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)) = 0 Then
'               adoaccrpt412.Fields(intCounter).Value = Null
'            Else
'               adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)), 2), FAmount)
'            End If
'        'add by sonia 2016/2/16 intCounter = 10,105年起為法務部,改同智權部及管理部做法,105年以前為投法仍維持原做法
'        'Modify by Amy 2020/04/21 條件起迄若含智慧所更名日前資料,以舊格式顯示 ex:10903~10904/條件下L公司 顯示L公司格式
'        ElseIf intCounter = 10 And Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
'            If Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)) = 0 Then
'               adoaccrpt412.Fields(intCounter).Value = Null
'            Else
'               adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)), 2), FAmount)
'            End If
'        'end 2016/2/16
'        Else
'             If intCounter = 12 Then
'                    '全所
'                    adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) + Val(stTotal4(intCounter) - stTotal5(intCounter)), 2), FAmount)
'             Else
'                  If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) + Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)) = 0 Then
'                     adoaccrpt412.Fields(intCounter).Value = Null
'                  Else
'                     adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) + Val(stTotal4(intCounter)) - Val(stTotal5(intCounter)), 2), FAmount)
'                  End If
'            End If
'        End If
'   Next intCounter
'   'adoaccrpt412.Fields(11).Value = Null   '2014/3/6 cancel by sonia 上面計算
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   PaintLine ReportSum(8)
'   adoaccrpt412.UpdateBatch
'   adoaccrpt412.Close
'   'Modify by Amy 2020/04/21 +排除單下 L公司不需分攤(條件起迄若含智慧所更名日前資料,以舊格式顯示)
'   If strCmp <> "L" Then
'    UpdAccrpt412 'Add by Amy 2016/07/25
'   End If
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt412Delete()
    'Modify by Amy 2020/08/12 改暫存檔
    'adoTaie.Execute "delete from accrpt412 Where r41201='" & strUserNum & "' "
    adoTaie.Execute "Delete From Accrpt44H0 Where ID='" & strUserNum & "' "
End Sub

'*************************************************
'  儲存資料表(部門損益比較表暫存檔)
'
'*************************************************
'Add by Amy 2020/08/12 欄位改抓變數
Private Sub Accrpt412Save()
Dim intCounter As Integer
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("ID").Value = strUserNum
   adoaccrpt412.Fields("R001").Value = Counter
   adoaccrpt412.Fields("R002").Value = "" & adoacc010.Fields("a0101") '會計科目代碼
   If IsNull(adoacc010.Fields("a0102").Value) Then
      adoaccrpt412.Fields(adoaccrpt412.Fields(GetValue("會計科目")).Name).Value = Null
   Else
      adoaccrpt412.Fields(adoaccrpt412.Fields(GetValue("會計科目")).Name).Value = adoacc010.Fields("a0102").Value
   End If
   For intCounter = GetValue("專利") To GetValue("全所")
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value
   
   If Mid(adoacc010.Fields("a0101").Value, 1, 1) = "9" Then
      adoaccrpt412.Fields(GetValue("管理")).Value = 0
      For intCounter = GetValue("專利") To GetValue("智權部")
         If IsNull(adoaccrpt412.Fields(intCounter).Value) = False Then
            adoaccrpt412.Fields(GetValue("管理")) = Val(adoaccrpt412.Fields(GetValue("管理"))) + Val(adoaccrpt412.Fields(intCounter))
         End If
      Next intCounter
   End If
   adoaccrpt412.UpdateBatch
End Sub

'Mark by Amy 2020/08/12
Private Sub Accrpt412Save_Old()
'Dim intCounter As Integer
'
'   adoaccrpt412.AddNew
'   adoaccrpt412.Fields("r41201").Value = strUserNum
'   adoaccrpt412.Fields("r41202").Value = Counter
'   adoaccrpt412.Fields("r41215").Value = "" & adoacc010.Fields("a0101") 'Add by Amy 2015/03/04 +會計科目代碼
'   If IsNull(adoacc010.Fields("a0102").Value) Then
'      adoaccrpt412.Fields("r41203").Value = Null
'   Else
'      adoaccrpt412.Fields("r41203").Value = adoacc010.Fields("a0102").Value
'   End If
'   For intCounter = 3 To 12
'      adoaccrpt412.Fields(intCounter).Value = 0
'   Next intCounter
'   Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value
'   If Mid(adoacc010.Fields("a0101").Value, 1, 1) = "9" Then
'      adoaccrpt412.Fields("r41213").Value = 0
'      For intCounter = 3 To 11
'         If IsNull(adoaccrpt412.Fields(intCounter).Value) = False Then
'            adoaccrpt412.Fields("r41213").Value = Val(adoaccrpt412.Fields("r41213").Value) + Val(adoaccrpt412.Fields(intCounter).Value)
'         End If
'      Next intCounter
'   End If
'   adoaccrpt412.UpdateBatch
End Sub

'*************************************************
'  計算各部門小計金額
'
'*************************************************
'Add by Amy 2020/08/12 改暫存檔,改動態欄位
Private Sub Calculate(strAccNo1 As String, strAccNo2 As String)
    Dim strDebit As String, strA As String, strA2 As String, stFN As String
    Dim intCounter As Integer, strSql As String
      
    intCounter = 3
   
    '公司別
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
        Else
            strSql = " and a0403 = '" & strCmp & "'"
        End If
    End If
    '年月
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
        strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402)  >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
         strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
    End If
    
    If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
    End If
    If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
    End If
    
    '因109年發現Acc090 分攤部門不會有CFL及FCL,若條件下105年前資料會抓不到
    If Val(Mid(MaskEdBox1.Text, 1, 3)) < 105 Then
        strA = " Union Select 'FCL','FCL' From Dual "
    End If
    
    strA = "Select a0901,a0902 From Acc090 Where a0904 = 'Y' " & strA & "Order by a0901 asc"
    If adoacc090.State = adStateOpen Then adoacc090.Close
    adoacc090.CursorLocation = adUseClient
    adoacc090.Open strA, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc090.EOF = False
        '智慧所更名日後,L公司或畫面條件公司別為空,顯示「L」(L部門),但1與J公司不顯示「L」(產生Excel時,隱藏欄位)
        '105年起只有「法務部」(L部門),放在原「投法」的位置(因只有L部門,FCL及CFL沒使用)
        If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
            If InStr("" & adoacc090.Fields("a0901"), "L") > 0 And "" & adoacc090.Fields("a0901") <> "SAL" Then
                strA2 = "select sum(a0408) from acc040 where a0404 in ('" & adoacc090.Fields("a0901").Value & "','FCL','CFL')" & strSql
            Else
                strA2 = "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql
            End If
        '105年前「法務」(L部門)「投法」(CFL+FCL部門)
        Else
            '因109年發現Acc090 分攤部門不會有CFL及FCL,若條件下105年前資料會抓不到
            If adoacc090.Fields("a0901").Value = "FCL" Then
                strA2 = "select sum(a0408) from acc040 where a0404 in ('CFL','FCL')" & strSql
            Else
                strA2 = "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql
            End If
        End If
        
        If adoacc040.State = adStateOpen Then adoacc040.Close
        adoacc040.CursorLocation = adUseClient
        adoacc040.Open strA2, adoTaie, adOpenStatic, adLockReadOnly
        If adoacc040.RecordCount <> 0 Then
            If IsNull(adoacc040.Fields(0).Value) Then
               strDebit = "0"
            Else
               strDebit = Format(adoacc040.Fields(0).Value, FAmount)
            End If
            
            stFN = "" & adoacc090.Fields("a0901")
            Select Case stFN
                Case "P"
                    stFN = "專利"
                Case "T"
                    stFN = "商標"
                Case "L"
                    '若條件為 智慧所更名日 且公司不為L 或空白仍需顯示 法務部
                    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
                        stFN = "法務部"
                    End If
                Case "FCL"
                    stFN = "投法"
                Case "W"
                    stFN = "ACS"
                Case "SAL"
                    stFN = "智權部"
                Case "M"
                    stFN = "管理"
                Case "TOT"
                    stFN = "全所"
            End Select
          
            adoaccrpt412.Fields(adoaccrpt412.Fields(GetValue(stFN)).Name) = strDebit
        End If
        adoacc040.Close
        adoacc090.MoveNext
    Loop
    
    Select Case Mid(strAccNo1, 1, 1)
        Case "6"
        Case "9"
            '會計科目不分攤 ex:9997-9999
            adoaccrpt412.Fields(GetValue("管理")).Value = 0
        Case Else
            adoaccrpt412.Fields(GetValue("管理")).Value = 0
            For intCounter = GetValue("專利") To GetValue("智權部")
                adoaccrpt412.Fields(GetValue("管理")).Value = Val(Format(adoaccrpt412.Fields(GetValue("管理")).Value, FAmount)) - Val(Format(adoaccrpt412.Fields(intCounter).Value, FAmount))
            Next intCounter
            adoaccrpt412.Fields(GetValue("管理")).Value = Format(Val(Format(adoaccrpt412.Fields(GetValue("管理")).Value, FAmount)) + Val(Format(adoaccrpt412.Fields(GetValue("全所")).Value, FAmount)), FAmount)
    End Select
    adoacc090.Close
End Sub

'Mark by Amy 2020/08/12
Private Sub Calculate_Old(strAccNo1 As String, strAccNo2 As String)
''Modify by Amy 2016/07/25 因只下 2.智權10501-06 會造成多重步驗操作…的錯誤
''Dim douDebit As Double
'Dim strDebit As String
''end 2016/07/25
'Dim intCounter As Integer, strSql As String
'
'   intCounter = 3
'   adoacc090.CursorLocation = adUseClient
'   adoacc090.Open "select * from acc090 where a0904 = 'Y' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc090.EOF = False
'      adoacc040.CursorLocation = adUseClient
'      strSql = MsgText(601)
'      'Modify by Amy 2020/04/21 公司別改下拉 原:Text5
'      If strCmp <> MsgText(601) Then
'         If InStr(strCmp, "+") > 0 Then
'            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
'         Else
'            '2014/1/23 modify by sonia
'            'strSql = " and a0403 = '" & Text5 & "'"
'            strSql = " and a0403 = '" & strCmp & "'"
'            '2014/1/23 end
'         End If
'      End If
'      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
'         strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402)  >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
'      End If
'      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
'         strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
'      End If
'      If strAccNo1 <> MsgText(601) Then
'         strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
'      End If
'      If strAccNo2 <> MsgText(601) Then
'         strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
'      End If
'      'add by sonia 2016/1/27 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
'      If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'         'Modify by Amy 2020/04/21 條件起迄若含智慧所更名日前資料,以舊格式顯示 ex:10903~10904/條件下L公司 顯示FCL,CFL欄
'         If adoacc090.Fields("a0901").Value = "L" And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
'            adoacc040.Open "select sum(a0408) from acc040 where a0404 in ('" & adoacc090.Fields("a0901").Value & "','FCL','CFL')" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'         Else
'            adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'         End If
'      'end 2016/1/27
'      'MODIFY BY SONIA 2013/11/7 102/10 CFL的416102會因為此段最下方的計算跑到總所/管理
'      'adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      ElseIf adoacc090.Fields("a0901").Value = "FCL" Then
'         adoacc040.Open "select sum(a0408) from acc040 where a0404 in ('" & adoacc090.Fields("a0901").Value & "','CFL')" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      Else
'         adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      End If
'      '2013/11/7 end
'      'Modify by Amy 2016/07/25 改 douDebit為strDebit
'      If adoacc040.RecordCount <> 0 Then
'         If IsNull(adoacc040.Fields(0).Value) Then
'            strDebit = "0"
'         Else
'            strDebit = Format(adoacc040.Fields(0).Value, FAmount)
'         End If
'         Select Case adoacc090.Fields("a0901").Value
'            Case "P"
'               adoaccrpt412.Fields(3).Value = strDebit
'            Case "T"
'               adoaccrpt412.Fields(4).Value = strDebit
'            Case "L"
'               'modify by sonia 2016/1/27 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
'               'adoaccrpt412.Fields(5).Value = douDebit
'               'Modify by Amy 2020/04/21 條件起迄若含智慧所更名日前資料,以舊格式顯示 ex:10903~10904/條件下L公司 顯示FCL,CFL欄
'               If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
'                  adoaccrpt412.Fields(10).Value = strDebit
'               Else
'                  adoaccrpt412.Fields(5).Value = strDebit
'               End If
'               'end 2016/1/27
'            'Add by Amy 2020/04/21
'            Case "CFL"
'               adoaccrpt412.Fields(15).Value = strDebit
'            Case "CFP"
'               adoaccrpt412.Fields(6).Value = strDebit
'            Case "CFT"
'               adoaccrpt412.Fields(7).Value = strDebit
'            Case "FCP"
'               adoaccrpt412.Fields(8).Value = strDebit
'            Case "FCT"
'               adoaccrpt412.Fields(9).Value = strDebit
'            Case "FCL"
'               adoaccrpt412.Fields(10).Value = strDebit
'            Case "SAL"
'               adoaccrpt412.Fields(11).Value = strDebit
'            Case "TOT"
'               adoaccrpt412.Fields(12).Value = strDebit
'            Case "M"
'               adoaccrpt412.Fields(13).Value = strDebit
'         End Select
'      End If
'      'end 2016/07/25
'      adoacc040.Close
'      adoacc090.MoveNext
'   Loop
'   Select Case Mid(strAccNo1, 1, 1)
'      Case "6"
'      Case "9"
'         adoaccrpt412.Fields(13).Value = 0
'      Case Else
'         adoaccrpt412.Fields(13).Value = 0
'         For intCounter = 3 To 11
'            adoaccrpt412.Fields(13).Value = Val(Format(adoaccrpt412.Fields(13).Value, FAmount)) - Val(Format(adoaccrpt412.Fields(intCounter).Value, FAmount))
'         Next intCounter
'         adoaccrpt412.Fields(13).Value = Format(Val(Format(adoaccrpt412.Fields(13).Value, FAmount)) + Val(Format(adoaccrpt412.Fields(12).Value, FAmount)), FAmount)
'   End Select
'   adoacc090.Close
End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String, Optional strAccNo As String)
Dim intCounter As Integer
    
   'Modify by Amy 2015/03/04 +會計科目欄位
   'Modify by Amy 2020/08/12 改抓變數及新暫存檔
   adoaccrpt412.Fields("R002").Value = strAccNo & LeftB(strSign, 2) 'TB欄位 varchar(8)
   adoaccrpt412.Fields("R003").Value = strSign
   For intCounter = GetValue("專利") To GetValue("全所")
        adoaccrpt412.Fields(intCounter).Value = strSign
   Next intCounter
   'end 2020/08/12
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/21 公司別改下拉
'   Text5 = ""
'   Text6 = "台一　專利商標/智權"
   'CboCmp = "" 'Mark by Amy 2023/02/21 法律所成立後,不會有L公司
   'end 2020/04/21
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   CboCmp.SetFocus 'Modify by Amy 2020/04/21
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2020/04/21 +bolShowMsg
Public Function FormCheck(bolShowMsg As Boolean) As Boolean
   'Add by Amy 2020/04/21
   Dim bCancel As Boolean
   
   If Trim(CboCmp) <> MsgText(601) Then
        Call CboCmp_Validate(bCancel)
        If bCancel = True Then
            bolShowMsg = True
            Exit Function
        End If
   End If
   'end 2020/04/21
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
       If MaskEdBox2 <> Mid(MsgText(29), 1, 6) Then
            If Val(Replace(MaskEdBox1, "/", "")) > Val(Replace(MaskEdBox2, "/", "")) Then
                MsgBox Label4.Caption & " 起日不可大於迄日"
                bolShowMsg = True
                Exit Function
            End If
      End If
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      If Val(Replace(MaskEdBox1, "/", "")) > Val(Replace(MaskEdBox2, "/", "")) Then
            MsgBox Label4.Caption & " 起日不可大於迄日"
            bolShowMsg = True
            Exit Function
      End If
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add by Amy 2020/08/12 +ACS 並改新暫存檔,抓欄位使用別名
Private Sub ExcelSave()
   Dim xlsAnnuity As New Excel.Application
   Dim wksAnnuity As New Worksheet
   Dim strWkName As String '工作表名稱為中文
   Dim bol105YA As Boolean '是否為105年後資料
   Dim intSeq As Integer
   Dim strFileName As String, strTemp As String, strTemp2 As String
   Dim strStartRow As String, strEndRow As String '合計起/迄始位置
   Dim strTotal(2) As String, strDSum(1) As String '加總列號(0:其他 1:智權部及總所/管理部 2:全所-不計入)/0:營業收入1:支出加總列號
   Dim strTotPos(1 To 2) As String '全所損益欄位
   Dim strVSum(1) '各部門加總列號(0:其他/1:全所)
   Dim strOSum(1) As String '營業外收入/支出加總列號
   Dim strICS As String, strVal1 As String, strVal2 As String '各部門損益列/計算比率用
   Dim str9998 As String, str9999 As String 'Add by Amy 2021/03/11 分攤 管理/智權 費用公式
    
   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then bol105YA = True
 
On Error GoTo ErrHnd
    
   intField = 65:  intCounter = 1
   strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "-" & Val(Replace(MaskEdBox2.Text, "/", "")) & "部門綜合損益表" & IIf(strCmp <> MsgText(601), strCmp & "公司", "") & ServerDate & MsgText(43)
   If Dir(strExcelPath & strFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
       End If
   Else
       Kill strExcelPath & strFileName
   End If
   
   xlsAnnuity.SheetsInNewWorkbook = 3
   xlsAnnuity.Workbooks.add
   '工作表名稱改為中文
   If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
   Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & "1")
   wksAnnuity.Activate
   Call SetTitle(wksAnnuity, 1)
   
   intTitleRow = intCounter: intCounter = intCounter + 1: strStartRow = intCounter
   With wksAnnuity
        Do While adoaccrpt412.EOF = False
            If InStr("" & adoaccrpt412.Fields("AccNo"), "－") > 0 Then
                '空行
            '分攤科目-公式
            ElseIf "" & adoaccrpt412.Fields("AccNo") = "9997" Or "" & adoaccrpt412.Fields("AccNo") = "9998" Or "" & adoaccrpt412.Fields("AccNo") = "9999" Then
                For ii = UBound(strFieldN) To LBound(strFieldN) Step -1
                    If GetValue("會計科目") = ii Then
                        strTemp = "" & adoaccrpt412.Fields("AccName")
                    ElseIf GetValue("全所") = ii Then
                        Select Case "" & adoaccrpt412.Fields("AccNo")
                            Case "9997"
                                strTemp = "=" & Chr(intField + GetValue("法務部")) & strDSum(1)
                            Case "9998"
                                strTemp = "=" & Chr(intField + GetValue("管理")) & strDSum(1)
                            Case Else
                                strTemp = "=" & Chr(intField + GetValue("智權部")) & strDSum(1)
                        End Select
                        'Modify by Amy 2023/02/16 可能有負值,故有 減項 需加ABS
                        'Add by Amy 2021/03/11 分攤公式(判斷若任一營業收入<0需於全所實際收入減掉,但ACS/智權部/管理 的實際收入固定減)
                        '分攤管理部門費用
                        If "" & adoaccrpt412.Fields("AccNo") = "9998" Then
                            'Memo by Amy 專利不會有負的情況,故不需判斷
                            For jj = GetValue("商標") To GetValue("管理")
                                If jj = GetValue("ACS") Or jj = GetValue("智權部") Or jj = GetValue("管理") Then
                                    '固定減
                                    If jj = GetValue("ACS") And Not ((Val(Replace(MaskEdBox1, "/", "")) <= 10908 And Val(Replace(MaskEdBox2, "/", ""))) >= 10908 _
                                                                                Or Val(Replace(MaskEdBox1, "/", "")) >= 10908 Or Val(Replace(MaskEdBox2, "/", "")) >= 10908) Then
                                        '畫面條件起月未包含10908年後資料,不會show ACS
                                    Else
                                        str9998 = str9998 & "-ABS(" & Chr(intField + jj) & strDSum(0) & ")"
                                    End If
                                Else
                                    '營業收入<0才需減
                                    If jj = GetValue("法務部") And Not (Val(Replace(MaskEdBox1, "/", "")) >= 10501 And Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100) Then
                                        '畫面條件起月不是 105年後且智慧所更名日前,不會show 法務部
                                    ElseIf jj = GetValue("投法") And Not (Val(Replace(MaskEdBox1, "/", "")) < 10501) Then
                                        '畫面條件起月不是 105年前資料,不會show 法投
                                    ElseIf jj = GetValue("L") And (Val(Replace(MaskEdBox1, "/", "")) >= 10501 Or (Val(Replace(MaskEdBox1, "/", "")) >= Val(Left(智慧所更名日, 6)) - 191100 And (strCmp <> "L" Or strCmp <> MsgText(601)))) Then
                                        '畫面條件起月包含105年後資料,不會show  L
                                    Else
                                        str9998 = str9998 & "-if(" & Chr(intField + jj) & strDSum(0) & "<0," & "ABS(" & Chr(intField + jj) & strDSum(0) & "),0)"
                                    End If
                                End If
                            Next jj
                        '分攤智權部門費用
                        ElseIf "" & adoaccrpt412.Fields("AccNo") = "9999" Then
                            For jj = GetValue("商標") To GetValue("管理")
                                If jj = GetValue("FCP") Or jj = GetValue("FCT") Then
                                    '原本就固定減
                                ElseIf jj = GetValue("ACS") Or jj = GetValue("智權部") Or jj = GetValue("管理") Then
                                    '固定減
                                    If jj = GetValue("ACS") And Not ((Val(Replace(MaskEdBox1, "/", "")) <= 10908 And Val(Replace(MaskEdBox2, "/", ""))) >= 10908 _
                                                                                Or Val(Replace(MaskEdBox1, "/", "")) >= 10908 Or Val(Replace(MaskEdBox2, "/", "")) >= 10908) Then
                                        '畫面條件起月未包含10908年後資料,不會show ACS
                                    Else
                                        str9999 = str9999 & "-ABS(" & Chr(intField + jj) & strDSum(0) & ")"
                                    End If
                                Else
                                    '營業收入<0才需減
                                    If jj = GetValue("法務部") And Not (Val(Replace(MaskEdBox1, "/", "")) >= 10501 And Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100) Then
                                        '畫面條件起月不是 105年後且智慧所更名日前,不會show 法務部
                                    ElseIf jj = GetValue("投法") And Not (Val(Replace(MaskEdBox1, "/", "")) < 10501) Then
                                        '畫面條件起月不是 105年前資料,不會show 法投
                                    ElseIf jj = GetValue("L") And (Val(Replace(MaskEdBox1, "/", "")) >= 10501 Or (Val(Replace(MaskEdBox1, "/", "")) >= Val(Left(智慧所更名日, 6)) - 191100 And (strCmp <> "L" Or strCmp <> MsgText(601)))) Then
                                        '畫面條件起月包含105年後資料,不會show  L
                                    Else
                                        str9999 = str9999 & "-if(" & Chr(intField + jj) & strDSum(0) & "<0," & "ABS(" & Chr(intField + jj) & strDSum(0) & "),0)"
                                    End If
                                End If
                            Next jj
                        End If
                        'end 2021/03/11
                        'end 2023/02/16
                    ElseIf ii = GetValue("差額") Then
                        strTemp = "=" & Chr(intField + GetValue("全所")) & intCounter & "-Sum(" & Chr(intField + GetValue("專利")) & intCounter & ":" & Chr(intField + GetValue("管理")) & intCounter & ")"
                    ElseIf GetValue("專利") = ii Then
                        '與其他欄位一樣用算的再加總全所會造成小數位與XX部門費用合計不符
                        strTemp = "=Round(" & Chr(intField + GetValue("全所")) & intCounter & "-Sum(" & Chr(intField + GetValue("商標")) & intCounter & ":" & Chr(intField + GetValue("管理")) & intCounter & "),2)"
                    '避免除數為0時,計算比率時會錯,故判斷值為0時,顯示0
                    ElseIf .Range(Chr(intField + GetValue("全所")) & strDSum(0)).Value = 0 Then
                        strTemp = "0"
                    ElseIf "" & adoaccrpt412.Fields("AccNo") = "9997" Then
                        '法務部門費用總合*(該部門當月實際收入/全所實際收入)
                        strTemp = "=Round(" & Chr(intField + GetValue("法務部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & Chr(intField + GetValue("全所")) & strDSum(0) & "),2)"
                    'Modify by Amy 2021/03/11 管理部門費用,智權部固定不分攤,分攤比例計算時,任一合計費用<0需於全所實際收入減掉
                    ElseIf "" & adoaccrpt412.Fields("AccNo") = "9998" Then
                        '管理部門費用總合*(該部門當月實際收入/(全所實際收入-ACS 收入-智權部 收入-總所/管理 收入-其他部門收入為負值))
                        'Modify by Amy 2023/02/16 +管理部固定不分攤
                        If GetValue("智權部") = ii Or GetValue("管理") = ii Or Val(.Range(Chr(intField + ii) & strDSum(0)).Value) <= 0 Then
                            '智權部固定不分攤,任一收入合計<0 不分攤
                            strTemp = "0"
                        Else
                            '管理部門費用總合*(該部門當月實際收入/全所實際收入-if(商標實際收入<0,商標實際收入,0)-...-智權部實際收入) +strShareCost
                            strTemp = "=Round(" & Chr(intField + GetValue("管理")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/(" & Chr(intField + GetValue("全所")) & strDSum(0) & str9998 & ")),2)"
                        End If
                    '智權部費用9999
                    Else
                        If GetValue("FCP") = ii Or GetValue("FCT") = ii Or GetValue("投法") = ii Or GetValue("智權部") = ii Or GetValue("管理") = ii _
                          Or Val(.Range(Chr(intField + ii) & strDSum(0)).Value) <= 0 Then
                            'FCP、FCT(國外部)/投法/智權部/管理/任一收入合計<0  不分攤
                            strTemp = "0"
                        Else
                            '智權部費用總合*(該部門當月實際收入/(全所實際收入-FCP收入-FCT收入-105前投法))-2021/03/11未改前
                            '智權部費用總合*(該部門當月實際收入/(全所實際收入-FCP收入-FCT收入-105前投法-if(商標實際收入<0,商標實際收入,0)-...-智權部實際收入))
                            'Modify by Amy 2023/02/16 +ABS
                            strTemp = "=Round( " & Chr(intField + GetValue("智權部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & _
                                                "(" & Chr(intField + GetValue("全所")) & strDSum(0) & "-ABS(" & Chr(intField + GetValue("FCP")) & strDSum(0) & ")-ABS(" & Chr(intField + GetValue("FCT")) & strDSum(0) & ")" & _
                                                "-" & IIf(bol105YA = True, "0", "-ABS(" & Chr(intField + GetValue("投法")) & strDSum(0) & ")") & _
                                                 str9998 & ")),2)"
                        End If
                    'end 2021/03/11
                    End If
                    
                    If GetValue("會計科目") <> ii Then
                        .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
                    End If
                    .Range(Chr(intField + ii) & intCounter).Value = strTemp
                Next ii
            '利潤率=該部門營業損益/該部門實際營業收入
            ElseIf "" & adoaccrpt412.Fields("AccNo") = "PMS" Then
                For ii = LBound(strFieldN) To UBound(strFieldN)
                    strTemp = ""
                    strVal1 = .Range(Chr(intField + ii) & strICS).Value
                    strVal2 = .Range(Chr(intField + ii) & strDSum(0)).Value
                    If ii = GetValue("會計科目") Then
                        strTemp = "" & adoaccrpt412.Fields("AccName")
                    '未判斷是否為0程式會錯
                    ElseIf GetValue("差額") <> ii And Val(strVal1) <> 0 And Val(strVal2) <> 0 Then
                        strTemp = "=" & Chr(intField + ii) & strICS & "/" & Chr(intField + ii) & strDSum(0)
                    End If
                    If GetValue("會計科目") <> ii Then
                        .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "0.00%;[紅色]-0.00%"
                    End If
                    .Range(Chr(intField + ii) & intCounter).Value = strTemp
                Next ii
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strFieldN)) & intCounter).Interior.ColorIndex = 20 'Add by Amy 2020/09/14 淺藍
                '更新合計起始位置
                strStartRow = intCounter + 2
            '依欄位抓資料
            Else
                For ii = LBound(strFieldN) To UBound(strFieldN)
                    If ii = GetValue("會計科目") Then
                        strTemp = "" & adoaccrpt412.Fields("AccName")
                    ElseIf ii = GetValue("差額") Then
                        strTemp = "=" & Chr(intField + GetValue("全所")) & intCounter & "-Sum(" & Chr(intField + GetValue("專利")) & intCounter & ":" & Chr(intField + GetValue("管理")) & intCounter & ")"
                    '總所損益
                    ElseIf "" & adoaccrpt412.Fields("AccNo") = "ZZZZZZZZ" Then
                        Select Case ii
                            'Modify by Amy 2020/09/14 法務部拆開判斷,10904月後單下L公司或公司空白公式不為營業外收入-營業外支出
                            Case GetValue("法務部")
                                If Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100 And strCmp <> "L" And strCmp <> MsgText(601) Then
                                    '營業外收入-營業外支出
                                    'Modify by Amy 2023/03/16 +ABS
                                    strTemp = "=" & Chr(intField + ii) & strOSum(0) & "-ABS(" & Chr(intField + ii) & strOSum(1) & ")"
                                Else
                                    strTemp = "=" & Replace(strTotal(0), ",", "," & Chr(intField + ii))
                                    strTemp = "=Sum(" & Right(Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)), Len(strTemp) - 1) & ")"
                                End If
                            Case GetValue("智權部"), GetValue("管理")
                            'end 2020/09/14
                                '營業外收入-營業外支出
                                'Modify by Amy 2023/03/16 +ABS
                                strTemp = "=" & Chr(intField + ii) & strOSum(0) & "-ABS(" & Chr(intField + ii) & strOSum(1) & ")"
                            Case GetValue("全所")
                                '排除不計入全所的列(分攤合計列)
                                strTemp = "=" & Replace(Replace(strTotal(0), strTotal(2), ""), ",", "," & Chr(intField + ii))
                                strTotPos(1) = Chr(intField + ii): strTotPos(2) = intCounter
                            Case Else
                                strTemp = "=" & Replace(strTotal(0), ",", "," & Chr(intField + ii))
                        End Select
                        '營業外收入 -營業外支出
                        ''Modify by Amy 2020/09/14
                        If GetValue("智權部") <> ii And GetValue("管理") <> ii And GetValue("法務部") <> ii Then
                            strTemp = "=Sum(" & Right(Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)), Len(strTemp) - 1) & ")"
                        End If
                        
                    '會計科目大項合計
                    ElseIf InStr("" & adoaccrpt412.Fields("AccNo"), "S") > 0 Then
                        .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
                        
                        '部門損益
                        If "" & adoaccrpt412.Fields("AccNo") = "DS" Then
                            'Modify by Amy 2023/02/16 +ABS 避免負值變成加
                            strTemp = "=" & Chr(intField + ii) & strDSum(0) & "-ABS(" & Chr(intField + ii) & strDSum(1) & ")"
                        '各部門營業損益
                        ElseIf "" & adoaccrpt412.Fields("AccNo") = "ICS" Then
                            '法務部也是0
                            'Modify by Amy 2020/09/14 法務拆開判斷,10904月後單下L公司也要設公式
                            If GetValue("智權部") = ii Or GetValue("管理") = ii Then
                                strTemp = 0
                            ElseIf GetValue("法務部") = ii Then
                                '109
                                If Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100 And strCmp <> "L" And strCmp <> MsgText(601) Then
                                    strTemp = 0
                                Else
                                    strTemp = Replace(strVSum(0), ",", "," & Chr(intField + ii))
                                    strTemp = "=Sum(" & Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)) & ")"
                                End If
                            'end 2020/09/14
                            ElseIf GetValue("全所") = ii Then
                                strTemp = "=" & Chr(intField + ii) & strVSum(1)
                            Else
                                strTemp = Replace(strVSum(0), ",", "," & Chr(intField + ii))
                                strTemp = "=Sum(" & Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)) & ")"
                            End If
                        Else
                            strTemp = "=Sum(" & Chr(intField + ii) & strStartRow & ":" & Chr(intField + ii) & intCounter - 1 & ")"
                        End If
                        
                        'strDSum():部門損益 / strTotal():全所損益 計算欄位 / strVSum():各部門營業損益
                        strTemp2 = "" & adoaccrpt412.Fields("AccNo")
                        If LBound(strFieldN) + 1 = ii Then
                            If strTemp2 = "ICS" Then
                                strICS = intCounter
                            Else
                                Select Case Left(strTemp2, 1)
                                    Case "4" '營業收入
                                        strDSum(0) = intCounter
                                    Case "6" '營業支出
                                         strDSum(1) = intCounter
                                    Case "7" '營業外收支
                                        If Left(strTemp2, 2) = "71" Then
                                            strOSum(0) = intCounter
                                            strTotal(0) = strTotal(0) & "," & intCounter
                                            strTotal(1) = strTotal(1) & "," & intCounter
                                        Else
                                            strOSum(1) = intCounter
                                            strTotal(0) = strTotal(0) & ",-" & intCounter
                                            strTotal(1) = strTotal(1) & ",-" & intCounter
                                        End If
                                    Case "9" '分攤費用
                                        strVSum(0) = strVSum(0) & ",-" & intCounter
                                        strTotal(0) = strTotal(0) & ",-" & intCounter
                                        strTotal(2) = ",-" & intCounter
                                    Case "D" '部門損益
                                        strVSum(0) = strVSum(0) & "," & intCounter
                                        strVSum(1) = intCounter
                                        strTotal(0) = strTotal(0) & "," & intCounter
                                    Case Else
                                End Select
                            End If
                        End If
                        '更新合計起始位置
                        If GetValue("全所") = ii And "" & adoaccrpt412.Fields("AccNo") <> "ICS" Then
                            strStartRow = intCounter + 2
                        End If
                    '資料
                    Else
                        
                        If InStr("" & adoaccrpt412.Fields("AccNo"), "－") > 0 Or InStr("" & adoaccrpt412.Fields("AccNo"), "E") > 0 Or _
                            "" & adoaccrpt412.Fields("AccNo") = "" Or "" & adoaccrpt412.Fields("AccNo") = "＝" Then
                            '－/＝不印
                        Else
                            .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00" '數字資料設小數2位
                            strTemp = Val("" & adoaccrpt412.Fields(strFieldTB(ii)))
                        End If
                    End If
                    .Range(Chr(intField + ii) & intCounter).Value = strTemp
            
                Next ii
                'Add by Amy 2020/09/14 +顏色
                If "" & adoaccrpt412.Fields("AccNo") = "ICS" Then
                    .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strFieldN)) & intCounter).Interior.ColorIndex = 20
                ElseIf "" & adoaccrpt412.Fields("AccNo") = "4S" Or "" & adoaccrpt412.Fields("AccNo") = "6S" Then
                    .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strFieldN)) & intCounter).Interior.ColorIndex = 19 '淺黃
                End If
                'end 2020/09/14
            End If
            '總所損益
            If "" & adoaccrpt412.Fields("AccNo") = "ZZZZZZZZ" Then
                'Add by Amy 2020/09/14 +顏色
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strFieldN)) & intCounter).Interior.ColorIndex = 19 '淺黃
                Call SetExcelLine(1, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
                strTotal(0) = intCounter
                '會計科目大項合計
            ElseIf InStr("" & adoaccrpt412.Fields("AccNo"), "S") > 0 Then
                Call SetExcelLine(4, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            Else
                Call SetExcelLine(3, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            End If
            
            adoaccrpt412.MoveNext
            intCounter = intCounter + 1
        Loop
               
        '利潤佔比=全所損益/各部門損益
        intCounter = intCounter + 1
        For ii = LBound(strFieldN) To GetValue("全所")
            strTemp = ""
            If GetValue("會計科目") = ii Then
                strTemp = "利潤佔比"
            '未判斷是否為0程式會錯
            ElseIf Val(.Range(Chr(intField + ii) & strTotal(0)).Value) <> 0 And Val(.Range(Chr(intField + GetValue("全所")) & strTotal(0)).Value) <> 0 Then
                 strTemp = "=" & Chr(intField + ii) & strTotal(0) & "/" & Chr(intField + GetValue("全所")) & strTotal(0)
            End If
            If GetValue("會計科目") <> ii Then
                .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "0.00%;[紅色]-0.00%"
            End If
            .Range(Chr(intField + ii) & intCounter).Value = strTemp
        Next ii
        
        '備註
        intCounter = intCounter + 2
        .Range(Chr(intField) & intCounter).Value = "備註："
        intCounter = intCounter + 1
        '智慧所更名後 "分攤法務部門費用"不顯示
        intSeq = 1
        If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
            .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤法務部門費用: 法務部費用總合＊各該部門當月實際收入／全所實際收入"
            intCounter = intCounter + 1: intSeq = intSeq + 1
        End If
        
        'Modify by Amy 2023/02/21 2021/03/11改公式時,說明未更正
        '.Range(Chr(intField) & intCounter).Value = intSeq & ".分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／全所實際收入"
        .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／(全所實際收入-ＡＣＳ收入-智權部收入-總所/管理收入-其他部門收入負值者)"
        intCounter = intCounter + 1: intSeq = intSeq + 1
        
        strTemp = "全所實際收入－ＦＣＰ收入－ＦＣＴ收入"
        If Val(Mid(MaskEdBox1.Text, 1, 3)) < 105 Then
            strTemp = strTemp & "－投法"
        End If
        'Modify by Amy 2023/02/21 2021/03/11改公式時,說明未更正
        '.Range(Chr(intField) & intCounter).Value = intSeq & ".分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（" & strTemp & "）"
        .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（" & strTemp & "-ＡＣＳ收入-智權部收入-總所/管理收入-其他部門收入負值者）"
        intCounter = intCounter + 1: intSeq = intSeq + 1
        
        If (Val(Replace(MaskEdBox1, "/", "")) < 10908 Or Val(Replace(MaskEdBox2, "/", "")) > 10908) Then
            .Range(Chr(intField) & intCounter).Value = intSeq & ".10908月前ACS收入列於智權部，分攤管理費用也列於智權部，造成各部門營業損益加總不等於全所"
        End If
   End With
   'Excel字型大小設定
   With wksAnnuity.Range(Chr(intField) & "1:" & Chr(UBound(strFieldN) + intField) & intCounter)
        .Font.Name = "新細明體"
        .Font.Size = 10
   End With
   '刪除不需顯示欄位及修改欄位名稱
   Call SetAndDelField(wksAnnuity)
   
   With wksAnnuity
       .PageSetup.PaperSize = 9 '設A4
       .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
       If UBound(strFieldN) > 10 Then
            .PageSetup.Orientation = xlLandscape '橫印
            .PageSetup.Zoom = 100
            .PageSetup.CenterHorizontally = True
            .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.2) '上
            .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.2) '下
            .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0) '左邊界
            .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0) '右邊界
       Else
            .PageSetup.Orientation = xlPortrait '直印
            .PageSetup.Zoom = 80 '縮放比例為80%,列印頁面水平置中
            .PageSetup.CenterHorizontally = True
            
            .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.78) '上
            .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.78) '下
            .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0) '左邊界
            .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0) '右邊界
       End If
   End With
   '產生sheet2抓區間結餘
    intCounter = 1
    Call ExcelSave2(xlsAnnuity, wksAnnuity, bol105YA)
    'Modify by Amy 2020/08/12 原:Chr(GetValue("CFT") + intField)
    With wksAnnuity.Range(Chr(intField) & "1:" & Chr(UBound(strFieldN) + intField) & intCounter)
        .Font.Name = "新細明體"
        .Font.Size = 10
   End With
   With wksAnnuity
       .PageSetup.PaperSize = 9 'A4
       .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
       .PageSetup.Orientation = xlLandscape '橫印
       .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.78) '上
       .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.78) '下
       .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.78) '左邊界
       .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.5) '右邊界
   End With
   '判斷版本
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
   End If
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   Set wksAnnuity = Nothing
   Set xlsAnnuity = Nothing
   MsgBox "檔案已產生~"
   Exit Sub
  
ErrHnd:
   If adoaccrpt412.State = adStateOpen Then adoaccrpt412.Close
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
   End If
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   Set wksAnnuity = Nothing
   Set xlsAnnuity = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Mark by Amy 2020/08/12 +ACS 並改新暫存檔,抓欄位使用別名
'Add by Amy 2015/03/04 產生Excel
Private Sub ExcelSave_Old()
'   Dim xlsAnnuity As New Excel.Application
'   Dim wksAnnuity As New Worksheet
'   Dim strFileName As String, strTemp As String
'   Dim strStartRow As String, strEndRow As String '合計起/迄始位置
'   Dim strTotal(2) As String, strDSum(1) As String '加總列號(0:其他 1:智權部及總所/管理部 2:全所-不計入)/營業收入/支出加總列號
'   Dim strTotPos(1 To 2) As String 'Added by Lydia 2016/01/30 全所損益欄位
'   Dim strVSum(1) '各部門加總列號(0:其他/1:全所) 'Added by Lydia 2016/02/17
'   'Add by Amy 2016/07/27
'   Dim strOSum(1) As String '營業外收入/支出加總列號
'   Dim bol105YA As Boolean '是否為105年後資料
'   Dim strWkName As String 'Add by Amy 2017/09/25 for 2010 工作表名稱為中文
'   'Add by Amy 2020/04/21
'   Dim intSeq As Integer, strAllF As String
'
'   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then bol105YA = True
'
'   'Modify by Amy 2020/04/21 欄位改變動抓,增加L公司調整欄位,SetField設定大小及TB對應欄位
'   'modify by sonia 2016/1/27 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
'   'strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
'   'intWidth = Array(13, 10, 10, 10, 10, 13, 10, 10, 10, 10, 10, 10)
'   If strCmp = "L" Or (Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(智慧所更名日, 6) - 191100) And strCmp = MsgText(601)) Then
'        strAllF = "會計科目,專利,商標,法務,CFP,CFT,FCP,FCT,FCL,CFL,智權部,總所/管理,全所"
'   ElseIf Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(智慧所更名日, 6) - 191100) Then
'        strAllF = "會計科目,專利,商標,法務,CFP,CFT,FCP,FCT,智權部,總所/管理,全所"
'   ElseIf Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'   'end 2020/04/14
'        strAllF = "會計科目,專利,商標,法務,CFP,CFT,FCP,FCT,法務部,智權部,總所/管理,全所"
'   Else
'        strAllF = "會計科目,專利,商標,法務,CFP,CFT,FCP,FCT,投法,智權部,總所/管理,全所"
'   End If
'   strFieldN = Split(strAllF, ",")
'   Call SetField
'   'end 2016/1/27 end
'   'end 2020/04/21
'
'On Error GoTo ErrHnd
'
'   intField = 65:  intCounter = 1
'   strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "-" & Val(Replace(MaskEdBox2.Text, "/", "")) & "部門綜合損益表" & ServerDate & MsgText(43)
'   If Dir(strExcelPath & strFileName) = MsgText(601) Then
'       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'       End If
'   Else
'       Kill strExcelPath & strFileName
'   End If
'
'   xlsAnnuity.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
'   xlsAnnuity.Workbooks.add
'   'Modify by Amy 2017/09/25 for 工作表名稱改為中文
'   If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
'   Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & "1")
'   'end 2017/09/25
'   wksAnnuity.Activate
'   Call SetTitle(wksAnnuity, 1)
'   'end 2017/02/15
'
'   intTitleRow = intCounter: intCounter = intCounter + 1: strStartRow = intCounter
'   With wksAnnuity
'       '列印資料
'       Do While adoaccrpt412.EOF = False
'           If "" & adoaccrpt412.Fields(2) <> "" Then
'               .Range(Chr(intField) & intCounter).Value = adoaccrpt412.Fields(2) '會計科目欄位
'           End If
'
'           If "" & adoaccrpt412.Fields("r41215") = "ZZZZZZZZ" Then
'                   For ii = 1 To UBound(strFieldN)
'                       Select Case ii
'                           'modify by sonia 2016/2/16 +法務部
'                           Case GetValue("智權部"), GetValue("總所/管理"), GetValue("法務部")
'                               'Modify by Amy 2016/07/25 改為營業外收入-營業外支出
'                               'strTemp = Replace(strTotal(1), ",", "," & Chr(intField + ii))
'                               strTemp = Chr(intField + ii) & strOSum(0) & "-" & Chr(intField + ii) & strOSum(1)
'                           Case GetValue("全所")
'                               '排除不計入全所的列(分攤合計列)
'                               strTemp = Replace(Replace(strTotal(0), strTotal(2), ""), ",", "," & Chr(intField + ii))
'                               'Added by Lydia 2016/01/30
'                               strTotPos(1) = Chr(intField + ii): strTotPos(2) = intCounter
'                           Case Else
'                               strTemp = Replace(strTotal(0), ",", "," & Chr(intField + ii))
'                       End Select
'                       'Modify by Amy 2016/07/27 改為營業外收入-營業外支出
'                       If GetValue("智權部") <> ii And GetValue("總所/管理") <> ii And GetValue("法務部") <> ii Then
'                          strTemp = "sum(" & Right(Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)), Len(strTemp) - 1) & ")"
'                       End If
'                       .Range(Chr(intField + ii) & intCounter).Formula = "=" & strTemp
'                       'end 2016/07/27
'                   Next ii
'                   'Add by Amy 2016/07/27 +框線
'                   Call SetExcelLine(1, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
'           ElseIf InStr("" & adoaccrpt412.Fields("r41215"), "S") > 0 Then
'                   'Add by Amy 2016/07/27 +框線
'                   If "" & adoaccrpt412.Fields("r41215") = "DS" Or "" & adoaccrpt412.Fields("r41215") = "VS" Then
'                        Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter - 1)
'                   Else
'                        If strStartRow = intTitleRow + 1 Then
'                            Call SetExcelLine(2, wksAnnuity, Chr(intField) & strStartRow & ":" & Chr(UBound(strFieldN) + intField) & strEndRow)
'                        Else
'                            Call SetExcelLine(2, wksAnnuity, Chr(intField) & strStartRow - 1 & ":" & Chr(UBound(strFieldN) + intField) & strEndRow)
'                        End If
'                   End If
'
'                   '*** 合計
'                   For ii = 1 To UBound(strFieldN)
'                       .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
'                       If "" & adoaccrpt412.Fields("r41215") = "DS" Then
'                           .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strDSum(0) & "-" & Chr(intField + ii) & strDSum(1)
'                       'Added by Lydia 2016/02/17 各部門營業損益
'                       ElseIf "" & adoaccrpt412.Fields("r41215") = "VS" Then
'                           'modify by sonia 2016/7/13 法務部也是0
'                           'If GetValue("智權部") = ii Or GetValue("總所/管理") = ii Then
'                           If GetValue("智權部") = ii Or GetValue("總所/管理") = ii Or GetValue("法務部") = ii Then
'                               .Range(Chr(intField + ii) & intCounter).Value = 0
'                           ElseIf GetValue("全所") = ii Then
'                               .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strVSum(1)
'                           Else
'                               strTemp = Replace(strVSum(0), ",", "," & Chr(intField + ii))
'                               .Range(Chr(intField + ii) & intCounter).Formula = "=sum(" & Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)) & ")"
'                           End If
'                       'end 2016/02/17
'                       Else
'                           .Range(Chr(intField + ii) & intCounter).Formula = "=sum(" & Chr(intField + ii) & strStartRow & ":" & Chr(intField + ii) & strEndRow & ")"
'                       End If
'                   Next ii
'                   'Add by Amy 2016/07/27 +框線
'                   Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
'
'                   'Added by Lydia 2016/02/17 + VSum
'                   'strDSum():部門損益 / strTotal():全所損益 計算欄位 / strVSum():各部門營業損益
'                   strTemp = Left("" & adoaccrpt412.Fields("r41215"), 1)
'                   Select Case strTemp
'                       Case "4" '營業收入
'                           strDSum(0) = intCounter
'                       Case "6" '營業支出
'                            strDSum(1) = intCounter
'                       Case "7" '營業外收支
'                           If Left(adoaccrpt412.Fields("r41215"), 2) = "71" Then
'                               strOSum(0) = intCounter 'Add by Amy 2016/07/27
'                               strTotal(0) = strTotal(0) & "," & intCounter
'                               strTotal(1) = strTotal(1) & "," & intCounter
'                           Else
'                               strOSum(1) = intCounter 'Add by Amy 2016/07/27
'                               strTotal(0) = strTotal(0) & ",-" & intCounter
'                               strTotal(1) = strTotal(1) & ",-" & intCounter
'                           End If
'                       Case "9" '分攤費用
'                           strVSum(0) = strVSum(0) & ",-" & intCounter 'Added by Lydia 2016/02/17
'                           strTotal(0) = strTotal(0) & ",-" & intCounter
'                           strTotal(2) = ",-" & intCounter
'                       Case "D" '部門損益
'                           'Added by Lydia 2016/02/17
'                           strVSum(0) = strVSum(0) & "," & intCounter
'                           strVSum(1) = intCounter
'                           'end 2016/02/17
'                           strTotal(0) = strTotal(0) & "," & intCounter
'                       Case Else
'                   End Select
'            'Add by Amy 2016/07/25 +判斷為分攤科目改為公式顯示-婧瑄
'            ElseIf "" & adoaccrpt412.Fields("r41215") = "9997" Or "" & adoaccrpt412.Fields("r41215") = "9998" Or "" & adoaccrpt412.Fields("r41215") = "9999" Then
'                For ii = UBound(strFieldN) To 1 Step -1
'                    .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
'                    If GetValue("全所") = ii Then
'                        Select Case "" & adoaccrpt412.Fields("r41215")
'                            Case "9997"
'                                strTemp = "=" & Chr(intField + GetValue("法務部")) & strDSum(1)
'                            Case "9998"
'                                strTemp = "=" & Chr(intField + GetValue("總所/管理")) & strDSum(1)
'                            Case Else
'                                strTemp = "=" & Chr(intField + GetValue("智權部")) & strDSum(1)
'                        End Select
'                    ElseIf GetValue("專利") = ii Then
'                        '與其他欄位一樣用算的再加總全所會造成小數位與XX部門費用合計不符
'                        strTemp = "=Round(" & Chr(intField + GetValue("全所")) & intCounter & "-Sum(" & Chr(intField + GetValue("商標")) & intCounter & ":" & Chr(intField + GetValue("總所/管理")) & intCounter & "),2)"
'                    ElseIf "" & adoaccrpt412.Fields("r41215") = "9997" Then
'                        strTemp = "=Round(" & Chr(intField + GetValue("法務部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & Chr(intField + GetValue("全所")) & strDSum(0) & "),2)"
'                    ElseIf "" & adoaccrpt412.Fields("r41215") = "9998" Then
'                        strTemp = "=Round(" & Chr(intField + GetValue("總所/管理")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & Chr(intField + GetValue("全所")) & strDSum(0) & "),2)"
'                    Else
'                        If GetValue("FCP") = ii Or GetValue("FCT") = ii Or GetValue("法務部") = ii Or GetValue("智權部") = ii Or GetValue("總所/管理") = ii Then
'                            strTemp = "0"
'                        Else
'                            strTemp = "=Round(" & Chr(intField + GetValue("智權部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/(" & Chr(intField + GetValue("全所")) & strDSum(0) & "-" & _
'                                                Chr(intField + GetValue("FCP")) & strDSum(0) & "-" & Chr(intField + GetValue("FCT")) & strDSum(0) & IIf(bol105YA = True, "", "-" & Chr(intField + GetValue("投法")) & strDSum(0)) & ")),2)"
'                        End If
'                    End If
'                    If InStr(strTemp, "=") > 0 Then
'                        .Range(Chr(intField + ii) & intCounter).Formula = strTemp
'                    Else
'                        .Range(Chr(intField + ii) & intCounter).Value = Val(strTemp)
'                    End If
'                Next ii
'           Else
'               '*** 資料
'               'Add by Amy 2016/07/25
'               If "" & adoaccrpt412.Fields("r41215") = "" Or InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Then
'                    intCounter = intCounter - 1
'               End If
'
'               For ii = 1 To UBound(strFieldN)
'                   If InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Then strEndRow = intCounter 'Modify by Amy 2016/07/25 原:- 1 '更新合計結束位置
'                   If InStr("" & adoaccrpt412.Fields("r41215"), "E") > 0 Then strStartRow = intCounter + 1 '更新合計起始位置
'                   If InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Or InStr("" & adoaccrpt412.Fields("r41215"), "E") > 0 Or _
'                       "" & adoaccrpt412.Fields("r41215") = "" Or "" & adoaccrpt412.Fields("r41215") = "＝" Then
'                       'Mark by Amy 2016/07/25 －/＝不印
'                       '.Range(Chr(intField + ii) & intCounter).Value = adoaccrpt412.Fields(ii + 2)
'                   Else
'                       .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00" '數字資料設小數2位
'                       '資料
'                       'Modify by Amy 2020/04/21 抓strFieldTB設定
''                       Select Case ii
''                           Case GetValue("總所/管理")
''                               .Range(Chr(intField + ii) & intCounter).Value = Val("" & adoaccrpt412.Fields(ii + 3))
''                           Case GetValue("全所")
''                               .Range(Chr(intField + ii) & intCounter).Value = Val("" & adoaccrpt412.Fields(ii + 1))
''                           Case Else
''                               .Range(Chr(intField + ii) & intCounter).Value = Val("" & adoaccrpt412.Fields(ii + 2))
''                       End Select
'                       .Range(Chr(intField + ii) & intCounter).Value = Val("" & adoaccrpt412.Fields(strFieldTB(ii)))
'                   End If
'               Next ii
'           End If
'           adoaccrpt412.MoveNext
'           intCounter = intCounter + 1
'       Loop
'      'Added by Lydia 2016/01/30 +利潤率(部門損益/全所損益)
'       .Range(Chr(intField) & intCounter).Value = "利潤率"
'        For ii = 1 To UBound(strFieldN)
'            'Modify by Amy 2020/04/21 條件起迄若含智慧所更名日前資料,以舊格式顯示 ex:10903~10904/條件下L公司 顯示FCL,CFL欄
'            If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
'               strExc(0) = "法務部,智權部,總所/管理,全所"
'            Else
'               strExc(0) = "智權部,總所/管理,全所"
'            End If
'            If InStr(strExc(0), strFieldN(ii)) = 0 Then
'               .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strTotPos(2) & "/$" & strTotPos(1) & "$" & strTotPos(2)
'               .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00%"
'            End If
'        Next ii
'        intCounter = intCounter + 2
'        'end 2016/01/30
'        'Add by Amy 2016/07/27 +備註
'        .Range(Chr(intField) & intCounter).Value = "備註："
'        intCounter = intCounter + 1
'        'Modiby by Amy 2020/04/21 +智慧所更後項目"分攤法務部門費用"不顯示
'        intSeq = 1
'        If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 And Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(智慧所更名日, 6) - 191100) And strCmp <> "L" Then
'            .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤法務部門費用: 法務部費用總合＊各該部門當月實際收入／全所實際收入"
'            intCounter = intCounter + 1: intSeq = intSeq + 1
'        End If
'        .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／全所實際收入"
'        intCounter = intCounter + 1: intSeq = intSeq + 1
'        .Range(Chr(intField) & intCounter).Value = intSeq & ".分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（全所實際收入－ＦＣＰ收入－ＦＣＴ收入）"
'   End With
'   'Add by Amy 2016/07/27 Excel字型大小設定
'   With wksAnnuity.Range(Chr(intField) & "1:" & Chr(UBound(strFieldN) + intField) & intCounter)
'        .Font.Name = "新細明體"
'        .Font.Size = 10
'   End With
'   With wksAnnuity
'       .PageSetup.PaperSize = 9 'Add by Amy 2016/07/27 設A4
'       .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
'       'Modify by Amy 2020/04/21 欄位太多無法直印,改橫印
'       If UBound(strFieldN) > 10 Then
'            .PageSetup.Orientation = xlLandscape '橫印
'            .PageSetup.Zoom = 100
'            .PageSetup.CenterHorizontally = True
'            .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.2) '上
'            .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.2) '下
'            .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0) '左邊界
'            .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0) '右邊界
'       Else
'            'Modified by Lydia 2017/06/06 橫印改直印 xlLandscape => xlPortrait
'            '.PageSetup.Orientation = xlLandscape '橫印
'            .PageSetup.Orientation = xlPortrait '直印
'            'Added by Lydia 2017/06/06 縮放比例為80%,列印頁面水平置中
'            .PageSetup.Zoom = 80
'            .PageSetup.CenterHorizontally = True
'            'end 2017/06/06
'
'            .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.78) '上
'            .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.78) '下
'            'Modified by Lydia 2017/06/06 為了印整張A4,左右邊界改為0
'            '.PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.78) '左邊界
'            '.PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.5) '右邊界
'            .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0) '左邊界
'            .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0) '右邊界
'       End If
'   End With
'   'Add by Amy 2017/02/15 加產生sheet2抓區間結餘
'    intCounter = 1
'    Call ExcelSave2(xlsAnnuity, wksAnnuity, bol105YA)
'    With wksAnnuity.Range(Chr(intField) & "1:" & Chr(GetValue("CFT") + intField) & intCounter)
'        .Font.Name = "新細明體"
'        .Font.Size = 10
'   End With
'   With wksAnnuity
'       .PageSetup.PaperSize = 9 'A4
'       .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
'       .PageSetup.Orientation = xlLandscape '橫印
'       .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.78) '上
'       .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.78) '下
'       .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.78) '左邊界
'       .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.5) '右邊界
'   End With
'    'end 2017/02/15
'   'Modify by Amy 2016/06/23 +判斷版本
'   If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
'   Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
'   End If
'   'end 2016/06/23
'   xlsAnnuity.Workbooks.Close
'   xlsAnnuity.Quit
'   Set wksAnnuity = Nothing
'   Set xlsAnnuity = Nothing
'   MsgBox "檔案已產生~"
'   Exit Sub
'
'ErrHnd:
'   If adoaccrpt412.State = adStateOpen Then adoaccrpt412.Close
'   'Modify by Amy 2016/06/23 +判斷版本
'   If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
'   Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
'   End If
'   'end 2016/06/23
'   xlsAnnuity.Workbooks.Close
'   xlsAnnuity.Quit
'   Set wksAnnuity = Nothing
'   Set xlsAnnuity = Nothing
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub
'end 2020/08/12

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
    
    For jj = 1 To UBound(strFieldN)
        If UCase(strFieldN(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2015/03/04

'Mark by Amy 2020/04/21 不使用
'Added by Lydia 2016/01/30 從AccReport改成Printer
Private Sub PrintData()
'Dim strRBase As String '全所損益
'Dim strRD(0 To 10) As String
'Dim ii As Integer
'    Printer.EndDoc
'    Printer.Orientation = 1 '1.直印 2.橫印
'    Printer.PaperSize = PUB_GetPaperSize(15) '美國標準
'
'    lngPageHeight = Printer.ScaleHeight
'    lngPageWidth = Printer.ScaleWidth
'    lngLineHeight = 300
'
'    iPage = 0
'    GetPleft
'    Erase strRD
'
'    PrintHeader '列印表頭
'    With adoaccrpt412
'        Do While Not .EOF
'        '列印明細
'           iPage = iPage + 1
'           strTemp(0) = "" & .Fields("R41203") '會計科目
'           strTemp(1) = "" & .Fields("R41204") '專利
'           strTemp(2) = "" & .Fields("R41205") '商標
'           strTemp(3) = "" & .Fields("R41206") '法務->105年以後併入"法務部"
'           strTemp(4) = "" & .Fields("R41207") 'CFP
'           strTemp(5) = "" & .Fields("R41208") 'CFT
'           strTemp(6) = "" & .Fields("R41209") 'FCP
'           strTemp(7) = "" & .Fields("R41210") 'FCT
'           strTemp(8) = "" & .Fields("R41211") '投法->105年以後"法務部"
'           strTemp(9) = "" & .Fields("R41212") '智權部
'           strTemp(10) = "" & .Fields("R41214") '總所/管理
'           strTemp(11) = "" & .Fields("R41213") '全所
'           If .Fields("R41215") = "ZZZZZZZZ" Then
'              strRBase = strTemp(11)
'              strRD(0) = "利潤率"
'              If Val(strRBase) <> 0 Then
'                 For ii = 1 To 10
'                    strRD(ii) = Format(Val(strTemp(ii)) / Val(strRBase), "##0.00%")
'                 Next
'              End If
'           End If
'           For ii = 0 To UBound(strFieldN)
'              If intWidth(ii) > 0 Then
'                 If strTemp(0) = "" And InStr(strTemp(1), "－") = 0 And InStr(strTemp(1), "＝") = 0 And Val(strTemp(11)) = 0 Then
'                    '空一行
'                    Exit For
'                 Else
'                    '靠左
'                    If ii = 0 Or InStr(strTemp(ii), "－") > 0 Or InStr(strTemp(ii), "＝") > 0 Then
'                        Printer.CurrentX = PLeft(ii) + 50
'                        Printer.CurrentY = iPrint
'                        If ii < UBound(strFieldN) Then
'                           Printer.Print strTemp(ii)
'                        Else
'                            If InStr(strTemp(ii), "－") > 0 Then
'                               Printer.Print String(7, "－")
'                            ElseIf InStr(strTemp(ii), "＝") > 0 Then
'                               Printer.Print String(7, "＝")
'                            End If
'                        End If
'                    '靠右
'                    Else
'                        Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(Format(Val(strTemp(ii)), "###,##0.00")) - ciColGap
'                        Printer.CurrentY = iPrint
'                        Printer.Print Format(Val(strTemp(ii)), "###,##0.00")
'                    End If
'                 End If
'              End If
'           Next
'           PrintNewLine
'
'           .MoveNext
'        Loop
'    End With
'
'    '利潤率
'    For ii = 0 To 10
'       If intWidth(ii) > 0 Then
'          '靠左
'          If ii = 0 Then
'              Printer.CurrentX = PLeft(ii)
'              Printer.CurrentY = iPrint
'              Printer.Print strRD(ii)
'          '靠右
'          Else
'              If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'                 strExc(0) = "法務部,智權部,總所/管理,全所"
'              Else
'                 strExc(0) = "智權部,總所/管理,全所"
'              End If
'              If InStr(strExc(0), strFieldN(ii)) = 0 Then
'                 Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(strRD(ii)) - ciColGap
'                 Printer.CurrentY = iPrint
'                 Printer.Print strRD(ii)
'              End If
'          End If
'       End If
'    Next
''Add by Amy 2016/07/27最後備註
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print "備註："
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print "1.分攤法務部門費用: 法務部費用總合＊各該部門當月實際收入／全所實際收入"
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print "2.分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／全所實際收入"
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print "3.分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（全所實際收入－ＦＣＰ收入－ＦＣＴ收入）"
''end 2016/07/27
'
'Printer.EndDoc
'ShowPrintOk

End Sub


Private Sub GetPleft() '明細表邊界
Dim inX As Integer

'105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "法務部", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 0, 9, 9, 9, 9, 9, 9, 9, 10)
Else
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 10)
End If
   
Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLeft
Erase PTitle
  
   PLeft(0) = ciStartX
   For inX = 1 To UBound(strFieldN)
       If intWidth(inX - 1) = 0 Then
           PLeft(inX) = PLeft(inX - 1)
       Else
           PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(intWidth(inX - 1), "A")) + ciColGap
       End If
   Next
   PLeft(12) = PLeft(11) + Printer.TextWidth(String(10, "A")) + ciColGap
End Sub
Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader

   End If
End Sub

Private Sub PrintHeader()
Dim strPTmp As String
Dim pa1 As Integer
Dim ii As Integer
iPrint = ciStartY
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False

'報表抬頭
strPTmp = ReportTitle(416)
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strPTmp

PrintNewLine
PrintNewLine

Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

'Modify by Amy 2020/04/21 公司別改抓變數
'strPTmp = "公司別：" & IIf(Text5 = "2", "J", Text5) & " " & IIf(Text6 = "", "台一　專利商標/智權", Text6)
strPTmp = "公司別：" & strCmp & " " & strCmpN
'end 2020/04/21
pa1 = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print strPTmp

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & CFDate(strSrvDate(2))

PrintNewLine

Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print "年　月：" & MaskEdBox1.Text & " ～ " & MaskEdBox2.Text

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

PrintNewLine

Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName

PrintNewLine

For ii = 0 To UBound(strFieldN)
   '顯示／欄位
   If intWidth(ii) > 0 And strFieldN(ii) <> "" Then
       strPTmp = strFieldN(ii)
       Printer.CurrentX = PLeft(ii) + ((PLeft(ii + 1) - PLeft(ii) - Printer.TextWidth(strPTmp)) / 2) - ciColGap
       Printer.CurrentY = iPrint
       Printer.Print strPTmp
   End If
Next

PrintNewLine

PrintLine

End Sub

Private Sub PrintLine()
   Printer.Line (PLeft(0) - 50, iPrint)-(PLeft(12) + 50, iPrint)
   iPrint = iPrint + 150
End Sub
'end 2016/01/30

'Added by Lydia 2016/02/17
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox1.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox1.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox1.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox1.SetFocus
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/02/17
Private Sub MaskEdBox2_LostFocus()
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox2.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox2.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox2.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox2.SetFocus
         End If
      End If
   End If
End Sub

'Mark by Amy 2020/08/12 Excel改公司後可不使用
'Add by Amy 2016/07/25 +更新分攤費用值(改成公式計算)-婧瑄
Private Sub UpdAccrpt412_Old()
'    Dim RsQ As New ADODB.Recordset
'    Dim strQ As String, strUpd As String
'    Dim i As Integer, intQ As Integer
'    Dim strVal(10) As String, strE As String '更新值(10:合計)/相對費用值
'    Dim bol105YA As Boolean '是否為105年後資料
'
'    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then bol105YA = True
'
'    strQ = "Select * From accrpt412 Where ID='" & strUserNum & "' And R41215 in ('9997','9998','9999') Order by R41215"
'    If adoaccrpt412.State <> adStateClosed Then adoaccrpt412.Close
'    adoaccrpt412.CursorLocation = adUseClient
'    adoaccrpt412.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    '抓取相關資料
'    strQ = "Select * From " & _
'            "(Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP,R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total " & _
'             "From accrpt412 Where ID='" & strUserNum & "' And R41215='4S')," & _
'            "(Select R41211 as LE,R41214 as ME,R41212 as SE From accrpt412 Where ID='" & strUserNum & "' And R41215='6S' )"
'    If RsQ.State <> adStateClosed Then RsQ.Close
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoaccrpt412.RecordCount > 0 And RsQ.RecordCount > 0 Then
'      With adoaccrpt412
'        .MoveFirst
'        Do While Not .EOF
'            Select Case "" & .Fields("R41215")
'                Case "9997"
'                    strE = Val("" & RsQ.Fields("LE"))
'                Case "9998"
'                    strE = Val("" & RsQ.Fields("ME"))
'                Case "9999"
'                    strE = Val("" & RsQ.Fields("SE"))
'            End Select
'            For i = 1 To 9
'               If "" & .Fields("R41215") = "9999" Then
'                    If bol105YA = False Then
'                        If i + 4 >= 9 And i + 4 <= 14 And i + 4 <> 11 Then
'                            strVal(i) = "0"
'                        Else
'                            '105年以前需剔除「投法」
'                            strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / (Val(RsQ.Fields("Total")) - Val(RsQ.Fields("FCP")) - Val(RsQ.Fields("FCT")) - Val(RsQ.Fields("Law")))), 2), FAmount)
'                        End If
'                    Else
'                        If i + 4 >= 9 And i + 4 <= 14 Then
'                            strVal(i) = "0"
'                        Else
'                            strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / (Val(RsQ.Fields("Total")) - Val(RsQ.Fields("FCP")) - Val(RsQ.Fields("FCT")))), 2), FAmount)
'                        End If
'                    End If
'               Else
'                    strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / Val(RsQ.Fields("Total"))), 2), FAmount)
'               End If
'                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
'                strVal(0) = Format(Round(Val(strVal(0)) + Val(strVal(i)), 2), FAmount)
'            Next i
'            strVal(0) = Format(Round(Val(strE) - Val(strVal(0)), 2), FAmount)
'            '更新
'            If strUpd <> MsgText(601) Then
'                strUpd = "Update Accrpt412 Set R41213='" & strE & "',R41204='" & strVal(0) & "'" & strUpd & " Where ID='" & strUserNum & "' And R41215='" & .Fields("R41215") & "'"
'                cnnConnection.Execute strUpd
'                strUpd = ""
'            End If
'            strVal(0) = ""
'             .MoveNext
'        Loop
'       End With
'    End If
'    adoaccrpt412.Close
'    RsQ.Close
'
'     '更新 分攤費用(9S)
'    For i = 0 To 10
'        strVal(i) = ""
'    Next i
'    '抓取相關資料
'    strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
'               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total " & _
'             "From accrpt412 Where ID='" & strUserNum & "' And R41215 in ('9997','9998','9999')"
'    If RsQ.State <> adStateClosed Then RsQ.Close
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    If RsQ.RecordCount > 0 Then
'      strUpd = ""
'      With RsQ
'        .MoveFirst
'        Do While Not .EOF
'            For i = 0 To 9
'                strVal(i) = Format(Round(Val(strVal(i)) + Val("" & .Fields(i)), 2), FAmount)
'            Next i
'            strVal(10) = Format(Round(Val(strVal(10)) + Val("" & .Fields("Total")), 2), FAmount)
'            .MoveNext
'        Loop
'        If strVal(10) <> MsgText(601) Then
'            For i = 0 To 9
'                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
'            Next i
'            If strUpd <> MsgText(601) Then
'                strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where ID='" & strUserNum & "' And R41215='9S' "
'                cnnConnection.Execute strUpd
'            End If
'        End If
'      End With
'    End If
'    RsQ.Close
'    '更新 各部門營業損益(VS)
'     strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
'               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total From accrpt412 Where ID='" & strUserNum & "' And R41215 ='DS' "
'    If RsQ.State <> adStateClosed Then RsQ.Close
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    If RsQ.RecordCount > 0 Then
'      strUpd = ""
'      With RsQ
'            .MoveFirst
'            Do While Not .EOF
'                For i = 0 To 9
'                    If bol105YA = False Then
'                        If i >= 8 And i <= 9 Then
'                            strVal(i) = "0"
'                        Else
'                            strVal(i) = Format(Round(Val("" & .Fields(i)) - Val(strVal(i)), 2), FAmount)
'                        End If
'                    Else
'                        If i >= 7 And i <= 9 Then
'                            strVal(i) = "0"
'                        Else
'                            strVal(i) = Format(Round(Val("" & .Fields(i)) - Val(strVal(i)), 2), FAmount)
'                        End If
'                    End If
'                Next i
'                strVal(10) = Val("" & .Fields("Total"))
'                .MoveNext
'            Loop
'            For i = 0 To 9
'                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
'            Next i
'            strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where ID='" & strUserNum & "' And R41215='VS' "
'            cnnConnection.Execute strUpd
'      End With
'    End If
'    '更新 全所損益(ZZZZZZZZ)
'     strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
'               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total,R41215 From accrpt412 " & _
'               "Where ID='" & strUserNum & "' And R41215 in ('71S','72S') Order by R41215"
'    If RsQ.State <> adStateClosed Then RsQ.Close
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    If RsQ.RecordCount > 0 Then
'      strUpd = ""
'      With RsQ
'            .MoveFirst
'            Do While Not .EOF
'                For i = 0 To 10
'                    If .Fields("R41215") = "71S" Then
'                        '+ 營業外收入
'                        strVal(i) = Format(Round(Val(strVal(i)) + Val("" & .Fields(i)), 2), FAmount)
'                    Else
'                        '- 營業外支出
'                        strVal(i) = Format(Round(Val(strVal(i)) - Val("" & .Fields(i)), 2), FAmount)
'                    End If
'                Next i
'                .MoveNext
'            Loop
'            For i = 0 To 9
'                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
'            Next i
'            strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where ID='" & strUserNum & "' And R41215='ZZZZZZZZ' "
'            cnnConnection.Execute strUpd
'      End With
'    End If
End Sub

'增加框線設定-婉莘
Private Sub SetExcelLine(intChoose As Integer, ByRef m_Xls As Worksheet, strField As String)

    With m_Xls.Range(strField)
        Select Case intChoose
            Case 0 '抬頭/合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 1 '最後合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThick
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlThick '粗線
             Case 2 '合計
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
            'Add by Amy 2020/08/12
            Case 3 '資料內容
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlHairline '虛線
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 4 '表1 大項 合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin '細線
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End Select
    End With
End Sub
'end 2016/07/27

'Add by Amy 2017/02/15
'產生區間結餘 Excel
Private Sub ExcelSave2(ByRef xlsAnnuity As Excel.Application, ByRef wksAnnuity As Worksheet, ByVal bol105YA As Boolean)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strValue As String
    Dim intQ As Integer
    Dim bolFormat As Boolean
    Dim strWkName As String 'Add by Amy 2017/09/25 for 2010 工作表名稱為中文
    Dim strQ2 As String, strWhere As String, strL As String, strOL As String, strAllF As String 'Add by Amy 2020/08/12
        
    'Modify by Amy 2020/04/21 +公司別條件
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strQ = "And a0201 In('" & Replace(strCmp, "+", "','") & "') "
        Else
            strQ = "And a0201='" & strCmp & "' "
        End If
    End If
    'Modify by Amy 2020/08/12
    strWhere = "And ax201(+) = a0201 And ax202(+) = a0202 And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
                    " And a0205>=" & Val(Mid(MaskEdBox1, 1, 3) & Mid(MaskEdBox1, 5, 2)) & "01" & _
                    " And a0205<=" & Val(Mid(MaskEdBox2, 1, 3) & Mid(MaskEdBox2, 5, 2)) & "31" & _
                    " And SubStr(ax205,1,1)='4' And InStr(ax213||' ','結餘')>0 And InStr(ax212,'轉撥')=0" & _
                    " And SubStr(ax205,1,4)=a0101(+) " & strQ
    '畫面條件起月是 105年後且公司別為空或L法務資料顯示於L欄
    If Val(Replace(MaskEdBox1, "/", "")) >= 10501 And (strCmp = "L" Or strCmp = MsgText(601)) Then
        strQ2 = " Union Select a0102,'L' as ax204,Sum(ax207-ax206) as Amt,SubStr(ax205,1,4) as AccNo From acc021,acc020,acc010 " & _
                     "Where InStr(ax204,'L')>0 And ax204<>'SAL' " & strWhere & " Group by SubStr(ax205,1,4),a0102,ax204 "
        strOL = ",L"
    '畫面條件起月是 105年後且智慧所更名日前資料且公司別不為空或L顯示於法務部欄
    ElseIf Val(Replace(MaskEdBox1, "/", "")) >= 10501 And Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100 Then
        strQ2 = " Union Select a0102,'法務部' as ax204,Sum(ax207-ax206) as Amt,SubStr(ax205,1,4) as AccNo From acc021,acc020,acc010 " & _
                     "Where InStr(ax204,'L')>0 And ax204<>'SAL' " & strWhere & " Group by SubStr(ax205,1,4),a0102,ax204 "
        strOL = ",法務部"
    '畫面條件起月是 105前L部門顯示法務,其他法務部門顯示投法
    ElseIf Val(Replace(MaskEdBox1, "/", "")) < 10501 Then
        strQ2 = " Union Select a0102,'法務' as ax204,Sum(ax207-ax206) as Amt,SubStr(ax205,1,4) as AccNo From acc021,acc020,acc010 " & _
                    "Where ax204='L' " & strWhere & " Group by SubStr(ax205,1,4),a0102,ax204 " & _
                    " Union Select a0102,'投法' as ax204,Sum(ax207-ax206) as Amt,SubStr(ax205,1,4) as AccNo From acc021,acc020,acc010 " & _
                    "Where ax204 In ('CFL','FCL') " & strWhere & " Group by SubStr(ax205,1,4),a0102,ax204 "
        strL = ",法務": strOL = ",投法"
    End If
    strQ = "Select  a0102,ax204,Sum(ax207-ax206) as Amt,SubStr(ax205,1,4) as AccNo From acc021,acc020,acc010 " & _
              "Where ax204 Not in('L','CFL','CFL') " & strWhere & " Group by SubStr(ax205,1,4),a0102,ax204 " & strQ2 & _
              "Order by AccNo,ax204"
    
    strAllF = "會計科目,專利,商標" & strL & ",CFP,CFT" & strOL
    strFieldN = Split(strAllF, ",")
    
    ReDim strFieldTB(UBound(strFieldN))
    'end 2020/08/12
    'end 2020/04/21
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        'Modify by Amy 2017/09/25 for 工作表名稱改為中文
        If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
        Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & "2")
        'end 2017/09/25
        wksAnnuity.Activate
        Call SetTitle(wksAnnuity, 2)
        intTitleRow = intCounter: intCounter = intCounter + 1
        
        With wksAnnuity
            Do While RsQ.EOF = False
                'Modify by Amy 2020/08/12 原: ii = 0 To GetValue("CFT")
                For ii = LBound(strFieldN) To UBound(strFieldN)
                    bolFormat = True
                    If ii = GetValue("會計科目") Then
                        bolFormat = False
                        strValue = "" & RsQ.Fields("a0102")
                    ElseIf ii = GetValue(Left(RsQ.Fields("a0102"), 2)) Or ii = GetValue(Left(RsQ.Fields("a0102"), 3)) Then
                        strValue = RsQ.Fields("Amt")
                    Else
                        strValue = "0"
                    End If
                    .Range(Chr(ii + intField) & intCounter).Value = strValue
                    If bolFormat = True Then
                        .Range(Chr(ii + intField) & intCounter).NumberFormatLocal = "#,##0"
                    End If
                Next ii
                RsQ.MoveNext
                intCounter = intCounter + 1
            Loop
            Call SetExcelLine(2, wksAnnuity, Chr(intField) & intTitleRow + 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            intCounter = intCounter + 1
            For ii = LBound(strFieldN) To UBound(strFieldN)
                If ii = 0 Then
                    .Range(Chr(ii + intField) & intCounter).Value = "合　　計"
                Else
                    .Range(Chr(ii + intField) & intCounter).Formula = "=Sum(" & Chr(ii + intField) & intTitleRow + 1 & ":" & _
                                                                                                        Chr(ii + intField) & intCounter - 1 & ")"
                End If
            Next ii
            Call SetExcelLine(1, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            'end 2020/08/12
        End With
    End If
End Sub

'原程式寫成Sub
Private Sub SetTitle(ByRef wksAnnuity As Worksheet, ByVal intChoose As Integer)
    Dim strField As String
    
    'Modify by Amy 2020/08/12 GetValue(,+1)
    If intChoose = 1 Then
        strField = Chr(Fix(UBound(strFieldN) / 2) + intField)
    Else
        strField = Chr(Fix(GetValue("CFT") / 2) + intField)
    End If
    
    With wksAnnuity
        '***表頭設定***
        .Range(strField & intCounter).Value = "部門綜合損益表 " & IIf(intChoose = 2, "- 結餘", "")
        .Range(strField & intCounter).HorizontalAlignment = xlCenter
        .Range(strField & intCounter).VerticalAlignment = xlCenter
        intCounter = intCounter + 1
        
        'Modify by Amy 2020/04/21 公司別改抓變數 原:IIf(Text6 = "", "台一　專利商標/智權", Text6)
        If intChoose = 1 Then
            .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "公司別："
            .Range(strField & intCounter).Value = strCmp & " " & strCmpN
        Else
             .Range(strField & intCounter).Value = "公司別：" & strCmp & " " & strCmpN
        End If
        'end 2020/04/21
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        If intChoose = 1 Then
            .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "年　月："
            .Range(strField & intCounter).Value = MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
        Else
            .Range(strField & intCounter).Value = "年　月：" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
        End If
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "列印人員：" & strUserName
        intCounter = intCounter + 1
       
        If intChoose = 1 Then
            For ii = LBound(strFieldN) To UBound(strFieldN)
                .Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = intWidth(ii)
                .Range(Chr(intField + ii) & intCounter).Value = strFieldN(ii)
                .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
            Next ii
            'Add by Amy 2016/07/27 +框線
            Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
        Else
            'Modify by Amy 2020/08/12 原: ii = 0 To GetValue("CFT")
            For ii = LBound(strFieldN) To UBound(strFieldN)
                .Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = 13
                .Range(Chr(intField + ii) & intCounter).Value = strFieldN(ii)
                .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
            Next ii
            Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            'end 2020/08/12
        End If
    End With
End Sub
'end 2017/02/15

'Modify by Amy 2020/08/12
Private Sub SetField(ByRef stQ2 As String)
    Dim ii As Integer, intW As Integer
    Dim strTBF As String, strAllF As String
    
    stQ2 = ""
    
    strAllF = "會計科目,專利,商標,L,CFP,CFT,FCP,FCT,ACS,投法,法務部,智權部,管理,全所,差額"
    strFieldN = Split(strAllF, ",")
    
    ReDim strFieldTB(UBound(strFieldN))
    ReDim intWidth(UBound(strFieldN))
    
    For ii = LBound(strFieldN) To UBound(strFieldN)
        intW = 0
        strTBF = ""
        Select Case strFieldN(ii)
            Case "會計科目"
                strTBF = "R003"
                intW = 13
            Case "專利"
                 strTBF = "R011"
                intW = 10
            Case "商標"
                 strTBF = "R012"
                intW = 10
            Case "L"
                strTBF = "R013"
                intW = 10
            Case "CFP"
                strTBF = "R014"
                intW = 11
            Case "CFT"
                strTBF = "R015"
                intW = 10
            Case "FCP"
                strTBF = "R016"
                intW = 11
            Case "FCT"
                strTBF = "R017"
                intW = 10
            Case "投法"
                strTBF = "R018"
                intW = 10
            Case "法務部"
                strTBF = "R019"
                intW = 10
            Case "ACS"
                strTBF = "R020"
                intW = 10
            Case "智權部"
                strTBF = "R021"
                intW = 10
            Case "管理"
                strTBF = "R022"
                intW = 10
            Case "全所"
                strTBF = "R023"
                intW = 13
            Case "差額"
                intW = 10
        End Select
        strFieldTB(ii) = strTBF
        intWidth(ii) = intW
        '差額為公式不需抓
        If strFieldN(ii) <> "差額" Then stQ2 = stQ2 & "," & strTBF
    Next ii
    If stQ2 <> MsgText(601) Then stQ2 = Mid(stQ2, 2) & ",ID,R001,R002"
End Sub

'Add by Amy 2020/04/21 設定欄位及對應TB欄位,從ExcelSave搬來修改
Private Sub SetField_Old()
'    Dim ii As Integer, intW As Integer
'    Dim strTBF As String
'
'    ReDim strFieldTB(UBound(strFieldN))
'    ReDim intWidth(UBound(strFieldN))
'
''    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
''      strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "法務部", "智權部", "總所/管理", "全所")
''      'Modified by Lydia 2017/06/06 調整大小
''      'intWidth = Array(13, 10, 10, 0, 10, 13, 10, 10, 10, 10, 10, 10)
''      intWidth = Array(13, 10, 10, 0, 11, 10, 11, 10, 10, 10, 10, 13)
''    Else
''      strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
''      intWidth = Array(13, 10, 10, 10, 10, 13, 10, 10, 10, 10, 10, 13)
''    End If
'    For ii = LBound(strFieldN) To UBound(strFieldN)
'        intW = 0
'        strTBF = ""
'        Select Case strFieldN(ii)
'            Case "會計科目"
'                strTBF = "R41203"
'                intW = 13
'            Case "專利"
'                 strTBF = "R41204"
'                intW = 10
'            Case "商標"
'                 strTBF = "R41205"
'                intW = 10
'            Case "法務"
'                strTBF = "R41206"
'                intW = 10
'                If Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(智慧所更名日, 6) - 191100) Then
'                    If strCmp <> "L" And strCmp <> MsgText(601) Then
'                        intW = 0
'                    End If
'                ElseIf Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'                    intW = 0
'                End If
'            Case "CFP"
'                strTBF = "R41207"
'                intW = 11
'            Case "CFT"
'                strTBF = "R41208"
'                intW = 10
'            Case "FCP"
'                strTBF = "R41209"
'                intW = 11
'            Case "FCT"
'                strTBF = "R41210"
'                intW = 10
'            Case "FCL", "投法", "法務部"
'                strTBF = "R41211"
'                intW = 10
'            Case "CFL"
'                strTBF = "R41216"
'                intW = 10
'            Case "智權部"
'                strTBF = "R41212"
'                intW = 10
'            Case "總所/管理"
'                strTBF = "R41214"
'                intW = 10
'            Case "全所"
'                strTBF = "R41213"
'                intW = 13
'        End Select
'        strFieldTB(ii) = strTBF
'        intWidth(ii) = intW
'    Next ii
End Sub

'Add by Amy 2020/08/12 刪除欄位
Private Sub SetAndDelField(ByRef m_Xls As Worksheet)
    Dim bolDel As Boolean
    
    'Memo 刪除順序依後至前刪,避免刪錯欄
    For ii = UBound(strFieldN) - 1 To LBound(strFieldN) + 1 Step -1
        bolDel = False
        
        Select Case strFieldN(ii)
            Case "管理"
                m_Xls.Range(Chr(ii + intField) & intTitleRow).Value = "總所/管理"
            Case "法務部"
                '畫面條件起月不是 105年後 且 公司別為L公司 或 空白,顯示為 法務 (非法務部)
                If Val(Replace(MaskEdBox1, "/", "")) >= 10501 And (strCmp = "L" Or strCmp = MsgText(601)) Then
                    m_Xls.Range(Chr(ii + intField) & intTitleRow).Value = "L"
                '畫面條件起月不是 105年後且智慧所更名日前資料,刪除 法務部 欄位
                ElseIf Not (Val(Replace(MaskEdBox1, "/", "")) >= 10501 And Val(Replace(MaskEdBox1, "/", "")) < Val(Left(智慧所更名日, 6)) - 191100) Then
                    bolDel = True
                End If
            Case "投法"
                '畫面條件起月不是 105年前資料,刪除 投法 欄位
                If Not (Val(Replace(MaskEdBox1, "/", "")) < 10501) Then
                    bolDel = True
                End If
            Case "ACS"
                '畫面條件起月未包含10908年後資料,刪除 ACS 欄位
                If Not ((Val(Replace(MaskEdBox1, "/", "")) <= 10908 And Val(Replace(MaskEdBox2, "/", ""))) >= 10908 _
                  Or Val(Replace(MaskEdBox1, "/", "")) >= 10908 Or Val(Replace(MaskEdBox2, "/", "")) >= 10908) Then
                    bolDel = True
                End If
            Case "L"
                '畫面條件起月包含105年後資料,刪除 L欄位
                If Val(Replace(MaskEdBox1, "/", "")) >= 10501 Or (Val(Replace(MaskEdBox1, "/", "")) >= Val(Left(智慧所更名日, 6)) - 191100 And (strCmp <> "L" Or strCmp <> MsgText(601))) Then
                    bolDel = True
                '畫面條件起月是 105年前資料,L欄位名稱改為法務
                Else
                    m_Xls.Range(Chr(ii + intField) & intTitleRow).Value = "法務"
                End If
                
        End Select
        If bolDel = True Then
            m_Xls.Range(Chr(ii + intField) & ":" & Chr(ii + intField)).Delete Shift:=xlToLeft
        End If
    Next ii
End Sub

