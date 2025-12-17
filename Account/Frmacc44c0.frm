VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44c0 
   AutoRedraw      =   -1  'True
   Caption         =   "年度綜合損益統計表 "
   ClientHeight    =   1524
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5064
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1524
   ScaleWidth      =   5064
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   210
      Width           =   3500
   End
   Begin VB.CommandButton cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
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
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   720
      Width           =   1215
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
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1005
      _ExtentX        =   1778
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
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1005
      _ExtentX        =   1778
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "區間不可下跨年"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1100
      Width           =   1995
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "區間"
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
      Left            =   120
      TabIndex        =   10
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "~"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "半年期"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc44c0"
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
Public adoaccrpt413 As New ADODB.Recordset
Dim lngCounter As Long
Dim douTotal1(8), douTotal2(8), douTotal3(8) As Double
Dim dllaccrpt413 As Object
'Add by Amy 2018/03/02
Const MskFormat As String = "###/##"
Dim ado44c0 As New ADODB.Recordset, adoAccList As New ADODB.Recordset
Dim bolExcel As Boolean, bolNYAcc As Boolean  '是否產生Excel/是否跨年度
Dim strQ As String, strAccNoT As String
Dim i As Integer, intCounter As Integer, intField As Integer
Dim strTp1 As String, strTp2 As String, stTemp(1) As String
Dim strY(1) As String, strM(1) As String
Dim strF, intWidth, arrAcOld, arrAcNext, arrTName
Dim strPreYM_S As String, strPreYM_E As String '去年起/迄
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/4/27
Dim bol4999 As Boolean 'Add by Amy 2024/05/28

'Add by Sindy 2020/4/27
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

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
'end 2020/4/27

'Add by Amy 2018/03/02 +Excel
Private Sub Cmd_Excel_Click()

    bolExcel = True
    If FormCheck = False Then
        bolExcel = True
        Exit Sub
    End If
    
    Call SetCompN 'Add by Sindy 2020/4/27
    
    bolNYAcc = False
    If Left(Replace(MaskEdBox1.Text, "/", ""), 3) <> Left(Replace(MaskEdBox2.Text, "/", ""), 3) Then
        bolNYAcc = True
    End If
   bol4999 = False 'Add by Amy 2024/05/28
    Screen.MousePointer = vbHourglass
    ProduceData_New
    
    'Modify by Amy 2018/11/22 +去年該當月累計/比較/前年該當月累計/比較
    'strQ = "Select R002 as AccNo,R003 as AccN,Decode(SubStr(R004,Length(R004)-1,2),'13',Decode(length(R004),5,SubStr(R004,1,3),SubStr(R004,1,2))||'年合計',R004) as YM,R005 as Amount " & _
           "From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' ORDER BY R002,R004 "
    strExc(0) = Replace(MaskEdBox2.Text, "/", "")
    If Len(strExc(0)) = 5 Then
         strExc(0) = Left(strExc(0), 3)
    Else
         strExc(0) = Left(strExc(0), 2)
    End If
    
    'Modify by Amy 2024/05/28 L公司會抓2字頭[來自智慧所收入(4999)],並列此科目後面 原:Order by AccNo,R004
    strQ = "Select AccNo,AccN, Decode(SubStr(R004,Length(R004)-1,2),'13', '當年累計', '14','" & Val(strExc(0)) - 1 & "年" & Val(Right(strPreYM_S, 2)) & "-" & Val(Right(strPreYM_E, 2)) & "月','15','" & strExc(0) & "vs" & Val(strExc(0)) - 1 & "','16','" & Val(strExc(0)) - 2 & "年" & Val(Right(strPreYM_S, 2)) & "-" & Val(Right(strPreYM_E, 2)) & "月','17','" & Val(strExc(0)) & "vs" & Val(strExc(0)) - 2 & "', R004) as YM,R005 as Amount From " & _
                "(Select Distinct R002 as AccNo,R003 as AccN From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R003 is not null), " & _
                "(Select R002,R003,R004,R005 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And ((R004>=" & Val(Replace(MaskEdBox1.Text, "/", "")) & " And R004<=" & Val(Replace(MaskEdBox2.Text, "/", "")) & ") Or R004=" & Val(strExc(0) & "13") & ") " & _
       "Union Select R002,R003, " & Val(strExc(0)) & "14,R005 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R004=" & Val(strExc(0) - 1) & "13 " & _
       "Union Select R002,R003, " & Val(strExc(0)) & "15,R005 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R004=" & Val(strExc(0) - 1) & "13 " & _
       "Union Select R002,R003, " & Val(strExc(0)) & "16,R005 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R004=" & Val(strExc(0) - 2) & "13 " & _
       "Union Select R002,R003, " & Val(strExc(0)) & "17,R005 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R004=" & Val(strExc(0) - 2) & "13) " & _
       "Where AccNo=R002(+) Order by Decode(Substr(AccNo,1,4),'2407','4999'||AccNo,AccNo),R004"
    'end 2018/11/22
    If ado44c0.State = adStateOpen Then ado44c0.Close
    ado44c0.CursorLocation = adUseClient
    ado44c0.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If ado44c0.RecordCount = 0 Then
        MsgBox "無資料！"
    Else
        If SaveExcel = True Then
            bolExcel = False
        End If
    End If
    ado44c0.Close
    Screen.MousePointer = vbDefault
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData_New()
    Dim strWhere As String

On Error GoTo Checking
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
    lngCounter = 0
    
    '刪除暫存檔(與總帳共用)
    strQ = "Delete From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1'"
    adoTaie.Execute strQ
    
    strPreYM_S = Val(Replace(MaskEdBox1.Text, "/", "")) + 191000 - 191100
    strPreYM_E = Val(Replace(MaskEdBox2.Text, "/", "")) + 191000 - 191100
    
    'Modify By Sindy 2020/4/27
'    If Text6 <> "" Then
'        strWhere = "And (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') "
'    End If
    If strCmp <> "" Then
      If InStr(strCmp, "+") > 0 Then
         strWhere = "And (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) "
      Else
         strWhere = "And (a0109 is null or a0109 = '" & strCmp & "') "
      End If
    End If
    '2020/4/27 END
    
'-------------------------------------------------
' 實際營業收入
'-------------------------------------------------
    strQ = "Select * From acc010 Where a0101 >= '4' and a0101 < '5' And a0104 = '3' and instr(a0102,'不用')=0 " & strWhere & _
               "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
       SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
       adoacc010.MoveNext
    Loop
    adoacc010.Close
  
'-------------------------------------------------
' 實際營業支出
'-------------------------------------------------
    strQ = "Select * From acc010 Where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0  " & strWhere & _
               "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'-------------------------------------------------
' 營業外收入
'-------------------------------------------------
    strQ = "Select * From acc010 Where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 " & strWhere & _
               "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'-------------------------------------------------
' 營業外支出
'-------------------------------------------------
    strQ = "Select * From acc010 Where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 " & strWhere & _
               "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
        adoacc010.MoveNext
    Loop
    adoacc010.Close
   
'-------------------------------------------------
' 合計
'-------------------------------------------------
    '會計科目年合計
    'Modify by Amy 2018/11/22 不論是否下一年都需顯示年合計
    'If bolNYAcc = True And bolExcel = True Then SaveSum "", "", "年合計"
    SaveSum "", "", "年合計"
    
    '營業毛利/營業收入
    SaveSum "4", "499999", "4T"

    '營業支出
    SaveSum "6", "699999", "6T"
    
    '營業損益
    SaveSum "4T", "6T", "6ZT"

    '營業外收入
    SaveSum "71", "719999", "71T"
    
    '營業外支出
    SaveSum "72", "729999", "72T"
  
'-------------------------------------------------
' 本期損益
'-------------------------------------------------
    SaveSum "ZZT", "ZZT", "ZZT"
    
    'Add by Amy 2024/05/28 L公司有4999資料,更新4141及2407 開頭名稱
    If bol4999 = True Then
      '更新4141xx名稱
       strSql = "Update Accrpt44c0 Set R003=(Select Replace(Replace(Replace(a0102,'法務收入',''),'-',''),' ','')||'收入' From Acc010 Where R002=a0101) " & _
                      "Where ID='" & strUserNum & "' And R001='1' And R002 Like '4141%' "
       adoTaie.Execute strSql
       '更新2407開頭名稱(不含 智慧所補)
       strSql = "Update Accrpt44c0 Set R003=(Select '　'||Replace(Replace(Replace(a0102,'代收款項',''),'-',''),' ','')||'收入' From Acc010 Where R002=a0101) " & _
                      "Where ID='" & strUserNum & "' And R001='1' And R002 Like '2407%' And R002<>'240799' "
       adoTaie.Execute strSql
       '更新4999 (來自智慧所收入) 名稱
       strSql = "Update Accrpt44c0 Set R003='來自智慧所收入' Where ID='" & strUserNum & "' And R001='1' And R002='4999' "
       adoTaie.Execute strSql
       '更新240799 (智慧所補) 名稱
       strSql = "Update Accrpt44c0 Set R003='　　智慧所補' Where ID='" & strUserNum & "' And R001='1' And R002='240799' "
       adoTaie.Execute strSql
   End If
       
    StatusClear
     
Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub SaveAcc040(strAccNo1 As String, strAccNo2 As String, strA0102 As String)
    Dim strSql As String
    Dim strWhere As String, strWhere2 As String 'Add by Amy 2018/11/22
    Dim ii As Integer, intNum As Integer, strCP As String 'Add by Amy 2024/05/28
    
    If bolExcel = True Then
        strSql = strSql & " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) >= " & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                    " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) <= " & Val(Replace(MaskEdBox2.Text, "/", ""))
        'Add by Amy 2018/11/22 +去年/前年
        strWhere = strWhere & " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) >= " & Val(strPreYM_S) & _
                    " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) <= " & Val(strPreYM_E)
        strWhere2 = strWhere2 & " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) >= " & Val(strPreYM_S + 191000 - 191100) & _
                    " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) <= " & Val(strPreYM_E + 191000 - 191100)
    Else
        '年度
        If Mid(Combo2, 1, 1) = "1" Then
            strSql = strSql & " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) >= " & Val(Text3 & "01") & _
                    " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) <= " & Val(Text3 & "06")
        Else
            strSql = strSql & " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) >= " & Val(Text3 & "07") & _
                    " And a0401||Decode(Length(a0402),1,'0'||a0402,a0402) <= " & Val(Text3 & "12")

        End If
    End If
    'Add by Amy 2024/05/28 10904月以後,L公司4141開頭會科抓6碼顯示(原:固定抓4碼)
    intNum = 4
    If Left(strCmp, 1) = "L" And Left(strAccNo1, 4) = "4141" _
      And (Left(Val(Replace(MaskEdBox1.Text, "/", "")) + 191100, 4) = "2020" And Val(Replace(MaskEdBox2.Text, "/", "")) >= 10904 Or Val(Replace(MaskEdBox1.Text, "/", "")) >= 10904) Then
         intNum = 6
         strAccNo1 = strAccNo1 & "00"
         strAccNo2 = strAccNo2 & "99"
    End If
    'end 2024/05/28
    
    '會計編號起迄
    'Modify by Amy 2018/11/22 +前一年
    'Modify by Amy 2024/05/28 L公司4字頭會科抓6碼顯示(原:固定抓4碼)
    If strAccNo1 <> MsgText(601) Then
        strSql = strSql & " and SubStr(a0405, 1, " & intNum & ") >= '" & strAccNo1 & "'"
        strWhere = strWhere & " and SubStr(a0405, 1, " & intNum & ") >= '" & strAccNo1 & "'"
        strWhere2 = strWhere2 & " and SubStr(a0405, 1, " & intNum & ") >= '" & strAccNo1 & "'"
    End If
    If strAccNo2 <> MsgText(601) Then
        strSql = strSql & " and SubStr(a0405, 1, " & intNum & ") <= '" & strAccNo2 & "'"
        strWhere = strWhere & " and SubStr(a0405, 1, " & intNum & ") <= '" & strAccNo2 & "'"
        strWhere2 = strWhere2 & " and SubStr(a0405, 1, " & intNum & ") <= '" & strAccNo2 & "'"
    End If
    'end 2024/05/28
    '公司別
    'Modify By Sindy 2020/4/27
'    If Text6 <> MsgText(601) Then
'        strSql = strSql & "And a0403 = '" & IIf(Text6 = "2", "J", "1") & "' "
'        strWhere = strWhere & "And a0403 = '" & IIf(Text6 = "2", "J", "1") & "' "
'        strWhere2 = strWhere2 & "And a0403 = '" & IIf(Text6 = "2", "J", "1") & "' "
'    End If
    If strCmp <> MsgText(601) Then
      If InStr(strCmp, "+") > 0 Then
         strSql = strSql & "And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
         strWhere = strWhere & "And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
         strWhere2 = strWhere2 & "And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
      Else
         strSql = strSql & "And a0403 = '" & strCmp & "' "
         strWhere = strWhere & "And a0403 = '" & strCmp & "' "
         strWhere2 = strWhere2 & "And a0403 = '" & strCmp & "' "
      End If
    End If
    '2020/4/27 END
    'end 2018/11/22
    
    'Modify by Amy 2018/11/22
    'Modify by Amy 2024/05/28 L公司4字頭會科抓6碼顯示(原:固定抓4碼)
    strSql = "Insert Into Accrpt44c0 (ID,R001,R002,R003,R004,R005) " & _
             "Select '" & strUserNum & "','1',SubStr(a0405, 1, " & intNum & "),'" & strA0102 & "',a0401||Decode(Length(a0402),1,'0'||a0402,a0402),Sum(Nvl(a0408,0)) " & _
             "From Acc040 Where a0404 = '" & MsgText(55) & "' " & strSql & " Group by a0401,a0402,SubStr(a0405, 1, " & intNum & ") " & _
    "Union Select '" & strUserNum & "','1',SubStr(a0405, 1, " & intNum & "),'" & strA0102 & "',a0401||Decode(Length(a0402),1,'0'||a0402,a0402),Sum(Nvl(a0408,0)) " & _
             "From Acc040 Where a0404 = '" & MsgText(55) & "' " & strWhere & " Group by a0401,a0402,SubStr(a0405, 1, " & intNum & ") " & _
    "Union Select '" & strUserNum & "','1',SubStr(a0405, 1, " & intNum & "),'" & strA0102 & "',a0401||Decode(Length(a0402),1,'0'||a0402,a0402),Sum(Nvl(a0408,0)) " & _
             "From Acc040 Where a0404 = '" & MsgText(55) & "' " & strWhere2 & " Group by a0401,a0402,SubStr(a0405, 1, " & intNum & ") "
    adoTaie.Execute strSql
    
    'Add by Amy 2024/05/28 L公司有414102,增加抓 來自智慧所收入(2407開頭)的收款傳票(ACC1P0資料)及智慧所補
    If intNum = 6 And HaveData("And R002='414102' ") = True Then
      strCP = "And cp01='LA' And cp02='999999' And cp05>=Date1 And cp05<=Date2 And cp16>0 "
      
      '來自智慧所收入(2407開頭)的收款傳票
      'Memo by Amy 因目前2407只用於L公司,避免加了L公司條件,不會知道其他特殊狀況,故先不加 A1p01='L' -秀玲
      strSql = ""
      For ii = -1 To 1
         strSql = strSql & "Select Substr(cp05,1,6)-191100 as YM,a1p05,Sum(a1p07) as a1p07 From Caseprogress,Acc1p0,Acc0m0,Acc0l0 " & _
                     "Where a1p05 LIKE '2407%' And cp60=a0m02(+) And a0m01=a0l01(+) And a0m01=a1p04(+) " & _
                     Replace(Replace(UCase(strCP), "DATE1", strPreYM_S + (191100 - ii * 100) & "01"), "DATE2", strPreYM_E + (191100 - ii * 100) & "31") & _
                     " Group by Substr(cp05,1,6)-191100,a1p05 "
         If ii < 1 Then strSql = strSql & " Union "
      Next ii
      '避免起迄日無2407開頭資料,不會顯示(ex:10901~10906),有414102 一定出現2407會計科目
      strSql = "Insert Into Accrpt44c0 (ID,R001,R002,R004,R005) " & _
                     "Select '" & strUserNum & "','1',a0101,YM,a1p07 From Acc010,(" & strSql & ") Where a0101=a1p05(+) And a0101 LIKE '2407%' And a0104='4' And (a0109 is null Or a0109='L') " & _
                     "And Exists (Select * From Accrpt44c0 Where ID='A2004' And R001='1' And R002='414102') "
      adoTaie.Execute strSql
      '避免沒資料也會Run Excel,故先抓414102 的值
      'Memo by Amy [智慧所補]於10904月(法律所成立)開始有資料,但會計科目2407於10909月才於Acc1p0有資料,故抓 414102值
      strSql = "Select '" & strUserNum & "','1','4999',R004,Sum(R005) From Accrpt44c0 Where ID='A2004' And R001='1' And R002='414102' Group by SubStr(R002,1,4),R004 "
      strSql = "Insert Into Accrpt44c0 (ID,R001,R002,R004,R005) " & strSql
      adoTaie.Execute strSql
      
      '智慧所補
      strCP = "Select '240799' as AccNo,Substr(cp05,1,6)-191100 as YM,Sum(Nvl(cp16,0)) as Amt From CaseProgress Where " & Mid(strCP, 4)
      strSql = ""
      For ii = -1 To 1
         strSql = strSql & Replace(Replace(UCase(strCP), "DATE1", strPreYM_S + (191100 - ii * 100) & "01"), "DATE2", strPreYM_E + (191100 - ii * 100) & "31") & _
                        " Group by Substr(cp05,1,6)-191100 "
         If ii < 1 Then strSql = strSql & " Union "
      Next ii
      '有414102會計科目資料才寫入智慧所補會計科目
      strSql = "Insert Into Accrpt44c0 (ID,R001,R002,R004,R005) " & _
                     "Select '" & strUserNum & "','1',AccNo,YM,Amt From (" & strSql & ") " & _
                     "Where Exists(Select * From Accrpt44c0 Where ID='A2004' And R001='1' And R002='414102') "
      adoTaie.Execute strSql
      bol4999 = True
    End If
    'end 2024/05/28
End Sub

Private Sub SaveSum(strAccNo1 As String, strAccNo2 As String, strSumNo As String)
    Dim stSQL As String, GetSumName As String, intIns As Integer
    
    If InStr(strSumNo, "T") > 0 Then
        Select Case strSumNo
            Case "4T"
                GetSumName = "營業收入"
            '營業支出
            Case "6T"
                GetSumName = ReportSum(2)
            '營業損益
            Case "6ZT"
                GetSumName = ReportSum(3)
            '營業外收入
            Case "71T"
                GetSumName = ReportSum(5)
            '營業外支出
            Case "72T"
                GetSumName = ReportSum(6)
            '稅前淨損益
            Case "ZZT"
                GetSumName = ReportSum(7)
        End Select
        GetSumName = Replace(GetSumName, ":", "")
    End If
    If strSumNo = "6ZT" Then
        '營業損益(6ZT)=營業收入(4T)-營業支出(6T)
        stSQL = "Select '" & strUserNum & "','1','" & strSumNo & "','" & GetSumName & "',RYM,Nvl(SZT,0)-Nvl(SE,0) From " & _
                  "(Select R004 as RYM,Sum(R005) as SZT From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R002='4T' Group by R004)," & _
                  "(Select R004 as EYM,Sum(R005) as SE From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R002='6T' Group by R004) " & _
                "Where RYM=EYM(+) "
    ElseIf strSumNo = "ZZT" Then
        '稅前淨損益(ZZT)=營業損益(6ZT)+營業外收入(71T)-營業外支出(72T)
        stSQL = "Select '" & strUserNum & "','1','" & strSumNo & "','" & GetSumName & "',RYM,Nvl(SZT,0)-Nvl(T72,0) From " & _
                  "(Select R004 as RYM,Sum(R005) as SZT From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And (R002='6ZT' or R002='71T') Group by R004)," & _
                  "(Select R004 as EYM,Sum(R005) as T72 From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' And R002='72T' Group by R004) " & _
                "Where RYM=EYM(+) "
    ElseIf strSumNo = "年合計" Then
        '會計科目年合計(跨年區間若其中一年無值也需新增,加總才會有值)
        stSQL = "Select '" & strUserNum & "','1',AccNo,'',YM||'13',Nvl(R005,0) From " & _
                "(Select AccNo,YM From " & _
                        "(Select Distinct Decode(Length(R004),5,SubStr(R004,1,3),SubStr(R004,1,2)) as YM From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1')," & _
                        "(Select Distinct R002 as AccNo From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1')" & _
                ")," & _
                "(Select R002,Decode(Length(R004),4,SubStr(R004,1,2),SubStr(R004,1,3)) as R004,Sum(R005) as R005 From Accrpt44c0 " & _
                    "Where ID='" & strUserNum & "' And R001='1' Group by R002,Decode(Length(R004),4,SubStr(R004,1,2),SubStr(R004,1,3))" & _
                ") Where AccNo=R002(+) And YM=R004(+)"
    Else
        stSQL = "Select '" & strUserNum & "','1','" & strSumNo & "','" & GetSumName & "',R004,Sum(R005) From Accrpt44c0 " & _
                "Where ID='" & strUserNum & "' And R001='1' And R002>='" & strAccNo1 & "' And R002<='" & strAccNo2 & "' Group by R004 "
    End If
    stSQL = "Insert Into Accrpt44c0 (ID,R001,R002,R003,R004,R005) " & stSQL
    adoTaie.Execute stSQL, intIns
    'Add by Amy 2018/03/07 若沒資料可合計也需要新增,要照損益表格式顯示-婧瑄
    If intIns = 0 And strSumNo <> "年合計" Then
        stSQL = "Insert Into Accrpt44c0 (ID,R001,R002,R003,R004,R005) " & _
                    "Select Distinct '" & strUserNum & "','1','" & strSumNo & "','" & GetSumName & "',R004,0 From Accrpt44c0 Where ID='" & strUserNum & "' and R001='1' "
        adoTaie.Execute stSQL, intIns
    End If
End Sub

Private Function SaveExcel() As Boolean
    Dim xlsAgentPoint As New Excel.Application, wksrpt As New Worksheet
    Dim j As Integer, intTitleR As Integer
    Dim stStartC As String, stPAndL As String '會計科目年度合計起始欄/營業損益列
    Dim intStartR As Integer '合計起始列
    Dim xlsFileName As String, strOldAccNo As String
    Dim bolSum As Boolean, bolB As Boolean, stStyle As String '加總欄/粗體/設定線
    Dim strColorField As String 'Add by Amy 2018/03/07 顏色欄
    Dim strFormat As String 'Add by Amy 2018/11/22 格式
    Dim strNField As String 'Add by Amy 2021/10/19 目前位置
    Dim strL414102 As String, strL4999 As String, int2407StartR As Integer 'Add by Amy 2024/05/24 顧問收入/來自智慧所收入 位置/2407開頭起始位置
    
On Error GoTo ErrHand
'*** Memo 10904 法律法成立,格式(4字頭會計科目)與其他公司不同 ***

    'Modify By Sindy 2020/4/27
'    xlsFileName = Replace(MaskEdBox1, "/", "") & "-" & Replace(MaskEdBox2, "/", "") & "年度" & Replace(Replace(IIf(Text7 = "", "台一　專利商標/智權", Text7), "　", ""), "/", "") & _
'                  "公司綜合損益表" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    'Modify by Amy 2024/06/20 檔名改與表單名相同,因太多損益表名稱一樣
    xlsFileName = Replace(MaskEdBox1, "/", "") & "-" & Replace(MaskEdBox2, "/", "") & "年度" & Replace(Replace(strCmpN, "　", ""), "/", "") & _
                  "公司" & Trim(Replace(ReportTitle(413), "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & xlsFileName
    End If
    
    xlsAgentPoint.SheetsInNewWorkbook = 3 'Modify by Amy 2021/10/19 改回 原:1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    intField = 65:  intCounter = 1
    'xlsAgentPoint.Visible = True
    '*** 抬頭
    Call SetTitleArr
    '欄位超過Z欄的error
    strColorField = GetFieldStr(GetValue(Val(Replace(MaskEdBox2, "/", ""))), intField)
    '報表名稱/公司別/區間改至頁首顯示
    wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "列印人員：" & StaffQuery(strUserNum)
    wksrpt.Range(GetFieldStr(9, intField) & intCounter).Value = "列印日期：" & CFDate(strSrvDate(2))
    intCounter = intCounter + 1
    For i = LBound(strF) To UBound(strF)
        'Modify by Amy 2021/10/19 年月中間加/
        If i > GetValue("科目名稱") And i < GetValue("當年累計") Then
            wksrpt.Range(GetFieldStr(i, intField) & intCounter).Value = Val(Mid(Val(strF(i)) + 191100, 1, 4)) - 1911 & "/" & Right(strF(i), 2)
        Else
            wksrpt.Range(GetFieldStr(i, intField) & intCounter).Value = strF(i)
        End If
        wksrpt.Range(GetFieldStr(i, intField) & intCounter).ColumnWidth = intWidth(i)
        'Add by Amy 2018/11/22 欄位名稱也顯示欄色-婧瑄
        If strColorField = GetFieldStr(GetValue("" & strF(i)), intField) Then
            wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & intCounter).Interior.ColorIndex = 20 '設置儲存格填充色(藍)
            wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & intCounter).Interior.tintandshade = 0.5 '設深淺
        End If
        'wksrpt.Range(GetFieldStr(i, intField) & intCounter).Font.Bold = True  'Mark by 2018/11/22 拿掉粗體-婧瑄
    Next i
    wksrpt.Range(GetFieldStr(LBound(strF), intField) & intCounter & ":" & GetFieldStr(UBound(strF), intField) & intCounter).HorizontalAlignment = xlCenter
    
    intTitleR = intCounter: stStartC = GetFieldStr(1, intField): intStartR = intTitleR + 1
    ado44c0.MoveFirst
    Do While ado44c0.EOF = False
        stTemp(0) = "": stTemp(1) = "": bolSum = False
        strFormat = "#,##0.00_ ;[紅色]-#,##0.00"
        'Modify by Amy 2018/11/22 +if 二年比較,固定公式
        If InStr("" & ado44c0.Fields("YM"), "vs") > 0 Then
            'Modify by Amy 2021/10/19 下099會錯
             strNField = "" & ado44c0.Fields("YM")
             If Mid(strNField, 1, 1) = "0" Then strNField = Mid(strNField, 2)
            stTemp(0) = GetFieldStr(GetValue("當年累計"), intField) & intCounter & "/" & GetFieldStr(GetValue(strNField) - 1, intField) & intCounter & "-1"
            strFormat = "0.00%;[紅色]-0.00%"
            stStyle = "xlContinuous" '單線
            If "" & ado44c0.Fields("AccNo") = "6ZT" Or "" & ado44c0.Fields("AccNo") = "ZZT" Then
                stStyle = "xlDouble" '雙線
            End If
            wksrpt.Range(GetFieldStr(GetValue(strNField), intField) & intCounter).Value = "=" & stTemp(0)
            wksrpt.Range(GetFieldStr(GetValue(strNField), intField) & intCounter).NumberFormatLocal = strFormat
            'wksrpt.Range(GetFieldStr(GetValue(ado44c0.Fields("YM")), intField) & intCounter).Font.Bold = True 'Mark by 2018/11/22 拿掉粗體-婧瑄
            If InStr("" & ado44c0.Fields("AccNo"), "T") > 0 Then
                wksrpt.Range(GetFieldStr(GetValue(strNField), intField) & intCounter).Interior.ColorIndex = 40  '設置儲存格填充色(膚)
                wksrpt.Range(GetFieldStr(GetValue(strNField), intField) & intCounter).Interior.tintandshade = 0.5  '設深淺
            End If
            If GetValue(strNField) = UBound(strF) Then stStartC = GetFieldStr(1, intField)
            'end 2021/10/19
        '合計
        ElseIf InStr("" & ado44c0.Fields("AccNo"), "T") > 0 Then
            bolSum = True
            '只設第一筆合計(因J公司有些月份會無資料 ex:產生10305-10403,10305營業外收入無資料,無法顯示加總公式
            If strOldAccNo <> "" & ado44c0.Fields("AccNo") Then
                intCounter = intCounter + 1
                '記錄營業收入 or 營業支出 or 營業外收入
                If "" & ado44c0.Fields("AccNo") = "4T" Or "" & ado44c0.Fields("AccNo") = "6T" Or "" & ado44c0.Fields("AccNo") = "71T" Or "" & ado44c0.Fields("AccNo") = "72T" Then
                    If "" & ado44c0.Fields("AccNo") = "71T" Then
                        stPAndL = stPAndL & "+" & GetFieldStr(LBound(strF) + 1, intField) & intCounter
                    Else
                        stPAndL = stPAndL & "-" & GetFieldStr(LBound(strF) + 1, intField) & intCounter
                    End If
                    'Modify by Amy 2024/05/28 +if L公司4T 設定顧問收文及來自智慧所收入
                    If bol4999 = True And "" & ado44c0.Fields("AccNo") = "4T" Then
                        '4T
                        stTemp(1) = "Sum(" & GetFieldStr(LBound(strF) + 1, intField) & intStartR & ":" & GetFieldStr(LBound(strF) + 1, intField) & strL4999 & ")"
                    'Add by Amy 2018/03/07 若沒資料合計也需顯示-婧瑄
                    ElseIf intStartR < intCounter Then
                        stTemp(1) = "Sum(" & GetFieldStr(LBound(strF) + 1, intField) & intStartR & ":" & GetFieldStr(LBound(strF) + 1, intField) & intCounter - 1 & ")"
                    Else
                        stTemp(1) = "0"
                    End If
                    stStyle = "xlContinuous" '單線
                '記錄營業損益/營業外支出/本期損益
                ElseIf "" & ado44c0.Fields("AccNo") = "6ZT" Or "" & ado44c0.Fields("AccNo") = "ZZT" Then
                    stTemp(1) = Mid(stPAndL, 2)
                    stPAndL = "+" & GetFieldStr(LBound(strF) + 1, intField) & intCounter
                    stStyle = "xlDouble" '雙線
                End If
                wksrpt.Range(GetFieldStr(LBound(strF), intField) & intCounter).Value = "" & ado44c0.Fields("AccN")
                For i = LBound(strF) + 1 To UBound(strF)
                    'Modify by Amy 2024/05/28 +if L公司4T 設定顧問收文及來自智慧所收入
                    If bol4999 = True And "" & ado44c0.Fields("AccNo") = "4T" Then
                        If InStr(strF(i), "vs") > 0 Then
                        ElseIf strF(i) = "當年累計" Then
                           '來自智慧所收入
                           strExc(2) = "=Sum(" & stStartC & strL4999 & ":" & GetFieldStr(GetValue("" & strF(i)) - 1, intField) & strL4999 & ")"
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).Value = strExc(2)
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).NumberFormatLocal = strFormat
                        Else
                           '來自智慧所收入=2407開頭加總
                           strExc(2) = "=Sum(" & GetFieldStr(GetValue("" & strF(i)), intField) & int2407StartR & ":" & GetFieldStr(GetValue("" & strF(i)), intField) & intCounter - 1 & ")"
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).Value = strExc(2)
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).NumberFormatLocal = strFormat
                           '顧問收入=顧問收入 餘額-來自智慧所收入
                           strExc(2) = wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL414102).Value
                           If strExc(2) <> MsgText(601) Then
                              strExc(2) = "=" & strExc(2) & "-" & GetFieldStr(GetValue("" & strF(i)), intField) & strL4999
                           End If
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL414102).Value = strExc(2)
                           wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL414102).NumberFormatLocal = strFormat
                           If strColorField = GetFieldStr(GetValue("" & strF(i)), intField) Then
                              wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).Interior.ColorIndex = 20 '設置儲存格填充色(藍)
                              wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL4999).Interior.tintandshade = 0.5 '設深淺
                              wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL414102).Interior.ColorIndex = 20 '設置儲存格填充色(藍)
                              wksrpt.Range(GetFieldStr(GetValue("" & strF(i)), intField) & strL414102).Interior.tintandshade = 0.5 '設深淺
                           End If
                        End If
                    End If
                    'end 2024/05/28
                    stTemp(0) = Replace(stTemp(1), GetFieldStr(LBound(strF) + 1, intField), GetFieldStr(i, intField))
                     
                    'Modify by Amy 2018/03/07 若沒資料合計也需顯示-婧瑄
                    wksrpt.Range(GetFieldStr(i, intField) & intCounter).Value = IIf(stTemp(0) <> "0", "=", "") & stTemp(0)
                    wksrpt.Range(GetFieldStr(i, intField) & intCounter).NumberFormatLocal = strFormat 'Modify by Amy 2018/11/22 改為變數
                    'wksrpt.Range(GetFieldStr(i, intField) & intCounter).Font.Bold = True 'Mark by 2018/11/22 拿掉粗體-婧瑄
                    'Add by Amy 2018/03/07
                    wksrpt.Range(GetFieldStr(i, intField) & intCounter).Interior.ColorIndex = 40  '設置儲存格填充色(膚)
                    wksrpt.Range(GetFieldStr(i, intField) & intCounter).Interior.tintandshade = 0.5  '設深淺
                Next i
                If stStyle <> MsgText(601) Then
                    If stStyle = "xlDouble" Then
                        wksrpt.Range(GetFieldStr(0, intField) & intCounter & ":" & GetFieldStr(UBound(strF), intField) & intCounter).Borders(xlEdgeBottom).LineStyle = xlDouble
                    Else
                        wksrpt.Range(GetFieldStr(0, intField) & intCounter & ":" & GetFieldStr(UBound(strF), intField) & intCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    End If
                    stStyle = ""
                End If
                intStartR = intCounter + 1
            End If
        '內容
        Else
            If strOldAccNo <> "" & ado44c0.Fields("AccNo") And "" & ado44c0.Fields("AccN") <> MsgText(601) Then
                intCounter = intCounter + 1
                wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "" & ado44c0.Fields("AccN")
                'Add by Amy 2024/05/28 L公司有 414102(顧問收入)及 來自智慧所收入文(4999) 需記錄位置
                If strCmp = "L" And ("" & ado44c0.Fields("AccNo") = "414102" Or "" & ado44c0.Fields("AccNo") = "4999") Then
                  If "" & ado44c0.Fields("AccNo") = "414102" Then
                     strL414102 = intCounter
                  Else
                     strL4999 = intCounter
                     int2407StartR = intCounter + 1
                  End If
                End If
            End If
            'Modify by Amy 2024/05/28 L公司有 來自智慧所收入文(4999) 於營收入時顯示
            If Not (strCmp = "L" And "" & ado44c0.Fields("AccNo") = "4999") Then
               stTemp(1) = "" & ado44c0.Fields("YM")
               '會計科目年度合計
               'Modify by Amy 2018/11/22  原:年合計
               If Right(stTemp(1), 4) = "當年累計" Then
                  stTemp(0) = "=Sum(" & stStartC & intCounter & ":" & GetFieldStr(GetValue(stTemp(1)) - 1, intField) & intCounter & ")"
                  'bolB = True 'Mark by 2018/11/22 拿掉粗體-婧瑄
               Else
                  stTemp(0) = "" & ado44c0.Fields("Amount")
               End If
               wksrpt.Range(GetFieldStr(GetValue(stTemp(1)), intField) & intCounter).Value = stTemp(0)
               wksrpt.Range(GetFieldStr(GetValue(stTemp(1)), intField) & intCounter).NumberFormatLocal = strFormat  'Modify by Amy 2018/11/22 改為變數
                  
               'Mark by 2018/11/22 拿掉粗體-婧瑄
   '            If bolB = True Then
   '                wksrpt.Range(GetFieldStr(GetValue(stTemp(1)), intField) & intCounter).Font.Bold = True
   '                bolB = False
   '            End If
               'Add by Amy 2018/03/07
               If strColorField = GetFieldStr(GetValue(stTemp(1)), intField) Then
                  wksrpt.Range(GetFieldStr(GetValue(stTemp(1)), intField) & intCounter).Interior.ColorIndex = 20 '設置儲存格填充色(藍)
                  wksrpt.Range(GetFieldStr(GetValue(stTemp(1)), intField) & intCounter).Interior.tintandshade = 0.5 '設深淺
               End If
            
            End If
            
        End If
        
        strOldAccNo = "" & ado44c0.Fields("AccNo")
        ado44c0.MoveNext
    Loop
    'Add by Amy 2024/05/28 L公司 且畫面條件小於等於111年,依年顯示備註
    If bol4999 = True And Val(strPreYM_S) - 191000 <= 2022 Then
      intCounter = intCounter + 2: intStartR = intCounter
      wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "備註：報表提示文字說明"
      intCounter = intCounter + 1
      wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　113.04新修改報表程式,因110年以前收款的規則有變動 , 若重新列印損益表時損益相符 , 但有部份科目不符"
      intCounter = intCounter + 1
      'Mark by Amy 111/2月初因過年,所以法律所於111/1/28先收文2月的資料,秀玲將111/1/28收文之AB1004251的收文日期改為111/2/7
      '                                         請辜調整文字內容,以下原因辜說不需顯示
'      wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "二、核對不符原因說明如下："
'      For i = 1 To -1 Step -1
'         strTp1 = Left(Val(strPreYM_S) + (191100 - i * 100), 4)
'         '109年
'         If strTp1 = "2020" Then
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　109年"
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　法律所 109.04開始 , 故前三個月無資料"
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　4~8月 收款來源處傳送至總帳產生傳票分錄不同 , 總帳資料人為變更分錄及數字"
'         ElseIf strTp1 = "2021" Then
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　110年"
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　1~6月 收款來源處都有1筆$500,000 , 但傳送至總帳故意改把科目為「累積損益」(3223)"
'         ElseIf strTp1 = "2022" Then
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　111年"
'            intCounter = intCounter + 1
'            wksrpt.Range(GetFieldStr(0, intField) & intCounter).Value = "　　1~2 收據的收文號日期都落在1月(因農曆年自1月底至2月初 , 可能是這個原因提早收文)"
'         End If
'      Next i
      wksrpt.Range(GetFieldStr(0, intField) & intStartR & ":" & GetFieldStr(0, intField) & intCounter).HorizontalAlignment = xlLeft
    End If
    
    wksrpt.Range(GetFieldStr(0, intField) & "1:" & GetFieldStr(UBound(strF), intField) & intCounter).RowHeight = 18
    wksrpt.Range(GetFieldStr(0, intField) & "1:" & GetFieldStr(UBound(strF), intField) & intCounter).Font.Size = 10
    'end 2018/03/27
    
    wksrpt.PageSetup.PaperSize = 9 '設定紙張 A4
    wksrpt.PageSetup.Orientation = xlPortrait 'Modify by Amy 2018/11/22 原:xlLandscape '橫印
    
    'Add by Amy 2018/11/22 報表名稱/公司別/區間從上搬下來
    'Modify By Sindy 2020/4/27
'    wksrpt.PageSetup.CenterHeader = Trim(Replace(ReportTitle(413), "*", "")) & Chr(10) & _
'                                    IIf(Text7 = "", "台一　國際專利商標/智權公司", Replace(Text7, Text6, "")) & Chr(10) & _
'                                    "日期" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
    wksrpt.PageSetup.CenterHeader = Trim(Replace(ReportTitle(413), "*", "")) & Chr(10) & _
                                    strCmpN & Chr(10) & _
                                    "日期" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
    '2020/4/27 END
    
    wksrpt.PageSetup.PrintTitleColumns = "$A:$A"  '表頭保留欄
    'end 2018/11/22
    wksrpt.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
    wksrpt.PageSetup.PrintGridlines = True '列印格線
    wksrpt.PageSetup.CenterHorizontally = True '水平置中(版面設定->邊界->水平置中)
    '邊界
    'Modify by Amy 2018/11/22 修改邊界、比例
    wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.19)
    wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.19)
    wksrpt.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.9)
    wksrpt.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0)
    wksrpt.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.38)
    wksrpt.PageSetup.FooterMargin = xlsAgentPoint.InchesToPoints(0)
    wksrpt.PageSetup.Zoom = 80 '縮放比例
    'end 2018/11/22
    
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    MsgBox "Excel已產生！"
    SaveExcel = True
    FormClear
    Exit Function
        
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
End Function

Private Sub SetTitleArr()
    Dim j As Integer

    strY(0) = Val(Replace(MaskEdBox1, "/", ""))
    strM(0) = Right(strY(0), 2)
    strY(0) = Val(Left(strY(0), IIf(Len(CStr(strY(0))) = 5, 3, 2)))

    strY(1) = Val(Replace(MaskEdBox2, "/", ""))
    strM(1) = Right(strY(1), 2)
    strY(1) = Val(Left(strY(1), IIf(Len(CStr(strY(1))) = 5, 3, 2)))

    strTp1 = "": strTp2 = ""
    If Val(strY(0)) = Val(strY(1)) Then
        For i = Val(strM(0)) To Val(strM(1))
            strTp1 = strTp1 & "," & strY(0) & IIf(Len(CStr(i)) = 2, i, "0" & i)
            strTp2 = strTp2 & "," & "10"
        Next i
        strTp1 = strTp1 & ",當年累計," & Val(strY(0)) - 1 & "年" & Val(strM(0)) & "-" & Val(strM(1)) & "月" & "," & Val(strY(0)) & "vs" & Val(strY(0)) - 1
        strTp2 = strTp2 & ",10.75,10.75,7"
        strTp1 = strTp1 & "," & Val(strY(0)) - 2 & "年" & Val(strM(0)) & "-" & Val(strM(1)) & "月" & "," & Val(strY(0)) & "vs" & Val(strY(0)) - 2
        strTp2 = strTp2 & ",10.75,7"
    'Mark by Amy 2018/11/22 不可下跨年
'    Else
'        For i = Val(strY(0)) To Val(strY(1))
'            If i = Val(strY(0)) Then
'                For j = Val(StrM(0)) To 12
'                    strTp1 = strTp1 & "," & i & IIf(Len(CStr(j)) = 2, j, "0" & j)
'                    strTp2 = strTp2 & "," & "13"
'                Next j
'                strTp1 = strTp1 & "," & i & "當年累計"
'                strTp2 = strTp2 & "," & "12"
'            ElseIf i = Val(strY(1)) Then
'                For j = 1 To Val(StrM(1))
'                    strTp1 = strTp1 & "," & i & IIf(Len(CStr(j)) = 2, j, "0" & j)
'                    strTp2 = strTp2 & "," & "13"
'                Next j
'                strTp1 = strTp1 & "," & i & "當年累計"
'                strTp2 = strTp2 & "," & "12"
'            Else
'                For j = 1 To 12
'                    strTp1 = strTp1 & "," & i & IIf(Len(CStr(j)) = 2, j, "0" & j)
'                    strTp2 = strTp2 & "," & "13"
'                Next j
'                strTp1 = strTp1 & "," & i & "當年累計"
'                strTp2 = strTp2 & "," & "12"
'            End If
'        Next i
    End If
    If strTp1 <> MsgText(601) Then
        strTp1 = "科目名稱" & strTp1
        strTp2 = "13" & strTp2
    End If

    strF = Split(strTp1, ",")
    intWidth = Split(strTp2, ",")
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(strF) To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2018/03/02

'Mark by Amy 2018/03/02 不使用
Private Sub Command1_Click()
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   Accrpt413Delete
'   ProduceData
'   '2014/2/20 modify by sonia
'   'dllaccrpt413.Acc44c0 ReportTitle(413), Text6, Text7, Text3, Combo2, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   dllaccrpt413.Acc44c0 ReportTitle(413), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), Text3, Combo2, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5190
   Me.Height = 2000 'Modify by Amy 2018/03/02 原:2200-婧瑄:拿掉 列印
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/4/27 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/4/27
   
   'Add by Amy 2018/03/02 +區間產生Excel
   MaskEdBox1.Mask = MskFormat
   MaskEdBox2.Mask = MskFormat
   'end 2018/03/02
   
   Combo2.AddItem ComboItem(151)
   Combo2.AddItem ComboItem(152)
   Combo2 = ComboItem(151)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'Set dllaccrpt413 = CreateObject("AccReport.ReportSelect") 'Mark by Amy 2018/03/02
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt413 = Nothing
   Set Frmacc44c0 = Nothing
End Sub

'Add by Amy 2018/03/02
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If Trim(Replace(MaskEdBox1, "/", "")) = MsgText(601) Or MaskEdBox1.Text = "___/__" Then Exit Sub
    
    If IsDate(ChangeTStringToWDateString(Replace(MaskEdBox1, "/", "") & "01")) = False Then
        MsgBox "起始區間輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If Trim(Replace(MaskEdBox2, "/", "")) = MsgText(601) Or MaskEdBox2.Text = "___/__" Then Exit Sub
        
    If IsDate(ChangeTStringToWDateString(Replace(MaskEdBox2, "/", "") & "01")) = False Then
        MsgBox "截止區間輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
End Sub
'end 2018/03/02

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Modify by Sindy 2020/4/27 公司別改下拉
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
'Private Sub Text6_Change()
'   '2014/2/20 modify by sonia
'   'If Text6 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text7 = A0802Query(Text6)
'   Select Case Text6
'      Case "1"
'         Text7 = A0802Query(Text6)
'      Case "2"
'         Text7 = A0802Query("J")
'      Case ""
'         Text7 = "台一　專利商標/智權"
'   End Select
'   '2014/2/20 end
'End Sub
'
''2014/2/20 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/2/20 end

''*************************************************
''  產生報表資料
''
''*************************************************
'Private Sub ProduceData()
'Dim intCounter As Integer
'
'On Error GoTo Checking
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   lngCounter = 0
'   adoaccrpt413.CursorLocation = adUseClient
'   adoaccrpt413.Open "select * from accrpt413", adoTaie, adOpenDynamic, adLockBatchOptimistic
''-------------------------------------------------
'' 實際營業收入
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 END
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt413Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(4)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   adoaccrpt413.Fields("r41303").Value = ReportSum(1)
'   If Mid(Combo2, 1, 1) = "1" Then
'      Calculate "4", "499999", 1, 6
'   Else
'      Calculate "4", "499999", 7, 12
'   End If
'   For intCounter = 3 To 8
'      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
'         douTotal1(intCounter) = 0
'      Else
'         douTotal1(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'
'   'Add By Cheng 2002/01/18
'   PaintLine ReportSum(8)
'
'   adoaccrpt413.UpdateBatch
''-------------------------------------------------
'' 實際營業支出
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 END
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt413Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(4)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   adoaccrpt413.Fields("r41303").Value = ReportSum(2)
'   If Mid(Combo2, 1, 1) = "1" Then
'      Calculate "6", "699999", 1, 6
'   Else
'      Calculate "6", "699999", 7, 12
'   End If
'   For intCounter = 3 To 8
'      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
'         douTotal2(intCounter) = 0
'      Else
'         douTotal2(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'
'   'Add By Cheng 2002/01/18
'   PaintLine ReportSum(8)
'
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(4)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   adoaccrpt413.Fields("r41303").Value = ReportSum(3)
'   For intCounter = 3 To 8
'      If douTotal1(intCounter) - douTotal2(intCounter) = 0 Then
'         adoaccrpt413.Fields(intCounter).Value = Null
'      Else
'         adoaccrpt413.Fields(intCounter).Value = douTotal1(intCounter) - douTotal2(intCounter)
'      End If
'   Next intCounter
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'
'   'Add By Cheng 2002/01/18
'   PaintLine ReportSum(8)
'
'   adoaccrpt413.UpdateBatch
''-------------------------------------------------
'' 營業外收入及支出
''-------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 END
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      Accrpt413Save
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(4)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   '2008/1/14 modify by sonia
'   'adoaccrpt413.Fields("r41303").Value = ReportSum(5)
'   adoaccrpt413.Fields("r41303").Value = ReportSum(19)
'   If Mid(Combo2, 1, 1) = "1" Then
'      Calculate "7", "799999", 1, 6
'   Else
'      Calculate "7", "799999", 7, 12
'   End If
'   For intCounter = 3 To 8
'      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
'         douTotal3(intCounter) = 0
'      Else
'         douTotal3(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
'      End If
'   Next intCounter
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'
'   'Add By Cheng 2002/01/18
'   PaintLine ReportSum(8)
'
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(4)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   adoaccrpt413.Fields("r41303").Value = ReportSum(11)
'   For intCounter = 3 To 8
'      If douTotal1(intCounter) - douTotal2(intCounter) + douTotal3(intCounter) = 0 Then
'         adoaccrpt413.Fields(intCounter).Value = Null
'      Else
'         adoaccrpt413.Fields(intCounter).Value = douTotal1(intCounter) - douTotal2(intCounter) + douTotal3(intCounter)
'      End If
'   Next intCounter
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   PaintLine ReportSum(8)
'   adoaccrpt413.UpdateBatch
'   adoaccrpt413.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt413Delete()
'   adoTaie.Execute "delete from accrpt413"
End Sub

''*************************************************
''  儲存資料表(部門損益比較表暫存檔)
''
''*************************************************
'Private Sub Accrpt413Save()
'Dim intCounter As Integer
'
'   adoaccrpt413.AddNew
'   adoaccrpt413.Fields("r41301").Value = strUserNum
'   adoaccrpt413.Fields("r41302").Value = Counter
'   If IsNull(adoacc010.Fields("a0102").Value) Then
'      adoaccrpt413.Fields("r41303").Value = Null
'   Else
'      adoaccrpt413.Fields("r41303").Value = adoacc010.Fields("a0102").Value
'   End If
'   If Mid(Combo2, 1, 1) = "1" Then
'      Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, 1, 6
'   Else
'      Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, 7, 12
'   End If
'   adoaccrpt413.UpdateBatch
'End Sub
'
''*************************************************
''  計算各月份小計金額
''
''*************************************************
'Private Sub Calculate(strAccNo1, strAccNo2 As String, intStartM, intEndM As Integer)
'Dim douDebit, douCredit As Double
'Dim intCounter, intMonth As Integer, strSql As String
'
'   If Text3 <> MsgText(601) Then
'      strSql = " and a0401 = " & Val(Text3) & ""
'   End If
'   If Text6 <> MsgText(601) Then
'      '2014/2/20 modify by sonia
'      'strSql = strSql & " and a0403 = '" & Text6 & "'"
'      strSql = strSql & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      '2014/2/20 end
'   End If
'   If strAccNo1 <> MsgText(601) Then
'      strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
'   End If
'   If strAccNo2 <> MsgText(601) Then
'      strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
'   End If
'   intCounter = 3
'   For intMonth = intStartM To intEndM
'      adoacc040.CursorLocation = adUseClient
'      '2008/1/14 modify by sonia 科目7XXX營業外收支應區分借貸方科目
'      'adoacc040.Open "select sum(a0408) from acc040 where a0402 = " & intMonth & " and a0404 = '" & MsgText(55) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      If strAccNo1 >= "7" And strAccNo2 <= "799999" Then
'         adoacc040.Open "select sum(decode(a0103,'1',a0408*-1,a0408)) from acc040,acc010 where a0405=a0101(+) and a0402 = " & intMonth & " and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      Else
'         adoacc040.Open "select sum(a0408) from acc040 where a0402 = " & intMonth & " and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      End If
'      '2008/1/14 end
'      If adoacc040.RecordCount <> 0 Then
'         If IsNull(adoacc040.Fields(0).Value) Then
'            adoaccrpt413.Fields(intCounter).Value = Null
'         Else
'            If adoacc040.Fields(0).Value = 0 Then
'               adoaccrpt413.Fields(intCounter).Value = Null
'            Else
'               adoaccrpt413.Fields(intCounter).Value = adoacc040.Fields(0).Value
'            End If
'         End If
'      Else
'         adoaccrpt413.Fields(intCounter).Value = Null
'      End If
'      intCounter = intCounter + 1
'      adoacc040.Close
'   Next intMonth
'End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   For intCounter = 3 To 8
      adoaccrpt413.Fields(intCounter).Value = strSign
   Next intCounter
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
   'Add by Amy 2018/03/07
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = MskFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = MskFormat
   'end 2018/03/07
'   Text6 = ""
'   Text7 = ""
   Text3 = ""
   'Combo2 = ""  '2014/2/20 cancel by sonia
'   Text6.SetFocus
   'Add By Sindy 2020/4/27
   CboCmp.ListIndex = -1
   CboCmp.SetFocus
   '2020/4/27 END
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2017/11/09 +Excel判斷
Public Function FormCheck() As Boolean
   Dim bolCancel As Boolean
   Dim strDateS As String, strDateE As String 'Add by Amy 2018/11/22
   
   FormCheck = False
   
   'Add by Sindy 2020/4/27 +公司別判斷
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bolCancel)
      If bolCancel = True Then
          Exit Function
      End If
   End If
   'end 2020/4/27
    
    If bolExcel = False Then
        If Text3 = MsgText(601) Then
            MsgBox "年度不可為空！", , MsgText(5)
            Exit Function
        End If
        If Combo2 = MsgText(601) Then
            MsgBox "請選擇年期！", , MsgText(5)
            Combo2.SetFocus
            Exit Function
        End If
    '產生Excel需判斷日期格式
    Else
        'Modify by Sindy 2020/4/27
        'If Trim(Replace(MaskEdBox1, "/", "")) = MsgText(601) Then
        If Val(FCDate(MaskEdBox1.Text)) = 0 Then
        '2020/4/27 END
            MsgBox "起始區間不可為空！", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
        MaskEdBox1_Validate (bolCancel)
        If bolCancel = True Then Exit Function
        
        'Modify by Sindy 2020/4/27
        'If Trim(Replace(MaskEdBox2, "/", "")) = MsgText(601) Then
        If Val(FCDate(MaskEdBox2.Text)) = 0 Then
        '2020/4/27 END
            MsgBox "截止區間不可為空！", , MsgText(5)
            MaskEdBox2.SetFocus
            Exit Function
        End If
        MaskEdBox2_Validate (bolCancel)
        If bolCancel = True Then Exit Function
        
        If Val(Replace(MaskEdBox1, "/", "")) > Val(Replace(MaskEdBox2, "/", "")) Then
            MsgBox "起始區間不可大於截止區間！", , MsgText(5)
            MaskEdBox2.SetFocus
            Exit Function
        End If
        
        'Add by Amy 2018/11/22 去年該當月累計,程式寫法不允許下跨年-已和婧瑄確認不會下跨年
        strDateS = Val(Replace(MaskEdBox1, "/", ""))
        strDateE = Val(Replace(MaskEdBox2, "/", ""))
        If Val(IIf(Len(strDateS) = 5, Left(strDateS, 3), Left(strDateS, 2))) <> Val(IIf(Len(strDateE) = 5, Left(strDateE, 3), Left(strDateE, 2))) Then
            MsgBox "起始區間不可跨年！", , MsgText(5)
            MaskEdBox2.SetFocus
            Exit Function
        End If
    End If
    FormCheck = True
End Function

'Add by Amy 2024/05/28 判斷L公司是否有414102會計科目且有值
Private Function HaveData(ByVal strWhere As String) As Boolean
   Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
   
   HaveData = False
   strQ = "Select * From Accrpt44c0 Where ID='" & strUserNum & "' And R001='1' " & strWhere
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      HaveData = True
   End If
   Set RsQ = Nothing
End Function
