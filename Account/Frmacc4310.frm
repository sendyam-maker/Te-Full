VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacc4310 
   AutoRedraw      =   -1  'True
   Caption         =   "傳票轉外帳作業"
   ClientHeight    =   3550
   ClientLeft      =   50
   ClientTop       =   350
   ClientWidth     =   5670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3550
   ScaleWidth      =   5670
   Begin VB.ComboBox CboComp 
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
      Text            =   "CboComp"
      Top             =   180
      Width           =   3520
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3210
      Width           =   4995
      _ExtentX        =   8819
      _ExtentY        =   459
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Top             =   1680
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行(&E)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1200
      Width           =   5000
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   11.5
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
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc4310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

'2005/8/3整理,2006/7/10 新規則
Public adoacc020 As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim stra0201 As String
Dim stra0202 As String
Dim stra0208 As String
Dim stra0211 As String
Dim lnga0205 As Long
Dim lnga0206 As Long
Dim lnga0207 As Long
Dim lnga0209 As Long
Dim lnga0210 As Long
Dim strax201 As String
Dim strax202 As String
Dim strax203 As String
Dim strax204 As String
Dim strax205 As String
Dim strax205W As String
Dim strax208 As String
Dim strax209 As String
Dim strax211 As String
Dim strax212 As String
Dim strax213 As String
Dim strax214 As String
Dim strax215 As String
Dim douax206 As Double
Dim douax207 As Double
Dim lngax210 As Long
Dim strSql As String
Dim strCom As String
Dim m_A1P26 As String    '2005/6/17 ADD BY SONIA
Dim strCaseNo As String
Dim strAccNo As String   '2006/7/12 ADD BY SONIA
Dim strCU10 As String    '2012/4/12 add by sonia 申請人國籍
Dim strNation As String  '2012/4/12 add by sonia 申請國家
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/4/17

'Add by Sindy 2020/4/17
Private Sub SetCompN()
    strCmpN = ""
    If Trim(cboComp) <> MsgText(601) Then
        strCmp = cboComp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

'Add by Sindy 2020/4/17
Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(cboComp) = MsgText(601) Then Exit Sub
    
    strCmp = cboComp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        cboComp.SetFocus
        Exit Sub
    ElseIf Len(Trim(cboComp)) = 1 Then
        cboComp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/4/17

Private Sub Command1_Click()
   '2005/11/22 ADD BY SONIA
   adoacc020.CursorLocation = adUseClient
   '2014/2/17 MODIFY BY SONIA加公司別條件
   'adoacc020.Open "select A0C05 from acc0C0 ", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2020/4/17
   Call SetCompN
   'If Text5 = "1" Then
'modify by sonia 2020/5/12 依公司別分別由內轉外,1->P,L->L,J->J
'   If strCmp = "1" Then
'   '2020/4/17 END
'      adoacc020.Open "select A0C05 from acc0C0 where a0c04<>'J' ", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc020.Open "select A0C05 from acc0C0 where a0c04='J' ", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'modify by sonia 2023/5/22 依公司別分別由內轉外,改以作帳公司的ACC0C0來判斷1->0->7,L->K->L,J->J
   adoacc020.Open "select A0C05 from acc0C0 where a0c04='" & IIf(strCmp = "1", "7", strCmp) & "' ", adoTaie, adOpenStatic, adLockReadOnly
'end 2020/5/12
   '2014/2/17 end
   If adoacc020.RecordCount <> 0 Then
      If IsNull(adoacc020.Fields("A0C05").Value) = False And adoacc020.Fields("A0C05").Value >= Val(FCDate(MaskEdBox1.Text)) Then
         adoacc020.Close
         Screen.MousePointer = vbDefault
         MsgBox "此日期傳票已轉至外帳, 不可重覆轉檔", , MsgText(21)
         Exit Sub
      End If
      If IsNull(adoacc020.Fields("A0C05").Value) = False And adoacc020.Fields("A0C05").Value >= Val(FCDate(MaskEdBox2.Text)) Then
         adoacc020.Close
         Screen.MousePointer = vbDefault
         MsgBox "此日期傳票已轉至外帳, 不可重覆轉檔", , MsgText(21)
         Exit Sub
      End If
   End If
   adoacc020.Close
   '2005/11/22 END
   '2014/2/17 MODIFY BY SONIA加公司別條件
   'Transfer
   'Modify By Sindy 2020/4/17
   'If Text5 = "1" Then
'modify by sonia 2020/5/12 依公司別分別由內轉外,1->P,L->L(同1公司規則),J->J
'   If strCmp = "1" Then
'   '2020/4/17 END
'      Transfer
'   Else
'      TransferJcomp
'   End If
'modify by sonia 2020/7/21 6公司未結束前,1公司仍維持轉0公司,故將1,L公司分開寫
   'If strCmp <> "J" Then
   '   TransferNew
   'Else
   '   TransferJcomp
   'End If
   Select Case strCmp
      Case "1"
         Transfer
      Case "L"
         TransferLcomp
      Case "J"
         TransferJcomp
   End Select
'end 2020/7/21
'end 2020/5/12
   '2014/2/17 end
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5790
   Me.Height = 3975
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Text5 = "":   Text6 = "" '2014/2/17 ADD BY SONIA
   'Add by Sindy 2020/4/17
   cboComp.Clear
   cboComp.AddItem "", 0
   Call Pub_SetCboCmp(cboComp, False, False, False, , 1)
   'end 2020/4/17
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4310 = Nothing
End Sub


'*************************************************
'  將傳票主檔資料放置系統變數中
'
'*************************************************
Private Sub Acc030Save()
   stra0201 = adoacc020.Fields("a0201").Value
   stra0202 = adoacc020.Fields("a0202").Value
   If IsNull(adoacc020.Fields("a0205").Value) Then
      lnga0205 = 0
   Else
      lnga0205 = adoacc020.Fields("a0205").Value
   End If
   If IsNull(adoacc020.Fields("a0206").Value) Then
      lnga0206 = 0
   Else
      lnga0206 = adoacc020.Fields("a0206").Value
   End If
   If IsNull(adoacc020.Fields("a0207").Value) Then
      lnga0207 = 0
   Else
      lnga0207 = adoacc020.Fields("a0207").Value
   End If
   If IsNull(adoacc020.Fields("a0208").Value) Then
      stra0208 = MsgText(601)
   Else
      stra0208 = adoacc020.Fields("a0208").Value
   End If
   If IsNull(adoacc020.Fields("a0209").Value) Then
      lnga0209 = 0
   Else
      lnga0209 = adoacc020.Fields("a0209").Value
   End If
   If IsNull(adoacc020.Fields("a0210").Value) Then
      lnga0210 = 0
   Else
      lnga0210 = adoacc020.Fields("a0210").Value
   End If
   If IsNull(adoacc020.Fields("a0211").Value) Then
      stra0211 = MsgText(601)
   Else
      stra0211 = adoacc020.Fields("a0211").Value
   End If
End Sub

'*************************************************
'  將傳票交易檔資料放置系統變數中
'
'*************************************************
Private Sub Acc031Save()
'add by nickc 2007/02/08
Dim strTemp1
Dim strJamt   As String   'add by sonia 2016/8/18
Dim strJax212 As String   'add by sonia 2016/8/18
   
   '2012/4/12 ADD BY SONIA 先依案號抓申請人國籍及申請國家
   strCU10 = "": strNation = ""
   If IsNull(adoacc021.Fields("ax214").Value) Then
      strCaseNo = "ZZZZZZZZZZZZ"
   Else
      strCaseNo = adoacc021.Fields("ax214").Value
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
      Case "P", "CFP", "FCP"        '專利
         '因D096011232之CFP011423001EPC指定國家故改抓PA04='00'
         adoquery.Open "select NVL(cu10,FA10) cu10, pa09 as nation from patent,customer,FAGENT where pa01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and pa02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and pa03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and pa04 = '00' and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and substr(pa75,1,8)=FA01(+) and substr(pa75,9,1)=FA02(+) ", adoTaie, adOpenStatic, adLockReadOnly
      Case "T", "TF", "CFT", "FCT"  '商標
         adoquery.Open "select NVL(cu10,FA10) cu10, tm10 as nation from trademark,customer,FAGENT where tm01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and tm02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and tm03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and tm04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and substr(TM44,1,8)=FA01(+) and substr(TM44,9,1)=FA02(+) ", adoTaie, adOpenStatic, adLockReadOnly
      Case "L", "CFL", "FCL"        '法務
         adoquery.Open "select NVL(cu10,FA10) cu10, lc15 as nation from lawcase,customer,FAGENT where lc01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and lc02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and lc03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and lc04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and substr(LC22,1,8)=FA01(+) and substr(LC22,9,1)=FA02(+) ", adoTaie, adOpenStatic, adLockReadOnly
      Case "LA"                     '顧問
         adoquery.Open "select cu10, '000' as nation from HIREcase,customer where Hc01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and Hc02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and Hc03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and Hc04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and substr(HC05,1,8)=cu01(+) and substr(HC05,9,1)=cu02(+) ", adoTaie, adOpenStatic, adLockReadOnly
      Case Else                     '服務
         adoquery.Open "select NVL(cu10,FA10) cu10, SP09 as nation from SERVICEPRACTICE,customer,FAGENT where SP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and SP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and SP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and SP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and substr(SP08,1,8)=cu01(+) and substr(SP08,9,1)=cu02(+) and substr(SP26,1,8)=FA01(+) and substr(SP26,9,1)=FA02(+) ", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   If adoquery.RecordCount <> 0 Then
      strCU10 = "" & adoquery.Fields("cu10").Value
      strNation = "" & adoquery.Fields("nation").Value
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   '2012/4/12 END

   strax201 = adoacc021.Fields("ax201").Value
   strax202 = adoacc021.Fields("ax202").Value
   strax203 = adoacc021.Fields("ax203").Value
   strax204 = "TOT"
   '2012/4/13 從最下面移上來,因為6130的F5542要判斷
   If IsNull(adoacc021.Fields("ax208").Value) Then
      strax208 = MsgText(601)
   Else
      strax208 = adoacc021.Fields("ax208").Value
   End If
   '2012/4/13 END
   If IsNull(adoacc021.Fields("ax212").Value) Then
      strax212 = MsgText(601)
   Else
      strax212 = Replace(adoacc021.Fields("ax212").Value, "'", "''")
      '2008/4/23 add by sonia 取消 （客戶出名） 字樣
      strax212 = Replace(adoacc021.Fields("ax212").Value, "（客戶出名）", "")
      '2008/4/23 end
      'add by sonia 2025/2/26 1公司、L公司費用科目之摘要取消 補助 字樣
      If (strax201 = "1" Or strax201 = "L") And Mid(adoacc021.Fields("ax205").Value, 1, 1) = "6" Then
         strax212 = Replace(strax212, "補助", "")
      End If
      'end 2025/2/26
   End If
   If IsNull(adoacc021.Fields("ax205").Value) Then
      strax205 = MsgText(601)
   Else
      Select Case Mid(adoacc021.Fields("ax205").Value, 1, 1)
         Case "6"
            strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
      End Select
      '2005/8/3 ADD BY SONIA : FCL 案件 摘要 X 字後面的都不要 strax212='FCL010380000/請款/X09311276/USD991.00'
      If Not IsNull(adoacc021.Fields("ax214").Value) And Mid(adoacc021.Fields("ax214").Value, 1, 3) = "FCL" Then
         If InStr(strax212, "X") > 0 Then
            strax212 = Left(strax212, InStr(strax212, "X") - 2)
         End If
      End If
      '2005/8/3 END
      
      Select Case Mid(adoacc021.Fields("ax205").Value, 1, 4)
         Case "2201"
            strax205 = adoacc021.Fields("ax205").Value
            If adoacc021.Fields("A1P02") = "I" Then
               strax205 = "6136"
            End If
            '2006/7/20 ADD BY SONIA 抵帳借方
            If adoacc021.Fields("A1P02") = "K" And adoacc021.Fields("AX206").Value > 0 Then
               strax205 = "6136"
            End If
            '2006/7/20 END
            '2012/4/12 ADD BY SONIA 抵帳貸方D101031271的P097900000退費,瑞婷說2201科目申請國家非台灣都改為6136
            '2012/10/16 modify by sonia 抵帳規費貸方改判斷a1p23(D101090908為抵帳收款,抵帳程式也配合修改放a1p23)
            'If adoacc021.Fields("A1P02") = "K" And adoacc021.Fields("AX207").Value > 0 And strNation <> "000" Then
            '   strax205 = "6136"
            If adoacc021.Fields("A1P02") = "K" And adoacc021.Fields("AX207").Value > 0 Then
               If Left(adoacc021.Fields("A1P23"), 1) <> "X" Or IsNull(adoacc021.Fields("A1P23")) Then strax205 = "6136"
            '2012/10/16 end
            End If
            '2012/4/12 END
            'add by sonia 2016/8/17 J公司結匯或抵帳之收款傳票,摘要之金額要除以1.05 (D105010030)
            If (strax201 = "J" And adoacc021.Fields("AX206").Value > 0) And (adoacc021.Fields("A1P02") = "I" Or adoacc021.Fields("A1P02") = "K") Then
               If Not IsNull(strax212) Then
                  strJamt = Mid(strax212, (InStr(strax212, "/")) + 1, InStr(strax212, " ") - InStr(strax212, "/"))
                  strJamt = Format(Val(Format(strJamt, "########0")) / 1.05, "###,###,##0")
                  strax212 = Left(strax212, (InStr(strax212, "/"))) + strJamt + Mid(strax212, (InStr(strax212, " ")))
               End If
            End If
            'end 2016/8/17
         Case "6120"
            strax205 = "6119"
         Case "4141"
            If strCmp <> "L" Then    'ADD BY SONIA 2020/5/13 L公司不轉換
               If IsNull(adoacc021.Fields("ax214").Value) = False Then
                  strax205 = adoacc021.Fields("ax205").Value         '92.10.18 ADD BY SONIA
                  Select Case Mid(adoacc021.Fields("ax214").Value, 1, 3)
                     Case "FCT"
                        strax205 = "4101"
                     Case "FCP"
                        strax205 = "4111"
                  End Select
               Else
                  strax205 = "4111"
               End If
            'add by sonia 2020/5/13
            Else
               strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
            End If
            'end 2020/5/13
         Case "4171"
            '2009/4/17 MODIFY BY SONIS
            'strax205 = "4111"
            'modify by sonia 2016/8/1 因加417104,417105,417109,故改寫法
            'If adoacc021.Fields("ax205").Value = "417101" Then
            '   strax205 = "4111"   '專利
            'Else
            '   strax205 = "4101"   '大陸做商標收入
            'End If
            Select Case adoacc021.Fields("ax205").Value
               'modify by sonia 2019/1/19 +417103,D107122447
               Case "417101", "417104", "417105", "417109", "417103"
                  strax205 = "4111"   '專利
               Case Else
                  strax205 = "4101"   '大陸做商標收入
            End Select
            'end 2016/8/1
            '2009/4/17 END
         Case "4161"
            If strCmp <> "L" Then   'ADD BY SONIA 2020/5/13 L公司不轉換
               Select Case Mid(adoacc021.Fields("AX214").Value, 1, 3)
                  Case "FCT"
                     strax205 = "4101"
                  Case Else
                     strax205 = "4111"
               End Select
            'add by sonia 2020/5/13
            Else
               strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
            End If
            'end 2020/5/13
         Case "4172"
            strax205 = "4101"
         'add by sonia 2023/12/12 預算收文收款240602改回收入科目D112110719
         Case "2406"
            If adoacc021.Fields("ax205").Value = "240602" And Val(adoacc021.Fields("ax207").Value) > 0 Then
               If Mid(adoacc021.Fields("AX214").Value, 1, 3) = "PS" Or Mid(adoacc021.Fields("AX214").Value, 1, 3) = "CPS" Then
                  strax205 = "4111"
               Else
                  strax205 = "4101"
               End If
            Else
               strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
            End If
         'end 2023/12/12
         Case Else
            If adoacc021.Fields("ax205").Value = "610103" Then
               If IsNull(adoacc021.Fields("a1p26").Value) Or adoacc021.Fields("a1p26").Value <> MsgText(602) Then
                  If IsNull(adoacc021.Fields("a1p22").Value) Then
                     strax205 = "0002"
                  Else
                     strax205 = adoacc021.Fields("ax205").Value
                  End If
               Else
                  strax205 = adoacc021.Fields("ax205").Value
               End If
            Else
               strax205 = adoacc021.Fields("ax205").Value
            End If
      End Select
      
      'add by sonia 2024/11/18 FCP之OA委外翻譯，結匯傳票扣收入，轉至外帳收入改為6130(D113041277)
      If (Mid(strax205, 1, 4) = "4171" Or Mid(strax205, 1, 4) = "4111") And adoacc021.Fields("A1P02") = "I" And adoacc021.Fields("AX206").Value > 0 Then
         If InStr(strax212, "OA委外翻譯") > 0 Then strax205 = "6130"
      End If
      'end 2024/11/18
      
      '收入科目及7121都做相同處理
      If (strax205 >= "4" And strax205 < "5") Or Mid(strax205, 1, 4) = "7121" Then
         '2006/7/12 ADD BY SONIA
         '摘要最後'/+數字'者,將'/+數字'刪除
         If Not IsNull(strax212) And InStr(strax212, "/") > 0 Then
            strTemp1 = Split(strax212, "/")
            If IsNumeric(strTemp1(UBound(strTemp1))) = True Then
               strax212 = Mid(strax212, 1, (Len(strax212) - (Len(strTemp1(UBound(strTemp1))) + 1)))
            End If
         End If
         '2006/7/12 END
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2011/11/17 考慮拆收據情形改語法,但若同一次收款同一案號的收據若有一個以上的公司別時抓號碼小的
         'adoquery.Open "select a0m03, a0k11, a0k23, a0k01, cp01||cp02 as CaseNo, a0k30 from acc1p0, acc0m0, acc0k0, caseprogress where a1p04 = a0m01 and a0m02 = a0k01 and a0m02 = cp60 (+) and a1p02 = 'A' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         'modify by sonia 2020/5/14 +substr(a0j02, 1, Length(a0j02) - 9) cp01
         adoquery.Open "select a0m03, a0k11, a0k23, a0k01,substr(a0j02, 1, Length(a0j02) - 9) cp01 from acc1p0, acc0m0, acc0k0,acc0j0 where a1p04 = a0m01 and a0m02 = a0k01 and a0j13(+)=a0k01 and a0j02=a1p17 and a1p17='" & strCaseNo & "' and a1p02 = 'A' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' order by a0k01", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a0m03").Value) = False Then
               Select Case Mid(adoquery.Fields("a0m03").Value, 1, 1)
                  Case "E"
                     If IsNull(adoquery.Fields("a0k11").Value) = False Then
                        Select Case adoquery.Fields("a0k11").Value
                           Case "1"
                              strax205 = "4101"
                           Case "2"
                              strax205 = "4111"
                           Case "3"
                              strax205 = "4181"
                           Case "5"
                              strax205 = "4183"
                           Case "6"
                              strax205 = "4184"
                           Case "7"
                              strax205 = "4185"
                           Case "8"
                              strax205 = "4186"
                           '2006/8/18 ADD BY SONIA
                           Case "9"
                              strax205 = "4201"
                           '2006/8/18 END
                           'add by sonia 2020/5/14
                           Case "L"
                              If adoquery.Fields("cp01").Value = "CFL" Then
                                 strax205 = "4161"
                              Else
                                 strax205 = "4141"
                              End If
                           'end 2020/5/14
                        End Select
                     End If
                  Case Else
                     strax205 = "4201"    '2006/7/11 科目0003->4201,作帳公司9
               End Select
            Else
               If IsNull(adoquery.Fields("a0k11").Value) = False Then
                  Select Case adoquery.Fields("a0k11").Value
                     Case "1"
                        strax205 = "4101"
                     Case "2"
                        strax205 = "4111"
                     Case "3"
                        strax205 = "4181"
                     Case "5"
                        strax205 = "4183"
                     Case "6"
                        strax205 = "4184"
                     Case "7"
                        strax205 = "4185"
                     Case "8"
                        strax205 = "4186"
                     '2006/8/18 ADD BY SONIA
                     Case "9"
                        strax205 = "4201"
                     '2006/8/18 END
                     'add by sonia 2020/5/14
                     Case "L"
                        If adoquery.Fields("cp01").Value = "CFL" Then
                           strax205 = "4161"
                        Else
                           strax205 = "4141"
                        End If
                     'end 2020/5/14
                  End Select
               End If
            End If
         '2005/8/10 ADD BY SONIA
         Else
            Select Case adoacc021.Fields("a1P02").Value
               Case "F", "K" '2006/7/18 加入 K抵帳
                  Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                     Case "T", "TF", "CFT", "FCT"
                        If strNation <> "000" Then
                           strax205 = "4101"
                        End If
                  ''''''''''''
                     Case "P", "CFP", "FCP"
                        If strNation <> "000" Then
                           '2006/7/28 MODIFY BY SONIA 改CFP規則
                           'strax205 = "4101"
                           If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" Then
                              strax205 = "4111"
                              Select Case strNation
                                 '2006/10/16 MODIFY BY SONIA
                                 'Case "011"
                                 '   If Mid(strCaseNo, Len(strCaseNo) - 8, 6) >= "011051" Then
                                 '      strax205 = "4101"
                                 '   End If
                                 'Case "221"
                                 '   If Mid(strCaseNo, Len(strCaseNo) - 8, 6) >= "011051" Or Mid(strCaseNo, Len(strCaseNo) - 8, 6) < "016183" Then
                                 '      strax205 = "4101"
                                 '   End If
                                 Case "011", "221"
                                    strax205 = GetACCNO(strNation, strCaseNo, strax205)
                                 '2006/10/16 END
                              End Select
                           Else
                              '2006/10/16 ADD BY SONIA 改P規則
                              If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" Then
                                 strax205 = GetACCNO(strNation, strCaseNo, strax205)
                              Else
                              '2006/10/16 END
                                 strax205 = "4101"
                              End If
                           End If
                        End If
                     Case "L", "CFL", "FCL"
                        If strNation <> "000" Then
                           strax205 = "4101"
                        End If
                     Case "LA"
                     Case Else
                        If strNation <> "000" Then
                           strax205 = "4101"
                        End If
                  End Select
               '2005/11/17 ADD BY SONIA 銷退傳票也做和收入傳票相同處理
               Case "Z"
                  If adoquery.State = adStateOpen Then
                     adoquery.Close
                  End If
                  adoquery.CursorLocation = adUseClient
                  'Modified by Morgan 2011/11/17 考慮拆收據情形改語法,但若同一次收款同一案號的收據若有一個以上的公司別時抓號碼小的
                  'adoquery.Open "select a0m03, a0k11, a0k23, a0k01, cp01||cp02 as CaseNo, a0k30 from acc1p0, acc0m0, acc0k0, caseprogress ,ACC0S0 where a1p04 = a0S01 and a0S02 = a0k01 and a0S02 = cp60 (+) AND A0S02=A0M02 and a1p02 = 'Z' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
                  'modify by sonia 2020/5/14 +substr(a0j02, 1, Length(a0j02) - 9) cp01
                  adoquery.Open "select a0m03, a0k11, a0k23, a0k01,substr(a0j02, 1, Length(a0j02) - 9) cp01 from acc1p0, acc0m0, acc0k0 ,ACC0S0,acc0j0 where a1p04 = a0S01 and a0S02 = a0k01 AND A0S02=A0M02 and a0j13(+)=a0k01 and a0j02=a1p17 and a1p17='" & strCaseNo & "' and a1p02 = 'Z' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' order by a0k01", adoTaie, adOpenStatic, adLockReadOnly
                  If adoquery.RecordCount <> 0 Then
                     If IsNull(adoquery.Fields("a0m03").Value) = False Then
                        Select Case Mid(adoquery.Fields("a0m03").Value, 1, 1)
                           Case "E"
                              If IsNull(adoquery.Fields("a0k11").Value) = False Then
                                 Select Case adoquery.Fields("a0k11").Value
                                    Case "1"
                                       strax205 = "4101"
                                    Case "2"
                                       strax205 = "4111"
                                    Case "3"
                                       strax205 = "4181"
                                    Case "5"
                                       strax205 = "4183"
                                    Case "6"
                                       strax205 = "4184"
                                    Case "7"
                                       strax205 = "4185"
                                    Case "8"
                                       strax205 = "4186"
                                    '2006/8/18 ADD BY SONIA
                                    Case "9"
                                       strax205 = "4201"
                                    '2006/8/18 END
                                    'add by sonia 2020/5/14
                                    Case "L"
                                       If adoquery.Fields("cp01").Value = "CFL" Then
                                          strax205 = "4161"
                                       Else
                                          strax205 = "4141"
                                       End If
                                    'end 2020/5/14
                                 End Select
                              End If
                           Case Else
                              strax205 = "4201"
                        End Select
                     Else
                        If IsNull(adoquery.Fields("a0k11").Value) = False Then
                           Select Case adoquery.Fields("a0k11").Value
                              Case "1"
                                 strax205 = "4101"
                              Case "2"
                                 strax205 = "4111"
                              Case "3"
                                 strax205 = "4181"
                              Case "5"
                                 strax205 = "4183"
                              Case "6"
                                 strax205 = "4184"
                              Case "7"
                                 strax205 = "4185"
                              Case "8"
                                 strax205 = "4186"
                              '2006/8/18 ADD BY SONIA
                              Case "9"
                                 strax205 = "4201"
                              '2006/8/18 END
                              'add by sonia 2020/5/14
                              Case "L"
                                 If adoquery.Fields("cp01").Value = "CFL" Then
                                    strax205 = "4161"
                                 Else
                                    strax205 = "4141"
                                 End If
                              'end 2020/5/14
                           End Select
                        End If
                     End If
                  End If
            'add by sonia 2023/11/16 預算收入的沖銷收入傳票A1P02='P',收入科目改0001
            Case "P"
               If Left(adoacc021.Fields("ax205").Value, 1) = "4" Then
                  strax205 = "0001"
               End If
            'end 2023/11/16
               '2005/11/17 END
            End Select
         '2005/8/10 END
         End If
         adoquery.Close
      Else
         strax205W = ""
         'modify by sonia 2022/9/20 再加L公司的2403科目
         'If Mid(adoacc021.Fields("ax205").Value, 1, 4) = "2201" Then
         If Mid(adoacc021.Fields("ax205").Value, 1, 4) = "2201" Or (adoacc021.Fields("ax201").Value = "L" And Mid(adoacc021.Fields("ax205").Value, 1, 4) = "2403") Then
            '2006/7/20 ADD BY SONIA 貸方摘要最後'/+數字'者,將'/+數字'刪除
            If adoacc021.Fields("ax207").Value > 0 Then
               strTemp1 = Split(strax212, "/")
               If IsNumeric(strTemp1(UBound(strTemp1))) = True Then
                  strax212 = Mid(strax212, 1, (Len(strax212) - (Len(strTemp1(UBound(strTemp1))) + 1)))
               End If
            End If
            '2006/7/20 END
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
            adoquery.CursorLocation = adUseClient
            'Modified by Morgan 2011/11/17 考慮拆收據情形改語法,但若同一次收款同一案號的收據若有一個以上的公司別時抓號碼小的
            'adoquery.Open "select a0m03, a0k11, a0k23, a0k01, cp01||cp02 as CaseNo, a0k30 from acc1p0, acc0m0, acc0k0, caseprogress where a1p04 = a0m01 and a0m02 = a0k01 and a0m02 = cp60 (+) and a1p02 = 'A' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2020/5/14 +substr(a0j02, 1, Length(a0j02) - 9) cp01
            adoquery.Open "select a0m03, a0k11, a0k23, a0k01, substr(a0j02, 1, Length(a0j02) - 9) cp01 from acc1p0, acc0m0, acc0k0,acc0j0 where a1p04 = a0m01 and a0m02 = a0k01 and a0j13(+)=a0k01 and a0j02=a1p17 and a1p17='" & strCaseNo & "' and a1p02 = 'A' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' order by a0k01", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If IsNull(adoquery.Fields("a0m03").Value) = False Then
                  Select Case Mid(adoquery.Fields("a0m03").Value, 1, 1)
                     Case "E"
                        If IsNull(adoquery.Fields("a0k11").Value) = False Then
                           Select Case adoquery.Fields("a0k11").Value
                              Case "1"
                                 strax205W = "4101"
                              Case "2"
                                 strax205W = "4111"
                              Case "3"
                                 strax205W = "4181"
                              Case "5"
                                 strax205W = "4183"
                              Case "6"
                                 strax205W = "4184"
                              Case "7"
                                 strax205W = "4185"
                              Case "8"
                                 strax205W = "4186"
                              '2006/8/18 ADD BY SONIA
                              Case "9"
                                 strax205W = "4201"
                              '2006/8/18 END
                               'add by sonia 2020/5/14
                              Case "L"
                                 If adoquery.Fields("cp01").Value = "CFL" Then
                                    strax205W = "4161"
                                 Else
                                    strax205W = "4141"
                                 End If
                              'end 2020/5/14
                          End Select
                        End If
                     Case Else
                        strax205W = "4201"    '2006/7/11 科目0003->4201,作帳公司9
                  End Select
               Else
                  If IsNull(adoquery.Fields("a0k11").Value) = False Then
                     Select Case adoquery.Fields("a0k11").Value
                        Case "1"
                           strax205W = "4101"
                        Case "2"
                           strax205W = "4111"
                        Case "3"
                           strax205W = "4181"
                        Case "5"
                           strax205W = "4183"
                        Case "6"
                           strax205W = "4184"
                        Case "7"
                           strax205W = "4185"
                        Case "8"
                           strax205W = "4186"
                        '2006/8/18 ADD BY SONIA
                        Case "9"
                           strax205W = "4201"
                        '2006/8/18 END
                        'add by sonia 2020/5/14
                        Case "L"
                           If adoquery.Fields("cp01").Value = "CFL" Then
                              strax205W = "4161"
                           Else
                              strax205W = "4141"
                           End If
                        'end 2020/5/14
                     End Select
                  End If
               End If
            Else
               Select Case adoacc021.Fields("a1P02").Value
                  Case "I", "K"   '2006/7/13 MODIFY BY SONIA 收據開發票之結匯或抵帳其作帳公司9,非發票之結匯才依收據公司決定作帳公司
                     strCom = ""  '2012/11/2 ADD BY SONIA 為下面判斷用,所以先清空
                     If adoquery.State = adStateOpen Then
                        adoquery.Close
                     End If
                     adoquery.CursorLocation = adUseClient
                     'Modified by Morgan 2011/11/17 考慮拆收據情形改語法,但若同一次收款同一案號的收據若有一個以上的公司別時抓號碼小的
                     'adoquery.Open "select A0M03,A0K11 from acc1P0,ACC190,ACC151,CASEPROGRESS,ACC0K0,ACC0M0 where a1p02 = 'I' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and A1P17 = '" & strCaseNo & "' and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' AND SUBSTR(A1P04,1,INSTR(A1P04, 'Y',1)-1)=A1908(+) AND A1902=AXF01(+) AND AXF02=CP09(+) AND CP60=A0K01 AND CP60=A0M02(+) ", adoTaie, adOpenStatic, adLockReadOnly
                     adoquery.Open "select A0M03,A0K11 from acc1P0,ACC190,ACC151,acc0j0,ACC0K0,ACC0M0 where a1p02 = 'I' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and A1P17 = '" & strCaseNo & "' AND SUBSTR(A1P04,1,INSTR(A1P04, 'Y',1)-1)=A1908(+) AND A1902=AXF01(+) AND AXF02=a0j01(+) and a0j02=a1p17 and a0m02(+)=a0j13 and a0k01=a0j13 order by a0k01", adoTaie, adOpenStatic, adLockReadOnly
                     If adoquery.RecordCount <> 0 Then
                        If IsNull(adoquery.Fields("a0m03").Value) = False Then
                           If Mid(adoquery.Fields("a0m03").Value, 1, 1) = "E" Then
                              strCom = adoquery.Fields("A0K11").Value
                           Else
                              strCom = "9"
                           End If
                        Else
                           strCom = adoquery.Fields("A0K11").Value
                        End If
                     Else
                        '2012/11/2 ADD BY SONIA 無收據者先判斷是否該案號其他收文號有收據 D101031825(CFP-022865)
                        If adoquery.State = adStateOpen Then
                           adoquery.Close
                        End If
                        adoquery.CursorLocation = adUseClient
                        'modify by sonia 2025/5/6 SUBSTR(CP60,1,1)='E'改為SUBSTR(CP60,1,1) is not null(CFP-020941最新收文是開請款單)
                        adoquery.Open "select SUBSTR(MAX(CP05||A0K11), 9) from caseprogress,acc0j0,acc0k0 where a0j01(+)=cp09 and a0k01(+)=a0j13 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '00' AND SUBSTR(CP60,1,1) is not null", adoTaie, adOpenStatic, adLockReadOnly
                        If adoquery.RecordCount <> 0 Then
                           If IsNull(adoquery.Fields(0).Value) = False Then
                              strCom = adoquery.Fields(0).Value
                           End If
                        End If
                        If strCom = "" Then
                        '2012/11/2 END
                           '國外請款時,商標或大陸,香港案件為1公司,其他為2公司
                           Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                              Case "P", "CFP", "FCP"        '專利
                                 If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" And (strNation = "011" Or strNation = "221")) Then
                                    strCom = GetCOMPNO(strNation, strCaseNo)
                                 Else
                                    If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" And strNation <> "000" Then
                                       strCom = GetCOMPNO(strNation, strCaseNo)
                                    Else
                                       strCom = "2"
                                    End If
                                 End If
                                 '2010/10/19 ADD BY SONIA D099052947
                                 If strNation <> "000" Then
                                    m_A1P26 = "Y"
                                 End If
                                 '2010/10/19 END
                              Case "PS", "CPS", "FG"        '專利服務
                                 If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CPS" And (strNation = "011" Or strNation = "221")) Then
                                    strCom = GetCOMPNO(strNation, strCaseNo)
                                 Else
                                    If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "PS" And strNation <> "000" Then
                                       strCom = GetCOMPNO(strNation, strCaseNo)
                                    Else
                                       strCom = "2"
                                    End If
                                 End If
                                 '2010/10/19 ADD BY SONIA D099052947
                                 If strNation <> "000" Then
                                    m_A1P26 = "Y"
                                 End If
                                 '2010/10/19 END
                              Case "T", "TF", "CFT", "FCT"  '商標
                                 strCom = "1"
                                 '2010/10/19 ADD BY SONIA D099052947
                                 If strNation <> "000" Then
                                    m_A1P26 = "Y"
                                 End If
                                 '2010/10/19 END
                              Case "L", "CFL", "FCL"        '法務
                                 strCom = "2"
                                 '2010/10/19 ADD BY SONIA D099052947
                                 If strNation <> "000" Then
                                    m_A1P26 = "Y"
                                 End If
                                 '2010/10/19 END
                              Case "LA"                     '顧問
                                 strCom = "2"
                              Case Else                     '商標服務
                                 strCom = "1"
                                 '2010/10/19 ADD BY SONIA D099052947
                                 If strNation <> "000" Then
                                    m_A1P26 = "Y"
                                 End If
                                 '2010/10/19 END
                           End Select
                        End If   '2012/11/2 ADD BY SONIA
                     End If
                  '2005/6/17 ADD BY SONIA 國外部接非台灣案, 規費併入收入
                  Case "F"
                     strCom = "7" 'add by sonia 2021/8/26 1公司D110071624第015項次先收規費會沒有公司別
                     Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                        Case "T", "TF", "CFT", "FCT"
                           strax205W = "4101"
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "P", "CFP", "FCP"
                           strax205W = "4111"
                           If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" And (strNation = "011" Or strNation = "221")) Then
                              strCom = GetCOMPNO(strNation, strCaseNo)
                           Else
                              If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" And strNation <> "000" Then
                                 strCom = GetCOMPNO(strNation, strCaseNo)
                              Else
                                 strCom = "2"
                              End If
                           End If
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "L", "CFL", "FCL"
                           strax205W = "4111"
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "LA"
                           strax205W = "4111"
                        Case Else
                           '2006/7/14 MODIFY BY SONIA
                           'strax205W = "4101"
                           Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                              Case "PS", "CPS", "FG"    '專利服務
                                 strax205W = "4111"
                                 '2006/10/16 ADD BY SONIA
                                 If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CPS" And (strNation = "011" Or strNation = "221")) Then
                                    strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                                 Else
                                    If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "PS" And strNation <> "000" Then
                                       strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                                    End If
                                 End If
                                 '2006/10/16 END
                              Case Else
                                 strax205W = "4101"
                           End Select
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                     End Select
                  '2005/6/17 END
                     '2007/3/27 ADD BY SONIA 國外請款依公司別設定科目
                     Select Case strCom
                        Case "1"
                           strax205W = "4101"
                        Case "2"
                           strax205W = "4111"
                        Case "3"
                           strax205W = "4181"
                        Case "5"
                           strax205W = "4183"
                        Case "6"
                           strax205W = "4184"
                        Case "7"
                           strax205W = "4185"
                        Case "8"
                           strax205W = "4186"
                        Case "9"
                           strax205W = "4201"
                     End Select
                     '2007/3/27 END
                  '2005/11/17 ADD BY SONIA
                  Case "Z"
                     If adoquery.State = adStateOpen Then
                        adoquery.Close
                     End If
                     adoquery.CursorLocation = adUseClient
                     'Modified by Morgan 2011/11/17 已不需抓CP資料,但若多案合併開收據會有可能包含合併及不合併的收文資料,目前先抓最小的收文號設定
                     'adoquery.Open "select a0m03, a0k11, a0k23, a0k01, cp01||cp02 as CaseNo, a0k30 from acc1p0, acc0m0, acc0k0, caseprogress,ACC0S0 where a1p04 = a0S01 and a0S02 = a0k01 and a0S02 = cp60 (+) AND A0S02=A0M02 and a1p02 = 'Z' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
                     'modify by sonia 2020/5/14 +substr(a0j02, 1, Length(a0j02) - 9) cp01
                     adoquery.Open "select a0m03, a0k11, a0k23, a0k01, a0j07,substr(a0j02, 1, Length(a0j02) - 9) cp01 from acc1p0, acc0m0, acc0k0,ACC0S0,acc0j0 where a1p04 = a0S01 and a0S02 = a0k01 and a0j13(+)=a0k01 AND A0S02=A0M02 and a1p02 = 'Z' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' order by a0j01", adoTaie, adOpenStatic, adLockReadOnly
                     If adoquery.RecordCount <> 0 Then
                        If IsNull(adoquery.Fields("a0j07").Value) = False And adoquery.Fields("a0j07").Value = MsgText(602) Then
                           m_A1P26 = adoquery.Fields("a0j07").Value
                        End If
                        If IsNull(adoquery.Fields("a0k11").Value) = False Then
                           Select Case adoquery.Fields("a0k11").Value
                              Case "1"
                                 strax205W = "4101"
                              Case "2"
                                 strax205W = "4111"
                              Case "3"
                                 strax205W = "4181"
                              Case "5"
                                 strax205W = "4183"
                              Case "6"
                                 strax205W = "4184"
                              Case "7"
                                 strax205W = "4185"
                              Case "8"
                                 strax205W = "4186"
                              '2006/8/18 ADD BY SONIA
                              Case "9"
                                 strax205W = "4201"
                              '2006/8/18 END
                              'add by sonia 2020/5/14
                              Case "L"
                                 If adoquery.Fields("cp01").Value = "CFL" Then
                                    strax205W = "4161"
                                 Else
                                    strax205W = "4141"
                                 End If
                              'end 2020/5/14
                           End Select
                        End If
                     End If
                  '2006/7/14 ADD BY SONIA
                  Case Else
                     Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                        Case "T", "TF", "CFT", "FCT"
                           strax205W = "4101"
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "P", "CFP", "FCP"
                           strax205W = "4111"
                           '2006/10/16 ADD BY SONIA
                           If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" And (strNation = "011" Or strNation = "221")) Then
                              strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                           Else
                              If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" And strNation <> "000" Then
                                 strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                              End If
                           End If
                           '2006/10/16 END
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "L", "CFL", "FCL"
                           strax205W = "4111"
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                        Case "LA"
                           strax205W = "4111"
                        Case Else
                           '2006/7/14 MODIFY BY SONIA
                           'strax205W = "4101"
                           Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
                              Case "PS", "CPS", "FG"    '專利服務
                                 strax205W = "4111"
                                 '2006/10/16 ADD BY SONIA
                                 If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CPS" And (strNation = "011" Or strNation = "221")) Then
                                    strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                                 Else
                                    If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "PS" And strNation <> "000" Then
                                       strax205W = GetACCNO(strNation, strCaseNo, strax205W)
                                    End If
                                 End If
                                 '2006/10/16 END
                              Case Else
                                 strax205W = "4101"
                           End Select
                           If strNation <> "000" Then
                              m_A1P26 = "Y"
                           End If
                     End Select
               '2005/11/17 END
               End Select
            End If
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
         End If
      End If
      '2012/2/14 ADD BY SONIA 王雅萍F5542的翻譯費6130要抓作帳公司,抓該案號翻譯的收據(無翻譯抓新申請案的收據)D101012596
      If Mid(strax205, 1, 4) = "6130" And strax208 = "F5542" Then
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select 1,a0m03,a0k11,cp05,cp09 from caseprogress,acc0m0,acc0k0 where cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp10 in ('201','209','210') and substr(cp60,1,1)='E' and cp60=a0m02(+) and cp60=a0k01(+) union " & _
                       "select 2,a0m03,a0k11,cp05,cp09 from caseprogress,acc0m0,acc0k0 where cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp31='Y' and substr(cp60,1,1)='E' and cp60=a0m02(+) and cp60=a0k01(+) " & _
                       "order by 1,cp05,cp09", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a0m03").Value) = False Then
               If Mid(adoquery.Fields("a0m03").Value, 1, 1) = "E" Then
                  strCom = adoquery.Fields("A0K11").Value
               Else
                  strCom = "9"
               End If
            Else
               strCom = adoquery.Fields("A0K11").Value
            End If
         Else
            '國外請款時大陸,香港案件為1公司,其他為2公司
            Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
               Case "P", "CFP", "FCP"        '專利
                  If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" And (strNation = "011" Or strNation = "221")) Then
                     strCom = GetCOMPNO(strNation, strCaseNo)
                  Else
                     If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" And strNation <> "000" Then
                        strCom = GetCOMPNO(strNation, strCaseNo)
                     Else
                        strCom = "2"
                     End If
                  End If
               Case "PS", "CPS", "FG"        '專利服務
                  If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CPS" And (strNation = "011" Or strNation = "221")) Then
                     strCom = GetCOMPNO(strNation, strCaseNo)
                  Else
                     If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "PS" And strNation <> "000" Then
                        strCom = GetCOMPNO(strNation, strCaseNo)
                     Else
                        strCom = "2"
                     End If
                  End If
               Case "T", "TF", "CFT", "FCT"  '商標
                  strCom = "1"
               Case "L", "CFL", "FCL"        '法務
                  strCom = "2"
               Case "LA"                     '顧問
                  strCom = "2"
               Case Else                     '商標服務
                  strCom = "1"
            End Select
         End If
      '江蘇舜禹Y52268的翻譯費抓結匯資料且a1p04有Y52268000者
      ElseIf Mid(strax205, 1, 4) = "6130" Then
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select SUBSTR(A1P04,INSTR(A1P04, 'Y',1)) fagentNO from acc1P0 where a1p02 = 'I' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and A1P17 = '" & strCaseNo & "' and a1p05='6130' order by a1p03", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("fagentno")) = True Then GoTo Nextstep
            'If adoquery.Fields("fagentno") <> "Y52268000" Then GoTo Nextstep   '2012/9/5 cancel by sonia 瑞婷說只要Y都要加作帳公司
         '2012/3/8 ADD BY SONIA D101022599 不該加作帳公司
         Else
            GoTo Nextstep
         '2012/3/8 END
         End If
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select A0M03,A0K11 from acc1P0,ACC190,ACC151,acc0j0,ACC0K0,ACC0M0 where a1p02 = 'I' and a1p01 = '" & strax201 & "' and a1p22 = '" & strax202 & "' and A1P17 = '" & strCaseNo & "' AND SUBSTR(A1P04,1,INSTR(A1P04, 'Y',1)-1)=A1908(+) AND A1902=AXF01(+) AND AXF02=a0j01(+) and a0j02=a1p17 and a0m02(+)=a0j13 and a0k01=a0j13 order by a0k01", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a0m03").Value) = False Then
               If Mid(adoquery.Fields("a0m03").Value, 1, 1) = "E" Then
                  strCom = adoquery.Fields("A0K11").Value
               Else
                  strCom = "9"
               End If
            Else
               strCom = adoquery.Fields("A0K11").Value
            End If
         Else
            '國外請款時大陸,香港案件為1公司,其他為2公司
            Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
               Case "P", "CFP", "FCP"        '專利
                  If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CFP" And (strNation = "011" Or strNation = "221")) Then
                     strCom = GetCOMPNO(strNation, strCaseNo)
                  Else
                     If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "P" And strNation <> "000" Then
                        strCom = GetCOMPNO(strNation, strCaseNo)
                     Else
                        strCom = "2"
                     End If
                  End If
               Case "PS", "CPS", "FG"        '專利服務
                  If (Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "CPS" And (strNation = "011" Or strNation = "221")) Then
                     strCom = GetCOMPNO(strNation, strCaseNo)
                  Else
                     If Mid(strCaseNo, 1, Len(strCaseNo) - 9) = "PS" And strNation <> "000" Then
                        strCom = GetCOMPNO(strNation, strCaseNo)
                     Else
                        strCom = "2"
                     End If
                  End If
               Case "T", "TF", "CFT", "FCT"  '商標
                  strCom = "1"
               Case "L", "CFL", "FCL"        '法務
                  strCom = "2"
               Case "LA"                     '顧問
                  strCom = "2"
               Case Else                     '商標服務
                  strCom = "1"
            End Select
         End If
      End If
      '2012/2/14 END
      
      'add by sonia 2020/11/27 1公司收入科目及規費科目改新規則
      'modify by sonia 2023/11/16 配合預算收入的沖銷收入傳票A1P02='P',收入科目改0001加strax205 <> "0001"條件
      'If strax201 = "1" And Left(adoacc021.Fields("ax205").Value, 1) = "4" Then
      If strax201 = "1" And Left(adoacc021.Fields("ax205").Value, 1) = "4" And strax205 <> "0001" Then
         '4141改4111
         If Mid(adoacc021.Fields("ax205").Value, 1, 4) = "4141" Then
            strax205 = "4111"
         '4161改4112
         ElseIf Mid(adoacc021.Fields("ax205").Value, 1, 4) = "4161" Then
            strax205 = "4112"
         '無案號者抓前4碼，到帳務自行處理
         ElseIf IsNull(adoacc021.Fields("ax214").Value) Then
            strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
         Else
            '依本所案號轉科目
            Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
               Case "FCT"
                  strax205 = "4102"
               Case "ACS", "CFP", "CPS", "P", "PS"
                  strax205 = "4111"
               Case "FCP", "FG"
                  strax205 = "4112"
               '法務案抓前4碼 D109110266
               Case "L", "LA", "CFL", "FCL", "LIN"
                  'modify by sonia 2023/5/25 瑞婷說一律改4112 (D112050101之010項次)
                  'strax205 = Mid(adoacc021.Fields("ax205").Value, 1, 4)
                  strax205 = "4112"
               Case Else
                  strax205 = "4101"
            End Select
         End If
      'modify by sonia 2023/4/25 J公司也要判斷D110060016
      'ElseIf strax201 = "1" And Mid(adoacc021.Fields("ax205").Value, 1, 4) = "2201" And adoacc021.Fields("ax207").Value > 0 Then
      ElseIf (strax201 = "1" Or strax201 = "J") And Mid(adoacc021.Fields("ax205").Value, 1, 4) = "2201" And adoacc021.Fields("ax207").Value > 0 Then
         '220113改4111
         If adoacc021.Fields("ax205").Value = "220113" Then
            strax205W = "4111"
         '無案號者抓前4碼，到帳務自行處理
         ElseIf IsNull(adoacc021.Fields("ax214").Value) Then
            strax205W = Mid(adoacc021.Fields("ax205").Value, 1, 4)
         Else
            '依本所案號轉科目
            Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
               Case "FCT"
                  strax205W = "4102"
               Case "ACS", "CFP", "CPS", "P", "PS"
                  strax205W = "4111"
               Case "FCP", "FG"
                  strax205W = "4112"
               Case Else
                  strax205W = "4101"
            End Select
         End If
      End If
      'Debug.Print "strax205=" & strax205 & " strax205W=" & strax205W
      'end 2020/11/27
   End If
   
Nextstep:
   If IsNull(adoacc021.Fields("ax206").Value) Then
      douax206 = 0
   Else
      douax206 = Int(adoacc021.Fields("ax206").Value)
   End If
   If IsNull(adoacc021.Fields("ax207").Value) Then
      douax207 = 0
   Else
      douax207 = Int(adoacc021.Fields("ax207").Value)
   End If
   If IsNull(adoacc021.Fields("ax209").Value) Then
      strax209 = MsgText(601)
   Else
      strax209 = adoacc021.Fields("ax209").Value
   End If
   If IsNull(adoacc021.Fields("ax210").Value) Then
      lngax210 = 0
   Else
      lngax210 = adoacc021.Fields("ax210").Value
   End If
   If IsNull(adoacc021.Fields("ax211").Value) Then
      strax211 = MsgText(601)
   Else
      strax211 = adoacc021.Fields("ax211").Value
   End If
   If IsNull(adoacc021.Fields("ax213").Value) Then
      strax213 = MsgText(601)
   Else
      strax213 = adoacc021.Fields("ax213").Value
   End If
   If IsNull(adoacc021.Fields("ax214").Value) Then
      strax214 = MsgText(601)
   Else
      strax214 = adoacc021.Fields("ax214").Value
   End If
End Sub

'*************************************************
'  將內帳1公司傳票資料轉外帳0公司
'
'*************************************************
Private Sub Transfer()
Dim strNo As String
Dim strDocNo As String
Dim lngEff As Long    '2005/11/24 ADD BY SONIA

'add by nickc 2007/02/08
Dim StrSQLa As String

On Error GoTo Checking
   strSql = ""
   strNo = ""
   strAccNo = ""
   Text1 = ""
   Screen.MousePointer = vbHourglass
   adoTaie.Execute "delete from acc031 where ax301 = '0' and ax302 in (select a0302 from acc030 where a0301 = '0' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "')"
   adoTaie.Execute "delete from acc030 where a0301 = '0' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "'"
   adoTaie.Execute "delete from acc1z0 where a1z01 = '0' and a1z02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1z03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1z03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   
   Text1 = "正在將內帳1公司傳票轉至外帳0公司......"
   ProgressBar1.Value = 0
   adoacc020.CursorLocation = adUseClient
   
   '2014/2/18 modify by sonia 只抓1公司
   adoacc020.Open "select * from acc020 where a0201='1' and a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   '測試某傳票時用
   'adoacc020.Open "select * from acc020 where a0201='1' and A0202 IN ('D112060006') order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount <> 0 Then
      ProgressBar1.max = adoacc020.RecordCount
   End If
   Do While adoacc020.EOF = False
      DoEvents    '2012/4/20 add by sonia
      
      'add by sonia 2019/10/15 剔除結餘結算傳票及每月25號的固定傳票
      If adoquery.State = adStateOpen Then adoquery.Close
      adoquery.CursorLocation = adUseClient
      'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日.但瑞婷說張文嘉的440要傳故加入入AXD02
      'adoquery.Open "select distinct a1p02,a1p22,a1p18 from acc1p0 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select distinct a1p02,a1p22,axd03,AXD02 from acc1p0,acc0d1 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "' And A1p01=Axd01(+) And decode(a1p02,'U',Substr(A1p04,1,Length(A1p04)-7),null)=Axd02(+)", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         '結餘結算傳票
         If adoquery.Fields("a1p02").Value = "S" Then GoTo NextRecord
         '每月25號的固定傳票
         'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日,但瑞婷說張文嘉的440要傳
         'If adoquery.Fields("a1p02").Value = "U" And "" & adoquery.Fields("a1p18").Value = "25" Then GoTo NextRecord
         If adoquery.Fields("a1p02").Value = "U" And "" & adoquery.Fields("axd03").Value = "25" And "" & adoquery.Fields("axd02").Value <> "440" Then GoTo NextRecord
      End If
      adoquery.Close
      'end 2019/10/15
      
      Acc030Save
      '2007/6/1 MODIFY BY SONIA 新增日期人員時間改存操作人及系統時間
      'adoTaie.Execute "insert into acc030 (a0301, a0302, a0305, a0306, a0307, a0308) values ('0', " & CNULL(stra0202) & ", " & lnga0205 & ", " & lnga0206 & ", " & lnga0207 & ", " & CNULL(stra0208) & ")"
      adoTaie.Execute "insert into acc030 (a0301, a0302, a0305, a0306, a0307, a0308) values ('0', " & CNULL(stra0202) & ", " & lnga0205 & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
      adoacc021.CursorLocation = adUseClient
      '2014/2/18 modify by sonia 只抓1公司
      adoacc021.Open "select * from acc021, acc1p0 where ax201=a1p01(+) and ax202 = a1p22 (+) and ax203 = a1p03 (+) and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc021.EOF = False
         'Debug.Print stra0202 & "  " & adoacc021.Fields("aX203").Value
         strCom = ""
         '2005/6/17 ADD BY SONIA
         m_A1P26 = ""
         If IsNull(adoacc021.Fields("a1p26").Value) = False And adoacc021.Fields("a1p26").Value = MsgText(602) Then
            m_A1P26 = adoacc021.Fields("a1p26").Value
         End If
         '2005/6/17 END
         '2005/12/23 ADD BY SONIA
         If IsNull(adoacc021.Fields("a1p14").Value) = False And InStr(1, adoacc021.Fields("a1p14").Value, "代理人請款", 1) > 0 Then
            m_A1P26 = "Y"
         End If
         '2005/12/23 END
         Acc031Save
         If strNo <> strax202 Then
            Select Case Mid(strax205, 1, 1)
               Case "6"
                  If IsNull(adoacc021.Fields("ax215").Value) = False Then
                     strCom = adoacc021.Fields("ax215").Value     '費用科目作帳公司抓 AX215
                  End If
            End Select
            strNo = strax202
         End If
         strax215 = strCom
         '2006/7/13 ADD BY SONIA 改公司1->6,2->7
         If strax215 = "1" Then strax215 = "7"   'modify by sonia 2020/12/14 6公司不用了也改7公司
         If strax215 = "2" Then strax215 = "7"
         '2006/7/13 END
         strax205 = Mid(strax205, 1, 4)   '所有科目改為4碼
         '2006/3/30 ADD BY SONIA 固定改科目
         Select Case strax205
            '2006/7/12 MODIFY BY SONIA
            'Case "4141", "4161"
            '   strax205 = "4111"
            Case "4141"
               strax205 = "4111"
            Case "4161"
               strax205 = "4112"
            '2006/7/12 END
            Case "6102"
               strax205 = "6101"
            Case "6120"
               strax205 = "6119"
            Case "6124"
               strax205 = "6111"
            'add by sonia 2025/2/26 1105,1106,1911,1912,1913科目且為貸方時才改1101現金
            Case "1105", "1106", "1911", "1912", "1913"
               If douax207 > 0 Then strax205 = "1101"
            'end 2025/2/26
         End Select
         '2006/3/30 END
         Select Case strax205
            '2010/6/11 MODIFY BY SONIA 加TC-10505科目4151
            'Case "4101", "4102"
            Case "4101", "4102", "4151"
               strax215 = "7"   '2006/7/12 1->6  'modify by sonia 2020/12/14 6公司不用了也改7公司
            Case "4111", "4112"
               strax215 = "7"   '2006/7/12 2->7
            Case "4181"
               strax215 = "3"
            Case "4183"
               strax215 = "5"
            Case "4186"
               strax215 = "8"
            '2006/7/11 ADD BY SONIA
            Case "4201"
               strax215 = "9"
            '2006/7/11 END
         End Select
         '2006/7/12 ADD BY SONIA
         '外對內之收入科目改為4102或4112, 無申請人抓代理人 D095010545
         'If (strax205 >= "4" And strax205 < "5") And douax207 > 0 And strCaseNo <> "ZZZZZZZZZZZZ" Then
         'modify by sonia 2020/12/9 此段內部2201的控制已取消故取消2201的條件,另2020/12/3辜郵件FMP,FMT借方也要改科目故取消douax207 > 0條件
         'If ((strax205 >= "4" And strax205 < "5") Or strax205 = "2201") And douax207 > 0 And strCaseNo <> "ZZZZZZZZZZZZ" Then
         If strax205 >= "4" And strax205 < "5" And strCaseNo <> "ZZZZZZZZZZZZ" Then
            '申請人國籍非台灣者,收入科目改4102或4112
            If strCU10 > "010" Then
               '2006/7/27 ADD BY SONIA 智權人員為國外部才要改科目,規費科目因無對沖智權人員所以在最後更新  D095011762 不轉,D095010834要轉
               'Select Case strax205
               '   Case "4101"
               '      strax205 = "4102"
               '   Case "4111"
               '      strax205 = "4112"
               '   '2006/7/26 ADD BY SONIA
               '   Case "2201"
               '      strax205 = "2202"
               '   '2006/7/26 END
               'End Select
               If strax209 <> "" Then
                  adoquery.CursorLocation = adUseClient
                  adoquery.Open "select ST15 from STAFF where ST01 = '" & strax209 & "' ", adoTaie, adOpenStatic, adLockReadOnly
                  If Mid(adoquery.Fields("ST15").Value, 1, 1) = "F" Then
                     Select Case strax205
                        Case "4101"
                           strax205 = "4102"
                        Case "4111"
                           strax205 = "4112"
                     End Select
                  End If
                  adoquery.Close
               End If
               '2006/7/27 END
            End If
         End If
         '2006/7/12 END
         
         '2006/8/10 取消AX311
         '2007/3/5 加 AX316
         strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                  "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                  ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
         '考慮規費是否合併至收入
         If strax205 = "2201" And douax207 > 0 Then
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
            adoquery.CursorLocation = adUseClient
            '2005/6/23 MODIFY BY SONIA 相同案號才可合併
            'adoquery.Open "select ax305 from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "')", adoTaie, adOpenStatic, adLockReadOnly
            '2006/7/14 MODIFY BY SONIA 收入必須為貸方才抓 D095010336 不抓
            'adoquery.Open "select ax305,AX315 from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "')", adoTaie, adOpenStatic, adLockReadOnly
            adoquery.Open "select ax305,AX315,AX303 from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
            '2005/6/23 END
            If adoquery.RecordCount <> 0 Then
               If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                  If m_A1P26 = MsgText(602) Then
                     '2005/11/17 MODIFY BY SONIA
                     'strSQL = "update acc031 set ax307 = ax307 + " & douax207 & " where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "')"
                     '2006/8/18 MODIFY BY SONIA 改直接更新抓到的項次
                     'strSQL = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "')"
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  '2006/7/14 ADD BY SONIA 不合併時規費作帳公司同收入
                  Else
                     strax215 = "" & adoquery.Fields("ax315").Value
                     '2006/8/10 取消AX311
                     '2007/3/5  加入AX316
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
                  '2206/7/14 END
               Else
                  Select Case strax205W
                     Case "4101"
                        strax215 = "7"   '2006/7/12 1->6  'modify by sonia 2020/12/14 6公司不用了也改7公司
                     Case "4111", "4141", "4161"
                        strax215 = "7"   '2006/7/12 2->7
                     Case "4181"
                        strax215 = "3"
                     Case "4183"
                        strax215 = "5"
                     Case "4186"
                        strax215 = "8"
                     '2006/7/11 ADD BY SONIA
                     Case "4201"
                        strax215 = "9"
                     '2006/7/11 END
                  End Select
                  If strax205W <> "" And m_A1P26 = MsgText(602) Then
                     '2006/8/10 取消AX311
                     '2007/3/5  加入AX316
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
                  Else
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
               End If
            '2005/6/23 ADD BY SONIA
            Else
               Select Case strax205W
                  Case "4101"
                     strax215 = "7"   '2006/7/12 1->6  'modify by sonia 2020/12/14 6公司不用了也改7公司
                  Case "4111", "4141", "4161"
                     strax215 = "7"   '2006/7/12 2->7
                  Case "4181"
                     strax215 = "3"
                  Case "4183"
                     strax215 = "5"
                  Case "4186"
                     strax215 = "8"
                  '2006/7/11 ADD BY SONIA
                  Case "4201"
                     strax215 = "9"
                  '2006/7/11 END
               End Select
               If strax205W <> "" And m_A1P26 = MsgText(602) Then
                  '2006/8/10 取消AX311
                  '2007/3/5  加入AX316
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
               Else
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
               End If
            '2005/6/23 END
            End If
            adoquery.Close
         End If
         '2005/6/28 ADD BY SONIA 4101,4111,4121,4131其他對沖為結餘X者,科目改0004,作帳公司NULL
         '2012/4/13 再加收入科目,借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
         Select Case strax205
            Case "4101", "4111", "4102", "4112", "4121", "4131"    '2006/7/12加4102,4112
               If IsNull(strax213) = False And Mid(strax213, 1, 2) = "結餘" Then
                  '2006/8/10 取消AX311
                  '2007/3/5  加入AX316
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
               End If
               '2012/4/13 add by sonia 借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
               If douax206 + douax207 = 5000 And (InStr(strax212, "點作轉專業") > 0 Or InStr(strax212, "支援") > 0) Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('0', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
              End If
               '2012/4/13 end
            Case Else
         End Select
         '2011/4/18 add by sonia 非結餘的收款傳票414101法務收入科目,智權人員為M0100總所者此為複委託,應併入前一收入項次D100010157,要抓acc1p0才能確定為414101
         If strax205 >= "4" And strax205 < "5" And IsNull(strax213) = False And Mid(strax213, 1, 2) <> "結餘" Then
            If "" & adoacc021.Fields("a1p02").Value & adoacc021.Fields("a1p05").Value & adoacc021.Fields("a1p16").Value = "A414101M0100" Then
               If adoquery.State = adStateOpen Then
                  adoquery.Close
               End If
               adoquery.CursorLocation = adUseClient
               adoquery.Open "select ax305,AX315,AX303 from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
               If adoquery.RecordCount <> 0 Then
                  If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '0' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  End If
               End If
               adoquery.Close
            End If
         End If
         '2011/4/18 end
         
         '2005/6/28 END
         If strDocNo <> (adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value) Then
            adoTaie.Execute strSql
            strDocNo = adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value
         End If
         adoacc021.MoveNext
      Loop
      adoacc021.Close

NextRecord:   'add by sonia 2019/10/15
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
'add by sonia 2023/5/25 款項匯至華銀傳票(借方110208,貸方收入或規費,但傳票無1133科目者)
'1.刪除摘要有(匯差)二字的項次
'2.611301科目放作帳公司7公司
'3.差額調整貸方最小項次的收入
   Text1 = "正在處理款項匯至華銀傳票...."
   '1.刪除摘要有(匯差)二字的項次
   strSql = "DELETE ACC031 WHERE INSTR(AX312,'匯差')>0 AND (AX301,AX302) IN " & _
            "(SELECT DISTINCT '0',AX202 FROM ACC021 WHERE (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021 WHERE AX207>0 AND (AX205 LIKE '4%' OR AX205 LIKE '2201%') AND (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='110208' AND AX206>0)) AND (AX201,AX202) NOT IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='1133'))"
   adoTaie.Execute strSql
   '2.611301科目放作帳公司7公司
   strSql = "UPDATE ACC031 SET AX315='7' WHERE AX305='6113' AND (AX301,AX302) IN " & _
            "(SELECT DISTINCT '0',AX202 FROM ACC021 WHERE (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021 WHERE AX207>0 AND (AX205 LIKE '4%' OR AX205 LIKE '2201%') AND (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='110208' AND AX206>0)) AND (AX201,AX202) NOT IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='1133'))"
   adoTaie.Execute strSql
   '3.差額調整貸方最小項次的收入
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "SELECT AX301,AX302,SUM(AX306-AX307) FROM ACC031 WHERE (AX301,AX302) IN " & _
            "(SELECT DISTINCT '0',AX202 FROM ACC021 WHERE (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021 WHERE AX207>0 AND (AX205 LIKE '4%' OR AX205 LIKE '2201%') AND (AX201,AX202) IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='110208' AND AX206>0)) AND (AX201,AX202) NOT IN " & _
            "(SELECT DISTINCT AX201,AX202 FROM ACC021,ACC020 WHERE A0201='1' AND A0205>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=AX201(+) AND A0202=AX202(+) " & _
            "AND AX205='1133')) GROUP BY AX301,AX302 HAVING SUM(AX306-AX307)<>0", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      DoEvents
      If IsNull(adoacc020.Fields(2).Value) = False Then
         If adoacc020.Fields(2).Value <> 0 Then
            strSql = "update acc031 set ax307 = ax307 + " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 IN (" & _
                     "SELECT MIN(AX303) FROM ACC031 where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax307 > 0 AND SUBSTR(AX305,1,1)='4' )"
            adoTaie.Execute strSql
         End If
      End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
'end 2023/5/25

'cancel by sonia 2018/6/14 辜通知取消
'   '2006/7/12 ADD BY SONIA
'   '借方有110205或110222科目且金額>150之傳票,該傳票的借方增加6113雜費金額150元,第一個110205或110222減少150元
'   '6113之作帳公司若有該傳票有作帳公司7則6113作帳公司放7,否則放6
'   '2013/10/1 modify by sonia 150元改為120元**********
'   Text1 = "正在處理借方有110205或110222科目" & vbCrLf & "    且金額>120之傳票,該傳票的借方增加" & vbCrLf & "    6113雜費且金額120元時,則第一個" & vbCrLf & "    110205或110222減少120元......"
'   ProgressBar1.Value = 0
'   adoacc020.CursorLocation = adUseClient
'   '2014/2/18 modify by sonia 只抓1公司
'   adoacc020.Open "select DISTINCT AX201,AX202 from ACC021 where AX205 IN ('110205','110222') AND AX206>120 AND (AX201,AX202) IN ( " & _
'                  "select A0201,A0202 from acc020,ACC021 where A0202=AX202 AND AX205>='4' AND AX205<='5' AND AX207>0 AND A0201='1' AND a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "') GROUP BY AX201,AX202 order by AX201,AX202 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '測試某傳票時用
'   'adoacc020.Open "select DISTINCT AX202 from ACC021 where AX205 IN ('110205','110222') AND AX206>150 AND (AX201,AX202) IN ( " & _
'   '               "select A0201,A0202 from acc020,ACC021 where A0202 IN ('D095022377','D095010812','D095010834','D095010837','D095011547','D095011762','D095011957','D095012404','D095012414') AND A0202=AX202 AND AX205>='4' AND AX205<='5' AND AX207>0 AND A0201='1' AND a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "') GROUP BY AX201,AX202 order by AX201,AX202 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc020.RecordCount <> 0 Then
'      ProgressBar1.max = adoacc020.RecordCount
'   End If
'   Do While adoacc020.EOF = False
'      DoEvents    '2012/4/20 add by sonia
'      'Debug.Print adoacc020.Fields("AX202")
'      '減少借方第一個110205或110222金額150
'      '2013/10/1 modify by sonia 150元改為120元**********
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select MIN(AX203) from acc021 where ax201 = " & CNULL(adoacc020.Fields("AX201")) & " and ax202 = " & CNULL(adoacc020.Fields("AX202")) & " and ax205 IN ('110205','110222') and ax206>120 ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strSql = "UPDATE ACC031 SET AX306=AX306-120 WHERE AX301='0' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & " AND AX303=" & CNULL(adoquery.Fields(0)) & ""
'         adoTaie.Execute strSql
'      End If
'      adoquery.Close
'
'      '為要新增001項次,所以先將每一項次+1
'      strSql = "UPDATE ACC031 SET AX303=SUBSTR((AX303+1001),2,3) WHERE AX301='0' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & ""
'      adoTaie.Execute strSql
'
'      '再新增001項次6113雜項借方金額150,並決定作帳公司別
'      '2013/10/1 modify by sonia 150元改為120元**********
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select * from acc031 where AX301='0' AND ax302 = " & CNULL(adoacc020.Fields("AX202")) & " and ax305>='4' AND AX305<'5' AND AX307>0 AND AX315='7' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strax215 = "7"
'      Else
'         strax215 = "6"
'      End If
'      adoquery.Close
'      '2007/3/2  加入AX316
'      strSql = "insert into acc031 (ax301, ax302, ax303, AX304, ax305, AX316, ax306, ax307, AX312, ax315) " & _
'               "values ('0', " & CNULL(adoacc020.Fields("AX202")) & ", '001', 'TOT', '6113', '6113', 120, 0, '手續費', " & CNULL(strax215) & ")"
'      adoTaie.Execute strSql
'
'      ProgressBar1.Value = ProgressBar1.Value + 1
'      adoacc020.MoveNext
'   Loop
'   adoacc020.Close
'   '2006/7/12 END
'end  2018/6/14 辜通知取消
   
   '2006/7/26 ADD BY SONIA
   '自行輸入之借方6113且金額為150者,掛作帳公司,貸方有7公司者則放7,否則放6  D095010837
   '2013/10/1 modify by sonia 150元改為120元**********
   Text1 = "正在將自行輸入之借方6113且金額為120者" & vbCrLf & "    ,掛作帳公司,貸方有7公司者則放7公司," & vbCrLf & "    否則放6公司......"
   strSql = "UPDATE ACC031 B SET B.AX315= " & _
            "(SELECT MAX(A.AX315) FROM ACC031 A WHERE B.AX301=A.AX301 AND B.AX302=A.AX302 AND A.AX307>0 AND A.AX315 IS NOT NULL) " & _
            "WHERE (B.AX301,B.AX302,B.AX303) IN ( " & _
            "SELECT AX301,AX302,AX303 FROM ACC030,ACC031 WHERE A0301='0' AND A0305>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301 AND A0302=AX302 AND AX306=120 AND AX315 IS NULL AND AX305='6113')"
   adoTaie.Execute strSql
   
'cancel by sonia 2018/6/14 辜通知取消
'   '2013/10/2 ADD BY SONIA
'   '借方有110205或110222科目且金額>120,且貸方有2401且金額與借方相同之傳票,該傳票的借方增加6113雜費金額120元(作帳公司放7公司),第一個110205或110222減少120元
'   Text1 = "正在處理借方有110205或110222科目" & vbCrLf & "    且金額>120,且貸方有2401且金額與借" & vbCrLf & "    方相同之傳票,該傳票的借方增加" & vbCrLf & "    6113雜費且金額120元時,則第一個" & vbCrLf & "    110205或110222減少120元......"
'   ProgressBar1.Value = 0
'   adoacc020.CursorLocation = adUseClient
'   '2014/2/18 modify by sonia 只抓1公司
'   adoacc020.Open "SELECT DISTINCT A.AX201,A.AX202 from ACC020,ACC021 A,ACC021 B WHERE A0201='1' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=A.AX201(+) AND A0202=A.AX202(+) AND A.AX205 IN ('110205','110222') AND A.AX206>120 " & _
'                  "AND A.AX201=B.AX201(+) AND A.AX202=B.AX202(+) AND SUBSTR(B.AX205,1,4)='2401' AND A.AX206=B.AX207", adoTaie, adOpenStatic, adLockReadOnly
'   '測試某傳票時用
'   'adoacc020.Open "SELECT DISTINCT A.AX201,A.AX202 from ACC020,ACC021 A,ACC021 B WHERE A0201='1' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=A.AX201(+) AND A0202=A.AX202(+) " & _
'                   "AND A0202 IN ('D102090805','D102091120','D102091795','D102091872') AND A.AX205 IN ('110205','110222') AND A.AX206>120 " & _
'                   "AND A.AX201=B.AX201(+) AND A.AX202=B.AX202(+) AND SUBSTR(B.AX205,1,4)='2401' AND A.AX206=B.AX207", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc020.RecordCount <> 0 Then
'      ProgressBar1.max = adoacc020.RecordCount
'   End If
'   Do While adoacc020.EOF = False
'      DoEvents
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select MIN(AX203) from acc021 where ax201 = " & CNULL(adoacc020.Fields("AX201")) & " and ax202 = " & CNULL(adoacc020.Fields("AX202")) & " and ax205 IN ('110205','110222') and ax206>120 ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strSql = "UPDATE ACC031 SET AX306=AX306-120 WHERE AX301='0' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & " AND AX303=" & CNULL(adoquery.Fields(0)) & ""
'         adoTaie.Execute strSql
'      End If
'      adoquery.Close
'
'      '為要新增001項次,所以先將每一項次+1
'      strSql = "UPDATE ACC031 SET AX303=SUBSTR((AX303+1001),2,3) WHERE AX301='0' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & ""
'      adoTaie.Execute strSql
'
'      '再新增001項次6113雜項借方金額120,作帳公司別放7
'      strSql = "insert into acc031 (ax301, ax302, ax303, AX304, ax305, AX316, ax306, ax307, AX312, ax315) " & _
'               "values ('0', " & CNULL(adoacc020.Fields("AX202")) & ", '001', 'TOT', '6113', '6113', 120, 0, '手續費', '7')"
'      adoTaie.Execute strSql
'
'      ProgressBar1.Value = ProgressBar1.Value + 1
'      adoacc020.MoveNext
'   Loop
'   adoacc020.Close
'   '2013/10/2 END
'end  2018/6/14 辜通知取消
   
   '內帳借方有110204,110205,110222,113002科目之傳票,該傳票的貸方之收入及規費改為4102,4112,2202      D095010812
   Text1 = "正在將內帳借方有110204,110205,110222" & vbCrLf & "    ,113002科目之傳票,該傳票的貸方之收入" & vbCrLf & "    及規費改為4102,4112,2202......"
   '2014/2/18 modify by sonia 只抓1公司
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) WHERE AX301='0' AND AX307>0 AND AX302 IN ( " & _
            "SELECT A0202 FROM acc020,ACC021 WHERE A0201=AX201 AND A0202=AX202 AND A0201='1' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX205 IN ('110204','110205','110222','113002') AND AX206>0 GROUP BY A0202)"
   adoTaie.Execute strSql
   
   '內帳借方有113001科目且對沖客戶為V0001之傳票,該傳票的貸方之所有科目都改回0004科目且不放作帳公司      D095011547~48
   Text1 = "正在將內帳借方有113001科目且對沖客戶" & vbCrLf & "    為V0001之傳票,該傳票的貸方之所有科目" & vbCrLf & "    都改回0004科目且不放作帳公司......"
   '2014/2/18 modify by sonia 只抓1公司
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='0' AND AX307>0 AND AX302 IN ( " & _
            "SELECT A0202 FROM acc020,ACC021 WHERE A0201=AX201 AND A0202=AX202 AND A0201='1' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX205='113001' AND AX206>0 AND AX208='V0001' GROUP BY A0202)"
   adoTaie.Execute strSql
   
   '傳票有2211或2491科目不管借方或貸方,對方之收入及規費都改回0004科目且不放作帳公司      D095021629~30
   Text1 = "正在將傳票有2211或2491科目不管借方或貸方" & vbCrLf & "    ,對方之收入及規費都改回0004科目" & vbCrLf & "    且不放作帳公司......"
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='0' AND AX305>='2201' AND AX305<='2202' AND AX302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='0' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('2211','2491') GROUP BY A0302)"
   adoTaie.Execute strSql
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='0' AND AX305>='4' AND AX305<'5' AND AX302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='0' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('2211','2491') GROUP BY A0302)"
   adoTaie.Execute strSql
   
   '貸方有4102,4112科目之傳票,該傳票的貸方之收入及規費都改為4102,4112,2202
   Text1 = "正在將貸方有4102,4112科目之傳票,該傳票的" & vbCrLf & "    貸方之收入及規費都改為" & vbCrLf & "    4102,4112,2202......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) where AX301='0' and ((AX305>='4' AND AX305<'5') OR AX305='2201') and ax307>0 AND ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='0' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('4102','4112') AND AX307>0 GROUP BY A0302)"
   adoTaie.Execute strSql
   
   '借方有1915或1916科目之傳票,該傳票的貸方之收入及規費都改回原科目        D095012414
   Text1 = "正在將借方有1915或1916科目之傳票,該傳票" & vbCrLf & "    的貸方之收入及規費都改回原科目......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2202','2201','4102','4101','4112','4111',AX305)) where AX301='0' and ((AX305>='4' AND AX305<'5') OR AX305='2202') and ax307>0 AND ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='0' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('1915','1916') AND AX306>0 GROUP BY A0302)"
   adoTaie.Execute strSql
   '2006/7/26 END
   
   '2007/5/17 ADD BY SONIA
   '科目為6114機油費之傳票,該項次摘要改成 汽車加油
   Text1 = "正在將科目為6114機油費之傳票,該項次摘要" & vbCrLf & "    改成 汽車加油......"
   strSql = "UPDATE ACC031 SET AX312='汽車加油' where AX301='0' AND AX305='6114' and ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='0' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' GROUP BY A0302)"
   adoTaie.Execute strSql
   '2007/5/17 END
   
   Text1 = "正在調整因小數位捨去所造成的借貸方" & vbCrLf & "    差額......"
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select ax301, ax302, sum(ax307 - ax306) from acc031, acc030 where A0301='0' AND ax301 = a0301 and ax302 = a0302 and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by ax301, ax302", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      DoEvents    '2012/4/20 add by sonia
      If IsNull(adoacc020.Fields(2).Value) = False Then
         '2005/11/24 MODIFY BY SONIA 因小數位捨去會造成結匯傳票轉至外帳可能會比內帳少幾元,若改為更新AX307,會造成國外收款傳票錯誤,故若借方為現金才改新AX306否則更新AX307
         'adoTaie.Execute "update acc031 set ax306 = ax306 + " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax306 <> 0 and rownum < 2"
         If adoacc020.Fields(2).Value <> 0 Then
            strSql = "update acc031 set ax306 = ax306 + " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax306 <> 0 AND AX305='1101' and rownum < 2"
            adoTaie.Execute strSql, lngEff
            If lngEff = 0 Then
               '2007/5/11 MODIFY BY SONIA 改更新貸方非規費科目的最小項次
               'StrSQLa = "update acc031 set ax307 = ax307 - " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax307 <> 0 and rownum < 2"
               StrSQLa = "update acc031 set ax307 = ax307 - " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 IN (" & _
                         "SELECT MIN(AX303) FROM ACC031 where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax307 <> 0 AND SUBSTR(AX305,1,3)<>'220' )"
               cnnConnection.Execute StrSQLa, lngEff
            End If
         End If
         '2005/11/24 END
      End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   adoTaie.Execute "insert into acc1z0 select '0', a1r02, a1r03, a1r04, a1r04 from acc1r0 where a1r01 = 'D' and a1r02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1r03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1r03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   '2014/2/18 modify by sonia 更新非J公司所有資料
   'modify by sonia 2020/7/22 A0C04<>'J'改A0C04='" & IIf(strCmp = "1", "P", strCmp) & "'
   'modify by sonia 2023/5/23 A0C04直接改7公司
   'adoTaie.Execute "UPDATE ACC0C0 SET A0C05= " & Val(FCDate(MaskEdBox2.Text)) & " WHERE A0C04='" & IIf(strCmp = "1", "P", strCmp) & "'"     '2005/11/22 ADD BY SONIA
   adoTaie.Execute "UPDATE ACC0C0 SET A0C05= " & Val(FCDate(MaskEdBox2.Text)) & " WHERE A0C04='7'"     '2005/11/22 ADD BY SONIA

   Screen.MousePointer = vbDefault
   Text1 = "" ': Text5 = "": Text6 = ""
   cboComp.ListIndex = 0 'Modify By Sindy 2020/4/17
   MsgBox MsgText(23), , MsgText(21)
Checking:
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   If adoacc021.State = adStateOpen Then
      Text1 = "錯誤之傳票號碼: " & stra0201 & " 公司 " & stra0202 & " 項次: " & adoacc021.Fields("aX203").Value
      adoacc021.Close
   End If
   Screen.MousePointer = vbDefault
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'add by sonia 2020/5/12
'*************************************************
'  將內帳L公司傳票資料轉外帳K公司   2023/5/19先將內帳L公司轉帳務K公司，後面再轉帳務L公司
'                                   2023/05以前是直接將內帳L公司轉帳務L公司
'*************************************************
Private Sub TransferLcomp()
Dim strNo As String
Dim strDocNo As String
Dim lngEff As Long
Dim StrSQLa As String
Dim strFromCmp As String '內帳公司別
Dim strSaleNM As String  '2023/5/19 add by sonia
 
On Error GoTo Checking
   strSql = ""
   strNo = ""
   strAccNo = ""
   Text1 = ""
   Screen.MousePointer = vbHourglass
   
   strFromCmp = "L": strCmp = "K"
   
   adoTaie.Execute "delete from acc031 where ax301 = '" & strCmp & "' and ax302 in (select a0302 from acc030 where a0301 = '" & strCmp & "' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "')"
   adoTaie.Execute "delete from acc030 where a0301 = '" & strCmp & "' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "'"
   adoTaie.Execute "delete from acc1z0 where a1z01 = '" & strCmp & "' and a1z02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1z03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1z03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   
   Text1 = "正在將內帳L公司傳票轉外帳K公司......"
   ProgressBar1.Value = 0
   adoacc020.CursorLocation = adUseClient
   
   adoacc020.Open "select * from acc020 where a0201='" & strFromCmp & "' and a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   '測試某傳票時用
   'adoacc020.Open "select * from acc020 where a0201='" & strFromCmp & "' and A0202 IN ('D113010051') order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount <> 0 Then
      ProgressBar1.max = adoacc020.RecordCount
   End If
   Do While adoacc020.EOF = False
      DoEvents
      
      '剔除:1.結餘結算傳票 2.每月25號的固定傳票 3.應收票據兌現轉乙存(借1130XX貸1103XX) 4.應付票據兌現轉甲存(借2111貸110203)
      If adoquery.State = adStateOpen Then adoquery.Close
      adoquery.CursorLocation = adUseClient
      'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日
      'adoquery.Open "select distinct a1p02,a1p22,a1p18 from acc1p0 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select distinct a1p02,a1p22,axd03 from acc1p0,acc0d1 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "' And A1p01=Axd01(+) And decode(a1p02,'U',Substr(A1p04,1,Length(A1p04)-7),null)=Axd02(+)", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         '1.結餘結算傳票
         If adoquery.Fields("a1p02").Value = "S" Then GoTo NextRecord
         '2.每月25號的固定傳票
         'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日
         'If adoquery.Fields("a1p02").Value = "U" And Right(adoquery.Fields("a1p18").Value, 2) = "25" Then GoTo NextRecord
         If adoquery.Fields("a1p02").Value = "U" And "" & adoquery.Fields("axd03").Value = "25" Then GoTo NextRecord
         '票據分錄
         If adoquery.Fields("a1p02").Value = "L" Then
            '3.應收票據兌現轉乙存(借1102XX貸1130XX)
            adoacc021.CursorLocation = adUseClient
            adoacc021.Open "select distinct a1.a1p22 NO,a1.a1p05 accno1,a2.a1p05 accno2 from acc1p0 a1,acc1p0 a2 where a1.a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1.a1p22 = '" & adoacc020.Fields("a0202").Value & "' and substr(a1.a1p05,1,4)='1102' and a1.a1p07>0 and a1.a1p01=a2.a1p01(+) and a1.a1p04=a2.a1p04(+) and a2.a1p08>0 ", adoTaie, adOpenStatic, adLockReadOnly
            If adoacc021.RecordCount = 1 Then
               If "" & Left(adoacc021.Fields("accno2").Value, 4) = "1130" Then
                  adoacc021.Close
                  GoTo NextRecord
               End If
            End If
            adoacc021.Close
            '4.應付票據兌現轉甲存(借2111貸110203)
            adoacc021.CursorLocation = adUseClient
            adoacc021.Open "select distinct a1.a1p22 NO,a1.a1p05 accno1,a2.a1p05 accno2 from acc1p0 a1,acc1p0 a2 where a1.a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1.a1p22 = '" & adoacc020.Fields("a0202").Value & "' and a1.a1p05='2111' and a1.a1p07>0 and a1.a1p01=a2.a1p01(+) and a1.a1p04=a2.a1p04(+) and a2.a1p08>0 ", adoTaie, adOpenStatic, adLockReadOnly
            If adoacc021.RecordCount = 1 Then
               If "" & adoacc021.Fields("accno2").Value = "110203" Then
                  adoacc021.Close
                  GoTo NextRecord
               End If
            End If
            adoacc021.Close
         End If
      End If
      adoquery.Close
      
      Acc030Save
      adoTaie.Execute "insert into acc030 (a0301, a0302, a0305, a0306, a0307, a0308) values ('" & strCmp & "', " & CNULL(stra0202) & ", " & lnga0205 & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
      adoacc021.CursorLocation = adUseClient
      adoacc021.Open "select * from acc021, acc1p0 where ax201=a1p01(+) and ax202 = a1p22 (+) and ax203 = a1p03 (+) and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc021.EOF = False
         'Debug.Print stra0202 & "  " & adoacc021.Fields("aX203").Value
         m_A1P26 = ""
         If IsNull(adoacc021.Fields("a1p26").Value) = False And adoacc021.Fields("a1p26").Value = MsgText(602) Then
            m_A1P26 = adoacc021.Fields("a1p26").Value
         End If
         If IsNull(adoacc021.Fields("a1p14").Value) = False And InStr(1, adoacc021.Fields("a1p14").Value, "代理人請款", 1) > 0 Then
            m_A1P26 = "Y"
         End If
         'add by sonia 2023/5/22   律師出庭費15000或5000都要併入收入
         If adoacc021.Fields("a1p05").Value = "220113" And (adoacc021.Fields("a1p08").Value = 5000 Or adoacc021.Fields("a1p08").Value = 15000) Then
            m_A1P26 = "Y"
         End If
         'end 2023/5/22
         Acc031Save
         If strNo <> strax202 Then
            strNo = strax202
         End If
         
         'add by sonia 2023/5/19 取消摘要欄之'智權人員名縮寫'||'/',因智權人員不在智權公司編制內,但因摘要位置很難抓,故只抓第1,2碼及第13,14碼才做
         '不能以AX309來判斷,有可能沒有值但摘要有
         '第1,2碼
         If Mid(strax212, 2, 1) <> "/" Then GoTo stepL1314
         strSaleNM = Mid(strax212, 1, 1)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select * from salesno order by sn02", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoquery.EOF = False
            If adoquery.Fields("sn01").Value = strSaleNM Then
               strax212 = Mid(strax212, 3)
               Exit Do
            End If
            adoquery.MoveNext
         Loop
         adoquery.Close
         GoTo stepLNext
stepL1314:
         '第13,14碼  AX312=SUBSTR(AX312,1,12)||SUBSTR(AX312,15)
         If Mid(strax212, 14, 1) <> "/" Then GoTo stepLNext
         strSaleNM = Mid(strax212, 13, 1)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select * from salesno order by sn02", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoquery.EOF = False
            If adoquery.Fields("sn01").Value = strSaleNM Then
               strax212 = Mid(strax212, 1, 12) & Mid(strax212, 15)
               Exit Do
            End If
            adoquery.MoveNext
         Loop
         adoquery.Close
stepLNext:
         'end 2023/5/19
         
         strax215 = ""       '作帳公司僅收入科目才放  add by sonia 2023/5/23
         strax205 = Mid(strax205, 1, 4)   '所有科目改為4碼
         '固定改科目
         Select Case strax205
            Case "6102"
               strax205 = "6101"
            Case "6120"
               strax205 = "6119"
            Case "6124"
               strax205 = "6111"
            Case "1102"
               '結匯傳票科目銀行存款1102XX改現金1101
               If "" & adoacc021.Fields("a1p02").Value = "I" Then
                  strax205 = "1101"
               End If
            'add by sonia 2023/5/19
            Case "6119"
               If InStr(strax212, "稿費") > 0 Then     '摘要有"稿費"字樣
                  strax205 = "6135"
               Else
                  strax205 = "6111"
               End If
            Case "1105"
               strax205 = "1101"
            'add by sonia 2025/2/26 1106,1911,1912,1913科目且為貸方時也改1101現金
            Case "1106", "1911", "1912", "1913"
               If douax207 > 0 Then strax205 = "1101"
            'end 2025/2/26
            Case "2112"
               If douax207 > 0 Then strax205 = "1101"  '廠商付款貸方2112改成現金1101
            Case "2407"
               strax205 = "4141"
               strax215 = "L"                          '作帳公司僅收入科目才放  add by sonia 2023/5/23
            'modify by sonia 2024/2/20 +4161科目
            Case "4141", "4161"
               strax215 = "L"                          '作帳公司僅收入科目才放  add by sonia 2023/5/23
            'end 2023/5/19
         End Select
         
         '外對內之收入科目改為4102或4112, 無申請人抓代理人 D095010545
         If ((strax205 >= "4" And strax205 < "5") Or strax205 = "2201") And douax207 > 0 And strCaseNo <> "ZZZZZZZZZZZZ" Then
            '申請人國籍非台灣者,收入科目改4102或4112
            If strCU10 > "010" Then
               '智權人員為國外部才要改科目,規費科目因無對沖智權人員所以在最後更新
               If strax209 <> "" Then
                  adoquery.CursorLocation = adUseClient
                  adoquery.Open "select ST15 from STAFF where ST01 = '" & strax209 & "' ", adoTaie, adOpenStatic, adLockReadOnly
                  If Mid(adoquery.Fields("ST15").Value, 1, 1) = "F" Then
                     Select Case strax205
                        Case "4101"
                           strax205 = "4102"
                        Case "4111"
                           strax205 = "4112"
                     End Select
                  End If
                  adoquery.Close
               End If
            End If
         End If
         
         'modify by sonia 2023/5/23 +ax315
         strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                  "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                  ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
         '考慮規費是否合併至收入
         'modify by sonia 2022/9/20 再加L公司的2403科目
         'If strax205 = "2201" And douax207 > 0 Then
         If (strax205 = "2201" Or strax205 = "2403") And douax207 > 0 Then
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
            adoquery.CursorLocation = adUseClient
            '相同案號才可合併  '收入必須為貸方才抓
            'modify by sonia 2023/5/23 +AX315
            adoquery.Open "select ax305,AX315,AX303 from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                  If m_A1P26 = MsgText(602) Then
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  '不合併時規費作帳公司同收入
                  Else
                     strax215 = "" & adoquery.Fields("ax315").Value   'add by sonia 2023/5/23
                     'modify by sonia 2023/5/23 +ax315
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
               Else
                  If strax205W = "4141" Then strax215 = "L"           'add by sonia 2023/5/23
                  If strax205W <> "" And m_A1P26 = MsgText(602) Then
                     'modify by sonia 2023/5/23 +ax315
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
                  Else
                     'modify by sonia 2023/5/23 +ax315
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
               End If
            Else
               If strax205W = "4141" Then strax215 = "L"           'add by sonia 2023/5/23
               If strax205W <> "" And m_A1P26 = MsgText(602) Then
                  'modify by sonia 2023/5/23 +ax315
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
               Else
                  'modify by sonia 2023/5/23 +ax315
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
               End If
            End If
            adoquery.Close
         End If
         
         '4101,4111,4121,4131其他對沖為結餘X者,科目改0004,作帳公司NULL
         '收入科目,借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
         Select Case strax205
            Case "4101", "4111", "4102", "4112", "4121", "4131"
               If IsNull(strax213) = False And Mid(strax213, 1, 2) = "結餘" Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
               End If
               '借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
               If douax206 + douax207 = 5000 And (InStr(strax212, "點作轉專業") > 0 Or InStr(strax212, "支援") > 0) Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('" & strCmp & "', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
              End If
            Case Else
         End Select
         '非結餘的收款傳票414101法務收入科目,智權人員為M0100總所者此為複委託,應併入前一收入項次D100010157,要抓acc1p0才能確定為414101
         If strax205 >= "4" And strax205 < "5" And IsNull(strax213) = False And Mid(strax213, 1, 2) <> "結餘" Then
            If "" & adoacc021.Fields("a1p02").Value & adoacc021.Fields("a1p05").Value & adoacc021.Fields("a1p16").Value = "A414101M0100" Then
               If adoquery.State = adStateOpen Then
                  adoquery.Close
               End If
               adoquery.CursorLocation = adUseClient
               adoquery.Open "select ax305,AX303 from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
               If adoquery.RecordCount <> 0 Then
                  If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = '" & strCmp & "' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  End If
               End If
               adoquery.Close
            End If
         End If
         
         If strDocNo <> (adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value) Then
            adoTaie.Execute strSql
            strDocNo = adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value
         End If
         adoacc021.MoveNext
      Loop
      adoacc021.Close

NextRecord:
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   '內帳借方有110204,110205,110222,113002科目之傳票,該傳票的貸方之收入及規費改為4102,4112,2202
   Text1 = "正在將內帳借方有110204,110205,110222" & vbCrLf & "    ,113002科目之傳票,該傳票的貸方之收入" & vbCrLf & "    及規費改為4102,4112,2202......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) WHERE AX301='" & strCmp & "' AND AX307>0 AND AX302 IN ( " & _
            "SELECT A0202 FROM acc020,ACC021 WHERE A0201=AX201 AND A0202=AX202 AND A0201='" & strFromCmp & "' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX205 IN ('110204','110205','110222','113002') AND AX206>0 GROUP BY A0202)"
   adoTaie.Execute strSql
   
   '貸方有4102,4112科目之傳票,該傳票的貸方之收入及規費都改為4102,4112,2202
   Text1 = "正在將貸方有4102,4112科目之傳票,該傳票的" & vbCrLf & "    貸方之收入及規費都改為" & vbCrLf & "    4102,4112,2202......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) where AX301='" & strCmp & "' and ((AX305>='4' AND AX305<'5') OR AX305='2201') and ax307>0 AND ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='" & strCmp & "' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('4102','4112') AND AX307>0 GROUP BY A0302)"
   adoTaie.Execute strSql
   
   Text1 = "正在調整因小數位捨去所造成的借貸方" & vbCrLf & "    差額......"
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select ax301, ax302, sum(ax307 - ax306) from acc031, acc030 where A0301='" & strCmp & "' AND ax301 = a0301 and ax302 = a0302 and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by ax301, ax302", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      DoEvents
      If IsNull(adoacc020.Fields(2).Value) = False Then
         '因小數位捨去會造成結匯傳票轉至外帳可能會比內帳少幾元,若改為更新AX307,會造成國外收款傳票錯誤,故若借方為現金才改新AX306否則更新AX307
         If adoacc020.Fields(2).Value <> 0 Then
            strSql = "update acc031 set ax306 = ax306 + " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax306 <> 0 AND AX305='1101' and rownum < 2"
            adoTaie.Execute strSql, lngEff
            If lngEff = 0 Then
               '改更新貸方非規費科目的最小項次
               StrSQLa = "update acc031 set ax307 = ax307 - " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 IN (" & _
                         "SELECT MIN(AX303) FROM ACC031 where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax307 <> 0 AND SUBSTR(AX305,1,3)<>'220' )"
               cnnConnection.Execute StrSQLa, lngEff
            End If
         End If
      End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   adoTaie.Execute "insert into acc1z0 select '" & strCmp & "', a1r02, a1r03, a1r04, a1r04 from acc1r0 where a1r01 = '" & strFromCmp & "D' and a1r02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1r03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1r03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   adoTaie.Execute "UPDATE ACC0C0 SET A0C05= " & Val(FCDate(MaskEdBox2.Text)) & " WHERE A0C04='L'"

   Screen.MousePointer = vbDefault
   Text1 = ""
   cboComp.ListIndex = 0
   MsgBox MsgText(23), , MsgText(21)
   
Checking:
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   If adoacc021.State = adStateOpen Then
      Text1 = "錯誤之傳票號碼: " & stra0201 & " 公司 " & stra0202 & " 項次: " & adoacc021.Fields("aX203").Value
      adoacc021.Close
   End If
   Screen.MousePointer = vbDefault
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'end 2020/5/12

'2006/10/16 ADD BY SONIA P非台灣, CFP日本及EPC之會計科目為4101
Private Function GetACCNO(ByVal Nation As String, ByVal CaseNo As String, ByVal AccNo As String) As String
   GetACCNO = AccNo
   Select Case Mid(CaseNo, 1, Len(strCaseNo) - 9)
      Case "CFP"
         Select Case Nation
            Case "011"
               '2011/2/22 modify by sonia
               'If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" Then
               If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" And Mid(CaseNo, Len(CaseNo) - 8, 6) <= "023914" Then
                  GetACCNO = "4101"
               End If
            Case "221"
               If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" And Mid(CaseNo, Len(CaseNo) - 8, 6) < "016183" Then
                  GetACCNO = "4101"
               End If
         End Select
      Case "P"
         If Nation <> "000" Then
            GetACCNO = "4101"
            '2012/4/24 ADD BY SONIA 此日期起非台灣P案改用專利法律2公司
            If Mid(CaseNo, Len(CaseNo) - 8, 6) >= 101672 Then
               GetACCNO = "4111"
            End If
            '2012/4/24 END
         End If
   End Select
End Function
'2006/10/16 ADD BY SONIA P非台灣, CFP日本及EPC之作帳公司為6
Private Function GetCOMPNO(ByVal Nation As String, ByVal CaseNo As String) As String
   GetCOMPNO = "2"   '2007/6/1 ADD BY SONIA 預設7公司 D096031771
   Select Case Mid(CaseNo, 1, Len(CaseNo) - 9)
      Case "CFP", "CPS"
         Select Case Nation
            Case "011"
               '2011/2/22 modify by sonia
               'If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" Then
               If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" And Mid(CaseNo, Len(CaseNo) - 8, 6) <= "023914" Then
                  GetCOMPNO = "1"
               End If
            Case "221"
               If Mid(CaseNo, Len(CaseNo) - 8, 6) >= "011051" And Mid(CaseNo, Len(CaseNo) - 8, 6) < "016183" Then
                  GetCOMPNO = "1"
               End If
         End Select
      Case "P", "PS"
         If Nation <> "000" Then
            GetCOMPNO = "1"
            '2012/4/24 ADD BY SONIA 此日期起非台灣P案改用專利法律2公司
            If Mid(CaseNo, Len(CaseNo) - 8, 6) >= 101672 Then
               GetCOMPNO = "2"
            End If
            '2012/4/24 END
         End If
   End Select
End Function

'Modify by Sindy 2020/4/17 公司別改下拉
''2014/2/17 add by sonia
'Private Sub Text5_Change()
'   If Text5 = MsgText(601) Then
'      Exit Sub
'   End If
'   Text6 = A0802Query(Text5)
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub
'
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub Text5_Validate(Cancel As Boolean)
'   If Text5 = MsgText(601) Then
'      MsgBox MsgText(10) & Label3, , MsgText(5)
'      Cancel = True
'      Text5.SetFocus
'      Exit Sub
'   Else
'      If Text5 <> "1" And Text5 <> "J" Then
'         MsgBox "只可輸入 1 或 J", vbCritical
'         Cancel = True
'         Text5.SetFocus
'         Exit Sub
'      End If
'   End If
'End Sub

'*************************************************
'  將內帳J公司傳票資料直接轉外帳J公司
'
'*************************************************
Private Sub TransferJcomp()
Dim strNo As String
Dim strDocNo As String
Dim lngEff As Long
Dim StrSQLa As String
Dim strSaleNM As String  '2015/3/13 add by sonia

On Error GoTo CheckingJ
   strSql = ""
   strNo = ""
   strAccNo = ""
   Text1 = ""
   Screen.MousePointer = vbHourglass
   adoTaie.Execute "delete from acc031 where ax301 = 'J' and ax302 in (select a0302 from acc030 where a0301 = 'J' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "')"
   adoTaie.Execute "delete from acc030 where a0301 = 'J' and a0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "'"
   adoTaie.Execute "delete from acc1z0 where a1z01 = 'J' and a1z02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1z03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1z03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   
   Text1 = "正在將內帳J公司傳票轉至外帳J公司......"
   ProgressBar1.Value = 0
   adoacc020.CursorLocation = adUseClient
   
   adoacc020.Open "select * from acc020 where a0201='J' and a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   '測試某傳票時用
   'adoacc020.Open "select * from acc020 where a0201='J' and A0202 IN ('D108080101') order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount <> 0 Then
      ProgressBar1.max = adoacc020.RecordCount
   End If
   Do While adoacc020.EOF = False
      DoEvents
      
      'add by sonia 2019/10/15 剔除結餘結算傳票及每月25號的固定傳票
      If adoquery.State = adStateOpen Then adoquery.Close
      adoquery.CursorLocation = adUseClient
      'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日
      'adoquery.Open "select distinct a1p02,a1p22,a1p18 from acc1p0 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select distinct a1p02,a1p22,axd03 from acc1p0,acc0d1 where a1p01 = '" & adoacc020.Fields("a0201").Value & "' and a1p22 = '" & adoacc020.Fields("a0202").Value & "' And A1p01=Axd01(+) And decode(a1p02,'U',Substr(A1p04,1,Length(A1p04)-7),null)=Axd02(+)", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         '結餘結算傳票
         If adoquery.Fields("a1p02").Value = "S" Then GoTo NextJRecord
         '每月25號的固定傳票
         'modify by sonia 2021/9/1 應判斷AXD03為25號而不是判斷A1P18,否則25號假日傳票會產生在下一工作日
         'If adoquery.Fields("a1p02").Value = "U" And Right(adoquery.Fields("a1p18").Value, 2) = "25" Then GoTo NextJRecord
         If adoquery.Fields("a1p02").Value = "U" And "" & adoquery.Fields("axd03").Value = "25" Then GoTo NextJRecord
      End If
      adoquery.Close
      'end 2019/10/15
      
      Acc030Save
      adoTaie.Execute "insert into acc030 (a0301, a0302, a0305, a0306, a0307, a0308) values ('J', " & CNULL(stra0202) & ", " & lnga0205 & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
      adoacc021.CursorLocation = adUseClient
      adoacc021.Open "select * from acc021, acc1p0 where ax201=a1p01(+) and ax202 = a1p22 (+) and ax203 = a1p03 (+) and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc021.EOF = False
         'Debug.Print stra0202 & "  " & adoacc021.Fields("aX203").Value
         strCom = ""
         m_A1P26 = ""
         If IsNull(adoacc021.Fields("a1p26").Value) = False And adoacc021.Fields("a1p26").Value = MsgText(602) Then
            m_A1P26 = adoacc021.Fields("a1p26").Value
         End If
         If IsNull(adoacc021.Fields("a1p14").Value) = False And InStr(1, adoacc021.Fields("a1p14").Value, "代理人請款", 1) > 0 Then
            m_A1P26 = "Y"
         End If
         Acc031Save
'2014/2/18 AX215已不使用
'         If strNo <> strax202 Then
'            Select Case Mid(strax205, 1, 1)
'               Case "6"
'                  If IsNull(adoacc021.Fields("ax215").Value) = False Then
'                     strCom = adoacc021.Fields("ax215").Value     '費用科目作帳公司抓 AX215
'                  End If
'            End Select
'            strNo = strax202
'         End If
         strCom = "J"   '仍為J公司
         strax215 = ""  '無作帳公司
         
         '2015/3/9 ADD BY SONIA 取消摘要欄之'智權人員名縮寫'||'/',因智權人員不在智權公司編制內,但因摘要位置很難抓,故只抓第1,2碼及第13,14碼才做
         '不能以AX309來判斷,有可能沒有值但摘要有
         '第1,2碼
         If Mid(strax212, 2, 1) <> "/" Then GoTo step1314
         strSaleNM = Mid(strax212, 1, 1)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select * from salesno order by sn02", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoquery.EOF = False
            If adoquery.Fields("sn01").Value = strSaleNM Then
               strax212 = Mid(strax212, 3)
               Exit Do
            End If
            adoquery.MoveNext
         Loop
         adoquery.Close
         GoTo stepNext   'ADD BY SONIA 2016/8/18
step1314:
         '第13,14碼  AX312=SUBSTR(AX312,1,12)||SUBSTR(AX312,15)
         If Mid(strax212, 14, 1) <> "/" Then GoTo stepNext
         strSaleNM = Mid(strax212, 13, 1)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select * from salesno order by sn02", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoquery.EOF = False
            If adoquery.Fields("sn01").Value = strSaleNM Then
               strax212 = Mid(strax212, 1, 12) & Mid(strax212, 15)
               Exit Do
            End If
            adoquery.MoveNext
         Loop
         adoquery.Close
stepNext:
         '2015/3/9 END
         
         'add by sonia 2023/4/26 J公司收入換科目
         Select Case strax205
            Case "420101", "410102", "411102"  '創新及顧問
               strax205 = "4201"
'cancel by sonia 2023/5/9 下方批次會改
'            Case "410103"                      'CMT
'               strax205 = "4203"
'            Case "412101"                      'CFT
'               strax205 = "4204"
'            Case "411103"                      'CMP
'               strax205 = "4205"
'            Case "413101"                      'CFP
'               strax205 = "4206"
'end 2023/5/9
         End Select
         '貸方2141依案號改為收入科目
         If strax205 = "2141" And douax207 > 0 Then
            Select Case Mid(strCaseNo, 1, Len(strCaseNo) - 9)
               Case "P", "PS"                  'CMP
                  strax205 = "4205"
               Case "CFP", "CPS"               'CFP
                  strax205 = "4206"
               Case "CFT", "CFC", "S"          'CFT
                  strax205 = "4204"
               Case Else
                  Select Case Mid(strCaseNo, 1, 1)
                     Case "T"                  'CMT
                        strax205 = "4205"
                     Case Else                 '創新及顧問
                        strax205 = "4201"
                  End Select
            End Select
         End If
         'end 2023/4/26
         
         '所有科目改為4碼,銀行存款科目1103(全部) 保留全碼數（六碼）
         'modify by sonia 2015/3/25 辜說1130也要保留六碼
         If Mid(strax205, 1, 4) = "1103" Or Mid(strax205, 1, 4) = "1130" Then
         'add by sonia 2023/4/25 1141科目改為1133,220114科目改為2403
         ElseIf strax205 = "1141" Then
            strax205 = "1133"
         ElseIf strax205 = "220114" Then
            strax205 = "2403"
         'end 2023/4/25
         Else
            strax205 = Mid(strax205, 1, 4)
         End If
         Select Case strax205
            Case "4141"
               strax205 = "4111"
            Case "4161"
               strax205 = "4112"
            Case "6102"
               strax205 = "6101"
            Case "6120"
               strax205 = "6119"
            Case "6124"
               strax205 = "6111"
            'modify by sonia 2014/10/27 J公司6136固定轉4202, 10/29再改轉4301
            Case "6136"
               strax205 = "4301"
            'end 2014/10/27
         End Select
         '外對內之收入科目改為4102或4112, 無申請人抓代理人 D095010545
         If ((strax205 >= "4" And strax205 < "5") Or strax205 = "2201") And douax207 > 0 And strCaseNo <> "ZZZZZZZZZZZZ" Then
            '申請人國籍非台灣者,收入科目改4102或4112
            If strCU10 > "010" Then
               If strax209 <> "" Then
                  adoquery.CursorLocation = adUseClient
                  adoquery.Open "select ST15 from STAFF where ST01 = '" & strax209 & "' ", adoTaie, adOpenStatic, adLockReadOnly
                  If Mid(adoquery.Fields("ST15").Value, 1, 1) = "F" Then
                     Select Case strax205
                        Case "4101"
                           strax205 = "4102"
                        Case "4111"
                           strax205 = "4112"
                     End Select
                  End If
                  adoquery.Close
               End If
            End If
         End If
         strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                  "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                  ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
         '考慮規費是否合併至收入
         If strax205 = "2201" And douax207 > 0 Then
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select ax305,AX315,AX303 from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                  If m_A1P26 = MsgText(602) Then
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  '不合併時規費作帳公司同收入
                  Else
                     strax215 = "" & adoquery.Fields("ax315").Value
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
               Else
                  If strax205W <> "" And m_A1P26 = MsgText(602) Then
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
                  Else
                     strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, ax316) " & _
                              "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                              ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
                  End If
               End If
            Else
               If strax205W <> "" And m_A1P26 = MsgText(602) Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205W) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205W) & ")"
               Else
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, ax315, AX316) " & _
                           "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", " & CNULL(strax205) & ", " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", " & CNULL(strax215) & ", " & CNULL(strax205) & ")"
               End If
            End If
            adoquery.Close
         End If
         '4101,4111,4121,4131其他對沖為結餘X者,科目改0004,作帳公司NULL
         '收入科目,借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
         Select Case strax205
            Case "4101", "4111", "4102", "4112", "4121", "4131"    '2006/7/12加4102,4112
               If IsNull(strax213) = False And Mid(strax213, 1, 2) = "結餘" Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
               End If
               '借或貸金額5000,摘要有'點作轉專業'或'支援'字樣者,科目改0004,作帳公司NULL
               If douax206 + douax207 = 5000 And (InStr(strax212, "點作轉專業") > 0 Or InStr(strax212, "支援") > 0) Then
                  strSql = "insert into acc031 (ax301, ax302, ax303, ax304, ax305, ax306, ax307, ax308, ax309, ax310, ax312, ax313, ax314, AX316) " & _
                           "values ('J', " & CNULL(strax202) & ", " & CNULL(strax203) & ", " & CNULL(strax204) & ", '0004', " & douax206 & ", " & douax207 & ", " & CNULL(ChgSQL(strax208)) & "" & _
                           ", " & CNULL(strax209) & ", null, " & CNULL(ChgSQL(strax212)) & ", " & CNULL(ChgSQL(strax213)) & ", " & CNULL(strax214) & ", '0004' " & ")"
              End If
            Case Else
         End Select
         '非結餘的收款傳票414101法務收入科目,智權人員為M0100總所者此為複委託,應併入前一收入項次D100010157,要抓acc1p0才能確定為414101
         If strax205 >= "4" And strax205 < "5" And IsNull(strax213) = False And Mid(strax213, 1, 2) <> "結餘" Then
            If "" & adoacc021.Fields("a1p02").Value & adoacc021.Fields("a1p05").Value & adoacc021.Fields("a1p16").Value = "A414101M0100" Then
               If adoquery.State = adStateOpen Then
                  adoquery.Close
               End If
               adoquery.CursorLocation = adUseClient
               adoquery.Open "select ax305,AX315,AX303 from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax314 = '" & strax214 & "' AND AX307>0 )", adoTaie, adOpenStatic, adLockReadOnly
               If adoquery.RecordCount <> 0 Then
                  If Mid(adoquery.Fields("ax305").Value, 1, 1) = "4" Then
                     strSql = "update acc031 set ax306 = ax306 + " & douax206 & ", ax307 = ax307 + " & douax207 & " where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 in (select max(ax303) from acc031 where ax301 = 'J' and ax302 = '" & strax202 & "' and ax303 = '" & adoquery.Fields("ax303").Value & "')"
                  End If
               End If
               adoquery.Close
            End If
         End If
         
         If strDocNo <> (adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value) Then
            adoTaie.Execute strSql
            strDocNo = adoacc021.Fields("ax201").Value & adoacc021.Fields("ax202").Value & adoacc021.Fields("ax203").Value
         End If
         adoacc021.MoveNext
      Loop
      adoacc021.Close

NextJRecord:   'add by sonia 2019/10/15
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
'cancel by sonia 2018/6/14 辜通知取消
'   Text1 = "正在處理借方有110205或110222科目" & vbCrLf & "    且金額>120之傳票,該傳票的借方增加" & vbCrLf & "    6113雜費且金額120元時,則第一個" & vbCrLf & "    110205或110222減少120元......"
'   ProgressBar1.Value = 0
'   adoacc020.CursorLocation = adUseClient
'   adoacc020.Open "select DISTINCT AX201,AX202 from ACC021 where AX205 IN ('110205','110222') AND AX206>120 AND (AX201,AX202) IN ( " & _
'                  "select A0201,A0202 from acc020,ACC021 where A0201=AX201 AND A0202=AX202 AND AX205>='4' AND AX205<='5' AND AX207>0 AND A0201='J' AND a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "') GROUP BY AX201,AX202 order by AX201,AX202 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '測試某傳票時用
'   'adoacc020.Open "select DISTINCT AX201,AX202 from ACC021 where AX205 IN ('110205','110222') AND AX206>150 AND (AX201,AX202) IN ( " & _
'   '               "select A0201,A0202 from acc020,ACC021 where A0201=AX201 AND A0202 IN ('D095022377','D095010812','D095010834','D095010837','D095011547','D095011762','D095011957','D095012404','D095012414') AND A0202=AX202 AND AX205>='4' AND AX205<='5' AND AX207>0 AND A0201='J' AND a0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "') GROUP BY AX201,AX202 order by AX201,AX202 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc020.RecordCount <> 0 Then
'      ProgressBar1.max = adoacc020.RecordCount
'   End If
'   Do While adoacc020.EOF = False
'      DoEvents
'      'Debug.Print adoacc020.Fields("AX202")
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select MIN(AX203) from acc021 where ax201 = " & CNULL(adoacc020.Fields("AX201")) & " and ax202 = " & CNULL(adoacc020.Fields("AX202")) & " and ax205 IN ('110205','110222') and ax206>120 ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strSql = "UPDATE ACC031 SET AX306=AX306-120 WHERE AX301='J' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & " AND AX303=" & CNULL(adoquery.Fields(0)) & ""
'         adoTaie.Execute strSql
'      End If
'      adoquery.Close
'
'      '為要新增001項次,所以先將每一項次+1
'      strSql = "UPDATE ACC031 SET AX303=SUBSTR((AX303+1001),2,3) WHERE AX301='J' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & ""
'      adoTaie.Execute strSql
'
'      '再新增001項次6113雜項借方金額150,並決定作帳公司別
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select * from acc031 where AX301='J' AND ax302 = " & CNULL(adoacc020.Fields("AX202")) & " and ax305>='4' AND AX305<'5' AND AX307>0 AND AX315='7' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strax215 = "7"
'      Else
'         strax215 = "6"
'      End If
'      adoquery.Close
'      strSql = "insert into acc031 (ax301, ax302, ax303, AX304, ax305, AX316, ax306, ax307, AX312, ax315) " & _
'               "values ('J', " & CNULL(adoacc020.Fields("AX202")) & ", '001', 'TOT', '6113', '6113', 120, 0, '手續費', " & CNULL(strax215) & ")"
'      adoTaie.Execute strSql
'
'      ProgressBar1.Value = ProgressBar1.Value + 1
'      adoacc020.MoveNext
'   Loop
'   adoacc020.Close
'end  2018/6/14 辜通知取消
   
   'add by sonia 2023/5/5
   '有2141借方之傳票(跨月收款傳票)，先將該傳票與2141相同案號之貸方收入及規費刪除，再刪除借方之2141
   '但該案號若無貸方之收入或規費時則不刪除保留原樣D112040042
   Text1 = "正在處理有2141借方之傳票(跨月收款傳票)資料......"
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select a.* from " & _
                  "(select ax301,ax302,ax303,ax314 from acc030,acc031 where a0301='J' and a0305>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' " & _
                  "and a0301=ax301 and a0302=ax302 and ax305='2141' and ax306>0) a, " & _
                  "(select distinct ax301,ax302,ax314 from acc030,acc031 where a0301='J' and a0305>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' " & _
                  "and a0301=ax301 and a0302=ax302 and ax307>0 and (ax305 like '2201%' or ax305 like '4%')) b " & _
                  "where a.ax301=b.ax301(+) and a.ax302=b.ax302(+) and a.ax314=b.ax314(+) and b.ax314 is not null ", adoTaie, adOpenStatic, adLockReadOnly
Do While adoacc020.EOF = False
      DoEvents
         '貸方收入及規費刪除
         strSql = "delete ACC031 WHERE AX301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax314 = '" & adoacc020.Fields("ax314").Value & "' and ax307>0 and (ax305 like '2201%' or ax305 like '4%')"
         adoTaie.Execute strSql
         '刪除借方之2141
         strSql = "delete ACC031 WHERE AX301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 = '" & adoacc020.Fields("ax303").Value & "'"
         adoTaie.Execute strSql
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   Text1 = "正在處理收入傳票無貸方規費時，借方之2201,2211要刪除，差額減在貸方第一筆收入......"
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select ax301,ax302,ax303,ax314,ax306 from acc031,acc030,acc1p0 where A0301='J' AND ax301 = a0301 and ax302 = a0302 and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax305 in ('2201','2211') and ax306>0 " & _
                  "and ax301=a1p01(+) and ax302=a1p22(+) and ax305=substr(a1p05,1,4) and ax306=a1p07(+) and ax314=a1p17(+) and a1p02='A' ", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      DoEvents
            '差額減在貸方同案號最小項次收入
            strSql = "update acc031 set ax307 = ax307 - " & Val(adoacc020.Fields("ax306").Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 IN (" & _
                     "SELECT MIN(AX303) FROM ACC031 where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax314 = '" & adoacc020.Fields("ax314").Value & "' and ax307 > 0 AND AX305 like '4%' )"
            adoTaie.Execute strSql, lngEff
            If lngEff > 0 Then
               '刪除借方之2201,2211
               StrSQLa = "delete ACC031 WHERE AX301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 = '" & adoacc020.Fields("ax303").Value & "'"
               cnnConnection.Execute StrSQLa, lngEff
            End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   '部分收款之收入科目改為暫收款2401  D108080101
   Text1 = "正在將部分收款之收入科目改為暫收款2401......"
   strSql = "UPDATE ACC031 SET AX305='2401',AX316='2401' WHERE (AX301,AX302,AX303) IN ( " & _
            "SELECT DISTINCT AX301,AX302,AX303 FROM ACC030,ACC031,acc1p0,acc0m0,acc0k0,acc0j0 WHERE A0301='J' AND A0305>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301 AND A0302=AX302 and ax305 like '4%' " & _
            "and ax301=a1p01(+) and ax302=a1p22(+) and ax306=a1p07(+) and ax314=a1p17(+) and a1p02='A' and a1p05 like '4%'" & _
            "and a1p04=a0m01(+) and a0m02=a0k01(+) and a0m02=a0j13(+) and a1p17=a0j02(+) and a0j02 is not null and a0k37 is null)"
   adoTaie.Execute strSql
   'end 2023/5/5
   
   '自行輸入之借方6113且金額為150者,掛作帳公司,貸方有7公司者則放7,否則放6  D095010837
   Text1 = "正在將自行輸入之借方6113且金額為120者" & vbCrLf & "    ,掛作帳公司,貸方有7公司者則放7公司," & vbCrLf & "    否則放6公司......"
   strSql = "UPDATE ACC031 B SET B.AX315= " & _
            "(SELECT MAX(A.AX315) FROM ACC031 A WHERE B.AX301=A.AX301 AND B.AX302=A.AX302 AND A.AX307>0 AND A.AX315 IS NOT NULL) " & _
            "WHERE (B.AX301,B.AX302,B.AX303) IN ( " & _
            "SELECT AX301,AX302,AX303 FROM ACC030,ACC031 WHERE A0301='J' AND A0305>= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301 AND A0302=AX302 AND AX306=120 AND AX315 IS NULL AND AX305='6113')"
   adoTaie.Execute strSql
   
'cancel by sonia 2018/6/14 辜通知取消
'   '借方有110205或110222科目且金額>120,且貸方有2401且金額與借方相同之傳票,該傳票的借方增加6113雜費金額120元(作帳公司放7公司),第一個110205或110222減少120元
'   Text1 = "正在處理借方有110205或110222科目" & vbCrLf & "    且金額>120,且貸方有2401且金額與借" & vbCrLf & "    方相同之傳票,該傳票的借方增加" & vbCrLf & "    6113雜費且金額120元時,則第一個" & vbCrLf & "    110205或110222減少120元......"
'   ProgressBar1.Value = 0
'   adoacc020.CursorLocation = adUseClient
'   adoacc020.Open "SELECT DISTINCT A.AX201,A.AX202 from ACC020,ACC021 A,ACC021 B WHERE A0201='J' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=A.AX201(+) AND A0202=A.AX202(+) AND A.AX205 IN ('110205','110222') AND A.AX206>120 " & _
'                  "AND A.AX201=B.AX201(+) AND A.AX202=B.AX202(+) AND SUBSTR(B.AX205,1,4)='2401' AND A.AX206=B.AX207", adoTaie, adOpenStatic, adLockReadOnly
'   '測試某傳票時用
'   'adoacc020.Open "SELECT DISTINCT A.AX201,A.AX202 from ACC020,ACC021 A,ACC021 B WHERE A0201='J' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0201=A.AX201(+) AND A0202=A.AX202(+) " & _
'                   "AND A0202 IN ('D102090805','D102091120','D102091795','D102091872') AND A.AX205 IN ('110205','110222') AND A.AX206>120 " & _
'                   "AND A.AX201=B.AX201(+) AND A.AX202=B.AX202(+) AND SUBSTR(B.AX205,1,4)='2401' AND A.AX206=B.AX207", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc020.RecordCount <> 0 Then
'      ProgressBar1.max = adoacc020.RecordCount
'   End If
'   Do While adoacc020.EOF = False
'      DoEvents
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select MIN(AX203) from acc021 where ax201 = " & CNULL(adoacc020.Fields("AX201")) & " and ax202 = " & CNULL(adoacc020.Fields("AX202")) & " and ax205 IN ('110205','110222') and ax206>120 ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount > 0 Then
'         strSql = "UPDATE ACC031 SET AX306=AX306-120 WHERE AX301='J' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & " AND AX303=" & CNULL(adoquery.Fields(0)) & ""
'         adoTaie.Execute strSql
'      End If
'      adoquery.Close
'
'      '為要新增001項次,所以先將每一項次+1
'      strSql = "UPDATE ACC031 SET AX303=SUBSTR((AX303+1001),2,3) WHERE AX301='J' AND AX302=" & CNULL(adoacc020.Fields("AX202")) & ""
'      adoTaie.Execute strSql
'
'      '再新增001項次6113雜項借方金額120,作帳公司別放7
'      strSql = "insert into acc031 (ax301, ax302, ax303, AX304, ax305, AX316, ax306, ax307, AX312, ax315) " & _
'               "values ('J', " & CNULL(adoacc020.Fields("AX202")) & ", '001', 'TOT', '6113', '6113', 120, 0, '手續費', '7')"
'      adoTaie.Execute strSql
'
'      ProgressBar1.Value = ProgressBar1.Value + 1
'      adoacc020.MoveNext
'   Loop
'   adoacc020.Close
'end  2018/6/14 辜通知取消
   
   '內帳借方有110204,110205,110222,113002科目之傳票,該傳票的貸方之收入及規費改為4102,4112,2202      D095010812
   Text1 = "正在將內帳借方有110204,110205,110222" & vbCrLf & "    ,113002科目之傳票,該傳票的貸方之收入" & vbCrLf & "    及規費改為4102,4112,2202......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) WHERE AX301='J' AND AX307>0 AND AX302 IN ( " & _
            "SELECT A0202 FROM acc020,ACC021 WHERE A0201=AX201 AND A0202=AX202 AND A0201='J' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX205 IN ('110204','110205','110222','113002') AND AX206>0 GROUP BY A0202)"
   adoTaie.Execute strSql
   
   '內帳借方有113001科目且對沖客戶為V0001之傳票,該傳票的貸方之所有科目都改回0004科目且不放作帳公司      D095011547~48
   Text1 = "正在將內帳借方有113001科目且對沖客戶" & vbCrLf & "    為V0001之傳票,該傳票的貸方之所有科目" & vbCrLf & "    都改回0004科目且不放作帳公司......"
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='J' AND AX307>0 AND AX302 IN ( " & _
            "SELECT A0202 FROM acc020,ACC021 WHERE A0201=AX201 AND A0202=AX202 AND A0201='J' AND A0205 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0205 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX205='113001' AND AX206>0 AND AX208='V0001' GROUP BY A0202)"
   adoTaie.Execute strSql
   
   '傳票有2211或2491科目不管借方或貸方,對方之收入及規費都改回0004科目且不放作帳公司      D095021629~30
   Text1 = "正在將傳票有2211或2491科目不管借方或貸方" & vbCrLf & "    ,對方之收入及規費都改回0004科目" & vbCrLf & "    且不放作帳公司......"
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='J' AND AX305>='2201' AND AX305<='2202' AND AX302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='J' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('2211','2491') GROUP BY A0302)"
   adoTaie.Execute strSql
   strSql = "UPDATE ACC031 SET AX305='0004',AX315=NULL WHERE AX301='J' AND AX305>='4' AND AX305<'5' AND AX302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='J' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' AND A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('2211','2491') GROUP BY A0302)"
   adoTaie.Execute strSql
   
'cancel by sonia 2023/5/9 J公司不會有4102,4112
'   '貸方有4102,4112科目之傳票,該傳票的貸方之收入及規費都改為4102,4112,2202
'   Text1 = "正在將貸方有4102,4112科目之傳票,該傳票的" & vbCrLf & "    貸方之收入及規費都改為" & vbCrLf & "    4102,4112,2202......"
'   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2201','2202','4101','4102','4111','4112',AX305)) where AX301='J' and ((AX305>='4' AND AX305<'5') OR AX305='2201') and ax307>0 AND ax302 IN ( " & _
'            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='J' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('4102','4112') AND AX307>0 GROUP BY A0302)"
'   adoTaie.Execute strSql
'end 2023/5/9

   '借方有1915或1916科目之傳票,該傳票的貸方之收入及規費都改回原科目        D095012414
   Text1 = "正在將借方有1915或1916科目之傳票,該傳票" & vbCrLf & "    的貸方之收入及規費都改回原科目......"
   strSql = "UPDATE ACC031 SET AX305=(DECODE(AX305,'2202','2201','4102','4101','4112','4111',AX305)) where AX301='J' and ((AX305>='4' AND AX305<'5') OR AX305='2202') and ax307>0 AND ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='J' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND AX305 IN ('1915','1916') AND AX306>0 GROUP BY A0302)"
   adoTaie.Execute strSql
   
   '科目為6114機油費之傳票,該項次摘要改成 汽車加油
   Text1 = "正在將科目為6114機油費之傳票,該項次摘要" & vbCrLf & "    改成 汽車加油......"
   strSql = "UPDATE ACC031 SET AX312='汽車加油' where AX301='J' AND AX305='6114' and ax302 IN ( " & _
            "SELECT A0302 FROM acc030,ACC031 WHERE A0301='J' AND A0301=AX301 AND A0302=AX302 AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' GROUP BY A0302)"
   adoTaie.Execute strSql
   
   'add by sonia 2014/8/15 婧瑄說智權公司收入換新科目
   '將收入科目改用新科目 4101->4203,4121->4204,4111->4205,4131->4206
   Text1 = "正在將收入科目換新科目......"
   strSql = "UPDATE ACC031 SET AX305='4203',AX316='4203' where AX301='J' and AX305='4101' AND ax302 IN ( " & _
            "SELECT DISTINCT A0302 FROM ACC030,ACC031 WHERE A0301='J' AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301(+) AND A0302=AX302(+) AND AX305='4101')"
   adoTaie.Execute strSql
   strSql = "UPDATE ACC031 SET AX305='4204',AX316='4204' where AX301='J' and AX305='4121' AND ax302 IN ( " & _
            "SELECT DISTINCT A0302 FROM ACC030,ACC031 WHERE A0301='J' AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301(+) AND A0302=AX302(+) AND AX305='4121')"
   adoTaie.Execute strSql
   strSql = "UPDATE ACC031 SET AX305='4205',AX316='4205' where AX301='J' and AX305='4111' AND ax302 IN ( " & _
            "SELECT DISTINCT A0302 FROM ACC030,ACC031 WHERE A0301='J' AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301(+) AND A0302=AX302(+) AND AX305='4111')"
   adoTaie.Execute strSql
   strSql = "UPDATE ACC031 SET AX305='4206',AX316='4206' where AX301='J' and AX305='4131' AND ax302 IN ( " & _
            "SELECT DISTINCT A0302 FROM ACC030,ACC031 WHERE A0301='J' AND A0305 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and A0305 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' AND A0301=AX301(+) AND A0302=AX302(+) AND AX305='4131')"
   adoTaie.Execute strSql
   'end 2014/8/15
   
   Text1 = "正在調整因小數位捨去所造成的借貸方" & vbCrLf & "    差額......"
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select ax301, ax302, sum(ax307 - ax306) from acc031, acc030 where A0301='J' AND ax301 = a0301 and ax302 = a0302 and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by ax301, ax302", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      DoEvents
      If IsNull(adoacc020.Fields(2).Value) = False Then
         If adoacc020.Fields(2).Value <> 0 Then
            strSql = "update acc031 set ax306 = ax306 + " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax306 <> 0 AND AX305='1101' and rownum < 2"
            adoTaie.Execute strSql, lngEff
            If lngEff = 0 Then
               '更新貸方非規費科目的最小項次
               StrSQLa = "update acc031 set ax307 = ax307 - " & Val(adoacc020.Fields(2).Value) & " where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax303 IN (" & _
                         "SELECT MIN(AX303) FROM ACC031 where ax301 = '" & adoacc020.Fields("ax301").Value & "' and ax302 = '" & adoacc020.Fields("ax302").Value & "' and ax307 <> 0 AND SUBSTR(AX305,1,3)<>'220' )"
               cnnConnection.Execute StrSQLa, lngEff
            End If
         End If
      End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   
   adoTaie.Execute "insert into acc1z0 select 'J', a1r02, a1r03, a1r04, a1r04 from acc1r0 where a1r01 = 'JD' and a1r02 = " & (Val(Mid(MaskEdBox1.Text, 1, 3)) + 1911) & " and a1r03 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a1r03 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & ""
   adoTaie.Execute "UPDATE ACC0C0 SET A0C05= " & Val(FCDate(MaskEdBox2.Text)) & " where a0c04='J' "

   Screen.MousePointer = vbDefault
   Text1 = "" ': Text5 = "": Text6 = ""
   cboComp.ListIndex = 0 'Modify By Sindy 2020/4/17
   MsgBox MsgText(23), , MsgText(21)
CheckingJ:
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   If adoacc021.State = adStateOpen Then
      Text1 = "錯誤之傳票號碼: " & stra0201 & " 公司 " & stra0202 & " 項次: " & adoacc021.Fields("aX203").Value
      adoacc021.Close
   End If
   Screen.MousePointer = vbDefault
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'2014/2/17 end


