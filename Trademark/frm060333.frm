VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060333 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "外商每月請款點數統計表"
   ClientHeight    =   3765
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5685
   Begin VB.CheckBox Check1 
      Caption         =   "是否含離職人員"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3480
      TabIndex        =   20
      Top             =   765
      Width           =   1815
   End
   Begin VB.CommandButton CmdPrt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel(&E)"
      Default         =   -1  'True
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
      Left            =   600
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   3120
      Width           =   4692
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   10
      Left            =   3240
      TabIndex        =   19
      Top             =   2520
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   8
      Left            =   3240
      TabIndex        =   18
      Top             =   2105
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   7
      Left            =   3240
      TabIndex        =   17
      Top             =   1680
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   9
      Left            =   3600
      TabIndex        =   8
      Top             =   2452
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1402;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   8
      Left            =   2280
      TabIndex        =   7
      Top             =   2452
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1402;635"
      Value           =   "10012"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   7
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1402;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   6
      Left            =   2280
      TabIndex        =   5
      Top             =   2037
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1402;635"
      Value           =   "10701"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   4
      Left            =   2280
      TabIndex        =   3
      Top             =   1622
      Width           =   800
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1411;635"
      Value           =   "10001"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   5
      Left            =   3600
      TabIndex        =   4
      Top             =   1622
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1402;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
      VariousPropertyBits=   8388627
      Caption         =   "報表４統計年月："
      Size            =   "3413;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   2105
      Width           =   1935
      VariousPropertyBits=   8388627
      Caption         =   "報表３統計年月："
      Size            =   "3413;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1690
      Width           =   1935
      VariousPropertyBits=   8388627
      Caption         =   "報表２統計年月："
      Size            =   "3413;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Width           =   2415
      ForeColor       =   16711680
      VariousPropertyBits=   8388627
      Caption         =   "空白：表示全部"
      Size            =   "4260;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "6800;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   13
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "報表種類："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   3
      Left            =   3600
      TabIndex        =   2
      Top             =   1185
      Width           =   795
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1411;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   2
      Left            =   2280
      TabIndex        =   1
      Top             =   1185
      Width           =   800
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1411;635"
      Value           =   "10701"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1275
      Width           =   1935
      VariousPropertyBits=   8388627
      Caption         =   "報表１統計年月："
      Size            =   "3413;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frm060333"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2019/02/23 外商每月請款點數統計表
'Memo by Lydia 2019/02/23 使用Form 2.0 (Label和TextBox)
Option Explicit

Dim oText As MSForms.TextBox

Private Sub CmdPrt1_Click()

   If FormCheck = False Then Exit Sub
   
   'Added by Lydia 2021/06/28 另外整理外商員工檔
   'CREATE TABLE R060333_STAFF (STID VARCHAR2(8 CHAR),ST01 VARCHAR2(8 CHAR), ST02 VARCHAR2(12 CHAR),ST03 VARCHAR2(3 CHAR),
   'ST04 VARCHAR2(1 CHAR),ST16 VARCHAR2(3 CHAR),ST16NAME VARCHAR(20 CHAR) , ST70 VARCHAR2(1 CHAR),ST70NAME VARCHAR(20 CHAR));
   strSql = "DELETE FROM R060333_STAFF WHERE STID = '" & strUserNum & "' "
   cnnConnection.Execute strSql
   '在職
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01, ST02,ST03, '1'  AS ST04, ST16, nvl(ST70,6) ST70 " & _
               "FROM STAFF WHERE ST03 LIKE 'F1%' AND ST04='1' "
   cnnConnection.Execute strSql
   '留職停薪
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01, ST02,ST03, '1'  AS ST04, ST16, nvl(ST70,6) ST70 " & _
               "FROM STAFF,STAFF_CHANGE,(SELECT SC01 MNO,MAX(SC02) MDATE FROM STAFF_CHANGE WHERE SC04='F21' GROUP BY SC01) VT1 " & _
               "WHERE ST03 LIKE 'F1%' AND ST01=SC01(+) AND ST01=MNO(+) AND MNO=SC01(+) AND MDATE=SC02(+) AND '04'=SC03 " & _
               "AND ST01 NOT IN (SELECT ST01 FROM R060333_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   '離職
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01,  '*'||ST02 AS ST02,ST03, '2'  AS ST04, ST16, nvl(ST70,6) ST70 " & _
               "FROM STAFF WHERE ST03 LIKE 'F1%' AND ST04='2' "
   cnnConnection.Execute strSql

   '調職；曾經是外商，現在非外商(人員異動排除03離職,08退休,09撤職,10資遣)
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F10' AS ST03, '2'  AS ST04, ST16,nvl(ST70,6) ST70 " & _
               "FROM STAFF WHERE ST03 <>'F10' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F10' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060333_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F11' AS ST03, '2'  AS ST04, ST16,nvl(ST70,6) ST70 " & _
               "FROM STAFF WHERE ST03 <>'F11' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F11' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060333_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO R060333_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST70) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F12' AS ST03, '2'  AS ST04, ST16,nvl(ST70,6) ST70 " & _
               "FROM STAFF WHERE ST03 <>'F12' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F12' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060333_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   'end 2021/06/28
   
   Screen.MousePointer = vbHourglass
   CmdPrt1.Enabled = False
   
   If Trim(Combo1.Text) = "" Then '空白=全部
        Call Process1("1")
        Call Process1("2")
        Call Process1("3")
        Call Process1("4")
   Else
        Select Case Left(Combo1.Text, 1)
            Case "1"  '各區統計
                Call Process1("1")
            Case "2"  '代理人地理區
                Call Process1("2")
            Case "3"  '區別個人統計
                Call Process1(3)
            Case "4"  '組別個人年移動平均
                Call Process1(4)
        End Select
   End If
   '執行完不清除條件
   CmdPrt1.Enabled = True
   Screen.MousePointer = vbDefault
   'Modify by Amy 2021/06/22 原:strExcelPath 改中文字顯示
   MsgBox "Excel檔案產生完成！檔案位置：" & strExcelPathN
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
       MkDir strExcelPath
   End If
   
   '1 各區統計  10701~系統日前一個月
   '2 代理人地理區(只統計外商收文案件,不含含非外商收文但業務點收數歸外商人員)   10001~系統日前一個月
   '3 區別個人統計(因為107才開始分區故不算年移動平均    10701~系統日前一個月
   '4 組別個人年移動平均   10001~系統日前一個月  資料自10001開始抓但報表從10012開始
   txtFM2(3).Text = Left(TransDate(CompDate(1, -1, strSrvDate(1)), 1), 5)
   txtFM2(5).Text = txtFM2(3).Text
   txtFM2(7).Text = txtFM2(3).Text
   txtFM2(9).Text = txtFM2(3).Text
   
   Combo1.Clear
   Combo1.AddItem "1. 各區統計"
   Combo1.AddItem "2. 代理人地理區"
   Combo1.AddItem "3. 區別個人統計"
   Combo1.AddItem "4. 組別個人年移動平均"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060333 = Nothing
End Sub

' 畫面輸入檢查
Private Function FormCheck() As Boolean
Dim bolTmp As Boolean
Dim inA  As Integer

   FormCheck = False
   For Each oText In txtFM2
       txtFM2_Validate oText.Index, bolTmp
       If bolTmp = True Then
           Exit Function
       End If
   Next
   
   strExc(1) = Left(strSrvDate(2), 5)
   For intI = 2 To 9 Step 2
       inA = inA + 1
       If txtFM2(intI) = "" And txtFM2(intI + 1) = "" Then
            txtFM2(intI).SetFocus
            MsgBox "報表" & inA & "統計年月不可空白！", , MsgText(5)
            Exit Function
       End If
       If txtFM2(intI) = "" > txtFM2(intI + 1) = "" Then
            txtFM2(intI).SetFocus
            MsgBox "報表" & inA & "統計年月起值不可大於迄值！", , MsgText(5)
            Exit Function
       End If
       If txtFM2(intI) > strExc(1) Then
             txtFM2(intI).SetFocus
             MsgBox "報表" & inA & "統計年月起值不可大於系統日！", , MsgText(5)
             Exit Function
       ElseIf txtFM2(intI + 1) > strExc(1) Then
             txtFM2(intI).SetFocus
             MsgBox "報表" & inA & "統計年月迄值不可大於系統日！", , MsgText(5)
             Exit Function
       End If
       If txtFM2(intI) < "10001" Then
             txtFM2(intI).SetFocus
             MsgBox "報表" & inA & "統計年月起值不可小於10001！", , MsgText(5)
             Exit Function
       End If
   Next intI
   
   'Added by Lydia 2021/06/07 發生A8030的st16=null但是有設st70=2 , 資料已修正; 為了後面的作業能夠正常,增加檢查
   strExc(0) = "select st01,st02 from staff where st03 ='F21' and st16 is null and st01 not like 'F%' and st01>='91000' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       MsgBox RsTemp.Fields("ST01") & " " & RsTemp.Fields("ST02") & "尚未設定工程師組別，請連絡電腦中心！"
       Exit Function
   End If
   'end 2021/0/6/07
   
   FormCheck = True
End Function

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)

    Select Case Index
        Case 2, 3, 4, 5, 6, 7, 8, 9 '1.統計年月
            If PUB_CheckKeyInYYMM(txtFM2(Index)) = -1 Then
               GoTo EXITSUB
            End If
    End Select
    
    Exit Sub
    
EXITSUB:
    txtFM2(Index).SetFocus
    txtFM2_GotFocus Index
    Cancel = True
End Sub

'報表1-各區統計, 2-代理人地理區, 3-區別個人統計, 4-組別個人年移動平均
Private Sub Process1(ByVal aKind As String)
Dim stCon As String
Dim strMid01 As String, strMid02 As String
Dim intP As Integer '子查詢序號
Dim yycnt As Integer, mmcnt As Integer
Dim yymm As String
Dim yymm1 As String, mmcnt1 As Integer
Dim rsAD As New ADODB.Recordset
Dim StrSqlB As String, strSQLc As String
Dim strDate1 As String, StrDate2 As String '統計年月期限
Dim strTitle As String

   '統計年月
   Select Case aKind
        Case "1" '1各區統計
             strDate1 = txtFM2(2)
             StrDate2 = txtFM2(3)
             strTitle = "各區統計"
             'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
             ClearQueryLog (Me.Name)
             pub_QL05 = pub_QL05 & ";報表1.各區統計"
             pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(2) & "00-" & txtFM2(3) & "31"
            'end 2021/11/16
        Case "2" '2代理人地理區
             strDate1 = txtFM2(4)
             StrDate2 = txtFM2(5)
             strTitle = "代理人地理區"
             'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
             ClearQueryLog (Me.Name)
             pub_QL05 = pub_QL05 & ";報表2.代理人地理區"
             pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(4) & "00-" & txtFM2(5) & "31"
            'end 2021/11/16
        Case "3" '3區別個人統計
             strDate1 = txtFM2(6)
             StrDate2 = txtFM2(7)
             strTitle = "區別個人統計"
             'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
             ClearQueryLog (Me.Name)
             pub_QL05 = pub_QL05 & ";報表3.區別個人統計"
             pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(6) & "00-" & txtFM2(7) & "31"
            'end 2021/11/16
        Case "4" '4組別個人年移動平均 : 10001~系統日前一個月  資料自10001開始抓但報表從10012開始
             strDate1 = Left(TransDate(CompDate(1, -11, (txtFM2(8) + 191100) & "01"), 1), 5)
             StrDate2 = txtFM2(9)
             strTitle = "組別個人年移動平均"
             'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
             ClearQueryLog (Me.Name)
             pub_QL05 = pub_QL05 & ";報表4.組別個人年移動平均"
             pub_QL05 = pub_QL05 & ";統計年月:" & Left(TransDate(CompDate(1, -11, (txtFM2(8) + 191100) & "01"), 1), 5) & "00-" & txtFM2(9) & "31"
            'end 2021/11/16
   End Select
   
   If strDate1 <> "" Then
        stCon = stCon & " AND A1K02>=" & strDate1 & "00"
   End If
   If StrDate2 <> "" Then
        stCon = stCon & " AND A1K02<=" & StrDate2 & "31"
   End If
   
   cnnConnection.BeginTrans
       strSql = "DELETE FROM R060333 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' "
       cnnConnection.Execute strSql
       
       strSql = "DELETE FROM R060333_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' "
       cnnConnection.Execute strSql

'---固定表格
'CREATE TABLE R060333 (FORMNAME VARCHAR2(20),ID VARCHAR2(6),TKIND VARCHAR2(1),TNAME VARCHAR2(20 CHAR),A1N04 VARCHAR2(6),DN04 VARCHAR2(1),DN70 VARCHAR2(1));
'CREATE TABLE R060333_1 (FORMNAME VARCHAR2(20),ID VARCHAR2(6),TKIND VARCHAR2(1),A1N04 VARCHAR2(6),DN70 VARCHAR2(1),
'YY00 VARCHAR2(4),MM01 NUMBER(10,3),MM02 NUMBER(10,3),MM03 NUMBER(10,3),MM04 NUMBER(10,3),MM05 NUMBER(10,3),MM06 NUMBER(10,3),MM07 NUMBER(10,3),MM08 NUMBER(10,3),MM09 NUMBER(10,3),MM10 NUMBER(10,3),MM11 NUMBER(10,3),MM12 NUMBER(10,3),YYTOTAL NUMBER(13,3));
'ALTER TABLE R060333 ADD PRIMARY KEY (FORMNAME,ID,TKIND,A1N04,DN70);
'ALTER TABLE R060333_1 ADD PRIMARY KEY (FORMNAME,ID,TKIND,A1N04,DN70,YY00);
        'strSql = "INSERT INTO R060333_1(FORMNAME, ID, TKIND, A1N04, YY00, MM01, MM02, MM03, MM04, MM05, MM06, MM07, MM08, MM09, MM10, MM11, MM12 ) " & _
                    "SELECT FORMNAME, ID, TKIND, A1N04,'" & yycnt & "')"
                    
'-------------------------------------------
   
   '各區小計,合計
   '一區
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','1TOT','1')"
   cnnConnection.Execute strSql
   '二區
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','2TOT','2')"
   cnnConnection.Execute strSql
   '三區
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','3TOT','3')"
   cnnConnection.Execute strSql
   '英文組合計
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','4TOT','4')"
   cnnConnection.Execute strSql
   '日文組
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','5TOT','5')"
   cnnConnection.Execute strSql
   '未分組
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','6TOT','6')"
   cnnConnection.Execute strSql
   '外商合計
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','7TOT','7')"
   cnnConnection.Execute strSql
   '其他部門合計
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','8TOT','8')"
   cnnConnection.Execute strSql
   '總請款點數
   strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                    "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','TOTAL','9')"
   cnnConnection.Execute strSql
   
   If aKind = "1" Or aKind = "3" Or aKind = "4" Then
        '承辦人
        'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
        'strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                     "select '" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "', A1N04,decode(ST70,null,'6',ST70) st70 from acc1k0,acc1n0,staff where nvl(a1k12,0)=0 and a1k25 is null " & stCon & _
                     " and a1k01=a1n01(+) and '1'=a1n02(+) and a1n04=st01(+) and st03 like 'F1%' group by A1N04,decode(ST70,null,'6',ST70) "
        strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                     "select '" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "', A1N04,decode(ST70,null,'6',ST70) st70 from acc1k0,acc1n0,R060333_staff " & _
                     " where nvl(a1k12,0)=0 and a1k25 is null " & stCon & _
                     " and a1k01=a1n01(+) and '1'=a1n02(+) and a1n04=st01(+) AND '" & strUserNum & "'=STID(+) and st03 like 'F1%' group by A1N04,decode(ST70,null,'6',ST70) "
        cnnConnection.Execute strSql, intI
        If aKind = "3" Or aKind = "4" Then
            '因A4009黃咸達106/12為三區,107/7換至一區,該員依年月拆成二筆記錄,故工作檔要加區別DN70欄
            strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                             "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','A4009','3')"
            cnnConnection.Execute strSql, intI
            'Added by Lydia 2021/06/28 林靖傑A6015由英文組1改成英文組3(2021/7/1開始)
            If strSrvDate(1) >= "20210701" Then
                strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                                 "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','A6015','1')"
                cnnConnection.Execute strSql, intI
            End If
            'end 2021/06/28
        End If
   End If
   If aKind = "1" Or aKind = "2" Then
        '折讓點數
        strSql = "INSERT INTO R060333 (FORMNAME,ID,TKIND,TNAME,A1N04,DN70) " & _
                         "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & aKind & "', '" & strTitle & "','ADIST','A')"
        cnnConnection.Execute strSql
   End If
   
    '統計資料-暫存檔
    '--------------組合語法
    Select Case aKind
        Case "1", "3", "4" ' 1各區統計, 3區別個人統計, 4組別個人年移動平均
            stCon = "SELECT B.A1N04, ST02, ST01, B.DN70, B.DN04 "
            'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
            'StrSqlB = "FROM R060333 B,STAFF"
            'strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & aKind & "' AND B.A1N04=ST01(+)"
            StrSqlB = "FROM R060333 B,R060333_STAFF"
            strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & aKind & "' AND B.A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) "
            'end 2021/06/28
        Case "2"    '2代理人地理區
            stCon = "SELECT B.A1N04, B.DN70, B.DN04 "
            StrSqlB = "FROM R060333 B "
            strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & aKind & "' "
    End Select
    strMid01 = " AND " '組合語法- 抓所有欄位加起來>0
    
    intP = 1  '子查詢序號
    For yycnt = Val(Left(strDate1, 3)) To Val(Left(StrDate2, 3))
       For mmcnt = 1 To 12
          If mmcnt = 1 Then
             strSql = "INSERT INTO R060333_1(FORMNAME, ID, TKIND, A1N04,DN70, YY00 ) " & _
                         "SELECT FORMNAME, ID, TKIND, A1N04,DN70, '" & yycnt & "' FROM R060333 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' "
             cnnConnection.Execute strSql, intI
             StrSqlB = StrSqlB & ", (SELECT * FROM R060333_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND YY00='" & yycnt & "') X" & intP
             strSQLc = strSQLc & " AND B.A1N04=X" & intP & ".A1N04 AND B.DN70=X" & intP & ".DN70"
          End If
          yymm = yycnt * 100 + mmcnt
          If yymm >= strDate1 And yymm <= Val(StrDate2) Then '超過期限不抓資料
                '組合語法
                If aKind <> "4" Then
                    stCon = stCon & ", X" & intP & ".MM" & Format(mmcnt, "00") & " AS D" & yymm 'ex: D10012 (100年12月)
                    strMid01 = strMid01 & "NVL(D" & yymm & ",0)+"
                ElseIf yymm >= txtFM2(8) Then '4組別個人年移動平均:報表起值
                    stCon = stCon & ", X" & intP & ".MM" & Format(mmcnt, "00") & " AS D" & yymm 'ex: D10012 (100年12月)
                    strMid01 = strMid01 & "NVL(D" & yymm & ",0)+"
                End If
                '------------------組合語法
                Select Case aKind
                      Case "1", "3", "4" ' 1各區統計, 3區別個人統計, 4組別個人年移動平均
                            '1.會分配給其他部門,故不限制只抓外商人員 2.只抓有點數的請款單 3.X10705262為內商收文但業務點收數歸外商A6005
                            'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                            'strSql = "select a1n04,sum(TOT) TOT,st70,st03 from " & _
                                     "(select a1n04,sum(a1n05) TOT,decode(ST70,null,'6',ST70) st70,st03 from acc1k0,acc1n0,staff where a1k01 in " & _
                                     "   (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%') " & _
                                     "    and a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0)>0 AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+)" & _
                                     "  group by a1n04,decode(ST70,null,'6',ST70),st03 " & _
                                     " union " & _
                                     " select a1n04,sum(a1n05) TOT,decode(ST70,null,'6',ST70) ST70,st03 from acc1k0,acc1n0,staff where a1k01 in " & _
                                     "   (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1') " & _
                                     "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) and st03 like 'F1%'" & _
                                     "  group by a1n04,decode(ST70,null,'6',ST70),st03 " & _
                                     ") group by a1n04,ST70,st03 order by a1n04"
                            strSql = "select a1n04,sum(TOT) TOT,st70,st03 from " & _
                                     "(select a1n04,sum(a1n05) TOT,decode(ST70,null,'6',ST70) st70,st03 from acc1k0,acc1n0,staff where a1k01 in " & _
                                     "   (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%') " & _
                                     "    and a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0)>0 AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) " & _
                                     "  group by a1n04,decode(ST70,null,'6',ST70),st03 " & _
                                     " union " & _
                                     " select a1n04,sum(a1n05) TOT,decode(ST70,null,'6',ST70) ST70,st03 from acc1k0,acc1n0,r060333_staff where a1k01 in " & _
                                     "   (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1') " & _
                                     "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) AND '" & strUserNum & "'=STID(+) and st03 like 'F1%'" & _
                                     "  group by a1n04,decode(ST70,null,'6',ST70),st03 " & _
                                     ") group by a1n04,ST70,st03 order by a1n04"
                      Case "2"  '2代理人地理區
                            '只抓有點數的請款單
                            strSql = "select decode(substr(fa10,1,3),'011','5TOT',decode(substr(na02,1,1),'A','1TOT','B','1TOT',decode(substr(na02,1,2),'C0','1TOT','C1','2TOT','C2','3TOT','C3','2TOT','C4','1TOT',na02))) a1n04,sum(a1n05) TOT " & _
                                     "from acc1k0,acc1n0,fagent,nation where a1k01 in " & _
                                     "(select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & "  and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%') " & _
                                     "   and a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0)>0 AND A1K01=A1N01(+) AND '1'=A1N02(+)  AND SUBSTR(A1K03,1,8)=FA01(+) AND SUBSTR(A1K03,9,1)=FA02(+) AND FA10=NA01(+) " & _
                                     "  group by decode(substr(fa10,1,3),'011','5TOT',decode(substr(na02,1,1),'A','1TOT','B','1TOT',decode(substr(na02,1,2),'C0','1TOT','C1','2TOT','C2','3TOT','C3','2TOT','C4','1TOT',na02))) order by 1"
                End Select
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                    RsTemp.MoveFirst
                    Select Case aKind
'--------------------------------------------------------------
                          Case "1", "3" ' 1各區統計, 3區別個人統計
                                Do While Not RsTemp.EOF
                                     '因A4009黃咸達106/12為三區,107/7換至一區,該員依年月拆成二筆記錄
                                     If "" & RsTemp.Fields("A1N04") = "A4009" Then
                                          If yymm >= 10707 Then
                                                 strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                             "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='1' AND YY00='" & yycnt & "' "
                                          Else
                                                 strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                             "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='3' AND YY00='" & yycnt & "' "
                                          End If
                                          cnnConnection.Execute strSql, intI
                                     'Added by Lydia 2021/06/28 林靖傑A6015由英文組1改成英文組3(2021/7/1開始)
                                     ElseIf "" & RsTemp.Fields("A1N04") = "A6015" Then
                                          If yymm >= 11007 Then
                                                 strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                             "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='3' AND YY00='" & yycnt & "' "
                                          Else
                                                 strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                             "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='1' AND YY00='" & yycnt & "' "
                                          End If
                                          cnnConnection.Execute strSql, intI
                                     'end 2021/06/28
                                     ElseIf Left("" & RsTemp.Fields("ST03"), 2) <> "F1" Then  '非外商人員只累加計8TOT
                                     Else
                                          strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                      "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='" & RsTemp.Fields("ST70") & "' AND YY00='" & yycnt & "' "
                                          cnnConnection.Execute strSql, intI
                                     End If
                                     '同時加入各區小計,四區陳鳳英,洪琬姿僅一人故除外
                                     If "" & RsTemp.Fields("ST70") <> "4" Then
                                         '因A4009黃咸達106/12為三區,107/7換至一區,該員依年月拆成二筆記錄
                                         If "" & RsTemp.Fields("A1N04") = "A4009" Then
                                              If yymm >= 10707 Then
                                                     strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                                 "AND A1N04='1TOT' AND DN70='1' AND YY00='" & yycnt & "' "
                                              Else
                                                     strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                                 "AND A1N04='3TOT' AND DN70='3' AND YY00='" & yycnt & "' "
                                              End If
                                              cnnConnection.Execute strSql, intI
                                         'Added by Lydia 2021/06/28 林靖傑A6015由英文組1改成英文組3(2021/7/1開始)
                                         ElseIf "" & RsTemp.Fields("A1N04") = "A6015" Then
                                              If yymm >= 11007 Then
                                                     strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                                 "AND A1N04='3TOT' AND DN70='3' AND YY00='" & yycnt & "' "
                                              Else
                                                     strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                                 "AND A1N04='1TOT' AND DN70='1' AND YY00='" & yycnt & "' "
                                              End If
                                              cnnConnection.Execute strSql, intI
                                         'end 2021/06/28
                                         ElseIf Left("" & RsTemp.Fields("ST03"), 2) <> "F1" Then  '非外商人員只累加計8TOT
                                              strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                          "AND A1N04='8TOT' AND DN70='8' AND YY00='" & yycnt & "' "
                                              cnnConnection.Execute strSql, intI
                                         Else
                                              strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                          "AND A1N04='" & RsTemp.Fields("ST70") & "TOT' AND DN70='" & RsTemp.Fields("ST70") & "' AND YY00='" & yycnt & "' "
                                              cnnConnection.Execute strSql, intI
                                         End If
                                     End If
                                     RsTemp.MoveNext
                                Loop
                                '英文組合計
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, STAFF where A1N04=ST01(+) and ST01 is not null AND DN70<='4' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='4TOT' AND DN70='4' AND YY00='" & yycnt & "' "
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, R060333_STAFF " & _
                                                    " where A1N04=ST01(+) AND '" & strUserNum & "'=STID(+)  and ST01 is not null AND DN70<='4' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='4TOT' AND DN70='4' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '外商合計
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, STAFF where A1N04=ST01(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='7TOT' AND DN70='7' AND YY00='" & yycnt & "' "
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, R060333_STAFF " & _
                                                    " where A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='7TOT' AND DN70='7' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '總請款點數, 要含非外商收文但業務點收數歸外商人員X10705262
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "select sum(總請款點數) from " & _
                                         "(select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總請款點數 from acc1k0 where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')" & _
                                         " union " & _
                                         " select sum(a1n05) 分配點數 from acc1k0,acc1n0,staff where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1')" & _
                                         "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) and st03 like 'F1%' " & _
                                         ")"
                                strSql = "select sum(總請款點數) from " & _
                                         "(select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總請款點數 from acc1k0 where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')" & _
                                         " union " & _
                                         " select sum(a1n05) 分配點數 from acc1k0,acc1n0,R060333_staff where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1')" & _
                                         "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) AND '" & strUserNum & "'=STID(+) and st03 like 'F1%' " & _
                                         ")"
                                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                If intI = 1 Then
                                    RsTemp.MoveFirst
                                    strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                "AND A1N04='TOTAL' AND DN70='9' AND YY00='" & yycnt & "' "
                                    cnnConnection.Execute strSql, intI
                                End If
                                If aKind = "1" Then '不計折讓: 3區別個人統計
                                    '折讓點數
                                    strSql = "select nvl(sum(nvl(a1k06,0)-nvl(a1k36,0))/1000,0) 折讓點數 from acc1k0 where a1k01 in " & _
                                             "(select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')"
                                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                    If intI = 1 Then
                                        RsTemp.MoveFirst
                                        strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                    "AND A1N04='ADIST' AND DN70='A' AND YY00='" & yycnt & "' "
                                        cnnConnection.Execute strSql, intI
                                    End If
                                End If
'--------------------------------------------------------------
                          Case "2" '2代理人地理區
                                Do While Not RsTemp.EOF
                                    strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='" & Left("" & RsTemp.Fields("A1N04"), 1) & "' AND YY00='" & yycnt & "' "
                                    cnnConnection.Execute strSql, intI
                                    RsTemp.MoveNext
                                Loop
                                '英文組合計
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1 where DN70<='4' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='4TOT' AND DN70='4' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '外商合計
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1 where DN70>='4' AND DN70<'7' and substr(A1N04,2)='TOT' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='7TOT' AND DN70='7' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '總請款點數
                                strSql = "select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總請款點數 from acc1k0 where a1k01 in " & _
                                         "(select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')"
                                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                If intI = 1 Then
                                    RsTemp.MoveFirst
                                    strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                "AND A1N04='TOTAL' AND DN70='9' AND YY00='" & yycnt & "' "
                                    cnnConnection.Execute strSql, intI
                                End If
                                '折讓點數
                                strSql = "select nvl(sum(nvl(a1k06,0)-nvl(a1k36,0))/1000,0) 折讓點數 from acc1k0 where a1k01 in " & _
                                         "(select distinct a1k01 from acc1k0,staff,caseprogress where  a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')"
                                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                If intI = 1 Then
                                    RsTemp.MoveFirst
                                    strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                "AND A1N04='ADIST' AND DN70='A' AND YY00='" & yycnt & "' "
                                    cnnConnection.Execute strSql, intI
                                End If
'--------------------------------------------------------------
                          Case "4"  '4組別個人年移動平均
                                Do While Not RsTemp.EOF
                                     If Left("" & RsTemp.Fields("ST03"), 2) = "F1" Then  '非外商人員只累加計8TOT
                                          strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                      "AND A1N04='" & RsTemp.Fields("A1N04") & "' AND DN70='" & RsTemp.Fields("ST70") & "' AND YY00='" & yycnt & "' "
                                          cnnConnection.Execute strSql, intI
                                     End If
                                     '同時加入各區小計,四區陳鳳英,洪琬姿僅一人故除外
                                     If "" & RsTemp.Fields("ST70") <> "4" Then
                                         If Left("" & RsTemp.Fields("ST03"), 2) <> "F1" Then  '非外商人員只累加計8TOT
                                              strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields("TOT") & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                          "AND A1N04='8TOT' AND DN70='8' AND YY00='" & yycnt & "' "
                                              cnnConnection.Execute strSql, intI
                                         Else
                                              strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                          "AND A1N04='" & RsTemp.Fields("ST70") & "TOT' AND DN70='" & RsTemp.Fields("ST70") & "' AND YY00='" & yycnt & "' "
                                              cnnConnection.Execute strSql, intI
                                         End If
                                     End If
                                     RsTemp.MoveNext
                                Loop
                                '英文組合計
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, STAFF where A1N04=ST01(+) and ST01 is not null AND DN70<='4' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='4TOT' AND DN70='4' AND YY00='" & yycnt & "' "
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, R060333_STAFF " & _
                                                 "where A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) and ST01 is not null AND DN70<='4' AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='4TOT' AND DN70='4' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '外商合計
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, STAFF where A1N04=ST01(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='7TOT' AND DN70='7' AND YY00='" & yycnt & "' "
                                strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060333_1, R060333_STAFF " & _
                                                    "where A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & yycnt & "') " & _
                                            "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04='7TOT' AND DN70='7' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '總請款點數, 要含非外商收文但業務點收數歸外商人員X10705262
                                'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF
                                'strSql = "select sum(總請款點數) from " & _
                                         "(select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總請款點數 from acc1k0 where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')" & _
                                         " union " & _
                                         " select sum(a1n05) 分配點數 from acc1k0,acc1n0,staff where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1')" & _
                                         "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) and st03 like 'F1%' " & _
                                         ")"
                                strSql = "select sum(總請款點數) from " & _
                                         "(select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總請款點數 from acc1k0 where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F1%')" & _
                                         " union " & _
                                         " select sum(a1n05) 分配點數 from acc1k0,acc1n0,R060333_staff where a1k01 in " & _
                                         "  (select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>=" & yymm & "01" & " and a1k02<=" & yymm & "31" & " and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and substr(cp12,1,2)<>'F1')" & _
                                         "    AND A1K01=A1N01(+) AND '1'=A1N02(+) and a1n04=st01(+) AND '" & strUserNum & "'=STID(+) and st03 like 'F1%' " & _
                                         ")"
                                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                If intI = 1 Then
                                    RsTemp.MoveFirst
                                    strSql = "UPDATE R060333_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  " & _
                                                "AND A1N04='TOTAL' AND DN70='9' AND YY00='" & yycnt & "' "
                                    cnnConnection.Execute strSql, intI
                                End If
                    End Select
                End If
          End If
       Next mmcnt
       intP = intP + 1
   Next yycnt
   
   '4組別個人年移動平均=> 年移動平均數：10012為10001~10012平均,10101為10002~10101平均
   If aKind = "4" Then
        For yycnt = Val(Left(StrDate2, 3)) To Val(Left(strDate1, 3)) Step -1
           For mmcnt = 12 To 1 Step -1
              yymm = yycnt * 100 + mmcnt
              '畫面起始值為報表起始值
              'If Val(yymm) >= Val(Left(strDate1, 3) & "12") And Val(yymm) <= Val(strDate2) Then
              If Val(yymm) >= strDate1 And Val(yymm) <= Val(StrDate2) Then
                 For mmcnt1 = 1 To 11 '往前推11個月
                    yymm1 = Val(yymm) - mmcnt1
                    If (Val(Right(yymm1, 2)) <= 0 Or Val(Right(yymm1, 2)) >= 90) Then '跨年
                       yymm1 = (yycnt - 1) * 100 + mmcnt - mmcnt1 + 12
                    End If
                    strSql = "UPDATE R060333_1 A SET A.MM" & Right(yymm, 2) & "=NVL(A.MM" & Right(yymm, 2) & ",0) " & _
                                 "+(SELECT NVL(B.MM" & Right(yymm1, 2) & ",0) FROM R060333_1 B WHERE A.FORMNAME=B.FORMNAME AND A.ID=B.ID AND A.TKIND=B.TKIND AND A.A1N04=B.A1N04 AND A.DN70=B.DN70 AND B.YY00='" & Mid(yymm1, 1, Len(yymm1) - 2) & "') " & _
                                 "WHERE A.FORMNAME = '" & Me.Name & "' AND A.ID = '" & strUserNum & "' AND A.TKIND = '" & aKind & "' AND A.YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
                    cnnConnection.Execute strSql, intI
                 Next mmcnt1
                 '平均
                 strSql = "UPDATE R060333_1 SET MM" & Right(yymm, 2) & "=MM" & Right(yymm, 2) & "/12 " & _
                             "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' AND YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
                 cnnConnection.Execute strSql, intI
              End If
           Next mmcnt
        Next yycnt
   End If
   
   '名單刪除離職人員但小計及合計是包含所有人的數字,故工作檔增加DN04,離職人員才有值
   If aKind = "1" Or aKind = "3" Or aKind = "4" Then   ' 1各區統計, 3區別個人統計, 4組別個人年移動平均
        'Modified by Lydia 2021/06/28 staff改用暫存檔; STAFF=>R060333_STAFF; 增加「是否含離職人員」的判斷
        'strSql = "UPDATE R060333 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "'  AND A1N04 NOT IN " & _
                 "(SELECT ST01 FROM STAFF,STAFF_CHANGE,(SELECT ST01 NO,MAX(SC02) MAXDATE FROM STAFF,STAFF_CHANGE WHERE ST03 LIKE 'F1%' AND ST01=SC01(+) GROUP BY ST01) " & _
                 "WHERE ST03 LIKE 'F1%' AND ST01=NO(+) AND NO=SC01(+) AND MAXDATE=SC02(+) AND '04'=SC03(+) AND (ST04='1' OR SC02 IS NOT NULL))"
        'cnnConnection.Execute strSql, intI
        If Check1.Value = 0 Then
            strSql = "UPDATE R060333 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & aKind & "' " & _
                        "AND A1N04 IN (SELECT ST01 FROM R060333_STAFF WHERE STID='" & strUserNum & "' AND ST04='2' AND ST03 LIKE 'F1%' ) "
            cnnConnection.Execute strSql, intI
        End If
        'end 2021/06/28
    End If
   cnnConnection.CommitTrans
   
   '產生Excel檔:
   Select Case aKind
        Case "1" '1各區統計
            strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
            strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
            '4英文組
            strExc(2) = "SELECT DECODE(DN70,'1','1英文組一區','2','2英文組二區','3','3英文組三區','4','4英文組','5','5日文組','6','6未分組',DN70) 組別,A1N04 編號,ST02 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  DN70='4' AND ST01 IS NOT NULL AND DN04 IS NULL"
            strExc(2) = strExc(2) & strMid01 & ">0"
            '各組-小計
            strExc(3) = "SELECT DECODE(DN70,'1','1英文組一區','2','2英文組二區','3','3英文組三區','4','4英文組合計','5','5日文組','6','6未分組',DN70) 組別,' ' 編號,'小計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04>'1' AND A1N04<'7' AND INSTR(A1N04,'TOT')=2 "
            '外商合計
            strExc(4) = "SELECT '7外商合計' 組別,' ' 編號,'外商合計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='7TOT' "
            '其他部門
            strExc(5) = "SELECT '8其他部門' 組別,' ' 編號,'其他部門' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='8TOT' "
            '外商總請款點數
            strExc(6) = "SELECT '0外商總請款點數' 組別,'' 編號,'外商總點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='TOTAL' "
            '折讓點數
            strExc(7) = "SELECT 'A折讓點數' 組別,'' 編號,'折讓點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='ADIST' "
            
            strSql = strExc(2) & " UNION " & strExc(3) & " UNION " & strExc(4) & " UNION " & strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7)
            strSql = strSql & "ORDER BY 1, 2"

        Case "2" '2代理人地理區
            strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
            strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
            '4英文組
            '(依A1K03之地理區, 不管分配點數所以沒有未分組及分配其他部門,只統計外商收文案件,不含非外商收文但業務點收數歸外商人員)
            strExc(2) = "SELECT DECODE(DN70,'1','1英文組亞洲(日本除外)','2','2英文組美洲+非洲','3','3英文組歐洲','4','4英文組合計','5','5日文組','6','6未分組',DN70) 地理區,'小計' 姓名," & _
                            strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04>'1' AND A1N04<'6' AND INSTR(A1N04,'TOT')=2"
            '外商合計
            strExc(4) = "SELECT '7外商合計' 組別,'外商合計' 姓名, " & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='7TOT' "
            '外商總請款點數
            strExc(6) = "SELECT '0外商總請款點數' 組別,'外商總點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='TOTAL' "
            '折讓點數
            strExc(7) = "SELECT 'A折讓點數' 組別,'折讓點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='ADIST' "
            
            strSql = strExc(2) & " UNION " & strExc(4) & " UNION " & strExc(6) & " UNION " & strExc(7)
            strSql = strSql & "ORDER BY 1, 2"
   
        Case "3" '3區別個人統計
            strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
            strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
            '個人: 因檔案內有組別小計,合計及總請款點數資料,故只抓讀得到員工檔且在職的資料
            strExc(2) = "SELECT DECODE(DN70,'1','1英文組一區','2','2英文組二區','3','3英文組三區','4','4英文組','5','5日文組','6','6未分組',DN70) 組別,A1N04 編號,ST02 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
            strExc(2) = strExc(2) & strMid01 & ">0"
            '各組-小計
            strExc(3) = "SELECT DECODE(DN70,'1','1英文組一區','2','2英文組二區','3','3英文組三區','4','4英文組合計','5','5日文組','6','6未分組',DN70) 組別,' ' 編號,'小計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04>'1' AND A1N04<'7' AND INSTR(A1N04,'TOT')=2 "
            '外商合計
            strExc(4) = "SELECT '7外商合計' 組別,' ' 編號,'外商合計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='7TOT' "
            '其他部門
            strExc(5) = "SELECT '8其他部門' 組別,' ' 編號,'其他部門' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='8TOT' "
            '外商總請款點數
            strExc(6) = "SELECT '9外商總請款點數' 組別,'' 編號,'外商總點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='TOTAL' "
            
            strSql = strExc(2) & " UNION " & strExc(3) & " UNION " & strExc(4) & " UNION " & strExc(5) & " UNION " & strExc(6)
            strSql = strSql & "ORDER BY 1, 2"
   
        Case "4" '組別個人年移動平均
            strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
            strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
            '個人: 因檔案內有組別小計,合計及總請款點數資料,故只抓讀得到員工檔且在職的資料
            strExc(2) = "SELECT DECODE(DN70,'1','1英文組','2','1英文組','3','1英文組','4','1英文組','5','5日文組','6','6未分組',DN70) 組別,A1N04 編號,ST02 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
            strExc(2) = strExc(2) & strMid01 & ">0"
            '各組-小計, 英文組小計只抓4TOT
            strExc(3) = "SELECT DECODE(DN70,'1','1英文組','2','1英文組','3','1英文組','4','1英文組','5','5日文組','6','6未分組',DN70) 組別,' ' 編號,'小計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04>'4' AND A1N04<'7' AND INSTR(A1N04,'TOT')=2 "
            '外商合計
            strExc(4) = "SELECT '7外商合計' 組別,' ' 編號,'外商合計' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='7TOT' "
            '其他部門
            strExc(5) = "SELECT '8其他部門' 組別,' ' 編號,'其他部門' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='8TOT' "
            '外商總請款點數
            strExc(6) = "SELECT '9外商總請款點數' 組別,'' 編號,'外商總點數' 姓名," & _
                              strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE A1N04='TOTAL' "

            strSql = strExc(2) & " UNION " & strExc(3) & " UNION " & strExc(4) & " UNION " & strExc(5) & " UNION " & strExc(6)
            strSql = strSql & "ORDER BY 1, 2"
   End Select
   
   If rsAD.State = adStateOpen Then
       rsAD.Close
   End If
   rsAD.CursorLocation = adUseClient
   rsAD.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsAD.RecordCount > 0 Then
         InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
         If aKind = "1" Or aKind = "2" Then
             Call ProcExcelSave1(aKind, rsAD, strDate1, StrDate2)
         ElseIf aKind = "3" Or aKind = "4" Then
             Call ProcExcelSave3(aKind, rsAD, strDate1, StrDate2)
         End If
   'Added by Lydia 2021/11/16
   Else
        InsertQueryLog (0)
   'end 2021/11/16
   End If

End Sub

'產生Excel檔案: 1各區統計, 2代理人地理區
Private Sub ProcExcelSave1(ByVal iType As String, ByRef m_Rst As ADODB.Recordset, ByVal mDate1 As String, ByVal mDate2 As String)
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String '組別
Dim intPage As Integer '工作表編號
Dim xCols As Integer '行位置
Dim MaxCols As Integer '最大行位置
Dim rowT1 As Integer '外商總點數的位置
Dim iCall As Integer '表1,2
Dim intJ As Integer
Dim strX As String
Dim strA1n04 As String
Dim tmpArray As Variant

On Error GoTo ErrHnd
   
   Select Case iType
        Case "1" '1各區統計
            '檔名：外商請款點數分析1各區統計xxxxx~xxxxx
            strExc(1) = strSrvDate(1) & "_外商請款點數1各區統計" & Val(mDate1) & "~" & Val(mDate2)
        Case "2" '2代理人地理區
            '檔名：外商請款點數分析2代理人地理區xxxxx~xxxxx
            strExc(1) = strSrvDate(1) & "_外商請款點數2代理人地理區" & Val(mDate1) & "~" & Val(mDate2)
   End Select
   
   strFileName = strExcelPath & strExc(1) & MsgText(43)
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint1.SheetsInNewWorkbook = 2 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint1.Workbooks.add
    xlsPoint1.Visible = False '預設不顯示

   For iCall = 1 To 2
       '備註
       If iType = "2" And iCall = 2 Then  '2代理人地理區
           iRow = iRow + 1
           wksPoint1.Range("A" & iRow).Value = "1.只統計外商收文案件"
           wksPoint1.Range("A" & iRow).Font.Color = vbRed
           iRow = iRow + 1
           wksPoint1.Range("A" & iRow).Value = "2.以案件之FC代理人統計"
           wksPoint1.Range("A" & iRow).Font.Color = vbRed
       End If
       
       m_Rst.MoveFirst
       iRow = 1
       xCols = 1
       intPage = iCall
       
       Set wksPoint1 = xlsPoint1.Worksheets(intPage)
       xlsPoint1.Sheets(intPage).Select '選擇工作表
       xlsPoint1.ActiveWindow.DisplayZeros = False '設工作表的零值不顯示
       xlsPoint1.Worksheets(intPage).Name = IIf(iCall = "1", "請款點數", "百分比") '工作表名稱
       '欄位抬頭
       For intJ = 0 To m_Rst.Fields.Count - 1
           strX = Pub_NumberToSystem26(xCols + intJ)
           If iType = "1" Then '1各區統計
                strExc(3) = "D2" '凍結窗格-位置
                strA1n04 = "組別"
                Select Case intJ
                     Case 0: '組別
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 16
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case 1: '編號
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 7
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case 2: '姓名
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 11
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case Else '年月統計量
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 10
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlRight
                End Select
           ElseIf iType = "2" Then '2代理人地理區
                strExc(3) = "C2" '凍結窗格-位置
                strA1n04 = "地理區"
                Select Case intJ
                     Case 0: '地理區
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 24
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case 1: '姓名
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 11
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case Else '年月統計量
                         wksPoint1.Range(strX & ":" & strX).ColumnWidth = 10
                         wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlRight
                End Select
           End If
           '年月標題
           strExc(2) = Replace("" & m_Rst.Fields(intJ).Name, "D", "")
           If Val(strExc(2)) > 0 Then
                If Val(strExc(2)) <= Val(mDate2) Then
                    wksPoint1.Range(strX & iRow).Value = strExc(2)
                    MaxCols = xCols + intJ
                End If
           Else
                'Added by Lydia 2021/06/30 姓名抬頭加註記
                If Check1.Value = 1 And iType = "1" And strExc(2) = "姓名" Then
                    strExc(2) = "姓名(*離職)"
                End If
                'end 2021/06/30
                wksPoint1.Range(strX & iRow).Value = strExc(2)
           End If
       Next intJ
       
       wksPoint1.Range(iRow & ":" & iRow).HorizontalAlignment = xlCenter '置中
       wksPoint1.Range(strExc(3)).Select
       xlsPoint1.ActiveWindow.FreezePanes = True '凍結窗格
       wksPoint1.Range("A1").Select
       ReDim tmpArray(1 To MaxCols)
       
       iRow = iRow + 1
       rowT1 = iRow
       Do While Not m_Rst.EOF
            If strGrp <> "" & m_Rst.Fields(strA1n04) And InStr("1,5,6,7,8,A", Left("" & m_Rst.Fields(strA1n04), 1)) > 0 Then
                If iCall = 2 And "" & m_Rst.Fields(strA1n04) = "A折讓點數" Then '百分比-不計算折讓點數
                   GoTo JumpToNext
                End If
                '不同組別多跳一行 (英文組分4組)
                iRow = iRow + 1
            End If
            wksPoint1.Range(Pub_NumberToSystem26(xCols) & iRow).Value = "" & m_Rst.Fields(strA1n04)
            
            If iCall = 1 Then '請款點數
                '因為office2013逐筆輸入過慢,改成陣列輸入
                For intJ = 1 To MaxCols - 1
                    tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                Next intJ
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                If iType = "1" Then
                    wksPoint1.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "#,##0.000"
                ElseIf iType = "2" Then
                    wksPoint1.Range(Pub_NumberToSystem26(xCols + 2) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "#,##0.000"
                End If
            ElseIf iCall = 2 Then '百分比
                For intJ = 1 To MaxCols - 1
                    strX = Pub_NumberToSystem26(xCols + intJ)
                    If (iType = "1" And intJ > 2) Or (iType = "2" And intJ > 1) Then
                         tmpArray(intJ) = "=IF(請款點數!" & strX & iRow & "="""","""",請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT1 & ")"
                    Else
                         tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                    End If
                Next intJ
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                If iType = "1" Then
                    wksPoint1.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "0.00%"
                ElseIf iType = "2" Then
                    wksPoint1.Range(Pub_NumberToSystem26(xCols + 2) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "0.00%"
                End If
            End If
            strGrp = "" & m_Rst.Fields(strA1n04)
            iRow = iRow + 1
            
JumpToNext:
           m_Rst.MoveNext
       Loop
   Next iCall
   
   xlsPoint1.Sheets(1).Select '選擇工作表
   
   '判斷版本
   If Val(xlsPoint1.Version) < 12 Then
        xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsPoint1.Workbooks.Close
   xlsPoint1.Quit
   Set wksPoint1 = Nothing
   Set xlsPoint1 = Nothing
  
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub

'產生Excel檔案: 3區別個人統計, 4組別個人年移動平均
Private Sub ProcExcelSave3(ByVal iType As String, ByRef m_Rst As ADODB.Recordset, ByVal mDate1 As String, ByVal mDate2 As String)
Dim xlsPoint3 As New Excel.Application
Dim wksPoint3 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String '組別
Dim intPage As Integer '工作表編號
Dim xCols As Integer '行位置
Dim MaxCols As Integer '最大行位置
Dim rowT1 As Integer, rowT2 As Integer  '外商合計、外商總點數的位置
Dim iCall As Integer '表1,2
Dim intJ As Integer
Dim strX As String
Dim strA1n04 As String
Dim tmpArray As Variant

On Error GoTo ErrHnd
   
   Select Case iType
        Case "3" '1區別個人統計
            '檔名：外商請款點數分析3區別個人統計xxxxx~xxxxx
            strExc(1) = strSrvDate(1) & "_外商請款點數3區別個人統計" & Val(mDate1) & "~" & Val(mDate2)
        Case "4" '2組別個人年移動平均
            '檔名：外商請款點數分析4組別個人年移動平均xxxxx~xxxxx
            strExc(1) = strSrvDate(1) & "_外商請款點數4組別個人年移動平均" & Val(mDate1) & "~" & Val(mDate2)
   End Select
   
   strFileName = strExcelPath & strExc(1) & MsgText(43)
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint3.SheetsInNewWorkbook = 2 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint3.Workbooks.add
    xlsPoint3.Visible = False '預設不顯示

   For iCall = 1 To 2
       m_Rst.MoveFirst
       iRow = 1
       xCols = 1
       intPage = iCall
       
       Set wksPoint3 = xlsPoint3.Worksheets(intPage)
       xlsPoint3.Sheets(intPage).Select '選擇工作表
       xlsPoint3.ActiveWindow.DisplayZeros = False '設工作表的零值不顯示
       xlsPoint3.Worksheets(intPage).Name = IIf(iCall = "1", "請款點數", "百分比") '工作表名稱
       '欄位抬頭
       For intJ = 0 To m_Rst.Fields.Count - 1
           strX = Pub_NumberToSystem26(xCols + intJ)
           strExc(3) = "D2" '凍結窗格-位置
           strA1n04 = "組別"
           Select Case intJ
                Case 0: '組別
                    wksPoint3.Range(strX & ":" & strX).ColumnWidth = 16
                    wksPoint3.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case 1: '編號
                    wksPoint3.Range(strX & ":" & strX).ColumnWidth = 7
                    wksPoint3.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case 2: '姓名
                    wksPoint3.Range(strX & ":" & strX).ColumnWidth = 11
                    wksPoint3.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case Else '年月統計量
                    wksPoint3.Range(strX & ":" & strX).ColumnWidth = 10
                    wksPoint3.Range(strX & ":" & strX).HorizontalAlignment = xlRight
           End Select
           '年月標題
           strExc(2) = Replace("" & m_Rst.Fields(intJ).Name, "D", "")
           If Val(strExc(2)) > 0 Then
                If Val(strExc(2)) <= Val(mDate2) Then
                    wksPoint3.Range(strX & iRow).Value = strExc(2)
                    MaxCols = xCols + intJ
                End If
           Else
                'Added by Lydia 2021/06/30 姓名抬頭加註記
                If Check1.Value = 1 And strExc(2) = "姓名" Then
                    strExc(2) = "姓名(*離職)"
                End If
                'end 2021/06/30
                wksPoint3.Range(strX & iRow).Value = strExc(2)
           End If
       Next intJ
       
       wksPoint3.Range(iRow & ":" & iRow).HorizontalAlignment = xlCenter '置中
       wksPoint3.Range(strExc(3)).Select
       xlsPoint3.ActiveWindow.FreezePanes = True '凍結窗格
       wksPoint3.Range("A1").Select
       ReDim tmpArray(1 To MaxCols)
       
       iRow = iRow + 1
       Do While Not m_Rst.EOF
            '不同組別分不同底色
            If strGrp <> "" & m_Rst.Fields(strA1n04) Then
                If iCall = 2 And Trim("" & m_Rst.Fields(strA1n04)) = "9外商總請款點數" Then
                    GoTo JumpToNext
                End If
                '表3-區分英文和非英文組,表4-區分外商和非外商
                'If InStr("5,6,7,8,9", Left("" & m_Rst.Fields(strA1n04), 1)) > 0 Then
                If (iType = "3" And InStr("5,6,7,8,9", Left("" & m_Rst.Fields(strA1n04), 1)) > 0) Or _
                      (iType = "4" And InStr("7,8,9", Left("" & m_Rst.Fields(strA1n04), 1)) > 0) Then
                    iRow = iRow + 1
                    If Trim("" & m_Rst.Fields(strA1n04)) = "7外商合計" Then
                        rowT1 = iRow
                    ElseIf Trim("" & m_Rst.Fields(strA1n04)) = "9外商總請款點數" Then
                        rowT2 = iRow
                    End If
                End If
                If InStr("1,2,3,5,6,7,8,9", Left("" & m_Rst.Fields(strA1n04), 1)) > 0 Or "" & m_Rst.Fields(strA1n04) = "4英文組合計" Then
                    wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = 22 '底色
                End If
                wksPoint3.Range(Pub_NumberToSystem26(xCols) & iRow).Value = "" & m_Rst.Fields(strA1n04)
            End If
            
            If iCall = 1 Then '請款點數
                '因為office2013逐筆輸入過慢,改成陣列輸入
                For intJ = 1 To MaxCols - 1
                    tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                Next intJ
                wksPoint3.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                If iType = "3" Then
                    wksPoint3.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "#,##0.000"
                ElseIf iType = "4" Then
                    wksPoint3.Range(Pub_NumberToSystem26(xCols + 2) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "#,##0.000"
                End If
            ElseIf iCall = 2 Then '百分比
                For intJ = 1 To MaxCols - 1
                    strX = Pub_NumberToSystem26(xCols + intJ)
                    If intJ > 2 Then
                         If strGrp <> "" & m_Rst.Fields(strA1n04) Then
                             'ex: =請款點數!D2/請款點數!D$35 (D35外專總點數)
                            tmpArray(intJ) = "=請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT2
                         Else
                            tmpArray(intJ) = "=IF(請款點數!" & strX & iRow & "="""","""",請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT1 & ")"
                         End If
                    Else
                        tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                    End If
                Next intJ
                wksPoint3.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                wksPoint3.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "0.00%"
            End If
            strGrp = "" & m_Rst.Fields(strA1n04)
            iRow = iRow + 1
            
JumpToNext:
           m_Rst.MoveNext
       Loop
   Next iCall
   
   xlsPoint3.Sheets(1).Select '選擇工作表
   
   '判斷版本
   If Val(xlsPoint3.Version) < 12 Then
        xlsPoint3.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint3.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsPoint3.Workbooks.Close
   xlsPoint3.Quit
   Set wksPoint3 = Nothing
   Set xlsPoint3 = Nothing
  
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub
