VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060332 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "外專暨日專工程師請款點數和OA發文統計表"
   ClientHeight    =   4605
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5685
   Begin VB.CheckBox Check1 
      Caption         =   "是否含離職人員"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3480
      TabIndex        =   29
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
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   3960
      Width           =   4692
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   10
      Left            =   3240
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      Value           =   "10601"
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
      Value           =   "10001"
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   4
      Left            =   3600
      TabIndex        =   21
      Top             =   3600
      Width           =   1305
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   3
      Left            =   2160
      TabIndex        =   20
      Top             =   3600
      Width           =   1305
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   2
      Left            =   3600
      TabIndex        =   19
      Top             =   3360
      Width           =   1305
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   3360
      Width           =   1305
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   15
      Left            =   2880
      TabIndex        =   17
      Top             =   2985
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
      Index           =   13
      Left            =   3240
      TabIndex        =   16
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
      Index           =   9
      Left            =   960
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "1931;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   14
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
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   2985
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "組　　別："
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
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Top             =   2910
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "882;635"
      Value           =   "4"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   2910
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "882;635"
      Value           =   "1"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Value           =   "10001"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   12
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
Attribute VB_Name = "frm060332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2018/12/20 外專工程師請款點數和OA發文統計表
'Memo by Lydia 2018/12/20 使用Form 2.0 (Label和TextBox)
'Memo by Lydia 2020/07/07 更名為「 外專暨日專」
Option Explicit

Dim oText As MSForms.TextBox

Private Sub CmdPrt1_Click()

   If FormCheck = False Then Exit Sub
   
   'Added by Lydia 2021/06/24 另外整理外專工程師員工檔; 因為A6013曾威誌改部門到W20
   'CREATE TABLE R060332_STAFF (STID VARCHAR2(8 CHAR),ST01 VARCHAR2(8 CHAR), ST02 VARCHAR2(12 CHAR),ST03 VARCHAR2(3 CHAR),
   'ST04 VARCHAR2(1 CHAR),ST16 VARCHAR2(3 CHAR),ST16NAME VARCHAR(20 CHAR) , ST70 VARCHAR2(1 CHAR),ST70NAME VARCHAR(20 CHAR));
   strSql = "DELETE FROM R060332_STAFF WHERE STID = '" & strUserNum & "' "
   cnnConnection.Execute strSql
   '在職(含外譯編號)F21,F81；考慮Process4.OTHP其他部門點數，抓全外專
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, ST02,ST03, '1'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 IN ('F21','F81','F23','F22') AND ST01 <> 'F4102' AND ST04='1' "
   cnnConnection.Execute strSql
   '留職停薪；考慮Process4.OTHP其他部門點數，抓全外專
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, ST02,ST03, '1'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF,STAFF_CHANGE,(SELECT SC01 MNO,MAX(SC02) MDATE FROM STAFF_CHANGE WHERE SC04='F21' GROUP BY SC01) VT1 " & _
               "WHERE ST03 IN ('F21','F81','F23','F22') AND ST01 <> 'F4102' AND ST01=SC01(+) AND ST01=MNO(+) AND MNO=SC01(+) AND MDATE=SC02(+) AND '04'=SC03 " & _
               "AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   '離職；考慮Process4.OTHP其他部門點數，抓全外專
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,ST03, '2' AS ST04 , ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 IN ('F21','F81','F23','F22')  AND ST01 <> 'F4102' AND ST04='2' AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   '調職；曾經是工程師，現在非工程師(人員異動排除03離職,08退休,09撤職,10資遣)
    strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F21' AS ST03, '2'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 <>'F21' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F21' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F81' AS ST03, '2'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 <>'F81' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F81' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F22' AS ST03, '2'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 <>'F22' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F22' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO R060332_STAFF (STID,ST01,ST02,ST03,ST04,ST16,ST16NAME,ST70,ST70NAME) " & _
               "SELECT '" & strUserNum & "', ST01, '*'||ST02 AS ST02,'F23' AS ST03, '2'  AS ST04, ST16, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) AS ST16NAME, ST70, DECODE(ST70,'1','3日文機電組','2','3日文化學組',NULL) AS ST70NAME " & _
               "FROM STAFF WHERE ST03 <>'F23' AND ST01 IN (SELECT SC01 FROM STAFF_CHANGE WHERE SC04='F23' AND SC03 NOT IN ('03','08','09','10') GROUP BY SC01) AND ST01 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' )"
   cnnConnection.Execute strSql
   'end 2021/06/24
   Screen.MousePointer = vbHourglass
   CmdPrt1.Enabled = False
   
   If Trim(Combo1.Text) = "" Then '空白=全部
        Call Process1
        Call Process2
        Call Process4("3")
        Call Process4("4")
   Else
        Select Case Left(Combo1.Text, 1)
            Case "1"  '請款點數
                Call Process1
            Case "2"  'OA發文數
                Call Process2
            Case "3"  '每季請款點數分析
                Call Process4(Left(Combo1.Text, 1))
            Case "4"  '每月請款點數分析
                Call Process4(Left(Combo1.Text, 1))
        End Select
   End If
   '執行完不清除條件
   CmdPrt1.Enabled = True
   Screen.MousePointer = vbDefault
   'Modify by Amy 2021/06/22 原:strExcelPath 改中文字顯示
   MsgBox "Excel檔案產生完成！檔案位置：" & strExcelPathN
End Sub

'回傳當季最大月份
Private Function GetSeasonL(ByVal pYYMM As String, Optional ByRef pSNo As String) As String
Dim mYY As String, mMM As String
    
    mYY = PUB_DBYEAR(pYYMM)
    mMM = PUB_DBMONTH(pYYMM)
    
    pSNo = mMM \ 3
    
    '1~2月回傳去年12月
    If Val(pSNo) = 0 Then
        GetSeasonL = (Val(mYY) - 1) & "1201"
    Else
       '回傳當季最大月份
       If mMM Mod 3 = 0 Then
           GetSeasonL = mYY & mMM & "01"
       Else
           GetSeasonL = mYY & Format(Val(pSNo) * 3, "00") & "01"
       End If
    End If
    '季度
    pSNo = Abs(Int(-(Val("" & Mid(GetSeasonL, 5, 2)) / 3)))
End Function


Private Sub Form_Load()

   MoveFormToCenter Me

   LblFM2C(1).Caption = "1." + PUB_GetFCPGrpName("1")
   LblFM2C(2).Caption = "2." + PUB_GetFCPGrpName("2")
   LblFM2C(3).Caption = "3." + PUB_GetFCPGrpName("3")
   LblFM2C(4).Caption = "4." + PUB_GetFCPGrpName("4")
   
   Call txtFM2_Validate(0, False)
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
       MkDir strExcelPath
   End If
   
   txtFM2(3).Text = Left(TransDate(CompDate(1, -1, strSrvDate(1)), 1), 5)
   txtFM2(5).Text = txtFM2(3).Text
   txtFM2(7).Text = Left(TransDate(GetSeasonL(CompDate(1, -1, strSrvDate(1))), 1), 5) '每季~設當季最大月份
   txtFM2(7).Tag = txtFM2(7).Text
   txtFM2(9).Text = txtFM2(3).Text
   
   Combo1.Clear
   Combo1.AddItem "1. 請款點數統計和年移動平均"
   Combo1.AddItem "2. OA發文件數統計和年移動平均"
   Combo1.AddItem "3. 各組每季請款點數分析"
   Combo1.AddItem "4. 各組每月請款點數分析"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060332 = Nothing
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
   
   If txtFM2(0) <> "" And txtFM2(1) <> "" And txtFM2(0) > txtFM2(1) Then
       MsgBox "組別起值不可大於迄值！", vbCritical
       txtFM2(0).SetFocus
       Call txtFM2_GotFocus(0)
       Exit Function
   End If
   
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
       'Added by Lydia 2019/02/27
       If txtFM2(intI) < "10001" Then
             txtFM2(intI).SetFocus
             MsgBox "報表" & inA & "統計年月起值不可小於10001！", , MsgText(5)
             Exit Function
       End If
   Next intI
   
   '每季
   If txtFM2(7).Text > txtFM2(7).Tag Then
        txtFM2(7).SetFocus
        MsgBox "報表3統計年月迄值不可大於" & txtFM2(7).Tag & "！", , MsgText(5)
        Exit Function
   End If
   
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
        Case 0, 1  '1.組別
            If txtFM2(Index).Text = "" Then Exit Sub
            If InStr("1,2,3,4", txtFM2(Index)) = 0 Then
                MsgBox "請輸入1~4 ！", vbCritical
                GoTo EXITSUB
            End If

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

'報表1-請款點數
Private Sub Process1()
Dim stCon As String
Dim strMid01 As String, strMid02 As String
Dim intP As Integer '子查詢序號
Dim yycnt As Integer, mmcnt As Integer
Dim yymm As String
Dim yymm1 As String, mmcnt1 As Integer
Dim rsAD As New ADODB.Recordset
Dim StrSqlB As String, strSQLc As String

   '統計年月
    If txtFM2(2) <> "" Then
        stCon = stCon & " AND A1K02>=" & txtFM2(2) & "00"
    End If
    If txtFM2(3) <> "" Then
        stCon = stCon & " AND A1K02<=" & txtFM2(3) & "31"
    End If
    'Added by Lydia 2021/11/16 查詢印表記錄檔欄位
    ClearQueryLog (Me.Name)
    pub_QL05 = pub_QL05 & ";報表1.請款點數統計"
    pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(2) & "00-" & txtFM2(3) & "31"
    pub_QL05 = pub_QL05 & ";組別:" & txtFM2(0) & "-" & txtFM2(1)
    'end 2021/11/16

   cnnConnection.BeginTrans
       strSql = "DELETE FROM R060332 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' "
       cnnConnection.Execute strSql
       
       strSql = "DELETE FROM R060332_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' "
       cnnConnection.Execute strSql

'---固定表格
'CREATE TABLE R060332 (FORMNAME VARCHAR2(20),ID VARCHAR2(6),TKIND VARCHAR2(1),TNAME VARCHAR2(20 CHAR),A1N04 VARCHAR2(6),DN04 VARCHAR2(1));
'CREATE TABLE R060332_1 (FORMNAME VARCHAR2(20),ID VARCHAR2(6),TKIND VARCHAR2(1),A1N04 VARCHAR2(6),
'YY00 VARCHAR2(4),MM01 NUMBER(10,3),MM02 NUMBER(10,3),MM03 NUMBER(10,3),MM04 NUMBER(10,3),MM05 NUMBER(10,3),MM06 NUMBER(10,3),MM07 NUMBER(10,3),MM08 NUMBER(10,3),MM09 NUMBER(10,3),MM10 NUMBER(10,3),MM11 NUMBER(10,3),MM12 NUMBER(10,3),YYTOTAL NUMBER(13,3));
'ALTER TABLE R060332 ADD PRIMARY KEY (FORMNAME,ID,TKIND,A1N04);
'ALTER TABLE R060332_1 ADD PRIMARY KEY (FORMNAME,ID,TKIND,A1N04,YY00);
        'strSql = "INSERT INTO R060332_1(FORMNAME, ID, TKIND, A1N04, YY00, MM01, MM02, MM03, MM04, MM05, MM06, MM07, MM08, MM09, MM10, MM11, MM12 ) " & _
                    "SELECT FORMNAME, ID, TKIND, A1N04,'" & yycnt & "')"
'-------------------------------------------
        strExc(1) = "請款點數統計"

        '工程師
        'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
        'strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) " & _
                    "SELECT '" & Me.Name & "', '" & strUserNum & "', '1', '" & strExc(1) & "', A1N04 " & _
                    "FROM ACC1K0,ACC1N0,STAFF WHERE NVL(A1K12,0)=0 AND A1K25 IS NULL " & _
                    "AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) AND ST03 IN ('F21','F81') AND A1N04<>'F4102'" & stCon & _
                    " GROUP BY A1N04"
        strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) " & _
                    "SELECT '" & Me.Name & "', '" & strUserNum & "', '1', '" & strExc(1) & "', A1N04 " & _
                    "FROM ACC1K0,ACC1N0,R060332_STAFF WHERE NVL(A1K12,0)=0 AND A1K25 IS NULL " & _
                    "AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') AND A1N04<>'F4102'" & stCon & _
                    " GROUP BY A1N04"
        cnnConnection.Execute strSql, intI
        '組別小計
        For intI = Val(txtFM2(0)) To Val(txtFM2(1))
            strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) VALUES " & _
                       "( '" & Me.Name & "', '" & strUserNum & "', '1', '" & strExc(1) & "', '" & intI & "TOT')"
            cnnConnection.Execute strSql
        Next intI
        '工程師合計
        strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) VALUES " & _
                    "( '" & Me.Name & "', '" & strUserNum & "', '1', '" & strExc(1) & "', 'SUBTOT')"
        cnnConnection.Execute strSql
        '外專總點數
        strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) VALUES " & _
                    "( '" & Me.Name & "', '" & strUserNum & "', '1', '" & strExc(1) & "', 'TOTAL')"
        cnnConnection.Execute strSql
       
        '統計資料-暫存檔
        stCon = "SELECT B.A1N04, ST02, ST01, ST16, DN04 "
        'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
        'StrSqlB = "FROM R060332 B,STAFF "
        'strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '1' AND B.A1N04=ST01(+)"
        StrSqlB = "FROM R060332 B,R060332_STAFF "
        strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '1' AND B.A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) "
        strMid01 = " AND " '抓所有欄位加起來>0
        intP = 1  '子查詢序號
        For yycnt = Val(Left(txtFM2(2), 3)) To Val(Left(txtFM2(3), 3))
           For mmcnt = 1 To 12
              If mmcnt = 1 Then
                 strSql = "INSERT INTO R060332_1(FORMNAME, ID, TKIND, A1N04, YY00 ) " & _
                             "SELECT FORMNAME, ID, TKIND, A1N04,'" & yycnt & "' FROM R060332 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' "
                 cnnConnection.Execute strSql, intI
                 StrSqlB = StrSqlB & ", (SELECT * FROM R060332_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  AND YY00='" & yycnt & "') X" & intP
                 strSQLc = strSQLc & " AND B.A1N04=X" & intP & ".A1N04"
              End If
              yymm = yycnt * 100 + mmcnt
              stCon = stCon & ", X" & intP & ".MM" & Format(mmcnt, "00") & " AS D" & yymm 'ex: D10012 (100年12月)
              strMid01 = strMid01 & "NVL(D" & yymm & ",0)+"
              'Modified by Lydia 2022/12/16 加判斷起始年月yymm >= Val(txtFM2(2))
              If yymm >= Val(txtFM2(2)) And yymm <= Val(txtFM2(3)) Then  '超過期限不抓資料
                    '工程師總點數扣除翻譯費201點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF, 注意A1K21為建檔人員,A1N04=點數分配之承辦人
                    'strSql = "select A.a1n04,(TOT-NVL(TFEE,0)) BAL,st16 from " & _
                             "(select a1n04,sum(a1n05) TOT,st16 from acc1k0,acc1n0,STAFF where a1k01 in " & _
                             " (select distinct a1k01 from acc1k0,STAFF,caseprogress where a1k02>='" & yymm & "00' and a1k02<='" & yymm & "31'  and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F2%') " & _
                             "     AND A1K01=A1N01(+) AND '2'=A1N02(+) and a1n04=st01(+) and st03 IN ('F21','F81') and a1n04<>'F4102'" & _
                             "   group by a1n04,st16) A," & _
                             "(select a1n04,sum(a1n05) TFEE from acc1k0,acc1n0,STAFF,caseprogress where a1k01 in " & _
                             " (select distinct a1k01 from acc1k0,STAFF,caseprogress where a1k02>='" & yymm & "00' and a1k02<='" & yymm & "31' and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F2%') " & _
                             "     AND A1K01=A1N01(+) AND '2'=A1N02(+) and a1n04=st01(+) and st03 IN ('F21','F81') and a1n06 is null and a1n03 is not null and a1n04<>'F4102' and a1n03=cp09(+) and cp01||cp10 in ('P201','FCP201') " & _
                             "   group by a1n04) B " & _
                             " WHERE A.a1n04=B.a1n04(+) order by A.a1n04"
                    strSql = "select A.a1n04,(TOT-NVL(TFEE,0)) BAL,st16 from " & _
                             "(select a1n04,sum(a1n05) TOT,st16 from acc1k0,acc1n0,R060332_staff where a1k01 in " & _
                             " (select distinct a1k01 from acc1k0,STAFF,caseprogress where a1k02>='" & yymm & "00' and a1k02<='" & yymm & "31'  and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) AND a1k01=cp60(+) and cp12 like 'F2%') " & _
                             "     AND A1K01=A1N01(+) AND '2'=A1N02(+) and a1n04=st01(+) and '" & strUserNum & "'=STID(+)  and st03 IN ('F21','F81') and a1n04<>'F4102'" & _
                             "   group by a1n04,st16) A," & _
                             "(select a1n04,sum(a1n05) TFEE from acc1k0,acc1n0,R060332_staff,caseprogress where a1k01 in " & _
                             " (select distinct a1k01 from acc1k0,Staff,caseprogress where a1k02>='" & yymm & "00' and a1k02<='" & yymm & "31' and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F2%') " & _
                             "     AND A1K01=A1N01(+) AND '2'=A1N02(+) and a1n04=st01(+) and '" & strUserNum & "'=STID(+) and st03 IN ('F21','F81') and a1n06 is null and a1n03 is not null and a1n04<>'F4102' and a1n03=cp09(+) and cp01||cp10 in ('P201','FCP201') " & _
                             "   group by a1n04) B " & _
                             " WHERE A.a1n04=B.a1n04(+) order by A.a1n04"
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                       RsTemp.MoveFirst
                       Do While Not RsTemp.EOF
                            strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  " & _
                                        "AND A1N04='" & RsTemp.Fields(0) & "' AND YY00='" & yycnt & "' "
                            cnnConnection.Execute strSql, intI
                            '同時加入各組別小計
                            strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  " & _
                                        "AND A1N04='" & RsTemp.Fields("ST16") & "TOT' AND YY00='" & yycnt & "' "
                            cnnConnection.Execute strSql, intI
                            RsTemp.MoveNext
                       Loop
                    End If
                    '工程師合計
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") " & _
                                "from R060332_1,STAFF where A1N04=ST01(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' AND YY00='" & yycnt & "') " & _
                                "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  AND A1N04='SUBTOT' AND YY00='" & yycnt & "' "
                    strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") " & _
                                "from R060332_1, R060332_STAFF where A1N04=ST01(+) and ST01 is not null and '" & strUserNum & "'=STID(+) AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' AND YY00='" & yycnt & "') " & _
                                "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  AND A1N04='SUBTOT' AND YY00='" & yycnt & "' "
                    cnnConnection.Execute strSql, intI
                    '外專總點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'Memo by Lydia 2021/06/24 注意A1K21為建檔人員
                    strSql = "select nvl(sum(a1k11-nvl(a1k06,0)-nvl(a1k09,0)+nvl(a1k36,0))/1000,0) 總點數 from acc1k0 where a1k01 in " & _
                             "(select distinct a1k01 from acc1k0,staff,caseprogress where a1k02>='" & yymm & "00' and a1k02<='" & yymm & "31' and nvl(a1k12,0)=0 and a1k25 is null and a1k21=st01(+) and a1k01=cp60(+) and cp12 like 'F2%')"
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                        strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(0) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' AND A1N04='TOTAL' AND YY00='" & yycnt & "' "
                        cnnConnection.Execute strSql
                    End If
              End If
           Next mmcnt
           intP = intP + 1
        Next yycnt
        
        '名單刪除離職人員但組別小計及工程師合計是包含所有人的數字,故工作檔增加DN04,離職人員才有值
        'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF ;  增加「是否含離職人員」的判斷
        'strSql = "UPDATE R060332 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  AND A1N04 NOT IN " & _
                 "(SELECT ST01 FROM STAFF,STAFF_CHANGE,(SELECT ST01 NO,MAX(SC02) MAXDATE FROM STAFF,STAFF_CHANGE WHERE ST03='F21' AND ST01=SC01(+) GROUP BY ST01) " & _
                 "WHERE ST03='F21' AND ST01<>'F4102' AND ST01=NO(+) AND NO=SC01(+) AND MAXDATE=SC02(+) AND '04'=SC03(+) AND (ST04='1' OR SC02 IS NOT NULL))"
        If Check1.Value = 0 Then
            strSql = "UPDATE R060332 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  " & _
                        "AND A1N04 IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' AND ST04='2' AND ST03='F21' AND ST01<>'F4102' ) "
            cnnConnection.Execute strSql, intI
        End If
        
        
        '剔除條件外的工程師組別資料
        'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
        'strSql = "DELETE FROM R060332 B WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '1'  AND INSTR(B.A1N04,'TOT')=0 " & _
                     "AND B.FORMNAME||B.ID||B.TKIND||B.A1N04 NOT IN (SELECT A.FORMNAME||A.ID||A.TKIND||A.A1N04 FROM R060332 A,STAFF WHERE A.A1N04=ST01(+) AND ST16>='" & txtFM2(0) & "' AND ST16<='" & txtFM2(1) & "') "
        strSql = "DELETE FROM R060332 B WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '1'  AND INSTR(B.A1N04,'TOT')=0 " & _
                     "AND B.FORMNAME||B.ID||B.TKIND||B.A1N04 NOT IN (SELECT A.FORMNAME||A.ID||A.TKIND||A.A1N04 FROM R060332 A,R060332_STAFF WHERE A.A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST16>='" & txtFM2(0) & "' AND ST16<='" & txtFM2(1) & "') "
        cnnConnection.Execute strSql, intI
        
   cnnConnection.CommitTrans
'-----------------------------------------
   '產生Excel檔: 外專工程師10001~107xx每月請款點數
       '在職工程師
       '因檔案內有組別小計,工程師合計及外專總點數資料,故只抓讀得到員工檔且在職的資料
       strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
       strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
       strExc(5) = "SELECT DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) 組別,A1N04 編號,ST02 姓名," & _
                        strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
       strExc(5) = strExc(5) & strMid01 & ">0"
       '各組小計
       strExc(6) = "SELECT DECODE(SUBSTR(A1N04,1,1),'1','1電子電機','2','2化學','3','3日文','4','4機械設計',A1N04) 組別,' ' 編號,'小計' 姓名," & _
                         strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04>'1' AND A1N04<'5' AND INSTR(A1N04,'TOT')=2"
       '工程師合計
       strExc(7) = "SELECT '5工程師合計' 組別,' ' 編號,'工程師合計' 姓名," & _
                         strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04='SUBTOT'"
       '外專總點數
       strExc(8) = "SELECT '6外專總點數' 組別,'' 編號,'外專總點數' 姓名," & _
                         strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04='TOTAL'"
       
       '先產生非年移動平均數.xls
JumpToRe01:
        If rsAD.State = adStateOpen Then
           rsAD.Close
        End If
        rsAD.CursorLocation = adUseClient
        rsAD.Open strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7) & " UNION " & strExc(8) & " ORDER BY 1,2 ", cnnConnection, adOpenStatic, adLockReadOnly
        If rsAD.RecordCount > 0 Then
             InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
             Call ProcExcelSave1("1", rsAD)
        'Added by Lydia 2021/11/16
        Else
             InsertQueryLog (0)
        'end 2021/11/16
        End If
'-----------------------------------------

   '改年移動平均數：10012為10001~10012平均,10101為10002~10101平均
   cnnConnection.BeginTrans
   For yycnt = Val(Left(txtFM2(3), 3)) To Val(Left(txtFM2(2), 3)) Step -1
      For mmcnt = 12 To 1 Step -1
         yymm = yycnt * 100 + mmcnt
         If Val(yymm) >= Val(Left(txtFM2(2), 3) & "12") And Val(yymm) <= Val(txtFM2(3)) Then
            For mmcnt1 = 1 To 11 '往前推11個月
               yymm1 = Val(yymm) - mmcnt1
               If (Val(Right(yymm1, 2)) <= 0 Or Val(Right(yymm1, 2)) >= 90) Then '跨年
                  yymm1 = (yycnt - 1) * 100 + mmcnt - mmcnt1 + 12
               End If
               strSql = "UPDATE R060332_1 A SET A.MM" & Right(yymm, 2) & "=NVL(A.MM" & Right(yymm, 2) & ",0) " & _
                            "+(SELECT NVL(B.MM" & Right(yymm1, 2) & ",0) FROM R060332_1 B WHERE A.FORMNAME=B.FORMNAME AND A.ID=B.ID AND A.TKIND=B.TKIND AND A.A1N04=B.A1N04 AND B.YY00='" & Mid(yymm1, 1, Len(yymm1) - 2) & "') " & _
                            "WHERE A.FORMNAME = '" & Me.Name & "' AND A.ID = '" & strUserNum & "' AND A.TKIND = '1' AND A.YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
               cnnConnection.Execute strSql, intI
            Next mmcnt1
            '平均
            strSql = "UPDATE R060332_1 SET MM" & Right(yymm, 2) & "=MM" & Right(yymm, 2) & "/12 " & _
                        "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1' AND YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
            cnnConnection.Execute strSql, intI
         End If
      Next mmcnt
   Next yycnt
 
   cnnConnection.CommitTrans

   '產生Excel檔: 外專工程師10012~107xx每月請款點數年移動平均
       '在職工程師
       '因檔案內有組別小計,工程師合計及外專總點數資料,故只抓讀得到員工檔且在職的資料
       strExc(5) = "SELECT DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) 組別,A1N04 編號,ST02 姓名," & _
                        Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(2), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
       strExc(5) = strExc(5) & strMid01 & ">0"
       '各組小計
       strExc(6) = "SELECT DECODE(SUBSTR(A1N04,1,1),'1','1電子電機','2','2化學','3','3日文','4','4機械設計',A1N04) 組別,' ' 編號,'小計' 姓名," & _
                        Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(2), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04>'1' AND A1N04<'5' AND INSTR(A1N04,'TOT')=2"
       '工程師合計
       strExc(7) = "SELECT '5工程師合計' 組別,' ' 編號,'工程師合計' 姓名," & _
                         Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(2), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04='SUBTOT'"
       '外專總點數
       strExc(8) = "SELECT '6外專總點數' 組別,'' 編號,'外專總點數' 姓名," & _
                         Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(2), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04='TOTAL'"
        
       '---產生年移動平均.xls
JumpToRe02:
        If rsAD.State = adStateOpen Then
           rsAD.Close
        End If
        rsAD.CursorLocation = adUseClient
        rsAD.Open strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7) & " UNION " & strExc(8) & " ORDER BY 1,2 ", cnnConnection, adOpenStatic, adLockReadOnly
        If rsAD.RecordCount > 0 Then
             Call ProcExcelSave1("2", rsAD)
        End If
        
        Set rsAD = Nothing

End Sub

'產生Excel檔案-請款點數
Private Sub ProcExcelSave1(ByVal iType As String, ByRef m_Rst As ADODB.Recordset)
'iType: 1.請款點數, 2.年移動平均
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String, strA1n04 As String '組別,工程師編號
Dim intPage As Integer '工作表編號
Dim xCols As Integer '行位置
Dim MaxCols As Integer '最大行位置
Dim rowT1 As Integer, rowT2 As Integer '工程師合計、外專總點數的位置
Dim iCall As Integer '表1,2
Dim intJ As Integer
Dim strX As String
Dim tmpArray As Variant 'Added by Lydia 2019/02/15

On Error GoTo ErrHnd


   '1.檔名：外專工程師10001~107xx每月請款點數 ; 2.檔名：外專工程師10012~107xx每月請款點數年移動平均
   If iType = "1" Then
       'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
       strExc(1) = strSrvDate(1) & "_外專暨日專工程師" & Val(txtFM2(2)) & "~" & Val(txtFM2(3)) & "每月請款點數"
   Else
       'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
       strExc(1) = strSrvDate(1) & "_外專暨日專工程師" & Val(Left(txtFM2(2), 3)) & "12" & "~" & Val(txtFM2(3)) & "每月請款點數年移動平均"
   End If
   strFileName = strExcelPath & strExc(1) & MsgText(43)
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint1.SheetsInNewWorkbook = 3 'Added by Lydia 2019/02/15 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint1.Workbooks.add
    xlsPoint1.Visible = False '預設不顯示

   For iCall = 1 To 2
       m_Rst.MoveFirst
       iRow = 1
       xCols = 1
       intPage = iCall
       
       Set wksPoint1 = xlsPoint1.Worksheets(intPage)
       xlsPoint1.Sheets(intPage).Select '選擇工作表
       xlsPoint1.ActiveWindow.DisplayZeros = False 'Added by Lydia 2019/02/15 設工作表的零值不顯示
       xlsPoint1.Worksheets(intPage).Name = IIf(iCall = "1", "請款點數", "百分比") '工作表名稱
       '欄位抬頭
       For intJ = 0 To m_Rst.Fields.Count - 1
           strX = Pub_NumberToSystem26(xCols + intJ)
           Select Case intJ
                Case 0: '組別
                    wksPoint1.Range(strX & ":" & strX).ColumnWidth = 13
                    wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case 1: '編號
                    wksPoint1.Range(strX & ":" & strX).ColumnWidth = 7
                    wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case 2: '姓名
                    wksPoint1.Range(strX & ":" & strX).ColumnWidth = 13
                    wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                Case Else '年月統計量
                    wksPoint1.Range(strX & ":" & strX).ColumnWidth = 11
                    wksPoint1.Range(strX & ":" & strX).HorizontalAlignment = xlRight
           End Select
           strExc(2) = Replace("" & m_Rst.Fields(intJ).Name, "D", "")
           If Val(strExc(2)) > 0 Then
                 'Modified by Lydia 2022/12/16 加判斷起始年月Val(strExc(2)) >= Val(txtFM2(2))
                If Val(strExc(2)) >= Val(txtFM2(2)) And Val(strExc(2)) <= Val(txtFM2(3)) Then
                    wksPoint1.Range(strX & iRow).Value = strExc(2)
                    MaxCols = xCols + intJ
                End If
           Else
                'Added by Lydia 2021/06/24 姓名抬頭加註記
                If Check1.Value = 1 And strExc(2) = "姓名" Then
                    strExc(2) = "姓名(*離職)"
                End If
                'end 2021/06/24
                wksPoint1.Range(strX & iRow).Value = strExc(2)
           End If
       Next intJ
       wksPoint1.Range(iRow & ":" & iRow).HorizontalAlignment = xlCenter '置中
       'Modified by Lydia 2021/11/29  Widen程式產生的excel會自動隱藏凍結窗格;雖然經理已刪除個人登錄檔,試著修改程式碼
       'wksPoint1.Range("D2").Select
       'xlsPoint1.ActiveWindow.FreezePanes = True '凍結窗格
       xlsPoint1.ActiveWindow.FreezePanes = False
       xlsPoint1.ActiveWindow.SplitColumn = 3
       xlsPoint1.ActiveWindow.SplitRow = 1
       xlsPoint1.ActiveWindow.FreezePanes = True
       'end 2021/11/29
       wksPoint1.Range("A1").Select
       'Added by Lydia 2020/07/07 統計年月不到一年會出錯
       If MaxCols = 0 Then
            GoTo JumpToExcept
       Else
       'end 2020/07/07
            ReDim tmpArray(1 To MaxCols) 'Added by Lydia 2019/02/15
       End If 'Added by Lydia 2020/07/07
       
       iRow = iRow + 1
       Do While Not m_Rst.EOF
            '不同組別分不同底色
            If strGrp <> "" & m_Rst.Fields("組別") Then
                If Val(Left("" & m_Rst.Fields("組別"), 1)) > 4 Then '合計,多跳一行
                    iRow = iRow + 1
                    If iCall = 1 Then
                       If Val(Left("" & m_Rst.Fields("組別"), 1)) = 5 Then
                           rowT1 = iRow
                       ElseIf Val(Left("" & m_Rst.Fields("組別"), 1)) = 6 Then
                           rowT2 = iRow
                       End If
                    ElseIf iCall = 2 Then
                       If Val(Left("" & m_Rst.Fields("組別"), 1)) = 6 Then
                            GoTo JumpToNext
                       End If
                    End If
                End If
                wksPoint1.Range(iRow & ":" & iRow).Interior.ColorIndex = 22 '底色
                wksPoint1.Range(Pub_NumberToSystem26(xCols) & iRow).Value = "" & m_Rst.Fields("組別")
            End If
            If iCall = 1 Then '請款點數
                'Modified by Lydia 2019/02/15 因為office2013逐筆輸入過慢,改成陣列輸入
'                For intJ = 1 To MaxCols - 1
'                    strX = Pub_NumberToSystem26(xCols + intJ)
'                    wksPoint1.Range(strX & iRow).Value = "" & m_Rst.Fields(intJ)
'                    If intJ >= 3 Then wksPoint1.Range(strX & iRow).NumberFormat = "#,##0.000"
'                Next intJ
                For intJ = 1 To MaxCols - 1
                    tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                Next intJ
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "#,##0.000"
                'end 2019/02/15
            ElseIf iCall = 2 Then '百分比
                
                'Modified by Lydia 2019/02/15 因為office2013逐筆輸入過慢,改成陣列輸入
'                For intJ = 1 To MaxCols - 1
'                    strX = Pub_NumberToSystem26(xCols + intJ)
'                    If intJ > 2 Then
'                         If strGrp <> "" & m_Rst.Fields("組別") Then
'                             'ex: =請款點數!D2/請款點數!D$53 (D53外專總點數)
'                            wksPoint1.Range(strX & iRow).Value = "=請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT2
'                         Else
'                            wksPoint1.Range(strX & iRow).Value = "=IF(請款點數!" & strX & iRow & "="""","""",請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT1 & ")"
'                         End If
'                         wksPoint1.Range(strX & iRow).NumberFormat = "0.00%"
'                    Else
'                        wksPoint1.Range(strX & iRow).Value = "" & m_Rst.Fields(intJ)
'                    End If
'                Next intJ
                For intJ = 1 To MaxCols - 1
                    strX = Pub_NumberToSystem26(xCols + intJ)
                    If intJ > 2 Then
                         If strGrp <> "" & m_Rst.Fields("組別") Then
                             'ex: =請款點數!D2/請款點數!D$53 (D53外專總點數)
                            tmpArray(intJ) = "=請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT2
                         Else
                            tmpArray(intJ) = "=IF(請款點數!" & strX & iRow & "="""","""",請款點數!" & strX & iRow & "/請款點數!" & strX & "$" & rowT1 & ")"
                         End If
                    Else
                        tmpArray(intJ) = "" & m_Rst.Fields(intJ)
                    End If
                Next intJ
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 1) & iRow).Value = tmpArray
                wksPoint1.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols) & iRow).NumberFormat = "0.00%"
                'end 2019/02/15
            End If
            strGrp = "" & m_Rst.Fields("組別")
            strA1n04 = "" & m_Rst.Fields("編號")
            iRow = iRow + 1
JumpToNext:

           m_Rst.MoveNext
       Loop
   Next iCall
   
JumpToExcept: 'Added by Lydia 2020/07/07
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

'報表2-OA發文數
Private Sub Process2()
Dim stCon As String
Dim strMid01 As String, strMid02 As String
Dim intP As Integer '子查詢序號
Dim yycnt As Integer, mmcnt As Integer
Dim yymm As String
Dim yymm1 As String, mmcnt1 As Integer
Dim rsAD As New ADODB.Recordset
Dim StrSqlB As String, strSQLc As String
Dim iRound As Integer '分兩次抓資料
   
   strExc(9) = ""
   strExc(10) = ""

   cnnConnection.BeginTrans
       strSql = "DELETE FROM R060332 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND LIKE '2%' "
       cnnConnection.Execute strSql
       
       strSql = "DELETE FROM R060332_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND LIKE '2%' "
       cnnConnection.Execute strSql
'-----------------分兩次抓資料
       strExc(1) = "OA發文件數"
       For iRound = 21 To 22
            '統計年月
            stCon = ""
            If txtFM2(4) <> "" Then
                stCon = stCon & " AND CP27>=" & Val(txtFM2(4)) + 191100 & "00"
            End If
            If txtFM2(5) <> "" Then
                stCon = stCon & " AND CP27<=" & Val(txtFM2(5)) + 191100 & "31"
            End If
            'Added by Lydia 2021/11/16 查詢印表記錄檔欄位
            If iRound = 21 Then
               ClearQueryLog (Me.Name)
               pub_QL05 = pub_QL05 & ";報表2.OA發文件數統計"
               pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(4) & "00-" & txtFM2(5) & "31"
               pub_QL05 = pub_QL05 & ";組別:" & txtFM2(0) & "-" & txtFM2(1)
            End If
            'end 2021/11/16
                
           '工程師
           'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
           strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) " & _
                       "SELECT '" & Me.Name & "', '" & strUserNum & "', '" & iRound & "', '" & strExc(1) & "', CP14 FROM CASEPROGRESS,R060332_STAFF "
           If iRound = 21 Then
                '工作表一：OA審查意見通知函1202+申復205+核駁1002,1006+再審申請107發文數
                strSql = strSql & "WHERE CP10 IN ('1202','1002','1006','205','107') "
           Else
                '工作表二：OA審查意見通知函1202+核駁1002,1006發文數
                strSql = strSql & "WHERE CP10 IN ('1202','1002','1006') "
           End If
           'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
           'strSql = strSql & stCon & " AND CP14=ST01(+) AND ST03 IN ('F21','F81') GROUP BY CP14"
           strSql = strSql & stCon & " AND CP14=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') GROUP BY CP14"
           cnnConnection.Execute strSql, intI
           '組別小計
           For intI = Val(txtFM2(0)) To Val(txtFM2(1))
                strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) VALUES " & _
                           "( '" & Me.Name & "', '" & strUserNum & "', '" & iRound & "', '" & strExc(1) & "', '" & intI & "TOT')"
                cnnConnection.Execute strSql
           Next intI
            '工程師合計
            strSql = "INSERT INTO R060332 (FORMNAME,ID,TKIND,TNAME,A1N04) VALUES " & _
                        "( '" & Me.Name & "', '" & strUserNum & "', '" & iRound & "', '" & strExc(1) & "', 'SUBTOT')"
            cnnConnection.Execute strSql
        
            '統計資料-暫存檔
            stCon = "SELECT B.A1N04, ST02, ST01, ST16, DN04 "
            'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
            'StrSqlB = "FROM R060332 B,STAFF"
            'strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & iRound & "' AND B.A1N04=ST01(+)"
            StrSqlB = "FROM R060332 B,R060332_STAFF "
            strSQLc = "WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & iRound & "' AND B.A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) "
            'end 2021/06/24
            strMid01 = " AND " '抓所有欄位加起來>0
            intP = 1  '子查詢序號
            For yycnt = Val(Left(txtFM2(4), 3)) To Val(Left(txtFM2(5), 3))
               For mmcnt = 1 To 12
                  If mmcnt = 1 Then
                     strSql = "INSERT INTO R060332_1(FORMNAME, ID, TKIND, A1N04, YY00 ) " & _
                                 "SELECT FORMNAME, ID, TKIND, A1N04,'" & yycnt & "' FROM R060332 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' "
                     cnnConnection.Execute strSql, intI
                     StrSqlB = StrSqlB & ", (SELECT * FROM R060332_1 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "'  AND YY00='" & yycnt & "') X" & intP
                     strSQLc = strSQLc & " AND B.A1N04=X" & intP & ".A1N04"
                  End If
                  yymm = yycnt * 100 + mmcnt
                  stCon = stCon & ", X" & intP & ".MM" & Format(mmcnt, "00") & " AS D" & yymm 'ex: D10012 (100年12月)
                  strMid01 = strMid01 & "NVL(D" & yymm & ",0)+"
                  'Modified by Lydia 2022/12/16 加判斷起始年月yymm >= Val(txtFM2(4))
                  If yymm >= Val(txtFM2(4)) And yymm <= Val(txtFM2(5)) Then  '超過期限不抓資料
                        If iRound = 21 Then
                            'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                            'strSql = "SELECT CP14,COUNT(*) CNT,st16 FROM CASEPROGRESS,STAFF " & _
                                       " WHERE CP27>='" & yymm + 191100 & "00' AND CP27<='" & yymm + 191100 & "31'  AND CP10 IN ('1202','1002','1006','205','107') " & _
                                       " AND CP14=ST01(+) AND ST03 IN ('F21','F81') GROUP BY CP14,st16 ORDER BY CP14,st16"
                            strSql = "SELECT CP14,COUNT(*) CNT,st16 FROM CASEPROGRESS,R060332_STAFF " & _
                                       " WHERE CP27>='" & yymm + 191100 & "00' AND CP27<='" & yymm + 191100 & "31'  AND CP10 IN ('1202','1002','1006','205','107') " & _
                                       " AND CP14=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') GROUP BY CP14,st16 ORDER BY CP14,st16"
                        Else
                            'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                            'strSql = "SELECT CP14,COUNT(*) CNT,st16 FROM CASEPROGRESS,STAFF " & _
                                       " WHERE CP27>='" & yymm + 191100 & "00' AND CP27<='" & yymm + 191100 & "31'  AND CP10 IN ('1202','1002','1006') " & _
                                       " AND CP14=ST01(+) AND ST03 IN ('F21','F81') GROUP BY CP14,st16 ORDER BY CP14,st16"
                            strSql = "SELECT CP14,COUNT(*) CNT,st16 FROM CASEPROGRESS,R060332_STAFF " & _
                                       " WHERE CP27>='" & yymm + 191100 & "00' AND CP27<='" & yymm + 191100 & "31'  AND CP10 IN ('1202','1002','1006') " & _
                                       " AND CP14=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') GROUP BY CP14,st16 ORDER BY CP14,st16"
                        End If
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           RsTemp.MoveFirst
                           Do While Not RsTemp.EOF
                                strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' " & _
                                            "AND A1N04='" & RsTemp.Fields(0) & "' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                '同時加入各組別小計
                                strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=nvl(MM" & Format(mmcnt, "00") & ",0)+" & RsTemp.Fields(1) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' " & _
                                            "AND A1N04='" & RsTemp.Fields("ST16") & "TOT' AND YY00='" & yycnt & "' "
                                cnnConnection.Execute strSql, intI
                                RsTemp.MoveNext
                           Loop
                        End If
                        '工程師合計
                        'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                        'strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060332_1,STAFF where A1N04=ST01(+) and ST01 is not null AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' AND YY00='" & yycnt & "') " & _
                                    "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "'  AND A1N04='SUBTOT' AND YY00='" & yycnt & "' "
                        strSql = "UPDATE R060332_1 SET MM" & Format(mmcnt, "00") & "=(select sum(MM" & Format(mmcnt, "00") & ") from R060332_1,R060332_STAFF where A1N04=ST01(+) and ST01 is not null and '" & strUserNum & "'=STID(+) AND FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' AND YY00='" & yycnt & "') " & _
                                    "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "'  AND A1N04='SUBTOT' AND YY00='" & yycnt & "' "
                        cnnConnection.Execute strSql, intI
                  End If
               Next mmcnt
               intP = intP + 1
            Next yycnt
         
         '名單刪除離職人員但組別小計及工程師合計是包含所有人的數字,故工作檔增加DN04,離職人員才有值
         'modify by sonia 2021/1/26 再排除F4104及F4105
         'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF ;  增加「是否含離職人員」的判斷
         'strSql = "UPDATE R060332 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' AND A1N04 NOT IN " & _
                  "(SELECT ST01 FROM STAFF,STAFF_CHANGE,(SELECT ST01 NO,MAX(SC02) MAXDATE FROM STAFF,STAFF_CHANGE WHERE ST03='F21' AND ST01=SC01(+) GROUP BY ST01) " & _
                  "WHERE ST03='F21' AND ST01<>'F4102' AND ST01<>'F4104' AND ST01<>'F4105' AND ST01=NO(+) AND NO=SC01(+) AND MAXDATE=SC02(+) AND '04'=SC03(+) AND (ST04='1' OR SC02 IS NOT NULL))"
         'cnnConnection.Execute strSql, intI
         If Check1.Value = 0 Then
             strSql = "UPDATE R060332 SET DN04='2' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '1'  " & _
                         "AND A1N04 IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' AND ST04='2' AND ST03='F21' AND ST01<>'F4102' AND ST01<>'F4104' AND ST01<>'F4105' ) "
             cnnConnection.Execute strSql, intI
         End If
         'end 2021/06/24
         
         '剔除條件外的工程師組別資料
         'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
         'strSql = "DELETE FROM R060332 B WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & iRound & "' AND INSTR(B.A1N04,'TOT')=0 " & _
                      "AND B.FORMNAME||B.ID||B.TKIND||B.A1N04 NOT IN (SELECT A.FORMNAME||A.ID||A.TKIND||A.A1N04 FROM R060332 A,STAFF WHERE A.A1N04=ST01(+) AND ST16>='" & txtFM2(0) & "' AND ST16<='" & txtFM2(1) & "') "
         strSql = "DELETE FROM R060332 B WHERE B.FORMNAME = '" & Me.Name & "' AND B.ID = '" & strUserNum & "' AND B.TKIND = '" & iRound & "' AND INSTR(B.A1N04,'TOT')=0 " & _
                      "AND B.FORMNAME||B.ID||B.TKIND||B.A1N04 NOT IN (SELECT A.FORMNAME||A.ID||A.TKIND||A.A1N04 FROM R060332 A,R060332_STAFF WHERE A.A1N04=ST01(+) AND '" & strUserNum & "'=STID(+) AND ST16>='" & txtFM2(0) & "' AND ST16<='" & txtFM2(1) & "') "
         cnnConnection.Execute strSql, intI
       
         '記錄-組合語法
         '在職工程師
         '因檔案內有組別小計,工程師合計及外專總點數資料,故只抓讀得到員工檔且在職的資料
         strMid01 = Mid(strMid01, 1, Len(strMid01) - 1) '抓所有欄位加起來>0
         strMid02 = Replace(Replace(Replace(Mid(strMid01, 5), "NVL(", ""), ",0)", ""), "+", ",") '抓所有欄位的別名, ex:D10001,D10002
         strExc(5) = "SELECT '" & iRound & "' as ORD1, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) 組別,A1N04 編號,ST02 姓名," & _
                          strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
         strExc(5) = strExc(5) & strMid01 & ">0"
         '各組小計
         strExc(6) = "SELECT '" & iRound & "' as ORD1, DECODE(SUBSTR(A1N04,1,1),'1','1電子電機','2','2化學','3','3日文','4','4機械設計',A1N04) 組別,' ' 編號,'小計' 姓名," & _
                           strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                          " WHERE A1N04>'1' AND A1N04<'5' AND INSTR(A1N04,'TOT')=2"
         '工程師合計
         strExc(7) = "SELECT '" & iRound & "' as ORD1, '5工程師合計' 組別,' ' 編號,'工程師合計' 姓名," & _
                           strMid02 & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                          " WHERE A1N04='SUBTOT'"
         If iRound = 21 Then
              strExc(9) = strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7)
         Else
              strExc(10) = strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7)
         End If
       Next iRound
'-----------------分兩次抓資料
   cnnConnection.CommitTrans
'-----------------------------------------
   '產生Excel檔: 外專工程師10001~107xx每月OA發文數
    '先產生非年移動平均數.xls
JumpToRe01:
    If rsAD.State = adStateOpen Then
       rsAD.Close
    End If
    rsAD.CursorLocation = adUseClient
    rsAD.Open strExc(9) & " UNION " & strExc(10) & " ORDER BY ord1,1,2 ", cnnConnection, adOpenStatic, adLockReadOnly
    If rsAD.RecordCount > 0 Then
         InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
         Call ProcExcelSave2("1", rsAD)
    'Added by Lydia 2021/11/16
    Else
         InsertQueryLog (0)
    'end 2021/11/16
    End If

'-----------------------------------------
   '改年移動平均數：10012為10001~10012平均,10101為10002~10101平均
   strExc(9) = ""
   strExc(10) = ""
   cnnConnection.BeginTrans
'-----------------分兩次抓資料
   For iRound = 21 To 22
        For yycnt = Val(Left(txtFM2(5), 3)) To Val(Left(txtFM2(4), 3)) Step -1
           For mmcnt = 12 To 1 Step -1
              yymm = yycnt * 100 + mmcnt
              If Val(yymm) >= Val(Left(txtFM2(4), 3) & "12") And Val(yymm) <= Val(txtFM2(5)) Then
                 For mmcnt1 = 1 To 11 '往前推11個月
                    yymm1 = Val(yymm) - mmcnt1
                    If (Val(Right(yymm1, 2)) <= 0 Or Val(Right(yymm1, 2)) >= 90) Then '跨年
                       yymm1 = (yycnt - 1) * 100 + mmcnt - mmcnt1 + 12
                    End If
                    strSql = "UPDATE R060332_1 A SET A.MM" & Right(yymm, 2) & "=NVL(A.MM" & Right(yymm, 2) & ",0) " & _
                                 "+(SELECT NVL(B.MM" & Right(yymm1, 2) & ",0) FROM R060332_1 B WHERE A.FORMNAME=B.FORMNAME AND A.ID=B.ID AND A.TKIND=B.TKIND AND A.A1N04=B.A1N04 AND B.YY00='" & Mid(yymm1, 1, Len(yymm1) - 2) & "') " & _
                                 "WHERE A.FORMNAME = '" & Me.Name & "' AND A.ID = '" & strUserNum & "' AND A.TKIND = '" & iRound & "' AND A.YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
                    cnnConnection.Execute strSql, intI
                 Next mmcnt1
                 '平均
                 strSql = "UPDATE R060332_1 SET MM" & Right(yymm, 2) & "=round(MM" & Right(yymm, 2) & "/12, 0) " & _
                             "WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iRound & "' AND YY00='" & Mid(yymm, 1, Len(yymm) - 2) & "' "
                 cnnConnection.Execute strSql, intI
              End If
           Next mmcnt
        Next yycnt
       '記錄-組合語法
       If iRound = 21 Then
           StrSqlB = Replace(StrSqlB, "TKIND = '22' ", "TKIND = '" & iRound & "' ")
           strSQLc = Replace(strSQLc, "TKIND = '22' ", "TKIND = '" & iRound & "' ")
       Else
           StrSqlB = Replace(StrSqlB, "TKIND = '21' ", "TKIND = '" & iRound & "' ")
           strSQLc = Replace(strSQLc, "TKIND = '21' ", "TKIND = '" & iRound & "' ")
       End If
       '在職工程師
       '因檔案內有組別小計,工程師合計及外專總點數資料,故只抓讀得到員工檔且在職的資料
       strExc(5) = "SELECT '" & iRound & "' as ORD1, DECODE(ST16,'1','1電子電機','2','2化學','3','3日文','4','4機械設計',ST16) 組別,A1N04 編號,ST02 姓名," & _
                        Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(4), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ") WHERE  ST01 IS NOT NULL AND DN04 IS NULL"
       strExc(5) = strExc(5) & strMid01 & ">0"
       '各組小計
       strExc(6) = "SELECT '" & iRound & "' as ORD1, DECODE(SUBSTR(A1N04,1,1),'1','1電子電機','2','2化學','3','3日文','4','4機械設計',A1N04) 組別,' ' 編號,'小計' 姓名," & _
                        Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(4), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04>'1' AND A1N04<'5' AND INSTR(A1N04,'TOT')=2"
       '工程師合計
       strExc(7) = "SELECT '" & iRound & "' as ORD1, '5工程師合計' 組別,' ' 編號,'工程師合計' 姓名," & _
                         Mid(strMid02, InStr(strMid02, "D" & Val(Left(txtFM2(4), 3)) & "12")) & " FROM (" & stCon & " " & StrSqlB & " " & strSQLc & ")" & _
                        " WHERE A1N04='SUBTOT'"
       If iRound = 21 Then
            strExc(9) = strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7)
       Else
            strExc(10) = strExc(5) & " UNION " & strExc(6) & " UNION " & strExc(7)
       End If
   Next iRound
'-----------------分兩次抓資料
   cnnConnection.CommitTrans

   '產生Excel檔: 外專工程師10012~107xx每月OA發文數年移動平均
   '---產生年移動平均.xls
JumpToRe02:
    If rsAD.State = adStateOpen Then
       rsAD.Close
    End If
    rsAD.CursorLocation = adUseClient
    rsAD.Open strExc(9) & " UNION " & strExc(10) & " ORDER BY ord1, 1,2 ", cnnConnection, adOpenStatic, adLockReadOnly
    If rsAD.RecordCount > 0 Then
         Call ProcExcelSave2("2", rsAD)
    End If

    Set rsAD = Nothing

End Sub

'產生Excel檔案-OA發文數
Private Sub ProcExcelSave2(ByVal iType As String, ByRef m_Rst As ADODB.Recordset)
'iType: 1.OA發文數, 2.年移動平均
Dim xlsPoint2 As New Excel.Application
Dim wksPoint2 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String, strA1n04 As String '組別,工程師編號
Dim intPage As Integer '工作表編號
Dim xCols As Integer '行位置
Dim MaxCols As Integer '最大行位置
Dim rowT1 As Integer, rowT2 As Integer '工程師合計、外專總點數的位置
Dim strTKind As String '表1,2
Dim intJ As Integer
Dim strX As String
Dim tmpArray As Variant 'Added by Lydia 2019/02/15

On Error GoTo ErrHnd

   '1.檔名：外專工程師10001~107xx每月OA發文數 ; 2.檔名：外專工程師10012~107xx每月OA發文數年移動平均
   If iType = "1" Then
       'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
       strExc(1) = strSrvDate(1) & "_外專暨日專工程師" & Val(txtFM2(4)) & "~" & Val(txtFM2(5)) & "每月OA發文數"
   Else
       'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
       strExc(1) = strSrvDate(1) & "_外專暨日專工程師" & Val(Left(txtFM2(4), 3)) & "12" & "~" & Val(txtFM2(5)) & "每月OA發文數年移動平均"
   End If
   strFileName = strExcelPath & strExc(1) & MsgText(43)
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint2.SheetsInNewWorkbook = 2 'Added by Lydia 2019/02/15 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint2.Workbooks.add
    xlsPoint2.Visible = False '預設不顯示
    
    m_Rst.MoveFirst
    Do While Not m_Rst.EOF
        '切換工作表
        If strTKind <> "" & m_Rst.Fields(0) Then
            strGrp = ""
            iRow = 1
            xCols = 1
            intPage = intPage + 1
            Set wksPoint2 = xlsPoint2.Worksheets(intPage)
            xlsPoint2.Sheets(intPage).Select '選擇工作表
            xlsPoint2.ActiveWindow.DisplayZeros = False 'Added by Lydia 2019/02/15 設工作表的零值不顯示
            xlsPoint2.Worksheets(intPage).Name = IIf(intPage = 1, "審查意見+申復+核駁+再審", "審查意見+核駁")   '工作表名稱
            '欄位抬頭
            For intJ = 1 To m_Rst.Fields.Count - 1
                strX = Pub_NumberToSystem26(xCols + intJ - 1)
                Select Case intJ
                     Case 1: '組別
                         wksPoint2.Range(strX & ":" & strX).ColumnWidth = 13
                         wksPoint2.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case 2: '編號
                         wksPoint2.Range(strX & ":" & strX).ColumnWidth = 7
                         wksPoint2.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case 3: '姓名
                         wksPoint2.Range(strX & ":" & strX).ColumnWidth = 13
                         wksPoint2.Range(strX & ":" & strX).HorizontalAlignment = xlLeft
                     Case Else '年月統計量
                         wksPoint2.Range(strX & ":" & strX).ColumnWidth = 7
                         wksPoint2.Range(strX & ":" & strX).HorizontalAlignment = xlRight
                End Select
                strExc(2) = Replace("" & m_Rst.Fields(intJ).Name, "D", "")
                If Val(strExc(2)) > 0 Then
                     'Modified by Lydia 2022/12/16 加判斷起始年月yymm >= Val(txtFM2(4))
                     If Val(strExc(2)) >= Val(txtFM2(4)) And Val(strExc(2)) <= Val(txtFM2(5)) Then
                         wksPoint2.Range(strX & iRow).Value = strExc(2)
                         MaxCols = xCols + intJ
                     End If
                Else
                     'Added by Lydia 2021/06/24 姓名抬頭加註記
                     If Check1.Value = 1 And strExc(2) = "姓名" Then
                         strExc(2) = "姓名(*離職)"
                     End If
                     'end 2021/06/24
                     wksPoint2.Range(strX & iRow).Value = strExc(2)
                End If
            Next intJ
            wksPoint2.Range(iRow & ":" & iRow).HorizontalAlignment = xlCenter '置中
            wksPoint2.Range("D2").Select
            xlsPoint2.ActiveWindow.FreezePanes = True '凍結窗格
            wksPoint2.Range("A1").Select
            'Added by Lydia 2020/07/07 統計年月不到一年會出錯
            If MaxCols = 0 Then
                 GoTo JumpToExcept
            Else
            'end 2020/07/07
                 ReDim tmpArray(1 To MaxCols - 1) 'Added by Lydia 2019/02/15
            End If 'Added by Lydia 2020/07/07
            iRow = iRow + 1
        End If

        '不同組別分不同底色
         If strGrp <> "" & m_Rst.Fields("組別") Then
             If Val(Left("" & m_Rst.Fields("組別"), 1)) > 4 Then '合計,多跳一行
                 iRow = iRow + 1
             End If
             wksPoint2.Range(iRow & ":" & iRow).Interior.ColorIndex = 22 '底色
             wksPoint2.Range(Pub_NumberToSystem26(xCols) & iRow).Value = "" & m_Rst.Fields("組別")
         End If
         'Modified by Lydia 2019/02/15 因為office2013逐筆輸入過慢,改成陣列輸入
'         For intJ = 2 To MaxCols - 1
'             strX = Pub_NumberToSystem26(xCols + intJ - 1)
'             wksPoint2.Range(strX & iRow).Value = "" & m_Rst.Fields(intJ)
'             If intJ > 3 Then wksPoint2.Range(strX & iRow).NumberFormat = "#,##0"
'         Next intJ
         For intJ = 2 To MaxCols - 1
             tmpArray(intJ - 1) = "" & m_Rst.Fields(intJ)
         Next intJ
         wksPoint2.Range(Pub_NumberToSystem26(xCols + 1) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 2) & iRow).Value = tmpArray
         wksPoint2.Range(Pub_NumberToSystem26(xCols + 3) & iRow & ":" & Pub_NumberToSystem26(xCols + MaxCols - 2) & iRow).NumberFormat = "#,##0"
        'end 2019/02/15
                
         strTKind = "" & m_Rst.Fields("ord1")
         strGrp = "" & m_Rst.Fields("組別")
         strA1n04 = "" & m_Rst.Fields("編號")
         iRow = iRow + 1
        m_Rst.MoveNext
    Loop
   
JumpToExcept: 'Added by Lydia 2020/07/07
   xlsPoint2.Sheets(1).Select '選擇工作表
   '判斷版本
   If Val(xlsPoint2.Version) < 12 Then
        xlsPoint2.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint2.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsPoint2.Workbooks.Close
   xlsPoint2.Quit
   Set wksPoint2 = Nothing
   Set xlsPoint2 = Nothing
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub

'報表4-每季/每月請款點數分析
Private Sub Process4(ByVal iType As String)
'iType: 3-每季, 4-每月
Dim stCon As String
Dim strMid01 As String
Dim yycnt As Integer, mmcnt As Integer
Dim yymm As String
Dim rsAD As New ADODB.Recordset
Dim strDate1 As String, strDate2 As String
   
'---固定表格 (依EXCEL欄位順序由上而下)
'CREATE TABLE R060332_3 (FORMNAME VARCHAR2(20),ID VARCHAR2(6),TKIND VARCHAR2(1),TKINDNAME VARCHAR2(30),YYMM VARCHAR2(6),
'總點數=A0K01P
'TOTALP NUMBER(16,5),
'工程師總點數=P01 , 若有組別不顯示在P01_X記錄為N
'P01 NUMBER(16,5), P01_1 VARCHAR2(13), P01_2 VARCHAR2(13), P01_3 VARCHAR2(13), P01_4 VARCHAR2(13),
'承辦請款：新案階段以外=承辦請款後續
'P02 NUMBER(16,5),
'承辦請款：新案階段=承辦請款新案
'P03 NUMBER(16,5),
'翻譯費201(不含核稿點數)=PS1-PS2
'P04 NUMBER(16,5),
'分配給其他部門=OTHP
'P05 NUMBER(16,5),
'註１：翻譯費201(含核稿點數)=翻譯費201含核稿點數
'PS1 NUMBER(16,5),
'註２：核稿點數=翻譯費之核稿點數
'PS2 NUMBER(16,5),
'折讓點數=A1K01P
'P06 NUMBER(16,5),
'FMP案安全基金=FMP安全基金101,102,103
'P07 NUMBER(16,5),
'OA委外翻譯費(累計)
'P08 NUMBER(16,5));
'ALTER TABLE R060332_3 ADD PRIMARY KEY (FORMNAME,ID,TKIND,YYMM);
   
   If iType = "3" Then '每季-統計期間
       strDate1 = txtFM2(6)
       strDate2 = txtFM2(7)
       strExc(1) = "每季請款點數分析"
        'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
        ClearQueryLog (Me.Name)
        pub_QL05 = pub_QL05 & ";報表3.各組每季請款點數"
        pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(6) & "00-" & txtFM2(7) & "31"
        pub_QL05 = pub_QL05 & ";組別:" & txtFM2(0) & "-" & txtFM2(1)
        'end 2021/11/16
   ElseIf iType = "4" Then   '每月-統計期間
       strDate1 = txtFM2(8)
       strDate2 = txtFM2(9)
       strExc(1) = "每月請款點數分析"
        'Added By Lydia 2021/11/16 查詢印表記錄檔欄位
        ClearQueryLog (Me.Name)
        pub_QL05 = pub_QL05 & ";報表4.各組每月請款點數"
        pub_QL05 = pub_QL05 & ";統計年月:" & txtFM2(8) & "00-" & txtFM2(9) & "31"
        pub_QL05 = pub_QL05 & ";組別:" & txtFM2(0) & "-" & txtFM2(1)
        'end 2021/11/16
   End If

   cnnConnection.BeginTrans
       strSql = "DELETE FROM R060332_3 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND ='" & iType & "' "
       cnnConnection.Execute strSql

        For yycnt = Val(Left(strDate1, 3)) To Val(Left(strDate2, 3))
           For mmcnt = 1 To 12
              yymm = yycnt * 100 + mmcnt
              stCon = " A1K02>=" & yymm & "00 AND A1K02<=" & yymm & "31 "
              'Modified by Lydia 2022/12/16 加判斷起始年月yymm >= Val(strDate1)
              If yymm >= Val(strDate1) And yymm <= Val(strDate2) Then '超過期限不抓資料
                    strSql = "SELECT B1.A1K01,B1.A1K02, A1K01P,A1K06P,A1L05 AS A1K201P,F21P,F21P201,A1N06P,OTHP,NEWCASE,FMPP,F201P FROM "
                    '總點數A1K01P,折讓點數A1K06P; Memo by Lydia 2021/06/24 注意A1K21為建檔人員
                    strMid01 = "SELECT A1K01,A1K02,A1K11-NVL(A1K06,0)-NVL(A1K09,0)+NVL(A1K36,0) AS A1K01P,NVL(A1K06,0)-NVL(A1K36,0) AS A1K06P FROM ACC1K0,STAFF,CASEPROGRESS " & _
                                  "WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F2%' " & _
                                  "GROUP BY A1K01,A1K02,A1K11-NVL(A1K06,0)-NVL(A1K09,0)+NVL(A1K36,0),NVL(A1K06,0)-NVL(A1K36,0) "
                    strSql = strSql & "(" & strMid01 & ") B1, "
                    'A1K201P翻譯費服務費
                    strSql = strSql & " (SELECT A1L01,NVL(A1L05,0) AS A1L05 FROM ACC1K0,ACC1L0 WHERE " & stCon & _
                                " AND A1K01=A1L01(+) AND A1L04='201' ) B2, "
                    'F21P工程師點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS F21P FROM ACC1K0,ACC1N0,STAFF WHERE " & stCon & _
                                " AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N04=ST01 AND ST03 IN ('F21','F81') GROUP BY A1N01) B3, "
                    strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS F21P FROM ACC1K0,ACC1N0,R060332_STAFF WHERE " & stCon & _
                                " AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N04=ST01 AND '" & strUserNum & "'=STID AND ST03 IN ('F21','F81') GROUP BY A1N01) B3, "
                    'F21P201工程師翻譯費及核稿點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS F21P201 FROM ACC1K0,ACC1N0,STAFF,CASEPROGRESS WHERE " & stCon & _
                                " AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N04=ST01 AND ST03 IN ('F21','F81') AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') GROUP BY A1N01) B4, "
                    strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS F21P201 FROM ACC1K0,ACC1N0,R060332_STAFF,CASEPROGRESS WHERE " & stCon & _
                                " AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N04=ST01 AND '" & strUserNum & "'=STID AND ST03 IN ('F21','F81') AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') GROUP BY A1N01) B4, "
                    'A1N06P翻譯費之核稿點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) A1N06P FROM ACC1K0,ACC1N0,R060332,CASEPROGRESS WHERE " & stCon & _
                                "AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N06 IS NOT NULL " & _
                                "AND A1N04=ST01 AND ST03 IN ('F21','F81') AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') GROUP BY A1N01) B5, "
                    strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) A1N06P FROM ACC1K0,ACC1N0,R060332_STAFF,CASEPROGRESS WHERE " & stCon & _
                                "AND A1K01=A1N01(+) AND A1N02='2' AND A1N04<>'F4102' AND A1N06 IS NOT NULL " & _
                                "AND A1N04=ST01 AND '" & strUserNum & "'=STID  AND ST03 IN ('F21','F81') AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') GROUP BY A1N01) B5, "
                    'OTHP其他部門點數
                    'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF
                    'strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS OTHP FROM ACC1K0,ACC1N0,STAFF WHERE " & stCon & _
                                "AND A1K01=A1N01(+) AND A1N02='2' AND A1N04=ST01 AND ST03 NOT IN ('F21','F81','F23','F22') GROUP BY A1N01) B6, "
                    'Modified by Lydia 2022/12/01 限制F2部門收文的點數
                    'strSql = strSql & " (SELECT A1N01,SUM(NVL(A1N05,0)) AS OTHP FROM ACC1K0,ACC1N0,R060332_STAFF WHERE " & stCon & _
                                "AND A1K01=A1N01(+) AND A1N02='2' AND A1N04=ST01 AND '" & strUserNum & "'=STID AND ST03 NOT IN ('F21','F81','F23','F22') GROUP BY A1N01) B6, "
                    strSql = strSql & " (SELECT A1N01,SUM(A1N05) AS OTHP FROM (SELECT A1N01,A1N04,A1N05 FROM ACC1K0,ACC1N0,CASEPROGRESS " & _
                                  "WHERE " & stCon & " AND A1K01=A1N01(+) AND A1N02='2' AND A1K01=CP60(+) AND CP12 LIKE 'F2%' " & _
                                  "AND A1N04<>'F4102' and A1N04 NOT IN (SELECT ST01 FROM R060332_STAFF WHERE STID='" & strUserNum & "' AND ST03 IN ('F21','F81','F23','F22')) " & _
                                  "GROUP BY A1N01,A1N04,A1N05) GROUP BY A1N01) B6, "
                    'NEWCASE是否新案請款單
                    strSql = strSql & " (SELECT A1K01,'Y' NEWCASE FROM ACC1K0,CASEPROGRESS WHERE " & stCon & _
                                "AND A1K01=CP60(+) AND CP01 IN ('P','FCP') AND CP10 IN ('101','102','103','109') GROUP BY A1K01) B7, "
                    'FMP安全基金101,102,103(20); Memo by Lydia 2021/06/24 注意A1K21為建檔人員
                    strSql = strSql & " (SELECT A1K01,(COUNT(*)*2) AS FMPP FROM ACC1K0 WHERE A1K01 IN ( " & _
                               "SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K13='P' " & _
                               "AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F%' AND CP09<'B' AND CP10 IN ('101','102','103')) GROUP BY A1K01) B8, "
                    'OA委外翻譯費(累計); Memo by Lydia 2021/06/24 注意A1K21為建檔人員
                    strSql = strSql & " (SELECT A1W01,SUM(A1P07) AS F201P FROM ACC1W0,CASEPROGRESS C1,CASEPROGRESS C2,ACC1P0 " & _
                                "WHERE SUBSTR(A1W02,1,1)='B' AND A1W02=C1.CP09(+) AND C1.CP01 IN ('P','FCP') AND C1.CP10='927' AND SUBSTR(C1.CP14,1,1)='F' " & _
                                "AND SUBSTR(C1.CP43,1,1)='C' AND C1.CP61||A1W02=A1P23 AND A1P07>0 AND C1.CP43=C2.CP09(+) AND A1W01 IN " & _
                                "(SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12>='F2' AND CP12<='F29')  GROUP BY A1W01) B9 "
                                
                    strSql = strSql & " WHERE B1.A1K01=B2.A1L01(+) AND B1.A1K01=B3.A1N01(+) AND B1.A1K01=B4.A1N01(+) AND B1.A1K01=B5.A1N01(+) AND B1.A1K01=B6.A1N01(+) AND B1.A1K01=B7.A1K01(+) AND B1.A1K01=B8.A1K01(+) AND B1.A1K01=B9.A1W01(+)"
                    
                    '每月統計
                    strSql = "SELECT SUM(A1K01P/1000) TOTALP,SUM(A1K06P/1000) A1K06P," & _
                                "SUM(DECODE(NEWCASE,'Y',0,NVL(A1K01P/1000,0)-NVL(A1K201P/1000,0)-NVL(F21P,0)+NVL(F21P201,0)-NVL(OTHP,0))) 承辦請款後續, " & _
                                "SUM(DECODE(NEWCASE,'Y',NVL(A1K01P/1000,0)-NVL(A1K201P/1000,0)-NVL(F21P,0)+NVL(F21P201,0)-NVL(OTHP,0),0)) 承辦請款新案, " & _
                                "SUM(OTHP) 分配給其他部門,SUM(A1K201P)/1000 翻譯費201含核稿點數,SUM(A1N06P) 翻譯費之核稿點數, SUM(FMPP) FMP安全基金,SUM(F201P)/1000 AS OA委外翻譯費 " & _
                                "FROM (" & strSql & ") "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                        strExc(0) = "INSERT INTO R060332_3 (FORMNAME,ID,TKIND,TKINDNAME,YYMM,TOTALP,P02,P03,P04,P05,PS1,PS2,P06,P07,P08)" & _
                                         "VALUES ('" & Me.Name & "', '" & strUserNum & "', '" & iType & "', '" & strExc(1) & "', '" & yymm & "', " & CNULL("" & RsTemp.Fields("TOTALP"), True) & ", " & _
                                         CNULL("" & RsTemp.Fields("承辦請款後續"), True) & ", " & CNULL("" & RsTemp.Fields("承辦請款新案"), True) & ", " & _
                                         CNULL(Val("" & RsTemp.Fields("翻譯費201含核稿點數")) - Val("" & RsTemp.Fields("翻譯費之核稿點數")), True) & ", " & CNULL("" & RsTemp.Fields("分配給其他部門"), True) & ", " & _
                                         CNULL("" & RsTemp.Fields("翻譯費201含核稿點數"), True) & ", " & CNULL("" & RsTemp.Fields("翻譯費之核稿點數"), True) & ", " & _
                                         CNULL("" & RsTemp.Fields("A1K06P"), True) & ", " & CNULL("" & RsTemp.Fields("FMP安全基金"), True) & ", " & _
                                         CNULL("" & RsTemp.Fields("OA委外翻譯費")) & ") "
                       cnnConnection.Execute strExc(0)

                       '工程師總點數扣除翻譯費201點數 (P01,P01_1,P01_2,P01_3,P01_4)
                       'Modified by Lydia 2021/06/24 staff改用暫存檔; STAFF=>R060332_STAFF, 注意A1K21為建檔人員,A1N04=點數分配之承辦人
                       ' strExc(0) = "SELECT A.ST16,(TOT-NVL(TFEE,0)) BAL FROM " & _
                                 "(SELECT ST16,SUM(A1N05) TOT FROM ACC1K0,ACC1N0,STAFF WHERE A1K01 IN " & _
                                 " (SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F2%') " & _
                                 "     AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) AND ST03 IN ('F21','F81') AND A1N04<>'F4102'" & _
                                 "   GROUP BY ST16) A," & _
                                 "(SELECT ST16,SUM(A1N05) TFEE FROM ACC1K0,ACC1N0,STAFF,CASEPROGRESS WHERE A1K01 IN " & _
                                 " (SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F2%') " & _
                                 "     AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) AND ST03 IN ('F21','F81') AND A1N06 IS NULL AND A1N03 IS NOT NULL AND A1N04<>'F4102' AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') " & _
                                 "   GROUP BY ST16) B " & _
                                 " WHERE A.ST16=B.ST16(+) ORDER BY A.ST16"
                        strExc(0) = "SELECT A.ST16,(TOT-NVL(TFEE,0)) BAL FROM " & _
                                 "(SELECT ST16,SUM(A1N05) TOT FROM ACC1K0,ACC1N0,R060332_STAFF WHERE A1K01 IN " & _
                                 " (SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F2%') " & _
                                 "     AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) and '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') AND A1N04<>'F4102'" & _
                                 "   GROUP BY ST16) A," & _
                                 "(SELECT ST16,SUM(A1N05) TFEE FROM ACC1K0,ACC1N0,R060332_STAFF,CASEPROGRESS WHERE A1K01 IN " & _
                                 " (SELECT DISTINCT A1K01 FROM ACC1K0,STAFF,CASEPROGRESS WHERE " & stCon & " AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K21=ST01(+) AND A1K01=CP60(+) AND CP12 LIKE 'F2%') " & _
                                 "     AND A1K01=A1N01(+) AND '2'=A1N02(+) AND A1N04=ST01(+) and '" & strUserNum & "'=STID(+) AND ST03 IN ('F21','F81') AND A1N06 IS NULL AND A1N03 IS NOT NULL AND A1N04<>'F4102' AND A1N03=CP09(+) AND CP01||CP10 IN ('P201','FCP201') " & _
                                 "   GROUP BY ST16) B " & _
                                 " WHERE A.ST16=B.ST16(+) ORDER BY A.ST16"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            RsTemp.MoveFirst
                            Do While Not RsTemp.EOF
                                 '工程師總點數
                                 strExc(2) = "UPDATE R060332_3 SET P01=NVL(P01,0) + " & Val(RsTemp.Fields(1)) & " WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iType & "' " & _
                                             "AND YYMM = '" & yymm & "' "
                                 cnnConnection.Execute strExc(2), intI
                                 '各組別小計
                                 If Val(RsTemp.Fields("ST16")) >= txtFM2(0) And Val(RsTemp.Fields("ST16")) <= txtFM2(1) Then
                                        strExc(2) = "UPDATE R060332_3 SET P01_" & RsTemp.Fields("ST16") & "='" & RsTemp.Fields(1) & "' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iType & "' " & _
                                                         "AND YYMM = '" & yymm & "' "
                                 Else '條件外的組別設為N
                                        strExc(2) = "UPDATE R060332_3 SET P01_" & RsTemp.Fields("ST16") & "='N' WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iType & "' " & _
                                                         "AND YYMM = '" & yymm & "' "
                                 End If
                                 cnnConnection.Execute strExc(2), intI
                                 RsTemp.MoveNext
                            Loop
                        End If
                    End If
              End If
           Next mmcnt

        Next yycnt

   cnnConnection.CommitTrans
'-----------------------------------------
   '產生Excel檔
JumpToRe01:
    If rsAD.State = adStateOpen Then
       rsAD.Close
    End If
    rsAD.CursorLocation = adUseClient
    If iType = "3" Then
        strExc(3) = "SELECT SUBSTR(YYMM,1,3) MYY, CEIL(TO_NUMBER(SUBSTR(YYMM,4,5))/3) MMM,SUM(TOTALP) TOTALP, SUM(P01) P01, " & _
                          "SUM(DECODE(P01_1,'N',-1,P01_1)) P01_1, SUM(DECODE(P01_2,'N',-1,P01_2)) P01_2, SUM(DECODE(P01_3,'N',-1,P01_3)) P01_3, SUM(DECODE(P01_4,'N',-1,P01_4)) P01_4, SUM(P02) P02,SUM(P03) P03,SUM(P04) P04, " & _
                          "SUM(P05) P05, SUM(PS1) PS1,SUM(PS2) PS2,SUM(P06) P06,SUM(P07) P07, SUM(P08) P08 " & _
                          "From R060332_3 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iType & "' " & _
                          "GROUP BY SUBSTR(YYMM,1,3), CEIL(TO_NUMBER(SUBSTR(YYMM,4,5))/3) ORDER BY 1, 2 "
        rsAD.Open strExc(3), cnnConnection, adOpenStatic, adLockReadOnly
        If rsAD.RecordCount > 0 Then
            InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
            Call ProcExcelSave3(rsAD)
        'Added by Lydia 2021/11/16
        Else
             InsertQueryLog (0)
        'end 2021/11/16
        End If
    Else
        strExc(3) = "SELECT YYMM,TOTALP,P01,P01_1,P01_2,P01_3,P01_4,P02,P03,P04,P05,PS1,PS2,DECODE(P06,0,'',P06) AS P06,P07,P08 " & _
                          "FROM R060332_3 WHERE FORMNAME = '" & Me.Name & "' AND ID = '" & strUserNum & "' AND TKIND = '" & iType & "'  ORDER BY YYMM "
        rsAD.Open strExc(3), cnnConnection, adOpenStatic, adLockReadOnly
        If rsAD.RecordCount > 0 Then
            InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
            Call ProcExcelSave4(rsAD)
        'Added by Lydia 2021/11/16
        Else
             InsertQueryLog (0)
        'end 2021/11/16
        End If
    End If

    Set rsAD = Nothing

End Sub

'產生Excel檔案-每季請款點數分析
Private Sub ProcExcelSave3(ByRef m_Rst As ADODB.Recordset)
Dim xlsPoint3 As New Excel.Application
Dim wksPoint3 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String, strGrp2 As String
Dim xCols As Integer, MaxCols As Integer '行位置 / 最大行位置
Dim intJ As Integer
Dim strX As String, strX2 As String, strBaseX As String
Dim strCellFormat As String
Dim strCellFormat2 As String
Dim sRows As Integer '資料起始列位置
Dim cInX As Integer
Dim inArr(1 To 5) As Integer
Dim inC1 As Integer, inC2 As Integer, inB1 As Integer, inB2 As Integer

On Error GoTo ErrHnd
    
    'Modified by Lydia 2019/11/07 更名
    'strExc(1) = strSrvDate(1) & "_外專工程師" & Val(txtFM2(6)) & "~" & Val(txtFM2(7)) & "每季請款點數分析"
    'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
    strExc(1) = strSrvDate(1) & "_外專暨日專" & Val(txtFM2(6)) & "~" & Val(txtFM2(7)) & "每季請款點數分析"
    strFileName = strExcelPath & strExc(1) & MsgText(43)
    sRows = 3
    cInX = 22  '底色
    strCellFormat = "#,##0.000"
    strCellFormat2 = "0.00%"
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint3.SheetsInNewWorkbook = 1 'Added by Lydia 2019/02/15 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint3.Workbooks.add
    xlsPoint3.Visible = False '預設不顯示
    Set wksPoint3 = xlsPoint3.Worksheets(1)
    xlsPoint3.Sheets(1).Select '選擇工作表
    xlsPoint3.ActiveWindow.DisplayZeros = False 'Added by Lydia 2019/02/15 設工作表的零值不顯示
    
    m_Rst.MoveFirst
    Do While Not m_Rst.EOF
    
        If strGrp = "" Then
            '預設左邊項目抬頭
            sRows = 3
            iRow = sRows
            wksPoint3.Range("A:A").ColumnWidth = 4
            wksPoint3.Range("A:A").HorizontalAlignment = xlLeft
            wksPoint3.Range("B:B").ColumnWidth = 8
            wksPoint3.Range("B:B").HorizontalAlignment = xlLeft
            wksPoint3.Range("C:C").ColumnWidth = 16
            wksPoint3.Range("C:C").HorizontalAlignment = xlLeft
            
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            wksPoint3.Range("B" & iRow & ":C" & iRow).WrapText = False
            wksPoint3.Range("B" & iRow & ":C" & iRow).HorizontalAlignment = xlCenter
            wksPoint3.Range("B" & iRow & ":C" & iRow).VerticalAlignment = xlBottom
                  
            wksPoint3.Range("A" & iRow).Value = "總點數＝1+2+3+4+5"
            wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint3.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行

            wksPoint3.Range("A" & iRow).Value = "1"
            wksPoint3.Range("B" & iRow).Value = "工程師"
            For intJ = 4 To 7
                 '條件內工程師組別,才顯示
                 If Val("" & m_Rst.Fields(intJ)) >= 0 Then
                     If inC1 = 0 Then inC1 = iRow
                     wksPoint3.Range("C" & iRow).Value = (intJ - 3) & PUB_GetFCPGrpName(intJ - 3)
                      iRow = iRow + 1
                 End If
            Next intJ
            inC2 = iRow - 1
            inArr(1) = iRow
            wksPoint3.Range("B" & iRow).Value = "工程師小計"
            wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            inArr(2) = iRow
            wksPoint3.Range("A" & iRow).Value = "2"
            wksPoint3.Range("B" & iRow).Value = "承辦請款：新案階段以外"
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            inArr(3) = iRow
            wksPoint3.Range("A" & iRow).Value = "3"
            wksPoint3.Range("B" & iRow).Value = "承辦請款：新案階段"
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint3.Range("B" & iRow).Value = "承辦請款小計"
            wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            'Modified by Lydia 2022/12/16 加說明
            wksPoint3.Range("B" & iRow + 1).Value = "(承辦請款含檢視中說50%)"
            wksPoint3.Range("B" & iRow + 1 & ":C" & iRow + 1).MergeCells = True
            'end 2022/12/16
            iRow = iRow + 2 '多空一行
            
            inArr(4) = iRow
            wksPoint3.Range("A" & iRow).Value = "4"
            wksPoint3.Range("B" & iRow).Value = "翻譯費201(不含核稿點數)"
            wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            inArr(5) = iRow
            wksPoint3.Range("A" & iRow).Value = "5"
            wksPoint3.Range("B" & iRow).Value = "分配給其他部門"
            wksPoint3.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            inB1 = iRow
            wksPoint3.Range("A" & iRow).Value = "註１：翻譯費201(含核稿點數)"
            wksPoint3.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            inB2 = iRow
            wksPoint3.Range("A" & iRow).Value = "註２：核稿點數"
            wksPoint3.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            wksPoint3.Range("A" & iRow).Value = "6"
            wksPoint3.Range("B" & iRow).Value = "折讓點數"
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint3.Range("A" & iRow).Value = "7"
            wksPoint3.Range("B" & iRow).Value = "FMP案安全基金"
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint3.Range("A" & iRow).Value = "8"
            wksPoint3.Range("B" & iRow).Value = "OA委外翻譯費支出"
            wksPoint3.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = sRows
            xCols = 2 '資料從D開始
        End If
        
        '切換年月
        If strGrp & strGrp2 <> "" & m_Rst.Fields("MYY") & m_Rst.Fields("MMM") Then
            xCols = xCols + 2
            MaxCols = xCols '最大行位置
            '年度為主,季度為明細
            If strGrp <> "" & m_Rst.Fields("MYY") Then
                wksPoint3.Range(Pub_NumberToSystem26(xCols) & "1").Value = "" & m_Rst.Fields("MYY")
                '合併,置中
                strX = Pub_NumberToSystem26(xCols) '每年起始
                If "" & m_Rst.Fields("MYY") < Left(txtFM2(7), 3) Then
                    strX2 = Pub_NumberToSystem26(xCols + 7)  '第4季止
                Else
                    strExc(10) = GetSeasonL(DBDATE(txtFM2(7) & "01"), strExc(9))
                    strX2 = Pub_NumberToSystem26(xCols + (Val(strExc(9)) - 1) * 2 + 1)  '第x季止
                End If
                With wksPoint3.Range(strX & "1:" & strX2 & "1")
                    .WrapText = False
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                End With
            End If
            '不同季度
            If strGrp2 <> "" & m_Rst.Fields("MMM") Then
                Select Case "" & m_Rst.Fields("MMM")
                   Case "1": strExc(10) = "一"
                   Case "2": strExc(10) = "二"
                   Case "3": strExc(10) = "三"
                   Case "4": strExc(10) = "四"
                End Select
                wksPoint3.Range(Pub_NumberToSystem26(xCols) & "2").Value = strExc(10)
                wksPoint3.Range(Pub_NumberToSystem26(xCols) & "2").HorizontalAlignment = xlCenter '置中
                wksPoint3.Range(Pub_NumberToSystem26(xCols + 1) & "2").Value = "%"
                wksPoint3.Range(Pub_NumberToSystem26(xCols + 1) & "2").HorizontalAlignment = xlCenter '置中
            End If
            iRow = sRows '起始位置
        End If
        
        '從上而下放資料
        strX = Pub_NumberToSystem26(xCols)
        strBaseX = Pub_NumberToSystem26(xCols) & "$" & iRow '除數=總點數
        strX2 = Pub_NumberToSystem26(xCols + 1)
        ' 總點數＝1+2+3+4+5 => 改成公式
        'wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("TOTALP")
        wksPoint3.Range(strX & iRow).Formula = "=" & strX & inArr(1) & "+" & strX & inArr(2) & "+" & strX & inArr(3) & "+" & strX & inArr(4) & "+" & strX & inArr(5)
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        
        iRow = iRow + 2 '多空一行

        For intJ = 4 To 7
             '條件內工程師組別,才顯示
             'If "" & m_Rst.Fields(intJ) <> "N" Then
             If Val("" & m_Rst.Fields(intJ)) >= 0 Then
                  wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields(intJ)
                  wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
                  wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
                  wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
                  wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
                  iRow = iRow + 1
             End If
        Next intJ
        '工程師小計=>改成公式
        'wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P01")
        wksPoint3.Range(strX & iRow).Formula = "=SUM(" & strX & inC1 & ":" & strX & inC2 & ")"
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P02")  '承辦請款：新案階段以外
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 1
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P03")  '承辦請款：新案階段
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 1
        
        wksPoint3.Range(strX & iRow).Formula = "=SUM(" & strX & iRow - 2 & ":" & strX & iRow - 1 & ")" '承辦請款小計
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        '翻譯費201(不含核稿點數) =>改成公式
        'wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P04")
        wksPoint3.Range(strX & iRow).Formula = "=" & strX & inB1 & "-" & strX & inB2
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P05")  '分配給其他部門
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint3.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint3.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint3.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("PS1")  '註１：翻譯費201(含核稿點數)
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("PS2")  '註２：核稿點數
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 2 '多空一行
        
        wksPoint3.Range(strX & iRow).Value = IIf("" & m_Rst.Fields("P06") = "", Empty, "" & m_Rst.Fields("P06")) '折讓點數
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P07")  'FMP案安全基金
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint3.Range(strX & iRow).Value = "" & m_Rst.Fields("P08")  'OA委外翻譯費支出
        wksPoint3.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        
        strGrp = "" & m_Rst.Fields("MYY")  '年度
        strGrp2 = "" & m_Rst.Fields("MMM")    '季度
        m_Rst.MoveNext
    Loop
    
    '框線
    wksPoint3.Range("A1:" & strX2 & iRow - 1).Select
    xlsPoint3.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsPoint3.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsPoint3.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsPoint3.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsPoint3.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsPoint3.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    '最後設年度左邊界加粗
    For intJ = 4 To MaxCols Step 8 '從D欄開始
        wksPoint3.Range(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).Select
        xlsPoint3.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsPoint3.Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Next intJ
    
    wksPoint3.Range("A1").Value = "項目"
    '合併,置中
    With wksPoint3.Range("A1:C2")
        .WrapText = False
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wksPoint3.Range("D3").Select
    xlsPoint3.ActiveWindow.FreezePanes = True '凍結窗格
    wksPoint3.Range("A1").Select
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

'產生Excel檔案-每月請款點數分析
Private Sub ProcExcelSave4(ByRef m_Rst As ADODB.Recordset)
Dim xlsPoint4 As New Excel.Application
Dim wksPoint4 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim strGrp As String
Dim xCols As Integer, MaxCols As Integer '行位置 / 最大行位置
Dim intJ As Integer
Dim strX As String, strX2 As String, strBaseX As String
Dim strCellFormat As String
Dim strCellFormat2 As String
Dim sRows As Integer '資料起始列位置
Dim cInX As Integer
Dim inArr(1 To 5) As Integer
Dim inC1 As Integer, inC2 As Integer, inB1 As Integer, inB2 As Integer

On Error GoTo ErrHnd
    
    'Modified by Lydia 2019/11/07 更名
    'strExc(1) = strSrvDate(1) & "_外專工程師" & Val(txtFM2(8)) & "~" & Val(txtFM2(9)) & "每月請款點數分析"
    'Modified by Lydia 2020/07/07 「外專」更名為「 外專暨日專」
    strExc(1) = strSrvDate(1) & "_外專暨日專" & Val(txtFM2(8)) & "~" & Val(txtFM2(9)) & "每月請款點數分析"
    strFileName = strExcelPath & strExc(1) & MsgText(43)
    sRows = 3
    cInX = 22  '底色
    strCellFormat = "#,##0.000"
    strCellFormat2 = "0.00%"
    
    If Dir(strFileName) <> "" Then
       Kill strFileName
    End If
    xlsPoint4.SheetsInNewWorkbook = 1 'Added by Lydia 2019/02/15 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
    xlsPoint4.Workbooks.add
    xlsPoint4.Visible = False '預設不顯示
    Set wksPoint4 = xlsPoint4.Worksheets(1)
    xlsPoint4.Sheets(1).Select '選擇工作表
    xlsPoint4.ActiveWindow.DisplayZeros = False 'Added by Lydia 2019/02/15 設工作表的零值不顯示
    
    m_Rst.MoveFirst
    Do While Not m_Rst.EOF
    
        If strGrp = "" Then
            '預設左邊項目抬頭
            sRows = 3
            iRow = sRows
            wksPoint4.Range("A:A").ColumnWidth = 4
            wksPoint4.Range("A:A").HorizontalAlignment = xlLeft
            wksPoint4.Range("B:B").ColumnWidth = 8
            wksPoint4.Range("B:B").HorizontalAlignment = xlLeft
            wksPoint4.Range("C:C").ColumnWidth = 16
            wksPoint4.Range("C:C").HorizontalAlignment = xlLeft
            
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            wksPoint4.Range("B" & iRow & ":C" & iRow).WrapText = False
            wksPoint4.Range("B" & iRow & ":C" & iRow).HorizontalAlignment = xlCenter
            wksPoint4.Range("B" & iRow & ":C" & iRow).VerticalAlignment = xlBottom
                  
            wksPoint4.Range("A" & iRow).Value = "總點數＝1+2+3+4+5"
            wksPoint4.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint4.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行

            wksPoint4.Range("A" & iRow).Value = "1"
            wksPoint4.Range("B" & iRow).Value = "工程師"
            For intJ = 3 To 6
                 '條件內工程師組別,才顯示
                 If "" & m_Rst.Fields(intJ) <> "N" Then
                     If inC1 = 0 Then inC1 = iRow
                     wksPoint4.Range("C" & iRow).Value = (intJ - 2) & PUB_GetFCPGrpName(intJ - 2)
                      iRow = iRow + 1
                 End If
            Next intJ
            inC2 = iRow - 1
            inArr(1) = iRow
            wksPoint4.Range("B" & iRow).Value = "工程師小計"
            wksPoint4.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行

            inArr(2) = iRow
            wksPoint4.Range("A" & iRow).Value = "2"
            wksPoint4.Range("B" & iRow).Value = "承辦請款：新案階段以外"
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            inArr(3) = iRow
            wksPoint4.Range("A" & iRow).Value = "3"
            wksPoint4.Range("B" & iRow).Value = "承辦請款：新案階段"
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint4.Range("B" & iRow).Value = "承辦請款小計"
            wksPoint4.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            'Add by Lydia 2022/12/16 加說明
            wksPoint4.Range("B" & iRow + 1).Value = "(承辦請款含檢視中說50%)"
            wksPoint4.Range("B" & iRow + 1 & ":C" & iRow + 1).MergeCells = True
            'end 2022/12/16
            iRow = iRow + 2 '多空一行
            
            inArr(4) = iRow
            wksPoint4.Range("A" & iRow).Value = "4"
            wksPoint4.Range("B" & iRow).Value = "翻譯費201(不含核稿點數)"
            wksPoint4.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            inArr(5) = iRow
            wksPoint4.Range("A" & iRow).Value = "5"
            wksPoint4.Range("B" & iRow).Value = "分配給其他部門"
            wksPoint4.Range(iRow & ":" & iRow).Interior.ColorIndex = cInX
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            inB1 = iRow
            wksPoint4.Range("A" & iRow).Value = "註１：翻譯費201(含核稿點數)"
            wksPoint4.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            inB2 = iRow
            wksPoint4.Range("A" & iRow).Value = "註２：核稿點數"
            wksPoint4.Range("A" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 2 '多空一行
            
            wksPoint4.Range("A" & iRow).Value = "6"
            wksPoint4.Range("B" & iRow).Value = "折讓點數"
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint4.Range("A" & iRow).Value = "7"
            wksPoint4.Range("B" & iRow).Value = "FMP案安全基金"
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = iRow + 1
            wksPoint4.Range("A" & iRow).Value = "8"
            wksPoint4.Range("B" & iRow).Value = "OA委外翻譯費支出"
            wksPoint4.Range("B" & iRow & ":C" & iRow).MergeCells = True
            iRow = sRows
            xCols = 2 '資料從D開始
        End If
        
        '切換年月
        If strGrp <> "" & m_Rst.Fields("yymm") Then
            xCols = xCols + 2
            MaxCols = xCols '最大行位置
            wksPoint4.Range(Pub_NumberToSystem26(xCols) & "1").Value = "" & m_Rst.Fields(0) '年月
            wksPoint4.Range(Pub_NumberToSystem26(xCols) & ":" & Pub_NumberToSystem26(xCols)).ColumnWidth = 10
            '合併,置中
            With wksPoint4.Range(Pub_NumberToSystem26(xCols) & "1:" & Pub_NumberToSystem26(xCols + 1) & "1")
                .WrapText = False
                .MergeCells = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            wksPoint4.Range(Pub_NumberToSystem26(xCols) & "2").Value = "點數"
            wksPoint4.Range(Pub_NumberToSystem26(xCols) & "2").HorizontalAlignment = xlCenter '置中
            wksPoint4.Range(Pub_NumberToSystem26(xCols + 1) & "2").Value = "%"
            wksPoint4.Range(Pub_NumberToSystem26(xCols + 1) & "2").HorizontalAlignment = xlCenter '置中
            iRow = sRows '起始位置
        End If
        
        '從上而下放資料
        strX = Pub_NumberToSystem26(xCols)
        strBaseX = Pub_NumberToSystem26(xCols) & "$" & iRow '除數=總點數
        strX2 = Pub_NumberToSystem26(xCols + 1)
        ' 總點數＝1+2+3+4+5 => 改成公式
        'wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("TOTALP")
        wksPoint4.Range(strX & iRow).Formula = "=" & strX & inArr(1) & "+" & strX & inArr(2) & "+" & strX & inArr(3) & "+" & strX & inArr(4) & "+" & strX & inArr(5)
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        
        iRow = iRow + 2 '多空一行

        For intJ = 3 To 6
             '條件內工程師組別,才顯示
             If "" & m_Rst.Fields(intJ) <> "N" Then
                  wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields(intJ)
                  wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
                  wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
                  wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
                  wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
                  iRow = iRow + 1
             End If
        Next intJ
        '工程師小計=>改成公式
        'wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P01")  '工程師小計
        wksPoint4.Range(strX & iRow).Formula = "=SUM(" & strX & inC1 & ":" & strX & inC2 & ")"
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P02")  '承辦請款：新案階段以外
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 1
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P03")  '承辦請款：新案階段
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 1
        
        wksPoint4.Range(strX & iRow).Formula = "=SUM(" & strX & iRow - 2 & ":" & strX & iRow - 1 & ")" '承辦請款小計
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        '翻譯費201(不含核稿點數) =>改成公式
        'wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P04")
        wksPoint4.Range(strX & iRow).Formula = "=" & strX & inB1 & "-" & strX & inB2
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P05")  '分配給其他部門
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        wksPoint4.Range(strX2 & iRow).Formula = "=" & strX & iRow & "/" & strBaseX
        wksPoint4.Range(strX2 & iRow).NumberFormat = strCellFormat2
        wksPoint4.Range(strX2 & iRow).HorizontalAlignment = xlCenter '百分比置中
        iRow = iRow + 2 '多空一行
        
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("PS1")  '註１：翻譯費201(含核稿點數)
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("PS2")  '註２：核稿點數
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 2 '多空一行
        
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P06")  '折讓點數
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P07")  'FMP案安全基金
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        wksPoint4.Range(strX & iRow).Value = "" & m_Rst.Fields("P08")  'OA委外翻譯費支出
        wksPoint4.Range(strX & iRow).NumberFormat = strCellFormat
        iRow = iRow + 1
        
        strGrp = "" & m_Rst.Fields("YYMM")
        m_Rst.MoveNext
    Loop
    
    '框線
    wksPoint4.Range("A1:" & strX2 & iRow - 1).Select
    xlsPoint4.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsPoint4.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsPoint4.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsPoint4.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsPoint4.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsPoint4.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    '最後設年月左邊界加粗
    For intJ = 4 To MaxCols Step 2 '從D欄開始
        wksPoint4.Range(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).Select
        xlsPoint4.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsPoint4.Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Next intJ
    
    wksPoint4.Range("A1").Value = "項目"
    '合併,置中
    With wksPoint4.Range("A1:C2")
        .WrapText = False
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wksPoint4.Range("D3").Select
    xlsPoint4.ActiveWindow.FreezePanes = True '凍結窗格
    wksPoint4.Range("A1").Select
   '判斷版本
   If Val(xlsPoint4.Version) < 12 Then
        xlsPoint4.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint4.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsPoint4.Workbooks.Close
   xlsPoint4.Quit
   Set wksPoint4 = Nothing
   Set xlsPoint4 = Nothing
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub



