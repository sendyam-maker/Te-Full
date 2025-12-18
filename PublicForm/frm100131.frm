VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100131 
   BorderStyle     =   1  '單線固定
   Caption         =   "程式修改公告查詢"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8580
   ControlBox      =   0   'False
   Icon            =   "frm100131.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8580
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2790
      TabIndex        =   40
      Top             =   1695
      Width           =   2595
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   930
         TabIndex        =   4
         Top             =   0
         Width           =   720
         VariousPropertyBits=   671105051
         MaxLength       =   5
         Size            =   "1270;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LbName 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   42
         Top             =   0
         Width           =   825
         BackColor       =   -2147483638
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "新增人員："
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   41
         Top             =   30
         Width           =   975
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "外商"
      Height          =   375
      Index           =   15
      Left            =   5415
      TabIndex        =   39
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "是"
      Height          =   255
      Left            =   3765
      TabIndex        =   8
      Top             =   2775
      Width           =   615
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細(&B)"
      Height          =   400
      Left            =   6720
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   45
      Width           =   800
   End
   Begin VB.CheckBox Check2 
      Caption         =   "否"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "是"
      Height          =   255
      Index           =   0
      Left            =   1095
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox cboDepName 
      Height          =   300
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   1155
      Width           =   2120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7560
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   45
      Width           =   800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "每日.每月批次"
      Height          =   375
      Index           =   14
      Left            =   6495
      TabIndex        =   38
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "電腦中心"
      Height          =   375
      Index           =   13
      Left            =   5415
      TabIndex        =   37
      Top             =   2760
      Width           =   1080
   End
   Begin VB.CheckBox Check1 
      Caption         =   "檔案室"
      Height          =   375
      Index           =   11
      Left            =   6495
      TabIndex        =   35
      Top             =   2355
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      Caption         =   "法務"
      Height          =   375
      Index           =   8
      Left            =   5415
      TabIndex        =   32
      Top             =   2355
      Width           =   840
   End
   Begin VB.CheckBox Check1 
      Caption         =   "外專"
      Height          =   375
      Index           =   6
      Left            =   6495
      TabIndex        =   30
      Top             =   1515
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "共同查詢"
      Height          =   375
      Index           =   4
      Left            =   6495
      TabIndex        =   28
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "薪資"
      Height          =   375
      Index           =   3
      Left            =   5415
      TabIndex        =   27
      Top             =   1080
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      Caption         =   "分所出納"
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   26
      Top             =   675
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "財務"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   24
      Top             =   675
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      Caption         =   "帳務"
      Height          =   375
      Index           =   1
      Left            =   6495
      TabIndex        =   25
      Top             =   675
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "專利處"
      Height          =   375
      Index           =   5
      Left            =   5415
      TabIndex        =   29
      Top             =   1515
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "商標處"
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   31
      Top             =   1515
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "承辦人"
      Height          =   375
      Index           =   9
      Left            =   6495
      TabIndex        =   33
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "收文"
      Height          =   375
      Index           =   10
      Left            =   7440
      TabIndex        =   34
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "人事"
      Height          =   375
      Index           =   12
      Left            =   7455
      TabIndex        =   36
      Top             =   2355
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   45
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2364
      Left            =   48
      TabIndex        =   18
      Top             =   3156
      Width           =   8436
      _ExtentX        =   14880
      _ExtentY        =   4170
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|上線日期|序號|需求部門|需求人員|請作單日期|摘要|內容|系統別"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "共 0 筆"
      Height          =   180
      Left            =   72
      TabIndex        =   43
      Top             =   5544
      Width           =   540
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   645
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1215
      TabIndex        =   5
      Top             =   2250
      Width           =   3180
      VariousPropertyBits=   671105051
      Size            =   "5609;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   1695
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   5
      Size            =   "1270;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   645
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "附件未簽回："
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   23
      Top             =   2775
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      Height          =   180
      Left            =   4500
      TabIndex        =   22
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "摘要或內容："
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "(民國年月日)"
      Height          =   255
      Index           =   6
      Left            =   3030
      TabIndex        =   20
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "是否公佈："
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2775
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "公佈系統別："
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   17
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2010
      TabIndex        =   16
      Top             =   675
      Width           =   135
   End
   Begin MSForms.Label LbName 
      Height          =   300
      Index           =   1
      Left            =   1830
      TabIndex        =   15
      Top             =   1695
      Width           =   795
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1402;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "上線日期："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   675
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "需求部門："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1155
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "需求人員："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1725
      Width           =   975
   End
End
Attribute VB_Name = "frm100131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/15 Form2.0已修改
'2013/03/25 Create by Amy
Option Explicit

Dim RbMain As New ADODB.Recordset, bp As New ADODB.Recordset

'執行各項功能的權限
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean
Dim m_bolFinalCheck As Boolean '最後檢查控制
Dim bolSelData As Boolean 'griddatalist
Dim i As Integer
Dim strField() As String 'Add by Amy 2024/08/19
Dim bolNoData As Boolean, intLimit As Integer, strCaseSys As String, strAccSys As String, arrCaseSys, arrAccSys 'Add by Amy 2024/08/27

Private Sub cmdDetail_Click()
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
Dim i As Integer
Me.Enabled = False
    For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            Dim Str01, Str02 As String
            grdDataList.col = 0
            grdDataList.Text = ""
            Call SetGridColor(0, strField) 'Modify by Amy 2024/08/19
            '取上線日
            grdDataList.col = 1
            Str01 = grdDataList.Text
            '取序號
             grdDataList.col = 2
             Str02 = grdDataList.Text
             
            If Not IsNull(grdDataList.Text) Then
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                frm100131_1.Show
                frm100131_1.Tag = Str01 & "," & Str02
                frm100131_1.StrMenu
                Screen.MousePointer = vbDefault
              End If
              Me.Enabled = True
              Exit Sub
           End If
           Next i
           Me.Enabled = True

End Sub

Private Sub cmdExit_Click()
    fnCloseAllFrm100
End Sub

'Modify By Amy 2013/05/08 原Private
Public Sub cmdSearch_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql, strTemp As String
    Dim strWhere, chkWhere, returnVal As String
    Dim idx, grdCountCol As Integer
    Dim oCheck As CheckBox
    
    returnVal = ""
    grdDataList.Clear
    grdDataList.Rows = 2
     '欄位驗證
    If Len(Text1(0)) > 0 Then
        If PUB_CheckKeyInDate(Text1(0)) = -1 Then
            Text1(0).SetFocus
            Text1_GotFocus 0
            Exit Sub
         End If
         
         If Len(Text1(1)) > 0 Then
            If PUB_CheckKeyInDate(Text1(1)) = -1 Then
                Text1(1).SetFocus
                Text1_GotFocus 1
                Exit Sub
            End If
            '上線日期終止日有值才確認 起始日不可大於終止日
            If Not nickChgRan(Text1(0), Text1(1), "日期") Then
               Text1(1).SetFocus
               Text1_GotFocus (1)
               Exit Sub
            End If
         Else
         '若上線日期終止日為空則代系統日
            Text1(1).Text = strSrvDate(2)
         End If
         m_bolFinalCheck = True
         strWhere = strWhere & " And BU01 Between '" & ChangeTStringToWString(Text1(0)) & "' And '" & ChangeTStringToWString(Text1(1)) & "'"
    Else
        MsgBox ("請輸入上線日期條件 !")
        Text1(0).SetFocus 'Add By Sindy 2021/12/27
        Exit Sub
    End If
    
    If Len(Text1(2)) > 0 Then
        If Len(Text1(2)) = 5 Then
            strTemp = GetStaffName(Text1(2), True)
            If Len(Trim(strTemp)) > 0 Then
                LbName(1).Caption = strTemp
                m_bolFinalCheck = True
                strWhere = strWhere & " And BU03='" & Text1(2).Text & "'"
            Else
                MsgBox ("無此員工")
                LbName(1).Caption = ""
                m_bolFinalCheck = False
                Exit Sub
            End If
       Else
            MsgBox ("員工編號輸入有誤")
            LbName(1).Caption = ""
            m_bolFinalCheck = False
            Exit Sub
        End If
    Else
            LbName(1).Caption = ""
            m_bolFinalCheck = True
     End If
     
     'Add by Amy 2015/02/10 +新增人員
     If Trim(Text1(4)) <> "" Then
          strWhere = strWhere & " And BU08='" & Text1(4) & "' "
     End If
     
    'Added by Lydia 2023/12/27
    If strSrvDate(1) >= 新部門啟用日 Then
       If Len(cboDepName.Text) > 0 Then strWhere = strWhere & " And A0921='" & Left(cboDepName.Text, 3) & "'"
    Else
    'end 2023/12/27
       If Len(cboDepName.Text) > 0 Then strWhere = strWhere & " And A0901='" & Left(cboDepName.Text, 3) & "'"
    End If
    '2013/04/29 Add by Amy 增加查詢摘要或內容模糊比對
    If Len(Text1(3).Text) > 0 Then
      'Modify by Amy 2025/06/25 +轉大寫 ex:上線日查1100101-1140630 摘要或內容 輸 invoice 大寫會出現6筆,小寫出現2筆
       strWhere = strWhere & " And (INSTR(UPPER(BU05),'" & UCase(Text1(3)) & "')>=1 OR INSTR(UPPER(BU14),'" & UCase(Text1(3)) & "')>=1)"
    End If
    
    If (Check2(0).Value = 1 And Check2(1).Value = 0) Or (Check2(0).Value = 0 And Check2(1).Value = 1) Then
        If Check2(0).Value = 1 Then
            strWhere = strWhere & " And BU06='1'"
        Else
            strWhere = strWhere & " And BU06='0'"
        End If
    End If
    
    'Add by Amy 2024/08/27 不可查詢之項目
    If intLimit <> 0 Then
         strWhere = strWhere & " And  bu07<>'Salary,'"
    End If
    
    'Add by Amy 2015/01/14 +if 增加查詢附件未簽回
    If Check3.Value = 1 Then
             strWhere = strWhere & " And Not Exists(Select * From ImgByteFile " & _
                                "Where IBF01=SubStr(BU01-19110000,1,3) And IBF02= SubStr(BU01-19110000,4)||Decode(length(BU02),1, '0'||BU02,BU02) And IBF03='0' And IBF04='00' And IBF05='5' ) "
    End If
    
    For Each oCheck In Check1
       idx = oCheck.Index
       If oCheck.Value Then
          Select Case idx
           Case 0
             returnVal = " Or instr(BU07,'Account,') >0"
           Case 1
             returnVal = " Or instr(BU07,'Finance,') >0"
           Case 2
            returnVal = " Or instr(BU07,'Casher,') >0"
           Case 3
              returnVal = " Or instr(BU07,'Salary,') >0"
           Case 4
              returnVal = " Or instr(BU07,'Query,') >0"
           Case 5
              returnVal = " Or instr(BU07,'Patpro,') >0"
           Case 6
              returnVal = " Or instr(BU07,'Patpro1,') >0"
           Case 7
              returnVal = " Or instr(BU07,'Trademark,') >0"
           Case 8
              returnVal = " Or instr(BU07,'Law,') >0"
           Case 9
             returnVal = " Or instr(BU07,'Promoter,') >0"
           Case 10
             returnVal = " Or instr(BU07,'Writer,') >0"
           Case 11
            returnVal = " Or instr(BU07,'File,') >0"
           Case 12
            returnVal = " Or instr(BU07,'Person,') >0"
           Case 13
             returnVal = " Or instr(BU07,'Computer,') >0"
           Case 14
             returnVal = " Or instr(BU07,'AutoBatch,') >0"
           'Add by Amy 2018/11/14 +外商
           Case 15
             returnVal = " Or instr(BU07,'Trademark1,') >0"
         End Select
         chkWhere = chkWhere & returnVal
       End If
    Next
            
    If Len(chkWhere) > 1 Then
         '去掉第一個Or
         chkWhere = Mid(chkWhere, 5)
         If InStr(chkWhere, "Or") > 0 Then
            strWhere = strWhere & " And (" & chkWhere & ")"
         Else
           strWhere = strWhere & " And " & chkWhere
         End If
    End If
    
    Screen.MousePointer = vbHourglass
    If m_bolFinalCheck = True Then
        'Added by Lydia 2023/12/27
        If strSrvDate(1) >= 新部門啟用日 Then
           If Pub_StrUserSt03 = "M51" Then '電腦中心才顯示是否公佈
               strSql = "Select '' as V,sqldatet(BU01) as 上線日, BU02 as 序號,DECODE(SIGN(TO_NUMBER(SUBSTR(BU01,1,6))-" & Val(Left(新部門啟用日, 6)) & "),-1,A0902, NVL(A0922,A0902)) as 需求部門,ST02 as 需求人員,sqldatet(BU04) as 請作單日期,Decode(BU06,1,'是','否') as 公佈,BU05 as 摘要, BU14 as 內容,BU15 as 時數,BU07 as 系統別 " & _
                        "From PGMBulletin,STAFF,ACC090,ACC090New " & _
                        "Where BU03=ST01 and ST03=A0901 AND ST93=A0921(+)" & strWhere
           Else
               strSql = "Select '' as V,sqldatet(BU01) as 上線日, BU02 as 序號,DECODE(SIGN(TO_NUMBER(SUBSTR(BU01,1,6))-" & Val(Left(新部門啟用日, 6)) & "),-1,A0902, NVL(A0922,A0902)) as 需求部門,ST02 as 需求人員,sqldatet(BU04) as 請作單日期,BU05 as 摘要,  BU14 as 內容,BU07 as 系統別 " & _
                        "From PGMBulletin,STAFF,ACC090,ACC090NEW " & _
                        "Where BU03=ST01 and ST03=A0901 AND ST93=A0921(+)" & strWhere
           End If
        Else
        'end if
           If Pub_StrUserSt03 = "M51" Then '電腦中心才顯示是否公佈
               'Modify by Amy 2014/07/16 +時數BU15
               strSql = "Select '' as V,sqldatet(BU01) as 上線日, BU02 as 序號,A0902 as 需求部門,ST02 as 需求人員,sqldatet(BU04) as 請作單日期,Decode(BU06,1,'是','否') as 公佈,BU05 as 摘要, BU14 as 內容,BU15 as 時數,BU07 as 系統別 From PGMBulletin,STAFF,ACC090 " & _
                           "Where BU03=ST01 and ST03=A0901" & strWhere
   
           Else
               strSql = "Select '' as V,sqldatet(BU01) as 上線日, BU02 as 序號,A0902 as 需求部門,ST02 as 需求人員,sqldatet(BU04) as 請作單日期,BU05 as 摘要,  BU14 as 內容,BU07 as 系統別 From PGMBulletin,STAFF,ACC090 " & _
                           "Where BU03=ST01 and ST03=A0901" & strWhere
           End If
        End If
        
        'Add by Amy 2014/07/21 +排序
        'modify by sonia 2018/8/20 +BU02排序
        strSql = strSql & " Order by BU01 Desc, BU02"
        
        lblCount = "共  筆" 'Added by Morgan 2024/3/11
        bolNoData = False 'Add by Amy 2024/08/19
        
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
            Set grdDataList.Recordset = rsTmp
            Call SetGridRowHide 'Add by Amy 2024/08/27 電腦中心 或 薪資系統可操作人員 可看薪資
            cmdDetail.Enabled = True
            grdCountCol = grdDataList.Cols - 1
            
            For i = 1 To grdDataList.Rows - 1
               If grdDataList.TextMatrix(i, grdCountCol) <> "" Then
                     Dim strAs, strAsVal As String
                     Dim strTmp() As String
                     Dim j As Integer
                     strAsVal = ""
                     strTmp = Split(grdDataList.TextMatrix(i, grdCountCol), ",")
                     For j = 0 To UBound(strTmp) - 1
                            Select Case strTmp(j)
                             Case "Account"
                                strAs = "財務"
                             Case "Finance"
                                strAs = "帳務"
                             Case "Casher"
                               strAs = "分所出納"
                             Case "Salary"
                               strAs = "薪資"
                             Case "Query"
                               strAs = "共同查詢"
                             Case "Patpro"
                               strAs = "專利處"
                             Case "Patpro1"
                               strAs = "外專"
                             Case "Trademark"
                               strAs = "商標處" 'Modify by Amy 2018/11/14 商標拆 2個
                             Case "Law"
                               strAs = "法務"
                             Case "Promoter"
                               strAs = "承辦人"
                             Case "Writer"
                               strAs = "收文"
                             Case "File"
                               strAs = "檔案室"
                             Case "Person"
                               strAs = "人事"
                             Case "Computer"
                               strAs = "電腦中心"
                             Case "AutoBatch"
                               strAs = "每日.每月批次"
                             'Add by Amy 2018/11/14
                             Case "Trademark1"
                               strAs = "外商"
                        End Select
                        strAsVal = strAsVal & "," & strAs
                     Next
                     strAsVal = Mid(strAsVal, 2)
                     grdDataList.TextMatrix(i, grdCountCol) = strAsVal
               End If
            Next
            Call SetGridColor(0, strField) 'Add by Amy 2024/08/19
                        
            '若查詢結果只有一筆資料
            If Me.grdDataList.Rows = 2 Then
               'Modify by Amy 2024/08/27 +if bolNoData
               If bolNoData = True Then
                  MsgBox "查無資料！"
               Else
                  grdDataList.row = 1
                  grdDataList.col = 1
                  If grdDataList.Text <> "" Then
                     '直接選定
                     bolSelData = True
                     grdDataList.Visible = False
                     grdDataList.row = 1
                     grdDataList.col = 0
                     grdDataList.Text = "V"
                     'Modify by Amy 2024/08/19 Gird變色改至SetGridColor
                     Call SetGridColor(1, strField)
                     grdDataList.Visible = True
                  End If
               End If
            End If
        Else
            SetGridWidth
            grdDataList.Rows = 2
            MsgBox "查無資料"
        End If
        'Modify by Amy 2024/08/27 筆數顯示從上面搬下來,改抓Grid列數 原:rsTmp.RecordCount
        lblCount = grdDataList.Rows - 1
        If bolNoData = True Then lblCount = "0"
        lblCount = "共 " & lblCount & " 筆" 'Added by Morgan 2024/3/11
        'end 2024/08/27
   End If 'm_bolFinalCheck = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetGridWidth 'Add by Amy 2024/08/19 從Form_Active搬過來
   '電腦中心才顯示是否公布及每日每月批次
   'Modify by Amy 2024/08/27 電腦中心 或 薪資系統可操作人員 可看薪資
   strCaseSys = "Patpro;Patpro1;Trademark;Law;Promoter;Writer;File;Computer;Trademark1" '未列AutoBatch;
   arrCaseSys = Split(strCaseSys, ";")
   strAccSys = "Account;Finance;Query;Casher"
   arrAccSys = Split(strAccSys, ";")
   If Pub_StrUserSt03 = "M51" Then
      Label2(5).Visible = True
      Check2(0).Visible = True
      Check2(1).Visible = True
      Check1(14).Visible = True
      'Add by Amy 2015/01/14 +附件未簽回
      Label2(8).Visible = True
      Check3.Visible = True
      'Modify By Sindy 2021/12/27 把3個欄位放入Frame1, Form2.0物件的Visible改變會影響TabIndex有關
'      'Add by Amy 2015/02/10 +新增人員
'      Label2(9).Visible = True
'      Text1(4).Visible = True
'      LbName(0).Visible = True
      Frame1.Visible = True
      '2021/12/27 END
      intLimit = 0 '全可看
   Else
      intLimit = 1 '薪資+案件系統就顯示
      Label2(5).Visible = False
      Check2(0).Visible = False
      Check2(1).Visible = False
      Check1(14).Visible = False
      'Add by Amy 2015/01/14 +附件未簽回
      Label2(8).Visible = False
      Check3.Visible = False
      Frame1.Visible = False 'Add By Sindy 2021/12/27
      Check2(0).Value = 1    'add by sonia 2024/2/5 一般使用者只可查要公告的資料
      If InStr(Pub_GetSpecMan("薪資系統可操作人員"), strUserNum) > 0 Then
         intLimit = 0 '全可看
      ElseIf Pub_StrUserSt03 = "M21" Then
         intLimit = 2 '有人事系統人事部都可看
      ElseIf Pub_StrUserSt03 = "M31" Then
         intLimit = 3 '有財務系統財務部都可看
      ElseIf UCase(App.EXEName) = "TECASHER" Or UCase(App.EXEName) = "CASHER" Then
         intLimit = 4 '以Casher系統登入,Casher系統別都可看(使用Casher系統都不是財務部門)
      End If
   End If
   'end 2024/08/27
   bolSelData = False
   bolToEndByNick = False
   SetComboData
   m_bolFinalCheck = True
   cmdDetail.Enabled = False
   LbName(0).Caption = ""
   LbName(1).Caption = ""
End Sub

Private Sub GrdDataList_Click()
    bolSelData = True
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
     If grdDataList.Text = "V" Then
       grdDataList.Text = ""
       Call SetGridColor(0, strField) 'Modify by Amy 2024/08/19
       bolSelData = False
     Else
      grdDataList.Text = "V"
      Call SetGridColor(1, strField) 'Modify by Amy 2024/08/19
     End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index))
   'add by sonia 2014/10/29
   Select Case Index
      Case 3
         OpenIme
      Case Else
         CloseIme
   End Select
   'end 2014/10/29
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strTemp As String
    Select Case Index
        Case 0, 1
            If PUB_CheckKeyInDate(Me.Text1(Index)) = -1 Then
            Text1(Index).SetFocus
            Text1_GotFocus Index
            m_bolFinalCheck = False
            Exit Sub
         End If
         If Index = 1 Then
           If Not nickChgRan(Text1(0), Text1(1), "日期") Then
               Text1(1).SetFocus
               Text1_GotFocus (1)
               m_bolFinalCheck = False
               Exit Sub
           End If
         End If
        'Modify by Amy 2015/02/10
        Case 2, 4
            If Len(Text1(Index)) > 0 Then
                If Len(Text1(Index)) = 5 Then
                    strTemp = GetStaffName(Text1(Index), True)
                    If Len(Trim(strTemp)) > 0 Then
                        If Index = 2 Then
                            LbName(1).Caption = strTemp
                        Else
                            LbName(0).Caption = strTemp
                        End If
                    Else
                        MsgBox ("無此員工")
                        m_bolFinalCheck = False
                        Exit Sub
                    End If
                Else
                    MsgBox ("員工編號輸入有誤")
                    m_bolFinalCheck = False
                    Exit Sub
                End If
            Else
                If Index = 2 Then
                    LbName(1).Caption = ""
                Else
                    LbName(0).Caption = ""
                End If
            End If
        'end 2015/02/10
    End Select
End Sub
Private Sub SetComboData()
'宣告變數
Dim Rs As New ADODB.Recordset

   'Added By Lydia 2023/12/27
   If strSrvDate(1) >= 新部門啟用日 Then
      Call SetST93Combo(cboDepName)
   Else
   'end 2023/12/27
      Me.cboDepName.Clear
      Rs.CursorLocation = adUseClient
      Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
               cnnConnection, adOpenStatic, adLockReadOnly
      Me.cboDepName.AddItem ""
      While Not Rs.EOF
         Me.cboDepName.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
         Rs.MoveNext
      Wend
      If Rs.State <> adStateClosed Then Rs.Close
      Set Rs = Nothing
   End If

End Sub

Private Sub SetGridWidth()
   '設欄寬:電腦中心才顯示是否公布
   With grdDataList
   If Pub_StrUserSt03 = "M51" Then
           'Modify by Amy 2014/07/16 +時數
           .FormatString = "V|上線日|序號|需求部門|需求人員|請作單日期|公佈|摘要|說明|時數|系統別"
           .ColWidth(0) = 200
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 885
           .ColAlignment(1) = flexAlignCenterCenter
           .ColWidth(2) = 0
      
           .ColWidth(3) = 960
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 885
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 1025
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 500
           .ColAlignment(6) = flexAlignLeftCenter
           .ColWidth(7) = 1350
           .ColAlignment(7) = flexAlignLeftCenter
           .ColWidth(8) = 1350
           .ColAlignment(8) = flexAlignLeftCenter
           'Add by Amy 2014/07/16 +時數
           .ColWidth(9) = 500
           .ColAlignment(9) = flexAlignLeftCenter
           'end 2014/07/16
           .ColWidth(10) = 1200
           .ColAlignment(10) = flexAlignLeftCenter
   Else
           .FormatString = .FormatString
           .ColWidth(0) = 200
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 885
           .ColAlignment(1) = flexAlignCenterCenter
           .ColWidth(2) = 0
      
           .ColWidth(3) = 960
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 885
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 1025
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 1350
           .ColAlignment(6) = flexAlignLeftCenter
           .ColWidth(7) = 1350
           .ColAlignment(7) = flexAlignLeftCenter
           .ColWidth(8) = 1200
           .ColAlignment(8) = flexAlignLeftCenter
  End If
  strField = Split(.FormatString, "|")  'Add by Amy 2024/08/19
 End With
End Sub

'無使用
Private Sub TxtClear()
   Dim txt As Object, Lbl As Object, Chk As Object
   For Each txt In Text1
      txt.Text = ""
   Next
   For Each Lbl In LbName
      Lbl = ""
   Next
   For Each Chk In Check2
      Chk.Value = 0
   Next
   For Each Chk In Check1
      Chk.Value = 0
   Next
End Sub

'Add by Amy 2024/08/19 Grid變色,避免有未改到拆出
'intChoose:0-未選取(依狀態設定顏色)/1-選取
Private Sub SetGridColor(intChoose As Integer, arrCol() As String)
   Dim j As Integer
   Dim iXRow As Integer 'Added by Morgan 2024/8/20
   
   '未選取
   If intChoose = 0 Then
      For j = 0 To grdDataList.Cols - 1
         grdDataList.col = j
         grdDataList.CellBackColor = QBColor(15)
      Next j
   '選取
   Else
      For j = 0 To grdDataList.Cols - 1
         grdDataList.col = j
         grdDataList.CellBackColor = &HFFC0C0
      Next j
   End If
   'Add by Amy 2024/06/27 電腦中心 不公告 顯示黃顏色
   If Pub_StrUserSt03 = "M51" Then
      iXRow = grdDataList.row 'Added by Morgan 2024/8/20
      For j = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(j, GetColVal(strField, "公佈", LBound(strField))) = "否" Then
            grdDataList.row = j
            grdDataList.col = GetColVal(strField, "公佈", LBound(strField))
            grdDataList.CellBackColor = vbYellow
         End If
      Next j
      grdDataList.row = iXRow 'Added by Morgan 2024/8/20
   End If
End Sub

'Add by Amy 2024/08/27
Private Sub SetGridRowHide()
'intLimit:0-全可看/1-薪資+案件系統就顯示/2-有人事系統人事部都可看/3-有財務系統財務部都可看/4-以Casher系統登入,Casher系統別都可看
   Dim j As Integer, k As Integer, n As Integer, strChkSys As String, bData As Boolean
   
   If intLimit = 0 Then Exit Sub
   
   grdDataList.Visible = False
   
   For j = grdDataList.Rows - 1 To 1 Step -1
      bData = False
      'Memo by Amy 可能要刪除,從前頭刪時當 j>目前刪後的列數會錯 ex:資料只有2筆,刪第1筆後,j=2會抓不到第2筆資料會錯
      strChkSys = grdDataList.TextMatrix(j, GetColVal(strField, "系統別", LBound(strField)))
      If InStr(strChkSys, "Salary,") > 0 Then
         '薪資+任一 案件系統就顯示 ex:請作單 1120601-01
         For k = LBound(arrCaseSys) To UBound(arrCaseSys)
            If InStr(strChkSys, arrCaseSys(k) & ",") > 0 Then
               bData = True
               Exit For
            End If
         Next k
         If bData = False Then
            '人事部門,人事系統都可看 ex:請作單 1120608-02
            If intLimit = 2 And InStr(strChkSys, "Person,") > 0 Then
               bData = True
               Exit For
            '財務部門,財務系統都可看 ex:請作單 1121005-01
            ElseIf intLimit = 3 Then
               For k = LBound(arrAccSys) To UBound(arrAccSys)
                  If InStr(strChkSys, arrAccSys(k) & ",") > 0 Then
                     bData = True
                     Exit For
                  End If
               Next k
            '以Casher登入,Casher系統都可看 ex:請作單 1051228-01
            ElseIf intLimit = 4 And InStr(strChkSys, "Casher,") > 0 Then
               bData = True
               Exit For
            End If
         End If
         If bData = False Then
            If grdDataList.Rows - 1 = 1 Then
               For n = 0 To grdDataList.Cols - 1
                  grdDataList.TextMatrix(j, n) = ""
                  bolNoData = True
               Next n
            Else
               grdDataList.RemoveItem (j)
            End If
         End If
      End If
   Next j
   grdDataList.Visible = True
End Sub
