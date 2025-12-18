VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100110_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案件查詢"
   ClientHeight    =   5730
   ClientLeft      =   90
   ClientTop       =   950
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   4
      Left            =   4965
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   3
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   0
      Left            =   6072
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7296
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4980
      Left            =   45
      TabIndex        =   4
      Top             =   720
      Width           =   9225
      _ExtentX        =   16281
      _ExtentY        =   8784
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   16
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   4350
      TabIndex        =   12
      Top             =   457
      Width           =   1605
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2831;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   990
      TabIndex        =   11
      Top             =   457
      Width           =   1605
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2831;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   1110
      TabIndex        =   10
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱 :"
      Height          =   180
      Left            =   3120
      TabIndex        =   7
      Top             =   457
      Width           =   1170
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   457
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "對造號數 :"
      Height          =   180
      Left            =   6330
      TabIndex        =   5
      Top             =   457
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   457
      Width           =   810
   End
End
Attribute VB_Name = "frm100110_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Label2、Label6
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim strSQL1 As String
'Add By Cheng 2002/07/01
Dim strSQL2 As String
'Add By Sindy 2010/02/05
Dim StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSwhSQL1 As String, strSwhSQL2 As String, strSwhSQL3 As String
Dim strSwhSQL4 As String
Dim strSubSQL1 As String, strSubSQL2 As String, strSubSQL3 As String
Dim strSubSQL4 As String
'2010/02/05 End
Dim strSwhSQL5 As String, strSwhSQL6 As String, strSubSQL5 As String, strSubSQL6 As String 'Add by Amy 2013/09/30 +查cp51 cp52
Dim strSql As String, i As Integer, j As Integer, s As Integer, intK As Integer
Dim strTemp As Variant, StrTest4 As String, STRTEMP12 As Variant
Dim StrTest2 As String, strTemp1 As String
Dim StrTest5 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'grdDataList.Cols = 17
Dim intField As Integer
intField = 23
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
grdDataList.ColWidth(2) = 1450
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 3: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(3) = 0
Else
    grdDataList.ColWidth(3) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(4) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "案件性質"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Modify By Cheng 2002/07/01
'grdDataList.Col = 5: grdDataList.Text = "審定號"
grdDataList.col = 6: grdDataList.Text = "證書審定號"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "承辦人"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2002/07/01
grdDataList.col = 8: grdDataList.Text = "智權人員"
grdDataList.ColWidth(8) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "對造號數"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "對造名稱"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2002/07/16
grdDataList.col = 11: grdDataList.Text = "對造案件名稱"
grdDataList.ColWidth(11) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Sindy 2010/02/12
grdDataList.col = 12: grdDataList.Text = "其他相關人"
grdDataList.ColWidth(12) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
'2010/02/12 End
grdDataList.col = 13: grdDataList.Text = "是否出名"
grdDataList.ColWidth(13) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 14: grdDataList.Text = "條款"
grdDataList.ColWidth(14) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
grdDataList.col = 15: grdDataList.Text = "CP09"
grdDataList.ColWidth(15) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/10
grdDataList.col = 16: grdDataList.Text = ""
grdDataList.ColWidth(16) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 16 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01

End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Dim Str01 As String
Dim StrTag As String

Select Case cmdState
Case 0 '相關卷號
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                 grdDataList.col = j
                 grdDataList.CellBackColor = QBColor(15)
            Next j
          grdDataList.col = 2
          If Not IsNull(grdDataList.Text) Then
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                frm100108_3.Show
                frm100108_3.Tag = Pub_RplStr(grdDataList.Text)
                frm100108_3.StrMenu1
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
           End If
       End If
       Next i
       Me.Enabled = True
Case 1 '回前畫面
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2 '結束
   fnCloseAllFrm100
Case 3 '案件基本資料
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
   grdDataList.col = 0
   grdDataList.row = i
   If Trim(grdDataList.Text) = "V" Then
    grdDataList.col = 0
    grdDataList.Text = ""
    For j = 0 To grdDataList.Cols - 1
        grdDataList.col = j
        grdDataList.CellBackColor = QBColor(15)
    Next j
     grdDataList.col = 2
     Str01 = SystemNumber(grdDataList, 1)
     If Mid(UCase(Str01), 1, 1) = "N" Then
         Str01 = Mid(Str01, 2, 3)
     End If
     If Not IsNull(grdDataList.Text) Then
        If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
        End If
         Select Case Pub_RplStr(Str01)
         Case "CFP", "FCP", "P"   '專利
               Screen.MousePointer = vbHourglass
               frm100101_3.Show
               frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_3.StrMenu
               Screen.MousePointer = vbDefault
         Case "CFT", "FCT", "T", "TF"   '商標
               Screen.MousePointer = vbHourglass
               frm100101_4.Show
               frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_4.StrMenu
               Screen.MousePointer = vbDefault
         'Modify By Sindy 2009/07/24 增加LIN系統類別
         'modify by sonia 2019/7/29 +ACS系統類別
         Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
               Screen.MousePointer = vbHourglass
               frm100101_5.Show
               frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_5.StrMenu
               Screen.MousePointer = vbDefault
         Case "LA"            '顧問
               Screen.MousePointer = vbHourglass
               frm100101_6.Show
               frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_6.StrMenu
               Screen.MousePointer = vbDefault
         Case Else                  '服務
              Select Case Pub_RplStr(Str01)
                  Case "TB"    '條碼
                     Screen.MousePointer = vbHourglass
                     frm100101_7.Show
                     frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_7.StrMenu
                     Screen.MousePointer = vbDefault
                  Case "TM"
                     Screen.MousePointer = vbHourglass
                     frm100101_8.Show
                     frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_8.StrMenu
                     Screen.MousePointer = vbDefault
                  Case "TD"
                        Screen.MousePointer = vbHourglass
                      frm100101_9.Show
                     frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_9.StrMenu
                     Screen.MousePointer = vbDefault
                  Case "TC", "CFC"
                     Screen.MousePointer = vbHourglass
                      frm100101_A.Show
                     frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_A.StrMenu
                     Screen.MousePointer = vbDefault
                  Case Else
                     Screen.MousePointer = vbHourglass
                      frm100101_B.Show
                     frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_B.StrMenu
                     Screen.MousePointer = vbDefault
               End Select
        End Select
        Me.Enabled = True
        Exit Sub
     End If
   End If
   Next i
   Me.Enabled = True
Case 4 '案件進度
   Me.Enabled = False
   StrTag = ""
   For i = 1 To grdDataList.Rows - 1
   grdDataList.col = 0
   grdDataList.row = i
   If Trim(grdDataList.Text) = "V" Then
    grdDataList.col = 0
    grdDataList.Text = ""
    For j = 0 To grdDataList.Cols - 1
        grdDataList.col = j
        grdDataList.CellBackColor = QBColor(15)
    Next j
       grdDataList.col = 2
       If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
          Screen.MousePointer = vbHourglass
          frm100101_2.Show
          frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
          frm100101_2.StrMenu
          Screen.MousePointer = vbDefault
          Me.Enabled = True
          Exit Sub
       End If
   End If
   Next i
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()        '對造名稱或案件名稱或號數
Dim ii As Integer
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

On Error Resume Next

'Dim fs, f
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set f = fs.OpenTextFile(App.Path & "\testfile.txt", 8, True, 0)
'f.WriteLine Me.Name & " : StrMenu-【" & frm100110_1.txt1(1) & "：" & frm100110_1.txt1(11) & "：" & frm100110_1.txt1(2) & "】"

ClearQueryLog ("frm100110_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
Label2.Caption = frm100110_1.txtFM2(0)  'Modified by Lydia 2022/01/05  txt1(1)改成txtFM2(0)
Label4.Caption = frm100110_1.txt1(2)
'Add By Cheng 2002/07/16
Me.Label6.Caption = frm100110_1.txtFM2(1)  'Modified by Lydia 2022/01/05  txt1(11)改成txtFM2(1)
Me.Enabled = False
'Add By Sindy 2010/02/05 增加查詢Lawcase,Hirecase,Serverpractice
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
strSwhSQL1 = ""
strSwhSQL2 = ""
strSwhSQL3 = ""
strSwhSQL4 = ""
strSwhSQL5 = "" 'Add by Amy 2013/09/30 +查cp51 cp52
strSwhSQL6 = "" 'Add by Amy 2013/09/30
strSubSQL1 = ""
strSubSQL2 = ""
strSubSQL3 = ""
strSubSQL4 = ""
strSubSQL5 = "" 'Add by Amy 2013/09/30
strSubSQL6 = "" 'Add by Amy 2013/09/30
strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr("", 5) & ") "

'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
m_AllSys = GetAllSysKind(, "ALL")
intCufaCnt = 0
'end 2019/11/01
    
'收文起日
If Len(Trim(frm100110_1.txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    StrSQL4 = StrSQL4 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
End If
'收文迄日
If Len(Trim(frm100110_1.txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    strSQL2 = strSQL2 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    StrSQL4 = StrSQL4 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
End If
If Len(Trim(frm100110_1.txt1(4))) <> 0 Or Len(Trim(frm100110_1.txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label5 & frm100110_1.txt1(4) & "-" & frm100110_1.txt1(5) 'Add By Sindy 2010/11/3
End If
'案件性質
If Len(Trim(frm100110_1.txt1(6))) <> 0 Then
    strSQL1 = strSQL1 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
    strSQL2 = strSQL2 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
    StrSQL3 = StrSQL3 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
    StrSQL4 = StrSQL4 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
    strSQL5 = strSQL5 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
    pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(0) & frm100110_1.txt1(6) 'Add By Sindy 2010/11/3
End If

'申請人起
'Modified by Lydia 2019/11/01 改成申請人1~5
'If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
'    strSQL1 = strSQL1 + " AND TM23>='" & frm100110_1.txt1(7) & "' "
'    strSQL2 = strSQL2 + " AND PA26>='" & frm100110_1.txt1(7) & "' "
'    StrSQL3 = StrSQL3 + " AND LC11>='" & frm100110_1.txt1(7) & "' "
'    StrSQL4 = StrSQL4 + " AND HC05>='" & frm100110_1.txt1(7) & "' "
'    strSQL5 = strSQL5 + " AND SP08>='" & frm100110_1.txt1(7) & "' "
'End If
''申請人迄
'If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
'    strSQL1 = strSQL1 & " AND TM23<='" & frm100110_1.txt1(8) & "' "
'    strSQL2 = strSQL2 & " AND PA26<='" & frm100110_1.txt1(8) & "' "
'    StrSQL3 = StrSQL3 & " AND LC11<='" & frm100110_1.txt1(8) & "' "
'    StrSQL4 = StrSQL4 & " AND HC05<='" & frm100110_1.txt1(8) & "' "
'    strSQL5 = strSQL5 & " AND SP08<='" & frm100110_1.txt1(8) & "' "
'End If
   'Memo by Lydia 2019/11/01 改成變數
   strExc(1) = ""
   If Len(Trim(frm100110_1.txt1(7))) <> 0 And Len(Trim(frm100110_1.txt1(8))) <> 0 Then
       strExc(1) = " AND ((PA26>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA26<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA27>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA27<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA28>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA28<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA29>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA29<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA30>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA30<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )) "
   Else
       If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
            strExc(1) = " AND ((PA26>='" & GetNewFagent(frm100110_1.txt1(7)) & "' )" & _
                                      " OR (PA27>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA28>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA29>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA30>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )) "
       End If
       If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
            strExc(1) = " AND ((PA26<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                      " OR (PA27<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA28<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA29<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA30<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )) "
       End If
   End If
   'end 2019/11/01
   
If Len(Trim(frm100110_1.txt1(7))) <> 0 Or Len(Trim(frm100110_1.txt1(8))) <> 0 Then
      'Added by Lydia 2019/11/01 合併SQL
      strSQL1 = strSQL1 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "TM23"), "PA27", "TM78"), "PA28", "TM79"), "PA29", "TM80"), "PA30", "TM81")
      strSQL2 = strSQL2 & strExc(1)
      StrSQL3 = StrSQL3 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "LC11"), "PA27", "LC43"), "PA28", "LC44"), "PA29", "LC45"), "PA30", "LC46")
      StrSQL4 = StrSQL4 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "HC05"), "PA27", "HC24"), "PA28", "HC25"), "PA29", "HC26"), "PA30", "HC27")
      strSQL5 = strSQL5 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "SP08"), "PA27", "SP58"), "PA28", "SP59"), "PA29", "SP65"), "PA30", "SP66")
      'end 2019/11/01
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(1) & frm100110_1.txt1(7) & "-" & frm100110_1.txt1(8) 'Add By Sindy 2010/11/3
End If
'代理人起
If Len(Trim(frm100110_1.txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND TM44>='" & frm100110_1.txt1(9) & "' "
    strSQL2 = strSQL2 + " AND PA75>='" & frm100110_1.txt1(9) & "' "
    StrSQL3 = StrSQL3 + " AND LC22>='" & frm100110_1.txt1(9) & "' "
    'strSQL4 = 無代理人
    strSQL5 = strSQL5 + " AND SP26>='" & frm100110_1.txt1(9) & "' "
End If
'代理人迄
If Len(Trim(frm100110_1.txt1(10))) <> 0 Then
    strSQL1 = strSQL1 & " AND TM44<='" & frm100110_1.txt1(10) & "' "
    strSQL2 = strSQL2 & " AND PA75<='" & frm100110_1.txt1(10) & "' "
    StrSQL3 = StrSQL3 & " AND LC22<='" & frm100110_1.txt1(10) & "' "
    'strSQL4 = 無代理人
    strSQL5 = strSQL5 & " AND SP26<='" & frm100110_1.txt1(10) & "' "
End If
If Len(Trim(frm100110_1.txt1(9))) <> 0 Or Len(Trim(frm100110_1.txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(2) & frm100110_1.txt1(9) & "-" & frm100110_1.txt1(10) 'Add By Sindy 2010/11/3
End If
'對造名稱
'Modified by Lydia 2022/01/05  txt1(1)改成txtFM2(0)
If Len(Trim(frm100110_1.txtFM2(0))) <> 0 Then
'    strSQL1 = strSQL1 + " AND (instr(CP40,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP41,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP42 ,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP50 ,'" & frm100110_1.txtfm2(0) & "')>0) "
'    strSQL2 = strSQL2 + " AND (instr(CP40,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP41,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP42 ,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP50 ,'" & frm100110_1.txtfm2(0) & "')>0) "
'    strSQL3 = strSQL3 + " AND (instr(CP40,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP41,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP42 ,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP50 ,'" & frm100110_1.txtfm2(0) & "')>0) "
'    strSQL4 = strSQL4 + " AND (instr(CP40,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP41,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP42 ,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP50 ,'" & frm100110_1.txtfm2(0) & "')>0) "
'    strSQL5 = strSQL5 + " AND (instr(CP40,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP41,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP42 ,'" & frm100110_1.txtfm2(0) & "')>0 OR instr(CP50 ,'" & frm100110_1.txtfm2(0) & "')>0) "
    strSwhSQL1 = " CP40>' ' "
    strSwhSQL2 = " CP41>' ' "
    strSwhSQL3 = " CP42>' ' "
    strSwhSQL4 = " CP50>' ' "
    strSwhSQL5 = " CP51>' ' " 'Add by Amy 2013/09/30
    strSwhSQL6 = " CP52>' ' " 'Add by Amy 2013/09/30
    strSubSQL1 = " AND instr(CP40,'" & frm100110_1.txtFM2(0) & "')>0 "
    strSubSQL2 = " AND instr(upper(CP41),'" & UCase(frm100110_1.txtFM2(0)) & "')>0 "
    strSubSQL3 = " AND instr(CP42,'" & frm100110_1.txtFM2(0) & "')>0 "
    strSubSQL4 = " AND instr(CP50,'" & frm100110_1.txtFM2(0) & "')>0 "
    strSubSQL5 = " AND instr(Upper(CP51),'" & UCase(frm100110_1.txtFM2(0)) & "')>0 " 'Add By Amy 2013/09/30
    strSubSQL6 = " AND instr(CP52,'" & frm100110_1.txtFM2(0) & "')>0 " 'Add By Amy 2013/09/30
    pub_QL05 = pub_QL05 & ";" & frm100110_1.Option1(1).Caption & frm100110_1.txtFM2(0) 'Add By Sindy 2010/11/3
End If
'end 2022/01/05

'Add By Cheng 2002/07/16
'對造案件名稱
'Modified by Lydia 2022/01/05  txt1(11)改成txtFM2(1)
If Len(Trim(frm100110_1.txtFM2(1))) <> 0 Then
    If Len(Trim(frm100110_1.txt1(1))) <> 0 Then
       strSQL1 = strSQL1 + " AND (CP37 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' OR upper(CP38) LIKE '%" & UCase(frm100110_1.txtFM2(1).Text) & "%' OR CP39 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' ) "
       strSQL2 = strSQL2 + " AND (CP37 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' OR upper(CP38) LIKE '%" & UCase(frm100110_1.txtFM2(1).Text) & "%' OR CP39 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' ) "
       StrSQL3 = StrSQL3 + " AND (CP37 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' OR upper(CP38) LIKE '%" & UCase(frm100110_1.txtFM2(1).Text) & "%' OR CP39 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' ) "
       StrSQL4 = StrSQL4 + " AND (CP37 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' OR upper(CP38) LIKE '%" & UCase(frm100110_1.txtFM2(1).Text) & "%' OR CP39 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' ) "
       strSQL5 = strSQL5 + " AND (CP37 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' OR upper(CP38) LIKE '%" & UCase(frm100110_1.txtFM2(1).Text) & "%' OR CP39 LIKE '%" & frm100110_1.txtFM2(1).Text & "%' ) "
    Else
       strSwhSQL1 = " CP37>' ' "
       strSwhSQL2 = " CP38>' ' "
       strSwhSQL3 = " CP39>' ' "
       strSwhSQL4 = " 1<>1 "
       strSubSQL1 = " AND instr(CP37,'" & frm100110_1.txtFM2(1) & "')>0 "
       strSubSQL2 = " AND instr(upper(CP38),'" & UCase(frm100110_1.txtFM2(1)) & "')>0 "
       strSubSQL3 = " AND instr(CP39,'" & frm100110_1.txtFM2(1) & "')>0 "
       strSubSQL4 = " AND 1<>1 "
    End If
    pub_QL05 = pub_QL05 & ";" & frm100110_1.lbl2 & frm100110_1.txtFM2(1) 'Add By Sindy 2010/11/3
End If
'end 2022/01/05
'對造號數
If Len(Trim(frm100110_1.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP36='" & frm100110_1.txt1(2) & "' "
   strSQL2 = strSQL2 + " AND CP36='" & frm100110_1.txt1(2) & "' "
   StrSQL3 = StrSQL3 + " AND CP36='" & frm100110_1.txt1(2) & "' "
   StrSQL4 = StrSQL4 + " AND CP36='" & frm100110_1.txt1(2) & "' "
   strSQL5 = strSQL5 + " AND CP36='" & frm100110_1.txt1(2) & "' "
   '2010/9/23 add by sonia 只下對造號數程式會錯誤
   'Modified by Lydia 2022/01/05  txt1(1)改成txtFM2(0)、 txt1(11)改成txtFM2(1)
   If Len(Trim(frm100110_1.txtFM2(0))) = 0 And Len(Trim(frm100110_1.txtFM2(1))) = 0 Then
      strSwhSQL1 = " CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL2 = " CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL3 = " CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL4 = " CP36='" & frm100110_1.txt1(2) & "' "
   Else
      strSwhSQL1 = strSwhSQL1 & " AND CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL2 = strSwhSQL2 & " AND CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL3 = strSwhSQL3 & " AND CP36='" & frm100110_1.txt1(2) & "' "
      strSwhSQL4 = strSwhSQL4 & " AND CP36='" & frm100110_1.txt1(2) & "' "
   End If
   '2010/9/23 end
   pub_QL05 = pub_QL05 & ";" & frm100110_1.LBL1 & frm100110_1.txt1(2) 'Add By Sindy 2010/11/3
End If
'Add by Amy 2014/09/25 +申請國家
If Len(Trim(frm100110_1.txt1(14))) <> 0 Then
    strSQL1 = strSQL1 & " AND TM10>='" & frm100110_1.txt1(14) & "' "
    strSQL2 = strSQL2 & " AND PA09>='" & frm100110_1.txt1(14) & "' "
    StrSQL3 = StrSQL3 & " AND LC15>='" & frm100110_1.txt1(14) & "' "
    strSQL5 = strSQL5 & " AND SP09>='" & frm100110_1.txt1(14) & "' "
End If
If Len(Trim(frm100110_1.txt1(15))) <> 0 Then
    strSQL1 = strSQL1 & " AND TM10<='" & frm100110_1.txt1(15) & "' "
    strSQL2 = strSQL2 & " AND PA09<='" & frm100110_1.txt1(15) & "' "
    StrSQL3 = StrSQL3 & " AND LC15<='" & frm100110_1.txt1(15) & "' "
    strSQL5 = strSQL5 & " AND SP09<='" & frm100110_1.txt1(15) & "' "
End If
If Len(Trim(frm100110_1.txt1(14))) <> 0 Or Len(Trim(frm100110_1.txt1(15))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(3) & frm100110_1.txt1(14) & "-" & frm100110_1.txt1(15)
End If
'end 2014/09/25
'edit by nickc 2005/05/10
'strSQL = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSQL = strSQL + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
'strSQL = strSQL + " ORDER BY 收文日,本所案號 "
'Modify By Sindy 2010/02/05
'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
'strSql = strSql + " ORDER BY 收文日,FSort,本所案號 "
'2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
'商標
'Modified by Lydia 2019/11/01 +增加欄位 SeColTM
strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL3
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL4
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL5
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColTM & " FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL6

'專利
'Modified by Lydia 2019/11/01 +增加欄位 SeColPA
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL3
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL4
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL5
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColPA & " FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL6

'法務
'Modified by Lydia 2019/11/01 +增加欄位 SeColLC
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL3
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL4
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL5
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColLC & " FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL6

'顧問
'Modified by Lydia 2019/11/01 +增加欄位 SeColHC
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL3
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL4
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL5
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColHC & " FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL6

'服務
'Modified by Lydia 2019/11/01 +增加欄位 SeColSP
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL3
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL4
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL5
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & SeColSP & " FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL6

strSql = strSql + " ORDER BY 收文日,FSort,本所案號 "
CheckOC
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

    'Added by Lydia 2019/11/01 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
    End If
    'end 2019/11/01
    
    cmdOK(0).Enabled = True
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
'    f.WriteLine "5-1"
'    mdiMain.Timer1.Enabled = False
    'ShowNoData
    Pub_Can_Copy_Pic = True 'Added by Morgan 2011/12/26
    MsgBox "資料庫中搜尋不到符合的對造資料!!", vbInformation, "沒有資料 " & Now
    Pub_Can_Copy_Pic = False 'Added by Morgan 2011/12/26
'    mdiMain.Timer1.Enabled = True
'    f.WriteLine "5-2"
    cmdOK(0).Enabled = False
'    f.WriteLine "5-3"
    Me.Enabled = True
'    f.WriteLine "5-4"
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
'    f.WriteLine "5-5"
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
'    f.WriteLine "5-End"
'    f.Close
    Exit Sub
End If
'f.Close
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 15), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True
End Sub

'Add By Sindy 2012/7/18 For 以申請人查詢 呼叫
'Remove by Lydia 2020/01/03 已不再使用這段程式
'Function StrMenu_2(strChkName As String)        '對造名稱或案件名稱或號數
'Dim ii As Integer
'On Error Resume Next
'
'ClearQueryLog ("frm100110_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
'Label2.Caption = strChkName
'Me.Enabled = False
'
'strSQL1 = ""
'strSQL2 = ""
'StrSQL3 = ""
'StrSQL4 = ""
'strSQL5 = ""
'strSwhSQL1 = ""
'strSwhSQL2 = ""
'strSwhSQL3 = ""
'strSwhSQL4 = ""
'strSwhSQL5 = "": strSwhSQL6 = "" 'Add by Amy 2013/09/30
'strSubSQL1 = ""
'strSubSQL2 = ""
'strSubSQL3 = ""
'strSubSQL4 = ""
'strSubSQL5 = "": strSubSQL6 = "" 'Add by Amy 2013/09/30
'strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
'strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
'
''對造名稱
'If Len(Trim(strChkName)) <> 0 Then
'    strSwhSQL1 = " CP40>' ' "
'    strSwhSQL2 = " CP41>' ' "
'    strSwhSQL3 = " CP42>' ' "
'    strSwhSQL4 = " CP50>' ' "
'    strSwhSQL5 = " CP51>' ' "  'Add by Amy 2013/09/30
'    strSwhSQL6 = " CP52>' ' "  'Add by Amy 2013/09/30
'    strSubSQL1 = " AND instr(CP40,'" & ChgSQL(strChkName) & "')>0 "
'    strSubSQL2 = " AND instr(upper(CP41),'" & UCase(ChgSQL(strChkName)) & "')>0 "
'    strSubSQL3 = " AND instr(CP42,'" & ChgSQL(strChkName) & "')>0 "
'    strSubSQL4 = " AND instr(CP50,'" & ChgSQL(strChkName) & "')>0 "
'    strSubSQL5 = " AND instr(Upper(CP51),'" & UCase(ChgSQL(strChkName)) & "')>0 "  'Add by Amy 2013/09/30
'    strSubSQL6 = " AND instr(CP52,'" & ChgSQL(strChkName) & "')>0 "                         'Add by Amy 2013/09/30
'    pub_QL05 = pub_QL05 & ";對造名稱：" & strChkName
'End If
'
''Add by Amy 2013/09/30 +查cp51 cp52
''商標
'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL3
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL4
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL5 '2013/09/30
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL6 '2013/09/30
'
''專利
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL3
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL4
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL5  '2013/09/30
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL6  '2013/09/30
'
''法務
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL3
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL4
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL5  '2013/09/30
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(LC36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),LAWCASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL6  '2013/09/30
'
''顧問
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL3
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL4
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL5  '2013/09/30
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(HC20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 案件性質,' ' AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),HIRECASE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL6  '2013/09/30
'
''服務
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL1 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL2 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL3 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL3
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL4 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL4
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL5 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL5  '2013/09/30
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,' ' AS 案件性質,SP11 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP36 AS 對造號數,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,NVL(CP37,NVL(CP38,CP39)) AS 對造案件名稱,NVL(CP50,NVL(CP51,NVL(CP52,CP56))) as 其他相關人,DECODE(CP22,'N','否','Y','是','') AS 是否出名,CP49 AS 條款, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where " & strSwhSQL6 & "),SERVICEPRACTICE,STAFF S1,STAFF S2,casepropertymap WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL6  '2013/09/30
'
'strSql = strSql + " ORDER BY 收文日,FSort,本所案號 "
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'   InsertQueryLog (adoRecordset.RecordCount)
'   cmdOK(0).Enabled = True
'Else
'   InsertQueryLog (0) 'Add By Sindy 2010/11/3
'   Pub_Can_Copy_Pic = True 'Added by Morgan 2011/12/26
'   MsgBox "資料庫中搜尋不到符合的對造資料!!", vbInformation, "沒有資料 " & Now
'   Pub_Can_Copy_Pic = False 'Added by Morgan 2011/12/26
'   cmdOK(0).Enabled = False
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Function
'End If
'Me.grdDataList.Visible = False
'Set grdDataList.Recordset = adoRecordset
'SetDataListWidth
'For ii = 1 To Me.grdDataList.Rows - 1
'    Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 15), "1")
'Next ii
'Me.grdDataList.Visible = True
'CheckOC
'Me.Enabled = True
'End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm100110_3 = Nothing
End Sub

Private Sub grdDataList_SelChange()
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
grdDataList.Visible = True
End Sub
