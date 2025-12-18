VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100116_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以國別查詢"
   ClientHeight    =   5730
   ClientLeft      =   285
   ClientTop       =   1710
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7308
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6084
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8532
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5016
      Left            =   36
      TabIndex        =   4
      Top             =   684
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8837
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
   Begin VB.Label lbl1 
      Height          =   255
      Left            =   1284
      TabIndex        =   6
      Top             =   444
      Width           =   3216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   450
      Width           =   1080
   End
End
Attribute VB_Name = "frm100116_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String
Dim s As Integer, i As Integer, j As Integer, intK As Integer
Dim strSql As String, StrTest As String, strTemp As Variant
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String, Str05 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
'edit by nickc 2007/03/23
'grdDataList.Cols = 17
grdDataList.Cols = 18
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 2: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(2) = 0
Else
    grdDataList.ColWidth(2) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "申請人"
grdDataList.ColWidth(4) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "審定號/證書號數"
grdDataList.ColWidth(5) = 1600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "申請案號"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "申請國家"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "申請日"
grdDataList.ColWidth(8) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = ""
grdDataList.ColWidth(9) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = ""
grdDataList.ColWidth(10) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = ""
grdDataList.ColWidth(11) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = ""
grdDataList.ColWidth(12) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 13: grdDataList.Text = ""
grdDataList.ColWidth(13) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 14: grdDataList.Text = ""
grdDataList.ColWidth(14) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 15: grdDataList.Text = ""
grdDataList.ColWidth(15) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/13
grdDataList.col = 16: grdDataList.Text = ""
grdDataList.ColWidth(16) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2007/03/23
grdDataList.col = 17: grdDataList.Text = "PCT"
If frm100116_1.ChkPCT.Value = vbChecked Then
    grdDataList.ColWidth(17) = 620
Else
    grdDataList.ColWidth(17) = 0
End If
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      Screen.MousePointer = vbHourglass
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
        Dim Str01 As String
        grdDataList.col = 1
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
                  Me.Hide
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
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
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
     Me.Enabled = False
     Screen.MousePointer = vbDefault
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
         grdDataList.col = 1
         Screen.MousePointer = vbHourglass
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
     Screen.MousePointer = vbDefault
Case 2
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 1
        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
        Select Case Pub_RplStr(Str01)
            Case "CFP", "FCP", "P"   '專利
                  Screen.MousePointer = vbHourglass
                  frm100101_3.Show
                  'frm100101_3.Hide
                   
                  'Modify By Cheng 2002/04/26
'                  frm100101_3.Tag = grdDataList.Text
                  frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_3.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_3.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_3.Visible
                  Unload frm100101_3
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  'frm100101_4.Hide
                   
                  'Modify By Cheng 2002/04/26
'                  frm100101_4.Tag = grdDataList.Text
                  frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_4.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_4.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_4.Visible
                  Unload frm100101_4
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
                  Screen.MousePointer = vbHourglass
                  frm100101_5.Show
                  'frm100101_5.Hide
                            
                  'Modify By Cheng 2002/04/26
'                  frm100101_5.Tag = grdDataList.Text
                  frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_5.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_5.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_5.Visible
                  Unload frm100101_5
            Case "LA"            '顧問
                  Screen.MousePointer = vbHourglass
                  frm100101_6.Show
                  'frm100101_6.Hide
                      
                  'Modify By Cheng 2002/04/26
'                  frm100101_6.Tag = grdDataList.Text
                  frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_6.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_6.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_6.Visible
                  Unload frm100101_6
            Case Else                  '服務
                 Select Case Pub_RplStr(Str01)
                     Case "TB"    '條碼
                        Screen.MousePointer = vbHourglass
                        frm100101_7.Show
                        'frm100101_7.Hide
                         
                        'Modify By Cheng 2002/04/26
'                        frm100101_7.Tag = grdDataList.Text
                        frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_7.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_7.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_7.Visible
                         Unload frm100101_7
                     Case "TM"
                         Screen.MousePointer = vbHourglass
                         frm100101_8.Show
                        'frm100101_8.Hide
                         
                        'Modify By Cheng 2002/04/26
'                        frm100101_8.Tag = grdDataList.Text
                        frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_8.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_8.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_8.Visible
                         Unload frm100101_8
                     Case "TD"
                         Screen.MousePointer = vbHourglass
                         frm100101_9.Show
                        'frm100101_9.Hide
                         
                        'Modify By Cheng 2002/04/26
'                        frm100101_9.Tag = grdDataList.Text
                        frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_9.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_9.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_9.Visible
                         Unload frm100101_9
                     Case "TC", "CFC"
                         Screen.MousePointer = vbHourglass
                         frm100101_A.Show
                         'frm100101_A.Hide
                         
                        'Modify By Cheng 2002/04/26
'                        frm100101_A.Tag = grdDataList.Text
                        frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_A.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_A.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_A.Visible
                         Unload frm100101_A
                     Case Else
                         Screen.MousePointer = vbHourglass
                         frm100101_B.Show
                        'frm100101_B.Hide
                         
                        'Modify By Cheng 2002/04/26
'                        frm100101_B.Tag = grdDataList.Text
                        frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_B.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_B.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_B.Visible
                         Unload frm100101_B
                  End Select
        End Select
        End If
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
              grdDataList.col = j
              grdDataList.CellBackColor = QBColor(15)
         Next j
        
     End If
     Next i
     Screen.MousePointer = vbDefault
     Me.Enabled = True
     Me.Show
Case 1
     Me.Enabled = False
     Screen.MousePointer = vbDefault
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 1
         Screen.MousePointer = vbHourglass
         If Not IsNull(grdDataList.Text) Then
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            'frm100101_2.Hide
             
            'Modify By Cheng 2002/04/26
'            frm100101_2.Tag = grdDataList.Text
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm100101_2.Show
            Screen.MousePointer = vbDefault
            Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
            Loop Until Not frm100101_2.Visible
            Unload frm100101_2
         End If
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
              grdDataList.col = j
              grdDataList.CellBackColor = QBColor(15)
         Next j
     End If
     Next i
     Me.Enabled = True
     Screen.MousePointer = vbDefault
     Me.Show
Case 2
     Me.Hide
Case 3
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   If frm100116_1.Option1(0).Value = True Then
       Label1.Caption = "申請人國籍"
       lbl1.Caption = frm100116_1.txt1(0) + "-" + frm100116_1.txt1(1)
   Else
       Label1.Caption = "申請國家"
       lbl1.Caption = frm100116_1.txt1(2) + "-" + frm100116_1.txt1(3)
   End If
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()
   If frm100116_1.Option1(0).Value = True Then
      pub_QL05 = pub_QL05 & ";查詢" & frm100116_1.Option1(0).Caption 'Add By Sindy 2010/11/15
      StrMenu1        '申請人國籍
   Else
      pub_QL05 = pub_QL05 & ";查詢" & frm100116_1.Option1(1).Caption 'Add By Sindy 2010/11/15
      StrMenu2        '申請國家
   End If
End Sub

Sub StrMenu1()          '申請人國籍
Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
If Trim(frm100116_1.txt1(4)) = "1" Then
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP05>=" & Val(ChangeTStringToWString(frm100116_1.txt1(5))) & " "
   End If
   If Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(frm100116_1.txt1(6))) & " "
   Else
      If Len(frm100116_1.txt1(5).Text) > 0 Then
         StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
      End If
   End If
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Or Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & frm100116_1.Label3 & frm100116_1.txt1(5) & "-" & frm100116_1.txt1(6) 'Add By Sindy 2010/11/15
   End If
Else
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP27>=" & Val(ChangeTStringToWString(frm100116_1.txt1(5))) & " "
   End If
   If Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(frm100116_1.txt1(6))) & " "
   Else
      If Len(frm100116_1.txt1(5).Text) > 0 Then
         StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
      End If
   End If
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Or Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      pub_QL05 = pub_QL05 & ";發文" & frm100116_1.Label3 & frm100116_1.txt1(5) & "-" & frm100116_1.txt1(6) 'Add By Sindy 2010/11/15
   End If
End If
If Len(Trim(frm100116_1.txt1(7))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 1) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100116_1.Label4, 5) & frm100116_1.txt1(7) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100116_1.txt1(8))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10>='" & frm100116_1.txt1(8) & "' "
End If
If Len(Trim(frm100116_1.txt1(9))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10<='" & frm100116_1.txt1(9) & "' "
End If
If Len(Trim(frm100116_1.txt1(8))) <> 0 Or Len(Trim(frm100116_1.txt1(9))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100116_1.Label5(0) & frm100116_1.txt1(8) & "-" & frm100116_1.txt1(9) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100116_1.txt1(0))) <> 0 And Len(Trim(frm100116_1.txt1(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND ((cu1.CU10>='" & frm100116_1.txt1(0) & "' and cu1.CU10<='" & frm100116_1.txt1(1) & "z') or (cu2.CU10>='" & frm100116_1.txt1(0) & "' and cu2.CU10<='" & frm100116_1.txt1(1) & "z') or (cu3.CU10>='" & frm100116_1.txt1(0) & "' and cu3.CU10<='" & frm100116_1.txt1(1) & "z') or (cu4.CU10>='" & frm100116_1.txt1(0) & "' and cu4.CU10<='" & frm100116_1.txt1(1) & "z') or (cu5.CU10>='" & frm100116_1.txt1(0) & "' and cu5.CU10<='" & frm100116_1.txt1(1) & "z'))"
         strSQL2 = strSQL2 & " AND ((cu1.CU10>='" & frm100116_1.txt1(0) & "' and cu1.CU10<='" & frm100116_1.txt1(1) & "z') or (cu2.CU10>='" & frm100116_1.txt1(0) & "' and cu2.CU10<='" & frm100116_1.txt1(1) & "z') or (cu3.CU10>='" & frm100116_1.txt1(0) & "' and cu3.CU10<='" & frm100116_1.txt1(1) & "z') or (cu4.CU10>='" & frm100116_1.txt1(0) & "' and cu4.CU10<='" & frm100116_1.txt1(1) & "z') or (cu5.CU10>='" & frm100116_1.txt1(0) & "' and cu5.CU10<='" & frm100116_1.txt1(1) & "z'))"
         'Modify By Sindy 2014/12/25
         'StrSQL3 = StrSQL3 & " AND CU10>='" & frm100116_1.txt1(0) & "' AND CU10<='" & frm100116_1.txt1(1) & "z' "
         'StrSQL4 = StrSQL4 & " AND CU10>='" & frm100116_1.txt1(0) & "' AND CU10<='" & frm100116_1.txt1(1) & "z' "
         StrSQL3 = StrSQL3 & " AND ((cu1.CU10>='" & frm100116_1.txt1(0) & "' and cu1.CU10<='" & frm100116_1.txt1(1) & "z') or (cu2.CU10>='" & frm100116_1.txt1(0) & "' and cu2.CU10<='" & frm100116_1.txt1(1) & "z') or (cu3.CU10>='" & frm100116_1.txt1(0) & "' and cu3.CU10<='" & frm100116_1.txt1(1) & "z') or (cu4.CU10>='" & frm100116_1.txt1(0) & "' and cu4.CU10<='" & frm100116_1.txt1(1) & "z') or (cu5.CU10>='" & frm100116_1.txt1(0) & "' and cu5.CU10<='" & frm100116_1.txt1(1) & "z'))"
         StrSQL4 = StrSQL4 & " AND ((cu1.CU10>='" & frm100116_1.txt1(0) & "' and cu1.CU10<='" & frm100116_1.txt1(1) & "z') or (cu2.CU10>='" & frm100116_1.txt1(0) & "' and cu2.CU10<='" & frm100116_1.txt1(1) & "z') or (cu3.CU10>='" & frm100116_1.txt1(0) & "' and cu3.CU10<='" & frm100116_1.txt1(1) & "z') or (cu4.CU10>='" & frm100116_1.txt1(0) & "' and cu4.CU10<='" & frm100116_1.txt1(1) & "z') or (cu5.CU10>='" & frm100116_1.txt1(0) & "' and cu5.CU10<='" & frm100116_1.txt1(1) & "z'))"
         '2014/12/25 END
         strSQL5 = strSQL5 & " AND ((cu1.CU10>='" & frm100116_1.txt1(0) & "' and cu1.CU10<='" & frm100116_1.txt1(1) & "z') or (cu2.CU10>='" & frm100116_1.txt1(0) & "' and cu2.CU10<='" & frm100116_1.txt1(1) & "z') or (cu3.CU10>='" & frm100116_1.txt1(0) & "' and cu3.CU10<='" & frm100116_1.txt1(1) & "z') or (cu4.CU10>='" & frm100116_1.txt1(0) & "' and cu4.CU10<='" & frm100116_1.txt1(1) & "z') or (cu5.CU10>='" & frm100116_1.txt1(0) & "' and cu5.CU10<='" & frm100116_1.txt1(1) & "z'))"
Else
      If Len(Trim(frm100116_1.txt1(0))) <> 0 Then
         strSQL1 = strSQL1 & " AND (cu1.CU10>='" & frm100116_1.txt1(0) & "' or cu2.CU10>='" & frm100116_1.txt1(0) & "' or cu3.CU10>='" & frm100116_1.txt1(0) & "' or cu4.CU10>='" & frm100116_1.txt1(0) & "' or cu5.CU10>='" & frm100116_1.txt1(0) & "')"
         strSQL2 = strSQL2 & " AND (cu1.CU10>='" & frm100116_1.txt1(0) & "' or cu2.CU10>='" & frm100116_1.txt1(0) & "' or cu3.CU10>='" & frm100116_1.txt1(0) & "' or cu4.CU10>='" & frm100116_1.txt1(0) & "' or cu5.CU10>='" & frm100116_1.txt1(0) & "')"
         'Modify By Sindy 2014/12/25
         'StrSQL3 = StrSQL3 & " AND CU10>='" & frm100116_1.txt1(0) & "' "
         'StrSQL4 = StrSQL4 & " AND CU10>='" & frm100116_1.txt1(0) & "' "
         StrSQL3 = StrSQL3 & " AND (cu1.CU10>='" & frm100116_1.txt1(0) & "' or cu2.CU10>='" & frm100116_1.txt1(0) & "' or cu3.CU10>='" & frm100116_1.txt1(0) & "' or cu4.CU10>='" & frm100116_1.txt1(0) & "' or cu5.CU10>='" & frm100116_1.txt1(0) & "')"
         StrSQL4 = StrSQL4 & " AND (cu1.CU10>='" & frm100116_1.txt1(0) & "' or cu2.CU10>='" & frm100116_1.txt1(0) & "' or cu3.CU10>='" & frm100116_1.txt1(0) & "' or cu4.CU10>='" & frm100116_1.txt1(0) & "' or cu5.CU10>='" & frm100116_1.txt1(0) & "')"
         '2014/12/25 END
         strSQL5 = strSQL5 & " AND (cu1.CU10>='" & frm100116_1.txt1(0) & "' or cu2.CU10>='" & frm100116_1.txt1(0) & "' or cu3.CU10>='" & frm100116_1.txt1(0) & "' or cu4.CU10>='" & frm100116_1.txt1(0) & "' or cu5.CU10>='" & frm100116_1.txt1(0) & "')"
      End If
      If Len(Trim(frm100116_1.txt1(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND (cu1.CU10<='" & frm100116_1.txt1(1) & "z' or cu2.CU10<='" & frm100116_1.txt1(1) & "z' or cu3.CU10<='" & frm100116_1.txt1(1) & "z' or cu4.CU10<='" & frm100116_1.txt1(1) & "z' or cu5.CU10<='" & frm100116_1.txt1(1) & "z') "
         strSQL2 = strSQL2 & " AND (cu1.CU10<='" & frm100116_1.txt1(1) & "z' or cu2.CU10<='" & frm100116_1.txt1(1) & "z' or cu3.CU10<='" & frm100116_1.txt1(1) & "z' or cu4.CU10<='" & frm100116_1.txt1(1) & "z' or cu5.CU10<='" & frm100116_1.txt1(1) & "z') "
         'Modify By Sindy 2014/12/25
         'StrSQL3 = StrSQL3 & " AND CU10<='" & frm100116_1.txt1(1) & "z' "
         'StrSQL4 = StrSQL4 & " AND CU10<='" & frm100116_1.txt1(1) & "z' "
         StrSQL3 = StrSQL3 & " AND (cu1.CU10<='" & frm100116_1.txt1(1) & "z' or cu2.CU10<='" & frm100116_1.txt1(1) & "z' or cu3.CU10<='" & frm100116_1.txt1(1) & "z' or cu4.CU10<='" & frm100116_1.txt1(1) & "z' or cu5.CU10<='" & frm100116_1.txt1(1) & "z') "
         StrSQL4 = StrSQL4 & " AND (cu1.CU10<='" & frm100116_1.txt1(1) & "z' or cu2.CU10<='" & frm100116_1.txt1(1) & "z' or cu3.CU10<='" & frm100116_1.txt1(1) & "z' or cu4.CU10<='" & frm100116_1.txt1(1) & "z' or cu5.CU10<='" & frm100116_1.txt1(1) & "z') "
         '2014/12/25 END
         strSQL5 = strSQL5 & " AND (cu1.CU10<='" & frm100116_1.txt1(1) & "z' or cu2.CU10<='" & frm100116_1.txt1(1) & "z' or cu3.CU10<='" & frm100116_1.txt1(1) & "z' or cu4.CU10<='" & frm100116_1.txt1(1) & "z' or cu5.CU10<='" & frm100116_1.txt1(1) & "z') "
      End If
End If
If Len(Trim(frm100116_1.txt1(0))) <> 0 Or Len(Trim(frm100116_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100116_1.Option1(0).Caption & frm100116_1.txt1(0) & "-" & frm100116_1.txt1(1) 'Add By Sindy 2010/11/15
End If
If frm100116_1.ChkPCT.Value = vbChecked Then
   pub_QL05 = pub_QL05 & ";顯示PCT 案" 'Add By Sindy 2010/11/15
End If
'Modify By Cheng 2002/04/26
'若已閉卷, 在本所案號後加"*"號
'edit by nick 2004/07/29 因為太慢了
'edit by nickc 2007/03/23 加入PCT 欄
'2010/9/15 MODIFY BY SONIA 所有日期欄若需排序改百年日期排序問題
'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
         strSql = "SELECT distinct '' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)) AS 申請人,TM15 AS 審定號或證書號數,TM12 AS 申請案號,NVL(NA03,TM10) AS 申請國家,SUBSTR(' '||sqldatet(TM11),-9) AS 申請日,'','','','','','','',TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort,'' FROM TRADEMARK,CASEPROGRESS,CUSTOMER cu1,CUSTOMER cu2,CUSTOMER cu3,CUSTOMER cu4,CUSTOMER cu5,NATION WHERE TM01=CP01(+) and TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND " & SQLNewFag("TM23", "cu1.CU") & " AND " & SQLNewFag("tm78", "CU2.CU") & " AND " & SQLNewFag("tm79", "CU3.CU") & " AND " & SQLNewFag("tm80", "CU4.CU") & " AND " & SQLNewFag("tm81", "CU5.CU") & " AND TM10=NA01(+) " & strSQL2 & StrSQL6
strSql = strSql + " union  select '' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)) AS 申請人,PA22 AS 審定號或證書號數,PA11 AS 申請案號,NVL(NA03,PA09) AS 申請國家,SUBSTR(' '||sqldatet(PA10),-9) AS 申請日,'','','','','','','',PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort,pa46 FROM PATENT,CASEPROGRESS,CUSTOMER cu1,CUSTOMER cu2,CUSTOMER cu3,CUSTOMER cu4,CUSTOMER cu5,NATION WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & SQLNewFag("PA30", "CU5.CU") & " AND " & SQLNewFag("PA26", "CU1.CU") & " AND " & SQLNewFag("PA27", "CU2.CU") & " AND " & SQLNewFag("PA28", "CU3.CU") & " AND " & SQLNewFag("PA29", "CU4.CU") & " AND PA09=NA01(+) " & strSQL1 & StrSQL6
strSql = strSql + " union  select '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)) AS 申請人,SP14 AS 審定號或證書號數,SP11 AS 申請案號,NVL(NA03,SP09) AS 申請國家,SUBSTR(' '||sqldatet(SP10),-9) AS 申請日,'','','','','','','',SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort,'' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER cu1,CUSTOMER cu2,CUSTOMER cu3,CUSTOMER cu4,CUSTOMER cu5,NATION WHERE SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) AND " & SQLNewFag("SP08", "CU1.CU") & " AND " & SQLNewFag("SP58", "CU2.CU") & " AND " & SQLNewFag("SP59", "CU3.CU") & " AND " & SQLNewFag("sp65", "CU4.CU") & " AND " & SQLNewFag("sp66", "CU5.CU") & " AND SP09=NA01(+) " & strSQL5 & StrSQL6
strSql = strSql + " union  select '' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NVL(NA03,LC15) AS 申請國家,'' AS 申請日,'','','','','','','',LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort,'' FROM LAWCASE,CASEPROGRESS,CUSTOMER cu1,CUSTOMER cu2,CUSTOMER cu3,CUSTOMER cu4,CUSTOMER cu5,NATION WHERE LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) AND " & SQLNewFag("LC11", "cu1.CU") & " AND " & SQLNewFag("LC43", "CU2.CU") & " AND " & SQLNewFag("LC44", "CU3.CU") & " AND " & SQLNewFag("LC45", "CU4.CU") & " AND " & SQLNewFag("LC46", "CU5.CU") & " AND LC15=NA01(+) " & StrSQL3 & StrSQL6
strSql = strSql + " union  select '' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NA03 AS 申請國家,'' AS 申請日,'','','','','','','',HC01||'-'||HC02||'-'||HC03||'-'||HC04 as FSort,'' FROM HIRECASE,CASEPROGRESS,CUSTOMER cu1,CUSTOMER cu2,CUSTOMER cu3,CUSTOMER cu4,CUSTOMER cu5,NATION WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND " & SQLNewFag("HC05", "cu1.CU") & " AND " & SQLNewFag("HC24", "CU2.CU") & " AND " & SQLNewFag("HC25", "CU3.CU") & " AND " & SQLNewFag("HC26", "CU4.CU") & " AND " & SQLNewFag("HC27", "CU5.CU") & " AND '000'=NA01(+) " & StrSQL4 & StrSQL6

strSql = strSql + " ORDER BY FSort,本所案號 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/15
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/15
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True
End Sub

Sub StrMenu2()          '申請國家
Me.Enabled = False
'DoEvents
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
If Trim(frm100116_1.txt1(4)) = "1" Then
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP05>=" & Val(ChangeTStringToWString(frm100116_1.txt1(5))) & " "
   End If
   If Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(frm100116_1.txt1(6))) & " "
   'Add By Cheng 2002/03/18
   Else
      If Len(frm100116_1.txt1(5).Text) > 0 Then
         StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
      End If
   End If
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Or Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & frm100116_1.Label3 & frm100116_1.txt1(5) & "-" & frm100116_1.txt1(6) 'Add By Sindy 2010/11/15
   End If
Else
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP27>=" & Val(ChangeTStringToWString(frm100116_1.txt1(5))) & " "
   End If
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(frm100116_1.txt1(6))) & " "
   'Add By Cheng 2002/03/18
   Else
      If Len(frm100116_1.txt1(5).Text) > 0 Then
         StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
      End If
   End If
   If Len(Trim(frm100116_1.txt1(5))) <> 0 Or Len(Trim(frm100116_1.txt1(6))) <> 0 Then
      pub_QL05 = pub_QL05 & ";發文" & frm100116_1.Label3 & frm100116_1.txt1(5) & "-" & frm100116_1.txt1(6) 'Add By Sindy 2010/11/15
   End If
End If
If Len(Trim(frm100116_1.txt1(7))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(frm100116_1.txt1(7), 1) & ") "
'   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(frm100116_1.txt1(7), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(frm100116_1.txt1(7), 3) & ") "
'   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(frm100116_1.txt1(7), 4) & ") "
'   StrSQL5 = StrSQL5 & " AND CP01 IN (" & SQLGrpStr(frm100116_1.txt1(7), 5) & ") "
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 1) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100116_1.txt1(7).Text <> "ALL", frm100116_1.txt1(7).Text, GetAllSysKind(frm100116_1.txt1(7))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100116_1.Label4, 5) & frm100116_1.txt1(7) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100116_1.txt1(8))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10>='" & frm100116_1.txt1(8) & "' "
End If
If Len(Trim(frm100116_1.txt1(9))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10<='" & frm100116_1.txt1(9) & "' "
End If
If Len(Trim(frm100116_1.txt1(8))) <> 0 Or Len(Trim(frm100116_1.txt1(9))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100116_1.Label5(0) & frm100116_1.txt1(8) & "-" & frm100116_1.txt1(9) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100116_1.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA09>='" & frm100116_1.txt1(2) & "' "
   strSQL2 = strSQL2 & " AND TM10>='" & frm100116_1.txt1(2) & "' "
   StrSQL3 = StrSQL3 & " AND LC15>='" & frm100116_1.txt1(2) & "' "
   'StrSQL4 = StrSQL4 & " AND CU10>='" & frm100116_1.txt1(02) & "' "
   strSQL5 = strSQL5 & " AND SP09>='" & frm100116_1.txt1(2) & "' "
End If
If Len(Trim(frm100116_1.txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA09<='" & frm100116_1.txt1(3) & "' "
   strSQL2 = strSQL2 & " AND TM10<='" & frm100116_1.txt1(3) & "' "
   StrSQL3 = StrSQL3 & " AND LC15<='" & frm100116_1.txt1(3) & "' "
   'StrSQL4 = StrSQL4 & " AND CU10<='" & frm100116_1.txt1(3) & "' "
   strSQL5 = strSQL5 & " AND SP09<='" & frm100116_1.txt1(3) & "' "
End If
If Len(Trim(frm100116_1.txt1(2))) <> 0 Or Len(Trim(frm100116_1.txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100116_1.Option1(1).Caption & frm100116_1.txt1(2) & "-" & frm100116_1.txt1(3) 'Add By Sindy 2010/11/15
End If
If frm100116_1.ChkPCT.Value = vbChecked Then
   pub_QL05 = pub_QL05 & ";顯示PCT 案" 'Add By Sindy 2010/11/15
End If
'Modify By Cheng 2002/04/26
'若已閉卷, 則在本所案號後加"*"號
'Modify by Morgan 2004/7/23
'加CP05>19110000條件，以免跑不出來
'strSQL = "SELECT distinct '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,TM15 AS 審定號或證書號數,TM12 AS 申請案號,NVL(NA03,TM10) AS 申請國家,SUBSTR(' '||sqldatet(TM11),-9) AS 申請日,0,0,'','','','','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND CP01=TM01(+) and CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) " & strSQL2 & StrSQL6
'strSQL = strSQL + " union all select '' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,PA22 AS 審定號或證書號數,PA11 AS 申請案號,NVL(NA03,PA09) AS 申請國家,SUBSTR(' '||sqldatet(PA10),-9) AS 申請日,0,0,'','','','','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,SP14 AS 審定號或證書號數,SP11 AS 申請案號,NVL(NA03,SP09) AS 申請國家,SUBSTR(' '||sqldatet(SP10),-9) AS 申請日,0,0,'','','','','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5 & StrSQL6
'strSQL = strSQL + " union all select '' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NVL(NA03,LC15) AS 申請國家,'' AS 申請日,0,0,'','','','','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND " & SQLNewFag("LC11", "CU") & " AND LC15=NA01(+) " & StrSQL3 & StrSQL6
'If frm100116_1.txt1(2) <= "000" And frm100116_1.txt1(3) >= "000" Then
'   strSQL = strSQL + " union all select '' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NA03 AS 申請國家,'' AS 申請日,CP05,CP27,CP10,'','','','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND " & SQLNewFag("HC05", "CU") & " AND '000'=NA01(+) " & StrSQL4 & StrSQL6
'End If
'edit by nick 2004/07/29 因為太慢了將由基本檔為主
'edit by nickc 2005/05/13
'strSQL = "SELECT distinct '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,TM15 AS 審定號或證書號數,TM12 AS 申請案號,NVL(NA03,TM10) AS 申請國家,SUBSTR(' '||sqldatet(TM11),-9) AS 申請日,0,0,'','','','','' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND TM01=CP01(+) and TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) " & strSQL2 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,PA22 AS 審定號或證書號數,PA11 AS 申請案號,NVL(NA03,PA09) AS 申請國家,SUBSTR(' '||sqldatet(PA10),-9) AS 申請日,0,0,'','','','','' FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,SP14 AS 審定號或證書號數,SP11 AS 申請案號,NVL(NA03,SP09) AS 申請國家,SUBSTR(' '||sqldatet(SP10),-9) AS 申請日,0,0,'','','','','' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NVL(NA03,LC15) AS 申請國家,'' AS 申請日,0,0,'','','','','' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) AND " & SQLNewFag("LC11", "CU") & " AND LC15=NA01(+) " & StrSQL3 & StrSQL6
'If frm100116_1.txt1(2) <= "000" And frm100116_1.txt1(3) >= "000" Then
'   strSQL = strSQL + " union  select '' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NA03 AS 申請國家,'' AS 申請日,CP05,CP27,CP10,'','','','' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND " & SQLNewFag("HC05", "CU") & " AND '000'=NA01(+) " & StrSQL4 & StrSQL6
'End If
'strSQL = strSQL + " ORDER BY 本所案號 "
'edit by nickc 2007/03/23 加入 PCT 欄
'strSQL = "SELECT distinct '' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,TM15 AS 審定號或證書號數,TM12 AS 申請案號,NVL(NA03,TM10) AS 申請國家,SUBSTR(' '||sqldatet(TM11),-9) AS 申請日,0,0,'','','','','',TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND TM01=CP01(+) and TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) " & strSQL2 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,PA22 AS 審定號或證書號數,PA11 AS 申請案號,NVL(NA03,PA09) AS 申請國家,SUBSTR(' '||sqldatet(PA10),-9) AS 申請日,0,0,'','','','','',PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,SP14 AS 審定號或證書號數,SP11 AS 申請案號,NVL(NA03,SP09) AS 申請國家,SUBSTR(' '||sqldatet(SP10),-9) AS 申請日,0,0,'','','','','',SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5 & StrSQL6
'strSQL = strSQL + " union  select '' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NVL(NA03,LC15) AS 申請國家,'' AS 申請日,0,0,'','','','','',LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) AND " & SQLNewFag("LC11", "CU") & " AND LC15=NA01(+) " & StrSQL3 & StrSQL6
'If frm100116_1.txt1(2) <= "000" And frm100116_1.txt1(3) >= "000" Then
'   strSQL = strSQL + " union  select '' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NA03 AS 申請國家,'' AS 申請日,CP05,CP27,CP10,'','','','',HC01||'-'||HC02||'-'||HC03||'-'||HC04 as FSort FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND " & SQLNewFag("HC05", "CU") & " AND '000'=NA01(+) " & StrSQL4 & StrSQL6
'End If
strSql = "SELECT distinct '' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人,TM15 AS 審定號或證書號數,TM12 AS 申請案號,NVL(NA03,TM10) AS 申請國家,SUBSTR(' '||sqldatet(TM11),-9) AS 申請日,0,0,'','','','','',TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort,'' FROM TRADEMARK,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND TM01=CP01(+) and TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) " & strSQL2 & StrSQL6
strSql = strSql + " union  select '' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人,PA22 AS 審定號或證書號數,PA11 AS 申請案號,NVL(NA03,PA09) AS 申請國家,SUBSTR(' '||sqldatet(PA10),-9) AS 申請日,0,0,'','','','','',PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort,pa46 FROM PATENT,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) " & strSQL1 & StrSQL6
strSql = strSql + " union  select '' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人,SP14 AS 審定號或證書號數,SP11 AS 申請案號,NVL(NA03,SP09) AS 申請國家,SUBSTR(' '||sqldatet(SP10),-9) AS 申請日,0,0,'','','','','',SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort,'' FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5 & StrSQL6
strSql = strSql + " union  select '' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NVL(NA03,LC15) AS 申請國家,'' AS 申請日,0,0,'','','','','',LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort,'' FROM LAWCASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) AND " & SQLNewFag("LC11", "CU") & " AND LC15=NA01(+) " & StrSQL3 & StrSQL6
If frm100116_1.txt1(2) <= "000" And frm100116_1.txt1(3) >= "000" Then
   strSql = strSql + " union  select '' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人,'' AS 審定號或證書號數,'' AS 申請案號,NA03 AS 申請國家,'' AS 申請日,CP05,CP27,CP10,'','','','',HC01||'-'||HC02||'-'||HC03||'-'||HC04 as FSort,'' FROM HIRECASE,CASEPROGRESS,CUSTOMER,NATION WHERE CP05>19110000 AND HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND " & SQLNewFag("HC05", "CU") & " AND '000'=NA01(+) " & StrSQL4 & StrSQL6
End If
strSql = strSql + " ORDER BY FSort,本所案號 "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/15
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/15
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    'Modify By Cheng 2003/07/30
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100116_2 = Nothing
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
