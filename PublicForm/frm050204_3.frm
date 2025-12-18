VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050204_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件性質統計"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   990
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8385
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   0
      Left            =   4410
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   5940
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7170
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5160
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9102
      _Version        =   393216
      Cols            =   15
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
      _Band(0).Cols   =   15
   End
End
Attribute VB_Name = "frm050204_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL6 As String
Dim strSql As String, i As Integer, j As Integer, strTemp As Variant, strTemp1 As Variant, s As Integer
Dim StrTag As String, intK As Integer
Public cmdState As Integer 'Add by Amy 2018/02/27 記錄按鈕

Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
grdDataList.ColWidth(2) = 1800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(3) = 1800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件性質"
grdDataList.ColWidth(4) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "承辦人"
grdDataList.ColWidth(5) = 900
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "智權人員"
grdDataList.ColWidth(6) = 900
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "發文日"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "申請人"
grdDataList.ColWidth(8) = 900
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "申請國家"
grdDataList.ColWidth(9) = 900
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdok_Click(Index As Integer)
'Modify by Amy 2018/02/27 按了下列按鈕會造成無法關閉表單,故改寫法
cmdState = Index
PubShowNextData
Exit Sub

'Mark by Amy 2018/02/27 改寫至PubShowNextData
'Select Case Index
'Case 0
'      Me.Enabled = False
'      For i = 1 To grdDataList.Rows - 1
'      grdDataList.col = 0
'      grdDataList.row = i
'      If Trim(grdDataList.Text) = "V" Then
'        Dim Str01 As String
'        grdDataList.col = 2
'        Str01 = SystemNumber(grdDataList, 1)
'        If Mid(UCase(Str01), 1, 1) = "N" Then
'            Str01 = Mid(Str01, 2, 3)
'        End If
'        If Not IsNull(grdDataList.Text) Then
'        Select Case Replace(Replace(Replace(Replace(Replace(Replace(Str01, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
'            Case "CFP", "FCP", "P"   '專利
'                  Screen.MousePointer = vbHourglass
'                  frm100101_3.Show
'                  'frm100101_3.Hide
'                  'Modify By Cheng 2002/06/25
''                  frm100101_3.Tag = grdDataList.Text
'                  frm100101_3.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "") ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
'                  frm100101_3.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Hide
'                  'frm100101_3.Show
'                  Do
'                  DoEvents
'                  If bolToEndByNick = True Then Unload Me: Exit Sub
'                  Loop Until Not frm100101_3.Visible
'                  Unload frm100101_3
'            Case "CFT", "FCT", "T", "TF"   '商標
'                  Screen.MousePointer = vbHourglass
'                  frm100101_4.Show
'                  'frm100101_4.Hide
'                  'Modify By Cheng 2002/06/25
''                  frm100101_4.Tag = grdDataList.Text
'                  frm100101_4.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "") ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
'                  frm100101_4.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Hide
'                  'frm100101_4.Show
'                  Do
'                  DoEvents
'                  If bolToEndByNick = True Then Unload Me: Exit Sub
'                  Loop Until Not frm100101_4.Visible
'                  Unload frm100101_4
'            'Modify By Sindy 2009/07/24 增加LIN系統類別
'            Case "CFL", "FCL", "L", "LIN"          '法務
'                  Screen.MousePointer = vbHourglass
'                  frm100101_5.Show
'                  'frm100101_5.Hide
'                  'Modify By Cheng 2002/06/25
''                  frm100101_5.Tag = grdDataList.Text
'                  frm100101_5.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "") ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
'                  frm100101_5.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Hide
'                  'frm100101_5.Show
'                  Do
'                  DoEvents
'                  If bolToEndByNick = True Then Unload Me: Exit Sub
'                  Loop Until Not frm100101_5.Visible
'                  Unload frm100101_5
''            Case "LA"            '顧問
''                  Screen.MousePointer = vbHourglass
''                  frm100101_6.Show
''                  'frm100101_6.Hide
''                  'Modify By Cheng 2002/06/25
'''                  frm100101_6.Tag = grdDataList.Text
''                  frm100101_6.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                  frm100101_6.StrMenu
''                  Screen.MousePointer = vbDefault
''                  Me.Hide
''                  'frm100101_6.Show
''                  Do
''                  DoEvents
''                  If bolToEndByNick = True Then Unload Me: Exit Sub
''                  Loop Until Not frm100101_6.Visible
''                  Unload frm100101_6
''            Case Else                  '服務
''                 select case Replace(Replace(Replace(Replace(Replace(Replace(str01, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
''                     Case "TB"    '條碼
''                         Screen.MousePointer = vbHourglass
''                        frm100101_7.Show
''                        'frm100101_7.Hide
''                        'Modify By Cheng 2002/06/25
'''                        frm100101_7.Tag = grdDataList.Text
''                        frm100101_7.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                        frm100101_7.StrMenu
''                        Screen.MousePointer = vbDefault
''                        Me.Hide
''                        'frm100101_7.Show
''                        Do
''                         DoEvents
''                         If bolToEndByNick = True Then Unload Me: Exit Sub
''                         Loop Until Not frm100101_7.Visible
''                         Unload frm100101_7
''                     Case "TM"
''                         Screen.MousePointer = vbHourglass
''                        frm100101_8.Show
''                        'frm100101_8.Hide
''                        'Modify By Cheng 2002/06/25
'''                        frm100101_8.Tag = grdDataList.Text
''                        frm100101_8.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                        frm100101_8.StrMenu
''                        Screen.MousePointer = vbDefault
''                        Me.Hide
''                        'frm100101_8.Show
''                         Do
''                         DoEvents
''                         If bolToEndByNick = True Then Unload Me: Exit Sub
''                         Loop Until Not frm100101_8.Visible
''                         Unload frm100101_8
''                     Case "TD"
''                         Screen.MousePointer = vbHourglass
''                        frm100101_9.Show
''                        'frm100101_9.Hide
''                        'Modify By Cheng 2002/06/25
'''                        frm100101_9.Tag = grdDataList.Text
''                        frm100101_9.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                        frm100101_9.StrMenu
''                        Screen.MousePointer = vbDefault
''                        Me.Hide
''                        'frm100101_9.Show
''                         Do
''                         DoEvents
''                         If bolToEndByNick = True Then Unload Me: Exit Sub
''                         Loop Until Not frm100101_9.Visible
''                         Unload frm100101_9
''                     Case "TC", "CFC"
''                         Screen.MousePointer = vbHourglass
''                        frm100101_A.Show
''                        'frm100101_A.Hide
''                        'Modify By Cheng 2002/06/25
'''                        frm100101_A.Tag = grdDataList.Text
''                        frm100101_A.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                        frm100101_A.StrMenu
''                        Screen.MousePointer = vbDefault
''                        Me.Hide
''                        'frm100101_A.Show
''                         Do
''                         DoEvents
''                         If bolToEndByNick = True Then Unload Me: Exit Sub
''                         Loop Until Not frm100101_A.Visible
''                         Unload frm100101_A
''                     Case Else
''                         Screen.MousePointer = vbHourglass
''                        frm100101_B.Show
''                        'frm100101_B.Hide
''                        'Modify By Cheng 2002/06/25
'''                        frm100101_B.Tag = grdDataList.Text
''                        frm100101_B.Tag = replace(replace(replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""),"＊",""),"v",""),"V","")     ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "")
''                        frm100101_B.StrMenu
''                        Screen.MousePointer = vbDefault
''                        Me.Hide
''                        'frm100101_B.Show
''                         Do
''                         DoEvents
''                         If bolToEndByNick = True Then Unload Me: Exit Sub
''                         Loop Until Not frm100101_B.Visible
''                         Unload frm100101_B
''                  End Select
'        End Select
'        End If
'        grdDataList.col = 0
'        grdDataList.Text = ""
'         For j = 0 To grdDataList.Cols - 1
'             grdDataList.col = j
'            grdDataList.CellBackColor = QBColor(15)
'         Next j
'     End If
'     Next i
'     Me.Enabled = True
'     Me.Show
'Case 1
'     Me.Enabled = False
'     StrTag = ""
'     For i = 1 To grdDataList.Rows - 1
'     grdDataList.col = 0
'     grdDataList.row = i
'     If Trim(grdDataList.Text) = "V" Then
'         grdDataList.col = 2
'         If Not IsNull(grdDataList.Text) Then
'            Screen.MousePointer = vbHourglass
'            frm100101_2.Show
'            'frm100101_2.Hide
'            'Modify By Cheng 2002/06/25
''            frm100101_2.Tag = grdDataList.Text ' StrTag
'            frm100101_2.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "") ' edit by nickc 2005/08/10       Replace(grdDataList.Text, "＊", "") ' StrTag
'            frm100101_2.StrMenu
'            Screen.MousePointer = vbDefault
'            Me.Hide
'            'frm100101_2.Show
'            Do
'            DoEvents
'            If bolToEndByNick = True Then Unload Me: Exit Sub
'            Loop Until Not frm100101_2.Visible
'            Unload frm100101_2
'            grdDataList.col = 0
'            grdDataList.Text = ""
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = QBColor(15)
'            Next j
'
'         End If
'     End If
'     Next i
'     Me.Enabled = True
'     Me.Show
'Case 2
'     Me.Hide
'Case 3
'     bolToEndByNick = True
'     Unload Me
'     Exit Sub
'Case Else
'End Select
End Sub

'Add by Amy 2017/02/27
Public Sub PubShowNextData()
    Select Case cmdState
        '案件基本資料/案件進度
        Case 0, 1
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
                        '案件基本資料
                        If cmdState = 0 Then
                            Dim Str01 As String
                            Str01 = SystemNumber(grdDataList, 1)
                            If Mid(UCase(Str01), 1, 1) = "N" Then
                                Str01 = Mid(Str01, 2, 3)
                            End If
                            Select Case Replace(Replace(Replace(Replace(Replace(Replace(Str01, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
                                Case "CFP", "FCP", "P"   '專利
                                      frm100101_3.Show
                                      frm100101_3.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
                                      frm100101_3.StrMenu
                                Case "CFT", "FCT", "T", "TF"   '商標
                                      Screen.MousePointer = vbHourglass
                                      frm100101_4.Show
                                      frm100101_4.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
                                      frm100101_4.StrMenu
                                'modify by sonia 2019/7/29 +ACS系統類別
                                Case "CFL", "FCL", "L", "LIN", "ACS" '法務
                                      Screen.MousePointer = vbHourglass
                                      frm100101_5.Show
                                      frm100101_5.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
                                      frm100101_5.StrMenu
                            End Select
                        '案件進度
                        Else
                            StrTag = ""
                            frm100101_2.Show
                            frm100101_2.Tag = Replace(Replace(Replace(Replace(Replace(Replace(grdDataList.Text, "*", ""), "△", ""), "N", ""), "＊", ""), "v", ""), "V", "")
                            frm100101_2.StrMenu
                        End If
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                    End If
                    
                End If
            
            Next i
            Me.Enabled = True
        '回前畫面
        Case 2
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
        '結束
        Case 3
            fnCloseAllFrm100
    End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   cmdState = -1 'Add by Amy 2018/02/27
End Sub
Sub StrMenu()

Me.Enabled = False
'只有專利,商標,法務
'檢查收發文
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL6 = ""
'系統類別
If Len(frm050204_1.txt1(10)) <> 0 Then
   strSQL1 = strSQL1 + " and cp01 in (" & SQLGrpStr(frm050204_1.txt1(10), 1) & ") "
   strSQL2 = strSQL2 + " and cp01 in (" & SQLGrpStr(frm050204_1.txt1(10), 2) & ") "
   StrSQL3 = StrSQL3 + " and cp01 in (" & SQLGrpStr(frm050204_1.txt1(10), 3) & ") "
End If
'無取消收文日
strSQL1 = strSQL1 + " AND CP57 IS NULL "
strSQL2 = strSQL2 + " AND CP57 IS NULL "
StrSQL3 = StrSQL3 + " AND CP57 IS NULL "
'收文
If frm050204_1.txt1(4) = "1" Then
   If Len(Trim(frm050204_1.txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
   End If
   If Len(Trim(frm050204_1.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
   End If
'發文
Else
   If Len(Trim(frm050204_1.txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.txt1(5))) & " "
   End If
   If Len(Trim(frm050204_1.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.txt1(6))) & " "
   End If
End If
'申請國家
If Len(Trim(frm050204_1.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 + " and PA09>='" & frm050204_1.txt1(0) & "' "
   strSQL2 = strSQL2 + " and TM10>='" & frm050204_1.txt1(0) & "' "
   StrSQL3 = StrSQL3 + " and LC15>='" & frm050204_1.txt1(0) & "' "
End If
If Len(Trim(frm050204_1.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " and PA09<='" & frm050204_1.txt1(1) & "' "
   strSQL2 = strSQL2 & " and TM10<='" & frm050204_1.txt1(1) & "' "
   StrSQL3 = StrSQL3 & " and LC15<='" & frm050204_1.txt1(1) & "' "
End If
'代理人國籍
If Len(Trim(frm050204_1.txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 + " and fa10>='" & frm050204_1.txt1(2) & "' "
End If
If Len(Trim(frm050204_1.txt1(3))) <> 0 Then
   StrSQL6 = StrSQL6 & " and fa10<='" & frm050204_1.txt1(3) & "z' "
End If
'代理人
If Len(frm050204_1.txt1(7)) <> 0 Then
    strSQL1 = strSQL1 & " and decode(pa09,'000',pa75,cp44)='" & GetNewFagent(frm050204_1.txt1(7)) & "' "
    strSQL2 = strSQL2 & " and decode(tm10,'000',tm44,cp44)='" & GetNewFagent(frm050204_1.txt1(7)) & "' "
    StrSQL3 = StrSQL3 & " and decode(lc15,'000',lc22,cp44)='" & GetNewFagent(frm050204_1.txt1(7)) & "' "
End If
strSQL1 = strSQL1 & " and decode(pa09,'000',pa75,cp44)='" & GetNewFagent(frm050204_3.Tag) & "' "
strSQL2 = strSQL2 & " and decode(tm10,'000',tm44,cp44)='" & GetNewFagent(frm050204_3.Tag) & "' "
StrSQL3 = StrSQL3 & " and decode(lc15,'000',lc22,cp44)='" & GetNewFagent(frm050204_3.Tag) & "' "

'案件性質
If Len(frm050204_1.txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10>='" & frm050204_1.txt1(8) & "' "
    strSQL2 = strSQL2 + " and cp10>='" & frm050204_1.txt1(8) & "' "
    StrSQL3 = StrSQL3 + " and cp10>='" & frm050204_1.txt1(8) & "' "
End If
If Len(frm050204_1.txt1(9)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10<='" & frm050204_1.txt1(9) & "' "
    strSQL2 = strSQL2 + " and cp10<='" & frm050204_1.txt1(9) & "' "
    StrSQL3 = StrSQL3 + " and cp10<='" & frm050204_1.txt1(9) & "' "
End If
strSQL1 = strSQL1 + " and cp10='" & frm050204_3.grdDataList.Tag & "' "
strSQL2 = strSQL2 + " and cp10='" & frm050204_3.grdDataList.Tag & "' "
StrSQL3 = StrSQL3 + " and cp10='" & frm050204_3.grdDataList.Tag & "' "
'Modify By Sindy 2011/2/16 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
'strSql = "select ' ' AS V," & SQLDate("NEW.CP05") & " AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.PA57,'Y','＊','') AS 本所案號,NVL(NEW.PA05,NVL(NEW.PA06,NEW.PA07)) AS 案件名稱,nvl(DECODE(NEW.PA09,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員," & SQLDate("NEW.CP27") & " AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
'          "from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,PA05,PA06,PA07,PA09,PA57,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, patent   ,CasePropertyMap where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
'strSql = strSql & " UNION ALL select ' ' AS V," & SQLDate("NEW.CP05") & " AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.TM29,'Y','＊','') AS 本所案號,NVL(NEW.TM05,NVL(NEW.TM06,NEW.TM07)) AS 案件名稱,nvl(DECODE(NEW.TM10,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員," & SQLDate("NEW.CP27") & " AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
'          "from (select substr(decode(TM10,'000',TM44,cp44),1,8) as a,decode(substr(decode(TM10,'000',TM44,cp44),9,1),'','0',substr(decode(TM10,'000',TM44,cp44),9,1)) as b,TM05,TM06,TM07,TM10,TM29,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, TRADEMARK   ,CasePropertyMap where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
'strSql = strSql & " UNION ALL select ' ' AS V," & SQLDate("NEW.CP05") & " AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.LC08,'Y','＊','') AS 本所案號,NVL(NEW.LC05,NVL(NEW.LC06,NEW.LC07)) AS 案件名稱,nvl(DECODE(NEW.LC15,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員," & SQLDate("NEW.CP27") & " AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
'          "from (select substr(decode(LC15,'000',LC22,cp44),1,8) as a,decode(substr(decode(LC15,'000',LC22,cp44),9,1),'','0',substr(decode(LC15,'000',LC22,cp44),9,1)) as b,LC05,LC06,LC07,LC15,LC08,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, LAWCASE   ,CasePropertyMap where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
strSql = "select ' ' AS V,sqldatet2(NEW.CP05) AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.PA57,'Y','＊','') AS 本所案號,NVL(NEW.PA05,NVL(NEW.PA06,NEW.PA07)) AS 案件名稱,nvl(DECODE(NEW.PA09,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員,sqldatet2(NEW.CP27) AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
          "from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,PA05,PA06,PA07,PA09,PA57,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, patent   ,CasePropertyMap where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
strSql = strSql & " UNION ALL select ' ' AS V,sqldatet2(NEW.CP05) AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.TM29,'Y','＊','') AS 本所案號,NVL(NEW.TM05,NVL(NEW.TM06,NEW.TM07)) AS 案件名稱,nvl(DECODE(NEW.TM10,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員,sqldatet2(NEW.CP27) AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
          "from (select substr(decode(TM10,'000',TM44,cp44),1,8) as a,decode(substr(decode(TM10,'000',TM44,cp44),9,1),'','0',substr(decode(TM10,'000',TM44,cp44),9,1)) as b,TM05,TM06,TM07,TM10,TM29,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, TRADEMARK   ,CasePropertyMap where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
strSql = strSql & " UNION ALL select ' ' AS V,sqldatet2(NEW.CP05) AS 收文日,NEW.CP01||'-'||NEW.CP02||'-'||NEW.CP03||'-'||NEW.CP04||DECODE(NEW.LC08,'Y','＊','') AS 本所案號,NVL(NEW.LC05,NVL(NEW.LC06,NEW.LC07)) AS 案件名稱,nvl(DECODE(NEW.LC15,'000',NEW.cpm03,NEW.cpm04),NEW.cp10) AS 案件性質,nvl(s1.st02,NEW.CP14) AS 承辦人,nvl(s2.st02,NEW.CP13) AS 智權人員,sqldatet2(NEW.CP27) AS 發文日,nvl(NA03,NA04) As 申請國家 " & _
          "from (select substr(decode(LC15,'000',LC22,cp44),1,8) as a,decode(substr(decode(LC15,'000',LC22,cp44),9,1),'','0',substr(decode(LC15,'000',LC22,cp44),9,1)) as b,LC05,LC06,LC07,LC15,LC08,CP01,CP02,CP03,CP04,CP05,CP10,CP13,CP14,CP27,cpm03,CPM04 from caseprogress, LAWCASE   ,CasePropertyMap where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & ") new,fagent,nation,SystemKind,STAFF S1,STAFF S2 where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL AND NEW.CP14=S1.ST01(+) AND NEW.CP13=S2.ST01(+) " & StrSQL6
If Trim(frm050204_1.txt1(4).Text) = "1" Then
    strSql = strSql + " ORDER BY 收文日,本所案號"
Else
    strSql = strSql + " ORDER BY 發文日,本所案號"
End If
CheckOC
Dim StrTest1 As String, StrTest2 As String
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   cmdok(0).Enabled = True
   cmdok(1).Enabled = True
Else
   ShowNoData
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Me.Hide
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
CheckOC
Me.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Modified by Lydia 2018/05/02
'Set frm050204_2 = Nothing
Set frm050204_3 = Nothing
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

