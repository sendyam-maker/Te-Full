VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090218_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "英文核稿查詢"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7560
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   4065
      Left            =   60
      TabIndex        =   3
      Top             =   870
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "暫停核稿記錄(&S)"
      Height          =   375
      Index           =   0
      Left            =   4710
      TabIndex        =   0
      Top             =   30
      Width           =   1515
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1230
      TabIndex        =   5
      Top             =   540
      Width           =   2790
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4921;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LBL1 
      Height          =   255
      Left            =   4590
      TabIndex        =   4
      Top             =   570
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核稿人："
      Height          =   180
      Index           =   25
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frm090218_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; Grd1改字型=新細明體-ExtB、Combo1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim i As Integer, strSql As String, ADORECORDSET66 As New ADODB.Recordset
Public cmdState As Integer

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Select Case cmdState
Case 0
     Me.Enabled = False
     For i = 1 To Grd1.Rows - 1
         Grd1.col = 0
         Grd1.row = i
         If Trim(Grd1.Text) = "V" Then
            Grd1.col = 0
            Grd1.Text = ""
            Me.Hide
            If Grd1.CellBackColor <> &HFF& Then
                For j = 0 To Grd1.Cols - 1
                    Grd1.col = j
                    Grd1.CellBackColor = QBColor(15)
                Next j
            End If
            'Modified by Morgan 2018/5/18
            'GRD1.col = 6
            Grd1.col = 7
            'end 2018/5/18
            Screen.MousePointer = vbHourglass
            frm090218_2.Show
            frm090218_2.Tag = Pub_RplStr(Grd1.Text)
            frm090218_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     Next i
     Me.Enabled = True
Case 1
        frm090218.Show
        Unload Me
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
End Sub

Private Sub Combo1_Click()
StrMenu
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Dim tmpInti As Integer

'2010/10/21 MODIFY BY SONIA 改抓專利處英文顧問
'*****改此部門條件要改四個畫面frm090201_2,frm090218,frm090218_1,frm100101_F
'strSql = "select st01||' ==> '||st02 from staff where st04='1' and st03='F62' and st01<>'99998' order by Decode(ST01,'99998','00000',ST01) "
strSql = "select st01||' ==> '||st02 from staff where st04='1' and st03='P14' and st01<>'99998' order by Decode(ST01,'99998','00000',ST01) "
i = 0
Combo1.Clear
'Combo1.AddItem "", 0
Set ADORECORDSET66 = New ADODB.Recordset
With ADORECORDSET66
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        .MoveFirst
        Do While .EOF = False
            Combo1.AddItem "" & .Fields(0), i
            i = i + 1
            .MoveNext
        Loop
        Combo1.Text = Combo1.List(0)
    Else
    End If
End With
If Trim(frm090218.Combo1.Text) <> "" Then
    For tmpInti = 0 To Combo1.ListCount - 1
        If Trim(Mid(frm090218.Combo1.Text, 1, 6)) = Trim(Mid(Combo1.List(tmpInti), 1, InStr(1, Combo1.List(tmpInti), "=") - IIf(InStr(1, Combo1.List(tmpInti), "=") = 0, 0, 1))) Then
            Combo1.Text = Combo1.List(tmpInti)
        End If
    Next tmpInti
    Combo1.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090218_1 = Nothing
End Sub

Sub StrMenu()
Dim i As Integer, j As Integer, oCount As Integer
Screen.MousePointer = vbHourglass
Grd1.MousePointer = flexArrowHourGlass
DoEvents
Grd1.Clear
Grd1.Rows = 2
SetGrd
strSql = ""
oCount = 0
If Trim(frm090218.Txt1(0)) <> "" Then
    strSql = strSql & " and ep09>=" & ChangeTStringToWString(frm090218.Txt1(0)) & " "
End If
If Trim(frm090218.Txt1(1)) <> "" Then
    strSql = strSql & " and ep09<=" & ChangeTStringToWString(frm090218.Txt1(1)) & " "
End If
If Trim(Combo1.Text) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090218.Label1(25) & frm090218.Combo1.Text 'Add By Sindy 2010/12/20
End If
If Trim(frm090218.Txt1(0)) <> "" Or Trim(frm090218.Txt1(1)) <> "" Then
   pub_QL05 = pub_QL05 & ";" & frm090218.Label2 & frm090218.Txt1(0) & "-" & frm090218.Txt1(1) 'Add By Sindy 2010/12/20
End If
'Modify By Sindy 2013/11/13 +增加電子簽核流程檔,完稿日改顯示送核日
'strSql = "select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),st02,sqldatet(ep09),EP32,ep02"
'Modify By Sindy 2013/12/25 +案件性質
'Modify By Sindy 2015/4/24 +and cp57 is null 已取消收文不需要再列出來
'Modified by Morgan 2018/5/18 改語法調整效能
'strSql = "select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),st02,sqldatet(eep.eep06),EP32,ep02" & _
         " from engineerprogress,caseprogress,patent,staff,casepropertymap" & _
         ",(select eep01,eep02,eep03,eep04,eep05,eep06 from empelectronprocess where eep01||eep02 in(select eep01||max(eep02) from empelectronprocess where eep04='" & EMP_送英核 & "' or instr(eep11,'" & EMP_送英核 & "')>0 group by eep01)) eep" & _
         " where ep03='" & Trim(Mid(Me.Combo1.Text, 1, 6)) & "' and ep02=cp09(+)" & _
         " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
         " and ep05=st01(+) and ep33 is null and ep02=eep.eep01 " & strSql & _
         " and cp01=cpm01(+) and cp10=cpm02(+) and cp57 is null" & _
         " order by ep09,cp01||'-'||cp02||'-'||cp03||'-'||cp04"

strSql = "select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),st02,sqldatet(eep06),EP32,ep02" & _
         " from engineerprogress,caseprogress,patent,staff,casepropertymap,empelectronprocess eep" & _
         " where ep03='" & Trim(Mid(Me.Combo1.Text, 1, 6)) & "' and ep02=cp09(+)" & _
         " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
         " and ep05=st01(+) and ep33 is null and eep01(+)=ep02 and eep02=(select max(b.eep02) from empelectronprocess b where b.eep01=ep02 and (b.eep04='" & EMP_送英核 & "' or instr(b.eep11,'" & "流程狀態:" & EMP_送英核 & "')>0))" & strSql & _
         " and cp01=cpm01(+) and cp10=cpm02(+) and cp57 is null" & _
         " order by ep09,cp01||'-'||cp02||'-'||cp03||'-'||cp04"
'end 2018/5/18
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/20
    Set Grd1.Recordset = adoRecordset
    SetGrd
    For i = 1 To Grd1.Rows - 1
        Grd1.col = 0
        Grd1.row = i
        For j = 0 To Grd1.Cols - 1
            Grd1.col = j
            Grd1.CellBackColor = QBColor(15)
        Next j
        Grd1.col = 5
        If Trim(Grd1.Text) = "Y" Then
            For j = 0 To Grd1.Cols - 1
                Grd1.col = j
                Grd1.CellBackColor = &HFF&
            Next j
        Else
            oCount = oCount + 1
        End If
     Next i
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/20
    ShowNoData
    If Combo1.Enabled = False Then
        cmdOK_Click 1
        Exit Sub
    End If
End If
CheckOC
LBL1.Caption = "共  " & Trim(oCount) & " 件 "
Grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2013/11/13 完稿日改顯示送核日
   'Modify By Sindy 2013/12/25 +案件性質
   arrGridHeadText = Array("V", "本所案號", "案件名稱", "案件性質", "承辦人", "送核日", "暫停核稿", "")
   arrGridHeadWidth = Array(200, 1400, 2000, 1000, 800, 850, 800, 0)
                        
   Grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To Grd1.Cols - 1
      Grd1.row = 0
      Grd1.col = iRow
      Grd1.Text = arrGridHeadText(iRow)
      Grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      Grd1.CellAlignment = flexAlignCenterCenter
   Next
End Sub


Private Sub grd1_SelChange()
Grd1.Visible = False
Grd1.row = Grd1.MouseRow
Grd1.col = 0
If Grd1.row <> 0 Then
    If Grd1.Text = "V" Then
         Grd1.Text = ""
         If Grd1.CellBackColor <> &HFF& Then
             For i = 0 To Grd1.Cols - 1
                  Grd1.col = i
                  Grd1.CellBackColor = QBColor(15)
             Next i
        End If
    Else
         Grd1.Text = "V"
         If Grd1.CellBackColor <> &HFF& Then
            For i = 0 To Grd1.Cols - 1
                Grd1.col = i
                Grd1.CellBackColor = &HFFC0C0
            Next i
        End If
    End If
End If
Grd1.Visible = True
End Sub
