VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090616_1 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "るσ"
   ClientHeight    =   5625
   ClientLeft      =   2100
   ClientTop       =   2775
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9240
   Begin VB.CommandButton cmdok 
      Caption         =   "挡(&X)"
      Height          =   345
      Index           =   1
      Left            =   8175
      TabIndex        =   1
      Top             =   45
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "玡礶(&U)"
      Height          =   345
      Index           =   0
      Left            =   6810
      TabIndex        =   0
      Top             =   45
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5070
      Left            =   135
      TabIndex        =   2
      Top             =   465
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   8943
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "穝灿砰-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm090616_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 эΘForm2.0 ; grd1э=穝灿砰-ExtB
'Memo By Morgan 2012/12/10 醇舦逆э
'2010/12/1 memo by sonia 絪腹逆э
'Memo by Morgan2010/8/17 ら戳逆э
Option Explicit

Dim SWPColor As String, SWPColor2 As String, SWPRow As String, SWPRow2 As String
Dim m_blnColOrderAsc As Boolean '逆戈パ逼
Dim PLeft(0 To 15) As Integer, iPrint As Integer, Page As Integer
Dim m_IsRun As Boolean
Dim m_ProState As String 'Add By Sindy 2017/8/10 癘魁ヘ玡舦


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
         frm090616_0.Show
         Unload Me
Case 1
         Unload frm090616_0
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 穝砞﹚舦
If m_IsRun = False Then
   m_IsRun = True
      If frm090616_0.txt1(3) = "2" Then
         Me.Hide
      End If
'      If ProState <> "2" Then
'      cmdok(0).Visible = False
'      End If
      Me.Hide
      Screen.MousePointer = vbHourglass
      DoEvents
      If StrMenu = False Then
         Screen.MousePointer = vbDefault
         cmdOK_Click 0
         Exit Sub
      End If
      Screen.MousePointer = vbDefault
      Me.Show
End If
End Sub

Private Sub Form_Load()
m_IsRun = False
MoveFormToCenter Me
m_ProState = ProState 'Add By Sindy 2017/8/10 癘魁ヘ玡舦
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090616_1 = Nothing
End Sub

Private Sub SetGrd1()
Dim j As Integer
With grd1
    .Visible = False
    If ProSysState = "1" Then
         .Cols = 13
         .row = 0
         .col = 0:   .Text = "┯快"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "ヘ夹"
         .ColWidth(1) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "祇ゅ"
         .ColWidth(2) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "祇ゅ"
         .ColWidth(3) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "祇ゅ"
         .ColWidth(4) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "祇ゅ"
         .ColWidth(5) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "祇ゅ"
         .ColWidth(6) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "┯快"
         .ColWidth(7) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "┯快"
         .ColWidth(8) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "┯快"
         .ColWidth(9) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "硉σ"
         .ColWidth(10) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 11:  .Text = "σ"
         .ColWidth(11) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = ""
         .ColWidth(12) = 0
         .CellAlignment = flexAlignCenterCenter
         .row = 1
         .col = 0:   .Text = "┯快"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "膀计"
         .ColWidth(1) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "膀计"
         .ColWidth(2) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "膀计笷Θ瞯%"
         .ColWidth(3) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "膀计眔だ"
         .ColWidth(4) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "翴计笷Θ瞯%"
         .ColWidth(5) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "翴计眔だ"
         .ColWidth(6) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "膀计"
         .ColWidth(7) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "膀计笷Θ瞯%"
         .ColWidth(8) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "膀计眔だ"
         .ColWidth(9) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "眔だ"
         .ColWidth(10) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 11:  .Text = "眔だ"
         .ColWidth(11) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = ""
         .ColWidth(12) = 0
         .CellAlignment = flexAlignCenterCenter
   Else
         .Cols = 17
         .row = 0
         .col = 0:   .Text = "酶瓜"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "ヘ夹"
         .ColWidth(1) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "祇ゅ秖"
         .ColWidth(2) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "祇ゅ秖"
         .ColWidth(3) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "祇ゅ秖"
         .ColWidth(4) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "祇ゅ眎计"
         .ColWidth(5) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "祇ゅ眎计"
         .ColWidth(6) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "祇ゅ翴计"
         .ColWidth(7) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "祇ゅ翴计"
         .ColWidth(8) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "┯快"
         .ColWidth(9) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "┯快"
         .ColWidth(10) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "┯快"
         .ColWidth(11) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 12:   .Text = "┯快"
         .ColWidth(12) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 13:   .Text = "┯快"
         .ColWidth(13) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 14:   .Text = "硉σ"
         .ColWidth(14) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 15:  .Text = "σ"
         .ColWidth(15) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 16:  .Text = ""
         .ColWidth(16) = 0
         .CellAlignment = flexAlignCenterCenter
         .row = 1
         .col = 0:   .Text = "酶瓜"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "膀计"
         .ColWidth(1) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "膀计"
         .ColWidth(2) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "笷Θ瞯%"
         .ColWidth(3) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "眔だ"
         .ColWidth(4) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "笷Θ瞯%"
         .ColWidth(5) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "眔だ"
         .ColWidth(6) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "笷Θ瞯%"
         .ColWidth(7) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "眔だ"
         .ColWidth(8) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "膀计"
         .ColWidth(9) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "膀计笷Θ瞯%"
         .ColWidth(10) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "膀计眔だ"
         .ColWidth(11) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 12:   .Text = "眎计笷Θ瞯%"
         .ColWidth(12) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 13:   .Text = "眎计眔だ"
         .ColWidth(13) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 14:   .Text = "眔だ"
         .ColWidth(14) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 15:  .Text = "眔だ"
         .ColWidth(15) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 16:  .Text = ""
         .ColWidth(16) = 0
         .CellAlignment = flexAlignCenterCenter
   End If
   .MergeCells = flexMergeRestrictRows
   .MergeRow(0) = True
   .MergeCol(0) = True

   .MergeCol(1) = True
    .Visible = True
End With
   With Me.grd1
      .row = 2
         For j = 1 To .Cols - 1
             .col = j
             .CellBackColor = &HFFC0C0
         Next j
      SWPColor2 = SWPColor
      SWPRow2 = .row
   End With

End Sub

Private Sub GRD1_DblClick()
Me.Enabled = False
Screen.MousePointer = vbHourglass
    If Me.grd1.MouseRow > 1 Then
        If Me.grd1.Rows > 2 Then
            SWPRow = str(grd1.MouseRow)
        End If
    End If
Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Strindex As Integer
Dim j As Integer
Dim oMouseCol As Integer
If Me.grd1.MouseRow <= 0 Then Exit Sub
If Button = 1 Then
    Screen.MousePointer = vbHourglass
    SWPRow = str(grd1.MouseRow)
    oMouseCol = grd1.MouseCol
    If Val(SWPRow) < 2 Then
        Select Case oMouseCol
        Case 0
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 5 '狜经
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 6 '经
                m_blnColOrderAsc = True
            End If
        Case Else
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 3 '狜经
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 4 '经
                m_blnColOrderAsc = True
            End If
        End Select
    End If
    Strindex = SWPRow
    With grd1
        DoEvents
        .Visible = False
         If Val(SWPRow) = 0 Or Val(SWPRow) = 1 Then
            For j = 2 To .Rows - 1
               .row = j
               .col = 1
               If .CellBackColor = &HFFC0C0 Then
                  SWPRow2 = j
                  .Visible = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            Next j
         End If
        If SWPRow2 <> "" Then
           .row = SWPRow2
           For j = 1 To .Cols - 1
               .col = j
               .CellBackColor = QBColor(15)
           Next j
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Or .row = 1 Then
            .row = 2
        End If
         For j = 1 To .Cols - 1
             .col = j
             .CellBackColor = &HFFC0C0
         Next j
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
End If
End Sub

Function StrMenu() As Boolean
StrMenu = True
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim CalMonth As Integer
Dim j As Integer
strSql = ""
strSQL1 = ""
strSQL2 = ""
CalMonth = 0
CalMonth = DateDiff("m", ChangeWStringToWDateString(Val(frm090616_0.txt1(0) & "01") + 19110000), ChangeWStringToWDateString(Val(frm090616_0.txt1(1) & "01") + 19110000)) + 1
If Len(Trim(frm090616_0.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " and ma01='" & frm090616_0.txt1(2) & "' "
   strSQL2 = strSQL2 & " and pe01='" & frm090616_0.txt1(2) & "' "
End If
strSQL1 = strSQL1 & " and ma03='" & ProSysState & "' "
'MODIFY BY SONIA 2014/4/11  pe02 in ('P','CFP') 縋ゅΤTヘ夹
'Modified by Morgan 2018/5/18 O12 琘ㄇΤ笲衡逆篈穦琌double(5)赣篈穦旧璓ず甧礚猭タ盽陪ボ(穦琌 "~00000001")э 0 笲衡タ
If ProSysState = "1" Then '┯快
      strSql = " select  A1+0 as A1,ma37+0 as ma37,decode(A1,0,0,round(ma37/A1 * 100,2))+0 as C1,0 as C2,decode(A2,0,0,round(ma40/A2 * 100,2))+0 as C3,0 as C4,ma43+0 as ma43,decode(A1,0,0,round(ma43/A1 * 100,2))+0 as C5,0,round(ma35/" & CalMonth & ",2)+0 as C6,0 as C7,st02  from (select pe01,sum(nvl(decode(pe02,'CFP',pe05*2,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe02 in ('P','CFP') And pe03>=" & Val(frm090616_0.txt1(0)) + 191100 & " and pe03<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL2 & " group by pe01) APE ,("
      'Added by Morgan 2019/3/19 108σ(筄戳ン计э–ンΙ0.5だぃ埃讽る笷Θ瞯)
      If Val(frm090616_0.txt1(0)) + 191100 >= Val(Left(PUB_108RuleDate, 6)) Then
         strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - decode(ma44,0,0,0.5*ma51),0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51 from monthassess,staff where ma01=st01(+) and ma02>=" & Val(frm090616_0.txt1(0)) + 191100 & " and ma02<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL1
      Else
      'end 2019/3/19
         strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - decode(ma44,0,0,((0.5/(ma44))*ma51 )),0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51 from monthassess,staff where ma01=st01(+) and ma02>=" & Val(frm090616_0.txt1(0)) + 191100 & " and ma02<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL1
      End If 'Added by Morgan 2019/3/19
      strSql = strSql & " group by st01,st02,ma03) AAA where AAA.st01=APE.pe01(+) order by st01"
Else
      strSql = " select  A3+0 as A3,ma37+0 as ma37,decode(A3,0,0,round(ma37/A3 * 100,2))+0 as C1,0 as C2,decode(A4,0,0,round(ma47/A4 * 100,2))+0 as C3,0 as C4,decode(A5,0,0,round(ma40/A5 * 100,2))+0 as C5,0 as C6,ma43+0 as ma43,decode(A3,0,0,round(ma43/A3 * 100,2))+0 as C7,0 as C8,decode(A4,0,0,round(ma52/A4 * 100,2))+0 as C9,0 as C10,round(ma35/2/" & CalMonth & ",2)+0 as C11,0 as C12,st02 from (select pe01,sum(nvl(decode(pe02,'CFP',pe05*2,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe02 in ('P','CFP') And pe03>=" & Val(frm090616_0.txt1(0)) + 191100 & " and pe03<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL2 & " group by pe01) APE ,("
      'Added by Morgan 2019/3/19 108σ(筄戳ン计э–ンΙ0.5だぃ埃讽る笷Θ瞯)
      If Val(frm090616_0.txt1(0)) + 191100 >= Val(Left(PUB_108RuleDate, 6)) Then
         strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - decode(ma44,0,0,0.5*ma51) ,0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51,sum(nvl(ma52,0)) as ma52 from monthassess,staff where ma01=st01(+) and ma02>=" & Val(frm090616_0.txt1(0)) + 191100 & " and ma02<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL1
      Else
         strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - decode(ma44,0,0,((0.5/(ma44))*ma51 )),0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51,sum(nvl(ma52,0)) as ma52 from monthassess,staff where ma01=st01(+) and ma02>=" & Val(frm090616_0.txt1(0)) + 191100 & " and ma02<=" & Val(frm090616_0.txt1(1)) + 191100 & " " & strSQL1
      End If 'Added by Morgan 2019/3/19
      strSql = strSql & " group by st01,st02,ma03) AAA where AAA.st01=APE.pe01(+) order by st01"
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
      Set grd1.Recordset = adoRecordset
      
      '衡眔だ
      strSql = "select * from assessrate where ar01 in (select max(ar01) from assessrate where ar01<=" & DBDATE(Trim(frm090616_0.txt1(0)) & "01") & ") "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
            With grd1
                  For j = 2 To grd1.Rows - 1
                     If ProSysState = "1" Then '┯快
                        .TextMatrix(j, 0) = .TextMatrix(j, 12)
                        '祇ゅ膀计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 3)) < 100 Then
                        If Val(.TextMatrix(j, 3)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 4) = Format((Val(.TextMatrix(j, 3)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar09").Value), "####0.00")
                        Else
                           .TextMatrix(j, 4) = Format((Val(.TextMatrix(j, 3)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar09").Value), "####0.00")
                        End If
                        '祇ゅ翴计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 5)) < 100 Then
                        If Val(.TextMatrix(j, 5)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar10").Value), "####0.00")
                        Else
                           .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar10").Value), "####0.00")
                           '翴计Τ
                           If Val(.TextMatrix(j, 6)) > ((AdoRecordSet3.Fields("ar10").Value) * 1.5) Then
                              .TextMatrix(j, 6) = Format((AdoRecordSet3.Fields("ar10").Value) * 1.5, "#####0.00")
                           End If
                        End If
                        '┯快眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 8)) < 100 Then
                        If Val(.TextMatrix(j, 8)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 9) = Format((Val(.TextMatrix(j, 8)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar11").Value), "####0.00")
                        Else
                           .TextMatrix(j, 9) = Format((Val(.TextMatrix(j, 8)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar11").Value), "####0.00")
                        End If
                        '硉σ程琌 0 だ
                        If Val(.TextMatrix(j, 10)) < 0 Then
                              .TextMatrix(j, 10) = "0.00"
                        End If
                        'σ眔だ
                        .TextMatrix(j, 11) = Format(Val(.TextMatrix(j, 4)) + Val(.TextMatrix(j, 6)) + Val(.TextMatrix(j, 9)) + Val(.TextMatrix(j, 10)), "#####0.00")
                     Else
                        .TextMatrix(j, 0) = .TextMatrix(j, 16)
                        '祇ゅ膀计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 3)) < 100 Then
                        If Val(.TextMatrix(j, 3)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 4) = Format((Val(.TextMatrix(j, 3)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar20").Value), "####0.00")
                        Else
                           .TextMatrix(j, 4) = Format((Val(.TextMatrix(j, 3)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar20").Value), "####0.00")
                        End If
                        '祇ゅ眎计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 5)) < 100 Then
                        If Val(.TextMatrix(j, 5)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar21").Value), "####0.00")
                        Else
                           .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar21").Value), "####0.00")
                        End If
                        '祇ゅ翴计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 7)) < 100 Then
                        If Val(.TextMatrix(j, 7)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 8) = Format((Val(.TextMatrix(j, 7)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar22").Value), "####0.00")
                        Else
                           .TextMatrix(j, 8) = Format((Val(.TextMatrix(j, 7)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar22").Value), "####0.00")
                           '翴计Τ
                           If Val(.TextMatrix(j, 8)) > ((AdoRecordSet3.Fields("ar22").Value) * 1.5) Then
                              .TextMatrix(j, 8) = Format((AdoRecordSet3.Fields("ar22").Value) * 1.5, "#####0.00")
                           End If
                        End If
                        '┯快膀计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 10)) < 100 Then
                        If Val(.TextMatrix(j, 10)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 11) = Format((Val(.TextMatrix(j, 10)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar23").Value), "####0.00")
                        Else
                           .TextMatrix(j, 11) = Format((Val(.TextMatrix(j, 10)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar23").Value), "####0.00")
                        End If
                        '┯快眎计眔だ
                        'Modified by Morgan 2019/3/19 108σ(眔だ(笷Θ瞯)^2璸衡よΑ)
                        'If Val(.TextMatrix(j, 12)) < 100 Then
                        If Val(.TextMatrix(j, 12)) < 100 And Val(frm090616_0.txt1(0)) + 191100 < Val(Left(PUB_108RuleDate, 6)) Then
                        'end 2019/3/19
                           .TextMatrix(j, 13) = Format((Val(.TextMatrix(j, 12)) / 100) ^ 2 * 0.8 * (AdoRecordSet3.Fields("ar24").Value), "####0.00")
                        Else
                           .TextMatrix(j, 13) = Format((Val(.TextMatrix(j, 12)) / 100) * 0.8 * (AdoRecordSet3.Fields("ar24").Value), "####0.00")
                        End If
                        '硉σ程だ琌 0 だ
                        If Val(.TextMatrix(j, 14)) < 0 Then
                              .TextMatrix(j, 14) = "0.00"
                        End If
                        'σ眔だ
                        .TextMatrix(j, 15) = Format(Val(.TextMatrix(j, 4)) + Val(.TextMatrix(j, 6)) + Val(.TextMatrix(j, 8)) + Val(.TextMatrix(j, 11)) + Val(.TextMatrix(j, 13)) + Val(.TextMatrix(j, 14)), "#####0.00")
                     End If
                  Next j
            End With
       End If
      grd1.col = grd1.Cols - 2
      grd1.Sort = 4
      SetGrd1
      If frm090616_0.txt1(3).Text = "2" Then '
         PrintData
         StrMenu = False
      End If
Else
   ShowNoData
   StrMenu = False
End If
End Function

Sub PrintData()
Dim iCol As Integer
Dim iRow As Integer
iPrint = 0
Page = 1
GetPleft
PrintTitle
With grd1
   For iRow = 2 To .Rows - 1
      .row = iRow
      For iCol = 0 To .Cols - 2
         If iCol = 0 Then
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print .TextMatrix(iRow, iCol)
         Else
            Printer.CurrentX = PLeft(iCol) + 600 - Printer.TextWidth(Format(Val(.TextMatrix(iRow, iCol)), "##0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val(.TextMatrix(iRow, iCol)), "##0.00")
         End If
      Next iCol
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
      End If
   Next iRow
End With
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
Erase PLeft
'﹚皚
If ProSysState = "1" Then '┯快
      PLeft(0) = 500
      PLeft(1) = 1500
      PLeft(2) = 2500
      PLeft(3) = 3500
      PLeft(4) = 5000
      PLeft(5) = 6250
      PLeft(6) = 7500
      PLeft(7) = 8750
      PLeft(8) = 10000
      PLeft(9) = 11250
      PLeft(10) = 12500
      PLeft(11) = 13750
Else
      PLeft(0) = 500
      PLeft(1) = 1500
      PLeft(2) = 2500
      PLeft(3) = 3500
      PLeft(4) = 4500
      PLeft(5) = 5500
      PLeft(6) = 6500
      PLeft(7) = 7500
      PLeft(8) = 8500
      PLeft(9) = 9500
      PLeft(10) = 10500
      PLeft(11) = 11500
      PLeft(12) = 12500
      PLeft(13) = 13500
      PLeft(14) = 14500
      PLeft(15) = 15500
End If
End Sub

Sub PrintTitle() '╋繷
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "灿砰"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print IIf(ProSysState = "1", "┯快", "酶瓜") & "るσ"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
Printer.Print "る" & Format(Format(str(Val(frm090616_0.txt1(0)) + 191100) & "01", "####/##/##"), "ee/MM") & "-" & Format(Format(str(Val(frm090616_0.txt1(1)) + 191100) & "01", "####/##/##"), "ee/MM")
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "" & strUserName
If ProSysState = "1" Then
   Printer.CurrentX = 13000
Else
   Printer.CurrentX = 14300
End If
Printer.CurrentY = iPrint
Printer.Print "ら戳" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print IIf(ProSysState = "1", "┯快", "酶瓜") & "" & IIf(Trim(frm090616_0.lbl1.Caption) = "", "┮Τ", frm090616_0.lbl1.Caption)
If ProSysState = "1" Then
   Printer.CurrentX = 13000
Else
   Printer.CurrentX = 14300
End If
Printer.CurrentY = iPrint
Printer.Print "    Ω" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
If ProSysState = "1" Then
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "┯快"
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "ヘ夹"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ"
      Printer.CurrentX = PLeft(3) + ((PLeft(4) - PLeft(3)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ膀计"
      Printer.Line (PLeft(3), iPrint + 290)-(PLeft(5) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(5) + ((PLeft(6) - PLeft(5)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ翴计"
      Printer.Line (PLeft(5), iPrint + 290)-(PLeft(7) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "┯快秖"
      Printer.Line (PLeft(7), iPrint + 290)-(PLeft(10) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "硉σ"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "σ"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print ""
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      ShowLine
   If iPrint >= 9000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       Exit Sub
   End If
Else
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "酶瓜"
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "ヘ夹"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ膀计"
      Printer.Line (PLeft(2), iPrint + 290)-(PLeft(5) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(5) + ((PLeft(6) - PLeft(5)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ眎计"
      Printer.Line (PLeft(5), iPrint + 290)-(PLeft(7) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(7) + ((PLeft(8) - PLeft(7)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "祇ゅ翴计"
      Printer.Line (PLeft(7), iPrint + 290)-(PLeft(9) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "┯快秖"
      Printer.Line (PLeft(9), iPrint + 290)-(PLeft(12) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(12) + ((PLeft(14) - PLeft(13)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "┯快眎计"
      Printer.Line (PLeft(12), iPrint + 290)-(PLeft(14) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "硉"
      Printer.CurrentX = PLeft(15)
      Printer.CurrentY = iPrint
      Printer.Print "σ"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print ""
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "膀计"
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "笷Θ瞯%"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      Printer.CurrentX = PLeft(15)
      Printer.CurrentY = iPrint
      Printer.Print "眔だ"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      ShowLine
   If iPrint >= 9000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       Exit Sub
   End If
End If
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If ProSysState = "1" Then
   Printer.Line (500, iPrint + 150)-(15000, iPrint + 150)
Else
   Printer.Line (500, iPrint + 150)-(16500, iPrint + 150)
End If
iPrint = iPrint + 300
End Sub
