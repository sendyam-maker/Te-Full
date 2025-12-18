VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090632 
   BorderStyle     =   1  '單線固定
   Caption         =   "預定會稿日異常案件查詢"
   ClientHeight    =   5715
   ClientLeft      =   3180
   ClientTop       =   2205
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   4230
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   3105
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1530
      MaxLength       =   2
      TabIndex        =   1
      Top             =   660
      Width           =   420
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1260
      Width           =   330
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1035
      MaxLength       =   2
      TabIndex        =   2
      Top             =   960
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   0
      Top             =   330
      Width           =   420
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4125
      Left            =   0
      TabIndex        =   9
      Top             =   1590
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   7276
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
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
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7260
      TabIndex        =   6
      Top             =   30
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8220
      TabIndex        =   7
      Top             =   30
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "發文日：                     -"
      Height          =   180
      Left            =   2385
      TabIndex        =   13
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否含已發文：        (Y:是)"
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   1320
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CFP案超過天數：            個工作天"
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   720
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "修改次數：            次以上"
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   1020
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P案超過天數：            個工作天"
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   2430
   End
End
Attribute VB_Name = "frm090632"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Create by Morgan 2010/10/11
Option Explicit

Private Sub SetDataListWidth()
   With grdDataList
   .Visible = False
   .Cols = 11
   .row = 0
   .col = 0: .Text = "承辦人"
   .ColWidth(0) = 650
   .CellAlignment = flexAlignCenterCenter
   .col = 1: .Text = "收文日"
   .ColWidth(1) = 800
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignCenterCenter
   .col = 2: .Text = "齊備日"
   .ColWidth(2) = 800
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignCenterCenter
   .col = 3: .Text = "發文"
   .ColWidth(3) = 300
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(3) = flexAlignCenterCenter
   .col = 4: .Text = "本所案號"
   .ColWidth(4) = 1400
   .CellAlignment = flexAlignCenterCenter
   .col = 5: .Text = "案件名稱"
   .ColWidth(5) = 1300
   .CellAlignment = flexAlignCenterCenter
   .col = 6: .Text = "案件性質"
   .ColWidth(6) = 800
   .CellAlignment = flexAlignCenterCenter
   .col = 7: .Text = "承辦期限"
   .ColWidth(7) = 800
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(7) = flexAlignCenterCenter
   .col = 8: .Text = "預會日"
   .ColWidth(8) = 800
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(8) = flexAlignCenterCenter
   .col = 9: .Text = "工作天"
   .ColWidth(9) = 600
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(9) = flexAlignCenterCenter
   .col = 10: .Text = "修改次數"
   .ColWidth(10) = 750
   .CellAlignment = flexAlignCenterCenter
   .ColAlignment(10) = flexAlignCenterCenter
   .Visible = True
   End With
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim CheckOk As Boolean
   Select Case Index
   Case 0
      If Val(Text1) + Val(Text2) + Val(Text3) = 0 Then
         MsgBox "條件不可全為空白!!"
         Text1.SetFocus
         Exit Sub
      ElseIf Text4 = "Y" And (txtDate(0) = "" Or txtDate(1) = "") Then
         MsgBox "含已發文時必須輸入發文日區間!!"
         If txtDate(0) = "" Then
            txtDate(0).SetFocus
         ElseIf txtDate(1) = "" Then
            txtDate(1).SetFocus
         End If
         Exit Sub
      End If
      
      Screen.MousePointer = vbHourglass
      Me.grdDataList.MousePointer = flexHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
      StrMenu
      Me.grdDataList.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
   Case 1
      Unload Me
   Case Else
   End Select
End Sub

Private Sub StrMenu()
   Dim stVTB As String
   Dim stConCP27 As String, stConEP30 As String
   
   stConCP27 = ""
   If Text4 = "Y" Then
      stConCP27 = " ( cp27>=" & DBDATE(txtDate(0)) & _
         " and cp27<=" & DBDATE(txtDate(1)) & ")"
      pub_QL05 = pub_QL05 & ";" & Left(Label5, 4) & txtDate(0) & "-" & txtDate(1) 'Add By Sindy 2010/12/20
      pub_QL05 = pub_QL05 & ";" & Left(Label4, 7) & Text4 'Add By Sindy 2010/12/20
   End If
   
   stConEP30 = ""
   If Val(Text3) > 0 Then
      stConEP30 = stConEP30 & " and ep30>=" & Val(Text3)
      pub_QL05 = pub_QL05 & ";" & Left(Label2, 5) & Text3 & "次以上" 'Add By Sindy 2010/12/20
   End If
   
   '為避免太耗時只抓預定會稿日1年內資料
   If Val(Text1) > 0 Then
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & " select ep02 X1,count(*) X2" & _
         " From engineerprogress, caseprogress, workday" & _
         " where ep28>to_char(sysdate-365,'yyyymmdd')" & stConEP30 & _
         " and cp09(+)=ep02 and cp01='P' and cp27 is null" & _
         " and wd01>=cp48 and wd01<=ep28" & _
         " group by ep02 having count(*)>" & Val(Text1)
         
      If stConCP27 <> "" Then
         stVTB = stVTB & " union select ep02 X1,count(*) X2" & _
            " From caseprogress,engineerprogress,workday" & _
            " where cp01='P' and " & stConCP27 & _
            " and ep02(+)=cp09 and wd01>=cp48 and wd01<=ep28" & stConEP30 & _
            " group by ep02 having count(*)>" & Val(Text1)
      End If
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 7) & Text1 & "個工作天" 'Add By Sindy 2010/12/20
   End If

   If Val(Text2) > 0 Then
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & " select ep02 X1,count(*) X2" & _
         " From engineerprogress, caseprogress, workday" & _
         " where ep28>to_char(sysdate-365,'yyyymmdd')" & stConEP30 & _
         " and cp09(+)=ep02 and cp01='CFP' and cp27 is null" & _
         " and wd01>=cp48 and wd01<=ep28" & _
         " group by ep02 having count(*)>" & Val(Text2)
         
      If stConCP27 <> "" Then
         stVTB = stVTB & " union select ep02 X1,count(*) X2" & _
            " From caseprogress,engineerprogress,workday" & _
            " where cp01='CFP' and " & stConCP27 & _
            " and ep02(+)=cp09 and wd01>=cp48 and wd01<=ep28" & stConEP30 & _
            " group by ep02 having count(*)>" & Val(Text2)
      End If
      pub_QL05 = pub_QL05 & ";" & Left(Label3, 9) & Text2 & "個工作天" 'Add By Sindy 2010/12/20
   End If

   If stConEP30 <> "" And stVTB = "" Then
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & " select ep02 X1,count(*) X2" & _
         " From engineerprogress, caseprogress, workday" & _
         " where ep28>to_char(sysdate-365,'yyyymmdd') and " & _
         " and cp09(+)=ep02 and cp01 in ('P','CFP')" & stConEP30 & _
         " and wd01>=cp48 and wd01<=ep28" & _
         " group by ep02"
         
      If stConCP27 <> "" Then
         stVTB = stVTB & " union select ep02 X1,count(*) X2" & _
            " From caseprogress,engineerprogress,workday" & _
            " where cp01 in ('P','CFP') and " & stConCP27 & _
            " and ep02(+)=cp09 and wd01>=cp48 and wd01<=ep28" & stConEP30 & _
            " group by ep02"
      End If
   End If


strSql = "select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
   ",nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),sqldatet(cp48),sqldatet(ep28),X2,ep30" & _
   " from ( " & stVTB & ") ,caseprogress,engineerprogress,patent,casepropertymap,staff" & _
   " where cp09(+)=X1 and ep02(+)=cp09" & _
   " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
   " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14" & _
   " order by cp14,cp09"

intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
Set grdDataList.Recordset = RsTemp.Clone
SetDataListWidth
If intI = 0 Then
   InsertQueryLog (0) 'Add By Sindy 2010/12/20
   MsgBox "無符合條件之資料 !", vbInformation
Else
   InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/20
End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   grdDataList.Clear
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090632 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_Change()
   If Text4 = "Y" Then
      txtDate(0).Enabled = True
      txtDate(1).Enabled = True
   Else
      txtDate(0).Text = ""
      txtDate(0).Enabled = False
      txtDate(1).Text = ""
      txtDate(1).Enabled = False
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If PUB_CheckKeyInDate(txtDate(Index)) = -1 Then
      txtDate(Index).SetFocus
      txtDate_GotFocus Index
      Cancel = True
      Exit Sub
   End If
   If Index = 1 Then
       If RunNick(txtDate(Index - 1), txtDate(Index)) Then
           txtDate(Index - 1).SetFocus
           txtDate_GotFocus (Index - 1)
           Cancel = True
           Exit Sub
       End If
   End If
End Sub
