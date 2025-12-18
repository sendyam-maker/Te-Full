VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170304 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月薪資資料查詢"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   1  '開啟
      Index           =   0
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   0
      Top             =   90
      Width           =   765
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   375
      Left            =   6705
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdeXit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   7830
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "9501"
      Top             =   390
      Width           =   585
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2070
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "9512"
      Top             =   390
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4440
      Left            =   135
      TabIndex        =   7
      Top             =   780
      Width           =   8685
      _ExtentX        =   15325
      _ExtentY        =   7832
      _Version        =   393216
      Cols            =   29
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frm170304.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   29
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   7155
      TabIndex        =   12
      Top             =   5325
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5175
      TabIndex        =   11
      Top             =   5325
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "平均基準月薪："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5715
      TabIndex        =   10
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "工作總天數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3915
      TabIndex        =   9
      Top             =   5370
      Width           =   1170
   End
   Begin MSForms.Label LblName 
      Height          =   285
      Left            =   2115
      TabIndex        =   8
      Top             =   120
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   420
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   1770
      X2              =   2190
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "frm170304"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Create by Morgan 2009/2/3
Option Explicit


Private Sub cmdExit_Click()
   Unload Me
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   Dim bCancel As Boolean
   If Text1(0) = "" Then
      MsgBox "員工編號不可空白!!"
      Text1(0).SetFocus
      Exit Function
   End If
   Text1_Validate 0, bCancel
   If bCancel = True Then
      Exit Function
   End If
   FormCheck = True
End Function

Private Sub QueryTable()
   Dim stCon As String
   Dim stVTB As String
   Dim strDays  As String, strSalary As String
   
   If Text1(0) <> "" Then
      stCon = stCon & " and (sm01='" & Text1(0) & "' or sm01='" & Left(Text1(0), 2) & "A" & Mid(Text1(0), 4) & "')"
   End If
   
   If Text1(1) <> "" Then
      stCon = stCon & " and sm02>=" & Val(Text1(1)) + 191100
   End If
   
   If Text1(2) <> "" Then
      stCon = stCon & " and sm02<=" & Val(Text1(2)) + 191100
   End If
   
On Error GoTo flgErr

   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/2/1 +sm43
   'Modified by Sindy 2020/6/23 +sm45
   'Modified by Morgan 2024/1/31 -acc090(沒用)
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stVTB = "select max((trunc(sm02/100)-1911)||'/'||substr(sm02,5)) sm02" & _
      ",to_char(sum(sm04),'9,999,999') sm04,to_char(sum(sm05),'9,999,999') sm05" & _
      ",to_char(sum(sm45),'9,999,999') sm45,to_char(sum(sm06),'9,999,999') sm06" & _
      ",to_char(sum(sm07),'9,999,999') sm07" & _
      ",to_char(sum(sm08),'9,999,999') sm08,to_char(sum(sm09),'9,999,999') sm09" & _
      ",to_char(sum(sm10),'9,999,999') sm10,sum(sm11) sm11" & _
      ",to_char(sum(sm12),'9,999,999') sm12,to_char(sum(sm13),'9,999,999') sm13" & _
      ",to_char(sum(sm19),'9,999,999') sm19,to_char(sum(sm20),'9,999,999') sm20" & _
      ",to_char(sum(sm21),'9,999,999') sm21,to_char(sum(sm22),'9,999,999') sm22" & _
      ",to_char(sum(sm23),'9,999,999') sm23,to_char(sum(sm24),'9,999,999') sm24" & _
      ",to_char(sum(sm14),'9,999,999') sm14,to_char(sum(sm15),'9,999,999') sm15" & _
      ",to_char(sum(sm16),'9,999,999') sm16,to_char(sum(sm17),'9,999,999') sm17" & _
      ",to_char(sum(sm18),'9,999,999') sm18,to_char(sum(sm43),'9,999,999') sm43" & _
      ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm45,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)" & _
      "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)) s1" & _
      ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
      "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)) s2" & _
      ",to_char(sum(sm30),'9,999,999') sm30,to_char(sum(sm25),'9,999,999') sm25" & _
      ",avg(sm27) sm27" & _
      ",st06,sm03,st01,max(st02) st01N" & _
      " from salarymonth,staff" & _
      " where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)" & stCon & _
      " group by st06,sm03,st01,sm02"
      
   'Modify By Sindy 2020/6/23 ,sm06 技術證照津貼 => ,sm45 證照津貼,sm06 技術津貼
   strExc(0) = "select sm02 年月,sm04 基本薪資,sm05 職務津貼" & _
      ",sm45 證照津貼,sm06 技術津貼,sm07 午餐津貼,sm08 差旅津貼,sm09 房租津貼,sm10 特支費" & _
      ",sm11 加班時數,sm12 加班費,sm13 其他所得,sm19 貸款還款,sm20 借支還款" & _
      ",sm21 缺勤扣款,sm22 未打卡,sm23 其他扣款,sm24 所得稅,sm14 勞保費" & _
      ",sm15 健保費,sm43 補充保費,sm16 勞退自提,sm17 婚喪戶助,sm18 互助會" & _
      ",to_char(nvl(s1,0),'9,999,999') 應發金額" & _
      ",to_char(nvl(s2,0),'9,999,999') 應扣金額" & _
      ",to_char(nvl(s1,0)-nvl(s2,0),'9,999,999') 實發金額" & _
      ",sm30 勞退公司提撥,sm25 年終獎金基準月薪" & _
      ",sm27 工作天數 from (" & stVTB & ") X order by 1"

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList.Recordset = RsTemp.Clone
   SetGrid
   If RsTemp.RecordCount > 0 Then
      strDays = 0
      strSalary = 0
      Do While Not RsTemp.EOF
         strDays = Val(strDays) + Val(RsTemp("工作天數"))
         strSalary = Val(strSalary) + Val(Format(RsTemp("年終獎金基準月薪")))
         RsTemp.MoveNext
      Loop
      Label1(1) = strDays
      Label1(2) = Format(strSalary / RsTemp.RecordCount, "#,###")
   Else
      Label1(1) = ""
      Label1(2) = ""
      MsgBox "查無資料！", vbInformation
   End If
   
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdok_Click()
   Screen.MousePointer = vbHourglass
   If FormCheck Then
      Me.Enabled = False
      QueryTable
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub SetGrid()
   Dim iCol As Integer
   With grdDataList
      .Visible = False
      .ColAlignment(0) = flexAlignCenterCenter
      For iCol = 1 To .Cols - 1
         .ColAlignment(iCol) = flexAlignRightCenter
      Next
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1(0) = ""
   Text1(1) = ""
   Text1(2) = ""
   LblName = ""
   Label1(1) = ""
   Label1(2) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170304 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 2 Then
      If Text1(1) <> "" Then Text1(2) = Text1(1)
   End If
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1, 2
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         cmdok_Click
      Case vbKeyEscape
         cmdExit_Click
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If Text1(Index) <> "" Then
         If ClsPDGetOtherIncomer(Text1(Index), strExc(1)) = True Then
            LblName.Caption = strExc(1)
         Else
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'If ChkStaffID(Replace(Text1(Index), "A", "0")) = True Then
            If ChkStaffID(Left(Text1(Index), 1) & Replace(Mid(Text1(Index), 2), "A", "0")) = True Then
               Cancel = True
            End If
            If Cancel = False Then
               If ClsPDGetStaffN(Text1(Index), strExc(1), , True) = False Then
                  Cancel = True
                  LblName.Caption = ""
               Else
                  LblName.Caption = strExc(1)
               End If
            End If
         End If
      Else
         LblName.Caption = ""
      End If
   End If
End Sub
