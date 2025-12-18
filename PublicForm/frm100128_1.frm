VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100128_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC客戶直接來所申請比率統計"
   ClientHeight    =   2670
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5500
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1770
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1170
      TabIndex        =   4
      Text            =   "ALL"
      Top             =   1440
      Width           =   3030
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   420
      Width           =   1320
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1110
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2835
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1110
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   1
      Top             =   780
      Width           =   492
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3285
      Left            =   60
      TabIndex        =   8
      Top             =   3300
      Width           =   8955
      _ExtentX        =   15804
      _ExtentY        =   5786
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   4485
      TabIndex        =   7
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   3585
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "（1. 國家   2. 洲）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1665
      TabIndex        =   18
      Top             =   1815
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "統計別："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   1830
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   " (ALL：全部)"
      Height          =   180
      Left            =   4230
      TabIndex        =   16
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "註：統計的案件性質專利為 101 ~ 103，商標為 101。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   2310
      Width           =   5130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "統計部門："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   1155
      Width           =   540
   End
   Begin VB.Line Line5 
      X1              =   2595
      X2              =   2715
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "輸入民國年"
      Height          =   180
      Left            =   4230
      TabIndex        =   11
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   825
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "（1. 收文   2. 發文）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1665
      TabIndex        =   9
      Top             =   825
      Width           =   1575
   End
End
Attribute VB_Name = "frm100128_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/ Form2.0不用改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Create by Morgan 2010/8/25
Option Explicit

Public cmdState As Integer

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetCombo1
   If Left(Pub_StrUserSt03, 2) = "F2" Then
      Combo1.ListIndex = 0
      Combo1.Enabled = False
   ElseIf Left(Pub_StrUserSt03, 2) = "F1" Then
      Combo1.ListIndex = 1
      Combo1.Enabled = False
   Else
      Combo1.ListIndex = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100128_1 = Nothing
End Sub

Private Sub SetCombo1()
   Combo1.Clear
   Combo1.AddItem "FCP", 0
   Combo1.AddItem "FCT", 1
End Sub

Sub PubShowNextData()
   Select Case cmdState
      Case 0
         cmdState = -1
         Screen.MousePointer = vbHourglass
         DoEvents
         If ConstrainCheck = True Then
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
            doQuery
         End If
         Screen.MousePointer = vbDefault
      Case 1
         fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Function ConstrainCheck() As Boolean
   If Combo1.ListIndex < 0 Then
      MsgBox "請選擇部門!!!"
      Combo1.SetFocus
      Exit Function
   End If
   If txt1(0) = "" Then
      MsgBox "查詢別不可空白!!!"
      txt1(0).SetFocus
      Exit Function
   End If
   If txt1(1) = "" Then
      MsgBox "日期(起)不可空白!!!"
      txt1(1).SetFocus
      Exit Function
   ElseIf Not ChkDate(txt1(1)) Then
      txt1(1).SetFocus
      Exit Function
   End If
   If txt1(2) = "" Then
      MsgBox "日期(迄)不可空白!!!"
      txt1(2).SetFocus
      Exit Function
   ElseIf Not ChkDate(txt1(2)) Then
      txt1(2).SetFocus
      Exit Function
   End If
   If txt1(3) = "" Then
      MsgBox "系統類別不可空白!!!"
      txt1(3).SetFocus
      Exit Function
   End If
   If txt1(4) = "" Then
      MsgBox "統計別不可空白!!!"
      txt1(4).SetFocus
      Exit Function
   End If
   ConstrainCheck = True
End Function

Private Sub doQuery()
   Dim stCon As String, stDate1 As String, stDate2 As String, ii As Integer
   Dim stCP10 As String, stTable As String, strSystemKind As String
   Dim arr1
   
   stDate1 = DBDATE(txt1(1))
   stDate2 = DBDATE(txt1(2))
   
   If txt1(3) <> "ALL" Then
      arr1 = Split(txt1(3), ",")
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      stCon = stCon & " AND CP01 IN ( " & strSystemKind & " ) "
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(3) 'Add By Sindy 2010/11/16
   End If
   
   If stDate1 <> "" Then
      If txt1(0) = "1" Then
         stCon = stCon & " and cp05>=" & stDate1
      Else
         stCon = stCon & " and cp27>=" & stDate1
      End If
   End If
   If stDate2 <> "" Then
      If txt1(0) = "1" Then
         stCon = stCon & " and cp05<=" & stDate2
      Else
         stCon = stCon & " and cp27<=" & stDate2
      End If
   End If
   If stDate1 <> "" Or stDate2 <> "" Then
      If txt1(0) = "1" Then
         pub_QL05 = pub_QL05 & ";收文" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/11/16
      Else
         pub_QL05 = pub_QL05 & ";發文" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/11/16
      End If
   End If
   
   'FCP
   If Combo1.ListIndex = 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & "FCP" 'Add By Sindy 2010/11/16
      stTable = "select cp09,decode(pa75,null,cu10,fa10) fa10,decode(pa75,null,'B',fa76) fa76,substr(decode(pa75,null,cu10,fa10),1,1) s" & _
         " From caseprogress, Patent, fagent, customer" & _
         " where cp01||'' in ('P','FCP','CFP') and cp10||'' in ('101','102','103') and cp09<'B'" & _
         " and cp57 is null and substr(cp12,1,2)='F2'" & stCon & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   'FCT
   Else
      pub_QL05 = pub_QL05 & ";" & Label3 & "FCT" 'Add By Sindy 2010/11/16
      stTable = " select cp09,decode(tm44,null,cu10,fa10) fa10,decode(tm44,null,'B',fa76) fa76,substr(decode(tm44,null,cu10,fa10),1,1) s" & _
         " From caseprogress, Trademark, fagent, customer" & _
         " where cp01 in ('T','FCT','CFT') and cp10='101' and cp09<'B'" & _
         " and cp57 is null and substr(cp12,1,2)='F1'" & stCon & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
         " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9)" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)"
   End If
   
   '洲
   If txt1(4) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & "2.洲" 'Add By Sindy 2010/11/16
      strExc(0) = "select '',decode(x1,'0','亞洲','1','美洲','歐非洲'),x2||'/'||x3,round(100*x2/x3)||'%',x2,x3" & _
         " from (select decode(s,'3','2',s) x1,sum(decode(fa76,'B',1,0)) x2" & _
         ",count(*) x3 From (" & stTable & ")" & _
         " group by decode(s,'3','2',s)) X" & _
         " order by x3 desc,3 desc,1"
   '國家
   Else
      pub_QL05 = pub_QL05 & ";" & Label6 & "1.國家" 'Add By Sindy 2010/11/16
      strExc(0) = "select '',na03,x2||'/'||x3,round(100*x2/x3)||'%',x2,x3" & _
         " from (select substr(fa10,1,3) x1,sum(decode(fa76,'B',1,0)) x2" & _
         ",count(*) x3 From (" & stTable & ")" & _
         " group by substr(fa10,1,3)) X,nation" & _
         " where na01(+)=x1 order by x3 desc,3 desc,1"
   End If
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If fnSaveParentForm(Me) Then
         InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/11/16
         With frm100128_2
         .SetGrid RsTemp, Val(txt1(4))
         If txt1(0) = "1" Then
            .lblCondition(0) = "收文"
         Else
            .lblCondition(0) = "發文"
         End If
         .lblCondition(1) = txt1(1) & " ~ " & txt1(2)
         .lblCondition(2) = Combo1.Text
         .lblCondition(3) = txt1(3)
         If IsUserHasRightOfFunction(Me.Name, strPrint, False) Then
            .cmdOK(2).Enabled = True
         Else
            .cmdOK(2).Enabled = False
         End If
         .lblMemo = Me.lblMemo
         End With
         Me.Hide
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/16
      ShowNoData
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      Select Case Index
         Case 0, 4
            If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
               KeyAscii = 0
               Beep
            End If
            
      End Select
   End If
End Sub
