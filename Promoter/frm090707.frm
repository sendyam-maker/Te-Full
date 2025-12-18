VERSION 5.00
Begin VB.Form frm090707 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖超時案件查詢"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3480
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   1
      Top             =   708
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1245
      MaxLength       =   7
      TabIndex        =   0
      Top             =   708
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   1485
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   750
      Width           =   1005
   End
   Begin VB.Line Line2 
      X1              =   1815
      X2              =   2970
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frm090707"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/28 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/9/25
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
   Case 0
      Unload Me
   Case 1
      Process
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090707 = Nothing
End Sub

Private Sub Process()
   Dim adoRst As ADODB.Recordset
   If txt1(0) = "" Then
      MsgBox "收文日條件不可空白！"
      txt1(0).SetFocus
      Exit Sub
   End If
   If txt1(1) = "" Then
      MsgBox "收文日條件不可空白！"
      txt1(1).SetFocus
      Exit Sub
   End If
   If Val(txt1(1)) < Val(txt1(0)) Then
      MsgBox "收文日條件輸入錯誤！"
      txt1(0).SetFocus
      Exit Sub
   End If
   
   strExc(0) = "select sqldatet(cp05) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C2,pa05 C3,cpm03 C4,s1.st02 C5" & _
      ",substr(cp64,1,2) C6,cp113 C7,cp103*cp104 C8,s2.st02 C9,nvl(cu04,cu05) C10" & _
      " from caseprogress,patent,casepropertymap,staff s1,staff s2,customer" & _
      " where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp10='943'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and s1.st01(+)=cp14 and s2.st01(+)=cp13" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Load frm090707_1
      With frm090707_1
      Set .GrdDataList.Recordset = adoRst.Clone
      .SetDataListWidth
      .txt1(0) = Me.txt1(0)
      .txt1(1) = Me.txt1(1)
      .Show
      End With
      Me.Hide
   Else
      MsgBox "查無資料！"
   End If
   Set adoRst = Nothing
End Sub
