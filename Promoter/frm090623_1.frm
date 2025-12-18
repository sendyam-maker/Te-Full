VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090623_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "選擇收文號"
   ClientHeight    =   3696
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3696
   ScaleWidth      =   5880
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3180
      Left            =   15
      TabIndex        =   2
      Top             =   480
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   5609
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
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
      Caption         =   "取消"
      Height          =   375
      Index           =   1
      Left            =   4950
      TabIndex        =   1
      Top             =   45
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4005
      TabIndex        =   0
      Top             =   45
      Width           =   870
   End
End
Attribute VB_Name = "frm090623_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Public oCP01 As String
Public oCP02 As String
Public oCP03 As String
Public oCP04 As String
Dim StrSqlB As String

Public Function Process() As Boolean
Process = False
CheckOC3
'2009/7/8 MODIFY BY SONIA 加入未取消收文的條件 CFP-019704
'modify by sonia 2014/7/3 P-108206
'StrSqlB = "select " & SQLDate("CP05") & " as 收文日,cp09 as 收文號,nvl(cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員 from caseprogress,casepropertymap,staff s1,staff s2 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) AND CP57 IS NULL "
'Modify By Sindy 2014/7/29 +order by cp05 desc
'Modified by Lydia 2025/10/08 +and nvl(cp16,0)>0 and nvl(cp20,'Y') <> 'N'=>在點選收文號或輸入收文號時，程式檢查不可為沒有收費的收文號。Ex. P-134497
StrSqlB = "select " & SQLDate("CP05") & " as 收文日,cp09 as 收文號,decode(pa09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員 from caseprogress,casepropertymap,staff s1,staff s2,patent " & _
          "where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
          "and nvl(cp16,0)>0 and nvl(cp20,'Y') <> 'N' order by cp05 desc"
With AdoRecordSet3
   .CursorLocation = adUseClient
   .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      Set GRD1.Recordset = AdoRecordSet3
   Else
      Exit Function
   End If
End With
Process = True
End Function

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer
Select Case Index
Case 0
      For i = 1 To GRD1.Rows - 1
         GRD1.row = i
         GRD1.col = 0
         If GRD1.CellBackColor = QBColor(13) Then
            frm090623.txtSH(14).Text = GRD1.TextMatrix(i, 1)
            Unload Me
            Exit Sub
         End If
      Next i
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub GRD1_DblClick()
Dim i As Integer
Dim j As Integer
Dim k  As Integer
j = GRD1.MouseRow
For i = 1 To GRD1.Rows - 1
   GRD1.row = i
   For k = 0 To GRD1.Cols - 1
      GRD1.col = k
      GRD1.CellBackColor = QBColor(15)
   Next k
Next i
GRD1.row = j
   
   For k = 0 To GRD1.Cols - 1
      GRD1.col = k
      GRD1.CellBackColor = QBColor(13)
   Next k
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090623_1 = Nothing
End Sub

