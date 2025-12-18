VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050324_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "內商-國外FC帳款明細表"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   2
      Left            =   3870
      TabIndex        =   4
      Top             =   630
      Width           =   1395
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   2025
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   360
      Width           =   2025
   End
End
Attribute VB_Name = "frm050324_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/02/23 Form2.0已修改 grd1
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
frm050324.Show
Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   lbl1(0).Caption = "系統類別: " & frm050324.Text1.Text
   lbl1(1).Caption = "帳款日期: " & frm050324.MaskEdBox1.Text & " ~ " & frm050324.MaskEdBox2.Text & " " & IIf(frm050324.Text3 = "2", "(應收帳款)", "(請款)")
   Set adoquery = New ADODB.Recordset
   adoquery.CursorLocation = adUseClient
   strSql = "select R21302,R21313,R21304||' '||R21305,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as FAName,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CUName,R21303 from accrpt213, acc1k0, fagent,customer where r21302 = a1k01 and substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and substr(R21314,1,8)=cu01(+) and substr(R21314,9,1)=cu02(+) and r21301 = '" & strUserNum & "' order by r21301 asc, a1k01 asc"
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount > 0 Then
      InsertQueryLog (adoquery.RecordCount) 'Add By Sindy 2010/10/4
      Set GRD1.Recordset = adoquery
      lbl1(2).Caption = "共 " & Trim(adoquery.RecordCount) & " 筆"
      adoquery.Close
      adoquery.CursorLocation = adUseClient
      'Modify By Sindy 2012/12/12
      'adoquery.Open "select sum(r21305) from accrpt213 where r21301 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select r21304,sum(r21305) from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 order by r21304", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         'lbl1(2).Caption = lbl1(2).Caption & "   TOTAL：USD " & Format(Val(CheckStr(adoquery.Fields(0).Value)), FDollar)
         adoquery.MoveFirst
         lbl1(2).Caption = "TOTAL："
         Do While adoquery.EOF = False
            lbl1(2).Caption = lbl1(2).Caption & adoquery.Fields(0).Value & " " & Format(Val(CheckStr(adoquery.Fields(1).Value)), FDollar) & " / "
            adoquery.MoveNext
         Loop
         lbl1(2).Caption = Trim(lbl1(2).Caption)
         lbl1(2).Caption = Left(lbl1(2).Caption, Len(lbl1(2).Caption) - 1)
      '2012/12/12 End
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/4
   End If
   adoquery.Close
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050324_1 = Nothing
End Sub

Private Sub SetDataListWidth()
With GRD1
    .Cols = 6
    .row = 0
    .col = 0: .Text = "請款編號"
    .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "本所案號"
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "請款外幣"
    .ColWidth(2) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "代理人"
    .ColWidth(3) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "申請人"
    .ColWidth(4) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "請款日期"
    .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
End With
End Sub
