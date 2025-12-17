VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc1441 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "不催款查詢"
   ClientHeight    =   5535
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6405
   Icon            =   "Frmacc1441.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   6435
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   5520
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   750
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4800
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   8467
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "＊：舊的名稱   $：有呆帳　⊕：不得代理"
      ForeColor       =   &H000000C0&
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Frmacc1441"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/24 Form2.0已修改
'Add By Sindy 2014/9/4
Option Explicit

Dim StrSQLa As String, StrSqlB As String
Dim rsA As New ADODB.Recordset


Private Sub cmdok_Click(Index As Integer)
    Unload Frmacc1441
End Sub

Private Sub Form_Activate()
    SetGridWidth
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   PUB_InitForm Me, Me.Width, Me.Height
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   tool3_enabled
   Frmacc1440.Show
   Set rsA = Nothing
   Set Frmacc1441 = Nothing
End Sub


'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   
On Error GoTo Checking
   
   '若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
   StrSqlB = "DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)) as 名稱,"
   
   '不催款者
   'Modify by Amy 2020/09/18 原:(cu140='N' and cu140 is not null) ,因不寄催款單改輸1-3
   strExc(0) = "Select CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,GetDizhang(CU142) AS 帳款處理情形" & _
                " From Customer,Nation" & _
               " Where cu140 is not null And CU10=NA01(+)"
   strExc(0) = "Select * From (" & strExc(0) & ") Order by 編號 "
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <> 0 And rsA.RecordCount > 0 Then
      Set grdDataList.Recordset = rsA
   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Set rsA = Nothing
End Sub

Private Sub SetGridWidth()
   '設欄寬
   With grdDataList
      .ColWidth(0) = 1000
      .ColAlignment(0) = flexAlignLeftCenter
      .ColWidth(1) = 2800
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 900
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 1300
      .ColAlignment(3) = flexAlignLeftCenter
   End With
End Sub
