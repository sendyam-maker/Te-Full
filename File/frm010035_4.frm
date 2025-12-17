VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010035_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "一週內新書公告"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7500
   Begin VB.CommandButton cmdButton 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   45
      Width           =   850
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2700
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   4763
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "編號|書名|ISBN|作者|譯者|類別|保管人|狀態|借閱人|借閱/延期日|上架日|出刊日"
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
      _Band(0).Cols   =   13
   End
End
Attribute VB_Name = "frm010035_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 GrdDataList
'2016/10/03 Create by Amy
Option Explicit

Dim i As Integer
Dim arrField, intWidth

Private Sub cmdButton_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    ReDim arrField(11)
    ReDim intWidth(11)
    arrField = Array("編號", "書名", "ISBN", "作者", "譯者", "類別", "保管人", "狀態", "借閱人", "借閱/延期日", "上架日", "出刊日")
    intWidth = Array(500, 1300, 1000, 1000, 1000, 500, 700, 1000, 700, 1000, 1000, 1000)
                                
    MoveFormToCenter Me
End Sub

Public Function QueryRecord() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strDate As String
    
    QueryRecord = False
    '抓取工作日5日內上架新書(FrmLogin有改也要改)
    strDate = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -5) + 19110000
    strQ = "Select BK01 as 編號,Decode(BK04,null,BK05,Decode(BK05,null,BK04,BK04||'('||BK05||')')) as 書名,Nvl(BK02,'') as ISBN," & _
                "Decode(BK06,null,BK07,Decode(BK07,null,BK06,BK06||'('||BK07||')')) as 作者,Nvl(BK08,'') as 譯者," & _
                "Decode(BK03,'1','專利','2','商標','3','法律','4','電腦','5','其他') as 類別,k.ST02 as 保管人,BK12 as  狀態," & _
                "Decode(LR02,'Z','',b.ST02) as 借閱人,Decode(LR02,'Z','',Decode(LR06,null,'',sqldatet(LR04))) as 借閱日," & _
                "sqldatet(BK09) as 上架日,Decode(BK13,null,'',sqldatet(BK13)) as 出刊日 " & _
                "From (Select * From BooksData Where BK09>=" & strDate & " And BK09<=" & strSrvDate(1) & ")," & _
                "(Select * From LoanRecord a Where LR01||LR02=(Select Max(LR01||LR02) From LoanRecord  Where a.LR03=LR03 ))" & _
                ",Staff k,Staff b Where BK01=LR03(+) And BK10=k.ST01(+) And LR08=b.ST01(+) "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    If RsQ.RecordCount > 0 Then
        QueryRecord = True
        grdDataList.Rows = 2
        Set grdDataList.Recordset = RsQ
        SetGridWidth
    End If
    
    RsQ.Close
    Set RsQ = Nothing
End Function

Private Sub SetGridWidth()
   
    '設欄寬
    With grdDataList
        .FormatString = .FormatString
        For i = LBound(intWidth) To UBound(intWidth)
            .ColWidth(i) = intWidth(i)
            If intWidth(i) <> 0 Then .ColAlignment(i) = flexAlignLeftCenter
        Next i
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm010035_4 = Nothing
End Sub
