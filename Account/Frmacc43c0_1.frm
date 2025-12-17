VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc43c0_1 
   AutoRedraw      =   -1  'True
   Caption         =   "轉撥檢查"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   7875
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7600
      _ExtentX        =   13414
      _ExtentY        =   8652
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   9
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "Frmacc43c0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/01 Form2.0已修改 grdDataList
'Create by Amy 2016/01/11
Option Explicit
Public m_YearMon As String '前畫面查詢年月(民國年)
Dim strColN, intWidth

Private Sub Form_Activate()
     strFormName = Name
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single

'    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 8000
    Me.Height = 5625
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'    Image1 = LoadPicture(strBackPicPath3)
'    sglWidth = Image1.Width
'    sglHeight = Image1.Height
'    For intX = 0 To Int(ScaleWidth / sglWidth)
'        For intY = 0 To Int(ScaleHeight / sglHeight)
'            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'        Next
'    Next
    'SetDataListWidth
    doQuery
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_YearMon = MsgText(601)
    tool3_enabled
    Frmacc43c0.Enabled = True
    Set Frmacc43c0_1 = Nothing
End Sub

Private Sub doQuery()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    strQ = "Select * From (" & _
                "Select a0902,st01,st02,sp19,sp20,sp40,sp41,1 as Sort,a0901 From SalesPoint,Acc090,Staff " & _
                "Where sp01=" & Val(m_YearMon) + 191100 & " And sp48=a0901(+) And sp02=st01(+) " & _
                " And (sp19<>0 Or (sp20 is not null And sp20>'') Or sp40<>0 Or (sp41 is not null And sp41>'')) " & _
    "Union Select '合計' as a0902,'' as st01,'' as st02,Sum(sp19) as sp19, '' as sp20,Sum(sp40) as sp40,'' as sp41,2 as Sort,'' as a0901 From SalesPoint " & _
                "Where sp01=" & Val(m_YearMon) + 191100 & ") Order by Sort,a0901,st02 "
 
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    If RsQ.RecordCount > 0 Then
        Set grdDataList.Recordset = RsQ
    Else
        grdDataList.Rows = 2
        MsgBox "資料庫中搜尋不到符合資料!!", , "沒有資料"
    End If
    SetDataListWidth
    RsQ.Close
End Sub

Private Sub SetDataListWidth()
    Dim iCol As Integer
   
    ReDim strColN(0 To 8)
    ReDim intWidth(0 To 8)
    strColN = Array("業務區", "員工編號", "姓名", "實績轉撥點數", "實績備註", "結餘轉撥點數", "結餘備註" _
                    , "Sort", "a0901")
    intWidth = Array(700, 800, 800, 1200, 1300, 1200, 1300, 0, 0)
    
    With grdDataList
        .Visible = False

        For iCol = 0 To UBound(strColN)
            .ColWidth(iCol) = intWidth(iCol)
            .TextMatrix(0, iCol) = strColN(iCol)
        Next
        .Visible = True
    End With
End Sub
