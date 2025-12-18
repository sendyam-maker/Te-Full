VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090113 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖形查名單未查覆統計表"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2970
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4125
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   7276
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "開始列印(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frm090113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Memo by Lydia 2015/06/16 隱藏不用;
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim iPrint As Integer
Dim PLeft(1) As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
        PrintData
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me '將畫面移至中央
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090113 = Nothing
End Sub

Private Sub PrintData()
Dim strSql As String
Dim arrTM09 As Variant
Dim strTM09 As String
Dim strField As String
Dim StrUno As String
Dim strUName As String
Dim i As Integer
Dim ii As Integer
Dim SeekExt As Boolean
Dim j As Integer
Dim k As Integer
grd1.Clear
grd1.Rows = 0
grd1.ColWidth(0) = 600
grd1.ColWidth(1) = 600
grd1.ColWidth(2) = 8200
strSql = "select st03,st01,st02 ,tmq10,tmq03,tmq09 from trademarkquery,staff where tmq10=st01(+) and tmq11 is null and tmq09 is not null and tmq09 >0 and tmq04>=20040301  order by st03,st01"
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
            StrUno = CheckStr(.Fields("st01").Value)
            strUName = CheckStr(.Fields("st02").Value)
            SeekExt = False
            For i = 0 To grd1.Rows - 1
                grd1.row = i
                grd1.col = 0
                If Trim(grd1.Text) = StrUno Then
                    SeekExt = True
                    Exit For
                End If
            Next i
            If SeekExt = False Then
                grd1.Rows = grd1.Rows + 1
                grd1.row = grd1.Rows - 1
                grd1.col = 0
                grd1.Text = StrUno
                grd1.col = 1
                grd1.Text = strUName
            End If
            grd1.col = 2
            If InStr(1, grd1.Text, CheckStr(.Fields("tmq03").Value)) = 0 Then
                grd1.Text = grd1.Text & CheckStr(.Fields("tmq03").Value) & "."
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/14
    End If
End With
CheckOC
'sort
InsertQueryLog (grd1.Rows) 'Add By Sindy 2010/12/14
With grd1
    For i = 0 To .Rows - 1
        .row = i
        .col = 2
        If Right(.Text, 1) = "." Then .Text = Left(.Text, Len(.Text) - 1)
        arrTM09 = Split(.Text, ".")
        For ii = 0 To UBound(arrTM09)
            For j = 0 To UBound(arrTM09)
                If j <> UBound(arrTM09) Then
                    For k = j To UBound(arrTM09)
                        If k <> UBound(arrTM09) Then
                            If arrTM09(k) > arrTM09(k + 1) Then
                                strTM09 = arrTM09(k)
                                arrTM09(k) = arrTM09(k + 1)
                                arrTM09(k + 1) = strTM09
                                .Text = Join(arrTM09, ".")
                            End If
                        End If
                    Next k
                End If
            Next j
        Next ii
    Next i
End With
PrintTitle
PrintDatil
Printer.EndDoc
MsgBox "列印成功！", , "感謝！"
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth("圖形查名單未查覆統計表") / 2)
Printer.CurrentY = iPrint
Printer.Print "圖形查名單未查覆統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：1"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub
Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1800
End Sub
Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "查名人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "未查覆組群"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Dim i As Integer
For i = 0 To grd1.Rows - 1
    grd1.row = i
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iPrint
    grd1.col = 1
    Printer.Print grd1.Text
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    grd1.col = 2
    Printer.Print grd1.Text
    iPrint = iPrint + 600
Next i
End Sub
