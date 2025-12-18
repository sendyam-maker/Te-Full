VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090801_10 
   BorderStyle     =   1  '單線固定
   Caption         =   "選取資料"
   ClientHeight    =   4230
   ClientLeft      =   2790
   ClientTop       =   3720
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7305
   Begin VB.CommandButton Command1 
      Caption         =   "回上一頁"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6120
      TabIndex        =   2
      Top             =   10
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   5040
      TabIndex        =   0
      Top             =   10
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   16
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
      _Band(0).Cols   =   16
   End
End
Attribute VB_Name = "frm090801_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (grdDataList)
'Create by Amy 2016/09/01
Option Explicit

Dim m_PrevForm As Form '前畫面
Public strTM12 As String '前畫面傳入申請案號
Public strTM15 As String '前畫面傳入審定號數

Dim i As Integer
Dim strColN(), intWidth()

Dim iStatus As String 'Added by Lydia 2020/12/15 模式：1 =>  601異議/603評定/605廢止新案商標案控制;
                                                                                    '2 => CFT緬甸重新申請案
Private Sub cmdOK_Click()
    For i = 1 To grdDataList.Rows - 1
        If grdDataList.TextMatrix(i, 0) = "V" Then
            With m_PrevForm 'frm090801
                .strCaseNo1 = grdDataList.TextMatrix(i, GetValue("tm01"))
                .strCaseNo2 = grdDataList.TextMatrix(i, GetValue("tm02"))
                .StrCaseNo3 = grdDataList.TextMatrix(i, GetValue("tm03"))
                .strCaseNo4 = grdDataList.TextMatrix(i, GetValue("tm04"))
                'Modified by Lydia 2020/12/15 +判斷
                '.strTM28 = grdDataList.TextMatrix(i, GetValue("tm28"))
                If iStatus = "1" Then .strTM28 = grdDataList.TextMatrix(i, GetValue("tm28"))
            End With
            Unload Me
            Exit Sub
        End If
    Next i
End Sub

Private Sub Command1_Click()
    'Added by Lydia 2020/12/15 CFT緬甸重新申請案：判斷未選取
    If iStatus = "2" And (m_PrevForm.strCaseNo1 = "" Or m_PrevForm.strCaseNo2 = "" Or m_PrevForm.StrCaseNo3 = "" Or m_PrevForm.strCaseNo4 = "") Then
        If MsgBox("未選取資料，是否重新選擇緬甸商標註冊案？", vbYesNo + vbInformation + vbDefaultButton1, "CFT緬甸重新申請案") = vbYes Then
           Exit Sub
        End If
    End If
    'end 2020/12/15
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strTM12 = ""
    strTM15 = ""
    Set m_PrevForm = Nothing
    Set frm090801_10 = Nothing
End Sub

Public Function GetTM1215Rec(Optional ByRef strCaseNo1 As String, Optional ByRef strCaseNo2 As String, Optional ByRef StrCaseNo3 As String, Optional ByRef strCaseNo4 As String, Optional ByRef strTM28 As String) As Integer
'Memo by Lydia 2020/12/15 601異議/603評定/605廢止新案商標案控制
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    SetDataListWidth
    GetTM1215Rec = 0: strCaseNo1 = "": strCaseNo2 = "": StrCaseNo3 = "": strCaseNo4 = "": strTM28 = ""
    strQ = "Select ' ' AS V,Decode(tm28,'1','','N')||tm01 ||'-'|| tm02 ||'-'|| tm03 ||'-'|| tm04||Decode(tm29,'Y','＊','')||Decode(length(Nvl(tm57,'')),null,'','●') AS 本所案號," & _
                "Decode(length(Nvl(tm73,'')),null,'','●')||tm34 as 分所號,Nvl(Nvl(tm05,tm06),tm07) AS 案件名稱,NA03 AS 申請國家,tm09 AS 商品類別," & _
                "Nvl(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日, " & _
                "tm01,tm02,tm03,tm04,tm28 " & _
                "From TradeMark,Nation,Customer " & _
                "Where SubStr(tm23,1,8)=CU01(+) And Decode(SubStr(tm23,9,1),'','0',SubStr(tm23,9,1))=CU02(+) And tm10=na01(+) " & _
                "And (tm12='" & strTM12 & "' Or tm15='" & strTM15 & "' )"
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If RsQ.RecordCount > 0 Then
        GetTM1215Rec = RsQ.RecordCount
        If RsQ.RecordCount = 1 Then
            grdDataList.Rows = 2
            strCaseNo1 = RsQ.Fields("tm01")
            strCaseNo2 = RsQ.Fields("tm02")
            StrCaseNo3 = RsQ.Fields("tm03")
            strCaseNo4 = RsQ.Fields("tm04")
            strTM28 = RsQ.Fields("tm28")
        End If
        Set grdDataList.Recordset = RsQ
    End If
    RsQ.Close
End Function

Private Sub GrdDataList_Click()
    Dim intRow As Integer
    Dim j As Integer
    
    grdDataList.Visible = False
    grdDataList.col = 0
    grdDataList.row = grdDataList.MouseRow
    intRow = grdDataList.row
    
    If grdDataList.row <> 0 Then
        For i = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(i, 0) = "V" And i <> intRow Then
                grdDataList.TextMatrix(i, 0) = ""
                grdDataList.row = i
                For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = QBColor(15)
                Next j
            ElseIf i = intRow Then
                grdDataList.TextMatrix(i, 0) = "V"
                grdDataList.row = i
                For j = 0 To grdDataList.Cols - 1
                    grdDataList.col = j
                    grdDataList.CellBackColor = &HFFC0C0
                Next j
            End If
        Next i
    End If
    grdDataList.Visible = True
End Sub

'Modified by Lydia 2020/12/15 模式
Public Sub SetParent(ByRef fm As Form, ByRef pType As String)
   Set m_PrevForm = fm
   iStatus = pType 'Added by Lydia 2020/12/15
End Sub

Private Function GetValue(pRowN As String) As Integer
    Dim j As Integer
 
    For j = 1 To UBound(strColN)
       If UCase(strColN(j)) = UCase(pRowN) Then
          GetValue = j
          Exit For
       End If
    Next j
End Function

Private Sub SetDataListWidth()
    Dim iCol As Integer
   
    ReDim strColN(12)
    ReDim intWidth(12)
    strColN = Array("V", "本所案號", "分所號", "案件名稱", "申請國家", "商品類別", "申請人", "專用期止日", _
                    "tm01", "tm02", "tm03", "tm04", "tm28")
    intWidth = Array(200, 1000, 1000, 2000, 600, 500, 800, 500, _
                    0, 0, 0, 0, 0)
    
    With grdDataList
        .Visible = False

        For iCol = 0 To UBound(strColN)
            .ColWidth(iCol) = intWidth(iCol)
            .TextMatrix(0, iCol) = strColN(iCol)
        Next
        .Visible = True
    End With
End Sub

'Added by Lydia 2020/12/15 CFT緬甸重新申請案
Public Function GetDataMode2(ByVal strCU01 As String, Optional ByRef strCaseNo1 As String, Optional ByRef strCaseNo2 As String, Optional ByRef StrCaseNo3 As String, Optional ByRef strCaseNo4 As String) As Integer
Dim RsQ As New ADODB.Recordset
Dim strQ As String
    
    'CFT緬甸新「申請」案，若該客戶有CFT緬甸案且申請日<20210101且有審定號者，在列印接洽單前詢問「是否有相關緬甸商標註冊案」，
    '智權可選擇「是」或「否」，如選擇「是」，若該客戶只有一筆符合條件的舊案則直接帶出該案號；若有多筆則開畫面讓智權人員選擇。
    'P.S 因為客戶通知函只要有專用期間就列入通知，所以不限制銷／閉卷或超過專用期間
    GetDataMode2 = 0: strCaseNo1 = "": strCaseNo2 = "": StrCaseNo3 = "": strCaseNo4 = ""
    strQ = "Select ' ' AS V,Decode(tm28,'1','','N')||tm01 ||'-'|| tm02 ||'-'|| tm03 ||'-'|| tm04||Decode(tm29,'Y','＊','')||Decode(length(Nvl(tm57,'')),null,'','●') AS 本所案號," & _
                "Decode(length(Nvl(tm73,'')),null,'','●')||tm34 as 分所號,Nvl(Nvl(tm05,tm06),tm07) AS 案件名稱,NA03 AS 申請國家,tm09 AS 商品類別," & _
                "Nvl(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日, " & _
                "tm01,tm02,tm03,tm04,tm28 " & _
                "From TradeMark,Nation,Customer " & _
                "Where tm01='CFT' and tm10='048' and tm23=" & CNULL(ChangeCustomerL(strCU01)) & " and tm11<20210101 and tm15 is not null " & _
                "And SubStr(tm23,1,8)=CU01(+) And Decode(SubStr(tm23,9,1),'','0',SubStr(tm23,9,1))=CU02(+) And tm10=na01(+) "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If RsQ.RecordCount > 0 Then
        GetDataMode2 = RsQ.RecordCount
        If RsQ.RecordCount = 1 Then  '單筆：預設
            strCaseNo1 = "" & RsQ.Fields("tm01")
            strCaseNo2 = "" & RsQ.Fields("tm02")
            StrCaseNo3 = "" & RsQ.Fields("tm03")
            strCaseNo4 = "" & RsQ.Fields("tm04")
        Else                                     '多筆：畫面選擇
            SetDataListWidth
            grdDataList.Rows = 2
            Set grdDataList.Recordset = RsQ
        End If
    End If
    RsQ.Close
    
End Function
