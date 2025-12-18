VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm071013_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "尚未到期的庭期資料"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8370
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7020
      TabIndex        =   1
      Top             =   60
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定取消庭期(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   5310
      TabIndex        =   0
      Top             =   60
      Width           =   1680
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3810
      Left            =   30
      TabIndex        =   2
      Top             =   480
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm071013_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
Option Explicit

Public m_CP01 As String
Public m_CP02 As String
Public m_CP03 As String
Public m_CP04 As String


Private Sub cmdok_Click(Index As Integer)
Dim i As Integer
   
   Select Case Index
      Case 0
         '檢查資料
         frm071013.m_strCancelRecv = ""
         For i = 1 To MSHFlexGrid1.Rows - 1
            If Trim(MSHFlexGrid1.TextMatrix(i, 0)) = "V" Then
               frm071013.m_strCancelRecv = frm071013.m_strCancelRecv & Trim(MSHFlexGrid1.TextMatrix(i, 2)) & ","
            End If
         Next i
         If frm071013.m_strCancelRecv = "" Then
            MsgBox "請勾選資料！", vbInformation
            Exit Sub
         End If
      Case 1 '回前畫面
   End Select
   Me.Hide
End Sub

Public Function doQuery() As Boolean
Dim LcTmp As String
   
   doQuery = False
   
   Me.Caption = m_CP01 + "-" + m_CP02
   If m_CP03 <> "0" Or m_CP04 <> "00" Then
      Me.Caption = Me.Caption + "-" + m_CP03 + "-" + m_CP04
   End If
   Me.Caption = Me.Caption + "尚未到期的庭期資料"
   
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   LcTmp = m_CP01 + m_CP02 + m_CP03 + m_CP04
   If m_CP01 = "L" Or m_CP01 = "FCL" Or m_CP01 = "CFL" Or m_CP01 = "LIN" Then
      strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
         "cp09,decode(lc15,020,cpm04,cpm03),decode(CP13," + _
         "S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02)," + _
         "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
         "decode(cp71,or01,or02),CP64,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
         "STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
         ChgLawcase(LcTmp) + " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
         "cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and cp09 in (select cp09 from caseprogress,courtyardperiod where CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' and cp09=cdp01 and cdp03>=" & strSrvDate(1) & " and cdp18 is null)"
   ElseIf m_CP01 = "LA" Then
      strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
         "cp09,cpm03,decode(CP13,S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02)," & _
         "decode(CP29,S3.ST01,S3.ST02)," + _
         "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
         "decode(cp71,or01,or02),CP64,hc05,hc06 from caseprogress, hirecase,STAFF S1,STAFF S2,STAFF S3, " + _
         "CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " & ChgHirecase(LcTmp) + _
         " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) " + _
         "and cp09 in (select cp09 from caseprogress,courtyardperiod where CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' and cp09=cdp01 and cdp03>=" & strSrvDate(1) & " and cdp18 is null)"
   End If
   strExc(0) = strExc(0) & " AND CP10<>'0' Order by cp27 DESC,CP09"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set MSHFlexGrid1.Recordset = RsTemp
      doQuery = True
   Else
      MSHFlexGrid1.Rows = 2
   End If
   GridHead
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm071013_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_SelChange()
Dim m_row As Integer, i As Integer
Dim m_mouseRow As Integer

MSHFlexGrid1.Visible = False
m_mouseRow = MSHFlexGrid1.MouseRow
MSHFlexGrid1.col = 0
If m_mouseRow <> 0 Then
'    If m_row <> 0 Then
''        grd1.row = m_row
'         For i = 0 To grd1.Cols - 1
'              grd1.col = i
'              If grd1.CellBackColor = &HFFC0C0 Then
'                grd1.CellBackColor = &H80000018
'                grd1.TextMatrix(m_row, 0) = ""
'              Else
'                grd1.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
'                grd1.TextMatrix(m_row, 0) = "V"
'              End If
'        Next i
'    End If
'    If m_row <> m_mouseRow Then
        MSHFlexGrid1.row = m_mouseRow
        m_row = m_mouseRow
         For i = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.col = i
              If MSHFlexGrid1.CellBackColor = &HFFC0C0 Then
                MSHFlexGrid1.CellBackColor = &H80000018
                MSHFlexGrid1.TextMatrix(m_row, 0) = ""
              Else
                MSHFlexGrid1.CellBackColor = &HFFC0C0
                MSHFlexGrid1.TextMatrix(m_row, 0) = "V"
              End If
        Next i
'    Else
'        m_row = 0
'    End If
End If
MSHFlexGrid1.Visible = True
End Sub

Private Sub GridHead()
   With MSHFlexGrid1
      .Visible = False
      .Cols = 14
      .row = 0
      .col = 0
      .Visible = True
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 900: .Text = "收文日"
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .col = 3: .ColWidth(3) = 1000: .Text = "案件性質"
      .col = 4: .ColWidth(4) = 900: .Text = "智權人員"
      .col = 5: .ColWidth(5) = 900: .Text = "承辦人"
      'Modified by Lydia 2015/10/05
      '.col = 6: .ColWidth(6) = 900: .Text = "法務人員"
      .col = 6: .ColWidth(6) = 900: .Text = "協辦人員"
      .col = 7: .ColWidth(7) = 900: .Text = "發文日"
      .col = 8: .ColWidth(8) = 1200: .Text = "法院"
      .col = 9: .ColWidth(9) = 1500: .Text = "進度備註"
      .col = 10: .ColWidth(10) = 0
      .col = 11: .ColWidth(11) = 0
      .col = 12: .ColWidth(12) = 0
      .col = 13: .ColWidth(13) = 0
      .CellAlignment = flexAlignCenterCenter
      '判斷是否有資料
      If .Rows > 1 Then .row = 1
      .Visible = True
   End With
End Sub
