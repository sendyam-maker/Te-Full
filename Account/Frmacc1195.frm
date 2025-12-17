VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc1195 
   AutoRedraw      =   -1  'True
   Caption         =   "銷帳退費收文號查詢"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   6375
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
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
      _Band(0).Cols   =   5
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -12
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc1195"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Create by Lydia 2020/04/16 銷帳退費收文號查詢(法務工作點數分配)
Option Explicit

Dim intLastRow As Integer '記錄勾選最後一筆
Public m_A0J13 As String '收據號碼


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 6500
   Me.Height = 4400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim intX As Integer
Dim strCP09 As String

   With MSHFlexGrid1
       For intX = 1 To .Rows - 1
          If .TextMatrix(intX, 0) = "v" And "" & .TextMatrix(intX, 2) <> "" Then
             strCP09 = strCP09 & "," & .TextMatrix(intX, 2)
          End If
       Next
   End With
   
   If strCP09 <> "" Then '選擇單筆或多筆收文號進入工作點數畫面
        Set frm071021.m_PrevForm = Frmacc1190
        frm071021.m_bolPrev = True
        frm071021.m_KeyList = Mid(strCP09, 2)
        Frmacc1190.Enabled = False
        frm071021.Show
   Else
       StatusClear
       tool1_enabled
       Frmacc1190.Enabled = True
       Frmacc1190.Show
   End If
   
   Set Frmacc1195 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         OpenTable
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub OpenTable()
Dim rsRead As New ADODB.Recordset
Dim intJ As Integer
Dim strQ1 As String

On Error GoTo Checking

    Call SetGrd(True) '清空
     
    If m_A0J13 = "" Then
          MsgBox "查無資料!!"
          Unload Me
          Exit Sub
    End If
    strQ1 = "select ' ' as v,  cp01||cp02||cp03||cp04 as caseno, cp09,decode(decode(lc01,null,'000',lc15),'000',cpm03,cpm04) as cpm0304,a0j09,a0j10 " & _
                "From acc0j0, caseprogress, casepropertymap, lawcase, hirecase " & _
                "where a0j13='" & m_A0J13 & "' and cp09(+)=a0j01 and cp01=cpm01(+) and cp10=cpm02(+) and nvl(cp18,0)>0 " & _
                "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) "
    strQ1 = strQ1 & "order by caseno,cp09 "
    intJ = 1
    Set rsRead = ClsLawReadRstMsg(intJ, strQ1)
    If intJ = 1 Then
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsRead
         Call SetGrd
    Else
          MsgBox "查無資料!!"
          Unload Me
          Exit Sub
    End If
    
    Set rsRead = Nothing
    Exit Sub
    
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer
Dim lngColor As Long

   'V,本所案號,總收文號,案件性質,服務費,規費
    arrGridHeadText = Array("v", "本所案號", "總收文號", "案件性質", "服務費", "規費")
   arrGridHeadWidth = Array(300, 1200, 1000, 1200, 900, 900)
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
    For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    Next

   For iR = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.row = iR
        For iRow = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.col = iRow
           MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
        Next iRow
   Next iR
   MSHFlexGrid1.Visible = True
End Sub

Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
       GridClick MSHFlexGrid1, intLastRow, 0, 1 '可複數
    End If
End Sub
