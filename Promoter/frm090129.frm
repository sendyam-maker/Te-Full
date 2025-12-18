VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm090129 
   Caption         =   "顯示圖檔／文字"
   ClientHeight    =   6792
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8808
   LinkTopic       =   "Form1"
   ScaleHeight     =   566
   ScaleMode       =   3  '像素
   ScaleWidth      =   734
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1455
      Left            =   5520
      ScaleHeight     =   117
      ScaleMode       =   3  '像素
      ScaleWidth      =   121
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox tmpPic 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   107
      ScaleMode       =   3  '像素
      ScaleWidth      =   227
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.Image tmpImg 
         Height          =   1140
         Left            =   0
         Top             =   0
         Width           =   2610
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      ExtentX         =   8916
      ExtentY         =   11668
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "無資料可讀取"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frm090129"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/10/15 Form2.0已檢查 (無需修改的物件)
Option Explicit
'保存表單的原始寬度
Private FormOldWidth As Long
'保存表單的原始高度
Private FormOldheight As Long

Dim mTQF01 As String, mTQFkind As String
Dim mTQF05 As String
Dim m_PrevForm As Form '前一畫面
Public iStiu As Integer  '1=讀取成功,0=失敗
Dim adoRst As New ADODB.Recordset
Dim stTempFile As String
Dim file_num As Integer
Dim bytes() As Byte
Dim bolInit As Boolean
Dim m_AttachPath As String
Dim m_STMF03 As String 'Added by Lydia 2024/10/15 查名單-網中：圖形查詢附件序號

'Modified by Lydia 2024/10/15 +pTMF03
Public Sub SetParent(ByRef fm As Form, fN01 As String, fN02 As Integer, fN03 As String, Optional ByVal pTMF03 As String)
   Set m_PrevForm = fm
   mTQF01 = fN01: mTQFkind = fN02
   mTQF05 = fN03
   m_STMF03 = pTMF03 'Added by Lydia 2024/10/15
   
End Sub
'在調用ResizeForm前先調用本函數
Public Sub ResizeInit(FormName As Form)
Dim obj As Control

    FormOldWidth = FormName.ScaleWidth
    FormOldheight = FormName.ScaleHeight

On Error Resume Next
    For Each obj In FormName
        obj.Tag = obj.Left & " " & obj.Top & " " _
        & obj.Width & " " & obj.Height & " "
    Next obj
On Error GoTo 0
End Sub

'按比例改變表單內各元件的大小，在調用ReSizeForm前先調用ReSizeInit函數
Public Sub ResizeForm(FormName As Form)
Dim Pos(4) As Double
Dim i As Long, TempPos As Long, StartPos As Long
Dim obj As Control
Dim ScaleX As Double, ScaleY As Double

On Error Resume Next
    '保存表單寬度縮放比例
    ScaleX = FormName.ScaleWidth / FormOldWidth
     '保存表單高度縮放比例
    ScaleY = FormName.ScaleHeight / FormOldheight
   

    For Each obj In FormName
        StartPos = 1
        For i = 0 To 4
            '讀取控制項的原始位置與大小
            TempPos = InStr(StartPos, obj.Tag, " ", vbTextCompare)
            If TempPos > 0 Then
            Pos(i) = Mid(obj.Tag, StartPos, TempPos - StartPos)
            StartPos = TempPos + 1
            Else
            Pos(i) = 0
            End If
            '根據控制項的原始位置及表單改變大小的比例對控制項重新置放與改變大小
            obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i
    Next obj
On Error GoTo 0

End Sub

Private Sub Form_Load()
Dim tmpW As Long, tmpH As Long
Dim tmpX As Single
Dim mStr As String

iStiu = 0
    If Len(mTQF01) > 0 Then
        m_AttachPath = App.path & "\" & strUserNum
        If Dir(m_AttachPath, vbDirectory) = "" Then
           MkDir m_AttachPath
        End If
      If GetAttachFile(mStr) = True Then
        '載入檔案,需要設定表單大小並重新記錄初始值
        bolInit = True
        iStiu = 1
           If UCase(Trim(mStr)) = "PDF" Then
              Me.tmpPic.Visible = False: Me.WebBrowser1.Visible = True
              FormOldWidth = 5280: FormOldheight = 6795
              Me.Width = 5400: Me.Height = 7200
              Call ResizeInit(Me)
           Else
              Me.WebBrowser1.Visible = False: Me.tmpPic.Visible = True
              tmpH = tmpImg.Height: tmpW = tmpImg.Width
            
              '圖檔過大,縮小比例
              tmpX = 1
              Do While tmpH >= 600 Or tmpW >= 800
                 tmpH = tmpH * (IIf(tmpX > 0.1, Format(tmpX - 0.1, "0.0"), Format(tmpX - 0.01, "0.00")))
                 tmpW = tmpW * (IIf(tmpX > 0.1, Format(tmpX - 0.1, "0.0"), Format(tmpX - 0.01, "0.00")))
                 tmpX = IIf(tmpX > 0.1, Format(tmpX - 0.1, "0.0"), Format(tmpX - 0.01, "0.00"))
              Loop
              If tmpX < 1 Then
                 tmpImg.Stretch = True
                 tmpImg.Height = tmpH * tmpX: tmpImg.Width = tmpW * tmpX
                 tmpH = tmpImg.Height: tmpW = tmpImg.Width
              End If
              tmpImg.Top = 3: tmpImg.Left = 3
              tmpPic.Top = 3: tmpPic.Left = 3
              tmpPic.Height = tmpH + 10: tmpPic.Width = tmpW + 10
              FormOldheight = tmpH + 10:   FormOldWidth = tmpW + 10
              Me.Width = Val(Format((tmpW + 20) * (Me.Width / Me.ScaleWidth), "0"))
              Me.Height = Val(Format((tmpH + 35) * (Me.Height / Me.ScaleHeight), "0"))
              Call ResizeInit(Me) '重新記錄初始值
              tmpImg.Stretch = True
           End If
        bolInit = False
        
        'Added by Lydia 2025/04/30 查名單-網中：圖形查詢附件序號
        If m_STMF03 <> "" Then
            Me.Caption = "顯示圖檔"
        Else
        'end 2025/04/30
            Me.Caption = "顯示" & IIf(mTQFkind = TMQ_AkindPic, "圖檔", "文字" & mTQFkind)
        End If
      Else
          bolInit = True
          Me.Caption = "檔案讀取失敗"
          Me.tmpPic.Visible = False: Me.WebBrowser1.Visible = False
          Label1.Visible = True: Label1.Left = 35: Label1.Top = 20
          FormOldWidth = 3280: FormOldheight = 1320
          Me.Width = 3600: Me.Height = 1500

          Call ResizeInit(Me)
          bolInit = False
      End If
    End If
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE '讓表單在最上層
End Sub

Private Sub Form_Resize()
    If bolInit = False Then Call ResizeForm(Me) '確保表單改變時控制項隨之改變
End Sub

Private Function GetAttachFile(ByRef outType As String) As Boolean

On Error GoTo ErrHnd
   
    GetAttachFile = False
    'Added by Lydia 2024/10/15 查名單-網中：圖形查詢附件
    If m_STMF03 <> "" Then
       '開啟時,無法刪除,預設下次開啟表單執行刪檔
       If Dir(m_AttachPath & "\H*.jpg") <> "" Then
          Kill m_AttachPath & "\H*.jpg"
       End If
       If Dir(m_AttachPath & "\HM*.JPG") <> "" Then
          Kill m_AttachPath & "\HM*.JPG"
       End If
  
       outType = UCase(Trim(mTQF05))
       stTempFile = m_AttachPath & "\" & mTQF01 & mTQFkind & m_STMF03 & "." & outType
       If PUB_TMQAppFileGet(m_AttachPath, stTempFile, mTQF01, mTQFkind, m_STMF03) = False Then
          MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
          Exit Function
       Else
          If outType <> "PDF" Then
             Set pic1.Picture = pvGetStdPicture(Trim(stTempFile))
             tmpImg.Stretch = False
             Set tmpImg.Picture = pic1.Picture
          Else
             WebBrowser1.Navigate stTempFile
          End If
       End If
    Else
    'end 2024/10/15
       '開啟時,無法刪除,預設下次開啟表單執行刪檔
       If Dir(m_AttachPath & "\HM*.jpg") <> "" Then
          Kill m_AttachPath & "\HM*.jpg"
       End If
       If Dir(m_AttachPath & "\HM*.pdf") <> "" Then
          Kill m_AttachPath & "\HM*.pdf"
       End If
       If Dir(m_AttachPath & "\HM*.JPG") <> "" Then
          Kill m_AttachPath & "\HM*.JPG"
       End If
       If Dir(m_AttachPath & "\HM*.PDF") <> "" Then
          Kill m_AttachPath & "\HM*.PDF"
       End If
   
       outType = UCase(Trim(mTQF05))
       stTempFile = m_AttachPath & "\" & mTQF01 & "_" & mTQFkind & "." & outType
       If PUB_TMQGetAFile("", stTempFile, mTQF01, TMQ_附件F02, mTQFkind, TMQ_附件F04, outType) = False Then
          MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
          Exit Function
       Else
          If outType <> "PDF" Then
             Set pic1.Picture = pvGetStdPicture(Trim(stTempFile))
             tmpImg.Stretch = False
             Set tmpImg.Picture = pic1.Picture
          Else
             WebBrowser1.Navigate stTempFile
          End If
       End If
    End If 'Added by Lydia 2024/10/15
    GetAttachFile = True
    Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
   If file_num > 0 Then Close #file_num
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090129 = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
      'Added by Lydia 2016/05/10 判斷表單是否開啟
      If PUB_CheckFormExist(TypeName(m_PrevForm)) Then
         m_PrevForm.Show
      End If
   End If
   Set m_PrevForm = Nothing
End Sub
