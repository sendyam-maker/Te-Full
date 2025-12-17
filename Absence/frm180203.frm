VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm180203 
   AutoRedraw      =   -1  'True
   Caption         =   "資料確認(預覽)"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8925
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   6630
      TabIndex        =   4
      Top             =   0
      Width           =   1965
      Begin VB.CommandButton cmdok 
         Caption         =   "結束(&X)"
         Height          =   405
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "確認(&O)"
         Height          =   405
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   228
      Left            =   630
      Max             =   20
      TabIndex        =   2
      Top             =   6150
      Visible         =   0   'False
      Width           =   8670
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5685
      Left            =   9600
      TabIndex        =   1
      Top             =   6390
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   5460
      Left            =   930
      ScaleHeight     =   5400
      ScaleWidth      =   8580
      TabIndex        =   0
      Top             =   6390
      Visible         =   0   'False
      Width           =   8640
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5805
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8925
      ExtentX         =   15743
      ExtentY         =   10239
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frm180203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/24 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Sindy 2010/9/2 日期欄已修改
Option Explicit

Public m_ImageH As Long, m_ImageW As Long, m_iPages As Integer

Dim m_iPageNow As Integer, m_iPageLast As Integer
Dim m_PicX As Long, m_PicY As Long
Dim iDotH As Integer, iDotV As Integer
Dim m_Pictures(1 To 2) As StdPicture
Public m_B1301 As String, m_B1302 As String, m_B1303 As String
Dim douExtRate As Double '字型位置縮放比


Private Sub cmdok_Click(Index As Integer)
Dim strB1303 As String, intB1303Seqno As Integer
   
On Error GoTo ErrHand
   
   Select Case Index
      Case 0 '確認
         Screen.MousePointer = vbHourglass
         cnnConnection.BeginTrans
         
         '個人資料明細只需員工確認即可
         If m_B1301 <> "05" Then
            strB1303 = m_B1303 '目的：為讓目前處理人員,可以Run到Do While迴圈
            '檢查下一處理人員是否離職,若是,則移轉下一處理人員
            Do While strB1303 = m_B1303 Or (strB1303 <> "" And ChkStaffST04(strB1303, False) = True)
               '審核主管確認
               intB1303Seqno = GetCurrB1303Seqno(m_B1301, m_B1302, strB1303)
               If intB1303Seqno = 1 Then
                  strSql = "Update ABS013 set B1304='Y' WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
                  cnnConnection.Execute strSql
               ElseIf intB1303Seqno = 2 Then
                  strSql = "Update ABS013 set B1305='Y' WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
                  cnnConnection.Execute strSql
               ElseIf intB1303Seqno = 3 Then
                  strSql = "Update ABS013 set B1306='Y' WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
                  cnnConnection.Execute strSql
               ElseIf intB1303Seqno = 4 Then
                  strSql = "Update ABS013 set B1307='Y' WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
                  cnnConnection.Execute strSql
               End If
               '讀取下一處理人員
               strB1303 = GetNextB1303(m_B1301, m_B1302)
            Loop
         End If
         If strB1303 = "" Then
            '已無下一處理人員時,代表已確認完畢即可刪除該筆資料
            strSql = "DELETE FROM ABS013 WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
         Else
            '送至下一處理人員
            strSql = "Update ABS013 set B1303='" & strB1303 & "' WHERE B1301='" & m_B1301 & "' and B1302='" & m_B1302 & "' "
         End If
         cnnConnection.Execute strSql
         
         cnnConnection.CommitTrans
         Screen.MousePointer = vbDefault
         
         Unload Me
         Exit Sub
         
      Case 1 '結束
         bolfrm180203ExitForm = True
'         If InStr(Me.Caption, "員工個人資料明細") > 0 Then
'            'frm160102.bolExitForm = True
'            'If pub_CallNextABSForm = False Then '不可少的判斷
''               Unload frm160102
'            'End If
'         ElseIf InStr(Me.Caption, "每月出缺勤統計") > 0 Then
'            'frm160201.bolExitForm = True
'            'If pub_CallNextABSForm = False Then '不可少的判斷
''               Unload frm160201
'            'End If
'         End If
         Unload Me
   End Select
   
   Exit Sub
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "更新失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   iDotH = 40: iDotV = 30
   
   With frm180203
      .Top = 0
      .Left = 0
      .Width = 9048 '11800 '12000 'Screen.Width - 100
      .Height = 6195 '5700 '6120 '7800 '6800 '8500 'Screen.Height - 1500
      .Picture1.Width = .ScaleWidth - .VScroll1.Width
      .Picture1.Height = .ScaleHeight - .HScroll1.Height
      .VScroll1.Left = .Picture1.Width
      .VScroll1.Height = .Picture1.Height
      .HScroll1.Top = .Picture1.Height
      .HScroll1.Width = .Picture1.Width
      .Frame1.Left = 6600 '9000
      .Frame1.BorderStyle = 0
      .Frame1.BackColor = &H80000005
   End With
   
   MoveFormToCenter Me
   
   '垂直捲軸
   If Picture1.Height >= m_ImageH Then
      VScroll1.Visible = False
   Else
      VScroll1.max = iDotV * m_iPages - Int(iDotV * Picture1.Height / m_ImageH)
   End If
   
   '水平捲軸
   If Picture1.Width >= m_ImageW Then
      HScroll1.Visible = False
   Else
      HScroll1.max = iDotH - Int(iDotH * Picture1.Width / m_ImageW)
   End If
      
   '載入第一張報表
   m_iPageLast = 0
   m_iPageNow = 1
   m_PicX = 0: m_PicY = 0
   PaintPic
   
   bolfrm180203ExitForm = False
   Forms(0).m_ChkIsOpenFrm180203 = True 'Add By Sindy 2013/7/8
End Sub

Private Function GetPic(idx As Integer, idx1 As Integer) As Boolean
Dim strPicFileName As String
   
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   If Dir(strPicFileName) = "" Then GetPic = False: Exit Function
   GetPic = True
   Set m_Pictures(idx1) = LoadPicture(strPicFileName)
   Picture1.AutoSize = True
   douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
End Function

Private Sub PaintPic()
   Picture1.Line (0, 0)-(Picture1.Width, Picture1.Height), QBColor(15), BF
   If m_iPageLast <> m_iPageNow Then
      If GetPic(m_iPageNow, 1) = False Then Exit Sub
      If m_iPages > m_iPageNow Then
         GetPic m_iPageNow + 1, 2
      End If
   End If
   
   Picture1.PaintPicture m_Pictures(1), m_PicX, m_PicY
   If m_iPages > m_iPageNow Then
      Picture1.PaintPicture m_Pictures(2), m_PicX, m_PicY + m_ImageH
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Forms(0).m_ChkIsOpenFrm180203 = False 'Add By Sindy 2013/7/8
   DelPic
   Set frm180203 = Nothing
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
End Sub

Private Sub HScroll1_Change()
   m_PicX = -1 * m_ImageW * HScroll1.Value / iDotH
   PaintPic
End Sub

Private Sub HScroll1_Scroll()
   m_PicX = -1 * m_ImageW * HScroll1.Value / iDotH
   PaintPic
End Sub

Private Sub VScroll1_Change()
   m_iPageLast = m_iPageNow
   m_iPageNow = 1 + VScroll1.Value \ iDotV
   m_PicY = -1 * m_ImageH * (VScroll1.Value Mod iDotV) / iDotV
   PaintPic
End Sub

Private Sub VScroll1_Scroll()
   m_iPageLast = m_iPageNow
   m_iPageNow = 1 + VScroll1.Value \ iDotV
   m_PicY = -1 * m_ImageH * (VScroll1.Value Mod iDotV) / iDotV
   PaintPic
End Sub
