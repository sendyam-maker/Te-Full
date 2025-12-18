VERSION 5.00
Begin VB.Form frm880015 
   BorderStyle     =   1  '單線固定
   Caption         =   "同案件同時段多筆發文設定"
   ClientHeight    =   4110
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4680
   Begin VB.TextBox txtPeriod 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   540
      Width           =   1410
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2355
      TabIndex        =   1
      Top             =   45
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      Top             =   45
      Width           =   1200
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      ItemData        =   "frm880015.frx":0000
      Left            =   120
      List            =   "frm880015.frx":0007
      Sorted          =   -1  'True
      Style           =   1  '項目包含核取方塊
      TabIndex        =   0
      Top             =   1380
      Width           =   4410
   End
   Begin VB.Label Label3 
      Caption         =   "PS：１．獨立申請書則請勾選算發文室件數　　　　　　　　２．無規費且非獨立申請書不可算發文室件數"
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   180
      TabIndex        =   8
      Top             =   3600
      Width           =   4500
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "是否算發文室件數　案件性質"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   4350
   End
   Begin VB.Label lblFund 
      Caption         =   "送件時段："
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "frm880015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/16 Form2.0已檢查 (無需修改的物件);
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

Public strCP09 As String '收文號
Public strCP09s As String '收文號(回傳)
Public strCP123s As String '是否經發文室-主管機關
Public strCP27 As String '發文日
Public bolBDefer As Boolean 'B類延期
Public bolOK As Boolean
Public strCP84 As String   'add by sonia 2014/6/23 發文規費
Public m_CP123s As String  'add by sonia 2014/6/23 是否經發文室-主管機關(暫存用)

Private Sub cmdOK_Click(Index As Integer)
   '確定
   If Index = 0 Then
      strCP123s = ""
      If lstData.Selected(0) = True Then
         strCP123s = strCP123s & "Y"
      Else
         strCP123s = strCP123s & "N"
      End If
      For intI = 1 To lstData.ListCount - 1
         strCP123s = strCP123s & ","
         If lstData.Selected(intI) = True Then
            strCP123s = strCP123s & "Y"
         Else
            strCP123s = strCP123s & "N"
         End If
      Next
      bolOK = True
   '回前畫面(取消)
   Else
      bolOK = False
   End If
   Me.Hide
End Sub

Public Function CheckShowList() As Boolean
   Dim stSQL As String, stCon As String, stSQL1 As String
   Dim stDesc As String, intIdx As Integer
   Dim bOther As Boolean
   
   lstData.Clear
   '未指定發文日時預設當日
   If strCP27 = "" Then
      strCP27 = strSrvDate(1)
   Else
      strCP27 = DBDATE(strCP27)
   End If
   
   txtPeriod = Format(strCP27 - 19110000, "###/##/##")
   
   strCP09s = strCP09
   strCP123s = ""
   If strCP09 <> "" Then
      If bolBDefer = True Then
         stSQL = "select a.CP09,'延期-'||b.CPM03 cpm03,a.CP01,a.CP02,a.CP03,a.CP04,a.cp43,a.cp10,sk02,'Y' flg" & _
            " from caseprogress a,casepropertymap b,casepropertymap c,systemkind" & _
            " where a.cp09='" & strCP09 & "' and b.cpm01(+)=a.cp01 and b.cpm02(+)=a.cp10 and sk01(+)=cp01"
      Else
         'Modified by Morgan 2013/1/10 取消抽換控制,補文件也可能備註抽換,但要經發文室 Ex.FCP-033029 --張靜芳,譚文容
         '抽換不經發文室,由總務直接給智慧局的承辦人員
         'stSQL = "select a.CP09,b.CPM03,a.CP01,a.CP02,a.CP03,a.CP04,a.cp43,a.cp10,sk02,b.cpm21 flg" & _
            " from caseprogress a,casepropertymap b,systemkind" & _
            " where a.cp09='" & strCP09 & "' and b.cpm01(+)=a.cp01 and b.cpm02(+)=a.cp10 and sk01(+)=cp01" & _
            " and b.cpm21 is not null and (a.cp64 is null or instr(a.CP64,'抽換')=0)"
         stSQL = "select a.CP09,b.CPM03,a.CP01,a.CP02,a.CP03,a.CP04,a.cp43,a.cp10,sk02,b.cpm21 flg" & _
            " from caseprogress a,casepropertymap b,systemkind" & _
            " where a.cp09='" & strCP09 & "' and b.cpm01(+)=a.cp01 and b.cpm02(+)=a.cp10 and sk01(+)=cp01" & _
            " and b.cpm21 is not null"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         '若為B類延期則收文號存檔時補
         If bolBDefer = True Then strCP09s = ""
         
         bOther = False
         '若案件性質為其他時要確認是否要經發文室
         '專利
         If RsTemp("flg") = "Y" Then
            If RsTemp("sk02") = "1" Or RsTemp("sk02") = "5" Then
               If RsTemp("cp10") = "910" Then
                  bOther = True
               End If
            '商標
            Else
               If RsTemp("cp10") = "706" Then
                  bOther = True
               End If
            End If
            If bOther = True Then
               If MsgBox("本程序為【其他】，請確認是否要經發文室？", vbYesNo + vbDefaultButton1) = vbNo Then
                  bolOK = True
                  Exit Function
               End If
            End If
         End If
         
         If Not IsNull(RsTemp("cp43")) Then
            stDesc = PUB_GetRelateCasePropertyName(strCP09, "1")
         Else
            stDesc = ""
         End If
         lstData.AddItem RsTemp("cp09") & " 　　　" & RsTemp("cpm03") & stDesc, 0
         lstData.Selected(0) = True
         strCP123s = "Y"
         Me.txtCaseNo = RsTemp("cp01") & "-" & RsTemp("cp02") & IIf(RsTemp("cp03") & RsTemp("cp04") = "000", "", "-" & RsTemp("cp03") & "-" & RsTemp("cp04"))
            
         '讀取分段時間
         stCon = ""
         'Modify by Morgan 2011/10/7 +判斷系統時間晚於分段時間才算
         stSQL = "select al05 from staff,applist where st01='" & strUserNum & "'" & _
            " and al02=substr(st03,1,2) and al01=" & strCP27 & " and al05<to_char(sysdate,'hh24miss') and rownum<2"
         intI = 1
         Set adoRecordset = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            '若有分段時間則表示為下午送的案件
            stCon = " and a.cp82>" & adoRecordset.Fields(0)
            txtPeriod = txtPeriod & " 下午"
         Else
            txtPeriod = txtPeriod & " 上午"
         End If
         
         '抓同日發文其他有經發文室的程序
         'modify by sonia 2014/6/23 +CP84發文規費, P-108903
         stSQL = " select a.cp09,b.CPM03,a.cp123,a.cp43,a.cp84 cp84" & _
            " from caseprogress a,casepropertymap b" & _
            " where a.cp01='" & RsTemp("cp01") & "' and a.cp02='" & RsTemp("cp02") & "'" & _
            " and a.cp03='" & RsTemp("cp03") & "' and a.cp04='" & RsTemp("cp04") & "'" & _
            " and a.cp27=" & strCP27 & " and a.cp09<>'" & strCP09 & "'" & stCon & _
            " and b.cpm01(+)=a.cp01 and b.cpm02(+)=a.cp10 and a.cp123 is not null" & _
            " order by a.cp82 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               If Not IsNull(RsTemp("cp43")) Then
                  stDesc = PUB_GetRelateCasePropertyName(RsTemp("cp09"), "1")
               Else
                  stDesc = ""
               End If
               intIdx = lstData.ListCount
               lstData.AddItem RsTemp("cp09") & " 　　　" & RsTemp("cpm03") & stDesc, intIdx
               strCP09s = strCP09s & "," & RsTemp("cp09")
               'modify by sonia 2014/6/23 +CP84發文規費, P-108903變更 發文之收文號有規費但其他已發文者為無發文規費卻算發文室件數,則改該筆不算案件數
               'strCP123s = strCP123s & "," & RsTemp("cp123")
               'If RsTemp("cp123") = "Y" Then
               m_CP123s = RsTemp("cp123")
               If Val(strCP84) > 0 And m_CP123s = "Y" And Val("" & RsTemp("cp84")) = 0 Then
                  If MsgBox(RsTemp("cp09") & " " & RsTemp("cpm03") & stDesc & " 無規費卻算發文室件數, 是否為獨立申請書？" & Chr(10) & Chr(13) & Chr(13) & _
                         "(若非獨立申請書, 則此收文號發文存檔後會將該筆改為不算發文室案件數！)", vbExclamation + vbYesNo) = vbYes Then
                     m_CP123s = "Y"
                  Else
                     m_CP123s = "N"
                  End If
               End If
               strCP123s = strCP123s & "," & m_CP123s
               If m_CP123s = "Y" Then
               'end 2014/6/23
                  lstData.Selected(intIdx) = True
               End If
               RsTemp.MoveNext
            Loop
            lstData.ListIndex = 0
            CheckShowList = True
         Else
            bolOK = True
         End If
      Else
         bolOK = True
      End If
   End If
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

