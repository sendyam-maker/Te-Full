VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060114 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子送件稽核"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4725
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   2745
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3585
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3540
      TabIndex        =   2
      Top             =   930
      Width           =   800
   End
   Begin VB.TextBox txtPA11 
      Height          =   270
      Left            =   1170
      MaxLength       =   25
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   630
      Width           =   2200
   End
   Begin VB.TextBox txtCP64 
      Height          =   270
      Left            =   1170
      MaxLength       =   12
      TabIndex        =   1
      Top             =   990
      Width           =   2200
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Left            =   1170
      TabIndex        =   14
      Top             =   2640
      Width           =   3195
      VariousPropertyBits=   27
      Size            =   "5636;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lb_Data 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   2
      Left            =   1170
      TabIndex        =   13
      Top             =   2310
      Width           =   3195
   End
   Begin VB.Label Lb_Data 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   1170
      TabIndex        =   12
      Top             =   1950
      Width           =   3195
   End
   Begin VB.Label Lb_Data 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   1170
      TabIndex        =   11
      Top             =   1590
      Width           =   3195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   210
      TabIndex        =   10
      Top             =   2640
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文文號:"
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   990
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   1590
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   210
      TabIndex        =   7
      Top             =   1950
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   2310
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   5
      Top             =   630
      Width           =   765
   End
End
Attribute VB_Name = "frm060114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/11 改成Form2.0 ; Lb_Data(3)=>lblFM2
'Create by Amy 2013/05/15
Option Explicit

Dim i  As Integer
Dim m_CP09 As String
Dim strSQL2 As String 'Added by Lydia 2018/03/22

Private Sub Form_Activate()
    txtPA11.SetFocus
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    cmdOK.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    
    '更新CP64
    Dim strUpd As String
    'Modified by Morgan 2018/2/23 --敏莉
    'strUpd = "Update CaseProgress set CP64=CP64||'電子送件已稽核;' Where CP09='" & m_CP09 & "'"
    strUpd = "Update CaseProgress set CP64='電子送件已稽核;'||CP64 Where CP09='" & m_CP09 & "'"
    cnnConnection.Execute strUpd
    MsgBox ("電子送件稽核完成!")
    cmdOK.Enabled = False
    CmdSearch.Default = True 'Add By Sindy 2014/5/9
    'Added by Lydia 2018/03/22  檢查同一案件進度已有智慧局收文文號但尚未稽核之案件性質
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSQL2)
    If intI = 1 Then
        strExc(2) = ""
        RsTemp.MoveFirst
        Do While Not RsTemp.EOF
              strExc(2) = strExc(2) & RsTemp.Fields("案件性質") & "、"
              RsTemp.MoveNext
        Loop
        MsgBox "本案尚有" & Mid(strExc(2), 1, Len(strExc(2)) - 1) & "尚未稽核！", vbInformation, "檢查"
    End If
    'end 2018/03/22
    
    Exit Sub
    
ErrHand:
    MsgBox "更新失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdSearch_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim idx As Integer
   Dim Lbl As LABEL
      
   TxtClear
   
   '欄位驗證
   'Modified by Lydia 2018/12/18 FCP-60108衍生設計案的申請編號超過9碼
   'If Len(txtPA11) <> 9 Then MsgBox "請輸入申請案號", vbInformation: txtPA11.SetFocus: Exit Sub
   If Len(txtPA11) < 9 Then MsgBox "請輸入申請案號", vbInformation: txtPA11.SetFocus: Exit Sub
   If txtCP64 = "" Then MsgBox "請輸入收文文號", vbInformation: txtCP64.SetFocus: Exit Sub
   
   'Modify By Sindy 2013/8/26
   'Modify by Amy 2017/09/21 原：CP118='Y'未判斷到A
   If Left(Trim(PUB_GetST03(strUserNum)), 2) = "F1" Then '外商
       'Modified by Lydia 2018/01/11 +CP43
      strSql = "Select Decode(TM28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(TM57,'')),null,'','●') AS 本所案號, " & _
               "NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質,CP27,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CP09,CP43 From Trademark,CaseProgress,CasePropertyMap,Customer " & _
               "Where TM12='" & txtPA11 & "' And TM10='000' And TM01='FCT' And CP01(+)=TM01 And CP02(+)=TM02 " & _
               "And CP03(+)=TM03 And CP04(+)=TM04 And SubStr(TM23,1,8)=CU01(+) And SubStr(TM23,9,1)=CU02(+) " & _
               "And CP01=CPM01(+) And CP10=CPM02(+) And CP118 is not null And CP27 >0 And  INSTR(CP64,'智慧局收文文號:" & txtCP64 & ";')>0 And Not INSTR(CP64,'電子送件已稽核;')>0"
   '2013/8/26 END
   Else
      'Modify By Sindy 2013/8/26 CU05 ==> NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90))
      'Modified by Lydia 2018/01/11 +CP43
      strSql = "Select Decode(PA23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號, " & _
               "NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,CP27,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CP09,CP43 From Patent,CaseProgress,CasePropertyMap,Customer " & _
               "Where PA11='" & txtPA11 & "' And PA09='000' And PA01='FCP' And CaseProgress.CP01(+)=Patent.PA01 And CaseProgress.CP02(+)=Patent.PA02 " & _
               "And CaseProgress.CP03(+)=Patent.PA03 And CaseProgress.CP04(+)=Patent.PA04 And SubStr(PA26,1,8)=CU01(+) And SubStr(PA26,9,1)=CU02(+) " & _
               "And CP01=CPM01(+) And CP10=CPM02(+) And CP118 is not null And CP27 >0 And  INSTR(CP64,'智慧局收文文號:" & txtCP64 & ";')>0 And Not INSTR(CP64,'電子送件已稽核;')>0"
   End If
   'Added by Lydia 2018/03/22  檢查同一案件進度已有智慧局收文文號但尚未稽核之案件性質
   strSQL2 = Replace(strSql, "INSTR(CP64,'智慧局收文文號:" & txtCP64 & ";')>0", "INSTR(CP64,'智慧局收文文號:')>0")
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        cmdOK.Enabled = True
        cmdOK.Default = True 'Add By Sindy 2014/5/9
        m_CP09 = rsTmp("CP09")
        lblFM2 = "" & rsTmp(3)  'Added by Lydia 2021/09/11 Lb_Data(3); 改成Form 2.0
        For Each Lbl In Lb_Data
            idx = Lbl.Index
            If IsNull(rsTmp(idx)) Then
                Lbl.Caption = ""
            Else
                If idx = 2 Then
                    Lbl.Caption = ChangeWStringToTDateString(rsTmp(idx)) '發文日
                'Added by Lydia 2018/01/11 案件性質+相關案號
                ElseIf idx = 1 Then
                    If "" & rsTmp.Fields("CP43") <> "" Then
                        Lbl.Caption = rsTmp(idx) & "-" & PUB_GetRelateCasePropertyName("" & rsTmp.Fields("CP09"), "1")
                    Else
                         Lbl.Caption = rsTmp(idx)
                    End If
                'end 2018/01/11
                Else
                    Lbl.Caption = rsTmp(idx)
                End If
            End If
        Next
    Else
        cmdOK.Enabled = False
        MsgBox ("查無資料或電子送件已稽核")
    End If
End Sub

Private Sub TxtClear()
   Dim Lbl As LABEL
   For Each Lbl In Lb_Data
       Lbl = ""
   Next
End Sub
