VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100106_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "列印聯絡單"
   ClientHeight    =   2484
   ClientLeft      =   360
   ClientTop       =   3876
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2484
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3948
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3156
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   48
      Width           =   756
   End
   Begin VB.OptionButton Option1 
      Caption         =   "承辦人"
      Height          =   180
      Index           =   2
      Left            =   156
      TabIndex        =   2
      Top             =   1170
      Width           =   852
   End
   Begin VB.OptionButton Option1 
      Caption         =   "智權人員"
      Height          =   180
      Index           =   1
      Left            =   156
      TabIndex        =   1
      Top             =   810
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "管制人"
      Height          =   180
      Index           =   0
      Left            =   156
      TabIndex        =   0
      Top             =   450
      Value           =   -1  'True
      Width           =   972
   End
   Begin MSForms.TextBox txt1 
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1830
      Visible         =   0   'False
      Width           =   4575
      VariousPropertyBits=   -1476378597
      ScrollBars      =   2
      Size            =   "8064;1080"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1710
      Width           =   4575
      VariousPropertyBits=   -1476378597
      ScrollBars      =   2
      Size            =   "8064;1080"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Top             =   1170
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   810
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Top             =   450
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   1500
      Width           =   540
   End
End
Attribute VB_Name = "frm100106_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: lbl1(index)、txt1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/13 日期欄已修改
Option Explicit
Dim i As Integer, j As Integer, strSql As String
Dim Str01 As String, Str02 As String, Str03 As String, strTemp As String
Dim StrR04001 As String
Dim StrR04002 As String
Dim StrR04003 As String
Dim StrR04004 As String
Dim StrR04005 As String
Dim StrR04006 As String
Dim StrR04007 As String
Dim StrR04008 As String
Dim StrR04009 As String
Dim StrR04010 As String
Dim StrR04011 As String
Dim StrR04012 As String
Dim StrR04013 As String
Dim StrR04014 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
    If StrMenu1 = False Then
      Exit Sub
    End If
    If DataEnvironment1.rsCommand2.State = 1 Then
      DataEnvironment1.rsCommand2.Close
    End If
    'Add by Morgan 2007/7/27 重新設語法
    DataEnvironment1.Commands(2).CommandText = "SELECT * FROM R100106 where id='" & strUserNum & "'"
    'end 2007/7/27
    DataEnvironment1.Command2
    
    PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
    datrptPublic1.PrintReport
    PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
    
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case Else
End Select
End Sub


 
Private Sub cmdGoInput_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
    If StrMenu1 = False Then
      Exit Sub
    End If
    If DataEnvironment1.rsCommand2.State = 1 Then
      DataEnvironment1.rsCommand2.Close
    End If
    DataEnvironment1.Command2
    'datrptPublic1.Show
    PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
    datrptPublic1.PrintReport
    PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
    Me.Hide
Case 1
     Me.Hide
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
For i = 0 To 2
    lbl1(i).Caption = ""
Next i
txt1(0).Text = "本案期限將至，請儘速作業，已利後續作業。"
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
Str01 = SystemNumber(Me.Tag, 1)
Str02 = SystemNumber(Me.Tag, 2)
Str03 = SystemNumber(Me.Tag, 3)
lbl1(0).Caption = Str01
lbl1(1).Caption = Str02
lbl1(2).Caption = Str03
End Sub
Function StrMenu1() As Boolean
CheckOC

'Modified by Lydia 2023/12/22 暫存檔欄位不足
'strSql = "SELECT * FROM R100106_T where id='" & strUserNum & "'"
strSql = "SELECT * FROM R100106_T,caseprogress where id='" & strUserNum & "' and r03003=cp09(+) "

adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Option1(0).Value = True Then
        'Modified by Lydia 2023/12/22 管制人
        'If Not IsNull(adoRecordset.Fields(1)) Then
        '    StrR04002 = adoRecordset.Fields(1)
        If Not IsNull(adoRecordset.Fields("r03002")) Then
            StrR04002 = adoRecordset.Fields("r03002")
        'end 2023/12/22
        Else
            StrR04002 = ""
        End If
    Else
        If Option1(1).Value = True Then
            'Modified by Lydia 2023/12/22 智權人員
            'If Not IsNull(adoRecordset.Fields(0)) Then
            '    StrR04002 = adoRecordset.Fields(0)
            If Not IsNull(adoRecordset.Fields("r03001")) Then
                StrR04002 = adoRecordset.Fields("r03001")
            'end 2023/12/22
            Else
                StrR04002 = ""
            End If
        Else
            If Option1(2).Value = True Then
                'Modified by Lydia 2023/12/22 承辦人員
                'If Not IsNull(adoRecordset.Fields(2)) Then
                '    StrR04002 = adoRecordset.Fields(2)
                If Not IsNull(adoRecordset.Fields("CP14")) Then
                    StrR04002 = adoRecordset.Fields("CP14")
                'end 2023/12/22
                Else
                    StrR04002 = ""
                End If
            Else
                StrR04002 = ""
            End If
        End If
    End If
    CheckOC2
    'Added by Lydia 2023/12/22
    If strSrvDate(1) >= 新部門啟用日 Then
       strSql = "SELECT ST02,NVL(A0921,ST03) ST03 FROM STAFF,ACC090NEW WHERE ST01='" & StrR04002 & "' AND ST93=A0921(+) "
    Else
    'end 2023/12/22
       strSql = "SELECT ST02,ST03 FROM STAFF WHERE ST01='" & StrR04002 & "'"
    End If
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If Not IsNull(adoRecordset1.Fields(0)) Then
            StrR04002 = adoRecordset1.Fields(0)
        Else
            StrR04002 = ""
        End If
        If Not IsNull(adoRecordset1.Fields(1)) Then
            StrR04001 = adoRecordset1.Fields(1)
        Else
            StrR04001 = ""
        End If
    Else
       StrR04001 = ""
    End If
    CheckOC2
    'Added by Lydia 2023/12/22
    If strSrvDate(1) >= 新部門啟用日 Then
        strSql = "SELECT A0923 FROM ACC090NEW WHERE A0921='" & StrR04001 & "'" & _
                 " UNION SELECT A0902 FROM ACC090 WHERE A0901='" & StrR04001 & "' AND A0901 NOT IN (SELECT A0921 FROM ACC090NEW)"
    Else
    'end 2023/12/22
        strSql = "SELECT A0902 FROM ACC090 WHERE A0901='" & StrR04001 & "'"
    End If
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If Not IsNull(adoRecordset1.Fields(0)) Then
            StrR04001 = adoRecordset1.Fields(0)
        Else
            StrR04001 = ""
        End If
    Else
        StrR04001 = ""
    End If
    CheckOC2
    If Not IsNull(adoRecordset.Fields(3)) Then
        StrR04003 = adoRecordset.Fields(3)
    Else
        StrR04003 = ""
    End If
    If Not IsNull(adoRecordset.Fields(4)) Then
        StrR04004 = adoRecordset.Fields(4)
    Else
        StrR04004 = ""
    End If
    If Not IsNull(adoRecordset.Fields(5)) Then
        StrR04005 = adoRecordset.Fields(5)
    Else
        StrR04005 = ""
    End If
    If Not IsNull(adoRecordset.Fields(6)) Then
        StrR04006 = adoRecordset.Fields(6)
    Else
        StrR04006 = ""
    End If
    If Not IsNull(adoRecordset.Fields(7)) Then
        StrR04007 = adoRecordset.Fields(7)
    Else
        StrR04007 = ""
    End If
    If Not IsNull(adoRecordset.Fields(8)) Then
        StrR04008 = adoRecordset.Fields(8)
    Else
        StrR04008 = ""
    End If
    StrR04009 = ""
    StrR04010 = ""
    StrR04011 = ""
    If Not IsNull(adoRecordset.Fields(9)) Then
        StrR04012 = adoRecordset.Fields(9)
    Else
        StrR04012 = ""
    End If
    If Not IsNull(adoRecordset.Fields(10)) Then
        StrR04013 = adoRecordset.Fields(10)
    Else
        StrR04013 = ""
    End If
Else
    StrR04001 = ""
    StrR04002 = ""
    StrR04003 = ""
    StrR04004 = ""
    StrR04005 = ""
    StrR04006 = ""
    StrR04007 = ""
    StrR04008 = ""
    StrR04009 = ""
    StrR04010 = ""
    StrR04011 = ""
    StrR04012 = ""
    StrR04013 = ""
    ShowNoData
    StrMenu1 = False
    Screen.MousePointer = vbDefault
    Exit Function
End If
CheckOC
strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & strUserNum & "'"
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Not IsNull(adoRecordset.Fields(0)) Then
        strTemp = adoRecordset.Fields(0)
    Else
        strTemp = strUserNum
    End If
Else
    strTemp = strUserNum
End If
CheckOC
StrR04014 = strTemp
cnnConnection.Execute "DELETE FROM R100106 WHERE id='" & strUserNum & "' "
cnnConnection.Execute "INSERT INTO R100106 VALUES ('" & ChgSQL(StrR04001) & "','" & ChgSQL(StrR04002) & "','" & ChgSQL(StrR04003) & "','" & ChgSQL(StrR04004) & "','" & ChgSQL(StrR04005) & "','" & ChgSQL(StrR04006) & "','" & ChgSQL(StrR04007) & "','" & ChgSQL(StrR04008) & "','" & ChgSQL(StrR04009) & "','" & ChgSQL(StrR04010) & "','" & ChgSQL(StrR04011) & "','" & ChgSQL(StrR04012) & "','" & ChgSQL(StrR04013) & "','" & ChgSQL(StrR04014) & "','" & strUserNum & "')"
StrMenu1 = True

End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm100106_5 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     If Option1(0).Value = True Then
        Option1(1).Value = False
        Option1(2).Value = False
     End If
Case 1
     If Option1(1).Value = True Then
        Option1(0).Value = False
        Option1(2).Value = False
     End If
Case 2
     If Option1(2).Value = True Then
        Option1(0).Value = False
        Option1(1).Value = False
     End If
Case Else
End Select
End Sub
