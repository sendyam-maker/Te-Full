Attribute VB_Name = "acc_cls"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit
'*************************************************
'  清除表單內容
'
'*************************************************
Public Sub Frmacc7100_Clear()
    With Frmacc7100
        'edit by nick 2004/08/19 新增時不預設
        If Frmacc7100.oState = "1" Then
                .MaskEdBox1.Text = "___/__/__"
                .MaskEdBox1.Mask = DFormat
        Else
            If .MaskEdBox1.Text = "" Or .MaskEdBox1.Text = "___/__/__" Then
                .MaskEdBox1.Mask = ""
                .MaskEdBox1.Text = CFDate(strSrvDate(2))
                .MaskEdBox1.Mask = DFormat
            End If
        End If
        .Text1.Text = "E"
        .Text2.Text = ""
        .Text13.Text = ""
        .Text14.Text = ""
        .Text15.Text = ""
        .Text16.Text = ""
        .Text3.Text = ""
        .Text4.Text = ""
        'add by nick 2004/08/19
        .Text10.Text = ""
        .Text11.Text = ""
        .Label18.Caption = "所別：" & " " & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
        .MaskEdBox2.Mask = ""
        .MaskEdBox2.Text = ""
        .MaskEdBox2.Mask = DFormat
        .Text5.Text = ""
        .Text6.Text = ""
        .Text7.Text = ""
        .MaskEdBox3.Mask = ""
        .MaskEdBox3.Text = ""
        .MaskEdBox3.Mask = DFormat
        .Text8.Text = ""
        .Text9.Text = ""
        If Frmacc7100.oState = "1" Or Frmacc7100.oState = "2" Then
            .MaskEdBox1.SetFocus
        End If
    End With
End Sub

'add by nick 2004/08/20
'*************************************************
'  鎖定表單內容
'
'*************************************************
Public Sub Frmacc7100_Lock()
    With Frmacc7100
        .MaskEdBox1.Enabled = False
        .MaskEdBox2.Enabled = False
        .Command1.Enabled = False
        .Text1.Enabled = False
        .Text2.Enabled = False
        .Text13.Enabled = False
        .Text14.Enabled = False
        .Text15.Enabled = False
        .Text16.Enabled = False
        .Text3.Enabled = False
        .Text4.Enabled = False
        .Text10.Enabled = False
        .Text11.Enabled = False
        .Text5.Enabled = False
        .Text6.Enabled = False
        .Text7.Enabled = False
        .MaskEdBox3.Enabled = False
        .Text8.Enabled = False
        .Text9.Enabled = False
    End With
End Sub

'add by nick 2004/08/20
'*************************************************
'  解鎖定表單內容
'
'*************************************************
Public Sub Frmacc7100_UnLock()
    With Frmacc7100
        .MaskEdBox1.Enabled = True
'        If Frmacc7100.oState = "2" Then
'            .Text1.Enabled = False
'            .Text2.Enabled = False
'        Else
            .Text1.Enabled = True
            .Text2.Enabled = True
'        End If
        .Text13.Enabled = True
        .Text14.Enabled = True
        .Text16.Enabled = True
        .Text3.Enabled = True
        .Text4.Enabled = True
        .Text11.Enabled = True
        .MaskEdBox2.Enabled = True
        .Text5.Enabled = True
        .Text6.Enabled = True
        .Text7.Enabled = True
        .MaskEdBox3.Enabled = True
        .Text8.Enabled = True
        .Text9.Enabled = True
    End With
End Sub

'Added by Lydia 2020/03/26 從account.aacc_cls複製
Public Sub Frmacc1130_Clear()
   With Frmacc1130
      .Text1 = "E"
      TextInverse .Text1
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox1.Mask = DFormat
      .Text9 = ""
      '.Text10 = "" 'Removed by Morgan 2011/11/24 一張收據可同時包含合併及不合併的收文資料
      .Text11 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Mask = DFormat
      .Text2 = ""
      .Text6 = ""
      .Text5 = ""
      .Text4 = ""
      .Text7 = ""
      .Text8 = ""
      .Text12 = ""
      .Text3 = ""
      .AdodcRefresh
      .Text1.SetFocus
   End With
End Sub

