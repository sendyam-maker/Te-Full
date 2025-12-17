Attribute VB_Name = "acc_cls"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  清除表單內容
'
'*************************************************

Public Sub Frmacc2150_Clear()
   With Frmacc2150
      .Text1 = ""      '2009/3/30 ADD BY SONIA 輸完存檔後再按新增不會清除
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = Mid(CFDate(ACDate(ServerDate)), 1, 3) & "/__/__"
      .MaskEdBox1.Mask = DFormat
      .Text6 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text8 = ""
      .AdodcRefresh
      .AdodcClear
      .Combo1 = ""    '2005/5/2 ADD BY SONIA
      '2010/12/6 ADD BY SONIA
      'Modified by Morgan 2013/10/23 考慮程序新人
      'If strUserNum = "73017" Or strUserNum = "81002" Then
      'Modified by Morgan 2016/10/19 +73 (品薇也要)
      'If PUB_GetST05(strUserNum) = "75" Then
      'Modofied by Morgan 2019/11/14 此處取消,改整批輸入自動新增時以案號設定 RMB/USD
      'If Pub_strUserST05 = "73" Or Pub_strUserST05 = "75" Then
      ''end 2013/10/23
      '   .Combo1 = "USD"
      'End If
      'end 2019/11/14
      
      .Text1.SetFocus
      .Label8 = ""    '2010/4/2 ADD BY SONIA
      .Check1.Value = vbUnchecked 'Added by Morgan 2019/3/12
      .Check2.Value = vbUnchecked 'Added by Morgan 2019/3/15
      'Added by Morgan 2023/4/20
      .Check3.Value = vbUnchecked
      .m_Comp = ""
      'end 2023/4/20
   End With
End Sub

Public Sub Frmacc2160_Clear()
   With Frmacc2160
      .Text2 = ""
      '.Text4 = ""
      .Text1 = ""
      .Text3 = ""
      .Text5 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = Mid(CFDate(ACDate(ServerDate)), 1, 3) & "/__/__"
      .MaskEdBox1.Mask = DFormat
      .Text7 = ""
      .Text9 = ""
      .AdodcRefresh
      .AdodcClear
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc21g0_Clear(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text9 = ""
      .Text10 = ""
      .Text11 = ""
      .Text13 = "" 'Added by Morgan 2018/3/7
      .Text14 = "" 'Added by Morgan 2018/3/7
      .Text15 = "" 'Added by Morgan 2018/3/7
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc21h0_Clear()
   With Frmacc21h0
      .Text1 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text5 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .AdodcRefresh
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc21i0_Clear()
   With Frmacc21i0
      .Text1 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text7 = ""
      .Text12 = ""
      .Text13 = ""
      .Text14 = ""
      .Text8 = ""
      .Text10 = ""
      .Text11 = ""
      .Text9 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc21j0_Clear()
   With Frmacc21j0
      .Text2 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text1 = ""
      .Text3 = ""
      .Text4 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text5 = ""
      .Text6 = ""
      .Text8 = ""
      .AdodcRefresh
      .Text2.SetFocus
   End With
End Sub

Public Sub Frmacc21k0_Clear()
   With Frmacc21k0
      .Text5 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = CFDate(strSrvDate(2))
      .MaskEdBox2.Mask = DFormat
      .Text1 = ""
      .Text7 = ""
      .Text8 = ""
      .Text9 = ""
      .Text6 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text10 = ""
      .AdodcRefresh
      .Text5.SetFocus
   End With
End Sub

Public Sub Frmacc21m0_Clear(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text5 = ""
      .MaskEdBox1.SetFocus
   End With
End Sub

Public Sub Frmacc21s0_Clear(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      .Combo1.Text = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .textDNR03.Text = ""
      .textDNR04.Text = ""
      .Combo1.SetFocus
   End With
End Sub
'Add By Sindy 2009/06/06
Public Sub Frmacc21o0_Clear(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      .Combo1 = ""
      .MaskEdBox1.Mask = ""
      'modify by sonia 2016/2/25 禧佩要求預設系統日
      '.MaskEdBox1.Text = ""
      .MaskEdBox1.Text = CFDate(strSrvDate(2))
      'end 2016/2/25
      .MaskEdBox1.Mask = DFormat
      .Text5 = ""
      .Combo1.SetFocus
      .txtBase = "1.03"
   End With
End Sub

