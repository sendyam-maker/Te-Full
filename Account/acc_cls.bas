Attribute VB_Name = "aacc_cls"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'*************************************************
'  清除表單內容
'
'*************************************************
Public Sub Frmacc1110_Clear()
   With Frmacc1110
      .AutoNoQuery
      .Text2 = ""
      .Text3 = ""
      .Text4 = "E"
      .Text5 = "E"
      .Text6 = "E"
      .Text7 = "E"
      .Text8 = ""
      .Text9 = ""
      .Text10 = "E"
      .Text11 = "E"
      .Text2.SetFocus
   End With
End Sub

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

Public Sub Frmacc1150_Clear()
   With Frmacc1150
      .Text2 = ""
      .Text1 = ""
      .Text3 = ""
      .AdodcRefresh
      .AdodcClear
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc11d0_Clear()
   With Frmacc11d0
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text10 = ""
      .Text11 = ""
      .Text12 = ""
      'Add by Morgan 2004/1/12
      .Text9 = ""
      .Text13 = ""
      'Add end----------------
      .Text14 = "" 'Add By Sindy 2024/3/14
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc11f0_Clear()
   With Frmacc11f0
      .Text1 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox1.Mask = DFormat
      .Text2 = ""
      .Text3 = ""
      .Text5 = ""
      .Text6 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text4 = ""
      .Text1.SetFocus
   End With
End Sub
Public Sub Frmacc2110_Clear()
   With Frmacc2110
      .Text2 = ""
      If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
         .MaskEdBox1.Mask = ""
         .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
         .MaskEdBox1.Mask = DFormat
      End If
'      .Text1 = ""
'      .Text3 = ""
      .Text4 = ""
      .Combo2.Clear
'      .Text2.SetFocus
   End With
End Sub

Public Sub Frmacc2120_Clear()
   With Frmacc2120
      .Text2 = ""
      .Text1 = ""
      .Text3 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox1.Mask = DFormat
      .Combo2 = ""
      .Text5 = ""
      .Text8 = ""
      .Text6 = ""
      .Text7 = ""
      .Text9 = ""
      .Text12 = ""
      .Text11 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc2130_Clear()
   With Frmacc2130
      .Text2 = ""
      .Text1 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text6 = ""
      .Combo1 = ""
      .Text10 = ""
      .Text7 = ""
      .Text8 = ""
      .Text11 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc2140_Clear()
   With Frmacc2140
      .Text2 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox1.Mask = DFormat
      .Text1 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text6 = ""
      .MaskEdBox1.SetFocus
   End With
End Sub

Public Sub Frmacc21e0_Clear()
   With Frmacc21e0
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text1 = ""
      .Combo2 = ""
      .Text9 = ""
      .Text4 = ""
      .Text2 = ""
      .Text10 = ""
      .Text3 = ""
      .Text11 = ""
      .Text13 = ""
      .AdodcRefresh
      .SumShow
      .MaskEdBox1.SetFocus
   End With
End Sub

Public Sub Frmacc21f0_Clear()
   With Frmacc21f0
      .Text9 = ""
      .Text2 = ""
      .Text6 = ""
      .Text3 = ""
      .Text4 = ""
      .Text7 = ""
      .Text8 = ""
      .Text10 = "Y"
      .Label12.Caption = ""
      .Text12 = ""
      .Text13 = ""
      .Text14 = ""
      .AdodcRefresh1
      .AdodcRefresh2
      .Text9.SetFocus
   End With
End Sub

Public Sub Frmacc21f1_Clear()
   With Frmacc21f1
      .Text1 = ""
      .Text2 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Combo3 = ""
      .Text4 = ""
      .Text5 = ""
      .Combo1 = ""
      .Text7 = ""
      .AdodcRefresh
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc21f2_Clear()
   With Frmacc21f2
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Combo3 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Combo1 = ""
      .Combo2 = ""
      .AdodcRefresh
      .MaskEdBox1.SetFocus
   End With
End Sub

Public Sub Frmacc21n0_Clear()
   With Frmacc21n0
      .Combo1 = ""
      .Text5 = ""
      .Combo1.SetFocus
   End With
End Sub

'Add By Cheng 2003/07/22
Public Sub Frmacc21q0_Clear()
   With Frmacc21q0
        .Text1.Text = ""
        .Text1.Tag = "" 'Added by Lydia 2025/11/06
        .Text2.Text = ""
        .Text3.Text = ""
        .Text4.Text = ""
        .Text5.Text = ""
        .Text6.Text = ""
        .Text7.Text = ""
        .Text8.Text = ""
        .Text9.Text = ""
        .Text10.Text = ""
        .Text11.Text = ""
        .Text12.Text = ""
        .Combo2.Text = ""
        .Text14.Text = ""
        .Text15.Text = ""
        .Combo1.Text = ""
        .Text13.Text = ""
        .Text16.Text = ""
        .Text17.Text = ""
        .Text18.Text = ""
        .Text19.Text = ""
        .Text20.Text = ""
        'Added by Lydia 2015/03/31
        .Combo3.Text = ""
        .Text21.Text = ""
        .Text22.Text = ""
        .Text23.Text = ""
        .lblCNT.Caption = "" 'Added by Lydia 2017/09/07
        'Added by Lydia 2025/08/21
        .txtFAddr.Text = ""
        .txtA2224.Text = ""
        .txtA2225.Text = ""
        .lblA2225Name.Caption = ""
        'end 2025/08/21
        .txtA2226.Text = "" 'Added by Lydia 2025/10/17
        .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc3130_Clear()
   With Frmacc3130
      .Text5 = ""
      .Text9 = ""
      .Text10 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Combo1 = ""
      .Text11 = ""
      .Text12 = ""
      .Text5.SetFocus
   End With
End Sub

Public Sub Frmacc3180_Clear()
   With Frmacc3180
      .Text1 = ""
      .Text10 = ""
      .Text2 = ""
      .Text4 = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox2.Text = ""
      .Text6 = ""
      .Text1.SetFocus
      'Add by Morgan 2007/2/7
      .Text3 = ""
      .Text5 = ""
      'End 2007/2/7
   End With
End Sub

Public Sub Frmacc3190_Clear()
   With Frmacc3190
      .Text5 = ""
      .Text6 = ""
      .Text1 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text2 = ""
      .Text3 = ""
      .Text10 = ""
      .Text4 = ""
      'Add by Amy 2013/08/09 +出名人及存款類別欄位
      .Text7 = ""
      .Text8 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc4130_Clear()
   With Frmacc4130
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text6 = ""
      .Text10 = ""
      .Text11 = ""
      .Text7 = ""
      .Text8 = ""
      .Text9 = ""
      '.Text12 = ""   '2008/12/15 cancel by sonia
      'Added by Lydia 2017/09/06
      .txtAddr1 = ""
      .txtAddr2 = ""
      'end 2017/09/06
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc4140_Clear()
   With Frmacc4140
      .Text1 = ""
      .Text3 = ""
      .Text4 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc4150_Clear()
'   With Frmacc4150
'      .Text1 = ""
'      .Text2 = ""
'      .Text15 = ""
'      .Text16 = ""
'      .Text17 = ""
'      .Text18 = ""
'      .Text3 = ""
'      .Text9 = ""
'      .Text4 = ""
'      .Text10 = ""
'      .Text5 = ""
'      .Text11 = ""
'      .Text6 = ""
'      .Text12 = ""
'      .Text7 = ""
'      .Text13 = ""
'      .Text8 = ""
'      .Text14 = ""
'      .Text1.SetFocus
'   End With
End Sub

Public Sub Frmacc4180_Clear()
   With Frmacc4180
      .Text1 = ""
      .Text5 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc4190_Clear()
   With Frmacc4190
      .Text1 = "1"
      .Text6 = ""
      .Text5 = ""
      .Text4 = ""
      .Text3 = ""
      .Text1.SetFocus
   End With
End Sub

Public Sub Frmacc5100_Clear()
   With Frmacc5100
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .MaskEdBox3.Mask = ""
      .MaskEdBox3.Text = ""
      .MaskEdBox3.Mask = DFormat
      .Text1 = ""
   End With
End Sub

Public Sub Frmacc5200_Clear()
   With Frmacc5200
      .Text4 = "1"
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .Text5 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text1.SetFocus
   End With
End Sub
