Attribute VB_Name = "网站配置"
'模块功能：批量网站配置
Option Explicit
'这个模块有如下功能和对应的快捷键
'===== 功能1：ctrl+w =====
'输入不带www.的网址信息(例如：wingink.com)，Ctrl+ w实现自动填充如下信息：
'www.wingink.com
'http://wingink.com
'http://www.wingink.com
'RH *wingink.com
'
'===== 功能2：ctrl+b =====
'根据商户ID自动填充卡种、MCC、通道、是否实体、Forter状态和ID
'
'下面两个功能是为了方便复查提供的
'===== 功能3：ctrl+n =====
'根据商户ID自动填充商户名称
'
'===== 功能4：ctrl+m =====
'根据商户名称自动填充商户ID
'
'===== 功能5：ctrl+d =====
'备份上传的应用条目并删除记录
Dim sht As Worksheet
Dim last_row

Sub OneClickAway():
Attribute OneClickAway.VB_ProcData.VB_Invoke_Func = "b\n14"
' 快捷键: Ctrl+b
    InducedByAppName
    TurnNameIntoID
    CheckMerchantByID
    CheckBuildingTools
    CheckChannels
    DeleteFullNameColumn
End Sub

Sub DeleteFullNameColumn():
    Dim sht As Worksheet
    Set sht = Worksheets("websites")
    If sht.Cells(1, 2) = "" Then
        sht.Cells(1, 2).EntireColumn.Delete
    End If
End Sub

Sub CheckBuildingTools()
    '检查建站工具是否填写
    Dim i, loadings
    Set loadings = Worksheets("websites")
    last_row = loadings.Cells(Rows.count, 1).End(xlUp).Row
    Dim app_col As Integer
    app_col = loadings.Range("1:1").Find("应用名称", LookIn:=xlValues).Column
    For i = 2 To last_row
        If loadings.Cells(i, app_col).Offset(0, 5) = "" Then
            loadings.Cells(i, app_col).Offset(0, 5).Interior.Color = RGB(0, 200, 200)
            MsgBox CStr(loadings.Cells(i, 1).Value) + "的建站工具喃？"
        Else
            loadings.Cells(i, app_col).Offset(0, 5).Interior.ColorIndex = 0
        End If
    Next i
    
End Sub


Sub CheckChannels():
'检查通道
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "[^,0-9]"
        .Global = True
    End With
    
    Dim sht As Worksheet
    Set sht = Worksheets("websites")
    Dim i, app_col, last_row As Integer
    app_col = sht.Range("A1").EntireRow.Find("应用名称", LookIn:=xlValues).Column
    last_row = Worksheets("websites").Cells(Rows.count, app_col).End(xlUp).Row
    For i = 2 To last_row
        Dim channels As String
        channels = sht.Cells(i, app_col).Offset(0, 4).Value
        If regex.Test(channels) Then
            sht.Cells(i, app_col).Offset(0, 4).Font.Color = vbRed
            sht.Cells(i, app_col).Offset(0, 4).Font.Bold = True
        Else
            sht.Cells(i, app_col).Offset(0, 4).Font.Color = vbBlack
            sht.Cells(i, app_col).Offset(0, 4).Font.Bold = False
        End If
    Next i

End Sub


Sub InducedByAppName()
'自动生成跟应用名称相关的单元格：应用网址、账单地址、隐私声明、条款&条件
    Dim app, billAdress, i, last_row, app_colno
    Dim sites As Worksheet
    Set sites = Worksheets("websites")
    app_colno = sites.Range("A1").EntireRow.Find("应用名称", LookIn:=xlValues).Column
    
    last_row = Worksheets("websites").Cells(Rows.count, app_colno).End(xlUp).Row
    For i = 2 To last_row
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Pattern = "[\s\u4e00-\u9fa5/（）]" 'Pattern - look for 空格、中文字符、斜杠（/）in the string
            .Global = True 'If False, would replace only first
        End With
    
        If sites.Cells(i, app_colno) <> "" Then
            app = sites.Cells(i, app_colno)
            app = LCase(app)
            app = regex.Replace(app, "")
            regex.Pattern = "(https?:)?(w{3}\.)?" 'Pattern - look for (http:)and (www.)in the string
            regex.Global = True
            app = regex.Replace(app, "")
            
            regex.Pattern = "\.comshopify"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.netshopify"
            regex.Global = True
            app = regex.Replace(app, ".net")
            
             regex.Pattern = "\.orgshopify"
            regex.Global = True
            app = regex.Replace(app, ".org")
            
            regex.Pattern = "\.pubshopify"
            regex.Global = True
            app = regex.Replace(app, ".pub")
            
            regex.Pattern = "\.saleshopify"
            regex.Global = True
            app = regex.Replace(app, ".sale")
            
            regex.Pattern = "\.deshopify"
            regex.Global = True
            app = regex.Replace(app, ".de")
            
             regex.Pattern = "\.techshopify"
            regex.Global = True
            app = regex.Replace(app, ".tech")
            
             regex.Pattern = "\.shopshopify"
            regex.Global = True
            app = regex.Replace(app, ".shop")
            
            regex.Pattern = "\.comxshoppy"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.comshopyy"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.comshoplazza"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.comshopbase"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.comfunpinpin"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            regex.Pattern = "\.comshopline"
            regex.Global = True
            app = regex.Replace(app, ".com")
            
            sites.Cells(i, app_colno) = app
            sites.Cells(i, app_colno).Offset(0, 2) = "http://" + app
            sites.Cells(i, app_colno).Offset(0, 7) = "RH *" + app
    '            sites.Cells(i, app_colno).Offset(0, 7) = app
            sites.Cells(i, app_colno).Offset(0, 10) = "http://" + app
            sites.Cells(i, app_colno).Offset(0, 11) = "http://" + app
        End If
    Next i
  
End Sub


Sub FillMechartInfo():
    '通过商户号，填充信息
    Dim websites As Worksheet
    Set websites = Worksheets("websites")
    
    Dim merchants As Worksheet
    Set merchants = Worksheets("merchant_info")
    
    Dim i, x, last_row_ws, last_row_info, app_col, mer_id
    app_col = 0
    mer_id = ""
    last_row_ws = websites.Cells(Rows.count, 1).End(xlUp).Row
    last_row_info = merchants.Cells(Rows.count, 1).End(xlUp).Row
    app_col = websites.Range("A1").EntireRow.Find("应用名称", LookIn:=xlValues).Column
    For i = 2 To last_row_ws
        If websites.Cells(i, 1) <> "" Then
            For x = 3 To last_row_info Step 2
                mer_id = merchants.Cells(x, 1)
                If websites.Cells(i, 1) = mer_id Then
                    websites.Cells(i, app_col).Offset(0, 1) = merchants.Cells(x, 3) '币种
                    websites.Cells(i, app_col).Offset(0, 3) = merchants.Cells(x, 5)  'MCC
                    websites.Cells(i, app_col).Offset(0, 4) = merchants.Cells(x, 6) '通道
                    websites.Cells(i, app_col).Offset(0, 6) = merchants.Cells(x, 8) '是否实体
                    websites.Cells(i, app_col).Offset(0, 8) = merchants.Cells(x, 10) 'Forter状态
                    websites.Cells(i, app_col).Offset(0, 9) = merchants.Cells(x, 11) 'ForterID
                    merchants.Cells(x, 1).Copy
                    websites.Cells(i, 1).EntireRow.PasteSpecial Paste:=xlPasteFormats
                    x = last_row_info '跳出内循环
                End If
                
            Next x
        End If
        
    Next i
    
    
End Sub


Sub TurnNameIntoNum():
'根据商户名获取商户ID
    Dim i, j, givenEn, lastrow
    Dim loadings As Worksheet
    Dim entities As Worksheet
    Dim rg As Range
    Set loadings = Worksheets("websites")
    Set entities = Worksheets("merchant_info")
    lastrow = loadings.Cells(Rows.count, 2).End(xlUp).Row
    Set rg = entities.Columns(2)

    
    For j = 2 To lastrow

        givenEn = Trim(loadings.Cells(j, 2)) '查看目标商户名称
        givenEn = LCase(givenEn)
        If Not rg.Find(givenEn, LookIn:=xlValues) Is Nothing Then
            Dim r
            r = rg.Find(givenEn, LookIn:=xlValues).Row
            entities.Cells(r + 1, 1).Copy
            loadings.Cells(j, 1).PasteSpecial Paste:=xlPasteFormats
            loadings.Cells(j, 1).PasteSpecial Paste:=xlPasteValues
        End If

    Next j

End Sub

Sub TurnNumIntoNam()
'
'
' 根据商户ID获取商户名称
'
'
'

    Dim i, j, givenEn, lastrow, lr
    Dim loadings As Worksheet
    Dim entities As Worksheet
    Dim rg As Range
    Set loadings = Worksheets("websites")
    Set entities = Worksheets("merchant_info")
    lr = entities.Cells(Rows.count, 1).End(xlUp).Row
    Set rg = entities.Range("A2:A" + CStr(lr))
    lastrow = loadings.Cells(Rows.count, 1).End(xlUp).Row
    For j = 2 To lastrow
        givenEn = loadings.Cells(j, 1) '查看Sheet1的目标商户ID
        If Not rg.Find(givenEn, LookIn:=xlValues) Is Nothing Then
            Dim r
            r = rg.Find(givenEn, LookIn:=xlValues).Row
            entities.Cells(r - 1, 1).Copy '商户名
            loadings.Cells(j, 2).PasteSpecial Paste:=xlPasteFormats
            loadings.Cells(j, 2).PasteSpecial Paste:=xlPasteValues
        End If
    Next j
End Sub

Sub TurnNameIntoID():
    '根据商户名称获取商户ID
    Dim sht As Worksheet
    Set sht = Worksheets("websites")
    
    Dim info As Worksheet
    Set info = Worksheets("merchants")
    Dim name_rg As Range
    Set name_rg = info.Columns("B:C")
    Dim merchant As String
    Dim lr, app_col As Integer
    Dim i As Integer
    lr = sht.Cells(Rows.count, 1).End(xlUp).Row
    app_col = sht.Range("1:1").Find("应用名称", LookIn:=xlValues).Column
    For i = 2 To lr
        merchant = Trim(CStr(sht.Cells(i, 1)))
        sht.Cells(i, 1) = merchant
        Dim r As Integer
        If Not name_rg.Find(merchant, LookIn:=xlValues) Is Nothing Then
            r = name_rg.Find(merchant, LookIn:=xlValues).Row
            sht.Cells(i, 1) = Trim(info.Cells(r, 1))
        End If
    Next i
End Sub


Sub CheckMerchantByID():
    '根据商户号获取【币种、MCC、通道ID、是否实体、Forter状态和ID】、
    '根据颜色判断新、老系统：灰色只加老系统、黄色新老系统都要加、粉色只加新系统
    Dim sht As Worksheet
    Set sht = Worksheets("websites")
    
    Dim info As Worksheet
    Set info = Worksheets("merchants")
    Dim num_rg As Range
    Set num_rg = info.Columns(1)
    Dim merchant As String
    Dim lr, app_col As Integer
    Dim i As Integer
    lr = sht.Cells(Rows.count, 1).End(xlUp).Row
    app_col = sht.Range("1:1").Find("应用名称", LookIn:=xlValues).Column
    For i = 2 To lr
        merchant = Trim(CStr(sht.Cells(i, 1)))
        sht.Cells(i, 1) = merchant
        Dim r As Integer
        If Not num_rg.Find(merchant, LookIn:=xlValues) Is Nothing Then
            r = num_rg.Find(merchant, LookIn:=xlValues).Row
            sht.Cells(i, app_col).Offset(0, 1) = info.Cells(r, 4) '卡种
            sht.Cells(i, app_col).Offset(0, 3) = info.Cells(r, 5) 'MCC
            sht.Cells(i, app_col).Offset(0, 4) = info.Cells(r, 6) '通道ID
            sht.Cells(i, app_col).Offset(0, 6) = info.Cells(r, 7) '是否实体
            sht.Cells(i, app_col).Offset(0, 8) = info.Cells(r, 8) 'Forter状态
            sht.Cells(i, app_col).Offset(0, 9) = info.Cells(r, 9) 'ForterID
            num_rg.Find(merchant, LookIn:=xlValues).Copy
            sht.Cells(i, app_col).EntireRow.PasteSpecial Paste:=xlPasteFormats
        End If
    Next i
End Sub






