Sub Gen_PPT()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim slideIndex As Integer
    Dim row As Long
    Dim vaName As String
    Dim ipPattern As String
    Dim regex As Object
    Dim match As Object
    Dim ipList As String
    Dim additionalText As String
    Dim cellValue As String
    
    ' 開啟 PowerPoint 應用程式
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    
    ' 新增一個簡報
    Set pptPres = pptApp.Presentations.Add
    
    ' 指定要讀取的 Excel 工作表
    Set ws = ThisWorkbook.Sheets("Name-Host-2")
    
    ' 找出 Excel 表格的最後一列，排除空白列
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' 正規表達式來匹配 IP 地址
    ipPattern = "\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b"
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = ipPattern
    
   
    ' 設定 PowerPoint 中的投影片索引
    slideIndex = 1
    
    ' 從第一列開始處理每一列資料
    For row = 1 To lastRow
        ' 讀取到的資料
        cellValue = ws.Cells(row, 1).Value
        
        ' 檢查是否為漏洞名稱
        If Not regex.Test(cellValue) Then
            ' 從這邊開始改
            ' 是漏洞名稱，調用 GetSearchResultHtml 並處理結果
            Dim htmlDoc As HTMLDocument
            Dim nessusUrl As String
            Set htmlDoc = GetSearchResultHtml(cellValue, nessusUrl)
            If Not htmlDoc Is Nothing Then
                ' 調用處理函數來整理 HTML 內容
                Dim processedContent As String
                additionalText = ProcessSearchResult(htmlDoc, nessusUrl)
            Else
                additionalText = "htmlDoc未正常運作"
            End If
        End If
        
        
        
        ' 檢查是否為 "列標籤" 或 "總計" 或空白行
        If cellValue = "列標籤" Or cellValue = "總計" Or Trim(cellValue) = "" Then
            If cellValue = "總計" Then
                Set pptSlide = pptPres.Slides.Add(slideIndex, ppLayoutText)
                pptPres.Slides(slideIndex).Delete
            End If
            GoTo ContinueLoop
        End If
        
        ' 初始化 IP 地址列表
        ipList = ""
        
        ' 檢查是否為漏洞名稱
        If regex.Test(cellValue) = False Then
            vaName = cellValue
            
            ' 新增一個投影片
            Set pptSlide = pptPres.Slides.Add(slideIndex, ppLayoutText)
            
            ' 在投影片中新增文字框並填入標題
            pptSlide.Shapes.title.TextFrame.TextRange.text = vaName
            
            ' 確保標題沒有項目符號
            pptSlide.Shapes.title.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = msoFalse
            
            If vaName = "Critical" Or vaName = "High" Or vaName = "Medium" Or vaName = "Low" Then
                ' 如果標題是這些，則跳過創建 pptSlide.Shapes(2)
                slideIndex = slideIndex + 1
                GoTo ContinueLoop
            End If
            
            

            ' 添加主要內容
            pptSlide.Shapes(2).TextFrame.TextRange.text = "Host：" & ipList & vbCrLf & additionalText
            
            ' 確保主要內容沒有項目符號
            pptSlide.Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = msoFalse
            
            ' 設置字體大小為 14
            With pptSlide.Shapes(2).TextFrame.TextRange
                .Font.Size = 14
            End With
            
            ' 從當前行開始，檢查下面的行是否為 IP 地址
            row = row + 1
            While row <= lastRow And regex.Test(ws.Cells(row, 1).Value)
                ' 如果找到 IP 地址，則加入到 IP 列表中
                If ipList <> "" Then
                    ipList = ipList & "、"
                End If
                ipList = ipList & ws.Cells(row, 1).Value
                row = row + 1
            Wend
            
            ' 如果有找到 IP 地址，則將其添加到本文
            If ipList <> "" Then
                ' 添加主要內容
                pptSlide.Shapes(2).TextFrame.TextRange.text = "Host：" & ipList & vbCrLf & additionalText
                
                ' 確保主要內容沒有項目符號
                pptSlide.Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = msoFalse
            End If
            
            ' 調整投影片索引
            slideIndex = slideIndex + 1
            
            ' 循環已經檢查到非 IP 地址的行，應該回退一行以檢查下一個潛在的漏洞名稱
            row = row - 1
        End If
        
ContinueLoop:
    Next row
    
    ' 如果第一頁是空白，則刪除
    If pptPres.Slides.Count >= 1 Then
        If pptPres.Slides(1).Shapes.Count = 0 Then
            pptPres.Slides(1).Delete
        End If
    End If
    
    ' 釋放物件
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set regex = Nothing

End Sub

Function GetSearchResultHtml(query As String, ByRef nessusUrl As String) As HTMLDocument
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")

    Dim apiKey As String, searchEngineId As String, apiUrl As String
    apiKey = "AIzaSyB8MbtxSx6uiwASlpm7_us-Uy6fV1uHSWY"  ' 您的API密鑰
    searchEngineId = "31b1d7ed47ae1422c"  ' 您的搜索引擎ID
    apiUrl = "https://www.googleapis.com/customsearch/v1?q=" & query & "&cx=" & searchEngineId & "&key=" & apiKey
    
    ' 發送HTTP請求以進行搜索
    With httpRequest
        .Open "GET", apiUrl, False
        .Send
    End With

    ' 解析API響應以獲取第一個搜索結果的URL
    Dim firstResultUrl As String
    If httpRequest.Status = 200 Then
        Dim jsonResponse As Object
        Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)
        
        If jsonResponse("items").Count > 0 Then
            firstResultUrl = jsonResponse("items")(1)("link")

            ' 判斷URL是否包含指定的子字串
            If InStr(firstResultUrl, "/changelog") > 0 Then
                firstResultUrl = Left(firstResultUrl, InStr(firstResultUrl, "/changelog") - 1)
            ElseIf InStr(firstResultUrl, "/dependencies") > 0 Then
                firstResultUrl = Left(firstResultUrl, InStr(firstResultUrl, "/dependencies") - 1)
            ElseIf InStr(firstResultUrl, "/dependents") > 0 Then
                firstResultUrl = Left(firstResultUrl, InStr(firstResultUrl, "/dependents") - 1)
            End If
    
            nessusUrl = firstResultUrl  ' 在這裡更新 nessusUrl 為第一個搜索結果的URL
            
            ' 使用新的HTTP請求來獲取第一個搜索結果的網頁內容
            With httpRequest
                .Open "GET", firstResultUrl, False
                .Send
                If .Status = 200 Then
                    ' 創建並返回一個HTMLDocument物件
                    Dim doc As HTMLDocument
                    Set doc = New HTMLDocument
                    doc.body.innerHTML = .responseText
                    Set GetSearchResultHtml = doc
                Else
                    Set GetSearchResultHtml = Nothing
                End If
            End With
        Else
            Set GetSearchResultHtml = Nothing
        End If
    Else
        Set GetSearchResultHtml = Nothing
    End If

    Set httpRequest = Nothing
End Function


Function ProcessSearchResult(htmlDoc As HTMLDocument, nessusUrl As String) As String
    Dim sectionTags As Object, elem As Object, item As Object
    Dim descript_str As String, solution_str As String, nessus_id_url As String, seealso_str As String
    Dim foundDescription As Boolean, foundSolution As Boolean, foundSeeAlso As Boolean
    
    nessus_id_url = nessusUrl
    
    ' 初始化布爾變數
    foundDescription = False
    foundSolution = False
    foundSeeAlso = False

    ' 使用getElementsByTagName尋找所有的<section>標籤
    Set sectionTags = htmlDoc.getElementsByTagName("section")
    
    For Each elem In sectionTags
        If InStr(elem.innerHTML, "Description") > 0 Or InStr(elem.innerHTML, "說明") > 0 Then
            descript_str = ""
            For Each item In elem.getElementsByTagName("span")
                If Trim(item.innerText) <> "" Then
                    descript_str = descript_str & item.innerText
                    If Not item Is elem.getElementsByTagName("span").item(elem.getElementsByTagName("span").Length - 1) Then
                        descript_str = descript_str & vbCrLf
                    End If
                End If
            Next item
            foundDescription = True
        ElseIf InStr(elem.innerHTML, "Solution") > 0 Or InStr(elem.innerHTML, "解決方案") > 0 Then
            solution_str = ""
            For Each item In elem.getElementsByTagName("span")
                If Trim(item.innerText) <> "" Then
                    solution_str = solution_str & item.innerText
                    If Not item Is elem.getElementsByTagName("span").item(elem.getElementsByTagName("span").Length - 1) Then
                        solution_str = solution_str & vbCrLf
                    End If
                End If
            Next item
            foundSolution = True
        ElseIf InStr(elem.innerHTML, "See Also") > 0 Or InStr(elem.innerHTML, "另請參閱") > 0 Then
            seealso_str = ""
            For Each item In elem.getElementsByTagName("p")
                seealso_str = seealso_str & item.innerText & vbCrLf
            Next item
            For Each item In elem.getElementsByTagName("span")
                seealso_str = seealso_str & item.innerText & vbCrLf
            Next item
            foundSeeAlso = True
        End If
    Next elem

    ' 翻譯部分
    Dim ErrorMessage As String

    ' 清理並翻譯 descript_str
    descript_str = CleanString(descript_str)
    If IsEnglish(descript_str) Then
        descript_str = TranslateText(descript_str, "zh-TW", ErrorMessage)
        If ErrorMessage <> "" Then
            descript_str = descript_str & " (英文未翻譯)"
        End If
    End If

    ' 清理並翻譯 solution_str
    solution_str = CleanString(solution_str)
    If IsEnglish(solution_str) Then
        solution_str = TranslateText(solution_str, "zh-TW", ErrorMessage)
        If ErrorMessage <> "" Then
            solution_str = solution_str & " (英文未翻譯)"
        End If
    End If

    ' seeAlso_str 不進行翻譯

    ' 在翻譯後處理 "-"，只處理 " - " 模式
    descript_str = Replace(descript_str, " - ", vbCrLf & " - ")
    solution_str = Replace(solution_str, " - ", vbCrLf & " - ")
    
    descript_str = Replace(descript_str, "請注意，", vbCrLf & "請注意，")
    solution_str = Replace(solution_str, "請注意，", vbCrLf & "請注意，")
    
    ' 組合文字
    Dim additionalText As String
    additionalText = "漏洞描述：" & descript_str & vbCrLf & _
                     "修補建議：" & solution_str & vbCrLf & _
                     "Nessus ID：" & nessus_id_url & vbCrLf & _
                     "相關連結：" & vbCrLf & seealso_str

    ' 返回處理後的字符串
    ProcessSearchResult = additionalText
End Function

Function IsEnglish(text As String) As Boolean
    Dim i As Long
    Dim englishCount As Double
    englishCount = 0
    Dim totalLength As Double
    totalLength = Len(text)

    ' 如果文本是空的，則直接返回 False
    If totalLength = 0 Then
        IsEnglish = False
        Exit Function
    End If

    For i = 1 To totalLength
        Dim c As String
        c = Mid(text, i, 1)
        If c Like "[a-zA-Z0-9 .,;:'""?!-]" Then
            englishCount = englishCount + 1
        End If
    Next i

    ' 計算英文字符比例並與閾值比較
    IsEnglish = (englishCount / totalLength) > 0.8
End Function

Function TranslateText(text As String, targetLanguage As String, ByRef errorMsg As String) As String
    Dim URL As String
    Dim objHTTP As Object
    Dim response As String
    Dim json As Object

    ' 初始化錯誤信息
    errorMsg = ""
    
    ' 在這裡定義您的 Google Translate API 金鑰
    Dim apiKey As String
    apiKey = "AIzaSyASb0YQ1eBsXHPVRuee7hTwKDHj9V0SBnc"

    On Error GoTo ErrorHandler
    ' Google翻譯API的URL
    URL = "https://translation.googleapis.com/language/translate/v2?q=" & _
          URLEncode(text) & "&target=" & targetLanguage & "&key=" & apiKey

    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "GET", URL, False
    objHTTP.Send

    If objHTTP.Status <> 200 Then
        errorMsg = "Error " & objHTTP.Status & ": " & objHTTP.statusText
        TranslateText = ""
        GoTo CleanUp
    End If

    response = objHTTP.responseText
    Set json = JsonConverter.ParseJson(response)
    TranslateText = json("data")("translations")(1)("translatedText")
    GoTo CleanUp

ErrorHandler:
    errorMsg = "HTTP Request Error: " & Err.Description
    TranslateText = ""
    Resume CleanUp

CleanUp:
    Set objHTTP = Nothing
End Function

Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long: StringLen = Len(StringVal)
    
    If StringLen > 0 Then
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String
        
        If SpaceAsPlus Then Space = "+" Else Space = "%20"
        
        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)
            
            Select Case CharCode
                Case 48 To 57, 65 To 90, 97 To 122
                    URLEncode = URLEncode & Char
                Case 32
                    URLEncode = URLEncode & Space
                Case Else
                    URLEncode = URLEncode & "%" & Right$("0" & Hex(CharCode), 2)
            End Select
        Next i
    End If
End Function

Function CleanString(str As String) As String
    Dim i As Long
    Dim cleanedStr As String
    cleanedStr = ""

    ' 遍歷字符串的每個字符
    For i = 1 To Len(str)
        Dim c As String
        c = Mid(str, i, 1)

        ' 檢查字符是否為標準ASCII字符範圍
        If Asc(c) >= 32 And Asc(c) <= 126 Then
            cleanedStr = cleanedStr & c
        Else
            ' 對於非標準字符，可以選擇替換為空格或其他字符
            cleanedStr = cleanedStr & " "
        End If
    Next i

    CleanString = cleanedStr
End Function

