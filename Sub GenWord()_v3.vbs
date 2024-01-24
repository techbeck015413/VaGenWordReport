Function InitializeWordApp() As Object
    Dim WordApp As Object
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    Set InitializeWordApp = WordApp
End Function

Function AddParagraphToWord(ByRef WordDoc As Object, ByVal vaName As String, ByVal hostList As String, ByRef vaDescript As String, ByRef vaSolution As String) As Boolean
    Dim para As Object
    If vaName <> "" And hostList <> "" Then
        Set para = WordDoc.Content.Paragraphs.Add
        para.Range.text = "漏洞名稱： " & vaName & vbCrLf & "漏洞來源：" & hostList & vbCrLf & "漏洞敘述：" & vbCrLf & vaDescript & vbCrLf & "修補建議：" & vbCrLf & vaSolution & vbCrLf & vbCrLf
        AddParagraphToWord = True
    Else
        AddParagraphToWord = False
    End If
End Function

'主函式
Sub MainGenWord()
    Dim ExcelApp As Object, Workbook As Object, Worksheet As Object
    Dim WordApp As Object, WordDoc As Object
    Dim lastRow As Long
    Dim vaName As String, hostList As String, vaDescript As String, vaSolution As String
    Dim cellValue As String, i As Long
    Dim searchResult As MSHTML.HTMLDocument

    Set ExcelApp = GetExistingExcelApp()
    Set Workbook = ExcelApp.ActiveWorkbook
    Set Worksheet = Workbook.Sheets("Name-Host-2")
    lastRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).row

    Set WordApp = InitializeWordApp()
    Set WordDoc = WordApp.Documents.Add

    vaName = ""
    hostList = ""
    vaDescript = ""
    vaSolution = ""

    For i = 1 To lastRow
        cellValue = Worksheet.Cells(i, 1).Value
        If Not IsNumeric(Left(cellValue, 1)) And cellValue <> "" Then
            If vaName <> "" And hostList <> "" Then
                Set searchResult = SearchVaName(vaName)
                vaDescript = ProcessVaDescript(searchResult)
                vaSolution = ProcessVaSolution(searchResult)
    
                If Not AddParagraphToWord(WordDoc, vaName, hostList, vaDescript, vaSolution) Then
                    MsgBox "無法將內容添加到 Word 文檔。"
                End If
            End If
            vaName = cellValue
            hostList = ""
            vaDescript = "[預覽描述2]"
            vaSolution = "[預覽解決方式2]"
        ElseIf IsNumeric(Left(cellValue, 1)) And cellValue <> "" Then
            If Len(hostList) > 0 Then
                hostList = hostList & "、"
            End If
            hostList = hostList & cellValue
        End If
    Next i

    If vaName <> "" And hostList <> "" Then
        searchResult = SearchVaName(vaName)
        vaDescript = ProcessVaDescript(searchResult)
        vaSolution = ProcessVaSolution(searchResult)
    
    If Not AddParagraphToWord(WordDoc, vaName, hostList, vaDescript, vaSolution) Then
        MsgBox "無法將內容添加到 Word 文檔。"
    End If
End If

    WordApp.Selection.HomeKey Unit:=wdStory

    CleanUp ExcelApp, Workbook, Worksheet, WordDoc, WordApp
End Sub

Sub CleanUp(ByRef ExcelApp As Object, ByRef Workbook As Object, ByRef Worksheet As Object, ByRef WordDoc As Object, ByRef WordApp As Object)
    Set Worksheet = Nothing
    Set Workbook = Nothing
    Set ExcelApp = Nothing
    Set WordDoc = Nothing
    Set WordApp = Nothing
End Sub

Function GetExistingExcelApp() As Object
    On Error Resume Next
    Set GetExistingExcelApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    If GetExistingExcelApp Is Nothing Then
        MsgBox "無法找到打開的 Excel 應用程式。"
End
End If
End Function

'搜尋功能
Function SearchVaName(vaName As String) As MSHTML.HTMLDocument
    Dim httpRequest As Object, htmlDoc As MSHTML.HTMLDocument
    Dim apiKey As String, searchEngineId As String, apiUrl As String, searchResult As String
    Dim json As Object, firstResultUrl As String

    ' API 設置
    apiKey = "AIzaSyB8MbtxSx6uiwASlpm7_us-Uy6fV1uHSWY"  ' 您的API密鑰
    searchEngineId = "31b1d7ed47ae1422c"  ' 您的搜索引擎ID
    apiUrl = "https://www.googleapis.com/customsearch/v1?q=" & _
             URLEncode(vaName) & "&cx=" & searchEngineId & "&key=" & apiKey

    ' 發送HTTP請求到Google Custom Search API
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", apiUrl, False
    httpRequest.Send

    ' 解析JSON回應並獲取第一個搜尋結果的URL
    If httpRequest.Status = 200 Then
        searchResult = httpRequest.responseText
        Set json = JsonConverter.ParseJson(searchResult)
        firstResultUrl = json("items")(1)("link")
    Else
        MsgBox "Error: " & httpRequest.Status & " - " & httpRequest.statusText
        Exit Function
    End If

    ' 使用第一個搜尋結果的URL發送另一個HTTP請求
    httpRequest.Open "GET", firstResultUrl, False
    httpRequest.Send

    ' 創建HTMLDocument對象
    Set htmlDoc = New MSHTML.HTMLDocument
    If httpRequest.Status = 200 Then
        ' 將獲取的HTML設置為HTMLDocument的內容
        htmlDoc.body.innerHTML = httpRequest.responseText
    Else
        MsgBox "Error: " & httpRequest.Status & " - " & httpRequest.statusText
    End If

    Set SearchVaName = htmlDoc
    Set httpRequest = Nothing
End Function
Function URLEncode(str As String) As String
    Dim reservedChars As String
    Dim i As Integer
    Dim ch As String
    URLEncode = ""
    reservedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"

    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If InStr(reservedChars, ch) <> 0 Then
            URLEncode = URLEncode & ch
        Else
            URLEncode = URLEncode & "%" & Right("0" & Hex(Asc(ch)), 2)
        End If
    Next
End Function

Function ProcessVaDescript(htmlDoc As MSHTML.HTMLDocument) As String
    Dim sectionElements As MSHTML.IHTMLElementCollection
    Dim sectionElement As MSHTML.IHTMLElement
    Dim descriptionText As String
    Dim spanElement As MSHTML.IHTMLElement
    Dim titleText As String

    descriptionText = ""

    ' 尋找所有的<section>元素
    Set sectionElements = htmlDoc.getElementsByTagName("section")

    ' 遍歷所有<section>元素
    For Each sectionElement In sectionElements
        ' 檢查<section>元素是否包含一個<h4>子元素，且其文本為"Description"或"說明"
        Dim h4Element As MSHTML.IHTMLElement
        For Each h4Element In sectionElement.getElementsByTagName("h4")
            titleText = h4Element.innerText
            If h4Element.ClassName = "border-bottom pb-1" And (titleText = "Description" Or titleText = "說明") Then
                ' 找到了Description或說明標題，接下來將提取所有SPAN元素的文本
                For Each spanElement In sectionElement.getElementsByTagName("span")
                    descriptionText = descriptionText & spanElement.innerText
                Next spanElement
                ' 添加一個段落間隔
                descriptionText = descriptionText & vbCrLf
                Exit For ' 已經找到並處理了相應的<section>，跳出循環
            End If
        Next h4Element
    Next sectionElement

    ProcessVaDescript = Trim(descriptionText) ' 移除最後的換行符號
End Function




' 處理修補建議的函式
Function ProcessVaSolution(htmlDoc As MSHTML.HTMLDocument) As String
    Dim sectionElements As MSHTML.IHTMLElementCollection
    Dim sectionElement As MSHTML.IHTMLElement
    Dim solutionText As String
    Dim spanElement As MSHTML.IHTMLElement
    Dim titleText As String

    solutionText = "" ' 修正變數名稱

    ' 尋找所有的<section>元素
    Set sectionElements = htmlDoc.getElementsByTagName("section")

    ' 遍歷所有<section>元素
    For Each sectionElement In sectionElements
        ' 檢查<section>元素是否包含一個<h4>子元素，且其文本為"Solution"或"解決方案"
        Dim h4Element As MSHTML.IHTMLElement
        For Each h4Element In sectionElement.getElementsByTagName("h4")
            titleText = h4Element.innerText ' 獲取標題文本
            If h4Element.ClassName = "border-bottom pb-1" And (titleText = "Solution" Or titleText = "解決方案") Then
                ' 找到了Solution或解決方案標題，接下來將提取所有SPAN元素的文本
                For Each spanElement In sectionElement.getElementsByTagName("span")
                    solutionText = solutionText & spanElement.innerText
                Next spanElement
                solutionText = solutionText & vbCrLf ' 添加段落間隔
                Exit For ' 已經找到並處理了相應的<section>，跳出循環
            End If
        Next h4Element
    Next sectionElement
    ProcessVaSolution = Trim(solutionText) ' 移除最後的換行符號
End Function




'翻譯功能(還未添加)
' ... 如果您未來有添加翻譯功能的代碼，可以在這裡加入 ...



