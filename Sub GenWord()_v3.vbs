Function InitializeWordApp() As Object
    Dim WordApp As Object
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    Set InitializeWordApp = WordApp
End Function

Function AddParagraphToWord(ByRef WordDoc As Object, ByVal vaName As String, ByVal hostList As String, ByRef vaDescript As String, vaSolution As String) As Boolean
    Dim para As Object
    If vaName <> "" And hostList <> "" Then
        Set para = WordDoc.Content.Paragraphs.Add
        para.Range.text = "漏洞名稱： " & vaName & vbCrLf & "漏洞來源：" & hostList & vbCrLf & "漏洞敘述：" & vaDescript & vbCrLf & "修補建議：" & vbCrLf & vaSolution & vbCrLf
        AddParagraphToWord = True
    Else
        AddParagraphToWord = False
    End If
End Function

Sub ProcessExcelData(ByRef Worksheet As Object, ByVal lastRow As Long, ByRef vaName As String, ByRef hostList As String, ByRef vaDescript As String, vaSolution As String, ByRef WordDoc As Object)
    ' 由於我們需要操作 Word 文檔，我們將 WordDoc 作為參數傳遞給 ProcessExcelData
    Dim cellValue As String, i As Long
    For i = 1 To lastRow
        cellValue = Worksheet.Cells(i, 1).Value
        If Not IsNumeric(Left(cellValue, 1)) And cellValue <> "" Then
            If vaName <> "" And hostList <> "" Then
                ' 確保 AddParagraphToWord 能夠接收 WordDoc 參數
                If Not AddParagraphToWord(WordDoc, vaName, hostList, vaDescript, vaSolution) Then
                    MsgBox "無法將內容添加到 Word 文檔。"
                End If
            End If
            vaName = cellValue
            hostList = ""
        ElseIf IsNumeric(Left(cellValue, 1)) And cellValue <> "" Then
            If Len(hostList) > 0 Then
                hostList = hostList & "、"
            End If
            hostList = hostList & cellValue
        End If
    Next i
    ' 確保最後一次循環的數據也被添加到 Word 文檔
    If vaName <> "" And hostList <> "" Then
        If Not AddParagraphToWord(WordDoc, vaName, hostList, vaDescript, vaSolution) Then
            MsgBox "無法將最後一個漏洞名稱和主機IP列表添加到 Word 文檔。"
        End If
    End If
End Sub

Sub MainGenWord()
    Dim ExcelApp As Object, Workbook As Object, Worksheet As Object
    Dim WordApp As Object, WordDoc As Object
    Dim lastRow As Long
    Dim vaName As String, hostList As String, vaDescript As String, vaSolution As String

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
    
    ' 確保 lastRow 是 Long 類型的變數
    lastRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).row
    
    ' 調用 ProcessExcelData 時，需要將 WordDoc 作為參數傳遞
    ProcessExcelData Worksheet, lastRow, vaName, hostList, vaDescript, vaSolution, WordDoc
    
     ' 如果在最後還有未添加的數據，則添加到 Word 文檔
    If vaName <> "" And hostList <> "" Then
        If Not AddParagraphToWord(WordDoc, vaName, hostList, vaDescript, vaSolution) Then
            MsgBox "添加數據到 Word 文檔時發生錯誤。"
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

