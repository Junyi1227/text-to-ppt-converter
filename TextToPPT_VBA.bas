' ========================================
' Text to PowerPoint Converter
' 文字轉 PowerPoint 轉換器
' ========================================
' 使用方式：
' 1. 在 PowerPoint 中按 Alt+F11 開啟 VBA 編輯器
' 2. 插入 > 模組，將此程式碼貼上
' 3. 執行 ConvertTextToPPT 巨集
' ========================================

Option Explicit

' 主程式：將文字轉換為 PowerPoint 投影片
Sub ConvertTextToPPT()
    Dim textContent As String
    Dim lines() As String
    Dim i As Integer
    Dim pres As Presentation
    Dim slide As slide
    
    ' 取得文字內容（可以從剪貼簿或檔案讀取）
    textContent = GetTextInput()
    
    If textContent = "" Then
        MsgBox "沒有輸入內容！", vbExclamation
        Exit Sub
    End If
    
    ' 分割文字為行
    lines = Split(textContent, vbCrLf)
    
    ' 使用目前的簡報
    Set pres = ActivePresentation
    
    ' 處理每一行
    Dim currentSlide As slide
    Dim slideType As String
    Dim content As String
    
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' 檢查是否為標題標記
        If Left(line, 2) = "##" Then
            ' 主題頁面
            content = Trim(Mid(line, 3))
            Set currentSlide = CreateTitleSlide(pres, content)
            
        ElseIf Left(line, 1) = "#" Then
            ' 內文頁面
            content = Trim(Mid(line, 2))
            Set currentSlide = CreateContentSlide(pres, content)
            
        ElseIf line <> "" And Not currentSlide Is Nothing Then
            ' 添加內容到目前投影片
            AddContentToSlide currentSlide, line
        End If
    Next i
    
    MsgBox "完成！已建立 " & pres.Slides.Count & " 張投影片", vbInformation
End Sub

' 建立主題投影片（## 標記）
Function CreateTitleSlide(pres As Presentation, title As String) As slide
    Dim sld As slide
    Dim titleShape As Shape
    
    ' 新增投影片（使用標題投影片版面配置）
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutTitle)
    
    ' 設定標題
    With sld.Shapes(1).TextFrame.TextRange
        .Text = title
        .Font.Size = 44
        .Font.Bold = msoTrue
        .Font.Name = "微軟正黑體"
        .ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    ' 設定背景顏色（淺藍色）
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.ForeColor.RGB = RGB(230, 240, 255)
    
    Set CreateTitleSlide = sld
End Function

' 建立內文投影片（# 標記）
Function CreateContentSlide(pres As Presentation, title As String) As slide
    Dim sld As slide
    
    ' 新增投影片（使用標題加內容版面配置）
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutText)
    
    ' 設定標題
    With sld.Shapes(1).TextFrame.TextRange
        .Text = title
        .Font.Size = 32
        .Font.Bold = msoTrue
        .Font.Name = "微軟正黑體"
    End With
    
    ' 清空內容區域
    If sld.Shapes.Count >= 2 Then
        sld.Shapes(2).TextFrame.TextRange.Text = ""
    End If
    
    ' 設定背景顏色（淺灰色）
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.ForeColor.RGB = RGB(245, 245, 245)
    
    Set CreateContentSlide = sld
End Function

' 新增內容到投影片
Sub AddContentToSlide(sld As slide, content As String)
    Dim contentShape As Shape
    
    ' 找到內容文字框
    If sld.Shapes.Count >= 2 Then
        Set contentShape = sld.Shapes(2)
        
        With contentShape.TextFrame.TextRange
            If .Text = "" Then
                .Text = content
            Else
                .Text = .Text & vbCrLf & content
            End If
            
            .Font.Size = 18
            .Font.Name = "微軟正黑體"
            .ParagraphFormat.Bullet.Visible = msoTrue
            .ParagraphFormat.Bullet.Type = ppBulletUnnumbered
        End With
    End If
End Sub

' 取得文字輸入（從剪貼簿）
Function GetTextInput() As String
    Dim dataObj As Object
    Dim textContent As String
    
    On Error Resume Next
    
    ' 建立 DataObject 來存取剪貼簿
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.GetFromClipboard
    
    ' 取得文字
    textContent = dataObj.GetText
    
    If Err.Number <> 0 Then
        MsgBox "無法讀取剪貼簿！請確保剪貼簿中有文字內容。", vbExclamation
        textContent = ""
    End If
    
    Set dataObj = Nothing
    GetTextInput = textContent
End Function

' ========================================
' 進階版本：從檔案讀取文字
' ========================================
Sub ConvertTextFileToPPT()
    Dim filePath As String
    Dim fileNum As Integer
    Dim textContent As String
    Dim line As String
    
    ' 選擇檔案
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "選擇文字檔案")
    
    If filePath = "False" Then Exit Sub
    
    ' 讀取檔案
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        textContent = textContent & line & vbCrLf
    Loop
    
    Close #fileNum
    
    ' 處理文字（重複使用上面的邏輯）
    If textContent <> "" Then
        ' 暫存到剪貼簿
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText textContent
        dataObj.PutInClipboard
        Set dataObj = Nothing
        
        ' 執行轉換
        ConvertTextToPPT
    End If
End Sub
