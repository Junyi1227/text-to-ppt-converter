' ========================================
' Text to PowerPoint Converter - Mac 相容版本
' 文字轉 PowerPoint 轉換器 (Mac Compatible)
' ========================================
' 適用於：Mac 版 PowerPoint
' 限制：不支援剪貼簿讀取，必須從檔案讀取
' ========================================

Option Explicit

' 主程式：從檔案讀取文字並轉換為 PowerPoint 投影片
Sub ConvertTextFileToPPT()
    Dim filePath As String
    Dim fileNum As Integer
    Dim textContent As String
    Dim line As String
    Dim pres As Presentation
    Dim currentSlide As Slide
    Dim content As String
    
    ' 選擇檔案（Mac 相容寫法）
    filePath = MacScript("return POSIX path of (choose file with prompt ""選擇文字檔案"" of type {""public.plain-text""})")
    
    If filePath = "" Then Exit Sub
    
    ' 讀取檔案
    On Error GoTo ErrorHandler
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' 使用目前的簡報
    Set pres = ActivePresentation
    
    ' 逐行處理
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        line = Trim(line)
        
        ' 檢查標記
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
    Loop
    
    Close #fileNum
    
    MsgBox "完成！已建立 " & pres.Slides.Count & " 張投影片", vbInformation
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

' 建立主題投影片（## 標記）
Function CreateTitleSlide(pres As Presentation, title As String) As Slide
    Dim sld As Slide
    Dim titleShape As Shape
    
    ' 新增投影片
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutTitle)
    
    ' 設定標題
    On Error Resume Next
    Set titleShape = sld.Shapes(1)
    If Not titleShape Is Nothing Then
        With titleShape.TextFrame.TextRange
            .Text = title
            .Font.Size = 44
            .Font.Bold = msoTrue
            .Font.Name = "微軟正黑體"
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End If
    On Error GoTo 0
    
    ' 設定背景顏色
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.ForeColor.RGB = RGB(230, 240, 255)
    
    Set CreateTitleSlide = sld
End Function

' 建立內文投影片（# 標記）
Function CreateContentSlide(pres As Presentation, title As String) As Slide
    Dim sld As Slide
    
    ' 新增投影片
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutText)
    
    ' 設定標題
    On Error Resume Next
    If sld.Shapes.Count >= 1 Then
        With sld.Shapes(1).TextFrame.TextRange
            .Text = title
            .Font.Size = 32
            .Font.Bold = msoTrue
            .Font.Name = "微軟正黑體"
        End With
    End If
    
    ' 清空內容區域
    If sld.Shapes.Count >= 2 Then
        sld.Shapes(2).TextFrame.TextRange.Text = ""
    End If
    On Error GoTo 0
    
    ' 設定背景顏色
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.ForeColor.RGB = RGB(245, 245, 245)
    
    Set CreateContentSlide = sld
End Function

' 新增內容到投影片
Sub AddContentToSlide(sld As Slide, content As String)
    Dim contentShape As Shape
    
    On Error Resume Next
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
    On Error GoTo 0
End Sub

' ========================================
' 從投影片備註欄讀取文字（Mac 替代方案）
' ========================================
Sub ConvertTextFromNotes()
    Dim pres As Presentation
    Dim notesSlide As Slide
    Dim textContent As String
    Dim lines() As String
    Dim i As Integer
    Dim currentSlide As Slide
    Dim content As String
    
    Set pres = ActivePresentation
    
    ' 檢查是否有投影片
    If pres.Slides.Count = 0 Then
        MsgBox "請先建立一張投影片，並在備註欄中輸入您的文字內容", vbExclamation
        Exit Sub
    End If
    
    ' 從第一張投影片的備註欄讀取文字
    On Error Resume Next
    textContent = pres.Slides(1).NotesPage.Shapes(2).TextFrame.TextRange.Text
    On Error GoTo 0
    
    If Trim(textContent) = "" Then
        MsgBox "備註欄沒有內容！" & vbCrLf & vbCrLf & _
               "使用方式：" & vbCrLf & _
               "1. 在第一張投影片下方的備註欄貼上您的文字" & vbCrLf & _
               "2. 執行此巨集" & vbCrLf & _
               "3. 新的投影片將會自動建立", vbInformation
        Exit Sub
    End If
    
    ' 刪除第一張投影片（用來輸入文字的）
    Dim shouldDelete As VbMsgBoxResult
    shouldDelete = MsgBox("是否刪除第一張投影片（輸入用）？", vbYesNo + vbQuestion)
    
    ' 分割文字
    lines = Split(textContent, vbLf)  ' Mac 使用 vbLf
    
    ' 處理每一行
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(Replace(lines(i), vbCr, ""))  ' 移除可能的 CR
        
        If Left(line, 2) = "##" Then
            content = Trim(Mid(line, 3))
            Set currentSlide = CreateTitleSlide(pres, content)
        ElseIf Left(line, 1) = "#" Then
            content = Trim(Mid(line, 2))
            Set currentSlide = CreateContentSlide(pres, content)
        ElseIf line <> "" And Not currentSlide Is Nothing Then
            AddContentToSlide currentSlide, line
        End If
    Next i
    
    ' 刪除輸入用的投影片
    If shouldDelete = vbYes Then
        pres.Slides(1).Delete
    End If
    
    MsgBox "完成！已建立投影片", vbInformation
End Sub

' ========================================
' 從剪貼簿讀取（Mac 替代方案 - 需要手動執行）
' ========================================
Sub ConvertTextFromClipboard_Mac()
    MsgBox "Mac 版本不支援直接從剪貼簿讀取。" & vbCrLf & vbCrLf & _
           "請使用以下替代方案：" & vbCrLf & _
           "1. 【推薦】使用 ConvertTextFileToPPT 從檔案讀取" & vbCrLf & _
           "2. 使用 ConvertTextFromNotes 從備註欄讀取" & vbCrLf & _
           "3. 使用 Python 版本（跨平台）", vbInformation
End Sub
