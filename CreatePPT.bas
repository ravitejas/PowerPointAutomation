Attribute VB_Name = "CreatePPT"
'1 inch = 72 Points
'Point: unit of measurement in VBA code
'Slide size in inches: 10 (width), 5.625 (height). In Points: 720, 405

' Settings
Const g_PPTContentFileName As String = "What can wash away"
Const g_TextFileFormat As String = ".txt"
Const g_ParagraphsPerSlide As Integer = 2
Const g_TextBoxLeftDistance As Integer = 10
Const g_TextBoxTopDistance As Integer = 10
Const g_TextBoxWidthPercent As Single = 0.975
Const g_SlideWidth As Integer = 720
Const g_SlideHeight As Integer = 405
Const g_ParagraphSeparatorTag As String = "" ' A blank line marks a new paragraph

' When 1 slide has 2 paragraphs (e.g. telugu and english)
' Allow different font properties for each para
Public Function GetParagraphInfos() As ParagraphInfo()
    Dim g_ParagraphInfos(1 To 2) As ParagraphInfo
    g_ParagraphInfos(1).FontName = "Calibri"
    g_ParagraphInfos(1).FontSize = 34
    g_ParagraphInfos(1).FontColor = vbBlack
    
    g_ParagraphInfos(2).FontName = "Nirmala UI"
    g_ParagraphInfos(2).FontSize = 44
    g_ParagraphInfos(2).FontColor = vbBlack

    GetParagraphInfos = g_ParagraphInfos
End Function


' ========================================================
' Main function (entry point)
' Create a presentation
Sub MainCreatePPT()
Dim activePres As Presentation
Dim templateSlide As Slide

Set activePres = Application.ActivePresentation
activePres.PageSetup.slideWidth = g_SlideWidth
activePres.PageSetup.slideHeight = g_SlideHeight
If activePres.Slides.Count < 1 Then
    Debug.Print "We need 1 slide in the initial PPT, as a template to create the remaining slides"
    Exit Sub
End If

' Delete other slides, so we can regenerate them
Do While activePres.Slides.Count > 1
    activePres.Slides(2).Delete
Loop

Set templateSlide = activePres.Slides(1)
Call CreateTextBoxes(templateSlide)
Call GenerateSlides(activePres, templateSlide)

Dim pptPath As String
pptPath = activePres.Path & "\" & g_PPTContentFileName
activePres.SaveCopyAs pptPath, ppSaveAsOpenXMLPresentation

End Sub ' MainCreatePPT


' ========================================================
' Create required number of TextBoxes on a slide, and set their properties
Sub CreateTextBoxes(templateSlide As Slide)

' Delete existing textboxes
With templateSlide.Shapes
    For intShape = .Count To 1 Step -1
        With .Item(intShape)
            If .Type = msoTextBox Then .Delete
        End With
    Next
End With

' Each paragraph goes into one TextBox.
Dim textBoxShape As Shape
Dim textBoxNumber As Integer
Dim textBoxTopPos As Integer
Dim textBoxHeight As Single
Dim oTxtRng As TextRange
Dim oTxtFont As font
Dim paragraphInfos() As ParagraphInfo
paragraphInfos = GetParagraphInfos()

textBoxHeight = g_SlideHeight / g_ParagraphsPerSlide
For textBoxNumber = 1 To g_ParagraphsPerSlide
    textBoxTopPos = g_TextBoxTopDistance + (textBoxNumber - 1) * textBoxHeight
    Set textBoxShape = templateSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=g_TextBoxLeftDistance, Top:=textBoxTopPos, Width:=g_TextBoxWidthPercent * g_SlideWidth, Height:=textBoxHeight)
    textBoxShape.TextFrame.TextRange.text = ""
    
    'distance of the text from the shape's border
    textBoxShape.TextFrame.MarginLeft = 0
    textBoxShape.TextFrame.MarginTop = 0
    
    Set oTxtRng = textBoxShape.TextFrame.TextRange
    oTxtRng.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    Set oTxtFont = oTxtRng.font
    oTxtFont.Size = paragraphInfos(textBoxNumber).FontSize
    oTxtFont.Color = paragraphInfos(textBoxNumber).FontColor
    oTxtFont.Name = paragraphInfos(textBoxNumber).FontName
    oTxtFont.Bold = msoFalse
    oTxtFont.Italic = msoFalse
    
Next textBoxNumber

End Sub 'CreateTextBoxes


' ========================================================
' Use a template slide to create new slides
' Read text from a file and add it to the slides.
Sub GenerateSlides(activePres As Presentation, templateSlide As Slide)
Dim curSlide As Slide
Dim stringCompareResult As Integer
Dim paragraphNumberInSlide As Integer
Dim textBoxNumberInSlide As Integer
Dim fullFilePath As String

fullFilePath = activePres.Path & "\" & g_PPTContentFileName & g_TextFileFormat
' does the file exist?
If Len(Dir$(fullFilePath)) = 0 Then
    Debug.Print "ppt content file " & fullFilePath & " could not be found"
    Exit Sub
End If

Set curSlide = templateSlide.Duplicate()(1)
Dim sFileContents As String
Call UnicodeTextReader.GetFileText(fullFilePath, "utf8", sFileContents)

Dim sLines() As String
Dim lineFromFile As Variant

sLines = Split(sFileContents, vbNewLine)
paragraphNumberInSlide = 1

For Each lineFromFile In sLines
    'Debug.Print lineFromFile
    stringCompareResult = StrComp(g_ParagraphSeparatorTag, lineFromFile, vbTextCompare)
    If stringCompareResult = 0 Then
        paragraphNumberInSlide = paragraphNumberInSlide + 1
        If paragraphNumberInSlide > g_ParagraphsPerSlide Then
            Set curSlide = templateSlide.Duplicate()(1)
            curSlide.MoveTo (activePres.Slides.Count)
            paragraphNumberInSlide = 1
        End If
        
    End If
    
    textBoxNumberInSlide = 0
    For Each curShape In curSlide.Shapes
        If curShape.Type = msoTextBox Then
            textBoxNumberInSlide = textBoxNumberInSlide + 1
            If textBoxNumberInSlide = paragraphNumberInSlide Then
                curShape.TextFrame.TextRange.text = curShape.TextFrame.TextRange.text & lineFromFile & vbCr
                Exit For
            End If
        End If
    Next curShape
    
Next lineFromFile

' Add the PPT Title to the template (first) slide
For Each curShape In templateSlide.Shapes
    If curShape.Type = msoTextBox Then
        curShape.TextFrame.TextRange.text = g_PPTContentFileName
        Exit For
    End If
Next curShape

End Sub ' GenerateSlides
