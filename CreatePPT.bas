Attribute VB_Name = "CreatePPT"
'1 inch = 72 Points
'Point: unit of measurement in VBA code
'Slide size in inches: 10 (width), 5.625 (height). In Points: 720, 405

' Settings
Const g_SlideWidth As Integer = 720
Const g_SlideHeight As Integer = 405
Const g_SlideSeparatorTag As String = ""
Const g_PPTContentFile As String = "ppt_content_telugu_song.txt"


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
Call GenerateSlides(activePres, templateSlide)
Call FormatTextInSlides(activePres)

activePres.Save

End Sub ' MainCreatePPT


' ========================================================
' Use a template slide to create new slides
' Read text from a file and add it to the slides.
Sub GenerateSlides(activePres As Presentation, templateSlide As Slide)
Dim curSlide As Slide
Dim stringCompareResult As Integer
Dim startNewSlide As Boolean
Dim fullFilePath As String

fullFilePath = activePres.Path & "\" & g_PPTContentFile
' does the file exist?
If Len(Dir$(fullFilePath)) = 0 Then
    Debug.Print "ppt content file " & fullFilePath & " could not be found"
    Exit Sub
End If

Set curSlide = templateSlide.Duplicate()(1)
startNewSlide = True

Dim sFileContents As String
Call UnicodeTextReader.GetFileText(fullFilePath, "utf8", sFileContents)

Dim sLines() As String
Dim lineFromFile As Variant

sLines = Split(sFileContents, vbNewLine)

For Each lineFromFile In sLines
    'Debug.Print lineFromFile
    stringCompareResult = StrComp(g_SlideSeparatorTag, lineFromFile, vbTextCompare)
    If stringCompareResult = 0 Then
        Set curSlide = templateSlide.Duplicate()(1)
        curSlide.MoveTo (activePres.Slides.Count)
        startNewSlide = True
    End If
    
    For Each curShape In curSlide.Shapes
        If curShape.HasTextFrame Then
            If startNewSlide Then
                curShape.TextFrame.TextRange.text = ""
                startNewSlide = False
            End If
            curShape.TextFrame.TextRange.text = curShape.TextFrame.TextRange.text & lineFromFile & vbCr
        End If
    Next curShape
Next lineFromFile

End Sub ' GenerateSlides
