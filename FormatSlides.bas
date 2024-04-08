Attribute VB_Name = "FormatSlides"

' Settings
Const g_FontSize As Integer = 32
Const g_FontName As String = "Nirmala UI"
' for telugu text "Nirmala UI"
' for english text "Calibri"
Const g_FontColor As Long = vbBlack
Const g_TextBoxWidthPercent As Single = 0.975


' ========================================================
' Apply a consistent format to the text in all slides
Sub FormatTextInSlides(activePres As Presentation)
Dim oSld As Slide
Dim oShp As Shape
Dim slideWidth As Single
Dim slideHeight As Single

For Each oSld In activePres.Slides
   For Each oShp In oSld.Shapes
        If oShp.HasTextFrame And oShp.TextFrame.HasText Then
            Call FormatText(oShp)
       End If
   Next oShp
Next oSld

End Sub 'FormatTextInSlides


' ========================================================
' Format text according to the settings
Sub FormatText(oShp As Object)
Dim oTxtRng As TextRange
Dim oTxtFont As font
Dim fontSize As Single
Dim shapeText As String
Dim shapeTextFormatted As String
Dim trimLine As String
Dim shapeTextLines() As String
Dim lineLength As Integer
Dim lineSizePixels As Integer
Dim lineCrossHalfWidth As Boolean

'distance of the Shape from the slide's border
oShp.Top = 10
oShp.Left = 10
oShp.Width = g_TextBoxWidthPercent * g_SlideWidth
        
'distance of the text from the shape's border
oShp.TextFrame.MarginLeft = 0
oShp.TextFrame.MarginTop = 0

Set oTxtRng = oShp.TextFrame.TextRange
oTxtRng.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft

Set oTxtFont = oTxtRng.font
oTxtFont.SIZE = g_FontSize
oTxtFont.Bold = msoFalse
oTxtFont.Italic = msoFalse
oTxtFont.Color = g_FontColor
oTxtFont.Name = g_FontName

fontSize = oTxtFont.SIZE
shapeText = oTxtRng.text
shapeTextLines = Split(shapeText, vbCr)
formattedTextLength = 0

For Each oLine In shapeTextLines
    trimLine = Trim(oLine)
    lineLength = Len(trimLine)
    If lineLength > 0 Then
        'TODO: calculate visible line size (width, height).
        'Use it to avoid text in the vertical or horizontal half of the Slide (for 4 TV split display)
        
        'lineCrossHalfWidth = (lineLength * fontSize > g_SlideWidth / 2)
        'lineSizePixels = GetLabelPixel(trimLine, oTxtFont.SIZE, oTxtFont.Name)
        'Debug.Print trimLine & " : " & lineLength & " : " & lineSizePixels
        
        shapeTextFormatted = shapeTextFormatted & trimLine & vbCr
    End If
Next

'remove last vbCr
shapeTextFormatted = Left(shapeTextFormatted, Len(shapeTextFormatted) - 1)
oTxtRng.text = Trim(shapeTextFormatted)

End Sub ' FormatText
