Attribute VB_Name = "FormatSlides"
Public Type ParagraphInfo
    FontSize As Integer
    FontName As String
    FontColor As Long
End Type

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
' Format text according to requirements
Sub FormatText(oShp As Object)
Dim oTxtRng As TextRange
Dim FontSize As Single
Dim shapeText As String
Dim shapeTextFormatted As String
Dim trimLine As String
Dim shapeTextLines() As String
Dim lineLength As Integer
Dim lineSizePixels As Integer
Dim lineCrossHalfWidth As Boolean

Set oTxtRng = oShp.TextFrame.TextRange
FontSize = oTxtRng.font.Size

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
