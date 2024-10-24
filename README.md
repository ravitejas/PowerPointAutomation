# PowerPointAutomation
VBA code to help create ppt files, add slides, format text, etc.

This project contains VBA code (.bas files) and examples which let a user easily create a PowerPoint presentation.
Demo: https://youtu.be/7WWRDV0Cxbc

Steps to use it:

1. Create a txt file with paragraphs of text (e.g., song lyrics).
2. Go to the examples folder and download the latest pptm file.
3. Right click on the pptm file, open Properties, and Unblock it.
4. Open it, go to the Developer tab, then open the Visual Basic Editor.
5. Enter the name of the txt file as the value of "g_PPTContentFileName"
6. Press F5 to run the Macro "MainCreatePPT"
7. Observe that new slides have been created, and text from the local text file is added to each slide.
8. A new pptx file is created with the same name as the txt file.

More info: 
1. Each paragraph (multiple lines) of text will be added to one Slide.
2. A blank line between paragraphs marks the transition to a new Slide.
3. Check each Module for Settings at the top of the file. Set values as per your preferences.
4. Run the procedure "MainCreatePPT".
5. This will read the .txt file specified in Settings, create slides, add text, and format text.
6. You can re-run it with different Settings or .txt files.

Update 4/13/2024:
For step 2 above: The number of paragraphs per Slide can be configured. 
Different font settings can be specified for each paragraph.

Update 5/3/2024:
Generated pptx will be saved as a new file in the same folder as the pptm.
pptx file name will be the same as that of the txt file.

Update 8/9/2024:
Added settings for Text Outline. 
By changing the color and weight of the outline, we can make the text more readable against backgrounds of different colors / shapes.

Update 10/23/2024:
Text box shape is filled with a transparent color, so text is more legible against the background image.
