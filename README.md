# PowerPointAutomation
VBA code to help create ppt files, add slides, format text, etc.

This project contains VBA code (.bas files) which let a user easily create a PowerPoint presentation.
Demo: https://youtu.be/7WWRDV0Cxbc

Steps to use it:

Quick version:
1. Go to the examples/ folder and open the pptm file.
2. Go to the Developer tab, then open the Visual Basic Editor.
3. Press F5 to run the Macro "MainCreatePPT"
4. Observe that new slides have been created, and text from the local text file is added to each slide.

Detailed version:
1. Create a .txt file in UTF-8 encoding, with the content of your PPT.
2. Each paragraph (multiple lines) of text will be added to one Slide.
3. A blank line between paragraphs marks the transition to a new Slide.
4. Create a .pptm file and add a template slide: e.g. with a picture background and a text box with some text.
5. Open the Visual Basic Editor and create modules (Import files) using the .bas files.
6. The entry point to the automation is in the Module "CreatePPT", procedure MainCreatePPT()
7. Check each Module for Settings at the top of the file. Set values as per your preferences.
8. Run the procedure "MainCreatePPT".
9. This will read the .txt file specified in Settings, create slides, add text, and format text.
10. You can re-run it with different Settings or .txt files.
11. When you are done, save the .pptm file as a .pptx file : i.e., a regular PowerPoint file without the VBA code.
12. Sample .pptm and .txt files are located in examples/ folder.

Update 4/13/2024:
For step 2 above: The number of paragraphs per Slide can be configured. 
Different font settings can be specified for each paragraph.

Update 5/3/2024:
Generated pptx will be saved as a new file in the same folder as the pptm.
pptx file name will be the same as that of the txt file.

Update 8/9/2024:
Added settings for Text Outline. 
By changing the color and weight of the outline, we can make the text more readable against backgrounds of different colors / shapes.
