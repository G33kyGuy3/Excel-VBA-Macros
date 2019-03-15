# VBA Macros
ActiveX Controls
 
Learn how to create ActiveX controls such as command buttons, text boxes, list boxes etc. To create an ActiveX control in Excel VBA, execute the following steps.

1. On the Developer tab, click Insert.

2. For example, in the ActiveX Controls group, click Command Button to insert a command button control.

Create an ActiveX control in Excel VBA


 
3. Drag a command button on your worksheet.

4. Right click the command button (make sure Design Mode is selected).

5. Click View Code.

View Code

Note: you can change the caption and name of a control by right clicking on the control (make sure Design Mode is selected) and then clicking on Properties. Change the caption of the command button to 'Apply Blue Text Color'. For now, we will leave CommandButton1 as the name of the command button.

The Visual Basic Editor appears.

6. Add the code line shown below between Private Sub CommandButton1_Click() and End Sub.

Add Code Lines

7. Select the range B2:B4 and click the command button (make sure Design Mode is deselected).

Result:


Run Code
