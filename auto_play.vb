' This is the code that was converted to vbaProject.bin and gets auto added
' to each xlsm file. xlsxwriter can't programmatically create vba code but
' it does have a utility to extract vba macros from existing xlsm files and
' then reattach them to new generated xlsm files.
Private Sub Workbook_Open()
    ' Declare Current as a worksheet object variable.
    Dim Current As Worksheet
    
    Do
        ' Loop through all of the worksheets in the active workbook.
        For Each Current In Worksheets
        
           ' Insert your code here.
           ' This line displays the worksheet name in a message box.
           Current.Activate
           DoEvents
        Next
    Loop While True
End Sub