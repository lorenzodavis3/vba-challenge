I used the following resources to assist me with my assignment:
    - Looking up the last row of the sheet: LastrowC = Worksheets("control").Cells(Rows.Count, "A").End(xlUp).Row -  https://superuser.com/questions/1109754/vba-find-last-row-in-very-large-spreadsheet-overflow-error
    - Formatting numbers to percentages and decimals: Range("A1").NumberFormat = "0.000" - https://stackoverflow.com/questions/21919712/formatting-a-number-in-vba?rq=3
    - Looping through all the worksheets within the workbook: https://stackoverflow.com/questions/21919712/formatting-a-number-in-vba?rq=3
    - Check minimum and maximum values in a loop: https://stackoverflow.com/questions/52191966/finding-min-and-max-in-a-range-in-a-column-vba?rq=3