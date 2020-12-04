

Function AddKeyCol(length)

Dim StatusRowNum, key

ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 9) = "Key"

If ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 1) = "Status" Then StatusRowNum = 1
If ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 2) = "Status" Then StatusRowNum = 2
If ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 3) = "Status" Then StatusRowNum = 3



For iter = 2 To length
    If ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, StatusRowNum) = "Open" Then ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, 9) = 1
    If ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, StatusRowNum) = "Pending" Then ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, 9) = 2
    If ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, StatusRowNum) = "Waiting on Third Party" Then ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, 9) = 3
    If ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, StatusRowNum) = "Resolved" Then ActiveWorkbook.Sheets("DoNotDelete").Cells(iter, 9) = 4
Next



End Function



Function Sorting(length)


''''Insertion Sort

Dim Temp, var1
For iter = 3 To length
    var1 = iter
    Do While var1 > 2 And (ActiveWorkbook.Sheets("DoNotDelete").Cells(var1, 9) < ActiveWorkbook.Sheets("DoNotDelete").Cells(var1 - 1, 9))
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(var1)) + ":I" + Trim(Str(var1))).Copy Destination:=ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(length + 10)) + ":I" + Trim(Str(length + 10)))
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(var1 - 1)) + ":I" + Trim(Str(var1 - 1))).Copy Destination:=ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(var1)) + ":I" + Trim(Str(var1)))
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(length + 10)) + ":I" + Trim(Str(length + 10))).Copy Destination:=ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(var1 - 1)) + ":I" + Trim(Str(var1 - 1)))
        var1 = var1 - 1
    Loop
 Next
 
 ''''''End of insertion sort block
 
 ''Delete helper row data
 ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(length + 10)) + ":I" + Trim(Str(length + 10))).Delete
 ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(length + 10)) + ":I" + Trim(Str(length + 10))).Borders.LineStyle = xlNone


End Function



Function coloring(length)


''''Top row coloring
ActiveWorkbook.Sheets("DoNotDelete").range("A:H").Interior.Color = xlNone
ActiveWorkbook.Sheets("DoNotDelete").range("A1:H1").Interior.Color = RGB(35, 58, 125)
ActiveWorkbook.Sheets("DoNotDelete").range("A1:H1").Font.Color = RGB(255, 255, 255)


''''Coloring of rows other than top row
'ActiveWorkbook.Sheets("DoNotDelete").range("A2:H2").Interior.Color = RGB(255, 179, 179)
'ActiveWorkbook.Sheets("DoNotDelete").range("A2:H2").Font.Color = RGB(255, 255, 255)


For iter = 2 To length
    If ActiveWorkbook.Sheets("DoNotDelete").range("I" + Trim(Str(iter))) = 1 Then
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(iter)) + ":H" + Trim(Str(iter))).Interior.Color = RGB(255, 102, 102)
    End If
    If ActiveWorkbook.Sheets("DoNotDelete").range("I" + Trim(Str(iter))) = 2 Then
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(iter)) + ":H" + Trim(Str(iter))).Interior.Color = RGB(230, 255, 230)
    End If
    If ActiveWorkbook.Sheets("DoNotDelete").range("I" + Trim(Str(iter))) = 3 Then
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(iter)) + ":H" + Trim(Str(iter))).Interior.Color = RGB(230, 230, 255)
    End If
    If ActiveWorkbook.Sheets("DoNotDelete").range("I" + Trim(Str(iter))) = 4 Then
        ActiveWorkbook.Sheets("DoNotDelete").range("A" + Trim(Str(iter)) + ":H" + Trim(Str(iter))).Interior.Color = RGB(242, 242, 242)
    End If
 Next


End Function


Function DelKeyCol()


ActiveWorkbook.Sheets("DoNotDelete").Columns(9).EntireColumn.Delete


End Function
