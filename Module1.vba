


Sub Button3_Click()

    Dim length, iter, FullRange
       
    Sheets("DoNotDelete").Rows(1).Columns("G") = "Comments"
    Sheets("DoNotDelete").Rows(1).Columns("H") = "Assigned To"
   
   
    length = RowsLen()
   
    For iter = 2 To length
        Sheets("DoNotDelete").Rows(iter).Columns("H") = "Team"
        If Sheets("Main").OLEObjects("CheckBox13").Object.Value = True Then
            If Sheets("DoNotDelete").Rows(iter).Columns("C") = "Pending" Then Sheets("DoNotDelete").Rows(iter).Columns("G") = "" + ActiveWorkbook.Sheets("Main").range("B24")
            If Sheets("DoNotDelete").Rows(iter).Columns("C") = "Open" Then Sheets("DoNotDelete").Rows(iter).Columns("G") = "" + ActiveWorkbook.Sheets("Main").range("B25")
            If Sheets("DoNotDelete").Rows(iter).Columns("C") = "Waiting on Third Party" Then Sheets("DoNotDelete").Rows(iter).Columns("G") = "" + ActiveWorkbook.Sheets("Main").range("B26")
            If Sheets("DoNotDelete").Rows(iter).Columns("C") = "Resolved" Then Sheets("DoNotDelete").Rows(iter).Columns("G") = "" + ActiveWorkbook.Sheets("Main").range("B27")
        End If
    Next
   
    Sheets("DoNotDelete").Rows(1).Font.Bold = True
   
   
    Sheets("DoNotDelete").Columns("A").AutoFit
    Sheets("DoNotDelete").Columns("B").ColumnWidth = 50
    Sheets("DoNotDelete").Columns("C").AutoFit
    Sheets("DoNotDelete").Columns("D").AutoFit
    Sheets("DoNotDelete").Columns("E").AutoFit
    Sheets("DoNotDelete").Columns("F").AutoFit
    Sheets("DoNotDelete").Columns("G").ColumnWidth = 40
    Sheets("DoNotDelete").Columns("H").AutoFit
   
    FullRange = "A1:H" + Mid(Str(length), 2)
   
    Sheets("DoNotDelete").range(FullRange).WrapText = True

    With Sheets("DoNotDelete").range(FullRange).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
    End With
   
   
    AddKeyCol (length)
    Sorting (length)
    coloring (length)
    DelKeyCol
   
    If Sheets("Main").OLEObjects("CheckBox5").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("A:A").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox6").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("B:B").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox7").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("C:C").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox8").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("D:D").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox9").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("E:E").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox10").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("F:F").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox11").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("G:G").EntireColumn.Delete
    If Sheets("Main").OLEObjects("CheckBox12").Object.Value = False Then ActiveWorkbook.Sheets("Donotdelete").range("H:H").EntireColumn.Delete

   
    OutlookEmail (length)
   
    Sheets("DoNotDelete").Cells.ClearContents
   
    Sheets("DoNotDelete").range(FullRange).Borders.LineStyle = xlNone
   
End Sub




Function RowsLen()

    RowsLen = 0
    Do While Not Sheets("DoNotDelete").Rows(RowsLen + 1).Columns("A") = ""
        RowsLen = RowsLen + 1
    Loop
   
End Function
Function OutlookEmail(length)
    Dim oOutApp As Object, oOutMail As Object
    Dim strbody As String, FixedHtmlBody As String
    Dim Ret, usrname As String

    usrname = Environ("Username")

    Ret = "C:\Users\" + usrname + "\AppData\Roaming\Microsoft\Signatures\" & Trim(CStr(ActiveWorkbook.Sheets("Settings").range("B3")))

   
    If Ret = False Then Exit Function

    FixedHtmlBody = FixHtmlBody(Ret)

    Set oOutApp = CreateObject("Outlook.Application")
    Set oOutMail = oOutApp.CreateItem(0)
   
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
   
    regex.Pattern = "\n"
    regex.Global = True

    strbody = "" + regex.Replace(ActiveWorkbook.Sheets("Main").range("B23"), "<BR>") & fncRangeToHtml("DoNotDelete", "A1:H" + CStr(length + 1)) & FixedHtmlBody

   
    On Error Resume Next
    With oOutMail
        .To = "" + ActiveWorkbook.Sheets("Main").range("B21")
        .CC = "" + ActiveWorkbook.Sheets("Main").range("B22")
        .BCC = ""
        .Subject = "" + Replace(ActiveWorkbook.Sheets("Main").range("B28"), "{today}", Format(Date, "dd/mm/yyyy"))
        .HTMLBody = .HTMLBody & strbody
        .Display
        If Sheets("Main").OLEObjects("CheckBox14").Object.Value = True Then .send
    End With
    On Error GoTo 0

    Set oOutMail = Nothing
    Set oOutApp = Nothing
End Function

Function FixHtmlBody(r As Variant) As String
    Dim FullPath As String, filename As String
    Dim FilenameWithoutExtn As String
    Dim foldername As String
    Dim MyData As String

    Open r For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1

    filename = GetFilenameFromPath(r)
    FilenameWithoutExtn = Left(filename, (InStrRev(filename, ".", -1, vbTextCompare) - 1))
    foldername = FilenameWithoutExtn & "_files"

    FullPath = Left(r, InStrRev(r, "\")) & foldername

      FullPath = Replace(FullPath, " ", "%20")

    FixHtmlBody = Replace(MyData, foldername, FullPath)
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then _
    GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
End Function


Function fncRangeToHtml( _
    strWorksheetName As String, _
    strRangeAddress As String) As String
   
    Dim objFilesytem As Object, objTextstream As Object, objShape As Shape
    Dim strFilename As String, strTempText As String
    Dim blnRangeContainsShapes As Boolean
   
    strFilename = Environ$("temp") & "\" & _
        Format(Now, "dd-mm-yy_h-mm-ss") & ".htm"
       
       
    ThisWorkbook.Sheets("DoNotDelete").Activate


    ThisWorkbook.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        filename:=strFilename, _
        Sheet:=strWorksheetName, _
        Source:=strRangeAddress, _
        HtmlType:=xlHtmlStatic).Publish True
       
    Set objFilesytem = CreateObject("Scripting.FileSystemObject")
    Set objTextstream = objFilesytem.GetFile(strFilename).OpenAsTextStream(1, -2)
    strTempText = objTextstream.ReadAll
    objTextstream.Close
   
    For Each objShape In Worksheets(strWorksheetName).Shapes
        If Not Intersect(objShape.TopLeftCell, Worksheets( _
            strWorksheetName).range(strRangeAddress)) Is Nothing Then
           
            blnRangeContainsShapes = True
            Exit For
           
        End If
    Next
   
    If blnRangeContainsShapes Then _
        strTempText = fncConvertPictureToMail(strTempText, Worksheets(strWorksheetName))
   
    fncRangeToHtml = strTempText
    fncRangeToHtml = Replace(fncRangeToHtml, " align=center", " align=left")


    Set objTextstream = Nothing
    Set objFilesytem = Nothing
   
    Kill strFilename
   
   
End Function

Function fncConvertPictureToMail(strTempText As String, objWorksheet As Worksheet) As String
    Const HTM_START = "") - lngPathLeft)
    strTemp = Replace(strTemp, HTM_START & Chr$(34), "")
    strTemp = Replace(strTemp, HTM_END & Chr$(34), "")
    strTemp = strTemp & "/"
   
    strTempText = Replace(strTempText, strTemp, Environ$("temp") & "" & strTemp)
    fncConvertPictureToMail = strTempText
   
End Function