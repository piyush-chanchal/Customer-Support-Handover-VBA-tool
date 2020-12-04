

Sub Button4_Click()

Dim StatusCount, StatusStr, Statusiter, GroupName, GidIter, GroupId


ActiveWorkbook.Sheets("DoNotDelete").Visible = True

StatusCount = 0

If Sheets("Main").OLEObjects("CheckBox2").Object.Value = True Then StatusCount = StatusCount + 1
If Sheets("Main").OLEObjects("CheckBox1").Object.Value = True Then StatusCount = StatusCount + 1
If Sheets("Main").OLEObjects("CheckBox3").Object.Value = True Then StatusCount = StatusCount + 1
If Sheets("Main").OLEObjects("CheckBox4").Object.Value = True Then StatusCount = StatusCount + 1

If StatusCount = 0 Then
    MsgBox ("Status not selected")
    Exit Sub
End If
   

If StatusCount > 0 Then


    If Sheets("Main").OLEObjects("CheckBox1").Object.Value = True Then StatusStr = StatusStr + "," + "3"
    If Sheets("Main").OLEObjects("CheckBox2").Object.Value = True Then StatusStr = StatusStr + "," + "2"
    If Sheets("Main").OLEObjects("CheckBox3").Object.Value = True Then StatusStr = StatusStr + "," + "7"
    If Sheets("Main").OLEObjects("CheckBox4").Object.Value = True Then StatusStr = StatusStr + "," + "4"

    StatusStr = Mid(StatusStr, 2)
   
    StatusStr = "(status:" + Replace(StatusStr, ",", "%20OR%20status:") + ")"

End If


strurl = "https://" + CStr(ActiveWorkbook.Sheets("Settings").Cells(4, 2)) + ".freshdesk.com/api/v2/search/tickets?query=%22agent_id:" + CStr(ActiveWorkbook.Sheets("Settings").Cells(2, 2)) + "%20AND%20" + StatusStr + "%22"





AuthKey = EncodeBase64("" + ActiveWorkbook.Sheets("Settings").Cells(1, 2))

Dim m_sStatusCode, response

Set m_XMLhttp = CreateObject("MSXML2.XMLHTTP.6.0")
m_XMLhttp.Open "GET", strurl, False, AuthKey
m_XMLhttp.SetRequestHeader "Content-Type", "application/json"
m_XMLhttp.SetRequestHeader "Accept", "application/json"
m_XMLhttp.SetRequestHeader "Authorization", "Basic " & AuthKey
m_XMLhttp.send

m_sStatusCode = m_XMLhttp.Status
If m_sStatusCode = 200 Then
Else
    MsgBox ("Access Failed")
End If


response = m_XMLhttp.ResponseText


Dim NumTckts
NumTckts = CInt(Len(response) - Len(Replace(response, "," & Chr(34) & "id" & Chr(34) & ":", ""))) / 6


Dim StrInd(), EndInd(), SubStr(), TcktId()

ReDim StrInd(1 To NumTckts)
ReDim EndInd(1 To NumTckts)
ReDim SubStr(1 To NumTckts)
ReDim TcktId(1 To NumTckts)

StrInd(1) = InStr(response, "{" & Chr(34) & "cc_emails" & Chr(34) & ":")

Dim iter

If NumTckts > 1 Then
    For iter = 2 To NumTckts
        StrInd(iter) = StrInd(iter - 1) + InStr(Mid(response, StrInd(iter - 1) + 6), "{" & Chr(34) & "cc_emails" & Chr(34) & ":") + 5
        EndInd(iter - 1) = StrInd(iter) - 1
    Next iter
    EndInd(NumTckts) = Len(response)
End If

If NumTckts = 1 Then EndInd(NumTckts) = Len(response)


For iter = 1 To NumTckts
    SubStr(iter) = Mid(response, StrInd(iter), EndInd(iter))
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 1) = "Ticket Id"
    ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 1) = Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "id" + Chr(34) + ":") + 6, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "id" + Chr(34) + ":") + 6), ",") - 1)
 
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 2) = "Subject"
    ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 2) = Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "subject" + Chr(34) + ":" + Chr(34)) + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "subject" + Chr(34) + ":" + Chr(34)) + 12), Chr(34) + "," + Chr(34) + "association_type" + Chr(34) + ":") - 1)
   
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 3) = "Status"
    If CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10), ",") - 1)) = 3 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 3) = "Pending"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10), ",") - 1)) = 2 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 3) = "Open"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10), ",") - 1)) = 7 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 3) = "Waiting on Third Party"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "status" + Chr(34) + ":") + 10), ",") - 1)) = 4 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 3) = "Resolved"
    End If
   
   
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 4) = "Priority"
    If CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12), ",") - 1)) = 1 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 4) = "Low"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12), ",") - 1)) = 2 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 4) = "Medium"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12), ",") - 1)) = 3 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 4) = "High"
    ElseIf CInt(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "priority" + Chr(34) + ":") + 12), ",") - 1)) = 4 Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 4) = "Urgent"
    End If
   
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 5) = "Group"
 
    GroupId = Null
    GroupName = Null
 
    GroupId = Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "group_id" + Chr(34) + ":") + 12, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "group_id" + Chr(34) + ":") + 12), ",") - 1)

       
    For GidIter = 7 To 21
        If CLngLng(ActiveWorkbook.Sheets("Settings").Cells(GidIter, 1)) = CLngLng(GroupId) Then
            GroupName = ActiveWorkbook.Sheets("Settings").Cells(GidIter, 2)
        End If
    Next
   
   
    If GroupName <> "" Then
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 5) = "'" + CStr(GroupName)
    Else
        ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 5) = "'" + CStr(GroupId)
    End If
     
     
     
    ActiveWorkbook.Sheets("DoNotDelete").Cells(1, 6) = "Last update time"
    ActiveWorkbook.Sheets("DoNotDelete").Cells(iter + 1, 6) = Replace(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "updated_at" + Chr(34) + ":" + Chr(34)) + 15, InStr(Mid(SubStr(iter), InStr(SubStr(iter), "," + Chr(34) + "updated_at" + Chr(34) + ":" + Chr(34)) + 15), ",") - 6), "T", " ")
   
Next iter


Button3_Click

ActiveWorkbook.Sheets("DoNotDelete").Visible = False
 ThisWorkbook.Sheets("Main").Activate
 
   
End Sub




Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function