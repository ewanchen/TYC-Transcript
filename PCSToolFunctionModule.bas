Attribute VB_Name = "PCSToolFunctionModule"
Public returnvalue As TextBox

'********************************'
'Input: none                     '
'Output: 0 - fail to get pcs url '
'        1 - got the url         '
'            set gv.pcsurl       '
'********************************'
Public Function getPCSurl()
On Error GoTo ErrLine
Dim conn As ADODB.Connection
Dim rs As ADODB.recordSet
Dim query As String

Set conn = New ADODB.Connection
Set rs = New ADODB.recordSet

conn.Open GV.masterDSN

'*** Change to Production Mode ***'
'query = "select pcsurl from pcsinfo_xtrain"
query = "select * from pcsinfo"
Set rs = conn.Execute(query)

If rs.EOF = False Then
    GV.pcsurl = rs!pcsurl
    GV.partnumber_length = CLng(rs!part_number_length)
    GV.boxbarcode_length = CLng(rs!box_barcode_length)
    getPCSurl = 1
Else
    getPCSurl = 0
    'MsgBox ("Can't get PCS URL. Check with support")
End If

ErrLine:
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing
    If Err.Number <> 0 Then
        Exit Function
    End If
End Function

'******************************************************************'
'to authenticate user                                              '
'input: username, password, employee barcode, project id and phase '
'output: PASS or the error message                                 '
'******************************************************************'

Public Function authenticateUser(ByVal username As String, ByVal password As String, ByVal empBarcode As String, ByVal projectid As String, ByVal phase As String)
Dim strquery As String
Dim returnvalue As String

strquery = "function=AuthenticateUser"

If Len(username) <> 0 Then
    strquery = strquery + "&username=" + username + "&password=" + password
Else
    strquery = strquery + "&barcode=" + empBarcode
End If

strquery = strquery + "&phase=" + phase + "&project_id=" + CStr(CInt(projectid))

authenticateUser = callAPI(strquery)


End Function

'***************************************************'
'Input: PCS query                                   '
'Output: Return string from PCS                     '
'***************************************************'

Public Function callAPI(ByVal strquery As String)
'   Dim objHTTP As MSXML.XMLHTTPRequest
'   Dim strEnvelope As String
'   Dim strReturn As String
'   Dim objReturn As MSXML.DOMDocument
'
'    Set objHTTP = New MSXML.XMLHTTPRequest
'    Set objReturn = New MSXML.DOMDocument
 Dim objHTTP As MSXML2.XMLHTTP60
   
   Dim strEnvelope As String
   Dim strReturn As String
   'Dim objReturn As MSXML6.DOMDocument
   Dim objReturn As MSXML2.DOMDocument60

    Set objHTTP = New MSXML2.XMLHTTP60
    Set objReturn = New MSXML2.DOMDocument60

    
   'Set up to post to our local server
   objHTTP.Open "post", GV.pcsurl + "?" + strquery, False
    
   'Set a standard SOAP/ XML header for the content-type
   objHTTP.setRequestHeader "Content-Type", "text/xml"

'   'Make the SOAP call
   objHTTP.send strquery

   'Get the return envelope
   strReturn = objHTTP.responseText
   callAPI = trimString(strReturn)

    Set objHTTP = Nothing
    Set objReturn = Nothing
End Function

'************************************'
'input: string that needs to be trim
'output: string that has been filtered '
'**************************************'

Public Function trimString(ByVal txt As String)
Dim result As String
Dim allowChar As String
Dim regExp, match, i, spec

allowChar = "[A-Z]|[0-9]|[a-z]|[ /,^;:\\|.()!@#_-]"

    txt = UCase$(txt)
    For i = 1 To Len(txt)
        spec = Mid$(txt, i, 1)
        Set regExp = New regExp
        regExp.Global = True
        regExp.IgnoreCase = False
        regExp.Pattern = allowChar

        Set match = regExp.Execute(spec)

        If match.count <> 0 Then
            result = result + spec
        End If

Set regExp = Nothing
Next

trimString = result
End Function

'**************************************'
'input: none                           '
'output: put list of project in listbox'
'**************************************'

Public Function getProjectList()
Dim strquery As String
Dim returnvalue As String
Dim projectName
Dim count As Long

strquery = "function=GetProjectList"

returnvalue = callAPI(strquery)
projectName = Split(returnvalue, "^")

count = 0

While count <= UBound(projectName)
    FormLogin.List1.AddItem projectName(count)
    count = count + 1
Wend

If FormLogin.List1.ListCount = 0 Then
    frmmain.RichTextBox1.Text = frmmain.RichTextBox1.Text + "ERROR: No project found" + vbCrLf
End If

End Function

'***********************************'
'input: box id
'output: <STATUS_NAME> or FAIL
'***********************************'
Public Function getBoxStatus(ByVal boxid As Long)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetBoxStatus&box_id=" + CStr(boxid)

returnvalue = trimString(callAPI(strquery))

If Len(returnvalue) <> 0 And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    getBoxStatus = returnvalue
Else
    getBoxStatus = "FAIL"
End If

End Function

'*************************************'
'input: box id and part number
'output: <STATUS_NAME> or FAIL
'*************************************'
Public Function getPartStatus(ByVal boxid As Long, ByVal partnumber As String)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetPartStatus&box_id=" + CStr(boxid) + "&part=" + CStr(CInt(partnumber))

returnvalue = trimString(callAPI(strquery))

If Len(returnvalue) <> 0 And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    getPartStatus = returnvalue
Else
    getPartStatus = "FAIL"
End If

End Function

'**********************************'
'input: project id
'output: project name
'**********************************'

Public Function getProjectName(ByVal projectid As Long)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetProjectName&project_id=" + CStr(projectid)

returnvalue = trimString(callAPI(strquery))

If Len(returnvalue) <> 0 And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    getProjectName = returnvalue
Else
    getProjectName = "FAIL"
End If

End Function

'************************************'
'input: project name
'output: project id
'************************************'

Public Function getProjectId(ByVal projectName As String)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetProjectID&project_name=" + projectName

returnvalue = trimString(callAPI(strquery))

If Len(returnvalue) <> 0 And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    getProjectId = returnvalue
Else
    getProjectId = "FAIL"
End If
End Function



'*************************************'
'input: phase and project id
'output: list of boxid, boxnumber and partnumber (eg 1001 means box number 1 and part number 001) put in formdialog.lstboxpart
'***************************************'
Public Function getPartsForPhase(ByVal phase As String, ByVal project_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata
Dim boxnumber As String

strquery = "function=GetPartsForPhase&phase=" + phase + "&project_id=" + CStr(project_id)

returnvalue = trimString(callAPI(strquery))
ddata = Split(returnvalue, "^")

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    If StrComp(GV.job, "View") <> 0 And StrComp(GV.job, "Process Individual") <> 0 Then
        While count <= UBound(ddata)
            If Len(ddata(count)) <> 0 Then
                bdata = Split(ddata(count), ":")
'                If InStr(boxnumber, bdata(1)) = 0 Then
                    FormDialog.List1.AddItem getBoxBarcode(bdata(0)) + "-" + bdata(2)
'                    boxnumber = boxnumber + "^" + bdata(1)
'                End If
            End If
            count = count + 1
        Wend
    Else
       While count <= UBound(ddata)
            If Len(ddata(count)) <> 0 Then
                bdata = Split(ddata(count), ":")
                FormDialog.List1.AddItem getBoxBarcode(bdata(0)) + "-" + bdata(2)
            End If
            count = count + 1
        Wend
    End If
Else
    getPartsForPhase = "FAIL"
End If

End Function

'*****************************'
'input: project id
'output: list of boxnumber and partnumber add to formdialog.lstboxpart
'******************************'

Public Function getPartsForproject(ByVal project_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

strquery = "function=GetPartsForProject&project_id=" + CStr(project_id)
returnvalue = trimString(callAPI(strquery))
ddata = Split(returnvalue, "^")

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    While count <= UBound(ddata)
        If Len(ddata(count)) <> 0 Then
            bdata = Split(ddata(count), ":")
            FormDialog.List1.AddItem bdata(1) + AddZero(bdata(2), 3) + "-" + bdata(0)
        End If
        count = count + 1
    Wend
Else
    getPartsForproject = "FAIL"
End If

End Function

'*******************************************************'
' input: box id, part number, user id, box id           '
' output: clock in the user and check out box part      '
'*******************************************************'
Public Function clockUserIn(ByVal box_id As Long, ByVal partnumber As String, ByVal user_id As Long, ByVal phase As String, ByVal trainer_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

If trainer_id = 0 Then
    strquery = "function=ClockUserIn&box_id=" + CStr(box_id) + "&part=" + CStr(CInt(partnumber)) + "&user_id=" + CStr(user_id) + "&phase=" + phase
Else
    strquery = "function=ClockUserIn&box_id=" + CStr(box_id) + "&part=" + CStr(CInt(partnumber)) + "&user_id=" + CStr(user_id) + "&phase=" + phase + "&trainer_id=" + CStr(trainer_id)
End If
returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    clockUserIn = "PASS"
Else
    clockUserIn = "FAIL"
End If
End Function


'*******************************************************'
' input: box id, part number, user id, complete         '
' output: clock out the user and check in box part      '
'*******************************************************'
Public Function clockUserOut(ByVal box_id As Long, ByVal partnumber As String, ByVal user_id As Long, ByVal state As String, ByVal trainer_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

If trainer_id = 0 Then
    strquery = "function=ClockUserOut&box_id=" + CStr(box_id) + "&part=" + CStr(CInt(partnumber)) + "&user_id=" + CStr(user_id) + "&state=" + state
Else
    strquery = "function=ClockUserOut&box_id=" + CStr(box_id) + "&part=" + CStr(CInt(partnumber)) + "&user_id=" + CStr(user_id) + "&state=" + state + "&trainer_id=" + CStr(trainer_id)
End If

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    clockUserOut = "PASS"
Else
    clockUserOut = "FAIL"
End If
End Function

'*******************************************************'
' input: box id, part number, status, user id, reason   '
' output: Return pass/fail after update box part status '
'*******************************************************'
Public Function UpdatePartStatus(ByVal box_id As Long, ByVal partnumber As String, ByVal Status As String, ByVal user_id As Long, ByVal reason As String)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

strquery = "function=UpdatePartStatus&box_id=" + CStr(box_id) + "&part=" + CStr(CInt(partnumber)) + "&status=" + Status + "&user_id=" + CStr(user_id) + "&reason=" + reason
returnvalue = trimString(callAPI(strquery))

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    UpdatePartStatus = "PASS"
Else
    UpdatePartStatus = "FAIL"
End If
End Function

'**************************************'
' input: project id, setting           '
' output: Return setting value/fail    '
'**************************************'
Public Function GetProjectSetting(ByVal project_id As Long, ByVal setting_value As String)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

strquery = "function=GetProjectSetting&project_id=" + CStr(project_id) + "&setting=" + setting_value
returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    GetProjectSetting = returnvalue
Else
    GetProjectSetting = "FAIL"
End If
End Function

'************************************'
'input: username
'output: user id
'************************************'

Public Function getUserId(ByVal username As String)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetUserID&username=" + username

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    getUserId = returnvalue
Else
    getUserId = "FAIL"
End If
End Function

'************************************'
'input: project id and box number or box barcode
'output: box id
'************************************'

Public Function getBoxId(ByVal boxbarcode As String, ByVal projectid As Long, ByVal boxnumber As Long)
Dim strquery As String
Dim returnvalue As String


If Len(boxbarcode) <> 0 Then
    strquery = "function=GetBoxID&barcode=" + boxbarcode
Else
    strquery = "function=GetBoxID&project_id=" + CStr(projectid) + "&box=" + CStr(boxnumber)
End If

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    getBoxId = returnvalue
Else
    getBoxId = "FAIL"
End If

End Function

'**************************************'
'input: boxid                          '
'output: list of parts and sections    '
'**************************************'
Public Function getPartsSectionsForBox(ByVal boxid As Long)
Dim strquery As String
Dim returnvalue As String
Dim temp
Dim parts
Dim result As String
Dim count As Long

strquery = "function=GetPartsForBox&box_id=" + CStr(boxid)

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    getPartsSectionsForBox = returnvalue
Else
    getPartsSectionsForBox = "FAIL"
End If
End Function

'***********************************************'
'input: boxid, part,phase and counts            '
'       Scanning                                '
'       count_1=Total number of scanned pages.  '
'                                               '
'       Image Sampling                          '
'       count_1=Total number of scanned pages.  '
'       count_2=Total number of failed pages.   '
'                                               '
'       Indexing                                '
'       count_1=Total number of indexed pages.  '
'       count_2=Total number of indexed fields. '
'                                               '
'       Index Sampling                          '
'       count_1=Total number of indexed pages.  '
'       count_2=Total number of failed fields.  '
'       count_3=Total number of indexed fields. '
'                                               '
'output: PASS/FAIL                              '
'***********************************************'

Public Function RecordPartCounts(ByVal boxid As Long, ByVal part As String, ByVal phase As String, ByVal count1 As String, ByVal count2 As String, ByVal count3 As String)
Dim strquery As String
Dim returnvalue As String
Dim temp
Dim parts
Dim result As String
Dim count As Long

strquery = "function=RecordPartCounts&box_id=" + CStr(boxid) + "&part=" + CStr(CInt(part)) + "&phase=" + phase

If phase = "SCANNING" Then
    strquery = strquery + "&count_1=" + count1
ElseIf phase = "IMAGE_SAMPLING" Or phase = "INDEXING" Then
    strquery = strquery + "&count_1=" + count1 + "&count_2=" + count2
ElseIf phase = "INDEX_SAMPLING" Then
    strquery = strquery + "&count_1=" + count1 + "&count_2=" + count2 + "&count_3=" + count3
End If

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    RecordPartCounts = resultvalue
Else
    RecordPartCounts = "FAIL"
End If
End Function


'*************************************'
' input: box_id                       '
' output: Client Ref ID               '
'*************************************'
Public Function GetBoxClientRefID(ByVal boxid As Long)
Dim strquery As String
Dim returnvalue As String
Dim result As String

strquery = "function=GetBoxClientRef&box_id=" + CStr(boxid)

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    GetBoxClientRefID = returnvalue
Else
    GetBoxClientRefID = "FAIL"
End If
End Function

'*******************************************************'
' input: box id, part number, user id, box id           '
' output: clock in the user and check out box part      '
'*******************************************************'
Public Function clockUserInBox(ByVal box_id As Long, ByVal user_id As Long, ByVal phase As String, ByVal trainer_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

If trainer_id = 0 Then
    strquery = "function=ClockUserIn&box_id=" + CStr(box_id) + "&user_id=" + CStr(user_id) + "&phase=" + phase
Else
    strquery = "function=ClockUserIn&box_id=" + CStr(box_id) + "&user_id=" + CStr(user_id) + "&phase=" + phase + "&trainer_id=" + CStr(trainer_id)
End If

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid(returnvalue, 1, 4), "FAIL") <> 0 Then
    clockUserInBox = "PASS"
Else
    clockUserInBox = "FAIL"
End If
End Function

'*******************************************************'
' input: box id, part number, user id, complete         '
' output: clock out the user and check in box part      '
'*******************************************************'
Public Function clockUserOutBox(ByVal box_id As Long, ByVal user_id As Long, ByVal state As String, ByVal trainer_id As Long)
Dim strquery As String
Dim returnvalue As String
Dim count As Long
Dim ddata
Dim bdata

If trainer_id = 0 Then
    strquery = "function=ClockUserOut&box_id=" + CStr(box_id) + "&user_id=" + CStr(user_id) + "&state=" + state
Else
    strquery = "function=ClockUserOut&box_id=" + CStr(box_id) + "&user_id=" + CStr(user_id) + "&state=" + state + "&trainer_id=" + CStr(trainer_id)
End If

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And returnvalue <> "FAIL" And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    clockUserOutBox = "PASS"
Else
    clockUserOutBox = "FAIL"
End If
End Function

'************************************'
'input: box id
'output: box barcode
'************************************'

Public Function getBoxBarcode(ByVal box_id As String)
Dim strquery As String
Dim returnvalue As String

strquery = "function=GetBoxBarcode&box_id=" + box_id

returnvalue = callAPI(strquery)

If Len(returnvalue) <> 0 And StrComp(Mid$(returnvalue, 1, 4), "FAIL") <> 0 Then
    getBoxBarcode = returnvalue
Else
    getBoxBarcode = "FAIL"
End If
End Function
