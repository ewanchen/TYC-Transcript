Attribute VB_Name = "EJRulesModule"
Private Type SYSTEMTIME
   vYear As Integer
   vMonth As Integer
   vDayOfWeek As Integer
   vDay As Integer
   vHour As Integer
   vMinute As Integer
   vSecond As Integer
   vMilliseconds As Integer
End Type


Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'********************************'
'Error Checking for each document'
'return true/false               '
'********************************'

Public Function error_checking()
Dim valid_date As Integer
Dim doctype As String
Dim i As Integer

    

ConvertToUpper
If GV.curr_phase = "INDEXING" Then
    If check_format("Current Last Name", "/.: ;'&#()-", 1, 1, Main.txtLastName.Text, 0, True) = False Then
        error_checking = False
        MsgBox "Current Last name has invalid character!! Only allow alphanumeric and /.: ;'&#()-"
        Exit Function
    End If
    
    If check_format("First Name", "/.: ;'&#()-", 1, 1, Main.txtFirstName.Text, 0, True) = False Then
        error_checking = False
        MsgBox "Current first name has invalid character!! Only allow alphanumeric and /.: ;'&#()-"
        Exit Function
    End If
    
'    If Len(Main.txtDD.Text) <> 0 Or Len(Main.txtMM.Text) <> 0 Or Len(Main.txtYY.Text) <> 0 Then
'    If check_date("DOB", Main.txtDD.Text, Main.txtMM.Text, Main.txtYY.Text, "", Format(Date, "mm-dd-yyyy")) <> 0 Then
'        error_checking = False
'        Exit Function
'    End If
'    End If
Else
    If Main.optFail.Value = False And Main.optPass.Value = False Then
        error_checking = False
        MsgBox "Please choose PASS/FAIL!"
        Exit Function
    End If
    
    If Main.optFail.Value = True And Len(Main.txtFailNum.Text) = 0 Then
        error_checking = False
        MsgBox "Please enter number of fields that fails"
        Exit Function
    End If
    
    If Main.optFail.Value = True And Len(Main.txtFailNum.Text) > 0 Then
        If Main.txtFailNum.Text > 3 Then
            error_checking = False
            MsgBox "Total number of fail fields cannot be greater than 3. Last name, First name and DOB each counts as 1 field."
            Exit Function
        End If
    End If
End If
    

    error_checking = True
End Function

Public Sub resetFields()
    Main.txtLastName.Text = ""
    Main.txtFirstName.Text = ""
'    Main.txtDD.Text = ""
    Main.txtMM.Text = ""
'    Main.txtYY.Text = ""
    Main.txtTYC.Text = ""
    Main.chkBadScan.Value = 0
    Main.chkDelete.Value = 0
    Main.chkFirst.Value = 0
    If GV.curr_phase = "INDEX_SAMPLING" Then
        Main.optPass.Value = True
        Main.txtFailNum.Text = ""
    End If
End Sub

Public Sub NextImage_toDo()
    ConvertToUpper
    Dim vTime As SYSTEMTIME
    Dim action As String
    GV.end_date = Date
    GetLocalTime vTime
    GV.end_time = Format(vTime.vHour) & ":" & Format(vTime.vMinute) & ":" & Format(vTime.vSecond) & ":" & Format(vTime.vMilliseconds)

    If GV.preview_flag = 1 Then
        MsgBox "You still in previewing image state!! Please go back to last indexed image " + GV.lastindexpage + " and select 'Unlock indexing field' from File drop list or use the short cut!"
        Exit Sub
    End If
    
    
        If Main.chkDelete.Value = 0 Then
            If error_checking() = False Then
                Exit Sub
            End If
        End If
        
        
        
            '*** Recording Activity ***'
        If GV.skipped = False Then
            If GV.curr_phase = "INDEXING" Then
                If CheckChangeField(False) = True Then
                    If Main.chkFirst.Value = 0 And Main.LabelIndexed.Caption = "Not Indexed" Then
                        MsgBox "It's not first page, indexed fields are not supposed to change. Please check!"
                        Exit Sub
                    End If
                End If
                If CheckChangeField(True) = True Then
                    Activity_Log GV.filename, "Save and Commit"
                Else
                    Activity_Log GV.filename, "Save without Change"
                End If
            Else
                If Main.optFail.Value = True Then
                    Activity_Log GV.filename, "Fail"
                Else
                    Activity_Log GV.filename, "Pass"
                End If
            End If
        End If
        
        If GV.curr_phase = "INDEXING" Then
            DocTypeInsert GV.boxnumber, GV.boxpart, GV.filename
        Else
            PassFailInsert GV.boxbarcode, GV.boxpart, GV.filename
        End If
        
        If GV.finish_sample = 0 Then
            ViewNextImage 1
            FetchData GV.boxnumber, GV.boxpart, GV.filename
            Main.txtLastName.SetFocus
            Main.txtLastName.BackColor = GV.FocusBColor
            Main.txtLastName.SelStart = 0
            Main.txtLastName.SelLength = Len(Main.txtLastName.Text)
        Else
            SampCalculate GV.boxbarcode, GV.boxpart
        End If

        GV.start_date = Date
        GetLocalTime vTime
        GV.start_time = Format(vTime.vHour) & ":" & Format(vTime.vMinute) & ":" & Format(vTime.vSecond) & ":" & Format(vTime.vMilliseconds)
        pre_page = False
        GV.lastindexpage = GV.filename
        If GV.end_flag = 0 Then
'        Main.chkFirst.Value = 1
        Else
            EnableTextField False
            MsgBox "You have finished this box part. Please switch to another box part or exit the tool."
        End If
End Sub

Public Sub PreviousImage_toDo()
Dim vTime As SYSTEMTIME
    If GV.preview_flag = 1 Then
        MsgBox "You still in previewing image state!! Please go back to last indexed image " + GV.lastindexpage + " and select 'Unlock indexing field' from File drop list or use the short cut!"
        Exit Sub
    End If
    
    GV.end_date = Date
    GetLocalTime vTime
    GV.end_time = Format(vTime.vHour) & ":" & Format(vTime.vMinute) & ":" & Format(vTime.vSecond) & ":" & Format(vTime.vMilliseconds)
    
    Activity_Log GV.filename, "Reverse"
    ViewPreviousImage
    FetchData GV.boxnumber, GV.boxpart, GV.filename
    Main.txtLastName.SetFocus
    Main.txtLastName.BackColor = GV.FocusBColor
    Main.txtLastName.SelStart = 0
    Main.txtLastName.SelLength = Len(Main.txtLastName.Text)
    GV.start_date = Date
    GetLocalTime vTime
    GV.start_time = Format(vTime.vHour) & ":" & Format(vTime.vMinute) & ":" & Format(vTime.vSecond) & ":" & Format(vTime.vMilliseconds)
End Sub

''***************************************'
'' Insert page pass/fail in to box table '
''***************************************'
'Public Sub PassFailInsert(ByVal boxnumber As String, ByVal partnumber As String, ByVal filename As String)
'On Error GoTo ErrLine
'    Dim conn As ADODB.Connection
'    Dim rs As ADODB.recordSet
'    Dim query As String
'    Dim imageindexedcount As Long
'
'    Set conn = New ADODB.Connection
'    Set rs = New ADODB.recordSet
'    conn.Open GV.DSN
'
'    '**** PCS: Sampling rule: sample pass return 2****'
'    query = "select * from temp_" + GV.box_table_name + partnumber + "_main where img_name='" + filename + "' and isPass=''"
'    Set rs = conn.Execute(query)
'    imageindexedcount = GetImageIndexedCount(GV.filename)
'    If rs.EOF = False Then
'        If Main.chkDelete.Value = 0 And Main.chkFirst.Value = 1 Then
'            GV.total_sample_field = GV.total_sample_field + imageindexedcount
'        End If
'    End If
'
'    If Main.optPass.Value = True Then
'        query = "update temp_" + GV.box_table_name + partnumber + "_main set isPass='PASS' where img_name='" + filename + "'"
'        GV.pass_count = GV.pass_count + imageindexedcount
'    Else
'        query = "update temp_" + GV.box_table_name + partnumber + "_main set isPass='FAIL' where img_name='" + filename + "'"
'        GV.fail_count = GV.fail_count + 1
'    End If
'    conn.Execute (query)
'
'
'    '**** sample fail return 1*****'
'    GV.fail_count = GetFailCount(boxnumber, partnumber)
'    If GV.fail_count >= GV.fail_rate Then
'        GV.finish_sample = 1
'    End If
'
'    If GV.finish_sample = 0 Then
'        If GV.total_sample_field >= GV.sample_field Then
'            GV.finish_sample = 2
'        End If
'    End If
'
'
'ErrLine:
'    conn.Close
'    Set conn = Nothing
'    If Err.Number <> 0 Then
'        MsgBox Err.Description + ": check PassFailInsert"
'        Exit Sub
'    End If
'End Sub

'***************************************'
' Insert page pass/fail in to box table '
'***************************************'
Public Sub PassFailInsert(ByVal boxnumber As String, ByVal partnumber As String, ByVal filename As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim imgSet As ADODB.recordSet
    Dim query As String
    
    Set conn = New ADODB.Connection
    Set imgSet = New ADODB.recordSet
    conn.Open GV.DSN
    
    '**** PCS: Sampling rule: sample pass return 2****'
    query = "select * from " + GV.box_table_name + "_QA where img_name='" + filename + "'"
    Set imgSet = conn.Execute(query)
    If imgSet.EOF = True Then
        GV.total_sample_field = GV.total_sample_field + GetImageIndexedCount(GV.filename)
        If Main.optPass.Value = True Then
            query = "insert into " + GV.box_table_name + "_QA values('" + boxnumber + "','" + partnumber + "','" + filename + "','" + "PASS" + " ',0,'','','')"
        ElseIf Main.optFail.Value = True Then
            query = "insert into " + GV.box_table_name + "_QA values('" + boxnumber + "','" + partnumber + "','" + filename + "','" + "FAIL" + "'," + Main.txtFailNum.Text + ",'','','')"
        End If
    Else
        If Len(imgSet!Status) = 0 Then
            GV.total_sample_field = GV.total_sample_field + GetImageIndexedCount(GV.filename)
        End If

        If Main.optPass.Value = True Then
            query = "update " + GV.box_table_name + "_QA set status='PASS', failcnt=0 where img_name='" + filename + "'"
        ElseIf Main.optFail.Value = True Then
            query = "update " + GV.box_table_name + "_QA set status='FAIL', failcnt=" + Main.txtFailNum.Text + " where img_name='" + filename + "'"
        End If
    End If
    conn.Execute (query)
    
    '**** sample fail return 1*****'
    GV.fail_count = GetFailCount(boxnumber, partnumber)
    If GV.fail_count >= GV.fail_rate Then
        GV.finish_sample = 1
    End If
    
    If GV.finish_sample = 0 Then
        If GV.total_sample_field >= GV.sample_field Then
            GV.finish_sample = 2
        End If
    End If
ErrLine:
    imgSet.Close: Set imgSet = Nothing
    conn.Close: Set conn = Nothing
End Sub

'********************************'
' Get Total indexed data count  '
'********************************'
Public Function GetIndexedCount(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim countSet As ADODB.recordSet
    Dim query As String
    Dim indexedCount As Long
    Dim temp
    
    Set conn = New ADODB.Connection
    Set countSet = New ADODB.recordSet
    conn.Open GV.DSN
    
    indexedCount = 0
    query = "select distinct(img_name) from " + GV.box_table_name + "_main where first_page=1 and delete_page='0' and part_number='" + partnumber + "' order by img_name"
    Set countSet = conn.Execute(query)
    
    While countSet.EOF = False
        indexedCount = indexedCount + GetImageIndexedCount(countSet!img_name)
        countSet.MoveNext
    Wend

    countSet.Close
    Set countSet = Nothing
    GetIndexedCount = indexedCount
ErrLine:
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check GetIndexedCount"
        Exit Function
    End If
End Function

'***************************************'
' Get individual image indexed count    '
'***************************************'
Public Function GetImageIndexedCount(ByVal img_name As String)
On Error GoTo ErrLine
Dim conn As ADODB.Connection
Dim rs As ADODB.recordSet
Dim count As Long

count = 3

'query = "select * from " + GV.box_table_name + "_main where img_name ='" + img_name + "'"
'
'Set rs = conn.Execute(query)
'
'While rs.EOF = False
'    If Len(rs!last_name) <> 0 Then
'        count = count + 1
'    End If
'    If Len(rs!first_name) <> 0 Then
'        count = count + 1
'    End If
'    If Len(rs!dob) <> 0 Then
'        count = count + 3
'    End If
'
'    rs.MoveNext
'Wend
    
    GetImageIndexedCount = count
ErrLine:
'rs.Close: Set rs = Nothing
'conn.Close: Set conn = Nothing


End Function

'***************************************'
'Get Sample Rate                        '
'***************************************'
Public Function GetSampleRate_toDo()
Dim field_count As Double
Dim rate As String
Dim ratetemp

field_count = GetIndexedCount(GV.boxnumber, GV.boxpart)
GV.indexed_field = field_count
rate = GetSampleRate(GV.projectid, GV.curr_phase, field_count)

If StrComp(rate, "0,0") = 0 Then
    MsgBox "Cannot do sampling!! Project sampling rate did not setup! Please contact administrator. Thank You."
    GetSampleRate_toDo = 0
    Exit Function
Else
    ratetemp = Split(rate, ",")
    GV.sample_field = CLng(ratetemp(0))
    GV.fail_rate = CLng(ratetemp(1))
    GetSampleRate_toDo = 1
End If

End Function
