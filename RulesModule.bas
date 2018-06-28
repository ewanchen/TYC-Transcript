Attribute VB_Name = "RulesModule_EC"
'***********************************************'
' Categorize the document type before insertion '
'***********************************************'
Public Sub DocTypeInsert(ByVal boxnumber As String, ByVal partnumber As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim boxtype As String
    Dim fieldstring As String
    Dim tablename As String
    Dim datefield As String
    Dim ssnnum As String
    Dim i As Integer
    Dim j As Integer
    Dim licnum As String
    
    tablename = "temp_" + GV.box_table_name + partnumber + "_main"

    DeleteData tablename, imagename
    
    fieldstring = "'" + Main.txtTYC.Text + "','" + Main.txtLastName.Text + "','" + Main.txtMM.Text + "','" + Main.txtFirstName.Text + "','" + imagename + "','" + boxnumber + "','" + partnumber + "','" + CStr(Main.chkBadScan.Value) + "','" + CStr(Main.chkDelete.Value) + "','" + CStr(Main.chkFirst.Value) + "','0','','','',''"
    SaveData tablename, fieldstring, GV.boxpath + "/" + imagename
    
    If Main.LabelIndexed.Caption <> "Indexed" Then
        GV.count_indexed = GV.count_indexed + 1
        GV.last_imageindexed = GV.filename
    End If
    
    If Main.chkDelete.Value = 0 Then
        GV.data_list = Main.txtLastName.Text + "^" + Main.txtFirstName.Text + "^" + Main.txtMM.Text + "^" + Main.txtTYC.Text
    End If
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check DocTypeInsert"
        Exit Sub
    End If
End Sub

'************'
' Fetch Data '
'************'
Public Sub FetchData(ByVal boxnumber As String, ByVal boxpart As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim recordSet As New ADODB.recordSet
    Dim query As String
    Dim boxtype As String
    Dim temp
    Dim datadate As String
    Dim month As String
    Dim year As String
    Dim day As String
    Dim i As Integer
    Dim j As Integer
    
    conn.Open GV.DSN
    'Set focus to first column when goes to new page
    Main.txtLastName.SetFocus

    query = "select * from temp_" + GV.box_table_name + boxpart + "_main where img_name='" + imagename + "' order by extra2"
    Set recordSet = conn.Execute(query)
    
    resetFields
    
    If recordSet.EOF = True Then
        Main.LabelIndexed.Caption = "Not Indexed"
        PasteData (GV.data_list)
    Else
        Main.LabelIndexed.Caption = "Indexed"
        If (recordSet!last_name <> "") Then
        Main.txtLastName.Text = recordSet!last_name
        End If
        If recordSet!first_name <> "" Then
        Main.txtFirstName.Text = recordSet!first_name
        End If

        Main.chkBadScan.Value = CInt(recordSet!bad_scan)
        Main.chkDelete.Value = CInt(recordSet!delete_page)
        Main.chkFirst.Value = CInt(recordSet!first_page)
    End If
    
    GV.prev_lastname = Main.txtLastName.Text
    GV.prev_firstname = Main.txtFirstName.Text
    GV.prev_TYC = Main.txtTYC.Text
    GV.prev_MM = Main.txtMM.Text

ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check FetchData"
        Exit Sub
    End If
End Sub


'********************************************'
' Paste previous page's data to current page '
'********************************************'
Public Sub PasteData(ByVal data_list As String)
On Error GoTo ErrLine

     Dim data() As String
    data = Split(GV.data_list, "^")
    
    If UBound(data) >= 0 Then
        Main.txtLastName.Text = data(0)
        Main.txtFirstName.Text = data(1)
        Main.txtMM.Text = data(2)
'        Main.txtDD.Text = data(3)
'        Main.txtYY.Text = data(4)
        Main.txtTYC = data(3)
    End If

ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check PasteData"
        Exit Sub
    End If
End Sub



'******************************'
' Convert to upper case        '
'******************************'
Public Sub ConvertToUpper()
On Error GoTo ErrLine

    Main.txtLastName.Text = UCase$(Main.txtLastName.Text)
    Main.txtFirstName.Text = UCase$(Main.txtFirstName.Text)
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox "Contact support: check ConvertToUpper"
        Exit Sub
    End If
End Sub

'********************************'
' Enable text field     '
'***********************'
Public Sub EnableTextField(ByVal Status As Boolean)
On Error GoTo ErrLine

    If GV.curr_phase = "INDEXING" Then
        Main.optFail.Visible = False
        Main.optPass.Visible = False
        Main.txtFailNum.Visible = False
    Else
        Main.optFail.Visible = True
        Main.optPass.Visible = True
        Main.txtFailNum.Visible = True
    End If

    Main.txtLastName.Enabled = Status
    Main.txtFirstName.Enabled = Status
'    Main.txtDD.Enabled = Status
    Main.txtMM.Enabled = Status
    Main.txtTYC.Enabled = Status
    
    If Status = True Then
        Main.txtLastName.SetFocus
    End If

    Main.cmdRL.Enabled = Status
    Main.cmdSaveImage.Enabled = Status
    Main.cmdZoomOut.Enabled = Status
    Main.cmdZoomIn.Enabled = Status
    Main.View_Full_Screen.Enabled = Status
    Main.cmdRR.Enabled = Status
    'Main.cmdSave.Enabled = status
    Main.Next_Image.Enabled = Status
    Main.Previous_Image.Enabled = Status
    Main.Save.Enabled = Status
    Main.Paste.Enabled = Status
    Main.Rotate_Right.Enabled = Status
    Main.cmdUpdate.Enabled = False
    Main.Zoom_In.Enabled = Status
    Main.Zoom_Out.Enabled = Status
    Main.Save_Image.Enabled = Status
    Main.OpenImage.Enabled = False
    Main.OpenImage.Visible = False
    Main.cmdOpenImage.Enabled = False
    Main.cmdOpenImage.Visible = False
    If StrComp(GV.tool_purpose, "Sampling") <> 0 Then
        Main.JumpFirst.Enabled = Status
    Else
        Main.JumpFirst.Enabled = False
    End If
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check EnableTextField"
        Exit Sub
    End If
End Sub

'***********************'
' Switch Box or part    '
'***********************'
Public Function SwitchImage(ByVal currentBox As String, ByVal newBox As String)
On Error GoTo ErrLine
    Dim boxnumber As String
    Dim newboxnumber As String
    Dim boxpart As String
    Dim newboxpart As String
    Dim preboxpart As String
    Dim temp
    
    SwitchImage = 0
    If currentBox <> "" Then
        temp = Split(currentBox, "\")
        boxnumber = temp(0)
        boxpart = Mid(temp(1), Len(temp(1)) - GV.partnumber_length + 1, GV.partnumber_length)
        temp = Split(newBox, "\")
        newboxnumber = temp(0)
        newboxpart = Mid(temp(1), Len(temp(1)) - GV.partnumber_length + 1, GV.partnumber_length)
        
    
            If StrComp(boxnumber, newboxnumber) <> 0 Then
                temp = Split(boxnumber, "-")
                ExitBoxPart boxnumber, boxpart
                EnableTextField False
                GV.data_list = ""
                SwitchImage = 1
                GV.pre_boxnum = currentBox
            Else
                If StrComp(boxpart, newboxpart) <> 0 Then
                    ExitBoxPart boxnumber, boxpart
                    EnableTextField False
                    GV.data_list = ""
                    SwitchImage = 1
                    GV.pre_boxnum = currentBox
                End If
            End If
            'GV.pre_boxnum = currentBox
    Else
        SwitchImage = 1
    End If
    
    
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox "Contact support: check SwitchImage"
        Exit Function
    End If
End Function


'**********************'
' Check Change fields  '
'**********************'
Public Function CheckChangeField(ByVal opt As Boolean)
On Error GoTo ErrLine

If GV.prev_lastname <> Main.txtLastName.Text Then
    CheckChangeField = True
    Exit Function
End If

If GV.prev_firstname <> Main.txtFirstName.Text Then
    CheckChangeField = True
    Exit Function
End If

If GV.prev_TYC <> Main.txtTYC.Text Then
    CheckChangeField = True
    Exit Function
End If

If opt = True Then
If GV.prev_TYC <> Main.txtTYC.Text Then
    CheckChangeField = True
    Exit Function
End If

If GV.prev_MM <> Main.txtMM.Text Then
    CheckChangeField = True
    Exit Function
End If

'If GV.prev_YY <> Main.txtYY.Text Then
'    CheckChangeField = True
'    Exit Function
'End If
End If

CheckChangeField = False
            
ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check CheckChangeField"
        CheckChangeField = 0
        Exit Function
    End If
End Function

'*********************************************************************'
' Get project sample rate                                             '
' Return project sample size and fail rate base on the phase          '
' exampe: (sameple size, fail rate) -- (2,1)                          '
' if the projectID does not exist in the database, it will return 0,0 '
'*********************************************************************'
Public Function GetSampleRate(ByVal projectid As String, ByVal phase As String, ByVal totalnum As Double)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim rateSet As ADODB.recordSet
    Dim letterSet As ADODB.recordSet
    Dim query As String
    Dim letter As String
    Dim rate As String
    Dim Status As String
    
    Set conn = New ADODB.Connection
    Set rateSet = New ADODB.recordSet
    Set letterSet = New ADODB.recordSet
    
    conn.Open GV.masterDSN
    query = "select max(minimum_lot+0) maxnum from sample_letter"
    Set letterSet = conn.Execute(query)
    If CLng(totalnum) > CLng(letterSet!maxnum) Then
        query = "select letter from sample_letter where minimum_lot='" + letterSet!maxnum + "'"
    Else
        query = "select letter from sample_letter where " + CStr(totalnum) + " >= (minimum_lot+0) and " + CStr(totalnum) + " <= (maximum_lot+0)"
    End If
    Set rateSet = conn.Execute(query)
    letter = rateSet!letter
    rateSet.Close
    
    query = "select * from project_sample_rate where project_id='" + projectid + "'"
    Set rateSet = conn.Execute(query)
    If rateSet.EOF = True Then
        GetSampleRate = "0,0"
    Else
        If StrComp(phase, "IMAGE SAMPLING") = 0 Then
            rate = rateSet!image_sample_rate
        Else
            rate = rateSet!index_sample_rate
        End If
        rateSet.Close
        
        '*** Depends on indexer's status, select sample rate from different table ***'
        Status = IndexerStatus(GV.boxbarcode, GV.boxpart)
        If Status = "False" Then
            GetSampleRate = "0,0"
        Else
            If Status = "NORMAL" Then
                query = "select * from sample_plan where letter='" + letter + "' and sample_rate='" + rate + "'"
            Else
                query = "select * from tightened_sample_plan where letter='" + letter + "' and sample_rate='" + rate + "'"
            End If
            Set rateSet = conn.Execute(query)
            If rateSet.EOF = True Then
                GetSampleRate = "0,0"
            Else
                GetSampleRate = rateSet!sample_size + "," + rateSet!fail_size
            End If
        End If
    End If
    

ErrLine:
    rateSet.Close
    Set rateSet = Nothing
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check GetSampleRate"
        Exit Function
    End If
End Function

'*******************************'
' find out the indexer's status '
'*******************************'
Public Function IndexerStatus(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim pconn As ADODB.Connection
    Dim IndexerSet As ADODB.recordSet
    Dim query As String
    
    Set conn = New ADODB.Connection
    Set IndexerSet = New ADODB.recordSet
    Set pIndexerSet = New ADODB.recordSet
    
    conn.Open GV.DSN
    
    query = "select * from user_boxpart where boxpart='" + boxnumber + "-" + partnumber + "' and phase='INDEXING' order by time desc"
    Set pIndexerSet = conn.Execute(query)
    If pIndexerSet.EOF = True Then
        IndexerStatus = "False"
    Else
        GV.indexer = pIndexerSet!username
        query = "select * from user_chart where username='" + GV.indexer + "'"
        Set IndexerSet = conn.Execute(query)
        If IndexerSet.EOF = True Then
            query = "insert into user_chart values('" + GV.indexer + "','NEW','0')"
            conn.Execute (query)
            IndexerStatus = "NEW"
        Else
            IndexerStatus = IndexerSet!Status
        End If
        IndexerSet.Close
        Set IndexerSet = Nothing
    End If
    pIndexerSet.Close
    Set pIndexerSet = Nothing
    
ErrLine:
    conn.Close
    Set conn = Nothing
   

    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check IndexerStatus: Possibly no records of the indexer of this boxpart"
        Exit Function
    End If
End Function

'*******************************************'
' Randomly generate a number for image      '
'*******************************************'
Public Function SampleImage()
On Error GoTo ErrLine
    Dim imgNum As Integer
    Dim totalImage As Integer
    Dim samplePage As Integer
    Dim conn As ADODB.Connection
    Dim rs As ADODB.recordSet
    Dim query As String
    
    Set conn = New ADODB.Connection
    conn.Open GV.DSN
    
    query = "select count(*) count from " + GV.box_table_name + "_main where first_page='1' and delete_page='0' and part_number='" + GV.boxpart + "'"
    Set rs = conn.Execute(query)
    
    totalImage = rs!count
    Randomize
    imgNum = Int((totalImage) * Rnd)
    samplePage = GV.sample_field / GV.indexed_field
    If (totalImage - (imgNum + 1)) < samplePage Then
        imgNum = imgNum - (samplePage - (totalImage - (imgNum + 1)))
    End If
    
    If (imgNum < 0) Then
        imgNum = 2
    End If
    
    SampleImage = imgNum
    GV.currentImage = imgNum

ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check SampleImage"
        Exit Function
    End If
End Function

'*******************************'
' Sampling Calculation          '
'*******************************'
Public Sub SampCalculate(ByVal boxnum As String, ByVal partnum As String)
On Error GoTo ErrLine

    
    If GV.finish_sample = 1 Then
        MsgBox boxnum + "-" + partnum + ": FAIL index sampling"
    ElseIf GV.finish_sample = 2 Then
        MsgBox boxnum + "-" + partnum + ": PASS index sampling"
    End If

    
ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check SampCalculate"
        Exit Sub
    End If
End Sub

'********************************'
' update user status             '
'********************************'
Public Sub UpdateUserStatus(ByVal username As String, ByVal pass_flag As Integer)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim Mconn As ADODB.Connection
    Dim userSet As ADODB.recordSet
    Dim query As String
    Dim passTimes As Long
    Dim currentPass As Long
    
    Set conn = New ADODB.Connection
    Set Mconn = New ADODB.Connection
    Set userSet = New ADODB.recordSet
    conn.Open GV.DSN
    Mconn.Open GV.masterDSN
    
    query = "select * from backup_variable where purpose='pass_times'"
    Set userSet = Mconn.Execute(query)
    If userSet.EOF = False Then
        passTimes = CLng(userSet!content)
    Else
        passTimes = 50
    End If
    userSet.Close
    
    '*** Check to make sure the user exist in the database ***'
    query = "select * from user_chart where username='" + username + "'"
    Set userSet = conn.Execute(query)
    If userSet.EOF = True Then
        query = "insert into user_chart values('" + username + "','NEW','0')"
        conn.Execute (query)
    Else
        If userSet!pass_times = "" Then
            currentPass = 0
        Else
            currentPass = CLng(userSet!pass_times)
        End If
        If pass_flag = 0 Then
            query = "update user_chart set pass_times=0, status='NEW' where username='" + username + "'"
            conn.Execute (query)
        Else
            currentPass = currentPass + 1
            If currentPass < passTimes Then
                query = "update user_chart set pass_times='" + CStr(currentPass) + "' where username='" + username + "'"
                conn.Execute (query)
            Else
                query = "update user_chart set pass_times='" + CStr(currentPass) + "', status='NORMAL' where username='" + username + "'"
                conn.Execute (query)
            End If
        End If
    End If
    userSet.Close
    Set userSet = Nothing

ErrLine:
    conn.Close
    Set conn = Nothing
    Mconn.Close
    Set Mconn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check UpdateUserStatus"
        Exit Sub
    End If
End Sub

Public Sub userBoxpartPhase(ByVal boxnumber As String, ByVal partnumber As String, ByVal username As String, ByVal phase As String)
On Error GoTo ErrLine
Dim conn As ADODB.Connection
Dim query As String

Set conn = New ADODB.Connection

conn.Open GV.DSN

query = "insert into user_boxpart values('" + username + "','" + boxnumber + "-" + partnumber + "','" + phase + "','" + GV.ToolVersion + "','" + CStr(Date) + " " + Format(Time, "HH:MM:SS") + "')"
conn.Execute query

ErrLine:
conn.Close: Set conn = Nothing

End Sub

'****************************************'
' Get fail Count                         '
'****************************************'
Public Function GetFailCount(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim failSet As ADODB.recordSet
    Dim query As String
    Dim count As Integer
    Dim temp
    
    Set conn = New ADODB.Connection
    Set failSet = New ADODB.recordSet
    conn.Open GV.DSN
    count = 0
    query = "select * from " + GV.box_table_name + "_QA where partnumber='" + partnumber + "' and status='FAIL'"
    Set failSet = conn.Execute(query)
    While failSet.EOF = False
        count = count + failSet!failcnt
        failSet.MoveNext
    Wend
    failSet.Close
    Set failSet = Nothing
    GetFailCount = count
ErrLine:
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check GetFailCount"
        GetFailCount = 0
        Exit Function
    End If
End Function

''*****************************************'
'' clear out pass/fail information  '
''**********************************'
'Public Sub clearout(ByVal boxnumber As String, ByVal partnumber As String)
'On Error GoTo ErrLine
'    Dim conn As ADODB.Connection
'    Dim query As String
'
'
'    Set conn = New ADODB.Connection
'    conn.Open GV.DSN
'    query = "update temp_" + GV.box_table_name + partnumber + "_main set isPass=''"
'    conn.Execute (query)
'
'ErrLine:
'    conn.Close
'    Set conn = Nothing
'    If Err.Number <> 0 Then
'        MsgBox Err.Description + ": check clearout"
'        Exit Sub
'    End If
'
'End Sub
