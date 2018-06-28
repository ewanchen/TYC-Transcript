Attribute VB_Name = "CheckModule"
'*******************************'
'Check validatity of input date '
'return 1 if error in month     '
'return 2 if error in day       '
'return 3 if error in year      '
'return 0 if no errors          '
'*******************************'
Public Function check_date(ByVal field As String, ByVal dd As String, ByVal mm As String, ByVal yyyy As String, ByVal mindate As String, ByVal maxdate As String)
On Error GoTo ErrLine
Dim startMonth As String
Dim endMonth As String
Dim startDay As String
Dim endDay As String
Dim leapyear As Integer
Dim currentDate As String
Dim tempDate As Date
Dim indexedDate As String
Dim err_string As String

leapyear = 0
startMonth = "01"
endMonth = "12"
startDay = "01"



Dim compResult As Integer
If Len(yyyy) <> 4 Then
    MsgBox field + ": " + "Please Enter year in YYYY format"
    check_date = 3
    Exit Function
 End If
 
 yyyy = CStr(CInt(yyyy))
indexedDate = mm + "-" + dd + "-" + yyyy
 
'Month validation
If Len(mm) <> 2 Then
    MsgBox field + ": Please enter date in mm-dd-yyyy format, eg 12-02-2005"
    check_date = 1
    Exit Function
ElseIf StrComp(mm, startMonth) = -1 Or StrComp(mm, endMonth) = 1 Then
    MsgBox field + ":" + mm + " is not a Valid month"
    check_date = 1
    Exit Function
End If

Dim year As Integer
year = CInt(yyyy)
If year Mod 4 = 0 Then
    leapyear = 1
End If
If leapyear = 1 And mm = "02" Then
    endDay = "29"
Else
    endDay = "28"
End If
If mm = "01" Or mm = "03" Or mm = "05" Or mm = "07" Or mm = "08" Or mm = "10" Or mm = "12" Then
    endDay = "31"
End If
If mm = "04" Or mm = "06" Or mm = "09" Or mm = "11" Then
    endDay = "30"
End If

'Check valid date of month
If Len(dd) <> 2 Then
   MsgBox field + ": Please enter date in mm-dd-yyyy format, eg 12-02-2005"
   check_date = 2
   Exit Function
ElseIf StrComp(dd, startDay) = -1 Or StrComp(dd, endDay) = 1 Then
    MsgBox field + ":" + dd + " is not a Valid Day for this month and year"
    check_date = 2
    Exit Function
End If

'Check valid year
If yyyy = "" Then
   MsgBox field + ": Year can't be empty"
   check_date = 3
   Exit Function
End If

If Len(yyyy) <> 4 Then
    MsgBox field + ": Year should be in yyyy format"
    check_date = 3
    Exit Function
End If

'Check date not exceed than specific day
Dim difference As Long
If maxdate <> "" Then
  difference = DateDiff("d", indexedDate, maxdate)
  If difference < 0 Then
    MsgBox field + ": is invalid " + indexedDate + " is ahead of " + maxdate
    check_date = 2
    Exit Function
  End If
End If

'Check date not earlier than specific day
If mindate <> "" Then
   difference = DateDiff("d", mindate, indexedDate)
   If difference < 0 Then
        MsgBox field + ": is invalid " + indexedDate + " is earlier than " + mindate
        check_date = 2
        Exit Function
   End If
End If

check_date = 0

ErrLine:
        If (Err.Number <> 0) Then
            err_string = Err.Description
            MsgBox "Error in validate date " + err_string + " " + Str(Err.Number)
            If Err.Number = 6 Then
                MsgBox "Error in Date"
                check_date = 0
            End If
            Exit Function
        End If
End Function

'*****************************'
'Check format of input string '
'return true/false            '
'*****************************'

Public Function check_format(ByVal field As String, ByVal spchar As String, ByVal allowAlpha As Integer, ByVal allowNumeric As Integer, ByVal instring As String, ByVal fixedlength As Integer, ByVal isRequired As Boolean)

Dim regExp, match, i, spec
Dim allowChar As String

If isRequired = True Then
   If Len(instring) = 0 Then
      MsgBox (field + " can not be empty")
      Exit Function
    End If
End If

If fixedlength <> 0 And isRequired = True Then
 If Len(instring) <> fixedlength Then
   MsgBox (field + " must have " + CStr(fixedlength) + " digit")
   Exit Function
 End If
End If

If allowAlpha = 1 Then
   allowChar = allowChar + "[A-Z]"
End If

If allowChar = "" Then
 If allowNumeric = 1 Then
   allowChar = allowChar + "[0-9]"
 End If
Else
   If allowNumeric = 1 Then
      allowChar = allowChar + "|" + "[0-9]"
   End If
End If

If spchar <> "" Then
'   For i = 1 To Len(spchar)
'       If i = 1 Then
        allowChar = allowChar + "|" + "[" + spchar + "]"
'   Next
End If


For i = 1 To Len(instring)
  spec = Mid(instring, i, 1)

Set regExp = New regExp
regExp.Global = True
regExp.IgnoreCase = False
regExp.Pattern = allowChar

Set match = regExp.Execute(spec)

If match.count = 0 Then
  MsgBox field + " has invalid character"
  check_format = False
Exit Function
End If

Set regExp = Nothing
Next
check_format = True

End Function


Public Function check_db(ByVal field As String, ByVal dbname As String, ByVal db_checkfield As String, ByVal instring As String)
Dim qryStrdb As String
Dim conn As New ADODB.Connection
Dim resultSet As New ADODB.recordSet
Dim db_checkvalue() As String
Dim instringvalue() As String
Dim i As Integer
    
conn.Open GV.DSN

db_checkvalue() = Split(db_checkfield, "^")
instringvalue() = Split(instring, "^")

qryStrdb = "select * from " + dbname + " where "
For i = 0 To UBound(db_checkvalue)
    If i <> 0 Then
       qryStrdb = qryStrdb + " AND "
    End If
    qryStrdb = qryStrdb + db_checkvalue(i) + "='" + instringvalue(i) + "'"
Next i

Set resultSet = conn.Execute(qryStrdb)

If resultSet.EOF = True Then
   check_db = False
   MsgBox field + ": " + instring + " isn't matched to the database"

   resultSet.Close: Set resultSet = Nothing
   conn.Close: Set conn = Nothing

   Exit Function
Else
   check_db = True
End If

resultSet.Close: Set resultSet = Nothing
conn.Close: Set conn = Nothing

End Function
