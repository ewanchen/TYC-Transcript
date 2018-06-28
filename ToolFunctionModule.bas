Attribute VB_Name = "ToolFunctionModule"

'************************************'
' This function open the image folder'
'************************************'
Public Function OpenImageFolder(ByVal pathname As String)
On Error GoTo ErrLine
    Dim temp As String
    
    temp = Dir(pathname, vbDirectory)
    If Len(temp) <> 0 Then
        FormDialog.DirFolder.path = pathname
        FormDialog.FileDialog.path = pathname
        OpenImageFolder = 1
    Else
        MsgBox "Image foler " + pathname + " does not exist!!"
        OpenImageFolder = 0
    End If
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check OpenImageFoler"
        Exit Function
    End If
End Function

'******************'
' Display the image'
'******************'
Public Sub DisplayImage()
On Error GoTo ErrLine
    Dim temp As String
    Dim used
    
    Main.ImgEdit1.image = GV.ImageFolder + "\" + FormDialog.FileDialog.filename
    GV.filepath = GV.ImageFolder + "\" + FormDialog.FileDialog.filename
    GV.currentImage = FormDialog.FileDialog.ListIndex
    Main.ImgEdit1.FitTo (0)
'    Main.ImgEdit1.Zoom = 30
    
    Main.ImgEdit1.Display
    RotateLeftImage
    GV.filename = FormDialog.FileDialog.filename
    GV.boxfilename = GV.boxbarcode + "\" + GV.boxnumber + GV.boxpart + "\" + GV.filename
    Main.Caption = GV.Tool + " " + GV.ToolVersion
    Main.LabelImage.Caption = GV.boxfilename
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check DisplayImage"
        Exit Sub
    End If
End Sub

'***************'
' Zoom in Image '
'***************'
Public Sub ZoomInImage()
On Error GoTo ErrLine
    Main.ImgEdit1.Zoom = Main.ImgEdit1.Zoom + 5
    If Not (Main.ImgEdit1.image = "") Then
        Main.ImgEdit1.Display
   End If
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check ZoomInImage"
        Exit Sub
    End If
End Sub

'****************'
' Zoom Out Image '
'****************'
Public Sub ZoomOutImage()
On Error GoTo ErrLine
    Main.ImgEdit1.Zoom = Main.ImgEdit1.Zoom - 5
    If Not (Main.ImgEdit1.image = "") Then
        Main.ImgEdit1.Display
    End If
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check ZoomOutImage"
        Exit Sub
    End If
End Sub

'************************'
' Rotate Image to left   '
'************************'
Public Sub RotateLeftImage()
On Error GoTo ErrLine
    Main.ImgEdit1.RotateLeft
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check RotateLeftImage"
        Exit Sub
    End If
End Sub

'*************************'
' Rotate Image to right   '
'*************************'
Public Sub RotateRightImage()
On Error GoTo ErrLine
    Dim pathname As String
    Main.ImgEdit1.RotateRight
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check RotateRightImage"
        Exit Sub
    End If
End Sub

'****************'
' Get Box Number '
' example of pathname: \9900\990001\990001aaa.tif'
'****************'
Public Function boxnumber(ByVal pathname As String)
On Error GoTo ErrLine
    Dim boxno As String
    Dim list As Variant
    Dim tok As Integer
    
    list = Split(pathname, "\")
    tok = UBound(list)
    boxnumber = list(tok - 2)
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check BoxNumber"
        Exit Function
    End If

End Function

'*****************'
' Get Part Number '
' example of pathname: \9900\990001\990001aaa.tif'
'*****************'
Public Function partnumber(ByVal pathname As String)
On Error GoTo ErrLine
    Dim partno As String
    Dim list As Variant
    Dim tok As Integer
    
    list = Split(pathname, "\")
    tok = UBound(list)
    partnumber = list(tok - 1)

ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check PartNumber"
        Exit Function
    End If

End Function



'****************************'
' Open Box part for indexing '
'****************************'
Public Function OpenBoxPart(ByVal boxnumber As String, ByVal boxpart As String, ByVal indexername As String, ByVal switch As Integer)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim boxSet As New ADODB.recordSet
    Dim boxquery As String
    Dim insertquery As String
    Dim updatequery As String
    Dim used As Integer
    Dim temp
    Dim field_count As Double
    Dim rate As String
    Dim ratetemp
    
    '*** PCSII: Clock user in/ check out box+part ***'
    If switch = 1 Then
            temp = clockUserIn(GV.boxid, boxpart, GV.userid, GV.curr_phase, 0)
            If StrComp(temp, "FAIL") = 0 Then
                MsgBox "This box part is currently using by other indexer. Please pick other box"
                OpenBoxPart = 0
                Exit Function
            End If

        CreateNewBoxTable boxnumber, boxpart
        CreateTempBoxTable boxnumber, boxpart
        OpenBoxPart = 1
    Else
        OpenBoxPart = 0
    End If
    
ErrLine:
    If (Err.Number <> 0) Then
        GV.next_state = "INDEXING"
        clockUserOut GV.boxid, boxpart, GV.userid, GV.next_state, 0
        Exit Function
    End If
End Function
'*****************'
' Exit Box Part   '
'*****************'
Public Sub ExitBoxPart(ByVal boxnumber As String, ByVal boxpart As String)
On Error GoTo ErrLine
    Dim query As String
    Dim temp
    Dim totalIndexCount As Double
    Dim purpose As String
    Dim partstatus
    Dim boximage As Double
    Dim box_id As String
    
    '***** generate error before exit the part *****'
If Len(boxnumber) <> 0 Then
    If GV.curr_phase = "INDEXING" Then
        If CheckInputCount(boxnumber, boxpart) = 1 Then
            '*** PCSII update boxpart status and clock the user out ***'
            '*** PCSII clock user out ***'
            If StrComp(GV.tool_purpose, "Production") = 0 Then
                GV.next_state = "INDEXED"
            Else
                GV.next_state = "INDEXING"
            End If
        Else
            GV.next_state = "INDEXING"
        End If
    Else
        If GV.finish_sample = 2 Then
            GV.next_state = "INDEX_SAMPLED"
            UpdateUserStatus GV.indexer, 1
            Pass_Fail_log boxnumber, boxpart, GV.uname, GV.indexer, "PASS"
        Else
            GV.next_state = "INDEXING"
            UpdateUserStatus GV.indexer, 0
            Pass_Fail_log boxnumber, boxpart, GV.uname, GV.indexer, "FAIL"
        End If
    End If
            '*** PCSII update boxpart status and clock the user out ***'
            '*** PCSII clock out the user ***'
            box_id = getBoxId(boxnumber, GV.projectid, 0)
            temp = clockUserOut(box_id, boxpart, GV.userid, GV.next_state, 0)
            If StrComp(temp, "FAIL") = 0 Then
                MsgBox "Cannot logout the user!! please report the box part number and your initial to your manager and contact support. Thank You!"
            End If
            
            '*** Update box part status to "INDEXING"
            If StrComp(GV.tool_purpose, "Train") <> 0 Then
                temp = UpdatePartStatus(box_id, boxpart, GV.next_state, GV.userid, "")
                If StrComp(temp, "FAIL") = 0 Then
                    MsgBox "Cannot update the box part status!! Please report the box part number and your initial to your manager and contact support.  Thank You!"
                End If
            End If

    
    GV.finish_sample = 0
    DropTempTable boxnumber, boxpart
    End If
ErrLine:
    If (Err.Number <> 0) Then
        GV.next_state = "INDEXING"
        clockUserOut GV.boxid, boxpart, GV.userid, GV.next_state, 0
        Exit Sub
    End If
End Sub


'***********************'
' Create box main table '
'***********************'
Public Sub CreateNewBoxTable(ByVal boxnumber As String, ByVal boxpart As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim query As String
    Dim deleteQuery As String

    conn.Open GV.DSN
        query = "create table " + GV.box_table_name + "_main select * from struct_main"
        conn.Execute (query)
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
'        MsgBox Err.Description + ": check CreateNewBoxTable"
'        Exit Sub
    End If
End Sub

'***************************'
' Create temp main table    '
'***************************'
Public Sub CreateTempBoxTable(ByVal boxnumber As String, ByVal boxpart As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    Dim deleteQuery As String
    
    conn.Open GV.DSN
On Error GoTo CreateMain
    query = "select count(*) from " + GV.box_table_name + "_main"
    conn.Execute (query)
CreateMain:
On Error GoTo ErrLine
    If (Err.Number <> 0) Then
            query = "create table " + GV.box_table_name + "_main select * from struct_main"
            conn.Execute (query)
    End If
    query = "create table temp_" + GV.box_table_name + boxpart + "_main select * from " + GV.box_table_name + "_main where part_number='" + boxpart + "'"
    conn.Execute query
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        Exit Sub
    End If
End Sub

'****************************************************************'
' Save all the information back to main table and drop temp table'
'****************************************************************'
Public Sub DropTempTable(ByVal boxnumber As String, ByVal boxpart As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    Dim boxtype As String
    Dim boxtable As String
    
    boxtable = Replace(boxnumber, "-", "_")
    
    conn.Open GV.DSN
    
    On Error GoTo CreateTable
    query = "drop table " + boxtable + boxpart + "_main_backup"
    conn.Execute query
    
CreateTable:
    query = "create table " + boxtable + boxpart + "_main_backup select * from temp_" + boxtable + boxpart + "_main"
    conn.Execute query
    
On Error GoTo CreateMain
    query = "delete from " + boxtable + "_main where part_number='" + boxpart + "'"
    conn.Execute query

CreateMain:
    If (Err.Number <> 0) Then
        query = "create table " + boxtable + "_main select * from new_struct_main"
        conn.Execute (query)
    End If
    
    query = "insert into " + boxtable + "_main select * from temp_" + boxtable + boxpart + "_main"
    conn.Execute query
    
    query = "drop table temp_" + boxtable + boxpart + "_main"
    conn.Execute query
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check DropTempTable"
        Exit Sub
    End If
End Sub

'********************'
' Save indexing data '
'********************'
Public Sub SaveData(ByVal tablename As String, ByVal insertvalues As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim query As String
    
    conn.Open GV.DSN
    temp = Split(insertvalues, ",")
    query = "insert into " + tablename + " values(" + insertvalues + ")"
    conn.Execute query

ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check SaveData"
        Exit Sub
    End If
End Sub

'*************************'
' Delete indexing data    '
'*************************'
Public Sub DeleteData(ByVal tablename As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim query As String
    
    conn.Open GV.DSN
    query = "delete from " + tablename + " where img_name='" + imagename + "'"
    conn.Execute query
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check DeleteData"
        Exit Sub
    End If
End Sub
'*************************'
'Save Updated Data        '
'*************************'
Public Sub saveUpdateData(ByVal tablename As String, ByVal updatevalues As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim query As String
    
    Set conn = New ADODB.Connection
    
    conn.Open GV.DSN
    query = "update " + tablename + " set " + updatevalues + " where img_name='" + imagename + "'"
    conn.Execute query

ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check saveUpdateData"
        Exit Sub
    End If
End Sub

'*****************'
' Save array data '
'*****************'
Public Sub SaveArrayData(ByVal tablename As String, ByVal insertvalues As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim query As String
    
    conn.Open GV.DSN
    
    query = "insert into " + tablename + " values(" + insertvalues + ")"
    conn.Execute query
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check SaveArrayData"
        Exit Sub
    End If
End Sub

'*******************'
' Delete array data '
'*******************'
Public Sub DeleteArrayData(ByVal tablename As String, ByVal imagename As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim query
    
    conn.Open GV.DSN
    
    query = "delete from " + tablename + " where img_name='" + imagename + "'"
    conn.Execute query
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check DeleteArrayData"
        Exit Sub
    End If
End Sub
'**********************'
' View Next Image      '
'**********************'
Public Sub ViewNextImage(ByVal flag As Integer)
On Error GoTo ErrLine
    Dim image As String
    Dim temp As String

FormDialog.FileDialog = GV.ImageFolder
    If GV.currentImage < FormDialog.FileDialog.ListCount - 1 Then
        GV.currentImage = GV.currentImage + 1
        image = GV.ImageFolder + "\" + FormDialog.FileDialog.list(GV.currentImage)
        GV.filepath = image
        Main.ImgEdit1.image = image
        Main.ImgEdit1.FitTo (0)
        Main.ImgEdit1.Zoom = 30
        Main.ImgEdit1.Display
        RotateLeftImage
        GV.filename = FormDialog.FileDialog.list(GV.currentImage)
        temp = GV.boxbarcode + "\" + GV.boxnumber + GV.boxpart + "\" + GV.filename
        Main.LabelImage.Caption = temp
        GV.end_flag = 0
    Else
        MsgBox "This is the last image"
        GV.end_flag = 1
    End If
    
    'close imagefolder
    'Unload FormDialog
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check ViewNextImage"
        Exit Sub
    End If
End Sub

'**************************'
' View Previous Image      '
'**************************'
Public Sub ViewPreviousImage()
On Error GoTo ErrLine
    Dim image As String
    Dim temp As String
    Dim count As Integer
    
    FormDialog.FileDialog = GV.ImageFolder
    count = GV.currentImage - 1
    If count >= 0 Then
        GV.currentImage = GV.currentImage - 1
        image = GV.ImageFolder + "\" + FormDialog.FileDialog.list(GV.currentImage)
        GV.filepath = image
        Main.ImgEdit1.image = image
        Main.ImgEdit1.FitTo (0)
        Main.ImgEdit1.Zoom = 30
        Main.ImgEdit1.Display
        RotateLeftImage
        GV.filename = CStr(FormDialog.FileDialog.list(GV.currentImage))
        temp = GV.boxbarcode + "\" + GV.boxnumber + GV.boxpart + "\" + GV.filename
        Main.LabelImage.Caption = temp
        GV.pre_page = True
    Else
        MsgBox "This is the first image"
    End If
    
    'close image folder
'    Unload FormDialog
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check ViewPreviousImage"
        Exit Sub
    End If
End Sub


'*********************'
' Last Index image    '
'*********************'
Public Function lastindex(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim resultSet As New ADODB.recordSet
    Dim imagename As String
    Dim query
    Dim temp
    
    conn.Open GV.DSN
    query = "select max(img_name) maxImage from temp_" + GV.box_table_name + partnumber + "_main"
    Set resultSet = conn.Execute(query)
    
    If resultSet!maxImage <> "" Then
        lastindex = resultSet!maxImage
    Else
        lastindex = GV.boxnumber + partnumber + "aaa.tif"
    End If
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check LastIndex"
        Exit Function
    End If
End Function
'*********************'
' First Index image    '
'*********************'
Public Function firstindex(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As New ADODB.Connection
    Dim resultSet As New ADODB.recordSet
    Dim imagename As String
    Dim query
    Dim temp
    
    conn.Open GV.DSN
    query = "select min(img_name) minImage from temp_" + GV.box_table_name + partnumber + "_main"
    Set resultSet = conn.Execute(query)
    
    If resultSet!minImage <> "" Then
        firstindex = resultSet!minImage
    Else
        temp = Split(boxnumber, "-")
        firstindex = CStr(CLng(temp(2))) + partnumber + "aaa.tif"
    End If
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check firstIndex"
        Exit Function
    End If
End Function
'*********************************'
'Get Index                        '
'*********************************'
Public Function GetIndex(ByVal imgname As String)
On Error GoTo ErrLine
Dim i As Long

i = 0

While i < FormDialog.FileDialog.ListCount
    If FormDialog.FileDialog.list(i) = imgname Then
        GetIndex = i
        Exit Function
    End If
    i = i + 1
Wend

GetIndex = -1

ErrLine:
    If Err.Number <> 0 Then
        Exit Function
    End If

End Function

'**********************'
' Add leading 0's      '
'**********************'
Public Function AddZero(ByVal num As String, ByVal Length As Integer)
On Error GoTo ErrLine
    Dim temp As String
    Dim zero As Integer
    
    temp = ""
    zero = Length - Len(num)
    While zero > 0
        temp = temp + "0"
        zero = zero - 1
    Wend
    
    AddZero = temp + num
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check AddZero"
        Exit Function
    End If
End Function

'****************************************'
' Remove Leading Zero's                  '
'****************************************'
Public Function removeZero(ByVal num As String)
On Error GoTo ErrLine
    Dim i As Integer
    Dim tempnum As String
    
    i = 1
    tempnum = ""
    Do
        If StrComp(Mid(num, i, 1), "0") <> 0 Then
            removeZero = Mid(num, i, Len(num) - i + 1)
            Exit Do
        End If
        i = i + 1
    Loop While i <= Len(num)

ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check removeZero"
        Exit Function
    End If
End Function

'*******************************************************'
'Check Special Characters and add \ before the character'
'*******************************************************'
Public Function CheckSpec(ByVal name As String)
On Error GoTo ErrLine
    Dim i As Integer
    Dim temp As String
    
    i = 1
    While i <= Len(name)
        If Mid(name, i, 1) = "\" Or Mid(name, i, 1) = "'" Then
            temp = temp + "\"
            temp = temp + Mid(name, i, 1)
        Else
            temp = temp + Mid(name, i, 1)
        End If
        i = i + 1
    Wend
        
    CheckSpec = temp
    
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check CheckSpec"
        Exit Function
    End If
End Function

'***************************'
' Save Image after rotation '
'***************************'
Public Sub SaveImage()
On Error GoTo ErrLine
'    Main.ImgEdit1.SaveAs (GV.filepath), , , Main.ImgEdit1.CompressionType, Main.ImgEdit1.CompressionInfo
    Main.ImgEdit1.SaveAs (GV.filepath), , 5, 8
ErrLine:
    If (Err.Number <> 0) Then
        MsgBox Err.Description + ": check SaveImage"
        Exit Sub
    End If
End Sub

'*********************************'
'View Image Full Screen Mode      '
'*********************************'
Public Sub viewmode()
If GV.fullscreen = False Then
'    Main.frmIndependent.Visible = False
    GV.frameWidth = Main.Width - Main.ImgFrame.Width - 250
    GV.imgWidth = Main.Width - Main.ImgEdit1.Width - 350
    GV.frameHeight = Main.Height - Main.ImgFrame.Height - 1500
    GV.imgHeight = Main.Height - Main.ImgEdit1.Height - 1500
    
    Main.ImgFrame.Width = Main.ImgFrame.Width + GV.frameWidth
    Main.ImgEdit1.Width = Main.ImgEdit1.Width + GV.imgWidth
    Main.ImgFrame.Height = Main.ImgFrame.Height + GV.frameHeight
    Main.ImgEdit1.Height = Main.ImgEdit1.Height + GV.imgHeight
    
    Main.ImgFrame.Top = Main.ImgFrame.Top - 50
    Main.ImgEdit1.Top = Main.ImgEdit1.Top - 200
    
    Main.ImgEdit1.Zoom = Main.ImgEdit1.Zoom + 35
    Main.IndexFrame.Visible = False
    
    Main.ImgEdit1.Display
    GV.fullscreen = True
    Main.cmdOpen.Enabled = False
    Main.cmdOpenImage.Enabled = False

Else
    
    Main.ImgFrame.Width = Main.ImgFrame.Width - GV.frameWidth
    Main.ImgEdit1.Width = Main.ImgEdit1.Width - GV.imgWidth
    Main.ImgFrame.Height = Main.ImgFrame.Height - GV.frameHeight
    Main.ImgEdit1.Height = Main.ImgEdit1.Height - GV.imgHeight

    Main.ImgEdit1.Top = Main.ImgEdit1.Top + 200
    Main.ImgFrame.Top = Main.ImgFrame.Top + 100
    
    Main.ImgEdit1.Zoom = Main.ImgEdit1.Zoom - 25
    Main.ImgEdit1.FitTo (0)
    Main.ImgEdit1.Display
    
'    Main.frmIndependent.Visible = True
    
    GV.fullscreen = False
    Main.cmdOpen.Enabled = True
    Main.cmdOpenImage.Enabled = True
    Main.IndexFrame.Visible = True
    EnableTextField True
End If

End Sub
'******************************'
' Get image path, export path  '
'******************************'
Public Sub GetLocation()
On Error GoTo ErrLine
    Dim temp
    
    '*** PCSII: Get location from PCS ***'
    temp = GetProjectSetting(GV.projectid, "input_location")
    If StrComp(temp, "FAIL") = 0 Then
        MsgBox "Cannot get input location! Please contact your manager and report this to IT support.  Thank You!"
        Exit Sub
    Else
        GV.imagepath = temp
    End If
    
    '*** PCSII: Get blank page threshold from PCS ***'
    temp = GetProjectSetting(GV.projectid, "blank_page_set")
    If StrComp(temp, "FAIL") = 0 Then
        MsgBox "Cannot get blank page threshold! Please contact your manager and report this to IT support.  Thank You!"
        Exit Sub
    Else
        GV.blank_page_threshold = temp
    End If
    
    '*** PCSII: Get version from PCS ***'
    temp = GetProjectSetting(GV.projectid, "IndexingTool")
    If StrComp(temp, "FAIL") = 0 Then
        MsgBox "Cannot get tool version! Please contact your manager and report this to IT support. Thanks!"
        Exit Sub
    Else
        GV.version = temp
    End If
    
ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check GetLocation"
        Exit Sub
    End If
End Sub




'***************************************************'
' Check image and indexed count (input)             '
' if image and indexed count match, return 1        '
' if image and indexed count don't match , return 0 '
' This function is used for auto checking the error '
'***************************************************'
Public Function CheckInputCount(ByVal boxnum As String, ByVal partnum As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim imgSet As ADODB.recordSet
    Dim query As String
    Dim boxtable As String
    
    Set conn = New ADODB.Connection
    Set imgSet = New ADODB.recordSet
    conn.Open GV.DSN
    
    boxtable = Replace(boxnum, "-", "_")
    
    If GV.tool_purpose <> "Sampling" Then
        ' Get the index count
        query = "select count(distinct(img_name)) imgc from temp_" + boxtable + partnum + "_main"
        Set imgSet = conn.Execute(query)
        If imgSet.EOF = False Then
            GV.indexcount = imgSet!imgc
        Else
            GV.indexcount = 0
        End If
        If GV.indexcount = GV.input_imagecount Then
            CheckInputCount = 1
        Else
            CheckInputCount = 0
        End If
    Else
        'Get the fail page count
        query = "select count(distinct(img_name)) imgc from temp_" + boxtable + partnum + "_main"
        Set imgSet = conn.Execute(query)
        If imgSet.EOF = False Then
            GV.failcount = CLng(imgSet!imgc)
        Else
            GV.failcount = 0
        End If
        CheckInputCount = 1
    End If
    
ErrLine:
    imgSet.Close
    Set imgSet = Nothing
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check CheckInputCount"
        Exit Function
    End If
    
End Function

'*****************************'
' Activity log Record         '
'*****************************'
Public Sub Activity_Log(ByVal imagename As String, ByVal action As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim query As String
    
    Set conn = New ADODB.Connection
    conn.Open GV.DSN
    query = "insert into activity_log values('" + GV.boxnumber + "/" + imagename + "','" + GV.uname + "','" + GV.end_date + "','" + GV.end_time + "','" + GV.start_date + "','" + GV.start_time + "','" + action + "')"
    conn.Execute (query)

ErrLine:
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check Activity_Log"
        Exit Sub
    End If
End Sub

'********************************'
' Get box part input image count '
'********************************'
Public Sub InputImageCount(ByVal boxnum As String, ByVal partnum As String)
On Error GoTo ErrLine
    
    GV.input_imagecount = FormDialog.List1.ListCount

ErrLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check InputImageCount"
        Exit Sub
    End If
End Sub

Public Sub testDSN()
On Error GoTo ErrLine
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.Open GV.DSN

ErrLine:
    conn.Close: Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox "TYCTRANSCRIPTSDSN is not setup correctly. Please contact support!!"
        Exit Sub
    End If
End Sub

Public Sub openBox_todo()
    Dim used
    Dim found
    Dim lastimage
    Dim foldername
    Dim count As Integer
    Dim switch As Integer
    Dim tempbox
    Dim path
    Dim temp
    Dim i As Integer
    Dim imageindex As Long
    
            GV.fail_reason = ""

            If GV.OpenImage = False And StrComp(FormDialog.List1.Text, GV.boxnumber + "-" + GV.boxpart) = 0 And GV.open_beyond_image = 0 Then
                MsgBox "You choose the same box part! Please pick another box part"
                Exit Sub
            End If
            GV.boxid = getBoxId("", GV.projectid, GV.boxnumber)
            If StrComp(CStr(GV.boxid), "FAIL") = 0 Then
                MsgBox "Cannot get box id from PCS! Cannot open box"
                Exit Sub
            End If
            GV.ImageFolder = GV.imagepath + "\" + GV.boxbarcode + "\" + GV.boxnumber + AddZero(GV.boxpart, GV.partnumber_length)
            FormDialog.FileDialog.path = GV.ImageFolder

            GV.box_table_name = Replace(GV.boxbarcode, "-", "_")
            '*** Check to see if the user trying to change box ***'
            switch = SwitchImage(GV.boxfilename, GV.boxbarcode + "\" + GV.boxnumber + GV.boxpart + "\" + FormDialog.FileDialog.filename)
            end_flag = 0
            '*** Check to see if the user can index the box part ***'
            used = OpenBoxPart(GV.boxnumber, GV.boxpart, GV.uname, switch)
            If used = 1 Then
                createQAtable GV.boxbarcode, GV.boxpart
                GV.input_imagecount = FormDialog.FileDialog.ListCount
                If GV.curr_phase = "INDEX_SAMPLING" Then
                    'initialize
                    GV.finish_sample = 0
                    GV.fail_count = 0
                    GV.pass_count = 0
                    GV.total_sample_field = 0
                    GV.currentImage = 0
                    clearout GV.boxnumber, GV.boxpart
                    GetSampleRate_toDo
                    FormDialog.FileDialog.Selected(SampleImage) = True
                Else
                    lastimage = lastindex(GV.boxbarcode, GV.boxpart)
                    imageindex = GetIndex(lastimage)
                    If imageindex = FormDialog.FileDialog.ListCount - 1 Then
                        If MsgBox("This boxpart has been indexed. Do you want to reindex?", vbYesNo) = vbNo Then
                            ExitBoxPart GV.boxnumber, GV.boxpart
                            Exit Sub
                        Else
                            Main.cmdGoTo.Visible = True
                            Main.txtGoTo.Visible = True
                        End If
                    End If
                    FormDialog.FileDialog.Selected(imageindex) = True
                End If
                DisplayImage
                userBoxpartPhase GV.boxbarcode, GV.boxpart, GV.uname, GV.curr_phase
                EnableTextField True
                FetchData GV.boxnumber, GV.boxpart, GV.filename
                If Main.LabelIndexed.Caption <> "Indexed" Then
                    GV.start_imageindexed = GV.filename
                End If
                If GV.currentImage = 0 Then
                    Main.chkDelete.Value = 1
                    GV.skipped = True
                    Main.Next_Image_Click
                End If

                GV.open_beyond_image = 0
            Else
                If switch <> 0 Then
                    GV.boxnumber = ""
                    GV.boxpart = ""
                End If
                EnableTextField False
                GV.open_beyond_image = 0
            End If
End Sub

Public Sub parseBoxInformation(ByVal boxnumber As String)
Dim temp
    temp = Split(boxnumber, "-")
    GV.projectid = CStr(CLng(temp(0)))
    GV.projectName = getProjectName(CLng(GV.projectid))
    GV.boxbarcode = Mid$(boxnumber, 1, GV.boxbarcode_length)
    GV.boxid = getBoxId(GV.boxbarcode, 0, 0)
    GV.boxnumber = CStr(CLng(temp(2)))
    GV.boxpart = AddZero(temp(3), GV.partnumber_length)
End Sub

'**********************************'
' insert activity to pass_fail_log '
'**********************************'
Public Sub Pass_Fail_log(ByVal boxnumber As String, ByVal partnumber As String, ByVal sampler As String, ByVal indexer As String, ByVal Status As String)

On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim query As String
    
    Set conn = New ADODB.Connection
    conn.Open GV.DSN
    
    query = "insert into pass_fail_log values('" + boxnumber + "','" + partnumber + "','" + CStr(Date) + " " + Format(Time, "HH:MM:SS") + "','" + sampler + "','" + indexer + "','" + Status + "')"

    conn.Execute (query)
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check Pass_Fail_log"
        Exit Sub
    End If
End Sub

'*****************************************'
' create QA table  '
'**********************************'
Public Sub createQAtable(ByVal boxnumber As String, ByVal partnumber As String)
On Error GoTo ErrLine
    Dim conn As ADODB.Connection
    Dim query As String
    
    
    Set conn = New ADODB.Connection
    conn.Open GV.DSN
    query = "create table " + GV.box_table_name + "_QA select * from struct_QA"
    conn.Execute query

    
ErrLine:
    conn.Close
    Set conn = Nothing

End Sub

'*****************************************'
' clear out pass/fail information  '
'**********************************'
Public Sub clearout(ByVal boxnumber As String, ByVal partnumber As String)

    Dim conn As ADODB.Connection
    Dim query As String
    
    
    Set conn = New ADODB.Connection
    conn.Open GV.DSN
On Error GoTo createQA
    query = "select * from " + GV.box_table_name + "_QA"
    conn.Execute (query)
createQA:
    If Err.Number <> 0 Then
        query = "create table " + GV.box_table_name + "_QA select * from struct_QA"
        conn.Execute query
    End If
On Error GoTo ErrLine
    query = "update " + GV.box_table_name + "_QA set status='',failcnt=0 where partnumber='" + partnumber + "'"
    conn.Execute query
    
ErrLine:
    conn.Close
    Set conn = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description + ": check clearout"
        Exit Sub
    End If
End Sub

