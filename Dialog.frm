VERSION 5.00
Begin VB.Form FormDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Box List"
   ClientHeight    =   6360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   5520
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.FileListBox samplefile 
      Height          =   480
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.FileListBox OutputFile 
      Height          =   480
      Left            =   8280
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.DirListBox OutputDir 
      Height          =   540
      Left            =   7200
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   4320
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.DirListBox Dir2 
      Height          =   765
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "List all boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.FileListBox FileDialog 
      Height          =   4965
      Left            =   8520
      Pattern         =   "*.TIF"
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.DirListBox DirFolder 
      Height          =   1440
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Box+Part List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FormDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub DirFolder_Change()
'    FileDialog.path = DirFolder
    Dim found
    Dim dirpath As String
    Dim boxpart
    
    boxpart = Split(FormDialog.List1.Text, "-")
    
If UBound(boxpart) >= 0 Then
    If GV.OpenImage = True Then
        dirpath = GV.imagepath + "\" + Mid(boxpart(0), 1, Len(boxpart(0)) - 3) + "\" + boxpart(0)
        found = Dir(dirpath, vbDirectory)
        If Len(found) = 0 Then
            MsgBox "Image path " + dirpath + " does not exist!! Cannot open image!"
            Exit Sub
        End If
        FileDialog.path = GV.imagepath + "\" + Mid(boxpart(0), 1, Len(boxpart(0)) - 3) + "\" + boxpart(0)
    Else
        If Len(List1.Text) = 0 Then
            dirpath = GV.imagepath + "\" + FormDialog.List1
            found = Dir(dirpath, vbDirectory)
            If Len(found) = 0 Then
                MsgBox "Image path " + dirpath + " does not exist!! Cannot open image!"
                Exit Sub
            End If
            FileDialog.path = GV.imagepath + "\" + FormDialog.List1
        Else
            dirpath = GV.imagepath + "\" + Mid(boxpart(0), 1, Len(boxpart(0)) - 3) + "\" + boxpart(0)
            found = Dir(dirpath, vbDirectory)
            If Len(found) = 0 Then
                MsgBox "Image path " + dirpath + " does not exist!! Cannot open image!"
                Exit Sub
            End If
            FileDialog.path = GV.imagepath + "\" + Mid(boxpart(0), 1, Len(boxpart(0)) - 3) + "\" + boxpart(0)
        End If
    End If
End If
End Sub

Private Sub List1_Click()
    Dim dirpath As String
    Dim found
    Dim boxpart

    parseBoxInformation FormDialog.List1.Text
    
    If GV.OpenImage = False Then
        dirpath = GV.imagepath + "\" + GV.boxbarcode + "\" + GV.boxnumber + AddZero(GV.boxpart, GV.partnumber_length)
        found = Dir(dirpath, vbDirectory)
        If Len(found) = 0 Then
            MsgBox "Image path " + dirpath + " does not exist!! Cannot open image!"
            Exit Sub
        End If
        FormDialog.FileDialog.path = GV.imagepath + "\" + GV.boxbarcode + "\" + GV.boxnumber + AddZero(GV.boxpart, GV.partnumber_length)

    End If
End Sub



Private Sub List2_Click()
    GV.current_index_page = List2.ListIndex
End Sub

Private Sub OKButton_Click()

    Dim vTime As SYSTEMTIME

    

    '*** PCSII: Get project Name ***'
    
        openBox_todo
            
        
        GV.start_date = Date
        GetLocalTime vTime
        GV.start_time = Format(vTime.vHour) & ":" & Format(vTime.vMinute) & ":" & Format(vTime.vSecond) & ":" & Format(vTime.vMilliseconds)

    Unload Me
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    Dim temp
    Dim temp2
    Dim boxcount As Integer
    Dim id_list
    Dim list As String
    
    If GV.from_view_next = 1 Then
        GV.from_view_next = 0
        If GV.tool_purpose = "Train" Or GV.tool_purpose = "Production" Then
            '*** PCSII: Get box and part list ready for index ***'
            temp = getPartsForPhase(GV.curr_phase, GV.projectid)
            If FormDialog.List1.ListCount = 0 Then
                MsgBox "No box list found in PCS"
                Exit Sub
            End If
            'FileDialog.Visible = True
            Label1.Visible = False
        End If
    End If
End Sub

Private Sub Option2_Click()
    Dim temp
    
    If Option2.Value = True Then
        '*** PCSII: Get full list of boxpart from PCS ***'
        temp = getPartsForproject(GV.projectid)
        If StrComp(temp, "FAIL") = 0 Then
            MsgBox "No box list found in the PCS!!"
            Exit Sub
        End If
    End If
End Sub
