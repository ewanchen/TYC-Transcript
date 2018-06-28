VERSION 5.00
Begin VB.Form IndexerDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1905
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtboxnum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Box Number"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox ListPart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "List of Box Part Number"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ListBox ListFolder 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "list of box status"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtuser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Indexer's Initial"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Box Status"
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
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Part Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Box Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "IndexerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim temp
    If Len(ListFolder.Text) = 0 Then
        MsgBox "Please select the folder"
        Exit Sub
    End If
    If Len(txtuser.Text) = 0 Then
        MsgBox "Please enter your user name"
        Exit Sub
    End If
    
    If Len(txtboxnum.Text) = 0 Then
        MsgBox "Please enter the box number"
        Exit Sub
    End If
    If Len(ListPart.Text) = 0 Then
        MsgBox "Please select the part number"
        Exit Sub
    End If
    GV.uname = UCase(txtuser.Text)
    GV.boxnumber = UCase(txtboxnum.Text)
    GV.boxpart = ListPart.Text
    If StrComp(ListFolder.Text, "Indexed") = 0 Then
        GV.imageLocation = GV.imagepath + "input\" + ListFolder.Text
    Else
        GV.imageLocation = GV.imagepath + "input"
    End If
    GV.ImageFolder = GV.imageLocation + "\" + GV.boxnumber + "\" + GV.boxnumber + GV.boxpart
    temp = OpenImageFolder(GV.ImageFolder)
    If temp = 1 Then
        FormDialog.Show
    Else
        GV.indexer = ""
        GV.boxnumber = ""
        GV.boxpart = ""
        GV.ImageFolder = ""
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    ListFolder.AddItem "Not Index"
    ListFolder.AddItem "Indexed"
End Sub
Private Sub ListFolder_LostFocus()
    If Len(txtboxnum) <> 0 Then
        Call txtboxnum_LostFocus
    End If
End Sub
Private Sub txtboxnum_LostFocus()
    Dim found
    If StrComp(ListFolder.Text, "Indexed") = 0 Then
        GV.imageLocation = GV.imagepath + "input\" + ListFolder.Text
    Else
        GV.imageLocation = GV.imagepath + "input"
    End If
    found = Dir(GV.imageLocation + "\" + txtboxnum.Text, vbDirectory)
    ListPart.Clear
    If Len(found) = 0 Then
        MsgBox GV.imageLocation + "\" + txtboxnum.Text + " does not exist. Try to open another box"
    Else
        FormDialog.DirFolder.path = GV.imageLocation + "\" + txtboxnum.Text
    
        Dim i As Integer
        Dim temp As Integer
        Dim partnumber
        Dim count As Integer
        Dim path As String
        i = 0
        temp = FormDialog.DirFolder.ListCount
        partnumber = Split(FormDialog.DirFolder.list(0), "\")
        count = UBound(partnumber)
        While i < temp
            partnumber = Split(FormDialog.DirFolder.list(i), "\")
            ListPart.AddItem Mid(partnumber(count), 6, 2)
            i = i + 1
        Wend
    End If
    Unload FormDialog
End Sub
