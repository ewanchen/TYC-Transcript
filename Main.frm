VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.7#0"; "IMGEDIT.OCX"
Begin VB.Form Main 
   Caption         =   "TYC Transcripts Indexing QA Tool 1.0 - Neubus"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdOpenImage 
      Caption         =   "Open Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "Save Image"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRR 
      Caption         =   "Rotate Right"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      ToolTipText     =   "Rotate Current Image to Right and Automatically Save the Image"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdRL 
      Caption         =   "Rotate Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Rotate Current Image to Left and Automatically Save the Image"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "Zoom Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      ToolTipText     =   "Zoom Out Current Image"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdZoomIn 
      Caption         =   "Zoom In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Zoom In Current Image"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   12
      ToolTipText     =   "Exit Indexing Tool"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Open or Switch Boxes"
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame IndexFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9975
      Left            =   8520
      TabIndex        =   5
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtTYC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "GO TO"
         Height          =   495
         Left            =   1680
         TabIndex        =   31
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtGoTo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtFailNum 
         Height          =   315
         Left            =   5280
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optFail 
         Caption         =   "Fail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton optPass 
         Caption         =   "Pass"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   2
         TabIndex        =   3
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   100
         TabIndex        =   2
         Top             =   2640
         Width           =   6135
      End
      Begin VB.CheckBox chkMisfile 
         Caption         =   "Misfile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1440
         Width           =   6135
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Delete Page"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkFirst 
         Caption         =   "First Page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkBadScan 
         Caption         =   "Bad Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Bad Scan Check Box"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "TYC Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   33
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "eg. 10001aac"
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Middle Initial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   5295
      End
   End
   Begin VB.Frame ImgFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8295
      Begin ImgeditLibCtl.ImgEdit ImgEdit1 
         Height          =   9855
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   8055
         _Version        =   131077
         _ExtentX        =   14208
         _ExtentY        =   17383
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   536870910
         OcrZoneVisibility=   -3692
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3756
         sReserved2      =   -3756
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
         Begin VB.Frame Frame1 
            Height          =   495
            Left            =   1320
            TabIndex        =   22
            Top             =   0
            Width           =   5415
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "Warning: You are indexing in the training mode."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   120
               Width           =   5055
            End
         End
      End
   End
   Begin VB.Label LabelIndexed 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      ToolTipText     =   "Index Status"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label LabelImage 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   10320
      TabIndex        =   13
      ToolTipText     =   "Image Name"
      Top             =   0
      Width           =   3615
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Next_Image 
         Caption         =   "Next Image"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Previous_Image 
         Caption         =   "Previous Image"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Zoom_In 
         Caption         =   "Zoom In"
         Shortcut        =   ^A
      End
      Begin VB.Menu Zoom_Out 
         Caption         =   "Zoom Out"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Rotate_Right 
         Caption         =   "Rotate Right"
         Shortcut        =   ^R
      End
      Begin VB.Menu Save 
         Caption         =   "Save Data"
         Shortcut        =   ^S
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Reset 
         Caption         =   "Clear All Field"
         Shortcut        =   {F9}
      End
      Begin VB.Menu OpenImage 
         Caption         =   "Open Image"
         Enabled         =   0   'False
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu FirstPage 
         Caption         =   "First Page"
         Shortcut        =   ^F
      End
      Begin VB.Menu View_Full_Screen 
         Caption         =   "View Full Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu DeletePage 
         Caption         =   "Delete Page"
         Enabled         =   0   'False
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu Save_Image 
         Caption         =   "Save Image"
         Shortcut        =   {F12}
      End
      Begin VB.Menu Update 
         Caption         =   "Update"
         Enabled         =   0   'False
         Shortcut        =   ^U
         Visible         =   0   'False
      End
      Begin VB.Menu JumpFirst 
         Caption         =   "Jump To First Page"
         Shortcut        =   ^J
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Bad_Scan_Click()
    If Main.chkBadScan.Value = 0 Then
        Main.chkBadScan.Value = 1
    Else
        Main.chkBadScan.Value = 0
    End If
End Sub

Private Sub cmdGoTo_Click()
Dim imgname As String
Dim imgindex As Long
imgname = Main.txtGoTo.Text
FormDialog.FileDialog.path = GV.ImageFolder
imgindex = GetIndex(imgname + ".tif")
If imgindex = -1 Then
    MsgBox "invalid imagename"
    Exit Sub
End If
    
    FormDialog.FileDialog.Selected(imgindex) = True
    DisplayImage
    FetchData GV.boxnumber, GV.boxpart, GV.filename
End Sub

Private Sub cmdOpenImage_Click()
    Dim temp
    If Len(GV.boxnumber) = 0 Then
        MsgBox "You cannot use open image for the first time"
    Else
        'FormDialog.txtIndexer = GV.indexer
        FormDialog.DirFolder.Enabled = False
        temp = OpenImageFolder(GV.ImageFolder)
        FormDialog.Show
    End If
End Sub



Private Sub cmdSaveImage_Click()
    SaveImage
End Sub


Private Sub txtLastName_GotFocus()
    Main.txtLastName.BackColor = GV.FocusBColor
    Main.txtLastName.SelStart = 0
    Main.txtLastName.SelLength = Len(Main.txtLastName.Text)
End Sub

Private Sub txtLastName_LostFocus()
    Main.txtLastName.BackColor = GV.LostBColor
    Main.txtLastName.Text = UCase$(Main.txtLastName.Text)
End Sub
Private Sub txtfirstname_GotFocus()
    Main.txtFirstName.BackColor = GV.FocusBColor
    Main.txtFirstName.SelStart = 0
    Main.txtFirstName.SelLength = Len(Main.txtFirstName.Text)
End Sub

Private Sub txtfirstname_LostFocus()
    Main.txtFirstName.BackColor = GV.LostBColor
    Main.txtFirstName.Text = UCase$(Main.txtFirstName.Text)
End Sub
Private Sub txtMM_GotFocus()
    Main.txtMM.BackColor = GV.FocusBColor
    Main.txtMM.SelStart = 0
    Main.txtMM.SelLength = Len(Main.txtMM.Text)
End Sub

Private Sub txtMM_LostFocus()
    Main.txtMM.BackColor = GV.LostBColor
    Main.txtMM.Text = UCase$(Main.txtMM.Text)
End Sub
Private Sub txtTYC_GotFocus()
    Main.txtTYC.BackColor = GV.FocusBColor
    Main.txtTYC.SelStart = 0
    Main.txtTYC.SelLength = Len(Main.txtTYC.Text)
End Sub

Private Sub txtTYC_LostFocus()
    Main.txtTYC.BackColor = GV.LostBColor
    Main.txtTYC.Text = UCase$(Main.txtTYC.Text)
End Sub


Private Sub DeletePage_Click()
    If Main.chkDelete.Value = 0 Then
        Main.chkDelete.Value = 1
    Else
        Main.chkDelete.Value = 0
    End If
End Sub

Private Sub FirstPage_Click()
    If Main.chkFirst.Value = 0 Then
        Main.chkFirst.Value = 1
    Else
        Main.chkFirst.Value = 0
    End If
End Sub

Private Sub Form_Load()
    GV.masterDSN = "MASTERDSN"
    Dim purpose
    If MsgBox("Are you going to use the tool for training?", vbYesNo) = vbYes Then
            GV.trainFlag = True
            GV.Tool = "TYC Transcripts Indexing Training Tool"
            GV.DSN = "TYCTRANSCRIPTSTrainDSN"
            purpose = "Train"
            GV.tool_purpose = "Train"
            Main.BackColor = &HFF&
            Main.Frame1.Visible = True
    Else
            GV.trainFlag = False
            GV.Tool = "TYC Transcripts Indexing Tool"  'Tool Name'
            GV.DSN = "TYCTRANSCRIPTSDSN"     'Database Name'
            purpose = "Production"
            GV.tool_purpose = "Production"
            Main.Frame1.Visible = False
    End If
    testDSN
    GV.job = "Indexing"
    
    GV.FocusBColor = &HFFFF80       'got focus background color'
    GV.LostBColor = &H80000005      'lost focus background color'

    EnableTextField False
End Sub


Private Sub cmdExit_Click()
    Unload Me
    Unload FormDialog
    Unload frmHelp
    Unload FormLogin
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MsgBox("Are you sure you want to exit this application?", vbYesNo + vbCritical) = vbYes Then
            ExitBoxPart GV.boxbarcode, GV.boxpart
            Unload Me
        Else
            Cancel = True
        End If
End Sub
Private Sub cmdOpen_Click()
    Dim temp

        GV.end_flag = 0
        GV.from_view_next = 1
        GV.OpenImage = False
        Unload FormDialog
'        temp = OpenImageFolder(GV.imagepath)
        FormDialog.Show
'    End If
End Sub

Private Sub cmdRL_Click()
    RotateLeftImage
End Sub

Private Sub cmdRR_Click()
    RotateRightImage
End Sub

Private Sub cmdZoomIn_Click()
    ZoomInImage
End Sub

Private Sub cmdZoomOut_Click()
    ZoomOutImage
End Sub



Private Sub JumpFirst_Click()

If MsgBox("Are you sure you want to go to first page?", vbYesNo) = vbYes Then
    FormDialog.FileDialog.path = GV.ImageFolder
    FormDialog.FileDialog.Selected(0) = True
    DisplayImage
    FetchData GV.boxnumber, GV.boxpart, GV.filename
End If
End Sub

Public Sub Next_Image_Click()

If GV.fullscreen = True Then
    viewmode
End If


NextImage_toDo
GV.skipped = False

If (FileLen(GV.filepath) < GV.blank_page_threshold And GV.currentImage Mod 2 = 1 And GV.end_flag = 0) Or GV.currentImage = 1 Then
    If Main.LabelIndexed.Caption <> "Indexed" Then
        Main.chkDelete.Value = 1
        Main.chkFirst.Value = 0
    End If
    GV.skipped = True
    NextImage_toDo
End If

End Sub


Private Sub OpenImage_Click()
    Dim temp
    If Len(GV.boxnumber) = 0 Then
        MsgBox "You cannot use open image for the first time"
    Else
        'FormDialog.txtIndexer = GV.uname
        FormDialog.DirFolder.Enabled = False
        temp = OpenImageFolder(GV.ImageFolder)
        FormDialog.Show
    End If
End Sub

Private Sub Paste_Click()
    PasteData data_list
End Sub





Private Sub Previous_Image_Click()
If GV.fullscreen = True Then
    viewmode
End If
    PreviousImage_toDo
If Main.chkDelete.Value = 1 Then
    PreviousImage_toDo
End If
End Sub

Private Sub Reset_Click()
    resetFields
End Sub

Private Sub Rotate_Left_Click()
    RotateLeftImage
End Sub

Private Sub Rotate_Right_Click()
    RotateRightImage
End Sub



Private Sub Save_Click()
    
    GV.end_date = Date
    GV.end_time = Format(Time, "HH:MM:SS")
    If MsgBox("This won't update all the pages!  Are you sure you only want to save this page's information?", vbYesNo) = vbYes Then
        ConvertToUpper
        If chkDelete.Value = 0 And chkBadScan.Value = 0 Then
            If error_checking() = False Then
                Exit Sub
            End If
        End If
        
        DocTypeInsert GV.boxnumber, GV.boxpart, GV.filename
        '*** Recording Activity ***'
        If pre_page = False Then
            Activity_Log GV.filename, "Save and Commit"
        Else
            Activity_Log GV.filename, "Reverse"
        End If
    End If
    GV.start_date = Date
    GV.start_time = Format(Time, "HH:MM:SS")
End Sub

Private Sub Save_Image_Click()
    SaveImage
End Sub


Private Sub View_Full_Screen_Click()
    viewmode
End Sub

Private Sub Zoom_In_Click()
    ZoomInImage
End Sub

Private Sub Zoom_Out_Click()
    ZoomOutImage
End Sub


