VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "TYC Transcripts Indexing Tool V 1.0 Login Page"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optIndexSampling 
      Caption         =   "Index Sampling"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton optIndexing 
      Caption         =   "Indexing"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox password 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox username 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Please select one of the project"
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
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please enter your username and password"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Menu help 
      Caption         =   "help"
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim temp
    Dim checkResult
    
    If Len(FormLogin.username.Text) = 0 Then
        MsgBox "Please enter username!"
        Exit Sub
    Else
        GV.uname = FormLogin.username.Text
    End If
    
    If Len(FormLogin.password.Text) = 0 Then
        MsgBox "Please enter password!"
        Exit Sub
    Else
        GV.pword = FormLogin.password.Text
    End If
    
    If FormLogin.optIndexing.Value = False And FormLogin.optIndexSampling.Value = False Then
        MsgBox "Please choose task Indexing/Index Sampling"
        Exit Sub
    Else
        If FormLogin.optIndexing.Value = True Then
            GV.curr_phase = "INDEXING"
        Else
            GV.curr_phase = "INDEX_SAMPLING"
        End If
    End If

    '*** PCSII: Get project ID ***'
    GV.projectid = getProjectId(GV.projectName)
    If StrComp(GV.projectid, "FAIL") = 0 Then
        MsgBox "Project " + GV.projectName + " does not exist in the PCS project list!! Please contact administrator!"
        Exit Sub
    End If
    
    '*** PCSII: Authenticate user ***'
    checkResult = authenticateUser(GV.uname, GV.pword, "", GV.projectid, GV.curr_phase)
    If StrComp(checkResult, "FAIL") = 0 Or StrComp(Mid(checkResult, 1, 4), "FAIL") = 0 Then
        MsgBox "You are not authorize for using this tool!! Please check your username and password!"
        Exit Sub
    Else
        GV.userid = CLng(checkResult)
    End If
    GetLocation
    
    GV.ToolVersion = "1.0"          'Tool version'
    If GV.ToolVersion <> GV.version Then
        MsgBox "You're opening version " + GV.ToolVersion + ". Latest version is " + GV.version + ". Please contact IT support. Thank you"
        Unload Me
        Exit Sub
    End If
        
        Main.Show
        Unload Me
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GV.masterDSN = "MASTERDSN"
    GV.projectName = "Texas Youth Commission - Transcripts"
    getPCSurl
    'getProjectList
End Sub

Private Sub help_Click()
    frmHelp.Show
End Sub
