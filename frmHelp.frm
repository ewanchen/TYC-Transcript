VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help - Neubus"
   ClientHeight    =   5865
   ClientLeft      =   2595
   ClientTop       =   2505
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'private for username
Private Const MAX_USERNAME As Long = 20

Private Declare Function GetUserName Lib "advapi32" _
   Alias "GetUserNameA" _
  (ByVal lpBuffer As String, _
   nSize As Long) As Long
      
Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long
  
  
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmHelp.Label1.Caption = "Program Name: TXSTATE REGISTRAR Indexing and QA Tool" + vbCrLf + vbCrLf
frmHelp.Label1.Caption = frmHelp.Label1.Caption + "Program Version: v1.9" + vbCrLf + vbCrLf
frmHelp.Label1.Caption = frmHelp.Label1.Caption + "Mysql server: <no username>@prod" + vbCrLf + vbCrLf
frmHelp.Label1.Caption = frmHelp.Label1.Caption + "DSN: MASTERDSN" + vbCrLf + vbCrLf
frmHelp.Label1.Caption = frmHelp.Label1.Caption + "DSN: TXSTATEDSN" + vbCrLf + vbCrLf
frmHelp.Label1.Caption = frmHelp.Label1.Caption + "Windows username: " + GetThreadUserName
End Sub

Private Function GetThreadUserName() As String

  'Retrieves the user name of the current
  'thread. This is the name of the user
  'currently logged onto the system. If
  'the current thread is impersonating
  'another client, GetUserName returns
  'the user name of the client that the
  'thread is impersonating.
   Dim buff As String
   Dim nSize As Long
   
   buff = Space$(MAX_USERNAME)
   nSize = Len(buff)

   If GetUserName(buff, nSize) = 1 Then

      GetThreadUserName = TrimNull(buff)
      Exit Function

   End If

End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function

