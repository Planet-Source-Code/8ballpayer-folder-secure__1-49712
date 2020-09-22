VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Other Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   4815
      Begin VB.CheckBox Check2 
         Caption         =   "Auto start Folder Security on system start up"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Encrypt Folder Log"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Password Options"
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Change App Password"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Folder Password Entry"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Again:"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Folder Entry Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton Option2 
         Caption         =   "Windows XP"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Windows 2000"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    
Private Sub Check1_Click()
If Check1.Value = Checked Then
    If folderpassentry = False Then
        MsgBox "Remember, you must click Save Changes for Folder Password Entry to work!", vbExclamation, "Folder Password Entry"
    End If
    Text1.Enabled = True
    Text2.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
    folderpassentry = True
Else
    Text1.Enabled = False
    Text2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False
    folderpassentry = False
    On Error GoTo error
    Kill App.Path & "\pass2.doc"
End If
error:
End Sub



Private Sub Check2_Click()
Dim strProgramPath   As String   ' The path of the executable file
Dim strGroup         As String
Dim strProgramIconTitle As String
Dim strProgramArgs   As String
Dim sParent          As String
If Check2.Value = Checked Then
    autostart = True
    If sh = False Then
        MsgBox "This also hides Folder Secure everytime it is started!", vbExclamation, "Warning"
    End If
    strProgramPath = App.Path & "\Folder Secure.exe"
    strGroup = "..\..\Start Menu\Programs\Startup"
    strProgramIconTitle = "Folder Secure.exe"
    strProgramArgs = ""
    sParent = "$(Programs)"
    CreateShellLink strProgramPath, strGroup, strProgramArgs, strProgramIconTitle, True, sParent
    
Else
    autostart = False
    sh = False
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = Checked Then
    encryptlog = True
Else
    encryptlog = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = Checked Then
    encryptlist = True
Else
    encryptlist = False
End If
End Sub

Private Sub Command1_Click()
Load Form4
Form4.Show
End Sub

Private Sub Command2_Click()
Form5.Hide
End Sub

Private Sub Command3_Click()
MsgBox "This also saves options on from main window! Continue?", vbInformation + vbYesNo, "Folder Secure"
If VbMsgBoxResult.vbYes = True Then
    
    If Check1.Value = Checked Then
        If Text1.Text <> "" Then
            If Text1.Text = Text2.Text Then
                Call savePass2
            Else
                MsgBox "Passwords do not match!  Folder Password Entry was not enabled!", vbExclamation, "Password Error"
                Exit Sub
            End If
        Else
            MsgBox "You did not enter a password!  Folder Password Entry was not enabled!", vbExclamation, "Folder Password Entry"
        End If
    End If
    Call saveConfig
    MsgBox "Changes saved successfully!", vbOKOnly, "Changes Saved"
End If
End Sub

Private Sub Form_Load()
If folderpassentry = False Then
    Label1.Enabled = False
    Label2.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
End If

End Sub

Private Sub Option1_Click()
winxp = False
End Sub

Private Sub Option2_Click()
winxp = True
End Sub
