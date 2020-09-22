VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Settings"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2925
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Again:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Please Enter Your New Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter Your Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim makepass As Boolean

Private Sub Command1_Click()
If Text2.Text = "" Then
    MsgBox ("Please Enter Passwords")
    Exit Sub
End If
If makepass = True Then
    If Text2.Text = Text3.Text Then
        
        Call savePass
        MsgBox "Password Created Successfully", vbOKOnly, "Folder Secure"
        Load Form1
        Form1.Show
        Form4.Hide
        Unload Form4
        
    Else
        MsgBox "Passwords Do Not Match!", vbCritical, "Folder Secure"
        Text2.Text = ""
        Text3.Text = ""
    End If
Else
    If Text1.Text = Text4.Text Then
        If Text2.Text = Text3.Text Then
            
            Call savePass
            MsgBox "New Password Created Successfully", vbOKOnly, "Folder Secure"
            Load Form1
            Form1.Show
            Form4.Hide
            Unload Form4
            
        Else
            MsgBox "Passwords Do Not Match!", vbCritical, "Folder Secure"
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
        End If
    Else
        MsgBox "Original Password Incorrect!", vbCritical, "Folder Secure"
    End If
End If
End Sub

Private Sub Command2_Click()
Text1.Enabled = True
Label1.Enabled = True
Form4.Hide
Unload Form4
End Sub

Private Sub Form_Load()
Call loadPass
makepass = True
If Text4.Text = "" Then

    Form4.Text1.Enabled = False
    Form4.Label1.Enabled = False

Else
    
    Dim decrypt
    For d = 1 To 5
        decrypt = DeCode(Text4.Text)
        Text4.Text = decrypt
    Next d
    
    Form4.Text1.Enabled = True
    Form4.Label1.Enabled = True
    makepass = False
    
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Command1_Click
End If
End Sub
