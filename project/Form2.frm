VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   -120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password:"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
If Text1.Text = Text2.Text Then
    Form1.Visible = True
    Unload Form2
Else
    If Form1.Check10.Value = Checked Then
        Form1.Timer2.Enabled = True
        Form2.Visible = False
        Unload Form2
    Else
        End
    End If
End If

End Sub

Private Sub Command2_Click()
Form2.Visible = False
Unload Form2
Form1.Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Call loadPass
If Form1.Check10 = "0" Then
    If Form5.Check2.Value = "1" Then
        this = True
        Form1.Check10.Value = Checked
        Form2.Hide
        Form1.Timer2.Enabled = True
        Call Running
    End If

End If

If Text2.Text = "" Then
    
    Form2.Hide
    Unload Form2
    Form4.Show
    Load Form4
    Form4.Text1.Enabled = False
    Form4.Label1.Enabled = False

Else
    
    Dim decrypt
    For d = 1 To 5
        decrypt = DeCode(Text2.Text)
        Text2.Text = decrypt
    Next d
    
End If


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Command1_Click
End If
End Sub

