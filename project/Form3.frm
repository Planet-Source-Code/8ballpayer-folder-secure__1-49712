VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Log"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Log"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   5175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call saveLog
MsgBox "Log saved successfully!", vbInformation, "Folder Secure"
End Sub

Private Sub Command2_Click()
Form3.Hide
End Sub

Private Sub Command3_Click()
MsgBox "Are you sure you want to start a new log?", vbYesNo, "FolderSecure"
If VbMsgBoxResult.vbYes Then
    Text1.Text = ""
    Text2.Text = ""
    Call saveLog
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Do you want to save the Folder Log?", vbYesNo, "Folder Secure"
If VbMsgBoxResult.vbYes Then
    Call saveLog
End If


End Sub

Private Sub Text2_Change()
Text1.Text = Text2.Text
End Sub
