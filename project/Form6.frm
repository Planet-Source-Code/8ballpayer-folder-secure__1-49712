VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Security"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Enter correct password to gain entry:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PROTECTED FOLDER"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_DISABLED = &H2&

Public Sub DisableX(Frm As Form)
Dim hMenu As Long
Dim nCount As Long
  hMenu = GetSystemMenu(Frm.hwnd, 0)
  nCount = GetMenuItemCount(hMenu)
  Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
  DrawMenuBar Frm.hwnd
End Sub

Private Sub Command1_Click()

Dim result As Long

If Text1.Text = Text2.Text Then
    SetOnTop Form6.hwnd, False
    DisableTrap Form6
    Form6.Hide
    Unload Form6
    plus = plus + 1
    Form1.List2.RemoveItem remove
    Form1.Timer1.Enabled = True


ElseIf Text1.Text = "" Then
    SetOnTop Form6.hwnd, False
    DisableTrap Form6
    result = PostMessage(lProcessID, WM_CLOSE, 0, 0)
    Form6.Hide
    Unload Form6
    Form1.Timer1.Enabled = True

Else
    SetOnTop Form6.hwnd, False
    DisableTrap Form6
    result = PostMessage(lProcessID, WM_CLOSE, 0, 0)
    Form6.Hide
    Unload Form6
    Form1.Timer1.Enabled = True
End If
End Sub


Private Sub Form_Load()
Dim result As Long
Dim d As Integer
Call loadPass2
DisableX Form6


Form1.Timer1.Enabled = False
SetOnTop Form6.hwnd, True
EnableTrap Form6
If Text2.Text <> "" Then
Dim decrypt
    For d = 1 To 5
        decrypt = DeCode(Text2.Text)
        Text2.Text = decrypt
    Next d
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetOnTop Form6.hwnd, True
EnableTrap Form6
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Command1_Click
End If
End Sub
