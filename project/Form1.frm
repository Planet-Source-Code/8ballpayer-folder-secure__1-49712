VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Secure (Not Running)"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      Height          =   1425
      Left            =   1560
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   120
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2760
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   720
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   120
      TabIndex        =   36
      Top             =   1440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Check10"
      Height          =   375
      Left            =   1200
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   720
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save List"
      Height          =   375
      Left            =   1080
      TabIndex        =   29
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   5175
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command4 
         Caption         =   "More Options"
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Disable Taskbar"
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Disable Ctrl Alt Del"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Text            =   "Alert"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "F8"
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   2040
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "F7"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   2040
         Width           =   615
      End
      Begin VB.OptionButton Option6 
         Caption         =   "F6"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   2040
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "F5"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide Program"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Auto Hide Program on Security Start"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Auto Save Security List on Update"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         Caption         =   "Log Options"
         Height          =   1335
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   5055
         Begin VB.CheckBox Check5 
            Caption         =   "Enable Folder Log"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Record All Folders Opened"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Record All Secured Folder Open Attempts"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "View Folder Log"
            Height          =   375
            Left            =   3480
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Minute(s)"
            Height          =   255
            Left            =   1920
            TabIndex        =   38
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Update Every"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make Invisible to Ctrl Alt Del"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtalertmsg 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "Access Denied"
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aggressiveness"
         Height          =   1935
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton Option4 
            Caption         =   "Super Aggressive"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Pretty Aggressive"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Normal"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Not Very"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Alert Title:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Key Pressed To Make Program Visible:"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Alert Message:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   720
   End
   Begin VB.CommandButton cmdretrieve 
      Caption         =   "Start Security"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Security List:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vKey As String
Dim DisableCAD As Boolean
Dim DisableTask As Boolean
Dim Ready As Boolean
Dim text3int As Integer
Dim Detect As Long
Dim timerint As Double


Private DTaskbar1 As Long
Private DTaskbar2 As Long
Private DTaskbar3 As Long
Dim thehwnd As Long
Dim FName As String
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Function GetTrayHandle(mType As Integer) As Long
    Dim Desktop As Long, mhandle As Long, temp As String * 16, SrchString As String
    Select Case mType
        Case 1
            SrchString = "Button"
        Case 2
            SrchString = "TrayNotifyWnd"
        Case 3
            SrchString = "ReBarWindow32"
    End Select
    Desktop = GetDesktopWindow()
    mhandle = GetWindow(Desktop, GW_CHILD)
    Do While mhandle <> 0
        GetClassName mhandle, temp, 14
        If Left(temp, 13) = "Shell_TrayWnd" Then
            If mType = 4 Then
                GetTrayHandle = mhandle
                Exit Do
            End If
            mhandle = GetWindow(mhandle, GW_CHILD)
            Do While mhandle <> 0
                GetClassName mhandle, temp, Len(SrchString) + 1
                If Left(temp, Len(SrchString)) = SrchString Then
                    GetTrayHandle = mhandle
                    Exit Function
                End If
                mhandle = GetWindow(mhandle, GW_NEXT)
            Loop
        End If
        mhandle = GetWindow(mhandle, GW_NEXT)
    Loop
End Function


Private Sub Check1_Click()

If Check1.Value = Checked Then
    App.TaskVisible = False
    itm = True
Else
    App.TaskVisible = True
    itm = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = Checked Then
    rsf = True
Else
    rsf = False
End If
End Sub

Private Sub Check3_Click()
If Check3 = "1" Then
    Label6.Enabled = True
    Label7.Enabled = True
    Text3.Enabled = True
    Timer3.Enabled = True
    ra = True
Else
    Label6.Enabled = False
    Label7.Enabled = False
    Text3.Enabled = False
    Timer3.Enabled = False
    ra = False
End If
End Sub

Private Sub Check5_Click()
If Check5 = "1" Then
    Check2.Enabled = True
    Check3.Enabled = True
    Label6.Enabled = True
    Label7.Enabled = True
    Text3.Enabled = True
    Text3.Text = "1"
    Timer3.Enabled = True
    el = True
Else
    Check2.Enabled = False
    Check3.Enabled = False
    Label6.Enabled = False
    Label7.Enabled = False
    Text3.Enabled = False
    Timer3.Enabled = False
    el = False
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = Checked Then
    If Me.Caption <> "Folder Secure (Running)" Then
        MsgBox "Task Manager will be disabled when security starts", vbInformation, "Folder Secure"
        MsgBox "Warning: May not work on all systems!", vbExclamation, "WARNING"
        DisableCAD = True
        dtm = True
    Else
        DisableCAD = True
        dtm = True
    End If
Else
    SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    DisableCAD = False
    dtm = False
End If
    
End Sub

Private Sub Check9_Click()
If Check9.Value = Checked Then
    If Me.Caption <> "Folder Secure (Running)" Then
        MsgBox "Taskbar will be disabled when security starts", vbInformation, "Folder Secure"
        DisableTask = True
        dtb = True
    Else
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_HIDEWINDOW
        DisableTask = True
        dtb = True
    End If
Else
    SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    DisableTask = False
    dtb = False
End If
    
End Sub

Private Sub cmdretrieve_Click()
Dim DTask1 As Long
Dim DTask2 As Long
Dim DTask3 As Long
Dim Visible As Boolean

If cmdretrieve.Caption = "Start Security" Then
    cmdretrieve.Caption = "Stop Security"
    Call Running
    
        
Else
    cmdretrieve.Caption = "Start Security"
    If DisableTask = True Then
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    End If
    Timer1.Enabled = False
    Me.Caption = "Folder Secure (Not Running)"
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "Please enter a file name", vbCritical, "Folder Secure"
Else
    List1.AddItem Text1.Text
    List2.AddItem Text1.Text
    If Check6.Value = Checked Then
        Call saveFiles
    End If
End If
Text1.Text = ""
End Sub

Private Sub Command2_Click()
List1.RemoveItem List1.ListIndex
On Error Resume Next
List2.RemoveItem List1.ListIndex + plus
If Check6.Value = Checked Then
    Call saveFiles
End If
End Sub

Private Sub Command3_Click()
Load Form3
Form3.Show
End Sub



Private Sub Command4_Click()
Load Form5
Form5.Show
End Sub

Private Sub Command5_Click()
Dim key As String

If vKey = "116" Then
    key = "F5"
ElseIf vKey = "117" Then
    key = "F6"
ElseIf vKey = "118" Then
    key = "F7"
ElseIf vKey = "119" Then
    key = "F8"
Else
    key = vKey
End If

MsgBox ("Remember your password and remember to press " & key & " to make the program visible again"), vbOKOnly, "Hide Program"
Timer2.Enabled = True
Me.Visible = False
End Sub

Private Sub Command6_Click()

Call saveFiles
MsgBox "List Saved Successfully", vbOKOnly, "Folder Secure"


End Sub

Private Sub Form_Load()
dtb = False
dtm = False
itm = False
el = False
rsf = False
ra = False
folderpassentry = False
autostart = False
encryptlog = False
Form5.Text1.Enabled = False
Form5.Text2.Enabled = False
Form5.Label1.Enabled = False
Form5.Label2.Enabled = False
plus = 0

Call loadFiles
Call loadConfig
Call loadLog

If List3.ListCount = 0 Then

Else
    List3.ListIndex = 0
    If List3.Text = "true" Then
        winxp = True
        Form5.Option2 = True
    Else
        winxp = False
        Form5.Option1 = True
    End If
    List3.ListIndex = 1
    If List3.Text = "true" Then
        folderpassentry = True
        Form5.Check1.Value = Checked
        Form5.Label1.Enabled = True
        Form5.Label2.Enabled = True
        Form5.Text1.Enabled = True
        Form5.Text2.Enabled = True
    Else
        folderpassentry = False
    End If
    List3.ListIndex = 2
    If List3.Text = "true" Then
        autostart = True
        sh = True
        Form5.Check2.Value = Checked
        
    Else
        autostart = False
    End If
    List3.ListIndex = 3
    If List3.Text = "true" Then
        encryptlog = True
        Form5.Check3.Value = Checked
    Else
        encryptlog = False
    End If
    On Error GoTo error
    List3.ListIndex = 4
    If List3.Text = "true" Then
        Check9.Value = Checked
    End If
    List3.ListIndex = 5
    If List3.Text = "true" Then
        Check8.Value = Checked
    End If
    List3.ListIndex = 6
    If List3.Text = "true" Then
        Check1.Value = Checked
    End If
    List3.ListIndex = 7
    If List3.Text = "true" Then
        Check5.Value = Checked
    End If
    List3.ListIndex = 8
    If List3.Text = "true" Then
        Check2.Value = Checked
    End If
    List3.ListIndex = 9
    If List3.Text = "true" Then
        Check3.Value = Checked
    End If
    List3.ListIndex = 10
    Text3.Text = List3.Text
    List3.ListIndex = 11
    Text2.Text = List3.Text
    List3.ListIndex = 12
    txtalertmsg.Text = List3.Text
error:
End If

    
DoEvents



    Dim decrypt
    Dim d
    For d = 1 To 5
        decrypt = DeCode(Form3.Text2.Text)
        Form3.Text2.Text = decrypt
    Next d
    Form3.Text1.Text = Form3.Text2.Text


Option5.Value = True
Text3.Text = "1"
text3int = Text3.Text
Timer3.Interval = text3int * 60000
Timer3.Enabled = True
Check5 = "1"
DisableCAD = False
Ready = True
vKey = "F5"
Label6.Enabled = False
Label7.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label8_Click()

End Sub

Private Sub List1_Click()
On Error Resume Next
List2.ListIndex = List1.ListIndex + plus
End Sub

Private Sub Option5_Click()
vKey = "F5"
End Sub

Private Sub Option6_Click()
vKey = "F6"
End Sub

Private Sub Option7_Click()
vKey = "F7"
End Sub

Private Sub Option8_Click()
vKey = "F8"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Command1_Click
End If
End Sub

Private Sub Text3_Change()

On Error GoTo error
timerint = Text3.Text * 10000
Timer3.Interval = timerint
Exit Sub

error:
        MsgBox "TIMER ERROR: INVALID INTERVAL " + "'" + Text3.Text + "'", vbCritical, "ERROR"

    Timer3.Enabled = False
    Check3 = 0
    Label6.Enabled = False
    Label7.Enabled = False
    Text3.Enabled = False
    Text3.Text = ""
    

End Sub

Private Sub Timer1_Timer()
Dim result As Long
Dim result2 As Long
Dim findfolder As Long



If List2.ListCount <> "0" Then

    If List2.ListIndex = "-1" Then
        List2.ListIndex = "0"
    End If
    
    If winxp = True Then
        lProcessID = FindWindow("CabinetWClass", List2.Text)
    Else
        lProcessID = FindWindow("ExploreWClass", List2.Text)
    End If
    
    If lProcessID <> "0" Then
        
        If folderpassentry = True Then
            folder = List2.Text
            remove = List2.ListIndex
            Form6.Show
        Else
            result = PostMessage(lProcessID, WM_CLOSE, 0, 0)
            MsgBox (txtalertmsg.Text), vbCritical, Text2.Text
        End If
        If Check5 = 1 Then
            If Check2 = 1 Then

                Form3.Text2.Text = Form3.Text2.Text + vbNewLine + "FOLDER:  " + List2.Text + "      OPEN ATTEMPT      " + " ON:  " + Date$ + "   AT:  " + Time$

            End If
        End If

    End If
    
    
    If List2.ListIndex = List2.ListCount - 1 Then
        List2.ListIndex = "0"
    Else
        List2.ListIndex = List2.ListIndex + 1
    End If

End If
lProcessID2 = FindWindow("#32770", "Windows Task Manager")
If DisableCAD = True Then
    If lProcessID2 <> "0" Then
        result2 = PostMessage(lProcessID2, WM_CLOSE, 0, 0)
    End If
End If
If DisableTask = True Then
    SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End If


End Sub

Private Sub Timer2_Timer()

If vKey = "F5" Then
    vKey = vbKeyF5
ElseIf vKey = "F6" Then
    vKey = vbKeyF6
ElseIf vKey = "F7" Then
    vKey = vbKeyF7
ElseIf vKey = "F8" Then
    vKey = vbKeyF8
End If

If GetAsyncKeyState(vKey) Then
    Load Form2
    Form2.Visible = True
    Check10.Value = Checked
    Timer2.Enabled = False
End If

End Sub
Private Sub DetectFolders()
                   
FName = String(250, Chr$(0))
thehwnd = 0
Do
    thehwnd = FindWindowEx(0, thehwnd, vbNullString, vbNullString)
    GetWindowText thehwnd, FName, 250
    
    If winxp = True Then
        Detect = FindWindow("CabinetWClass", FName)
    Else
        Detect = FindWindow("ExploreWClass", FName)
    End If
    
    
        If Detect <> "0" Then
            If List4.ListCount <> 0 Then
                List4.ListIndex = 0
                Do
                    
                    If List4.Text = thehwnd Then
                        GoTo endsub
                    End If
                    If List4.ListIndex = List4.ListCount - 1 Then
                        Exit Do
                    End If

                    List4.ListIndex = List4.ListIndex + 1

                    
                Loop Until List4.ListIndex = List4.ListCount
            End If
            DoEvents
            Form3.Text2.Text = Form3.Text2.Text + vbNewLine + "FOLDER:  " + FName
            Form3.Text2.Text = Form3.Text2.Text + "    DETECTED OPEN    " + " ON:  " + Date$ + "   AT:  " + Time$
            List4.AddItem thehwnd
        End If
    
Loop Until thehwnd = 0
endsub:
End Sub
Private Sub Timer3_Timer()

timerint = Text3.Text * 10000
Timer3.Interval = timerint
Call DetectFolders
End Sub

