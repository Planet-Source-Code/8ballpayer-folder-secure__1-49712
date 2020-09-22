Attribute VB_Name = "Module1"
Option Explicit
Public folder As String
Public lProcessID As Long
Public lProcessID2 As Long
Public line, test, folders, descrip, dirs, pass, windir As Variant
Public winxp, folderpassentry, autostart, encryptlog, this, sh, dtb, dtm, itm, el, rsf, ra As Boolean
Public count, pos, length, i As Integer
Public plus, remove As Integer


Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_CLOSE = &H10
Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function getwindir Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function WriteFileNO Lib "kernel32.dll" Alias "WriteFile" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Option Base 1
Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
         (ByVal lpstrFolderName As String, _
         ByVal lpstrLinkName As String, _
         ByVal lpstrLinkPath As String, _
         ByVal lpstrLinkArguments As String, _
         ByVal fPrivate As Long, _
         ByVal sParent As String) As Long

Public Const gstrQUOTE$ = """"
Public Const SPI_SCREENSAVERRUNNING = 97
Public Sub CreateShellLink(ByVal strLinkPath As String, _
         ByVal strGroupName As String, _
         ByVal strLinkArguments As String, _
         ByVal strLinkName As String, _
         ByVal fPrivate As Boolean, _
         sParent As String, _
         Optional ByVal fLog As Boolean = True)
Dim fSuccess As Boolean
Dim intMsgRet As Integer
Dim lREt       As Boolean
   strLinkName = strUnQuoteString(strLinkName)
   strLinkPath = strUnQuoteString(strLinkPath)
   
   If StrPtr(strLinkArguments) = 0 Then strLinkArguments = ""
   
   lREt = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments, _
         fPrivate, sParent)
End Sub


Public Function strUnQuoteString(ByVal strQuotedString As String)
    strQuotedString = Trim$(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then

            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function

Public Sub DisableTrap(CurForm As Form)
    
    Dim erg As Long
    Dim NewRect As RECT
    
    With NewRect
        .Left = 0&
        .Top = 0&
        .Right = Screen.Width / Screen.TwipsPerPixelX
        .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With

    erg& = ClipCursor(NewRect)
    
End Sub
Public Sub EnableTrap(CurForm As Form)
    
    Dim X As Long, Y As Long, erg As Long
    Dim NewRect As RECT
    
    X& = Screen.TwipsPerPixelX
    Y& = Screen.TwipsPerPixelY
    With NewRect
        .Left = CurForm.Left / X&
        .Top = CurForm.Top / Y&
        .Right = .Left + CurForm.Width / X&
        .Bottom = .Top + CurForm.Height / Y&
    End With
    erg& = ClipCursor(NewRect)
    
End Sub

Public Sub SetOnTop(ByVal hwnd As Long, ByVal bSetOnTop As Boolean)
    Dim lR As Long
    If bSetOnTop Then
        lR = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        lR = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub

Public Sub Running()
    
    If Form1.Option1 = True Then
        Form1.Timer1.Interval = "1000"
    ElseIf Form1.Option2 = True Then
        Form1.Timer1.Interval = "500"
    ElseIf Form1.Option3 = True Then
        Form1.Timer1.Interval = "100"
    ElseIf Form1.Option4 = True Then
        Form1.Timer1.Interval = "10"
    End If
   
    Form1.Caption = "Folder Secure (Running)"
    Form1.cmdretrieve.Caption = "Stop Security"
    Form1.Timer1.Enabled = True
    
    If Form1.Check7.Value = Checked Then
        Form1.Timer2.Enabled = True
        Form1.Visible = False
    End If
    
    
End Sub
Public Sub loadConfig()

   On Error GoTo error
    
    Open App.Path & "\config.fsf" For Input As #1
        Do Until EOF(1)
            Line Input #1, line
                Form1.List3.AddItem (line)
        Loop
    Close #1
        
error:
End Sub
Public Sub loadPass2()
                
    On Error GoTo error
             
    Open App.Path & "\pass2.fsf" For Input As #1
        Do Until EOF(1)
            Line Input #1, line
                Form6.Text2.Text = line
        Loop
        Close #1
error:
End Sub
Public Sub loadPass()
    
    On Error GoTo error
    
    Open App.Path & "\pass.fsf" For Input As #1
        Do Until EOF(1)
            Line Input #1, line
                Form2.Text2.Text = line
                Form4.Text4.Text = line
        Loop
    Close #1
    
error:
    
End Sub
Public Sub loadLog()
        
    
    On Error GoTo error
    Open App.Path & "\folderlog.fsf" For Input As #1
        Do Until EOF(1)
            Line Input #1, line
            If Form3.Text2.Text = "" Then
                Form3.Text2.Text = line
            Else
                Form3.Text2.Text = Form3.Text2.Text + vbNewLine + line
            End If
        Loop
    Close #1
    
error:
    
            
End Sub
Public Sub loadFiles()
    
    On Error GoTo error
    Open App.Path & "\folders.fsf" For Input As #1
        Do Until EOF(1)
            Line Input #1, line
            If line <> "" Then
                test = Left(line, 8)
                If test = "<start>=" Then
                    line = Mid(line, 9, (Len(line) - 8))
                    Form1.List1.AddItem (line)
                    Form1.List2.AddItem (line)
                End If
            End If
        Loop
    Close #1

error:
End Sub
Public Sub savePass()

    Dim e As Integer
    Dim encrypt
    Dim encrypt2
        For e = 1 To 5
            encrypt = Encode(Form4.Text3.Text)
            encrypt2 = encrypt
        Next e
      
    Close #1
    Open App.Path & "\pass.fsf" For Output As #1
        Print #1, encrypt2
    Close #1

End Sub
Public Sub savePass2()
    
    Dim e As Integer
    Dim encrypt
    Dim encrypt2
        For e = 1 To 5
            encrypt = Encode(Form5.Text2.Text)
            encrypt2 = encrypt
        Next e
        
    Close #1
    Open App.Path & "\pass2.fsf" For Output As #1
        Print #1, encrypt2
    Close #1
    
    
End Sub

Public Sub saveFiles()
    
    Dim data As Variant
    
    Open App.Path & "\folders.fsf" For Output As #1
        data = "[Grammer]" & vbCrLf & "type=cfg" & vbCrLf & "[<start>]" & vbCrLf
        Form1.List1.ListIndex = -1
        For i = 0 To Form1.List1.ListCount
            If Form1.List1.List(i) <> "" Then
                data = data & "<start>=" & Form1.List1.List(i) & vbCrLf
            End If
        Next i
        Print #1, data
    Close #1
    
End Sub
Public Sub saveConfig()

    Dim data As String

    Open App.Path & "\config.fsf" For Output As #1
        If winxp = True Then
            data = "true"
        Else
            data = "false"
        End If
        If folderpassentry = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If autostart = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If encryptlog = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If dtb = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If dtm = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If itm = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If el = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If rsf = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If ra = True Then
            data = data + vbNewLine + "true"
        Else
            data = data + vbNewLine + "false"
        End If
        If Form1.Check3.Value = 1 And Form1.Text3.Text <> "" Then
            data = data + vbNewLine + Form1.Text3.Text
        Else
            data = data + vbNewLine + ""
        End If
        data = data + vbNewLine + Form1.Text2.Text
        data = data + vbNewLine + Form1.txtalertmsg.Text


        Print #1, data
    Close #1
        
        
End Sub

Public Sub saveLog()
    Dim data As Variant

    Dim e As Integer
    Dim encrypt
    Dim encrypt2

      
    Open App.Path & "\folderlog.fsf" For Output As #1
        data = Form3.Text1.Text
        If encryptlog = True Then
            For e = 1 To 5
                encrypt = Encode(Form3.Text1.Text)
                encrypt2 = encrypt
            Next e
            Print #1, encrypt2
            Close #1
        Else
            Print #1, data
            Close #1
        End If
    
    
End Sub


Public Function DeCode(vText As String) As String
    On Error GoTo ErrHandler
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        varChr = Mid(vText, CurSpc, 3)
        Select Case varChr
            'lower case
            Case "coe"
                varChr = "a"
            Case "wer"
                varChr = "b"
            Case "ibq"
                varChr = "c"
            Case "am7"
                varChr = "d"
            Case "pm1"
                varChr = "e"
            Case "mop"
                varChr = "f"
            Case "9v4"
                varChr = "g"
            Case "qu6"
                varChr = "h"
            Case "zxc"
                varChr = "i"
            Case "4mp"
                varChr = "j"
            Case "f88"
                varChr = "k"
            Case "qe2"
                varChr = "l"
            Case "vbn"
                varChr = "m"
            Case "qwt"
                varChr = "n"
            Case "pl5"
                varChr = "o"
            Case "13s"
                varChr = "p"
            Case "c%l"
                varChr = "q"
            Case "w$w"
                varChr = "r"
            Case "6a@"
                varChr = "s"
            Case "!2&"
                varChr = "t"
            Case "(=c"
                varChr = "u"
            Case "wvf"
                varChr = "v"
            Case "dp0"
                varChr = "w"
            Case "w$-"
                varChr = "x"
            Case "vn&"
                varChr = "y"
            Case "c*4"
                varChr = "z"
            'numbers
            Case "aq@"
                varChr = "1"
            Case "902"
                varChr = "2"
            Case "2.&"
                varChr = "3"
            Case "/w!"
                varChr = "4"
            Case "|pq"
                varChr = "5"
            Case "ml|"
                varChr = "6"
            Case "t'?"
                varChr = "7"
            Case ">^s"
                varChr = "8"
            Case "<s^"
                varChr = "9"
            Case ";&c"
                varChr = "0"
            'caps
            Case "$)c"
                varChr = "A"
            Case "-gt"
                varChr = "B"
            Case "|p*"
                varChr = "C"
            Case "1" & Chr(34) & "r"
                varChr = "D"
            Case "c>:"
                varChr = "E"
            Case "@+x"
                varChr = "F"
            Case "v^a"
                varChr = "G"
            Case "]eE"
                varChr = "H"
            Case "aP0"
                varChr = "I"
            Case "{=1"
                varChr = "J"
            Case "cWv"
                varChr = "K"
            Case "cDc"
                varChr = "L"
            Case "*,!"
                varChr = "M"
            Case "fW" & Chr(34)
                varChr = "N"
            Case ".?T"
                varChr = "O"
            Case "%<8"
                varChr = "P"
            Case "@:a"
                varChr = "Q"
            Case "&c$"
                varChr = "R"
            Case "WnY"
                varChr = "S"
            Case "{Sh"
                varChr = "T"
            Case "_%M"
                varChr = "U"
            Case "}'$"
                varChr = "V"
            Case "QlU"
                varChr = "W"
            Case "Im^"
                varChr = "X"
            Case "l|P"
                varChr = "Y"
            Case ".>#"
                varChr = "Z"
            'Special characters
            Case "\" & Chr(34) & "]"
                varChr = "!"
            Case "cY,"
                varChr = "@"
            Case "x%B"
                varChr = "#"
            Case "a*v"
                varChr = "$"
            Case "'&T"
                varChr = "%"
            Case ";%R"
                varChr = "^"
            Case "eG_"
                varChr = "&"
            Case "Z/e"
                varChr = "*"
            Case "rG\"
                varChr = "("
            Case "]*F"
                varChr = ")"
            Case "@B*"
                varChr = "_"
            Case "+Hc"
                varChr = "-"
            Case "&|D"
                varChr = "="
            Case "(:#"
                varChr = "+"
            Case "SlW"
                varChr = "["
            Case "'QB"
                varChr = "]"
            Case "{D>"
                varChr = "{"
            Case "+c%"
                varChr = "}"
            Case "(s:"
                varChr = ":"
            Case "^a("
                varChr = ";"
            Case "16."
                varChr = "'"
            Case "s.*"
                varChr = Chr(34)
            Case "&?W"
                varChr = ","
            Case "GPQ"
                varChr = "."
            Case "SK*"
                varChr = "<"
            Case "RL^"
                varChr = ">"
            Case "40C"
                varChr = "/"
            Case "?#9"
                varChr = "?"
            Case "_?/"
                varChr = "\"
            Case "(_@"
                varChr = "|"
            Case "=#B"
                varChr = " "
        End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3
        DoEvents
    Loop
    DeCode = varFin
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function

Public Function Encode(vText As String)
    On Error GoTo ErrHandler
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        CurSpc = CurSpc + 1
        varChr = Mid(vText, CurSpc, 1)
        Select Case varChr
            'lower case
            Case "a"
                varChr = "coe"
            Case "b"
                varChr = "wer"
            Case "c"
                varChr = "ibq"
            Case "d"
                varChr = "am7"
            Case "e"
                varChr = "pm1"
            Case "f"
                varChr = "mop"
            Case "g"
                varChr = "9v4"
            Case "h"
                varChr = "qu6"
            Case "i"
                varChr = "zxc"
            Case "j"
                varChr = "4mp"
            Case "k"
                varChr = "f88"
            Case "l"
                varChr = "qe2"
            Case "m"
                varChr = "vbn"
            Case "n"
                varChr = "qwt"
            Case "o"
                varChr = "pl5"
            Case "p"
                varChr = "13s"
            Case "q"
                varChr = "c%l"
            Case "r"
                varChr = "w$w"
            Case "s"
                varChr = "6a@"
            Case "t"
                varChr = "!2&"
            Case "u"
                varChr = "(=c"
            Case "v"
                varChr = "wvf"
            Case "w"
                varChr = "dp0"
            Case "x"
                varChr = "w$-"
            Case "y"
                varChr = "vn&"
            Case "z"
                varChr = "c*4"
            'numbers
            Case "1"
                varChr = "aq@"
            Case "2"
                varChr = "902"
            Case "3"
                varChr = "2.&"
            Case "4"
                varChr = "/w!"
            Case "5"
                varChr = "|pq"
            Case "6"
                varChr = "ml|"
            Case "7"
                varChr = "t'?"
            Case "8"
                varChr = ">^s"
            Case "9"
                varChr = "<s^"
            Case "0"
                varChr = ";&c"
            'caps
            Case "A"
                varChr = "$)c"
            Case "B"
                varChr = "-gt"
            Case "C"
                varChr = "|p*"
            Case "D"
                varChr = "1" & Chr(34) & "r"
            Case "E"
                varChr = "c>:"
            Case "F"
                varChr = "@+x"
            Case "G"
                varChr = "v^a"
            Case "H"
                varChr = "]eE"
            Case "I"
                varChr = "aP0"
            Case "J"
                varChr = "{=1"
            Case "K"
                varChr = "cWv"
            Case "L"
                varChr = "cDc"
            Case "M"
                varChr = "*,!"
            Case "N"
                varChr = "fW" & Chr(34)
            Case "O"
                varChr = ".?T"
            Case "P"
                varChr = "%<8"
            Case "Q"
                varChr = "@:a"
            Case "R"
                varChr = "&c$"
            Case "S"
                varChr = "WnY"
            Case "T"
                varChr = "{Sh"
            Case "U"
                varChr = "_%M"
            Case "V"
                varChr = "}'$"
            Case "W"
                varChr = "QlU"
            Case "X"
                varChr = "Im^"
            Case "Y"
                varChr = "l|P"
            Case "Z"
                varChr = ".>#"
            'Special characters
            Case "!"
                varChr = "\" & Chr(34) & "]"
            Case "@"
                varChr = "cY,"
            Case "#"
                varChr = "x%B"
            Case "$"
                varChr = "a*v"
            Case "%"
                varChr = "'&T"
            Case "^"
                varChr = ";%R"
            Case "&"
                varChr = "eG_"
            Case "*"
                varChr = "Z/e"
            Case "("
                varChr = "rG\"
            Case ")"
                varChr = "]*F"
            Case "_"
                varChr = "@B*"
            Case "-"
                varChr = "+Hc"
            Case "="
                varChr = "&|D"
            Case "+"
                varChr = "(:#"
            Case "["
                varChr = "SlW"
            Case "]"
                varChr = "'QB"
            Case "{"
                varChr = "{D>"
            Case "}"
                varChr = "+c%"
            Case ":"
                varChr = "(s:"
            Case ";"
                varChr = "^a("
            Case "'"
                varChr = "16."
            Case Chr(34)
                varChr = "s.*"
            Case ","
                varChr = "&?W"
            Case "."
                varChr = "GPQ"
            Case "<"
                varChr = "SK*"
            Case ">"
                varChr = "RL^"
            Case "/"
                varChr = "40C"
            Case "?"
                varChr = "?#9"
            Case "\"
                varChr = "_?/"
            Case "|"
                varChr = "(_@"
            Case " "
                varChr = "=#B"
        End Select
        varFin = varFin & varChr
        DoEvents
    Loop
    Encode = varFin
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function
