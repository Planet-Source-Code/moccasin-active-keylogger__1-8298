Attribute VB_Name = "smtpmailer"
'This was programmed by  Joseph Ninan
' email - josephninan@crosswinds.net
' S4-Computer Science and engineering
' SCT College of engineering
' Trivandrum, Kerala, India
' Phone - 0091-471-594477
' www.jofu.8m.com

Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Private lpPrevWndProc As Long
Public mysock As Long
Public Progress, rtncode As Integer
Public Helo_OK, Green_Light, m_timeout, do_cancel As Boolean
Public e_err, e_errstr, timeout As Variant
'Here are the actual values of the rtncode used
'Private Const smtpStatus = 211
'Private Const smtpHelp = 214
'Private Const smtpReady = 220
'Private Const smtpClosing = 221
'Private Const smtpDone = 250
'Private Const smtpWillForward = 251
'Private Const smtpStartMail = 354
'Private Const smtpShuttingDown = 421
'Private Const smtpMailboxUnavailable = 450
'Private Const smtpLocalError = 451
'Private Const smtpNoSpace = 452
'P'rivate Const smtpSyntaxError = 500
'Private Const smtpArgError = 501
'Private Const smtpNoCommand = 502
'Private Const smtpBadSequence = 503
'Private Const smtpNoParamater = 504
'Private Const smtpMailboxUnavailable2 = 550
'Private Const smtpUserRejected = 551
'Private Const smtpTooBig = 552
'Private Const smtpInvalidMailboxName = 553
'Private Const smtpFailed = 554


Public Function Hook(ByVal hWnd As Long)
    'ok, we are going to catch ALL msg's sent
    'to the handle we are subclassing (form1)
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub UnHook(ByVal hWnd As Long)
    'if we dont un-subclass before we shutdown
    'the program, we get an illigal procedure error.
    'fun.
    Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Long, a As String
Dim wp As Integer
Dim temp As Variant
Dim ReadBuffer(1000) As Byte
'Debug.Print uMsg, wParam, lParam
    Select Case uMsg
        Case 1025:
            Debug.Print uMsg, wParam, lParam
            Log uMsg & "  " & wParam & "  " & lParam
            e_err = WSAGetAsyncError(lParam)
            e_errstr = GetWSAErrorString(e_err)
            
            If e_err <> 0 Then
                Log e_err & " - " & e_errstr
                Log "Terminating...."
                do_cancel = True
                'Exit Function
            End If
            Select Case lParam
            Case FD_READ: 'lets check for data
                    x = recv(mysock, ReadBuffer(0), 1000, 0) 'try to get some
                    If x > 0 Then 'was there any?
                        a = StrConv(ReadBuffer, vbUnicode) 'yep, lets change it to stuff we can understand
                        Log a
                        rtncode = Val(Mid(a, 1, 3))
                        'Log "Analysing code " & rtncode & "..."
                        Select Case rtncode
                        Case 354, 250
                            Progress = Progress + 1
                            Log ">>Progress becomes " & Progress
                        Case 220
                            Log "Recieved Greenlight"
                            Green_Light = True
                        Case 221
                            Progress = Progress + 1
                            Log ">>Progress becomes " & Progress
                        Case 550, 551, 552, 553, 554, 451, 452, 500
                            Log "There was some error at the server side"
                            Log "error code is " & rtncode
                            do_cancel = True
                        End Select
                    End If
            Case FD_CONNECT: 'did we connect?
                    mysock = wParam 'yep, we did! yayay
                    'Log WSAGetAsyncError(lParam) & "error code"
                    'Log mysock & " - Mysocket Value"

            Case FD_CLOSE: 'uh oh. they closed the connection
                    Call closesocket(wp)   'so we need to close
            End Select
    End Select
    'let the msg get through to the form
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Public Function GenerateMessageID(ByVal sHost As String) As String
    Dim idnum As Double
    Dim sMessageID As String
    sMessageID = "Message-ID: "
    ' this makes the randomize seed different every time
    Randomize Int(CDbl((Now))) + Timer
    idnum = GetRandom(9999999999999#, 99999999999999#)
    sMessageID = sMessageID & CStr(idnum)
    idnum = GetRandom(9999, 99999)
    sMessageID = sMessageID & "." & CStr(idnum) & ".qmail@" & sHost
    GenerateMessageID = sMessageID
End Function
Public Function GetRandom(ByVal dFrom As Double, ByVal dTo As Double) As Double

    Dim x As Double
    Randomize
    x = dTo - dFrom
    GetRandom = Int((x * Rnd) + 1) + dFrom
End Function
Public Sub Log(ByVal sText As String)
    ' this way it doesnt refresh the whole thing every time, no blinking...
    With Form1.txtstatus
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub


Public Function smtp(m_host, m_port, m_from, m_rcpt, name_from, name_rcpt, m_reply, m_subject, m_data As String) As Integer
'smtp = 1  Mail sent successfully
'smtp =-1  Mail sent met with some error
'smtp = 0  Timed Out
'Log mysock
Dim temp, timeout As Variant
    Progress = 0
    Green_Light = False
    do_cancel = False
    timeout = Timer + 60
    Log "Will timeout in 60 seconds"
    'make sure the port is closed!
    If mysock <> 0 Then Call closesocket(mysock)
    'let's connect!!!       host            port       handle
    temp = ConnectSock(m_host, m_port, 0, Form1.hWnd, True)
    Log "Connect socket return value" & temp
    Log "Connected to " & m_host & " at port " & m_port
    If temp = INVALID_SOCKET Then
        Log "Error -Invalid Socket"
        smtp = -1
        Exit Function
    End If
    While mysock = 0  'make sure we are connected
        DoEvents
        If do_cancel = True Then
            Log "Error .. No connection"
            smtp = -1
            Exit Function
        End If
    Wend
    timeout = Timer + 60
    Log "Connection Established..."
    While Green_Light = False
        
        DoEvents
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend

    Log "HELO " & Mid(m_from, InStr(1, m_from, "@") + 1, Len(m_from)) & vbCrLf
    Call SendData(mysock, "HELO " & Mid(m_from, InStr(1, m_from, "@") + 1, Len(m_from)) & vbCrLf) 'send the data
    While Progress < 1
        DoEvents
        
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            smtp = 0
            Log "Timeout at progress step " & Progress
            Exit Function
        End If
    Wend
    Log "MAIL FROM: <" & m_from & ">" & vbCrLf
    Call SendData(mysock, "MAIL FROM: <" & m_from & ">" & vbCrLf)
    While Progress < 2
        DoEvents
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend
    Log "RCPT TO: <" & m_rcpt & ">" & vbCrLf
    Call SendData(mysock, "RCPT TO: <" & m_rcpt & ">" & vbCrLf)
    While Progress < 3
        
        DoEvents
        
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend
    Log "DATA"
    Call SendData(mysock, "DATA" & vbCrLf)
    While Progress < 4
        DoEvents
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend
    Log "Beginning transfer of body..."
    temp = GenerateMessageID(Mid(m_from, InStr(1, m_from, "@") + 1, Len(m_from)))
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    temp = "DATE: " & Format(Now, "h:mm:ss")
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    temp = "FROM: " & name_from & " <" & m_from & ">"
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    temp = "TO: " & name_rcpt & " <" & m_rcpt & ">"
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    temp = "Reply-to: " & " <" & m_reply & ">"
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    temp = "SUBJECT: " & m_subject
    Log temp
    Call SendData(mysock, temp & vbCrLf)
    Log "MIME-Version: 1.0"
    Call SendData(mysock, "MIME-Version: 1.0" & vbCrLf)
    Log "Content-Type: text/plain; charset=us-ascii"
    Call SendData(mysock, "Content-Type: text/plain; charset=us-ascii" & vbCrLf)
    Log m_data
    Call SendData(mysock, m_data)
    Log vbCrLf & "." & vbCrLf
    Call SendData(mysock, vbCrLf & "." & vbCrLf)
    While Progress < 5
        DoEvents
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend
    Call SendData(mysock, "QUIT" & vbCrLf)
    While Progress < 6
        DoEvents
        If do_cancel = True Then
            Log "Error in between smtp - fatal"
            smtp = -1
            Exit Function
        End If
        
        If Timer > timeout Then
            m_timeout = True
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            smtp = 0
            Exit Function
        End If
    Wend
    
    Log "Mail sent succesfully"
    smtp = 1
    
End Function
