Attribute VB_Name = "Hide"
' hide module
' --------------
' those functions effectively hide/unhide the program
' from Win9x TaskList

Option Explicit

Public Declare Function GetCurrentProcessId Lib "kernel32" _
    () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" _
    (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    
Sub HideMe()

    Dim Ret As Long
    
    Ret = RegisterServiceProcess(GetCurrentProcessId, 1)

End Sub

Sub UnHideMe()  ' not used by this project but good to know :-)

    Dim Ret As Long
    
    Ret = RegisterServiceProcess(GetCurrentProcessId, 0)

End Sub
