Attribute VB_Name = "SubMain"
Option Explicit

'********************************************
'configuration variables
'********************************************
Public gravityStrength As Double       'default 20
Public numOfBarsToDraw As Byte         'default 1


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' Global variables.
Public Const rmConfigure = 1
Public Const rmScreenSaver = 2
Public Const rmPreview = 3
Public RunMode As Integer

Public Type Ball
    BallClr As Long
    BallR As Single
    BallX As Single
    BallY As Single
    BallVx As Single
    BallVy As Single
End Type

Public NumBalls As Integer
Public Balls() As Ball

' Private variables.
Private Const APP_NAME = "BouncingBalls"

' See if another instance of the program is
' running in screen saver mode.
Private Sub CheckShouldRun()
    ' If no instance is running, we're safe.
    If Not App.PrevInstance Then Exit Sub

    ' See if there is a screen saver mode instance.
    If FindWindow(vbNullString, APP_NAME) Then End

    ' Set our caption so other instances can find
    ' us in the previous line.
    frmCover.Caption = APP_NAME
End Sub

' Start the program.
Public Sub Main()
Dim args As String
Dim preview_hwnd As Long
Dim preview_rect As RECT
Dim window_style As Long

    ' Get the command line arguments.
    args = UCase$(Trim$(Command$))

    ' Examine the first 2 characters.
    Select Case Mid$(args, 1, 2)
        Case "/C"   ' Display configuration dialog.
            RunMode = rmConfigure
        Case "", "/S"   ' Run as a screen saver.
            RunMode = rmScreenSaver
        Case "/P"       ' Run in preview mode.
            RunMode = rmPreview
        Case Else       ' This shouldn't happen.
            RunMode = rmScreenSaver
    End Select

    Select Case RunMode
        Case rmConfigure    ' Display configuration dialog.
            frmConfig.Show
        
        Case rmScreenSaver  ' Run as a screen saver.
            ' Make sure there isn't another one running.
            CheckShouldRun

            ' Display the cover form.
            Load frmCover
            frmCover.Show
            ShowCursor False

        Case rmPreview      ' Run in preview mode.
            ' Get the preview area hWnd.
            preview_hwnd = GetHwndFromCommand(args)

            ' Get the dimensions of the preview area.
            GetClientRect preview_hwnd, preview_rect

            Load frmCover

            ' Set the caption for Windows 95.
            frmCover.Caption = "Preview"

            ' Get the current window style.
            window_style = GetWindowLong(frmCover.hwnd, GWL_STYLE)

            ' Add WS_CHILD to make this a child window.
            window_style = (window_style Or WS_CHILD)

            ' Set the window's new style.
            SetWindowLong frmCover.hwnd, _
                GWL_STYLE, window_style

            ' Set the window's parent so it appears
            ' inside the preview area.
            SetParent frmCover.hwnd, preview_hwnd

            ' Save the preview area's hWnd in
            ' the form's window structure.
            SetWindowLong frmCover.hwnd, _
                GWL_HWNDPARENT, preview_hwnd

            ' Show the preview.
            SetWindowPos frmCover.hwnd, _
                HWND_TOP, 0&, 0&, _
                preview_rect.Right, _
                preview_rect.Bottom, _
                SWP_NOZORDER Or SWP_NOACTIVATE Or _
                    SWP_SHOWWINDOW
    End Select
End Sub

' Get the hWnd for the preview window from the
' command line arguments.
Private Function GetHwndFromCommand(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

    ' Take the rightmost numeric characters.
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

