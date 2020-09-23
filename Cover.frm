VERSION 5.00
Begin VB.Form frmCover 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   DrawMode        =   4  'Mask Not Pen
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Timer restart 
      Interval        =   65000
      Left            =   480
      Top             =   960
   End
End
Attribute VB_Name = "frmCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'pieces of data
'********************************************
Private a As atom
Private b As atom
Private c As atom
Private d As atom

Private counter As Double

Private Sub loadPrefs()

    gravityStrength = GetSetting(App.EXEName, "runPrefs", "gravityStrength", 20)
    numOfBarsToDraw = GetSetting(App.EXEName, "runPrefs", "numOfBarsToDraw", 1)
    restart.Interval = GetSetting(App.EXEName, "runPrefs", "restart_Interval", 65000)
    restart.Enabled = True

End Sub

Private Sub randomizeStart()
    
    Me.Cls
    
    a.xV = 0
    b.xV = 0
    c.xV = 0

    a.yV = 0
    b.yV = 0
    c.yV = 0
    
    Randomize Timer

'*************************************************************
'if you want things to start really whacky then uncomment these.
'    a.xV = Rnd * 100
'    b.xV = Rnd * 100
'    c.xV = Rnd * 100
'
'    a.yV = Rnd * 100
'    b.yV = Rnd * 100
'    c.yV = Rnd * 100
'*************************************************************
    
    a.X = Rnd * Me.Width
    b.X = Rnd * Me.Width
    c.X = Rnd * Me.Width
    d.X = Rnd * Me.Width
    
    a.Y = Rnd * Me.Height
    b.Y = Rnd * Me.Height
    c.Y = Rnd * Me.Height
    d.Y = Rnd * Me.Height
    
End Sub

Private Sub Form_Load()

    loadPrefs
    
    randomizeStart

End Sub

Private Function step(which As atom)

    which.X = which.X - which.xV
    which.Y = which.Y - which.yV

End Function

Private Function invert(ByRef X As Double) As Double

    invert = X * -1

End Function

Private Function isOutX(which As atom)

    isOut = False
    If which.X < 0 Or which.X > Me.Width Then isOutX = True

End Function

Private Function isOutY(which As atom)

    isOut = False
    If which.Y < 0 Or which.Y > Me.Height Then isOutY = True

End Function

Private Function passTime()

'***************************************************
'these force balls to bounce around inside the screen
'uncomment if you want
'***************************************************
'    If isOutX(a) Then a.xV = invert(a.xV)
'    If isOutY(a) Then a.yV = invert(a.yV)
'
'    If isOutX(b) Then b.xV = invert(b.xV)
'    If isOutY(b) Then b.yV = invert(b.yV)
'
'    If isOutX(c) Then c.xV = invert(c.xV)
'    If isOutY(c) Then c.yV = invert(c.yV)
'
    If isOutX(d) Then d.xV = invert(d.xV)
    If isOutY(d) Then d.yV = invert(d.yV)
    
    performGravity a, b
    performGravity b, a
    performGravity b, c
    performGravity c, b
    performGravity c, d
    performGravity d, c
    performGravity d, a
    performGravity a, d

    step a
    step b
    step c
    step d
    
    drawSquare a.X, a.Y, b.X, b.Y, c.X, c.Y, d.X, d.Y

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = 2 Then End
        
    randomizeStart
    
    Do While True
        DoEvents
        passTime
    Loop

End Sub

Private Sub drawSquare(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double)
On Error Resume Next
    
    If numOfBarsToDraw > 0 Then
        l = (Abs(x1 - x2) / 50) Mod 256
        Me.Line (x1, y1)-(x2, y2), RGB(256 - l, l, l)
    End If
    
    If numOfBarsToDraw > 1 Then
        l = (Abs(x2 - x3) / 50) Mod 256
        Me.Line (x2, y2)-(x3, y3), RGB(l, 256 - l, l)
    End If
    
    If numOfBarsToDraw > 2 Then
        l = (Abs(x3 - x4) / 50) Mod 256
        Me.Line (x3, y3)-(x4, y4), RGB(l, l, 256 - l)
    End If
    
    If numOfBarsToDraw > 3 Then
        l = (Abs(x4 - x1) / 50) Mod 256
        Me.Line (x4, y4)-(x1, y1), RGB(256 - l, 256 - l, 256 - l)
    End If

End Sub

Private Function performGravity(ByRef one As atom, ByRef two As atom)
On Error Resume Next

    temp = (one.X - two.X) ^ 2 + (one.Y - two.Y) ^ 2
    
    one.xV = one.xV + (one.X - two.X) / temp * gravityStrength
    one.yV = one.yV + (one.Y - two.Y) / temp * gravityStrength
    
End Function

Private Sub restart_Timer()

    randomizeStart

End Sub
