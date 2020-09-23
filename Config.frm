VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   1845
   ClientLeft      =   3540
   ClientTop       =   2835
   ClientWidth     =   2640
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRestart 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   765
      Width           =   495
   End
   Begin VB.TextBox txtGravStrength 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   420
      Width           =   495
   End
   Begin VB.TextBox txtNumBars 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   75
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label labRestart 
      AutoSize        =   -1  'True
      Caption         =   "Secs Between Refreshes:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   810
      Width           =   1845
   End
   Begin VB.Label labGravStrength 
      AutoSize        =   -1  'True
      Caption         =   "Gravity Strength:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   465
      Width           =   1185
   End
   Begin VB.Label labNumBars 
      AutoSize        =   -1  'True
      Caption         =   "Number of Bars To Draw:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Save configuration information in the registry.
Private Sub SaveConfig()

    SaveSetting App.EXEName, "runPrefs", "gravityStrength", txtGravStrength
    SaveSetting App.EXEName, "runPrefs", "numOfBarsToDraw", txtNumBars
    SaveSetting App.EXEName, "runPrefs", "restart_Interval", txtRestart * 1000
    
End Sub

' Load configuration information from the registry.
Public Sub LoadConfig()

    txtNumBars = CInt(GetSetting(App.EXEName, "runPrefs", "numOfBarsToDraw", "1"))
    txtGravStrength = CLng(GetSetting(App.EXEName, "runPrefs", "gravityStrength", "20"))
    txtRestart = CLng(GetSetting(App.EXEName, "runPrefs", "restart_Interval", "65000")) / 1000

End Sub

' Save the new configuration values.
Private Sub cmdOk_Click()
    ' Get the new values.
    On Error Resume Next
    aaa = CInt(txtNumBars)
    aaa = CInt(txtGravStrength)
    aaa = CInt(txtRestart)
    On Error GoTo 0

    ' Save the new values.
    SaveConfig

    ' Unload this form.
    Unload Me
End Sub

' Fill in current values.
Private Sub Form_Load()
    ' Load the current configuration information.
    LoadConfig
    
End Sub

Private Sub txtNumBars_Change()
On Error Resume Next

    If CInt(txtNumBars.Text) <= 4 Then Exit Sub
    
    txtNumBars.Text = 4

End Sub

Private Sub txtGravStrength_Change()
On Error Resume Next

    Dim ab As Long

    ab = CLng(txtGravStrength)

    If ab = CLng(txtGravStrength) Then Exit Sub
    
    txtGravStrength.Text = 20

End Sub

Private Sub txtRestart_Change()
On Error Resume Next

    If CInt(txtRestart.Text) <= 65 Then Exit Sub
    
    txtNumBars.Text = 65

End Sub
