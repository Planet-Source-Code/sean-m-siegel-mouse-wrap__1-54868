VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00DFB794&
      Caption         =   "Show on starup"
      Height          =   315
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   1935
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'*Mouse Wrap By Sean Siegel 7-11-2004                                 *
'*This program could be usefull if you were extreamly lazy. Or simply *
'*annoying as hell. This code demonstrates how to get and set the     *
'*pointers cordinates.                                                *
'**********************************************************************
'*Feel free to make changes or use any portion of this code           *
'*Just let me know if you use it or if you have any ideas.            *
'***********************************************************************


Dim PointerXY As POINTAPI     'initialize the pointer cordinates variable
Dim ScreenHeight As Long      'initialize the screen height variable
Dim ScreenWidth As Long       'initialize the screen width variable

Private Sub Check1_Click()
    'save splash screen setting
    SaveSetting "mousewrap", "splash", "enabled", Check1.Value
End Sub

Private Sub Form_Click()
    'skip the timers interval and run its sub directly
    Timer1_Timer
End Sub

Private Sub Form_Load()
    'store some commonly used formulas
    ScreenWidth = Screen.Width \ Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height \ Screen.TwipsPerPixelY
    
    'set the form dimentions to the dimentions of the splash image
    Me.Width = 300 * Screen.TwipsPerPixelX
    Me.Height = 154 * Screen.TwipsPerPixelY

    'get the application path
    ap$ = App.Path
    
    'make sure the path ends in \
    If Right(ap$, 1) <> "\" Then ap$ = ap$ & "\"
    
    'load the splashscreen from the hard drive so it doesnt take up ram since this program runs in the background and is only needed once
    Me.Picture = LoadPicture(ap$ & "start.jpg")
End Sub

Private Sub Timer1_Timer()
    'hide the splash screen after 3 seconds.
    Me.Hide
    'disable the timer since the form is now hidden
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    'load the pointer position into the cordinate variable
    GetCursorPos PointerXY
    'check if the pointer touches the bottom of the screen and reposition it on the top of the screen
    If PointerXY.y = ScreenHeight - 1 Then SetCursorPos PointerXY.x, 1
    'check if the pointer touches the top of the screen and reposition it on the bottom of the screen
    If PointerXY.y = 0 Then SetCursorPos PointerXY.x, ScreenHeight - 2
    'check if the pointer touches the left of the screen and reposition it on the right of the screen
    If PointerXY.x = 0 Then SetCursorPos ScreenWidth - 2, PointerXY.y
    'check if the pointer touches the right of the screen and reposition it on the left of the screen
    If PointerXY.x = ScreenWidth - 1 Then SetCursorPos 1, PointerXY.y
End Sub
