VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer sleepTimer 
      Interval        =   3000
      Left            =   105
      Tag             =   "stores and compares the last time to see if the PC has slept"
      Top             =   1560
   End
   Begin VB.Timer settingsTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   90
      Tag             =   "settingsTimer for reading external changes to prefs"
      Top             =   1095
   End
   Begin VB.Timer rotationTimer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   90
      Top             =   615
   End
   Begin VB.Timer revealWidgetTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   135
   End
   Begin VB.Label Label4 
      Caption         =   "sleeptimer for testing awake from sleep"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   1635
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "Note: this invisible form is also the container for the large 128x128px project icon"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   4125
   End
   Begin VB.Label Label2 
      Caption         =   "settingsTimer for reading external changes to prefs"
      Height          =   195
      Left            =   705
      TabIndex        =   2
      Top             =   1170
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "rotationTimer for handling rotation of the screen"
      Height          =   195
      Left            =   690
      TabIndex        =   1
      Top             =   735
      Width           =   3570
   End
   Begin VB.Label Label 
      Caption         =   "revealWidgetTimer for revealing after a hide."
      Height          =   195
      Left            =   690
      TabIndex        =   0
      Top             =   270
      Width           =   3480
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ModuleWithoutFolder
Option Explicit












'---------------------------------------------------------------------------------------
' Procedure : revealWidgetTimer_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub revealWidgetTimer_Timer()
    On Error GoTo revealWidgetTimer_Timer_Error

    revealWidgetTimerCount = revealWidgetTimerCount + 1
    If revealWidgetTimerCount >= (minutesToHide * 12) Then
        revealWidgetTimerCount = 0

        fAlpha.gaugeForm.Visible = True
        revealWidgetTimer.Enabled = False
        PzGWidgetHidden = "0"
        sPutINISetting "Software\PzJustClock", "widgetHidden", PzGWidgetHidden, PzGSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

revealWidgetTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure revealWidgetTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : rotationTimer_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : for handling rotation of the screen in tablet mode
'---------------------------------------------------------------------------------------
'
Private Sub rotationTimer_Timer()
    On Error GoTo rotationTimer_Timer_Error

    screenHeightPixels = GetDeviceCaps(menuForm.hdc, VERTRES) ' we use the name of any form currently loaded
    screenWidthPixels = GetDeviceCaps(menuForm.hdc, HORZRES)
    
    ' will be used to check for orientation changes
    If (oldScreenHeightPixels <> screenHeightPixels) Or (oldScreenWidthPixels <> screenWidthPixels) Then
        
        ' move/hide onto/from the main screen
        Call mainScreen
        
        'store the resolution change
        oldScreenHeightPixels = screenHeightPixels
        oldScreenWidthPixels = screenWidthPixels
    End If

    On Error GoTo 0
    Exit Sub

rotationTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rotationTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : settingsTimer_Timer
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : if the unhide setting is set by another process it will unhide the widget
'---------------------------------------------------------------------------------------
'
Private Sub settingsTimer_Timer()
    
    On Error GoTo settingsTimer_Timer_Error

    PzGUnhide = fGetINISetting("Software\PzJustClock", "unhide", PzGSettingsFile)

    If PzGUnhide = "true" Then
        'overlayWidget.Hidden = False
        fAlpha.gaugeForm.Visible = True
        sPutINISetting "Software\PzJustClock", "unhide", vbNullString, PzGSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

settingsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sleepTimer_Timer
' Author    : beededea
' Date      : 21/04/2021
' Purpose   : timer that stores the last time
' if the current time is greater than the last time stored by more than 30 seconds we can assume the system
' has been sent to sleep, if the two are significantly different then we reorganise the dock
'---------------------------------------------------------------------------------------
'
Private Sub sleepTimer_Timer()
    Dim strTimeNow As Date: strTimeNow = #1/1/2000 12:00:00 PM#  'set a variable to compare for the NOW time
    Dim lngSecondsGap As Long: lngSecondsGap = 0  ' set a variable for the difference in time
    Static strTimeThen As Date
    
    On Error GoTo sleepTimer_Timer_Error

    If strTimeThen = "00:00:00" Then strTimeThen = Now(): Exit Sub
    sleepTimer.Enabled = False
    
    strTimeNow = Now()
    
    lngSecondsGap = DateDiff("s", strTimeThen, strTimeNow)
    strTimeThen = Now()

    If lngSecondsGap > 60 Then
        'MsgBox "System has just woken up from a sleep" ' awoken, awake
        fAlpha.gaugeForm.Refresh
        'MessageBox Me.hwnd, "System has just woken up from a sleep - animatedIconsRaised =" & animatedIconsRaised, "SteamyDock Information Message", vbOKOnly
    End If
    
    sleepTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

sleepTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sleepTimer_Timer of Form dock"

End Sub


