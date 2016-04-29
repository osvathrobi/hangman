Attribute VB_Name = "Game"
Public bRunning As Boolean

Public Sub Start()
Load frmMain.Sc_select.ListIndex + 1 'Debugging: this will load the selected scene
ST_CAM = False
'Make sure the animation won't start
ARUNNING = False
IA = 2
Do While bRunning
    '// Debugging: if you comment out the "PlayWithCamera" line,
    '// you can navigate using the following keys:
    '// DirectionKeys for movement, PageUp/PageDown for height, Insert/Delete for Pitching
    HandleKeyb
    PlayWithCamera
    PlayAnim
    Render
    DoEvents
Loop
Direct3D.TerminateDirect3D
End
End Sub
