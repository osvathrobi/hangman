VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangman in 3D (where available :)"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Options"
      ForeColor       =   &H80000000&
      Height          =   6735
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.CheckBox Op 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "&Enable Dramatic Lights"
         ForeColor       =   &H80000005&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Op 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "&Hardware T&&L"
         ForeColor       =   &H80000005&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Op 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "&Anti-aliasing"
         ForeColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.ComboBox Sc_select 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmMain.frx":E1D0C
         Left            =   1080
         List            =   "frmMain.frx":E1D28
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "&Full Screen"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Scenario:"
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label L_Caps 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label StartB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   6240
         Width           =   855
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
   End
   Begin VB.Menu Player 
      Caption         =   "&Player"
   End
   Begin VB.Menu Settings 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu fullscreen 
         Caption         =   "&Fullscreen"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
   End
   Begin VB.Menu helpb 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim AP As Byte '// Just used for Debugging scenes (holds "Animation Part" -> 1 to 8)

Private Sub Form_Load()
    Direct3D.PreCheck
    Me.Show
    Sc_select.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRunning = False
End Sub

Private Sub StartB_Click()
StartB.Caption = "Loading..."
DoEvents
Frame1.Visible = False
    If Check1.Value = 0 Then
        Game.bRunning = Direct3D.InitD3D(False)
    Else
        Game.bRunning = Direct3D.InitD3D(True)
    End If
Game.Start
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'// "Enter" will save coordinates: Camera Position, Camera Target, Character Position
If KeyAscii = vbKeyReturn Then
Open App.Path & "\coords.txt" For Append As #7
    Print #7, "(" & CStr(CP.X) & ", " & CStr(CP.Y) & ", " & CStr(CP.Z) & ")"
Close #7
Open App.Path & "\coords2.txt" For Append As #8
    Print #8, "(" & CStr(CT.X) & ", " & CStr(CT.Y) & ", " & CStr(CT.Z) & ")"
Close #8
Dim vTemp3d As D3DVECTOR 'cut
vTemp3d = Animated(0).CMesh.Get_Pos
Write_To_File App.Path & "\Guy.txt", "(" & CStr(vTemp3d.X) & ", " & CStr(vTemp3d.Y) & ", " & CStr(vTemp3d.Z) & ")"
End If

'// "Tab" wil play an Animation Clip
If KeyAscii = vbKeyTab Then
    AP = AP + 1
    If AP > 8 Then AP = 8
    Play_Anim_Clip AP
End If

'// "Space" will Camera mode change
If KeyAscii = vbKeySpace Then
    ST_CAM = ST_CAM + 1
    If ST_CAM > 2 Then ST_CAM = 0
End If
End Sub

