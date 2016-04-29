Attribute VB_Name = "Direct3D"
'-----------------------------------------
'Here I will implement Direct3D functions
'such as Initializing, Rendering,
'Device autodetection, Device termination
'-----------------------------------------
Option Explicit

Public DX As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public matProj As D3DMATRIX
Public matView As D3DMATRIX
Public matWorld As D3DMATRIX
Public matTemp As D3DMATRIX
Public MainFontDesc As IFont
Public MainFont As D3DXFont
Public Fnt As New StdFont
Public TextRect As RECT

Public Const PI = 3.14159
Public Const RAD = PI / 180
Public Const DEG = 180 / PI

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim Dispmode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DevCaps As D3DCAPS8
Dim DevIdent As D3DADAPTER_IDENTIFIER8
Dim Vp As D3DVIEWPORT8
Dim Lights As D3DLIGHT8
Dim BackBuf As Direct3DSurface8

Public Function InitD3D(FullSCR As Boolean) As Boolean
'--------------------------------------------------------------'
'Checks device caps. If available, then initializes them.      '
'If every initialization was succesfull, true is being returned'
'--------------------------------------------------------------'
On Error Resume Next
    '> Presentation Parameters
    If FullSCR Then
        D3DWindow.Windowed = 0
        D3DWindow.BackBufferFormat = Dispmode.Format
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferWidth = Dispmode.Width
        D3DWindow.BackBufferHeight = Dispmode.Height
    Else
        D3DWindow.Windowed = 1
        D3DWindow.BackBufferFormat = Dispmode.Format
    End If
    D3DWindow.hDeviceWindow = frmMain.hwnd
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY ' D3DSWAPEFFECT_COPY
    
    '> Stencil depth
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Dispmode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        D3DWindow.EnableAutoDepthStencil = 1
    Else
        D3DWindow.EnableAutoDepthStencil = 0
    End If
    
    'Anti-aliasing
    If D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Dispmode.Format, False, D3DMULTISAMPLE_2_SAMPLES) = 0 Then
        If frmMain.Op(0).Value = 1 Then 'Was it enabled?
            D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
            D3DWindow.MultiSampleType = D3DMULTISAMPLE_2_SAMPLES
        End If
    End If
    
    '> Hardware/Software VP & initialization
    If frmMain.Op(1).Value = 1 Then 'Hw T&L Option
        Set D3DDevice = D3D.CreateDevice(0, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
        If D3DDevice Is Nothing Then Set D3DDevice = D3D.CreateDevice(0, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    Else
        Set D3DDevice = D3D.CreateDevice(0, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    End If
    
    '> Setting up the World Matrix
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    '> Setting up the View Matrix
    D3DXMatrixLookAtLH matView, MakeVector(0, 0, -10), MakeVector(0, 0, 0), MakeVector(0, 1, 0) ' VIEW
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    '> Setting up the Projection Matrix
    D3DXMatrixPerspectiveFovLH matProj, PI / 2.2, PI / (4), 1, 1000 '10001 'PI/2 ' ' '''' Pi / 3, Pi / (3.4)
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
    '> Enabling ZBuffer
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    
    '> Positioning a camera variable
    CP = MakeVector(58, -17, -11) 'Position
    CT = MakeVector(57, -17, -11) 'Target
    
    '> Getting the viewport (optional)
    D3DDevice.GetViewport Vp
    
    '> Setting up Render States
    D3DDevice.SetRenderState D3DRS_AMBIENT, &H202020
    D3DDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID '
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZBIAS, 0
    If frmMain.Op(0).Value = 1 Then 'Antialias was enabled
        D3DDevice.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, True
    Else
        D3DDevice.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, False
    End If
    D3DDevice.SetRenderState D3DRS_EDGEANTIALIAS, 0
    
    '> Setting up Texture Stages
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, 2
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, 2
    D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, 2
    
    '> Setting up a Font for text rendering
    Fnt.Name = "Arial"
    Fnt.Size = 10
    Fnt.Bold = False
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
        
    '> Setting up lighting
    Lights.Type = D3DLIGHT_POINT
    Lights.diffuse = MakeD3DColorValue(255, 170 + 251, 170 + 224, 170 + 177)
    Lights.Ambient = MakeD3DColorValue(255, 23, 20, 10)
    Lights.specular = MakeD3DColorValue(255, 1, 1, 1)
    Lights.Falloff = 0.001
    Lights.Attenuation0 = 100
    Lights.Range = 1800
    Lights.Position = MakeVector(0, 20, 0)
    Lights.Direction = MakeVector(1, 0, 0)
    Lights.Theta = PI / 3 'Inner Cone
    Lights.Phi = PI / 2 'Outer Cone
    D3DDevice.LightEnable 0, 1
    D3DDevice.SetLight 0, Lights
    '> Setting up lighting2
    If frmMain.Op(2).Value = 1 Then 'Dramatic Light
        Lights.Type = D3DLIGHT_POINT
        Lights.diffuse = MakeD3DColorValue(100, -50, -50, 0)
        Lights.Ambient = MakeD3DColorValue(255, 1, 1, 1)
        Lights.specular = MakeD3DColorValue(255, 1, 1, 1)
        Lights.Falloff = 0.001
        Lights.Attenuation0 = 100
        Lights.Range = 180
        Lights.Position = MakeVector(-200, 20, -200)
        Lights.Direction = MakeVector(1, 0, 0)
        Lights.Theta = PI / 3 'Inner Cone
        Lights.Phi = PI / 2 'Outer Cone
        D3DDevice.LightEnable 1, 1
        D3DDevice.SetLight 1, Lights
    End If

InitD3D = True
Exit Function
BailOutInit:
InitD3D = False
MsgBox "There was an error initializing the graphics engine. Please start the Hangman troubleshooter, or contact custumer support (see readme.txt for details)", vbInformation
End Function

Public Sub TerminateDirect3D()
'-----------------'
'Unloading objects'
'-----------------'
Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing
End Sub

Public Function PreCheck() As Boolean
'--------------------------------------------'
'Displays information about graphics adapter,'
'and creates our main Dx Objects.            '
'--------------------------------------------'
Set DX = New DirectX8
Set D3D = DX.Direct3DCreate()
Set D3DX = New D3DX8
D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps
D3D.GetAdapterIdentifier D3DADAPTER_DEFAULT, D3DENUM_NO_WHQL_LEVEL, DevIdent
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Dispmode
Dim a As Single
frmMain.L_Caps.Caption = "> Graphics adapter: " & vbNewLine & "        "
For a = 0 To 511
frmMain.L_Caps.Caption = frmMain.L_Caps.Caption & Chr(DevIdent.Description(a))
Next a
frmMain.L_Caps.Caption = frmMain.L_Caps.Caption & vbNewLine & vbNewLine & "> Screen Resolution: " & vbNewLine & "       " & CStr(Dispmode.Width) & "x" & CStr(Dispmode.Height)
End Function

Public Sub Render()
'------------------------------------------------------------'
'The entire rendering of objects, and presentation on screen '
'------------------------------------------------------------'
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0, 1#, 0
D3DDevice.BeginScene

D3DXMatrixLookAtLH matView, CP, CT, MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView
D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetVertexShader FVF_VERTEX

D3DDevice.SetRenderState 8, 3
D3DDevice.SetRenderState D3DRS_LIGHTING, 0
Animated(0).Render
Animated(1).Render
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
Ambient(0).Render
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
Scenery.Render

'Show Debug Information
TextXY 1, 1, &HFFEEEEEE, "P: (" & CStr(CInt(CP.X)) & "," & CStr(CInt(CP.Y)) & "," & CStr(CInt(CP.Z)) & ")"
TextXY 1, 15, &HFFEEEEEE, "T: (" & CStr(CInt(CT.X)) & "," & CStr(CInt(CT.Y)) & "," & CStr(CInt(CT.Z)) & ")"
TextXY 1, 29, &HFFEEEEEE, "Waypoint: " & CStr(CWAYPOINT)
TextXY 1, 45, &HFFFFEEEE, "Speed: " & CStr(CSpeed)
TextXY 1, 60, &HFFFFFFFF, "Rot:" & CStr(Find_Angle(WAYPOINTS(CWAYPOINT, 0), WAYPOINTS(CWAYPOINT + 1, 0)) * DEG)
TextXY 1, 75, &HFFFFEEEE, "Animation Clip: " & CStr(C_A_F)
If C_A_F = 8 Then TextXY 250, 20, &HFFFF0000, "You've lost the game! Try Again !"

D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Public Sub Create_Parametric_Fog(Enabled As Boolean, Fog_Color As Long, RangeStart As Single, RangeEnd As Single, Ranged As Boolean, VertexMode As Long, PixelMode As Long)
    D3DDevice.SetRenderState D3DRS_FOGENABLE, Enabled 'set to 0 to disable
    D3DDevice.SetRenderState D3DRS_FOGTABLEMODE, PixelMode  'LINEAR 'dont use table fog
    D3DDevice.SetRenderState D3DRS_FOGVERTEXMODE, VertexMode 'use standard linear fog
    D3DDevice.SetRenderState D3DRS_FOGCOLOR, Fog_Color
    D3DDevice.SetRenderState D3DRS_RANGEFOGENABLE, Ranged 'enable range based fog, hw dependent
    D3DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(RangeStart)
    D3DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(RangeEnd)
End Sub
