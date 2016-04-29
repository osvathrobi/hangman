Attribute VB_Name = "Helpers"
'----------------------------------------------------
' This module will contain all the utility functions,
' I will use throughout the project
'----------------------------------------------------

Function FloatToDWord(f As Single) As Long
    Dim BUF As D3DXBuffer
    Dim L As Long
    Set BUF = D3DX.CreateBuffer(4)
    D3DX.BufferSetData BUF, 0, 4, 1, f
    D3DX.BufferGetData BUF, 0, 4, 1, L
    FloatToDWord = L
End Function

Public Sub TextXY(X As Single, Y As Single, COLOR As Long, TEXT As String)
'------------------------------------------------------------------------
'Displays a COLOUR-ed TEXT at positions (X,Y) on the surface of the screen
'------------------------------------------------------------------------
TextRect.Top = Y
TextRect.Left = X
TextRect.bottom = Y + Round(Len(TEXT) * (MainFontDesc.Size + 4))
TextRect.Right = X + Round(Len(TEXT) * (MainFontDesc.Size + 4))
D3DX.DrawText MainFont, COLOR, TEXT, TextRect, DT_TOP Or DT_LEFT
End Sub

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
'------------------------------
'Creates a 3 dimensional Vector
'------------------------------
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.Z = Z
End Function

Public Function MakeVector2(X As Single, Y As Single, Z As Single) As D3DVERTEX
'------------------------------
'Creates a 2 dimensional Vector
'------------------------------
    MakeVector2.X = X: MakeVector2.Y = Y: MakeVector2.Z = Z
End Function

Public Function Subs(P1 As D3DVECTOR, P2 As D3DVECTOR) As D3DVECTOR
'---------------------
'Substracts 2 Vectors
'---------------------
With Subs
    .X = P1.X - P2.X
    .Y = P1.Y - P2.Y
    .Z = P1.Z - P2.Z
End With
End Function

Public Function VecEQ(P1 As D3DVECTOR, P2 As D3DVECTOR) As Boolean
'-------------------------------
'Checks if two vectors are equal
'-------------------------------
If ((P1.X = P2.X) And (P1.Y = P2.Y)) And (P1.Z = P2.Z) Then
    VecEQ = True
Else
    VecEQ = False
End If
End Function

Public Function MakeD3DColorValue(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
'------------------------
'Makes a D3D Colour Value
'------------------------
MakeD3DColorValue.a = a '/ 255
MakeD3DColorValue.b = b '/ 255
MakeD3DColorValue.g = g '/ 255
MakeD3DColorValue.r = r '/ 255
End Function

Public Function Dist3D(P1 As D3DVECTOR, P2 As D3DVECTOR) As Single
'-------------------------------------
'Calculates distance between 2 vectors
'-------------------------------------
Dist3D = Sqr((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2 + (P2.Z - P1.Z) ^ 2)
End Function

Public Sub Position_Mesh(P As D3DVECTOR, S As D3DVECTOR, r As D3DVECTOR)
'-------------------------------------------------------------------
' Modifies the appropriate matrixes, resulting in the
'(P)ositioning, (S)caling and (R)otation of the next rendered mesh
'-------------------------------------------------------------------
D3DXMatrixIdentity matWorld
D3DXMatrixIdentity matTemp
D3DXMatrixRotationX matTemp, r.X * RAD
D3DXMatrixMultiply matWorld, matWorld, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixRotationY matTemp, r.Y * RAD
D3DXMatrixMultiply matWorld, matWorld, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixRotationZ matTemp, r.Z * RAD
D3DXMatrixMultiply matWorld, matWorld, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixScaling matTemp, S.X, S.Y, S.Z
D3DXMatrixMultiply matWorld, matWorld, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, P.X, P.Y, P.Z
D3DXMatrixMultiply matWorld, matWorld, matTemp
D3DDevice.SetTransform D3DTS_WORLD, matWorld
End Sub

Public Sub HandleKeyb()
'--------------------------
' Handles Keyboard Events
' faster then Form_KeyDown
'--------------------------
If GetKeyState(vbKeyUp) < 0 Then MoveCamera WALK_SPEED
If GetKeyState(vbKeyDown) < 0 Then MoveCamera -WALK_SPEED
If GetKeyState(vbKeyRight) < 0 Then
    RP = RP + TURN_SPEED
    RotateCam_XoZ RP
End If
If GetKeyState(vbKeyLeft) < 0 Then
    RP = RP - TURN_SPEED
    RotateCam_XoZ RP
End If
If GetKeyState(vbKeyPageDown) < 0 Then
    CP.Y = CP.Y - 0.5
    CT.Y = CT.Y - 0.5
End If
If GetKeyState(vbKeyPageUp) < 0 Then
    CP.Y = CP.Y + 0.5
    CT.Y = CT.Y + 0.5
End If
If GetKeyState(vbKeyEscape) < 0 Then
    Game.bRunning = False
End If
If GetKeyState(vbKeyDelete) < 0 Then
    'CP.Y = CP.Y - 0.5
    CT.Y = CT.Y - 0.05
End If
If GetKeyState(vbKeyInsert) < 0 Then
    'CP.Y = CP.Y + 0.5
    CT.Y = CT.Y + 0.05
End If

Dim vTemp3d As D3DVECTOR 'cut
vTemp3d = Animated(0).CMesh.Get_Pos
If GetKeyState(vbKeyQ) < 0 Then
    vTemp3d.X = vTemp3d.X + 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If
If GetKeyState(vbKeyA) < 0 Then
    vTemp3d.X = vTemp3d.X - 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If
If GetKeyState(vbKeyW) < 0 Then
    vTemp3d.Y = vTemp3d.Y + 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If
If GetKeyState(vbKeyS) < 0 Then
    vTemp3d.Y = vTemp3d.Y - 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If
If GetKeyState(vbKeyE) < 0 Then
    vTemp3d.Z = vTemp3d.Z + 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If
If GetKeyState(vbKeyD) < 0 Then
    vTemp3d.Z = vTemp3d.Z - 0.1
    Animated(0).CMesh.Set_Pos vTemp3d
End If

End Sub

Public Sub MoveCamera(Speed As Single)
'------------------------------
' Movement Forward/backwards
' Moves Ct and Cp, SPEED amount
'------------------------------
c.X = CT.X - CP.X: c.Y = CT.Y - CP.Y: c.Z = CT.Z - CP.Z
CP.X = CP.X + c.X * Speed: CP.Y = CP.Y + c.Y * Speed: CP.Z = CP.Z + c.Z * Speed
CT.X = CT.X + c.X * Speed: CT.Y = CT.Y + c.Y * Speed: CT.Z = CT.Z + c.Z * Speed
End Sub

Public Sub RotateCam_XoZ(Fi As Single)
'----------------------
' Turning Left/Right
' Rotates Ct aorund Cp
'----------------------
If Fi > 2 * PI Then Fi = Fi - 2 * PI
If Fi < 0 Then Fi = 2 * PI - Fi
CT.X = CP.X + Sin(Fi) * CAM_COORD_DIST
CT.Z = CP.Z + Cos(Fi) * CAM_COORD_DIST
End Sub

Public Sub RotateCam_XoY(Fi As Single)
'---------------
' Pitch up/Down
'---------------
If Fi > PI / 4 Then Fi = PI / 4
If Fi < -PI / 4 Then Fi = -PI / 4
CT.Y = CP.Y + Cos(Fi) * Sin(Fi) * CAM_COORD_DIST
End Sub

Public Function Find_Angle(P1 As D3DVECTOR, P2 As D3DVECTOR) As Single
'----------------------------------------
' Finds the correct Angle of two Vectors
'----------------------------------------
Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single, Yv As Single
x1 = P1.X
y1 = P1.Z
x2 = P2.X
y2 = P2.Z
If (x2 - x1) <> 0 Then Yv = (Atn((y2 - y1) / (x2 - x1)))
If (((y2 - y1) < 0) And ((x2 - x1) < 0)) Or (((y2 - y1) > 0) And ((x2 - x1) < 0)) Then Yv = Yv + PI
If Yv > 2 * PI Then Yv = Yv - 2 * PI
If Yv < 0 Then Yv = 2 * PI + Yv
' NAGY IFEK
Yv = Yv + 0.0001
Find_Angle = Yv
End Function

Public Sub Write_To_File(FName As String, CMD As String)
'-----------------------------
' Used for >Debugging<(tm) :)
'-----------------------------
Open FName For Append As #7
Print #7, CMD
Close #7
End Sub

Public Sub GetWaypoints(FName As String, V As Single)
'---------------------------------------------------------------------
' This procedure reads the coordinate values from a text file
' The format of the coordinates have to be like this: ex. (12,66,22)
' V is the number of the Vector to put the data into:
'       0 - Camera Positioning Vectors
'       1 - Camera Targeting Vector
' Theese values will be interpolated between them by "PlayWithCamera"
'---------------------------------------------------------------------
Dim TS As String
Dim P1 As Single, P2 As Single, c As Single
c = 0
Open FName For Input As #8
Do While Not EOF(8)
Line Input #8, TS
P1 = InStr(1, TS, ",", vbTextCompare)
P2 = InStr(P1 + 1, TS, ",", vbTextCompare)
WAYPOINTS(c, V) = MakeVector(Mid$(TS, 2, P1 - 1), Mid$(TS, P1 + 1, P2 - P1 - 1), Mid$(TS, P2 + 1, Len(TS) - P2 - 1))
c = c + 1
Loop
If c > NWAYPOINTS Then NWAYPOINTS = c
Close #8
End Sub

Public Sub PlayWithCamera()
'-------------------------------------------------------------------
' Interpolates camera based on the Information provided in:
' WAYPOINTS(), CWAYPOINT, NWAYPOINTS and WALK_AMOUNT
' Camera is being interpolated based on two waypoint coordinates,
' and the path is "rounded" with the vector in the Next+1 waypoint.
'-------------------------------------------------------------------
If INTERPOLATEAMOUNT >= 1 Then
    CWAYPOINT = CWAYPOINT + 1
    INTERPOLATEAMOUNT = 0
    If CWAYPOINT > (NWAYPOINTS - 2) Then
        CWAYPOINT = 0
    End If
End If
CSpeed = 1 / ((Dist3D(WAYPOINTS(CWAYPOINT, 0), WAYPOINTS(CWAYPOINT + 1, 0)) + 0.000001) * Dist3D(WAYPOINTS(0, 0), WAYPOINTS(1, 0)) * WALK_AMOUNT)
INTERPOLATEAMOUNT = INTERPOLATEAMOUNT + CSpeed
Dim vTemp3d As D3DVECTOR
Dim T1 As D3DVECTOR
Dim T2 As D3DVECTOR
D3DXVec3Lerp vTemp3d, WAYPOINTS(CWAYPOINT, 0), WAYPOINTS(CWAYPOINT + 1, 0), INTERPOLATEAMOUNT
CP = vTemp3d
D3DXVec3Lerp vTemp3d, WAYPOINTS(CWAYPOINT, 1), WAYPOINTS(CWAYPOINT + 1, 1), INTERPOLATEAMOUNT
CT = vTemp3d
If ST_CAM = 1 Then
        CP = WAYPOINTS(C_A_F, 3) 'ST_CAM_P
        CT = WAYPOINTS(C_A_F, 4) 'ST_CAM_T
End If
If ST_CAM = 2 Then
        CT = Animated(0).CMesh.Get_Pos
        CP = WAYPOINTS(C_A_F, 4)
End If
End Sub

Public Sub Play_Anim_Clip(id As Byte)
'// This is how you wanted it: pass the clips no. as a param, and it will be played
ARUNNING = True
IA = 0
C_A_F = id
Select Case id
Case 1
    SV = WAYPOINTS(0, 2)
    DV = WAYPOINTS(1, 2)
Exit Sub
Case 2
    SV = WAYPOINTS(1, 2)
    DV = WAYPOINTS(2, 2)
Exit Sub
Case 3
    SV = WAYPOINTS(2, 2)
    DV = WAYPOINTS(3, 2)
Exit Sub
Case 4
    SV = WAYPOINTS(3, 2)
    DV = WAYPOINTS(4, 2)
Exit Sub
Case 5
    SV = WAYPOINTS(4, 2)
    DV = WAYPOINTS(5, 2)
Exit Sub
Case 6
    SV = WAYPOINTS(5, 2)
    DV = WAYPOINTS(6, 2)
Exit Sub
Case 7
    SV = WAYPOINTS(6, 2)
    DV = WAYPOINTS(7, 2)
Exit Sub
Case 8
    SV = WAYPOINTS(7, 2)
    DV = WAYPOINTS(7, 2)
    Set Animated(0).CMesh = Death
    Set Animated(1) = Nothing
    ARUNNING = False
    IA = 1.1
Exit Sub
End Select
End Sub

Public Sub PlayAnim()
'// Calculates the animation clip's behaviour
If ARUNNING Then
    Dim vTemp3d As D3DVECTOR
    IA = IA + 0.005
    D3DXVec3Lerp vTemp3d, SV, DV, IA
    If VecEQ(vTemp3d, Animated(0).CMesh.Get_Pos) = False Then Animated(0).CMesh.Set_Pos vTemp3d
    'We are ready
End If
ARUNNING = (IA <= 1)
End Sub
