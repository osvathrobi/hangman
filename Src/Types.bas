Attribute VB_Name = "Types"
'----------------------------------------
'Module which groups all Global variables
'----------------------------------------
Type VERTEX
    X As Single
    Y As Single
    Z As Single
    COLOR As Long
    TX As Single
    TY As Single
End Type

Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const FVF_NVERTEX = (D3DFVF_NORMAL Or D3DFVF_TEX1 Or D3DFVF_XYZ)
Public Const WALK_SPEED = 0.3
Public Const TURN_SPEED = 0.05
Public Const CAM_COORD_DIST = 1

Public CP As D3DVECTOR 'CAMERA POSITION
Public CT As D3DVECTOR 'CAMERA TARGET
Public c As D3DVECTOR 'TEMP VECTOR FOR CAMERA
Public RP As Single 'ROTATION PARAMETER TEMP FOR TURNING
Public WAYPOINTS(0 To 80, 0 To 4) As D3DVECTOR 'WAIPOINTS FOR CAMERA
Public CWAYPOINT As Single 'CURRENT WAYPOINT
Public CSpeed As Single
Public NWAYPOINTS As Single 'NUMBER OF CAMERA WAIPOINTS
Public INTERPOLATEAMOUNT As Single
Public INTERPOLATEAMOUNT2 As Single
Public WALK_AMOUNT As Single
Public ST_CAM_P As D3DVECTOR
Public ST_CAM_T As D3DVECTOR
Public ST_CAM As Byte

'Animation Props
Public SV As D3DVECTOR 'Source Vector
Public DV As D3DVECTOR 'Source Vector
Public IA As Single 'Interpolate Amount
Public ARUNNING As Boolean 'Animation Running?
Public C_A_F As Byte 'Current Animation Frame
