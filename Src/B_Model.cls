VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------
' This CLASS can handle X-Format meshes to be used in the game
' The functions are:
'       > Loading
'       > Rendering
'       > Positioning/Rotating/Scaling
'-------------------------------------------------------------
Dim Mesh As D3DXMesh
Dim MaterialList() As D3DMATERIAL8
Dim TextureList() As Direct3DTexture8
Dim nMaterials As Long
Dim VertexList() As D3DVERTEX
Dim R1 As D3DVECTOR
Dim P1 As D3DVECTOR
Dim S1 As D3DVECTOR
Dim RP1 As D3DVECTOR
Dim K As Single

Public Sub CreateMeshFromFile(XfNAME As String)
Dim I As Long
Dim XBuffer As D3DXBuffer
Dim TextureFile As String
Dim hResult As Long
Set Mesh = D3DX.LoadMeshFromX(XfNAME, D3DXMESH_MANAGED, D3DDevice, Nothing, XBuffer, nMaterials)
ReDim MaterialList(0 To nMaterials - 1) As D3DMATERIAL8
ReDim TextureList(0 To nMaterials - 1) As Direct3DTexture8
For I = 0 To nMaterials - 1
D3DX.BufferGetMaterial XBuffer, I, MaterialList(I)
MaterialList(I).Ambient = MaterialList(I).diffuse
TextureFile = D3DX.BufferGetTextureName(XBuffer, I)
If TextureFile <> "" Then
    Set TextureList(I) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\data\tex\" & TextureFile, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, ByVal 0, ByVal 0)
End If
Next I
ReDim VertexList(0 To Mesh.GetNumVertices - 1) As D3DVERTEX
hResult = D3DXMeshVertexBuffer8GetData(Mesh, 0, Len(VertexList(0)) * Mesh.GetNumVertices, 0, VertexList(0))

S1 = MakeVector(1, 1, 1)
P1 = MakeVector(0, 0, 0)
R1 = MakeVector(0, 0, 0)
On Error GoTo bailthis
Open Mid(XfNAME, 1, Len(XfNAME) - 3) & "tex" For Input As #1
Line Input #1, TS ' ROTATION
Line Input #1, TS
R1.X = TS
Line Input #1, TS
R1.Y = TS
Line Input #1, TS
R1.Z = TS
Line Input #1, TS ' SCALE
Line Input #1, TS
S1.X = TS
Line Input #1, TS
S1.Y = TS
Line Input #1, TS
S1.Z = TS
Line Input #1, TS ' POSITION
Line Input #1, TS
P1.X = TS
Line Input #1, TS
P1.Y = TS
Line Input #1, TS
P1.Z = TS
If Not EOF(1) Then
    Line Input #1, TS ' Rotation Effect
    Line Input #1, TS
    RP1.X = TS
    Line Input #1, TS
    RP1.Y = TS
    Line Input #1, TS
    RP1.Z = TS
End If
Close #1
Exit Sub
bailthis:
End Sub

Public Sub Render()
    R1.X = R1.X + RP1.X: R1.Y = R1.Y + RP1.Y: R1.Z = R1.Z + RP1.Z
    If R1.X > 360 Then R1.X = 0: If R1.Y > 360 Then R1.Y = 0: If R1.Z > 360 Then R1.Z = 0
    
    D3DDevice.SetVertexShader FVF_NVERTEX
    Position_Mesh P1, S1, R1
    For K = 0 To (nMaterials - 1)
        D3DDevice.SetMaterial MaterialList(K)
        D3DDevice.SetTexture 0, TextureList(K)
        Mesh.DrawSubset K
    Next K
End Sub
Public Sub Render2()
    D3DDevice.SetVertexShader FVF_NVERTEX
    Position_Mesh P1, S1, R1
    For K = 0 To (nMaterials - 1)
        Mesh.DrawSubset K
    Next K
End Sub

Public Sub SetPosition(Pp As D3DVECTOR, Ss As D3DVECTOR, Rr As D3DVECTOR)
P1 = Pp
S1 = Ss
R1 = Rr
End Sub
