VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
' This CLASS can handle mesh files to be used in the game
' The functions are:
'       > Loading
'       > Rendering
'       > Positioning
'       > Read/Write Vertex Data
'       > Update VertexBuffers
'       > Set Angle-Facing
'---------------------------------------------------------

Option Explicit
Public Name As String

Dim Indices() As Integer
Dim VBuffer As Direct3DVertexBuffer8
Dim IBuffer As Direct3DIndexBuffer8
Dim p_Count As Single
Dim V() As VERTEX
Dim q As Single
Dim P As Single
Public N_Vertex As Long
Dim N_Faces As Long
Dim S_Line As Boolean
Dim a As Single
Dim c As Single
Dim d As Single
Dim Y As Single
Dim LastCHK As Long
Dim Texture As Direct3DTexture8
Dim R1 As D3DVECTOR
Dim P1 As D3DVECTOR
Dim S1 As D3DVECTOR

Private Function Get_Param(S As String, pName As String) As String
P = 0
P = InStr(1, S, pName, vbTextCompare)
If P <> 0 Then
    q = InStr(P + 1 + Len(pName), S, " ", vbTextCompare)
    If (q <> 0) Then
        Get_Param = Mid$(S, P + Len(pName), q - (P + Len(pName)))
    Else
        Get_Param = Mid$(S, P + Len(pName), 1 + Len(S) - (P + Len(pName)))
    End If
Else
    Get_Param = "0"
End If
End Function

Public Sub Load(FName As String, ByVal c As Long)
Dim TS As String
Open FName For Input As #1
Line Input #1, TS
Line Input #1, TS
Do While InStr(1, TS, "Named object", vbTextCompare) <> 1
    Line Input #1, TS
Loop
'Named object
Name = Get_Param(TS, "Named object: ")
Name = Mid$(Name, 2, Len(Name) - 2)

Line Input #1, TS
N_Vertex = CStr(Get_Param(TS, "Vertices:"))
N_Faces = CStr(Get_Param(TS, "Faces:"))
Line Input #1, TS
S_Line = False
If TS = "Mapped" Then
    Line Input #1, TS
    S_Line = True
End If

ReDim V(0 To N_Vertex - 1) As VERTEX
For a = 0 To N_Vertex - 1
    Line Input #1, TS
    With V(a)
        .X = (Get_Param(TS, "X:"))
        .Z = (Get_Param(TS, "Y:"))
        .Y = (Get_Param(TS, "Z:"))
        .TX = (Get_Param(TS, "U:"))
        .TY = 1 - (Get_Param(TS, "V:")) 'corrects texturing coordinates
        If c = 0 Then
            .COLOR = &HFFFFFFFF
        Else
            .COLOR = c
        End If
    End With
Next a
Line Input #1, TS
c = 0
ReDim Indices(0 To (N_Faces) * 3) As Integer
For a = 0 To N_Faces - 1
    Do
    Line Input #1, TS
    Loop Until (InStr(1, TS, "Face", vbTextCompare) = 1)
    Indices(c) = CInt(Get_Param(TS, "A:"))
    c = c + 1
    Indices(c) = CInt(Get_Param(TS, "B:"))
    c = c + 1
    Indices(c) = CInt(Get_Param(TS, "C:"))
    c = c + 1
Next a

Set VBuffer = D3DDevice.CreateVertexBuffer(N_Vertex * Len(V(0)), 0, FVF_VERTEX, D3DPOOL_DEFAULT)
D3DVertexBuffer8SetData VBuffer, 0, N_Vertex * Len(V(0)), 0, V(0)
Set IBuffer = D3DDevice.CreateIndexBuffer(N_Faces * 3 * Len(Indices(0)), 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
D3DIndexBuffer8SetData IBuffer, 0, (N_Faces * 3) * Len(Indices(0)), 0, Indices(0)

Close #1

'// Setting up some default values
S1 = MakeVector(0.5, 0.5, 0.5)
P1 = MakeVector(0, 0, 0)
R1 = MakeVector(0, 0, 0)

'// If the correct .TEX file is found, read the data
On Error GoTo BailOpen
Open Mid(FName, 1, Len(FName) - 3) & "tex" For Input As #1
Line Input #1, TS
If TS <> "" Then Set Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\data\tex\" & TS, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, ByVal 0, ByVal 0)
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
Close #1
BailOpen:

End Sub

Public Sub Render()
D3DDevice.SetVertexShader FVF_VERTEX
Position_Mesh P1, S1, R1
D3DDevice.SetTexture 0, Texture
D3DDevice.SetStreamSource 0, VBuffer, Len(V(0))
D3DDevice.SetIndices IBuffer, 0
D3DDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, N_Faces * 3, 0, N_Faces
End Sub

Public Function GetVertice(id As Single) As D3DVECTOR
GetVertice.X = V(id).X: GetVertice.Y = V(id).Y: GetVertice.Z = V(id).Z
End Function
Public Sub SetVertice(id As Single, cV As D3DVECTOR)
V(id).X = cV.X: V(id).Y = cV.Y: V(id).Z = cV.Z
End Sub
Public Sub UpdateMesh()
D3DVertexBuffer8SetData VBuffer, 0, N_Vertex * Len(V(0)), 0, V(0)
End Sub
Public Sub Set_Pos(Pp As D3DVECTOR)
P1 = Pp
End Sub
Public Sub Set_Scale(Sp As D3DVECTOR)
S1 = Sp
End Sub
Public Sub Set_RotY(q As Single)
R1.Y = q
End Sub
Public Function Get_Pos() As D3DVECTOR
Get_Pos = P1
End Function
