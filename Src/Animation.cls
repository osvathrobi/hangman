VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
' This CLASS can handle Animations
' The functions are:
'       > Loading
'       > Interpolating between keyframes
'       > Rendering
'---------------------------------------------------------

Dim Frames() As New Model
Dim Speed As Single
Dim LCHK As Long
Dim FrameCount As Single
Dim CFrame As Single
Public CMesh As New Model
Dim Interpolate_Amount As Single
Dim Loaded As Boolean

Public Sub Load(FName As String)
Dim TS As String
Dim a As Single
Open FName For Input As #3 '#3 so there would not be any interferance with other classes
Line Input #3, TS '// FRAME COUNT
Line Input #3, TS
FrameCount = CInt(TS)
ReDim Frames(0 To FrameCount - 1) As New Model
Line Input #3, TS '// KEYFRAME LIST
For a = 0 To FrameCount - 1
    Line Input #3, TS
    Frames(a).Load App.Path & "\data\objects\" & TS, &HFFFFFFFF
Next a
CMesh.Load App.Path & "\data\objects\" & TS, &HFFFFFFFF
Line Input #3, TS '// // ANIM SPEED (ms)
Line Input #3, TS
Speed = CInt(TS)
Close #3
CFrame = 0
Loaded = True
End Sub

Public Sub Animate()
Dim Vec1 As D3DVECTOR
Dim Vec2 As D3DVECTOR
Dim Vec3 As D3DVECTOR
Dim a As Single
For a = 0 To Frames(CFrame).N_Vertex - 1
    Vec1 = Frames(CFrame).GetVertice(a)
    Vec2 = Frames(CFrame + 1).GetVertice(a)
    D3DXVec3Lerp Vec3, Vec1, Vec2, Interpolate_Amount
    CMesh.SetVertice a, Vec3
Next a
CMesh.UpdateMesh
End Sub

Public Sub Render()
If Loaded Then
    If GetTickCount - LCHK > Speed Then
        LCHK = GetTickCount
        CFrame = CFrame + 1
        If CFrame > FrameCount - 2 Then CFrame = 0
    End If
        Interpolate_Amount = (GetTickCount - LCHK) / Speed
        TextXY 100, 100, &HFFFFFFFF, CStr(Interpolate_Amount)
        If ARUNNING And (FrameCount > 1) Then Animate
        CMesh.Set_RotY (180 + 180 - (Find_Angle(SV, DV) * DEG))
        CMesh.Render
End If
End Sub


