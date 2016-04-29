Attribute VB_Name = "MediaLoading"
'--------------------------------------------------------------------------------------
' In this module I will write a memory manager,
' so I could keep track of the object being loaded.
' The manager will be "slot" based.
' There will be:
'       1 static scenery mesh (town bulidings and the skybox)
'       2 animated character
'       1 animated object
'--------------------------------------------------------------------------------------

Public Scenery As New XModel 'Complex Scenery
Public Animated(0 To 1) As New Animation  'Animated scenery meshes
Public Ambient(0) As New XModel 'Animated Scenery model
Public Death As New Model

Public Sub Load(SceneSet As Byte)
'-------------------------------------
'Loads the selected scene set (1 to 8)
'-------------------------------------
Select Case SceneSet
Case 1
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc1.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    WALK_AMOUNT = 0.2  '0.5
    GetWaypoints App.Path & "\data\motion\Mo1.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT1.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_1.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_1_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_1_CT.txt", 4
    ST_CAM_P = MakeVector(12.52815, -6.225001, -53.50978)
    ST_CAM_T = MakeVector(13.09279, -6.475002, -52.68444)
    Death.Set_Pos MakeVector(26, -12, -37)
Exit Sub
Case 2
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc2.xsc"
    Ambient(0).CreateMeshFromFile App.Path & "\data\objects\Sc2b.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    WALK_AMOUNT = 0.05 '0.5
    GetWaypoints App.Path & "\data\motion\Mo2.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT2.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_2.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_2_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_2_CT.txt", 4
    Create_Parametric_Fog True, &HF8EA54, 20, 500, False, D3DFOG_LINEAR, D3DFOG_NONE
    ST_CAM_P = MakeVector(5.768689, 9.559979, 84.10186)
    ST_CAM_T = MakeVector(5.29667, 9.359979, 83.22027)
    Death.Set_Pos MakeVector(0, 8.399998, 66.40002)
Exit Sub
Case 3
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc3.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    WALK_AMOUNT = 0.7 '0.5
    GetWaypoints App.Path & "\data\motion\Mo3.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT3.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_3.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_3_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_3_CT.txt", 4
    Create_Parametric_Fog False, 0, 0, 0, False, D3DFOG_NONE, D3DFOG_NONE
    ST_CAM_P = MakeVector(-85.2048, -3.130016, 6.906557)
    ST_CAM_T = MakeVector(-84.3679, -3.680019, 6.3592)
    Animated(0).CMesh.Set_Scale MakeVector(0.1, 0.1, 0.1)
    Death.Set_Scale MakeVector(0.1, 0.1, 0.1)
    Death.Set_Pos MakeVector(-79.2, -7.799993, -1)
Exit Sub
Case 4
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc4.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d.asc", &HFFFFFFFF
    Create_Parametric_Fog True, &HF8EA54, 150, 650, True, D3DFOG_LINEAR, D3DFOG_NONE
    GetWaypoints App.Path & "\data\motion\Mo4.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT4.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_4.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_4_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_4_CT.txt", 4
    WALK_AMOUNT = 0.06
    ST_CAM_P = MakeVector(81.3639, 139.545, -17.64091)
    ST_CAM_T = MakeVector(80.3664, 139.245, -17.57018)
    Death.Set_Pos MakeVector(-1.300002, -8.299999, -42.69997)
Exit Sub
Case 5
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc5.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    GetWaypoints App.Path & "\data\motion\Mo5.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT5.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_5.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_5_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_5_CT.txt", 4
    WALK_AMOUNT = 0.1
    Create_Parametric_Fog True, &HF8EA54, 20, 900, True, D3DFOG_LINEAR, D3DFOG_NONE
    ST_CAM_P = MakeVector(-13.32054, -5.5, -57.6529)
    ST_CAM_T = MakeVector(-12.6389, -5.5, -56.92122)
    Death.Set_Pos MakeVector(-1.300002, -3.900002, -42.69997)
Exit Sub
Case 6
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc6.xsc"
    Animated(1).Load App.Path & "\data\objects\Test5.ANM" 'cowboy
    Animated(0).Load App.Path & "\data\objects\Test4.ANM" 'horse
    Animated(1).CMesh.Set_Pos MakeVector(-7.899996, 6.499996, 11.30001)
    Animated(0).CMesh.Set_Pos MakeVector(-6.899996, -13.10001, 11.30001)
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    Death.Set_Pos MakeVector(-6.899996, 6.499996, 11.30001)
    GetWaypoints App.Path & "\data\motion\Mo6.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT6.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_6.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_6_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_6_CT.txt", 4
    WALK_AMOUNT = 0.06
    ST_CAM_P = MakeVector(17.49784, 1.170451, 15.74823)
    ST_CAM_T = MakeVector(16.75215, 1.320479, 15.08194)
Exit Sub
Case 7
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc7.xsc"
    Animated(1).Load App.Path & "\data\objects\Test2.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d.asc", &HFFFFFFFF
    Animated(0).Load App.Path & "\data\objects\Test3.ANM"
    GetWaypoints App.Path & "\data\motion\Mo7.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT7.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_7.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_7_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_7_CT.txt", 4
    WALK_AMOUNT = 0.04
    ST_CAM_P = MakeVector(110.6825, -8.160007, 9.9662)
    ST_CAM_T = MakeVector(110.2853, -7.710005, 9.048445)
    Animated(1).CMesh.Set_Pos MakeVector(108.6989, -6.800001, -35.10004)
    Death.Set_Pos MakeVector(108.6989, -11.800001, -35.10004)
Exit Sub
Case 8
    Scenery.CreateMeshFromFile App.Path & "\data\objects\Sc8.xsc"
    Animated(0).Load App.Path & "\data\objects\Test.ANM"
    Death.Load App.Path & "\data\objects\anim\F_d7.asc", &HFFFFFFFF
    Create_Parametric_Fog True, &H222222, 20, 900, True, D3DFOG_LINEAR, D3DFOG_NONE
    GetWaypoints App.Path & "\data\motion\Mo8.txt", 0
    GetWaypoints App.Path & "\data\motion\MoT8.txt", 1
    GetWaypoints App.Path & "\data\motion\Chr_8.txt", 2
    GetWaypoints App.Path & "\data\motion\Chr_8_CP.txt", 3
    GetWaypoints App.Path & "\data\motion\Chr_8_CT.txt", 4
    WALK_AMOUNT = 0.2
    ST_CAM_P = MakeVector(-96.93266, -36, -55.27393)
    ST_CAM_T = MakeVector(-96.25101, -36, -54.54224)
    Death.Set_Pos MakeVector(-208.6996, -92.70013, 139.7999)
Exit Sub
End Select
End Sub
