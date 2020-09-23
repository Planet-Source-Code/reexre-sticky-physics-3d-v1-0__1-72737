VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCREA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------

Option Explicit

Public NP As Long 'Number of Points
Private P() As tPoint 'Points Array
Public NL As Long 'Number of Links
Private L() As tLink 'Links Array

Public NF As Long 'Number of Virtual Faces
Private F() As tVirtualFace

'Public NMuscles As Long 'Number of Muscles
'Private Muscle() As tMuscle 'Muscles Array

Public IntTimer As Long 'Internal Timer

Public DrawMode As Integer

Public MidPointX As Single
Public MidPointY As Single
Public MidPointZ As Single



Public Sub AddPoint(X, Y, Z)
    NP = NP + 1
    ReDim Preserve P(NP)
    
    With P(NP)
        .X = X
        .Y = Y
        .Z = Z
        
    End With
    
End Sub

Public Sub AddFace(P1, P2, p3, cR, cG, cB)
    NF = NF + 1
    ReDim Preserve F(NF)
    
    With F(NF)
        .A = P1
        .B = P2
        .C = p3
        .DrawOrder = NF
        
        .color.R = cR
        .color.G = cG
        .color.B = cB
        
    End With
    
End Sub


Public Sub AddLink(P1 As Long, P2 As Long, TENS, _
            Optional dPhase360 As Single = 0, Optional dAmp100 As Single = 50, Optional dSpeed As Single = 0)
    'TENS is the tension between the Points (or Link Stregth). MAXvalue=1
    'dPhase360 is the Dynamic Phase in Degrees
    'dAmp100 is the Dynamic (length) Amplitude: 100 Means then Maximum Length will be 2Times the MainLength and The Minimun Length will be 0
    'dSpeed is the Dynamic Speed
    
    If P1 = P2 Then Exit Sub
    
    NL = NL + 1
    ReDim Preserve L(NL)
    With L(NL)
        .P1 = P1
        .P2 = P2
        
        '.midX = (P(P1).x + P(P2).x) * 0.5
        '.midY = (P(P1).y + P(P2).y) * 0.5
        
        .MLeng = PointDist(P(P1), P(P2))
        
        .TENS = TENS
        
        If dSpeed <> 0 Then
            .IsDynamic = True
            .DynPhase = dPhase360
            
            .DynAmp = .MLeng * dAmp100 / 100
            .DynSpeed = dSpeed
        End If
        
        
    End With
    
End Sub


'Public Function ADDMuscle(L1, L2, Strength)
'
'
''Public Function ADDMuscle(L1, L2, Strength, _
'            Optional dPhase360 As Single = 0, Optional dAmp360 As Single = 20, Optional dSpeed As Single = 0) As Boolean
'    'dPhase360 is the Dynamic Phase in Degrees
'    'dAmp360 is the Dynamic Angle Amplitude in Degrees (from -Value to +Value)
'    'dSpeed is the Dynamic Speed
'
'    NMuscles = NMuscles + 1
'    ReDim Preserve Muscle(NMuscles)
'
'    ADDMuscle = False
'
'    With Muscle(NMuscles)
'        .L1 = L1
'        .L2 = L2
'        .f = Strength
'
'        If L(L1).P1 = L(L2).P1 Then
'            .P0 = L(L1).P1
'            .P1 = L(L1).P2
'            .P2 = L(L2).P2
'            ADDMuscle = True
'        End If
'
'        If L(L1).P1 = L(L2).P2 Then
'            .P0 = L(L1).P1
'            .P1 = L(L1).P2
'            .P2 = L(L2).P1
'            ADDMuscle = True
'        End If
'
'        If L(L1).P2 = L(L2).P1 Then
'            .P0 = L(L1).P2
'            .P1 = L(L1).P1
'            .P2 = L(L2).P2
'            ADDMuscle = True
'        End If
'
'        If L(L1).P2 = L(L2).P2 Then
'            .P0 = L(L1).P2
'            .P1 = L(L1).P1
'            .P2 = L(L2).P1
'            ADDMuscle = True
'        End If
'        'Stop
'
'    End With
'
'    If ADDMuscle = False Then
'        NMuscles = NMuscles - 1
'        ReDim Preserve Muscle(NMuscles)
'        Exit Function
'    End If
'
'
'
'    With Muscle(NMuscles)
'        .MainAX = Atan2(P(.P1).X - P(.P0).X, P(.P1).Z - P(.P0).Z) - _
'                Atan2(P(.P2).X - P(.P0).X, P(.P2).Z - P(.P0).Z)
'        .MainAY = Atan2(P(.P1).Y - P(.P0).Y, P(.P1).Z - P(.P0).Z) - _
'                Atan2(P(.P2).Y - P(.P0).Y, P(.P2).Z - P(.P0).Z)
'
'
'
'        If dSpeed <> 0 Then
'            .IsDynamic = True
'            .DynPhase = dPhase360
'            .DynAmp = dAmp360
'            .DynSpeed = dSpeed
'        End If
'
'    End With
'
'End Function


'Sub DoMuscles()
'    '*** DO NOT WORK!!!!
'
'    Dim X1 As Single
'    Dim y1 As Single
'    Dim z1 As Single
'
'    Dim X2 As Single
'    Dim y2 As Single
'    Dim z2 As Single
'
'    Dim AMX As Single
'    Dim AMY As Single
'
'    Dim Dx1 As Single
'    Dim Dy1 As Single
'    Dim DZ1 As Single
'
'    Dim Dx2 As Single
'    Dim Dy2 As Single
'    Dim Dz2 As Single
'
'    Dim RetP As tPoint
'
'    Dim Perc1 As Single
'    Dim Perc2 As Single
'
'    'Const Kfixed = 0.1
'    Const kMuscleSpeedLimit = 5
'
'    For M = 1 To NMuscles
'        With Muscle(M)
'            P(.P0).newX = P(.P0).X + P(.P0).vX
'            P(.P0).newY = P(.P0).Y + P(.P0).vY
'            P(.P0).newZ = P(.P0).Z + P(.P0).vZ
'
'            P(.P1).newX = P(.P1).X + P(.P1).vX
'            P(.P1).newY = P(.P1).Y + P(.P1).vY
'            P(.P1).newZ = P(.P1).Z + P(.P1).vZ
'
'            P(.P2).newX = P(.P2).X + P(.P2).vX
'            P(.P2).newY = P(.P2).Y + P(.P2).vY
'            P(.P2).newZ = P(.P2).Z + P(.P2).vZ '
'
'        End With
'    Next
'
'
'    Perc1 = 1
'    Perc2 = 1
'
'
'    For M = 1 To NMuscles
'
'        With Muscle(M)
'
'            a1 = Atan2(P(.P1).newX - P(.P0).newX, P(.P1).newZ - P(.P0).newZ)
'            a2 = Atan2(P(.P2).newX - P(.P0).newX, P(.P2).newZ - P(.P0).newZ)
'
'            AMX = (a1 - a2) - .MainAX
'
'            'If .IsDynamic Then                                  'pi/180
'            '    AMx = AMx + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp * 0.017453292
'            'End If
'
'
'            '--------------------------------------------------
'            'This Should Make Longer Link Move Slower
'            'and shorter move faster
'            '            X1 = P(.P0).newX + Cos(a1 - AMX) * L(.L1).MLeng
'            '            z1 = P(.P0).newZ + Sin(a1 - AMX) * L(.L1).MLeng
'            '
'            '            X2 = P(.P0).newX + Cos(a2 + AMX) * L(.L2).MLeng
'            '            z2 = P(.P0).newY + Sin(a2 + AMX) * L(.L2).MLeng
'            '
'            '            If L(.L1).MLeng > L(.L2).MLeng Then
'            '                Perc1 = L(.L2).MLeng / L(.L1).MLeng '(L(.L1).mleng + L(.L2).mleng)
'            '                Perc2 = 1 - Perc1
'            '            Else
'            '                Perc2 = L(.L1).MLeng / L(.L2).MLeng '(L(.L1).mleng + L(.L2).mleng)
'            '                Perc1 = 1 - Perc2
'            '            End If
'            '--------------------------------------------------
'
'            Dx1 = (P(.P1).newX - X1) * Perc1
'            DZ1 = (P(.P1).newZ - z1) * Perc1
'            Dx2 = (P(.P2).newX - X2) * Perc2
'            Dz2 = (P(.P2).newZ - z2) * Perc2
'
'
'            P(.P1).vX = P(.P1).vX - .F * Dx1
'            P(.P1).vZ = P(.P1).vZ - .F * DZ1
'            P(.P2).vX = P(.P2).vX - .F * Dx2
'            P(.P2).vZ = P(.P2).vZ - .F * Dz2
'           P(.P0).vX = P(.P0).vX + .F * (Dx1 + Dx2)
'           P(.P0).vZ = P(.P0).vZ + .F * (DZ1 + Dz2)
'
'           '-***************************************************
'
'
'            a1 = Atan2(P(.P1).newY - P(.P0).newY, P(.P1).newZ - P(.P0).newZ)
'            a2 = Atan2(P(.P2).newY - P(.P0).newY, P(.P2).newZ - P(.P0).newZ)
'
'            AMY = (a1 - a2) - .MainAY
'
'            'If .IsDynamic Then                                  'pi/180
'            '    AMy = AMy + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp * 0.017453292
'            'End If
'
'
'            '--------------------------------------------------
'            'This Should Make Longer Link Move Slower
'            'and shorter move faster
'            '            y1 = P(.P0).newY + Cos(a1 - AMY) * L(.L1).MLeng
'            '            z1 = P(.P0).newZ + Sin(a1 - AMY) * L(.L1).MLeng
'            '
'            '            y2 = P(.P0).newY + Cos(a2 + AMY) * L(.L2).MLeng
'            '            z2 = P(.P0).newY + Sin(a2 + AMY) * L(.L2).MLeng
'            '
'           '            If L(.L1).MLeng > L(.L2).MLeng Then
'           '                Perc1 = L(.L2).MLeng / L(.L1).MLeng '(L(.L1).mleng + L(.L2).mleng)
'           '                Perc2 = 1 - Perc1
'           '            Else
'           '                Perc2 = L(.L1).MLeng / L(.L2).MLeng '(L(.L1).mleng + L(.L2).mleng)
'           '                Perc1 = 1 - Perc2
'           '            End If
'           '--------------------------------------------------
'
'           Dy1 = (P(.P1).newY - y1) * Perc1
'           DZ1 = (P(.P1).newZ - z1) * Perc1
'           Dy2 = (P(.P2).newY - y2) * Perc2
'           Dz2 = (P(.P2).newZ - z2) * Perc2
'
'
'           P(.P1).vY = P(.P1).vY - .F * Dy1
'           P(.P1).vZ = P(.P1).vZ - .F * DZ1
'           P(.P2).vY = P(.P2).vY - .F * Dy2
'           P(.P2).vZ = P(.P2).vZ - .F * Dz2
'           P(.P0).vY = P(.P0).vY + .F * (Dy1 + Dy2)
'           P(.P0).vZ = P(.P0).vZ + .F * (DZ1 + Dz2)
'
'       End With
'
'   Next
'
'End Sub
Public Sub DoPhysics()
    
    
    Dim Xer  As Single
    Dim Yer  As Single
    Dim Zer  As Single
    
    Dim newX  As Single
    Dim newY As Single
    
    Dim Leng As Double 'Single
    
    Dim Leng2Go_X As Single
    Dim Leng2Go_Y As Single
    Dim Leng2Go_Z As Single
    
    
    Dim LengthTotal As Single
    Dim TempTime As Single
    Dim Fric
    
    Dim I As Long
    Dim P1 As Long
    Dim P2 As Long
    
    
    
    IntTimer = IntTimer + 1
    
    
    
    For I = 1 To NL
        
        With L(I)
            
            
            P1 = .P1
            P2 = .P2
            
            'This will calculate the links 'true' length.  That is the distance
            'between its 2 vertices.   The length stored in .linklength is
            'the length that the link 'should' be.  The link will push or pull
            'to bring the vertices back to that distance.
            LengthTotal = .MLeng
            
            Xer = (P(P2).X + P(P2).vX) - (P(P1).X + P(P1).vX)
            Yer = (P(P2).Y + P(P2).vY) - (P(P1).Y + P(P1).vY)
            Zer = (P(P2).Z + P(P2).vZ) - (P(P1).Z + P(P1).vZ)
            
            
            Leng = Sqr(Xer * Xer + Yer * Yer + Zer * Zer)
            
            
            If .IsDynamic Then
                'pi/180
                Leng = Leng + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp
                '                Stop
                
            End If
            
            
            
            
            'Calculate how far in each direction the vertices must go
            'in order to get the link back to its regular length
            Leng2Go_X = ((Leng - LengthTotal) / Leng) * Xer
            Leng2Go_Y = ((Leng - LengthTotal) / Leng) * Yer
            Leng2Go_Z = ((Leng - LengthTotal) / Leng) * Zer
            '            Stop
            
            
            P(P1).vX = P(P1).vX + (Leng2Go_X / 2) * .TENS
            P(P1).vY = P(P1).vY + (Leng2Go_Y / 2) * .TENS
            P(P1).vZ = P(P1).vZ + (Leng2Go_Z / 2) * .TENS
            
            P(P2).vX = P(P2).vX + (Leng2Go_X / 2) * -1 * .TENS
            P(P2).vY = P(P2).vY + (Leng2Go_Y / 2) * -1 * .TENS
            P(P2).vZ = P(P2).vZ + (Leng2Go_Z / 2) * -1 * .TENS
            
            
            
        End With
        
    Next
    
    MidPointX = 0
    MidPointY = 0
    MidPointZ = 0
    
    For I = 1 To NP
        
        With P(I)
            
            .vY = .vY - Gravity
            .vX = .vX + WINDx
            .vZ = .vZ + WINDz
            
            
            
            .vX = .vX * (1 - Atmosphere)
            .vY = .vY * (1 - Atmosphere)
            .vZ = .vZ * (1 - Atmosphere)
            
            
            .OldX = .X
            .OldY = .Y
            .OldZ = .Z
            
            .X = .X + .vX
            .Y = .Y + .vY
            .Z = .Z + .vZ
            
            MidPointX = MidPointX + .X
            MidPointY = MidPointY + .Y
            MidPointZ = MidPointZ + .Z
            
            
            If .Y < 0 Then 'floor
                
                '.Y = .OldY
                .Y = 0.001
                .vY = -(.vY * WallBounce)
                .vX = .vX * (1 - WallFriction)
                .vZ = .vZ * (1 - WallFriction)
            End If
            
            
        End With
    Next
    
    MidPointX = MidPointX / NP
    MidPointY = MidPointY / NP
    MidPointZ = MidPointZ / NP
    
End Sub




Public Sub Convert3DToScreen()
    Dim ScreenPoint As POINTAPI
    Dim V As tVector
    Dim I As Long
    
    UpdateCamera
    
    For I = 1 To NP
        
        V.X = P(I).X
        V.Y = P(I).Y
        V.Z = P(I).Z
        
        ScreenPoint = PointToScreen(V)
        
        P(I).ScreenX = ScreenPoint.X
        P(I).ScreenY = ScreenPoint.Y
        
    Next
    
End Sub
Private Sub DRAW_Links(picHdc As Long)
    Dim X1 As Long
    Dim y1 As Long
    
    Dim X2 As Long
    Dim y2 As Long
    
    Dim I As Long
    
    For I = 1 To NL
        
        X1 = P(L(I).P1).ScreenX
        y1 = P(L(I).P1).ScreenY
        
        X2 = P(L(I).P2).ScreenX
        y2 = P(L(I).P2).ScreenY
        
        If L(I).IsDynamic Then
            FastLine picHdc, X1, y1, X2, y2, 2, vbGreen
        Else
            FastLine picHdc, X1, y1, X2, y2, 2, vbWhite
        End If
        
    Next I
End Sub

Public Sub DRAW(Mode_0or1 As Integer, pHDC As Long)
    
    DrawMode = Mode_0or1
    
    If DrawMode = 0 Then DRAW_AllFaces pHDC Else DRAW_Links pHDC
    
End Sub

Private Sub DRAW_AllFaces(picHdc As Long)
    
    
    Dim I As Long
    
    Faces_CalculateDistances
    
    Faces_Sort 1, NF
    
    For I = 1 To NF
        
        Draw_SingleFace F(I).DrawOrder, picHdc
        
    Next
    
End Sub

Private Sub Draw_SingleFace(wF As Long, hDC As Long)
    'Taken from EGL_Dxf from PlanetSourceCode
    
    
    Dim DrawColorRGB As ColorRGB
    Dim DrawColorLong As Long
    
    Dim FaceLight As Single
    Dim FaceColor As ColorRGB
    
    Dim PenSelect As Long
    Dim BrushSelect As Long
    
    Dim IsVisible As Boolean
    
    
    Dim tmp(2) As POINTAPI
    
    tmp(0).X = P(F(wF).A).ScreenX
    tmp(0).Y = P(F(wF).A).ScreenY
    
    tmp(1).X = P(F(wF).B).ScreenX
    tmp(1).Y = P(F(wF).B).ScreenY
    
    tmp(2).X = P(F(wF).C).ScreenX
    tmp(2).Y = P(F(wF).C).ScreenY
    
    
    Face_CalculateNormal wF
    
    FaceLight = VectorDot(F(wF).Normal, Light)
    
    IsVisible = IIf(VectorDot(F(wF).Normal, VectorSub(camera.cFrom, camera.cTo)) > 0, True, False)
    
    If FaceLight < 0 Then FaceLight = -FaceLight
    'If FaceLight > 0 Then
    If IsVisible Then
                
        FaceColor = F(wF).color
        
        DrawColorRGB = ColorScale(FaceColor, FaceLight)
        DrawColorRGB = ColorScale(DrawColorRGB, 1.2)
        DrawColorLong = ColorRGBToLong(DrawColorRGB)
        
        'Case SolidFrame
        PenSelect = SelectObject(hDC, CreatePen(0, 1, vbWhite)) 'BackgroundColor(0)))
        'BrushSelect = SelectObject(hDC, CreateSolidBrush(lngColor))
        BrushSelect = SelectObject(hDC, CreateSolidBrush(DrawColorLong))
        Polygon hDC, tmp(0), 3
        
        'Case Solid
        '            PenSelect = SelectObject(hdc, CreatePen(PS_SOLID, 1, lngColor))
        '            BrushSelect = SelectObject(hdc, CreateSolidBrush(lngColor))
        '            Polygon hdc, tmp(0), 3
        
        DeleteObject PenSelect
        DeleteObject BrushSelect
    End If
    
End Sub



Public Sub Kill()
    Dim I As Long
    For I = 1 To NP
        P(I).vX = 0
        P(I).vY = 0
        P(I).vZ = 0
    Next
    
    IntTimer = 0
    
    NP = 0
    NL = 0
    NF = 0
    'NMuscles = 0
End Sub


Public Sub Face_CalculateNormal(wF)
    'Taken from EGL_Dxf from PlanetSourceCode
    
    Dim N As tVector
    Dim vA As tVector
    Dim vB As tVector
    Dim vC As tVector
    
    With F(wF)
        
        vA.X = P(.A).X
        vA.Y = P(.A).Y
        vA.Z = P(.A).Z
        
        vB.X = P(.B).X
        vB.Y = P(.B).Y
        vB.Z = P(.B).Z
        
        vC.X = P(.C).X
        vC.Y = P(.C).Y
        vC.Z = P(.C).Z
        
        N = VectorCross(VectorSub(vC, vB), VectorSub(vA, vB))
        
        .Normal = VectorNormalize(N)
        
    End With
    
End Sub

Private Sub Faces_CalculateDistances()
    Dim faceMid As tVector
    Dim I As Long
    
    Dim dA As Single
    Dim dB As Single
    Dim dC As Single
    
    For I = 1 To NF
        
        'With F(I)
        '    faceMid.X = (P(.a).X + P(.b).X + P(.C).X) / 3
        '    faceMid.Y = (P(.a).Y + P(.b).Y + P(.C).Y) / 3
        '    faceMid.Z = (P(.a).Z + P(.b).Z + P(.C).Z) / 3
        'End With
        
        'F(I).DistFromEYE = -VectorDist(faceMid, camera.cFrom)
        
        With F(I)
            dA = VectorDist(Vector(P(.A).X, P(.A).Y, P(.A).Z), camera.cFrom)
            dB = VectorDist(Vector(P(.B).X, P(.B).Y, P(.B).Z), camera.cFrom)
            dC = VectorDist(Vector(P(.C).X, P(.C).Y, P(.C).Z), camera.cFrom)
            
            F(I).DistFromEYE = GetMin(GetMin(-dA, -dB), -dC)
            
            
        End With
        
        
    Next
    
End Sub

Private Sub Faces_Sort(ByVal First As Long, ByVal Last As Long)
    'Taken from EGL_Dxf by Erkan Sanli
    'from PlanetSourceCode
    
    Dim FirstIdx  As Long
    Dim MidIdx As Long
    Dim LastIdx  As Long
    Dim MidVal As Single
    Dim TempOrder  As Long 'Order
    
    If (First < Last) Then
        
        MidIdx = (First + Last) \ 2
        MidVal = F(F(MidIdx).DrawOrder).DistFromEYE
        FirstIdx = First
        LastIdx = Last
        Do
            Do While F(F(FirstIdx).DrawOrder).DistFromEYE < MidVal
                FirstIdx = FirstIdx + 1
            Loop
            Do While F(F(LastIdx).DrawOrder).DistFromEYE > MidVal
                LastIdx = LastIdx - 1
            Loop
            If (FirstIdx <= LastIdx) Then
                TempOrder = F(LastIdx).DrawOrder
                F(LastIdx).DrawOrder = F(FirstIdx).DrawOrder
                F(FirstIdx).DrawOrder = TempOrder
                FirstIdx = FirstIdx + 1
                LastIdx = LastIdx - 1
            End If
        Loop Until FirstIdx > LastIdx
        
        If (LastIdx <= MidIdx) Then
            Faces_Sort First, LastIdx
            Faces_Sort FirstIdx, Last
        Else
            Faces_Sort FirstIdx, Last
            Faces_Sort First, LastIdx
        End If
    End If
    
End Sub
