Attribute VB_Name = "modCREA"
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

Public Type tPoint
    X As Single
    Y As Single
    Z As Single
    vX As Single
    vY As Single
    vZ As Single
    
    
    
    OldX As Single
    OldY As Single
    OldZ As Single
    
    newX As Single
    newY As Single
    newZ As Single
    
    ScreenX As Long
    ScreenY As Long
    
End Type


Public Type tLink
    
    P1 As Long
    P2 As Long
    
    MLeng As Single
    
    
    TENS As Single
    
    
    DynPhase As Single
    DynAmp As Single
    DynSpeed As Single
    
    IsDynamic As Boolean
    
End Type

Public Type tVirtualFace
    
    A As Long
    B As Long
    C As Long
    
    Normal As tVector
    
    DistFromEYE As Single
    
    DrawOrder As Long
    
    color As ColorRGB
    
End Type


Public Type tMuscle
    
    L1 As Integer '     Link1
    L2 As Integer '     Link2
    
    MainAX As Double '   Angle that should be between L1 and L2
    MainAY As Double '   Angle that should be between L1 and L2
    
    P0 As Integer '     Common point of L1 and L2
    P1 As Integer '     Other point on L1
    P2 As Integer '     Other point on L2
    
    F As Double '       Muscle Force(strength)
    
    'DynPhase As Single
    'DynAmp As Single
    'DynSpeed As Single
    
    'isDynamic As Boolean
    
    'isNotBroken As Boolean
    
    
End Type

Global Const PI = 3.14159265358979

'Global Variables:---------------------------------------------------
Global Gravity As Single
Global Atmosphere As Single
Global WallBounce As Single
Global WallFriction As Single
Global WINDx As Single
Global WINDz As Single

Public Light As tVector

Public Ncre As Long
Public CRE() As New clsCREA

Public Function PointDist(P1 As tPoint, P2 As tPoint) As Single
    
    Dim dX As Single
    Dim dY As Single
    Dim dZ As Single
    
    dX = P1.X - P2.X
    dY = P1.Y - P2.Y
    dZ = P1.Z - P2.Z
    
    PointDist = Sqr(dX * dX + dY * dY + dZ * dZ)
    
    
End Function


Public Function Atan2(ByVal dX As Double, ByVal dY As Double) As Double
    'This Should return Angle
    
    Dim theta As Double
    
    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            theta = 0#
        ElseIf (dY > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(dY / dX)
        
        If (dX < 0) Then
            If (dY >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If
    
    Atan2 = theta
End Function






Public Function GetMin(A As Single, B As Single) As Single
    
    GetMin = IIf(A < B, A, B)
    
End Function



