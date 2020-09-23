Attribute VB_Name = "modVector2"
'Taken from EGL_Dxf by Erkan Sanli
'from PlanetSourceCode

Option Explicit


Public Type tVector
    X As Single
    Y As Single
    Z As Single
    W As Single
End Type

Public Function VectorSet(X As Single, Y As Single, Z As Single) As tVector
    
    VectorSet.X = X
    VectorSet.Y = Y
    VectorSet.Z = Z
    
End Function

Public Function VectorSub(V1 As tVector, V2 As tVector) As tVector
    
    VectorSub.X = V1.X - V2.X
    VectorSub.Y = V1.Y - V2.Y
    VectorSub.Z = V1.Z - V2.Z
    
End Function

Public Function VectorAdd(V1 As tVector, V2 As tVector) As tVector
    
    VectorAdd.X = V1.X + V2.X
    VectorAdd.Y = V1.Y + V2.Y
    VectorAdd.Z = V1.Z + V2.Z
    
End Function

'------------------------------------------------
'Function:  VectorDot
'DotProduct
'Ýki noktanýn orijine göre konumlarýný belirliyor.
'------------------------------------------------
Public Function VectorDot(V1 As tVector, V2 As tVector) As Single
    
    VectorDot = (V1.X * V2.X) + _
            (V1.Y * V2.Y) + _
            (V1.Z * V2.Z)
    
End Function

'-----------------------------------------------
'Function:VectorCross
'CrossProduct
'iki nokta ve orijin ile oluþan üçgenin normali bulunuyor.
'Nasýl yaptýðýný tam olarak çözemedim.
'Geriye normalin koordinatýný döndürüyor.
'-----------------------------------------------

Public Function VectorCross(V1 As tVector, V2 As tVector) As tVector
    
    VectorCross.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
    VectorCross.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
    VectorCross.Z = (V1.X * V2.Y) - (V1.Y * V2.X)
    
End Function

'-----------------------------------
'Function: VectorNormalize
'Normalize tVector
'vektörün eðimini veriyor.Boyu 1 birim (mm) oluyor.
'-----------------------------------
Public Function VectorNormalize(V As tVector) As tVector
    
    Dim VLength As Single
    
    VLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))
    If VLength = 0 Then VLength = 1
    VectorNormalize.X = V.X / VLength
    VectorNormalize.Y = V.Y / VLength
    VectorNormalize.Z = V.Z / VLength
    
End Function

Public Function VectorScale(V As tVector, S As Single) As tVector
    
    VectorScale.X = V.X * S
    VectorScale.Y = V.Y * S
    VectorScale.Z = V.Z * S
    
End Function

'Public Sub CalculateNormal()
'    Dim I As Integer
'    With ObjPart
'        For I = 0 To .NumFaces
'            .Normal(I) = VectorCross _
'                    (VectorSub(.Vertices(.Faces(I).C), .Vertices(.Faces(I).B)), _
'                    VectorSub(.Vertices(.Faces(I).A), .Vertices(.Faces(I).B)))
'            .Normal(I) = VectorNormalize(.Normal(I))
'        Next
'        .NormalT = .Normal
'    End With
'
'End Sub

Public Function Vector(X, Y, Z) As tVector
    Vector.X = X
    Vector.Y = Y
    Vector.Z = Z
    Vector.W = 1
End Function

Public Function VectorLength(V As tVector) As Single
    
    VectorLength = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
    
End Function

Public Function VectorDist(V1 As tVector, V2 As tVector) As Single
    Dim dX As Single
    Dim dY As Single
    Dim dZ As Single
    dX = V1.X - V2.X
    dY = V1.Y - V2.Y
    dZ = V1.Z - V2.Z
    VectorDist = Sqr(dX * dX + dY * dY + dZ * dZ)
    
End Function

' def projection(self, vector):
'        k = (self.dot(vector)) / vector.length()
'        return k * vector.unit()

Public Function VectorProjection(V As tVector, Vto As tVector) As tVector
    Dim K As Single
    
    K = VectorDot(V, Vto) / VectorLength(Vto)
    
    Vto = VectorNormalize(Vto)
    
    VectorProjection = VectorScale(Vto, K)
    
End Function

