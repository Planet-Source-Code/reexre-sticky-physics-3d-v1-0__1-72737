Attribute VB_Name = "modCAMERA2"
'Taken From Here:
'http://local.wasp.uwa.edu.au/~pbourke/miscellaneous/transform/
'by Paul Bourke
'TRANSLATED BY ME
'[reexre]
'
'
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


'/* Camera definition */
'typedef struct {
'   XYZ from;
'   XYZ to;
'   XYZ up;
'   double angleh,anglev;
'   double zoom;
'   double NearPlane,FarPlane;
'   short projection;
'} CAMERA;
'
'#define DTOR 0.01745329252
'#define EPSILON 0.001
'#define PERSPECTIVE 0
'#define ORTHOGRAPHIC 1
Option Explicit

Public Enum eProjection
    PERSPECTIVE
    ORTHOGRAPHIC
End Enum

Public Type tCamera
    
    
    cFrom As tVector
    cTo As tVector
    cUp As tVector
    
    ANGh As Single
    ANGv As Single
    Zoom As Single
    NearPlane As Single
    FarPlane As Single
    Projection As eProjection
    
End Type

Public Type tScreen
    Center As tVector
    Size As tVector
End Type

Public camera As tCamera
Public Scree As tScreen

Public Const Epsilon = 0.001
Public Const DTOR = 0.01745329252


'/* Static globals */
'double tanthetah,tanthetav;
'XYZ basisa,basisb,basisc;

Private basisA As tVector
Private basisB As tVector
Private basisC As tVector

Private tanthetaH As Single
Private tanthetaV As Single

Public Function UpdateCamera()
    
    basisB.X = camera.cTo.X - camera.cFrom.X
    basisB.Y = camera.cTo.Y - camera.cFrom.Y
    basisB.Z = camera.cTo.Z - camera.cFrom.Z
    
    basisB = VectorNormalize(basisB)
    
    'CrossProduct(camera.up,basisb,&basisa);
    'Normalise(&basisa);
    
    
    basisA = VectorCross(camera.cUp, basisB)
    basisA = VectorNormalize(basisA)
    
    
    '   CrossProduct(basisb,basisa,&basisc);
    basisC = VectorCross(basisB, basisA)
    
    '   /* Calculate camera aperture statics, note: angles in degrees */
    '   tanthetah = tan(camera.angleh * DTOR / 2);
    '   tanthetav = tan(camera.anglev * DTOR / 2);
    
    tanthetaH = Tan(camera.ANGh * DTOR * 0.5)
    tanthetaV = Tan(camera.ANGv * DTOR * 0.5)
    
    
End Function

' Take a point in world coordinates and transform it to
'   a point in the eye coordinate system.

'void Trans_World2Eye(W, e, CAMERA)
'XYZ w;
'XYZ *e;
'CAMERA camera;
'{
'   /* Translate world so that the camera is at the origin */
'   w.x -= camera.from.x;
'   w.y -= camera.from.y;
'   w.z -= camera.from.z;
'
'   /* Convert to eye coordinates using basis vectors */
'   e->x = w.x * basisa.x + w.y * basisa.y + w.z * basisa.z;
'   e->y = w.x * basisb.x + w.y * basisb.y + w.z * basisb.z;
'   e->z = w.x * basisc.x + w.y * basisc.y + w.z * basisc.z;
'}
Public Function World2EYE(W As tVector) As tVector
    
    W.X = W.X - camera.cFrom.X
    W.Y = W.Y - camera.cFrom.Y
    W.Z = W.Z - camera.cFrom.Z
    
    World2EYE.X = W.X * basisA.X + W.Y * basisA.Y + W.Z * basisA.Z
    World2EYE.Y = W.X * basisB.X + W.Y * basisB.Y + W.Z * basisB.Z
    World2EYE.Z = W.X * basisC.X + W.Y * basisC.Y + W.Z * basisC.Z
    
    
    
End Function

' Clip a line segment in eye coordinates to the camera NearPlane
'   and FarPlane clipping planes. Return FALSE if the line segment
'   is entirely before or after the clipping planes.
'*/
'int Trans_ClipEye(e1,e2,camera)
'XYZ *e1,*e2;
'CAMERA camera;
'{
'   double mu;
'
'   /* Is the vector totally in NearPlane of the NearPlane cutting plane ? */
'   if (e1->y <= camera.NearPlane && e2->y <= camera.NearPlane)
'      return(FALSE);
'
'   /* Is the vector totally behind the FarPlane cutting plane ? */
'   if (e1->y >= camera.FarPlane && e2->y >= camera.FarPlane)
'      return(FALSE);
'
'   /* Is the vector partly in NearPlane of the NearPlane cutting plane ? */
'   if ((e1->y < camera.NearPlane && e2->y > camera.NearPlane) ||
'      (e1->y > camera.NearPlane && e2->y < camera.NearPlane)) {
'      mu = (camera.NearPlane - e1->y) / (e2->y - e1->y);
'      if (e1->y < camera.NearPlane) {
'         e1->x = e1->x + mu * (e2->x - e1->x);
'         e1->z = e1->z + mu * (e2->z - e1->z);
'         e1->y = camera.NearPlane;
'      } else {
'         e2->x = e1->x + mu * (e2->x - e1->x);
'         e2->z = e1->z + mu * (e2->z - e1->z);
'         e2->y = camera.NearPlane;
'      }
'   }
'
'   /* Is the vector partly behind the FarPlane cutting plane ? */
'   if ((e1->y < camera.FarPlane && e2->y > camera.FarPlane) ||
'      (e1->y > camera.FarPlane && e2->y < camera.FarPlane)) {
'      mu = (camera.FarPlane - e1->y) / (e2->y - e1->y);
'      if (e1->y < camera.FarPlane) {
'         e2->x = e1->x + mu * (e2->x - e1->x);
'         e2->z = e1->z + mu * (e2->z - e1->z);
'         e2->y = camera.FarPlane;
'      } else {
'         e1->x = e1->x + mu * (e2->x - e1->x);
'         e1->z = e1->z + mu * (e2->z - e1->z);
'         e1->y = camera.FarPlane;
'      }
'   }
'
'   return(TRUE);
'}

Public Function ClipEYE(E1 As tVector, E2 As tVector) As Boolean
    
    Dim Mu As Double
    '   /* Is the vector totally in NearPlane of the NearPlane cutting plane ? */
    If (E1.Y <= camera.NearPlane And E2.Y <= camera.NearPlane) Then ClipEYE = False: Exit Function
    
    '   /* Is the vector totally behind the FarPlane cutting plane ? */
    If (E1.Y >= camera.FarPlane And E2.Y >= camera.FarPlane) Then ClipEYE = False: Exit Function
    
    '   /* Is the vector partly in NearPlane of the NearPlane cutting plane ? */
    If ((E1.Y < camera.NearPlane And E2.Y > camera.NearPlane) Or _
            (E1.Y > camera.NearPlane And E2.Y < camera.NearPlane)) Then
    Mu = (camera.NearPlane - E1.Y) / (E2.Y - E1.Y)
    If (E1.Y < camera.NearPlane) Then
        E1.X = E1.X + Mu * (E2.X - E1.X)
        E1.Z = E1.Z + Mu * (E2.Z - E1.Z)
        E1.Y = camera.NearPlane
    Else
        E2.X = E1.X + Mu * (E2.X - E1.X)
        E2.Z = E1.Z + Mu * (E2.Z - E1.Z)
        E2.Y = camera.NearPlane
    End If
End If

'   /* Is the vector partly behind the FarPlane cutting plane ? */
If ((E1.Y < camera.FarPlane And E2.Y > camera.FarPlane) Or _
        (E1.Y > camera.FarPlane And E2.Y < camera.FarPlane)) Then
Mu = (camera.FarPlane - E1.Y) / (E2.Y - E1.Y)
If (E1.Y < camera.FarPlane) Then
    E2.X = E1.X + Mu * (E2.X - E1.X)
    E2.Z = E1.Z + Mu * (E2.Z - E1.Z)
    E2.Y = camera.FarPlane
Else
    E1.X = E1.X + Mu * (E2.X - E1.X)
    E1.Z = E1.Z + Mu * (E2.Z - E1.Z)
    E1.Y = camera.FarPlane
End If
End If

ClipEYE = True

End Function


'/*
'   Take a vector in eye coordinates and transform it into
'   normalised coordinates for a perspective view. No normalisation
'   is performed for an orthographic projection. Note that although
'   the y component of the normalised vector is copied from the eye
'   coordinate system, it is generally no longer needed. It can
'   however still be used externally for vector sorting.
'*/
'void Trans_Eye2Norm(e, n, CAMERA)
'XYZ e;
'XYZ *n;
'CAMERA camera;
'{
'    double d;
'
'   if (camera.projection == PERSPECTIVE) {
'    d = camera.zoom / e.y;
'      n->x = d * e.x / tanthetah;
'      n->y = e.y;;
'      n->z = d * e.z / tanthetav;
'   } else {
'      n->x = camera.zoom * e.x / tanthetah;
'      n->y = e.y;
'      n->z = camera.zoom * e.z / tanthetav;
'   }
'}

'/*
'   Take a vector in eye coordinates and transform it into
'   normalised coordinates for a perspective view. No normalisation
'   is performed for an orthographic projection. Note that although
'   the y component of the normalised vector is copied from the eye
'   coordinate system, it is generally no longer needed. It can
'   however still be used externally for vector sorting.
'*/
Public Function Eye2Norm(E As tVector) As tVector
    Dim D As Single
    
    If camera.Projection = PERSPECTIVE Then
        D = camera.Zoom / E.Y
        Eye2Norm.X = D * E.X / tanthetaH
        Eye2Norm.Y = E.Y
        Eye2Norm.Z = D * E.Z / tanthetaV
    Else
        Eye2Norm.X = camera.Zoom * E.X / tanthetaH
        Eye2Norm.Y = E.Y
        Eye2Norm.Z = camera.Zoom * E.Z / tanthetaV
    End If
    
    
End Function



'/*
'   Clip a line segment to the normalised coordinate +- square.
'   The y component is not touched.
'*/
'int Trans_ClipNorm(n1,n2)
'XYZ *n1,*n2;
'{
'   double mu;
'
'   /* Is the line segment totally right of x = 1 ? */
'   if (n1->x >= 1 && n2->x >= 1)
'      return(FALSE);
'
'   /* Is the line segment totally left of x = -1 ? */
'   if (n1->x <= -1 && n2->x <= -1)
'      return(FALSE);
'
'   /* Does the vector cross x = 1 ? */
'   if ((n1->x > 1 && n2->x < 1) || (n1->x < 1 && n2->x > 1)) {
'      mu = (1 - n1->x) / (n2->x - n1->x);
'      if (n1->x < 1) {
'         n2->z = n1->z + mu * (n2->z - n1->z);
'         n2->x = 1;
'      } else {
'         n1->z = n1->z + mu * (n2->z - n1->z);
'         n1->x = 1;
'      }
'   }
'
'   /* Does the vector cross x = -1 ? */
'   if ((n1->x < -1 && n2->x > -1) || (n1->x > -1 && n2->x < -1)) {
'      mu = (-1 - n1->x) / (n2->x - n1->x);
'      if (n1->x > -1) {
'         n2->z = n1->z + mu * (n2->z - n1->z);
'         n2->x = -1;
'      } else {
'         n1->z = n1->z + mu * (n2->z - n1->z);
'         n1->x = -1;
'      }
'   }
'
'   /* Is the line segment totally above z = 1 ? */
'   if (n1->z >= 1 &&; n2->z >= 1)
'      return(FALSE);
'
'   /* Is the line segment totally below z = -1 ? */
'   if (n1->z <= -1 && n2->z <= -1)
'      return(FALSE);
'
'   /* Does the vector cross z = 1 ? */
'   if ((n1->z > 1 && n2->z < 1) || (n1->z < 1 && n2->z > 1)) {
'      mu = (1 - n1->z) / (n2->z - n1->z);
'      if (n1->z < 1) {
'         n2->x = n1->x + mu * (n2->x - n1->x);
'         n2->z = 1;
'      } else {
'         n1->x = n1->x + mu * (n2->x - n1->x);
'         n1->z = 1;
'      }
'   }
'
'   /* Does the vector cross z = -1 ? */
'   if ((n1->z < -1 && n2->z > -1) || (n1->z > -1 && n2->z < -1)) {
'      mu = (-1 - n1->z) / (n2->z - n1->z);
'      if (n1->z > -1) {
'         n2->x = n1->x + mu * (n2->x - n1->x);
'         n2->z = -1;
'      } else {
'         n1->x = n1->x + mu * (n2->x - n1->x);
'         n1->z = -1;
'      }
'   }
'
'   return(TRUE);
'}
'


'/*
'   Clip a line segment to the normalised coordinate +- square.
'   The y component is not touched.
'*/
Public Function ClipNorm(ByRef n1 As tVector, n2 As tVector) As Boolean
    Dim Mu As Double
    
    '   /* Is the line segment totally right of x = 1 ? */
    If (n1.X >= 1 And n2.X >= 1) Then ClipNorm = False: Exit Function
    
    '   /* Is the line segment totally left of x = -1 ? */
    If (n1.X <= -1 And n2.X <= -1) Then ClipNorm = False: Exit Function
    
    '   /* Does the vector cross x = 1 ? */
    If ((n1.X > 1 And n2.X < 1) Or (n1.X < 1 And n2.X > 1)) Then
        Mu = (1 - n1.X) / (n2.X - n1.X)
        If (n1.X < 1) Then
            n2.Z = n1.Z + Mu * (n2.Z - n1.Z)
            n2.X = 1
        Else
            n1.Z = n1.Z + Mu * (n2.Z - n1.Z)
            n1.X = 1
        End If
    End If
    
    '   /* Does the vector cross x = -1 ? */
    If ((n1.X < -1 And n2.X > -1) Or (n1.X > -1 And n2.X < -1)) Then
        Mu = (-1 - n1.X) / (n2.X - n1.X)
        If (n1.X > -1) Then
            n2.Z = n1.Z + Mu * (n2.Z - n1.Z)
            n2.X = -1
        Else
            n1.Z = n1.Z + Mu * (n2.Z - n1.Z)
            n1.X = -1
        End If
    End If
    
    '   /* Is the line segment totally above z = 1 ? */
    If (n1.Z >= 1 And n2.Z >= 1) Then ClipNorm = False: Exit Function
    
    '   /* Is the line segment totally below z = -1 ? */
    If (n1.Z <= -1 And n2.Z <= -1) Then ClipNorm = False: Exit Function
    
    '   /* Does the vector cross z = 1 ? */
    If ((n1.Z > 1 And n2.Z < 1) Or (n1.Z < 1 And n2.Z > 1)) Then
        Mu = (1 - n1.Z) / (n2.Z - n1.Z)
        If (n1.Z < 1) Then
            n2.X = n1.X + Mu * (n2.X - n1.X)
            n2.Z = 1
        Else
            n1.X = n1.X + Mu * (n2.X - n1.X)
            n1.Z = 1
        End If
    End If
    
    '   /* Does the vector cross z = -1 ? */
    If ((n1.Z < -1 And n2.Z > -1) Or (n1.Z > -1 And n2.Z < -1)) Then
        Mu = (-1 - n1.Z) / (n2.Z - n1.Z)
        If (n1.Z > -1) Then
            n2.X = n1.X + Mu * (n2.X - n1.X)
            n2.Z = -1
        Else
            n1.X = n1.X + Mu * (n2.X - n1.X)
            n1.Z = -1
        End If
    End If
    
    
    ClipNorm = True
    
End Function



'/*
'   Take a vector in normalised Coordinates and transform it into
'   screen coordinates.
'*/
'void Trans_Norm2Screen(norm, projected, Screen)
'XYZ norm;
'Point *projected;
'SCREEN screen;
'{
'   projected->h = screen.center.h - screen.size.h * norm.x / 2;
'   projected->v = screen.center.v - screen.size.v * norm.z / 2;
'}

'/*
'   Take a vector in normalised Coordinates and transform it into
'   screen coordinates.
'*/
Public Function Norm2Screen(norm As tVector) As POINTAPI
    
    Norm2Screen.X = Scree.Center.X + Scree.Size.X * norm.X / 2
    Norm2Screen.Y = Scree.Center.Y - Scree.Size.Y * norm.Z / 2
    
End Function


'/*
'   Transform a point from world to screen coordinates. Return TRUE
'   if the point is visible, the point in screen coordinates is p.
'   Assumes Trans_Initialise() has been called
'*/
'int Trans_Point(w,p,screen,camera)
'XYZ w;
'Point *p;
'SCREEN screen;
'CAMERA camera;
'{
'   XYZ e,n;
'
'   Trans_World2Eye(w,&e,camera);
'   if (e.y >= camera.NearPlane && e.y <= camera.FarPlane) {
'      Trans_Eye2Norm(e,&n,camera);
'      if (n.x >= -1 && n.x <= 1 && n.z >= -1 && n.z <= 1) {
'         Trans_Norm2Screen(n,p,screen);
'         return(TRUE);
'      }
'   }
'   return(FALSE);
'}

Public Function PointToScreen(W As tVector) As POINTAPI
    
    
    Dim E As tVector
    Dim N As tVector
    
    E = World2EYE(W)
    
    If (E.Y >= camera.NearPlane And E.Y <= camera.FarPlane) Then
        N = Eye2Norm(E)
        
        'If (N.X >= -1 And N.X <= 1 And N.Z >= -1 And N.Z <= 1) Then
        PointToScreen = Norm2Screen(N)
        
        'Else
        '    PointToScreen.X = -99999
        '    PointToScreen.Y = -99999
        'End If
        
    Else
        PointToScreen.X = -99999
    End If
    
End Function


'/*
'   Transform and appropriately clip a line segment from
'   world to screen coordinates. Return TRUE if something
'   is visible and needs to be drawn, namely a line between
'   screen coordinates p1 and p2.
'   Assumes Trans_Initialise() has been called
'*/
'int Trans_Line(w1,w2,p1,p2,screen,camera)
'XYZ w1,w2;
'Point *p1,*p2;
'SCREEN screen;
'CAMERA camera;
'{
'   XYZ e1,e2,n1,n2;
'
'   Trans_World2Eye(w1,&e1,camera);
'   Trans_World2Eye(w2,&e2,camera);
'   if (Trans_ClipEye(&e1,&e2,camera)) {
'      Trans_Eye2Norm(e1,&n1,camera);
'      Trans_Eye2Norm(e2,&n2,camera);
'      if (Trans_ClipNorm(&n1,&n2)) {
'         Trans_Norm2Screen(n1,p1,screen);
'         Trans_Norm2Screen(n2,p2,screen);
'         return(TRUE);
'      }
'   }
'   return(FALSE);
'}
'/*
'   Transform and appropriately clip a line segment from
'   world to screen coordinates. Return TRUE if something
'   is visible and needs to be drawn, namely a line between
'   screen coordinates p1 and p2.
'   Assumes Trans_Initialise() has been called
'*/

Public Function LineToScreen(w1 As tVector, w2 As tVector, RetP1 As POINTAPI, RetP2 As POINTAPI) As Boolean
    
    Dim E1 As tVector
    Dim E2 As tVector
    Dim n1 As tVector
    Dim n2 As tVector
    
    E1 = World2EYE(w1)
    E2 = World2EYE(w2)
    If ClipEYE(E1, E2) Then
        n1 = Eye2Norm(E1)
        n2 = Eye2Norm(E2)
        If ClipNorm(n1, n2) Then
            RetP1 = Norm2Screen(n1)
            RetP2 = Norm2Screen(n2)
            LineToScreen = True
        Else
            LineToScreen = False
        End If
    Else
        LineToScreen = False
    End If
End Function
