Attribute VB_Name = "modColor"
'Taken from EGL_Dxf by Erkan Sanli
'from PlanetSourceCode


Option Explicit


Public Type ColorRGB
    R As Integer
    G As Integer
    B As Integer
End Type

Public Function ColorSet(R As Integer, G As Integer, B As Integer) As ColorRGB
    
    ColorSet.R = R
    ColorSet.G = G
    ColorSet.B = B
    
End Function

Public Function ColorAdd(C1 As ColorRGB, C2 As ColorRGB) As ColorRGB
    
    ColorAdd.R = C1.R + C2.R
    ColorAdd.G = C1.G + C2.G
    ColorAdd.B = C1.B + C2.B
    
End Function

Public Function ColorSub(C1 As ColorRGB, C2 As ColorRGB) As ColorRGB
    
    ColorSub.R = C1.R - C2.R
    ColorSub.G = C1.G - C2.G
    ColorSub.B = C1.B - C2.B
    
End Function

Public Function ColorScale(C As ColorRGB, S As Single) As ColorRGB
    
    ColorScale.R = C.R * S
    ColorScale.G = C.G * S
    ColorScale.B = C.B * S
    
End Function

Public Function ColorPlus(C As ColorRGB, L As Integer) As ColorRGB
    
    ColorPlus.R = C.R + L
    ColorPlus.G = C.G + L
    ColorPlus.B = C.B + L
    
End Function

Public Function ColorInvert(C As ColorRGB) As ColorRGB
    
    ColorInvert.R = 255 - C.R
    ColorInvert.G = 255 - C.G
    ColorInvert.B = 255 - C.B
    
End Function

Public Function ColorGray(R As Integer, G As Integer, B As Integer) As ColorRGB
    
    Dim Temp As Integer
    
    Temp = (R + G + B) / 3
    ColorGray.R = Temp
    ColorGray.G = Temp
    ColorGray.B = Temp
    
End Function

Function ColorRandom() As ColorRGB
    
    Randomize
    ColorRandom.R = Rnd * 255
    ColorRandom.G = Rnd * 255
    ColorRandom.B = Rnd * 255
    
End Function

Public Function ColorLongToRGB(lC As Long) As ColorRGB
    
    ColorLongToRGB.R = (lC And &HFF&)
    ColorLongToRGB.G = (lC And &HFF00&) / &H100&
    ColorLongToRGB.B = (lC And &HFF0000) / &H10000
    
End Function

Public Function ColorRGBToLong(C As ColorRGB) As Long
    
    ColorRGBToLong = RGB(C.R, C.G, C.B)
    
End Function

'Public Function ColorBlend(C1 As ColorRGB, C2 As ColorRGB, ByVal Ratio As Integer) As ColorRGB
'    ColorBlend.R = C1.R + Ratio * (C2.R - C1.R)
'    ColorBlend.G = C1.G + Ratio * (C2.G - C1.G)
'    ColorBlend.B = C1.B + Ratio * (C2.B - C1.B)
'
'End Function
