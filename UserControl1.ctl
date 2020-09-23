VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ww As Integer
Dim Ixy_angle, Iz_angle, dYYshift, dXXshift, csx, csy As Integer
Dim cosa, cosb, sina, sinb, coscosba, cossinba, sincosba, sinsinba, zoom, pi180 As Double
Private Sub UserControl_Initialize()
pi180 = 0.0174532925199 ' Pi / 180 ...
Public Ixy_angle
Ixy_angle = 280
Iz_angle = 90
cosa = Cos(Ixy_angle * pi180)
sina = Sin(Ixy_angle * pi180)
cosb = Cos(Iz_angle * pi180)
sinb = Sin(Iz_angle * pi180)
sinsinba = sinb * sina
sincosba = sinb * cosa
cossinba = sina * cosb
coscosba = cosb * cosa
dYYshift = 80
dXXshift = 80
zoom = 6#

  posxy -1, -1, -1: xxx = csx: yyy = csy:
    posxy -1, 1, -1: Line (xxx, yyy)-(csx, csy), QBColor(15): x = csx: y = csy
    posxy -1, 1, 1: Line (x, y)-(csx, csy), QBColor(15): x = csx: y = csy
    posxy -1, -1, 1: Line (x, y)-(csx, csy), QBColor(15): Line (csx, csy)-(xxx, yyy), QBColor(15)
    posxy 1, -1, -1: xxx = csx: yyy = csy:
    posxy 1, 1, -1: Line (xxx, yyy)-(csx, csy), QBColor(15): x = csx: y = csy
    posxy 1, 1, 1: Line (x, y)-(csx, csy), QBColor(15): x = csx: y = csy
    posxy 1, -1, 1: Line (x, y)-(csx, csy), QBColor(15): Line (csx, csy)-(xxx, yyy), QBColor(15)
    
    posxy 1, -1, -1: x = csx: y = csy: posxy -1, -1, -1: Line (x, y)-(csx, csy), QBColor(15)
    posxy 1, -1, 1: x = csx: y = csy: posxy -1, -1, 1: Line (x, y)-(csx, csy), QBColor(15)
    posxy 1, 1, 1: x = csx: y = csy: posxy -1, 1, 1: Line (x, y)-(csx, csy), QBColor(15)
    posxy 1, 1, -1: x = csx: y = csy: posxy -1, 1, -1: Line (x, y)-(csx, csy), QBColor(15)

End Sub
Private Sub posxy(x1 As Double, y1 As Double, z1 As Double)
    Dim Yy, Xx As Double
    Yy = zoom / (10# - (z1 * cosb + y1 * sinsinba - x1 * sincosba))
    Xx = 100# * (1# + (y1 * cosa + x1 * sina) * Yy)
    csx = Int(dXXshift) + Int(Xx)
    Xx = 100# * (1# + (y1 * cossinba - x1 * coscosba - z1 * sinb) * Yy)
    csy = Int(dYYshift) + Int(Xx)
End Sub



