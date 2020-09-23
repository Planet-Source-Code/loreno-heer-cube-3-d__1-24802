Attribute VB_Name = "Draw3D"
Const pi180 = 0.0174532925199
Dim frm As Form
Dim ww As Integer
Dim Ixy_angle, Iz_angle, dYYshift, dXXshift As Integer
Public csx, csy As Integer 'Return
Dim cosa, cosb, sina, sinb, coscosba, cossinba, sincosba, sinsinba, zoom As Double
Public Function SetForm(fr As Form)
Set frm = fr
End Function
Public Function SetView( _
ByVal xy_angle As Double, _
ByVal z_angle As Double, _
ByVal YYshift As Double, _
ByVal XXshift As Double, _
ByVal zm As Double)
Ixy_angle = xy_angle
Iz_angle = z_angle
cosa = Cos(Ixy_angle * pi180)
sina = Sin(Ixy_angle * pi180)
cosb = Cos(Iz_angle * pi180)
sinb = Sin(Iz_angle * pi180)
sinsinba = sinb * sina
sincosba = sinb * cosa
cossinba = sina * cosb
coscosba = cosb * cosa
dYYshift = YYshift
dXXshift = XXshift
zoom = zm
End Function
Private Sub posxy(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double)
    Dim Yy, Xx As Double
    cosa = Cos(Ixy_angle * pi180)
    sina = Sin(Ixy_angle * pi180)
    cosb = Cos(Iz_angle * pi180)
    sinb = Sin(Iz_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa
    Yy = zoom / (10# - (z1 * cosb + y1 * sinsinba - x1 * sincosba))
    Xx = 100# * (1# + (y1 * cosa + x1 * sina) * Yy)
    csx = Int(dXXshift) + Int(Xx)
    Xx = 100# * (1# + (y1 * cossinba - x1 * coscosba - z1 * sinb) * Yy)
    csy = Int(dYYshift) + Int(Xx)
End Sub
Public Function PSet3D(ByVal x As Double, ByVal y As Double, ByVal z As Double, Optional col As ColorConstants)
    posxy x, y, z
    frm.PSet (csx, csy), col
End Function
Public Function Line3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, Optional col As ColorConstants)
    posxy x1, y1, z1
    nx1 = csx: ny1 = csy
    posxy x2, y2, z2
    frm.Line (nx1, ny1)-(csx, csy), col
End Function
Public Function Rect3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, stp As Double, Optional col As ColorConstants)
Dim n1, n2, n3
For n1 = x1 To x2 Step stp
    For n2 = y1 To y2 Step stp
        For n3 = z1 To z2 Step stp
            PSet3D n1, n2, n3, col
        Next
    Next
Next
End Function
Public Function Text2D(ByVal x As Double, ByVal y As Double, ByVal z As Double, text As String, Optional col As ColorConstants)
    posxy x, y, z
    frm.CurrentX = csx
    frm.CurrentY = csy
    frm.Print text
End Function
