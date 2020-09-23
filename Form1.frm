VERSION 5.00
Begin VB.Form frmCube 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "3D Cube"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "frmCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ixy_angle, Iz_angle, dYYshift, dXXshift As Integer
Dim speed
Dim zoom As Double
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
    Ixy_angle = Ixy_angle + (speed * 2)
    Cube
ElseIf KeyCode = vbKeyRight Then
    Ixy_angle = Ixy_angle - (speed * 2)
    Cube
ElseIf KeyCode = vbKeyUp Then
    Iz_angle = Iz_angle + (speed * 2)
    Cube
ElseIf KeyCode = vbKeyDown Then
    Iz_angle = Iz_angle - (speed * 2)
    Cube
ElseIf KeyCode = vbKeyHome Then
    zoom = zoom + (speed / 10)
    Cube
ElseIf KeyCode = vbKeyEnd Then
    zoom = zoom - (speed / 10)
    Cube
ElseIf KeyCode = vbKeyW Then
    dYYshift = dYYshift - (speed)
    Cube
ElseIf KeyCode = vbKeyY Then
    dYYshift = dYYshift + (speed)
    Cube
ElseIf KeyCode = vbKeyA Then
    dXXshift = dXXshift - (speed)
    Cube
ElseIf KeyCode = vbKeyS Then
    dXXshift = dXXshift + (speed)
    Cube
ElseIf KeyCode = vbKeyAdd Then
    speed = speed + 0.1
    Cube
ElseIf KeyCode = vbKeySubtract Then
    speed = speed - 0.1
    Cube
End If
End Sub

Private Sub Form_Load()
Ixy_angle = 270
Iz_angle = 90
dYYshift = 80
dXXshift = 80
zoom = 6#
speed = 1
SetForm Me
Cube
End Sub
Private Sub Cube()
    Cls
    SetView Ixy_angle, Iz_angle, dYYshift, dXXshift, zoom
    Line3D 1, 1, 1, -1, 1, 1, vbBlue
    Line3D 1, 1, 1, 1, -1, 1, vbBlue
    Line3D 1, 1, 1, 1, 1, -1, vbBlue
    
    Line3D -1, 1, 1, -1, -1, 1, vbBlue
    Line3D -1, 1, 1, -1, 1, -1, vbBlue
    
    Line3D 1, -1, 1, -1, -1, 1, vbBlue
    Line3D 1, -1, 1, 1, -1, -1, vbBlue
    
    Line3D 1, 1, -1, -1, 1, -1, vbBlue
    Line3D 1, 1, -1, 1, -1, -1, vbBlue
    
    Line3D -1, -1, -1, 1, -1, -1, vbBlue
    Line3D -1, -1, -1, -1, -1, 1, vbBlue
    Line3D -1, -1, -1, -1, 1, -1, vbBlue
    
    Text2D -1, -1, -1, "(-1|-1|-1)"
    Text2D -1, -1, 1, "(-1|-1|1)"
    Text2D -1, 1, -1, "(-1|1|-1)"
    Text2D -1, 1, 1, "(-1|1|1)"
    Text2D 1, -1, -1, "(1|-1|-1)"
    Text2D 1, -1, 1, "(1|-1|1)"
    Text2D 1, 1, -1, "(1|1|-1)"
    Text2D 1, 1, 1, "(1|1|1)"
End Sub
