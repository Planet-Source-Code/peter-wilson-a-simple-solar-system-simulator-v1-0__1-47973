VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Solar System"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   435
      Left            =   3900
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton btnSolid 
      Caption         =   "Solid"
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton btnToggle 
      Caption         =   "Cls"
      Height          =   435
      Left            =   1980
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton btnZoomOut 
      Caption         =   "Zoom Out"
      Height          =   435
      Left            =   1020
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton btnZoomIn 
      Caption         =   "Zoom In"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ======================================================
' Simple Solar System Simulator
' Version: 1.0
'
' by Peter Wilson
' Copyright © 2003 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' ======================================================

Private Const g_sngPI As Single = 3.141594!
Private Const g_sngPIDivideBy180 As Single = 0.0174533!

Private blnSolidPlanets As Boolean
Private blnClearScreen As Boolean
Private sngGlobalZoom As Single


Private Sub DrawCrossHairs()

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Me.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    Me.ForeColor = RGB(0, 64, 64)
    Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
    
    ' Draw Vertical line
    Me.ForeColor = RGB(0, 92, 92)
    Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
    
    
    ' ==================
    ' Draw grid of dots.
    ' ==================
    Me.ForeColor = RGB(0, 220, 220)
    Dim sngX As Single, sngY As Single
    For sngX = 0 To Me.ScaleWidth Step 20
        For sngY = 0 To Me.ScaleHeight Step 20
        
            Me.PSet (sngX, sngY)    ' Draw the first quadrant...
            
            Me.PSet (-sngX, sngY)   ' ...then draw the others.
            Me.PSet (-sngX, -sngY)
            Me.PSet (sngX, -sngY)
            
        Next sngY
    Next sngX
End Sub


Private Sub DrawSolarSystem(lngRotationAngle As Long)

    Const EarthDistance = 12
    Const EarthMoonDistance = 4 ' ie. distance from Earth (not the Sun)
    
    Const MarsDistance = 25
    Const PhobosDistance = 3 ' ie. Mars's moon
    
    
    Dim sngEarthX As Single, sngMoonX As Single, sngMarsX As Single, sngPhobosX As Single
    Dim sngEarthY As Single, sngMoonY As Single, sngMarsY As Single, sngPhobosY As Single
    Dim sngRadians As Single
    
    
    If blnSolidPlanets = True Then
        Me.DrawStyle = vbSolid
        Me.FillStyle = vbFSSolid
    Else
        Me.DrawStyle = vbSolid
        Me.FillStyle = vbFSTransparent
    End If
    
    
    ' ===========================================================
    ' Draw our Sun the in the middle of the screen (never moves).
    ' ===========================================================
    Me.ForeColor = RGB(255, 255, 0)
    Me.FillColor = RGB(255, 255, 0)
    Me.Circle (0, 0), 3
    
    
    ' ==================
    ' Draw planet Earth.
    ' ==================
    sngRadians = ConvertDeg2Rad((lngRotationAngle) Mod 360)
    sngEarthX = (Sin(sngRadians) * EarthDistance)
    sngEarthY = (Cos(sngRadians) * EarthDistance)
    Me.ForeColor = RGB(0, 255, 220)
    Me.FillColor = RGB(0, 255, 220) ' Blue/Green
    Me.Circle (sngEarthX, sngEarthY), 1
    
    ' ==================
    ' Draw Earth's Moon.
    ' ==================
    sngRadians = ConvertDeg2Rad(lngRotationAngle * 2 Mod 360) ' ie. Moon rotates twice as fast?
    sngMoonX = (Sin(sngRadians) * EarthMoonDistance) + sngEarthX
    sngMoonY = (Cos(sngRadians) * EarthMoonDistance) + sngEarthY
    Me.ForeColor = RGB(127, 127, 127)
    Me.FillColor = RGB(127, 127, 127) ' Grey
    Me.Circle (sngMoonX, sngMoonY), 0.4
    
    
    ' ==========
    ' Draw Mars.
    ' ==========
    sngRadians = ConvertDeg2Rad((lngRotationAngle * 1.1) Mod 360)
    sngMarsX = (Sin(sngRadians) * MarsDistance)
    sngMarsY = (Cos(sngRadians) * MarsDistance)
    Me.ForeColor = RGB(255, 0, 0)
    Me.FillColor = RGB(255, 0, 0) ' The RED Planet!
    Me.Circle (sngMarsX, sngMarsY), 1
    
    
    ' ================
    ' Draw Mars' Moon.
    ' ================
    sngRadians = ConvertDeg2Rad((lngRotationAngle * 8) Mod 360)
    sngPhobosX = (Sin(sngRadians) * PhobosDistance) + sngMarsX
    sngPhobosY = (Cos(sngRadians) * PhobosDistance) + sngMarsY
    Me.ForeColor = RGB(255, 127, 0)
    Me.FillColor = RGB(255, 127, 0) ' The RED Planet!
    Me.Circle (sngPhobosX, sngPhobosY), 0.4
    
End Sub


Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub btnSolid_Click()
    blnSolidPlanets = Not blnSolidPlanets
End Sub

Private Sub btnToggle_Click()
    blnClearScreen = Not blnClearScreen
    If blnClearScreen = True Then
        Me.btnToggle.Caption = "Not Cls"
    Else
        Me.btnToggle.Caption = "Cls"
    End If
End Sub


Private Sub btnZoomIn_Click()
    sngGlobalZoom = sngGlobalZoom / 1.3
    Call Form_Resize
End Sub


Private Sub btnZoomOut_Click()
    sngGlobalZoom = sngGlobalZoom * 1.3
    Call Form_Resize
End Sub


Private Sub Form_Load()
    blnClearScreen = True
    sngGlobalZoom = 100
End Sub


Private Sub Form_Resize()

    On Error Resume Next
    
    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier!
    Dim sngAspectRatio As Single
    sngAspectRatio = Me.Width / Me.Height
    
    Me.ScaleLeft = -sngGlobalZoom / 2
    Me.ScaleWidth = sngGlobalZoom
    
    Me.ScaleHeight = sngGlobalZoom / sngAspectRatio
    Me.ScaleTop = -Me.ScaleHeight / 2

End Sub


Private Sub Timer1_Timer()

    ' Increment a counter.
    Static lngRotationAngle As Long
    lngRotationAngle = lngRotationAngle + 1
    If lngRotationAngle = 2 ^ 31 Then lngRotationAngle = 0
    
    
    ' =============
    ' Clear screen.
    ' =============
    If blnClearScreen = True Then Me.Cls
    
    
    ' ===========================
    ' Draw Crosshairs (optional).
    ' ===========================
    Call DrawCrossHairs
    
    
    ' Draw new planets and calculate positions.
    Call DrawSolarSystem(lngRotationAngle)
    
End Sub


Private Function ConvertDeg2Rad(Degress As Single) As Single
    
    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function


