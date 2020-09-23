VERSION 5.00
Begin VB.Form frmM 
   Caption         =   "Sticky Physics 3D"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chDrawMode 
      Caption         =   "Draw Physic Structure"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9480
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   431
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   567
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks to: Erkan Sanli (for NoDirectx 3D render) and  Paul Bourke (for Perspective 3D)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   9375
   End
End
Attribute VB_Name = "frmM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'reexre@gmail.com       Roberto Mior
'
'
'
Option Explicit

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & "   V" & App.Major & "." & App.Minor
    
    
    Gravity = 0.25
    Atmosphere = 0.001
    WallBounce = 0.9
    WallFriction = 0.2
    
    'WINDx = 0.08
    
    Ncre = 1
    ReDim CRE(Ncre)
    
    
    camera.cTo.X = 25
    camera.cTo.Y = 25
    camera.cTo.Z = 25
    
    camera.cFrom.X = 100
    camera.cFrom.Y = 200
    camera.cFrom.Z = -200
    
    camera.cUp.Y = 10
    
    camera.ANGh = 70
    camera.ANGv = 70
    
    camera.FarPlane = 2000
    camera.NearPlane = 1
    
    camera.Projection = PERSPECTIVE
    camera.Zoom = 1
    
    'camera.Projection = ORTHOGRAPHIC
    'camera.Zoom = 0.005
    
    Scree.Size.X = PIC.ScaleWidth
    Scree.Size.Y = PIC.ScaleHeight
    Scree.Center.X = PIC.ScaleWidth \ 2
    Scree.Center.Y = PIC.ScaleHeight \ 2
    
    Light.X = 1000
    Light.Y = 1000
    Light.Z = -1000
    
    Light = VectorNormalize(Light)

End Sub




Private Sub Command1_Click()
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim T As Single
    
    Command1.Caption = "ReStart"
    
    CRE(1).Kill
    
    '               4--------8
    '              /.       /|
    '             / .      / |
    '            3--------7  |
    '            |  2.....|..6
    '            | .      | /
    '            |.       |/
    '            1--------5
    '
    
    
    '  Y  Z
    '  ^ /
    '  |/
    '  --->X
    
    
    With CRE(1)
        
        
        For X = 0 To 50 Step 50
            For Y = 0 To 50 Step 50
                For Z = 0 To 50 Step 50
                    
                    .AddPoint X, 200 + Y, Z
                    
                Next
            Next
        Next
        
        T = 0.4 '0.24
        
        .AddLink 1, 2, T
        .AddLink 3, 4, T
        .AddLink 5, 6, T
        .AddLink 7, 8, T
        
        .AddLink 1, 3, T
        .AddLink 3, 7, T
        .AddLink 7, 5, T
        .AddLink 5, 1, T
        
        .AddLink 2, 4, T, 0, 80, 5
        .AddLink 4, 8, T
        .AddLink 8, 6, T, 0, 80, 5
        .AddLink 6, 2, T
        
        '------
        'Internal diagonals
        '    .AddLink 1, 8, T
        '    .AddLink 3, 6, T
        '    .AddLink 4, 5, T
        '    .AddLink 2, 7, T
        
        '-----------
        
        
        'Diagonals
        .AddLink 1, 7, T
        .AddLink 3, 5, T
        .AddLink 2, 8, T
        .AddLink 4, 6, T
        
        .AddLink 1, 4, T
        .AddLink 3, 2, T
        .AddLink 5, 8, T
        .AddLink 7, 6, T
        
        .AddLink 3, 8, T
        .AddLink 7, 4, T
        .AddLink 1, 6, T
        .AddLink 5, 2, T
        '
               
        
        '.........................
        '      .AddPoint 200, 300, 50
        '
        '      .AddLink 8, 9, t
        '--------------------------
        
        '(Front)Face Points must be Clockwise [BackFace is not drawn]
        
        .AddFace 7, 1, 3, 255, 0, 0
        .AddFace 1, 7, 5, 255, 0, 0
        .AddFace 8, 5, 7, 0, 255, 0
        .AddFace 5, 8, 6, 0, 255, 0
        .AddFace 4, 1, 2, 0, 0, 255
        .AddFace 1, 4, 3, 0, 0, 255
        .AddFace 6, 4, 2, 255, 255, 0
        .AddFace 4, 6, 8, 255, 255, 0
        .AddFace 1, 6, 2, 0, 255, 255
        .AddFace 1, 5, 6, 0, 255, 255
        .AddFace 8, 3, 4, 255, 0, 255
        .AddFace 3, 8, 7, 255, 0, 255
               
        
        'Pseudo Floor
        .AddPoint -150 + 50, 0, -150 + 50
        .AddPoint 150 + 50, 0, -150 + 50
        .AddPoint 150 + 50, 0, 150 + 50
        .AddPoint -150 + 50, 0, 150 + 50
        
        .AddFace .NP - 3, .NP - 1, .NP - 2, 100, 100, 100
        .AddFace .NP - 3, .NP, .NP - 1, 100, 100, 100
        
    End With
    
    Timer1.Enabled = True
    
End Sub


Private Sub Timer1_Timer()
    
    With CRE(1)
        
        .DoPhysics
        
        camera.cTo.X = CRE(1).MidPointX
        camera.cTo.Y = CRE(1).MidPointY
        camera.cTo.Z = CRE(1).MidPointZ
        
        .Convert3DToScreen
        
        
        BitBlt PIC.hDC, 0, 0, PIC.ScaleWidth, PIC.ScaleHeight, PIC.hDC, 0, 0, vbBlackness
        .DRAW IIf(chDrawMode.Value = Checked, 1, 0), PIC.hDC
        PIC.Refresh
        DoEvents
        
    End With
    
End Sub
