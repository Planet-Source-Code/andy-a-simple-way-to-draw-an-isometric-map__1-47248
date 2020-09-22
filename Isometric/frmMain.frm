VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isometric Map"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":000C
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   720
      MousePointer    =   1  'Arrow
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   720
      Width           =   1440
      Begin VB.Shape shp2 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'IsoEngine objects
Dim IsoEngine As IsoEngine, Cam As IECamera

'Tileset variables
Private Type TexInfo
    Names() As String
    Colors() As Long
    Types() As Byte
    NumTex As Integer
End Type
Dim TexInfo As TexInfo

'Map variables
Dim Map() As Integer, MapW As Integer, MapH As Integer

'Mouse variables
Dim mX As Integer, mY As Integer
Dim MDown As Boolean
Dim NoScroll As Boolean

Const Pi As Double = 3.14159265358979

Private Sub Form_Load()
    'Show the form
    Show
    MiniMap.Move 16, ScaleHeight - MiniMap.Height - 16
    shp.Move 0, 0, 0, 0
    shp2.Move 0, 0, 0, 0
    DoEvents
    
    'Initialize IsoEngine and set options
    Set IsoEngine = New IsoEngine
    With IsoEngine
        .Initialize hWnd, ScaleWidth, ScaleHeight, , , 1.5
        .SetAppPath App.Path & "\"
        .LoadTextureInfo "TexInfo.tex", TexInfo.Names, TexInfo.Colors, TexInfo.Types, TexInfo.NumTex
        .SetTileSize 128, 64
        
        'Load the tile textures, object textures, and other textures
        .LoadTileTextures "Graphics\TT", ".bmp"
        .LoadObjectTextures "Graphics\O"
        .LoadTexture "Graphics\Arrow.bmp", "Arrow", Magenta
        .LoadTexture "Graphics\Cursor.bmp", "Cursor", Magenta
        
        'Load the map
        .LoadMapFromFile "Map.map", Map, MapW, MapH
        
        'Create the fonts
        .CreateFont "Main", "MS Sans Serif", 12, True, False, False
        .CreateFont "Web", "MS Sans Serif", 24, True, False, False
        
        'Draw the minimap
        .DrawMiniMap MiniMap.hDC, 1.5
        MiniMap.Refresh
        
        'Initialize the camera
        Set Cam = New IECamera
        Cam.Move 0, 0
        
        'Our render loop
        .FadeIn 1000, 0, 0, 0
        Do
            DoEvents
            .Clear
            
            'Draw the map
            .DrawMap
            .DrawMapObjects
            
            'Draw the scroll arrow
            .Sprite_Begin
                If mX <= 32 Then
                    .DrawSprite GetTex("Arrow"), Pnt(0, mY - 16), Pnt(1, 1), Pi * 0.5, Green + TransMedFlag
                    Cam.X = Cam.X - (700 / .GetFPS)
                ElseIf mX >= ScaleWidth - 32 Then
                    .DrawSprite GetTex("Arrow"), Pnt(ScaleWidth - 32, mY - 16), Pnt(1, 1), Pi * 1.5, Green + TransMedFlag
                    Cam.X = Cam.X + (700 / .GetFPS)
                End If
                If mY <= 32 Then
                    .DrawSprite GetTex("Arrow"), Pnt(mX - 16, 0), Pnt(1, 1), Pi * 0, Green + TransMedFlag
                    Cam.Y = Cam.Y - (350 / .GetFPS)
                ElseIf mY >= ScaleHeight - 32 Then
                    .DrawSprite GetTex("Arrow"), Pnt(mX - 16, ScaleHeight - 32), Pnt(1, 1), Pi * 1, Green + TransMedFlag
                    Cam.Y = Cam.Y + (350 / .GetFPS)
                End If
                If NoScroll Then
                    .DrawSprite GetTex("Cursor"), Pnt(mX, mY), Pnt(1, 1), 0, Blue + TransMedFlag
                End If
            .Sprite_End
            
            'Draw the FPS
            .DrawBox RECT(0, 0, ScaleWidth, 22), Orange + TransHvyFlag, Orange + TransparentFlag, Orange + TransHvyFlag, Orange + TransparentFlag
            .DrawText 2, 1, "FPS: " & .GetFPS, GetFont("Main"), White
            
            'Update the minimap
            Dim l As Integer, l2 As Integer, t As Integer, w As Integer, h As Integer
            Cam.GetMiniMapCameraRect l, t, w, h, l2
            shp.Move l, t, w, h
            shp2.Move l2, t, w, h
            
            .RenderToScreen
        Loop
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mX = X
    mY = Y
    NoScroll = (mX > 32 And mX < ScaleWidth - 32 _
                And mY > 32 And mY < ScaleHeight - 32)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MiniMap.Visible = False
    
    Dim tmr As Single, col As Long
    tmr = Timer
    With IsoEngine
        Do
            DoEvents
            .Clear
            
            If Timer - tmr <= 3.5 Then
                'Draw the website URL
                If (Timer - tmr) / 3.5 * 255 > 255 Then col = -1 Else col = RGBA(255, 255, 255, (Timer - tmr) / 3.5 * 255)
                .DrawTextBox "Visit us on the web at" & vbCrLf _
                        & "www.firstproductions.com/isoengine", GetFont("Web"), col, RECT(0, 0, ScaleWidth, ScaleHeight), IETextAlign_HCenter + IETextAlign_VCenter
            Else
                'Draw the Thank You text
                If (Timer - (tmr + 3.5)) / 3.5 * 255 > 255 Then col = -1 Else col = RGBA(255, 255, 255, (Timer - (tmr + 3.5)) / 3.5 * 255)
                .DrawTextBox "Thank you for your interest in IsoEngine.", GetFont("Web"), col, RECT(0, 0, ScaleWidth, ScaleHeight), IETextAlign_HCenter + IETextAlign_VCenter
            End If
            
            'Draw the cursor
            .Sprite_Begin
            .DrawSprite GetTex("Cursor"), Pnt(mX, mY), Pnt(1, 1), 0, Blue + TransMedFlag
            .Sprite_End
                        
            IsoEngine.RenderToScreen
        Loop Until Timer - tmr >= 7
    End With
    
    IsoEngine.Unload
    Set IsoEngine = Nothing
    Set Cam = Nothing
    End
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDown = True
    Cam.ClickedMiniMap X, Y
End Sub

Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MDown Then Cam.ClickedMiniMap X, Y
End Sub

Private Sub MiniMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDown = False
End Sub
