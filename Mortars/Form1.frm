VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   179
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'is the cannonball flying?
Dim bCannonBallFlying As Boolean

'is explosion?
Dim bExplosion As Boolean
Dim bBigExplosion As Boolean
Dim bBigExplodeOnce As Boolean

'is the cannon shooting?
Dim bShooting As Boolean

'the color-fill and Lock RECT
Dim rFill As RECT

'the time moment in the cannonball's trajectory
Dim T As Double

'the cannonball's position in every time moment
Dim CBallX As Long
Dim CBallY As Long

'The time im milliseconds per loop
Dim mSperLoop As Long

'is the screen centered to the cannonball?
Dim CenterAtCBall As Boolean

'the current player
Dim CurrentPlayer As Integer

'this will protect our cannon beiing destroyed by us
'while shooting at nearby trees
Dim bCanBeDestroyed As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If bCannonBallFlying = False Then
    'angle change
    If KeyCode = vbKeyZ Then
        If CurrentPlayer = 1 Then
            P1.Angle = P1.Angle + 1
            If P1.Angle > 135 Then P1.Angle = 135
        Else
            P2.Angle = P2.Angle + 1
            If P2.Angle > 135 Then P2.Angle = 135
        End If
    End If
    If KeyCode = vbKeyX Then
        If CurrentPlayer = 1 Then
            P1.Angle = P1.Angle - 1
            If P1.Angle < 45 Then P1.Angle = 45
        Else
            P2.Angle = P2.Angle - 1
            If P2.Angle < 45 Then P2.Angle = 45
        End If
    End If
    'velocity change
    If KeyCode = vbKeyA Then
        If CurrentPlayer = 1 Then
            P1.GunPowder = P1.GunPowder + 1
            If P1.GunPowder > 1000 Then P1.GunPowder = 1000
        Else
            P2.GunPowder = P2.GunPowder + 1
            If P2.GunPowder > 1000 Then P2.Angle = 1000
        End If
    End If
    If KeyCode = vbKeyS Then
        If CurrentPlayer = 1 Then
            P1.GunPowder = P1.GunPowder - 1
            If P1.GunPowder < 10 Then P1.GunPowder = 10
        Else
            P2.GunPowder = P2.GunPowder - 1
            If P2.GunPowder < 10 Then P2.GunPowder = 10
        End If
    End If
    
End If
'fire key
If KeyCode = vbKeySpace Then
    If bCannonBallFlying = False And _
        bExplosion = False Then
        
        If CurrentPlayer = 1 Then
            CBallX = P1.Pos.X + ShootFromX
            CBallY = P1.Pos.Y + ShootFromY
        Else
            CBallX = P2.Pos.X + ShootFromX
            CBallY = P2.Pos.Y + ShootFromY
        End If
        
        T = 0
        bCannonBallFlying = True
        bShooting = True
    End If
End If

If KeyCode = vbKey1 Then CurrentPlayer = 1
If KeyCode = vbKey2 Then CurrentPlayer = 2

End Sub
Private Sub Form_Load()
'----------------------
'setting some variables
'----------------------
bCannonBallFlying = False
bExplosion = False
bBigExplosion = False
bShooting = False
CurrentPlayer = 1
P1.Destroyed = False
P2.Destroyed = False
bCanBeDestroyed = False
bBigExplodeOnce = False

'--------------------
'seting the DX and DD
'--------------------
Set DX = New DirectX7
Set DD = DX.DirectDrawCreate("")

'----------------------------------------
'Setting the resolution and the coop mode
'----------------------------------------
DD.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
DD.SetDisplayMode ResWidth, ResHeight, 32, 0, DDSDM_DEFAULT

'---------------------
'setting the colorkeys
'---------------------
keyMagenta.high = vbMagenta '&HFF00FF
keyMagenta.low = vbMagenta '&HFF00FF
keyGreen.high = vbGreen '&HFF00
keyGreen.low = vbGreen '&HFF00

'--------------------
'Setting the surfaces
'--------------------
'the primary surface
sdPrim.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
sdPrim.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
sdPrim.lBackBufferCount = 1
Set ddsPrim = DD.CreateSurface(sdPrim)

'the backbuffer
Dim Caps As DDSCAPS2
Caps.lCaps = DDSCAPS_BACKBUFFER
Set ddsBack = ddsPrim.GetAttachedSurface(Caps)
ddsBack.GetSurfaceDesc sdBack
ddsBack.SetForeColor vbBlack

Set Clip = DD.CreateClipper(0)
Clip.SetClipList 1, rClip
Clip.SetHWnd Me.hWnd
ddsBack.SetClipper Clip

'the terrain
sdTerrain.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdTerrain.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdTerrain.lWidth = MapWidth * TerrainTileWidth
sdTerrain.lHeight = MapHeight * TerrainTileHeight
Set ddsTerrain = DD.CreateSurface(sdTerrain)
rTerrain.Top = 0
rTerrain.Left = 0
rTerrain.Right = ResWidth
rTerrain.Bottom = ResHeight
ddsTerrain.SetColorKey DDCKEY_SRCBLT, keyGreen

'trees
sdTree.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdTree.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdTree.lWidth = TreeWidth
sdTree.lHeight = TreeHeight
Set ddsTree = DD.CreateSurface(sdTree)
rTree.Left = 0
rTree.Right = TreeWidth
rTree.Top = 0
rTree.Bottom = TreeHeight
ddsTree.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the terrain tile
sdTerrainTile.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdTerrainTile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdTerrainTile.lWidth = TerrainTileWidth
sdTerrainTile.lHeight = TerrainTileHeight
rTerrainTile.Top = 0
rTerrainTile.Left = 0
rTerrainTile.Right = sdTerrainTile.lWidth
rTerrainTile.Bottom = sdTerrainTile.lHeight

'the cannon
sdCannon.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdCannon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdCannon.lWidth = CannonTileSetWidth
sdCannon.lHeight = CannonTileSetHeight
Set ddsCannon = DD.CreateSurfaceFromFile(sApPath & "data\mortar.bmp", sdCannon)
'rCannon.Top = 0
'rCannon.Left = 0
'rCannon.Right = CannonTileWidth
'rCannon.Bottom = CannonTileHeight
ddsCannon.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the cannonball
sdCannonBall.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdCannonBall.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdCannonBall.lWidth = CannonBallWidth
sdCannonBall.lHeight = CannonBallHeight
Set ddsCannonBall = DD.CreateSurfaceFromFile(sApPath & "data\cannonball.bmp", sdCannonBall)
rCannonBall.Top = 0
rCannonBall.Left = 0
rCannonBall.Right = sdCannonBall.lWidth
rCannonBall.Bottom = sdCannonBall.lHeight
ddsCannonBall.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the blast hole
sdHole.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdHole.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdHole.lWidth = HoleWidth
sdHole.lHeight = HoleHeight
Set ddsHole = DD.CreateSurfaceFromFile(sApPath & "data\hole.bmp", sdHole)
rHole.Top = 0
rHole.Left = 0
rHole.Right = sdHole.lWidth
rHole.Bottom = sdHole.lHeight
ddsHole.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the explosion
sdExp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdExp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdExp.lWidth = ExpTileWidth
sdExp.lHeight = ExpTileHeight
Set ddsExp = DD.CreateSurfaceFromFile(sApPath & "data\explosion.bmp", sdExp)
'rExp.Top = 0
'rExp.Left = 0
'rExp.Right = ExpWidth
'rExp.Bottom = ExpHeight
ddsExp.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the arrow
sdArrow.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdArrow.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdArrow.lWidth = ArrowTileWidth
sdArrow.lHeight = ArrowTileHeight
Set ddsArrow = DD.CreateSurfaceFromFile(sApPath & "data\arrow.bmp", sdArrow)
'rArrow.Top = 0
'rArrow.Left = 0
'rArrow.Right = ArrowWidth
'rArrow.Bottom = ArrowHeight
ddsArrow.SetColorKey DDCKEY_SRCBLT, keyMagenta

'the sky
sdSky.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdSky.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdSky.lWidth = SkyWidth
sdSky.lHeight = SkyHeight
Set ddsSky = DD.CreateSurfaceFromFile(sApPath & "data\sky.bmp", sdSky)
'rSky.Top = 0
'rSky.Left = 0
'rSky.Right = ResWidth
'rSky.Bottom = ResHeight

'-----------------------
'Generate random terrain
'-----------------------
Dim I As Long
Dim I2 As Long

'first generate some random elevations
Randomize
For I = LBound(Elev) To UBound(Elev) Step 1
    Elev(I) = Int(Rnd * 7)
    If Elev(I) = 0 Then Elev(I) = 1
Next

'calculate the grid
Dim Pos(1 To MapWidth, 1 To MapHeight) As COORDS

For I = 1 To MapWidth Step 1
    For I2 = 1 To MapHeight Step 1
        Pos(I, I2).X = (I - 1) * TerrainTileWidth
        Pos(I, I2).Y = (MapHeight * TerrainTileHeight) - (I2 * TerrainTileHeight)
        Pos(I, I2).Tile = 999
    Next
Next

'generate the terrain
I2 = 1
For I = LBound(Elev) To UBound(Elev) Step 1
    'place the first part of the elevation on the map
    'check to see if the previous elev was higher or lower
    If I = 1 Then
        Pos((I * 2) - 1, Elev(I)).Tile = 0
    Else
        If Elev(I - 1) < Elev(I) Then
            'we're higher
            Pos((I * 2) - 1, Elev(I)).Tile = 1
        
        ElseIf Elev(I - 1) > Elev(I) Then
            'we're lower
            Pos((I * 2) - 1, Elev(I) + 1).Tile = 3
        
        ElseIf Elev(I - 1) = Elev(I) Then
            'we're going the same level
            Pos((I * 2) - 1, Elev(I)).Tile = 0
            
        End If
    End If
    
    'check to see if we're going up or down, or nowhere, and
    'place the second part of the elevation
    If I = UBound(Elev) Then
        Pos((I * 2), Elev(I)).Tile = 0
    Else
        If Elev(I) < Elev(I + 1) Then
            'we're lower
            Pos((I * 2), Elev(I) + 1).Tile = 4
            
        ElseIf Elev(I) > Elev(I + 1) Then
            'we're higher
            Pos((I * 2), Elev(I)).Tile = 2
        
        ElseIf Elev(I) = Elev(I + 1) Then
            'we're the same
            Pos((I * 2), Elev(I)).Tile = 0
            
        End If
    End If
Next

'fill the background with the transparent color key
ddsTerrain.BltColorFill rFill, keyGreen.high
Dim rDestTemp As RECT

'BLt the tiles
For I = 1 To MapWidth Step 1
    For I2 = 1 To MapHeight Step 1
        If Pos(I, I2).Tile <> 999 Then
            rDestTemp.Left = Pos(I, I2).X
            rDestTemp.Right = rDestTemp.Left + TerrainTileWidth
            rDestTemp.Top = Pos(I, I2).Y
            rDestTemp.Bottom = Pos(I, I2).Y + TerrainTileHeight
            
            Set ddsTerrainTile = DD.CreateSurfaceFromFile(sApPath & "data\terrain\terr_" & CStr(Pos(I, I2).Tile) & ".bmp", sdTerrainTile)
            
            ddsTerrain.Blt rDestTemp, ddsTerrainTile, rTerrainTile, DDBLT_WAIT
        End If
    Next
Next

'fill the empty terrain with tiles
For I = 1 To MapWidth Step 1
    For I2 = 1 To MapHeight Step 1
        If Pos(I, I2 + 1).Tile = 3 Then
            Set ddsTerrainTile = DD.CreateSurfaceFromFile(sApPath & "data\terrain\terr_6.bmp", sdTerrainTile)
        ElseIf Pos(I, I2 + 1).Tile = 4 Then
            Set ddsTerrainTile = DD.CreateSurfaceFromFile(sApPath & "data\terrain\terr_7.bmp", sdTerrainTile)
        ElseIf Pos(I, I2).Tile = 999 Then
            Set ddsTerrainTile = DD.CreateSurfaceFromFile(sApPath & "data\terrain\terr_8.bmp", sdTerrainTile)
        Else
            Exit For
        End If
        
        rDestTemp.Left = Pos(I, I2).X
        rDestTemp.Right = rDestTemp.Left + TerrainTileWidth
        rDestTemp.Top = Pos(I, I2).Y
        rDestTemp.Bottom = Pos(I, I2).Y + TerrainTileHeight
        ddsTerrain.Blt rDestTemp, ddsTerrainTile, rTerrainTile, DDBLT_WAIT
    Next
Next

'place the player's cannons on some random places
I = Int(Rnd * (UBound(Elev) / 2))
If I = 0 Then I = 1
I2 = Int(Rnd * (UBound(Elev) / 2))
If I2 = 0 Then I2 = 1
I2 = I2 + (UBound(Elev) / 2) + 1

P1.Elev = I
P2.Elev = I2
P1.Angle = 70
P2.Angle = 120
P1.GunPowder = 450
P2.GunPowder = 450
P1.Pos.X = (P1.Elev * 2 * TerrainTileWidth) - TerrainTileWidth - (CannonTileWidth / 2)
P1.Pos.Y = (MapHeight * TerrainTileHeight) - (Elev(P1.Elev) * TerrainTileHeight) - CannonTileHeight
P2.Pos.X = (P2.Elev * 2 * TerrainTileWidth) - TerrainTileWidth - (CannonTileWidth / 2)
P2.Pos.Y = (MapHeight * TerrainTileHeight) - (Elev(P2.Elev) * TerrainTileHeight) - CannonTileHeight

'place some random trees
Dim Rand As Integer
Dim TreeX As Long
Dim TreeY As Long
Dim TreeInterv As Long


TreeInterv = (((2 * TerrainTileWidth) - (NoTreesInterval * 2)) / MaxTreesAtElev)
For I = 1 To UBound(Elev) Step 1
    TreeY = (MapHeight * TerrainTileHeight) - (Elev(I) * TerrainTileHeight) - TreeHeight
    
    For I2 = 0 To MaxTreesAtElev - 1 Step 1
        TreeX = (((I * 2) - 2) * TerrainTileWidth) + NoTreesInterval + (I2 * TreeInterv)
        
        'will there be a tree?
        Rand = Int(Rnd * 2)
        
        'check to make sure there are no cannon at this position
        If TreeX + (TreeWidth / 2) > ((P1.Pos.X + (CannonTileWidth / 2)) - P1NoTrees1) And _
        TreeX < ((P1.Pos.X + (CannonTileWidth / 2)) + P1NoTrees2) Then Rand = 0
        
        If TreeX + (TreeWidth / 2) > ((P2.Pos.X + (CannonTileWidth / 2)) - P1NoTrees1) And _
        TreeX < ((P2.Pos.X + (CannonTileWidth / 2)) + P1NoTrees2) Then Rand = 0
        
        If Rand = 1 Then
            'what tree will there be?
            Rand = Int(Rnd * 5)
            
            Set ddsTree = DD.CreateSurfaceFromFile(sApPath & "data\terrain\tree_" & CStr(Rand + 1) & ".bmp", sdTree)
            ddsTree.SetColorKey DDCKEY_SRCBLT, keyMagenta

            rDestTree.Left = TreeX - (TreeWidth / 2) / 2
            rDestTree.Right = rDestTree.Left + TreeWidth
            rDestTree.Top = TreeY
            rDestTree.Bottom = rDestTree.Top + TreeHeight
            
            ddsTerrain.Blt rDestTree, ddsTree, rTree, DDBLT_WAIT Or DDBLT_KEYSRC
        End If
    Next
Next

'-------
'Show me
'-------
Me.Show

'-------------
'THE MAIN LOOP
'-------------
bLoop = True
Do While bLoop = True
    DoEvents
    MainLoop
Loop

'---
'end
'---
Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MouseX = X
'MouseY = Y
End Sub
Private Sub Form_Unload(Cancel As Integer)

bLoop = False

'returning the old resolution and coop level
DD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
DD.RestoreDisplayMode

'unloading the surfaces
Set ddsTree = Nothing
Set ddsSky = Nothing
Set ddsArrow = Nothing
Set ddsExp = Nothing
Set ddsHole = Nothing
Set ddsCannonBall = Nothing
Set ddsCannon = Nothing
Set ddsTerrainTile = Nothing
Set ddsTerrain = Nothing
Set ddsBack = Nothing
Set ddsPrim = Nothing

'unloading the cipper
'Set TerrClip = Nothing
Set Clip = Nothing

'unloading DX and DD
Set DD = Nothing
Set DX = Nothing

End Sub
Private Sub MainLoop()
'---------------------------------
'calculate the time per loop (ms.)
'---------------------------------
mSperLoop = CalculateTime

'----------------------
'Get the mouse position
'----------------------
MouseX = GetMouseX
MouseY = GetMouseY

'-------------
'Calculate FPS
'-------------
Static Tint As Long
Static FPS As Long
Static Prev As Single
If Tint >= 30 Then
    FPS = 30 / (Timer - Prev)
    Prev = Timer
    Tint = 0
End If
Tint = Tint + 1

'---------------------------------
'Keyboard and mouse input handling
'---------------------------------
'close program
If IsKeyDown(vbKeyEscape) = True Then bLoop = False

'scrolling
If CenterAtCBall = False Then
    If IsKeyDown(vbKeyLeft) = True Or MouseX <= 2 Then rTerrain.Left = rTerrain.Left - 20
        If rTerrain.Left < 0 Then rTerrain.Left = 0
    If IsKeyDown(vbKeyRight) = True Or MouseX >= ResWidth - 2 Then rTerrain.Left = rTerrain.Left + 20
        If rTerrain.Left + ResWidth > (MapWidth * TerrainTileWidth) Then rTerrain.Left = (MapWidth * TerrainTileWidth) - ResWidth
    If IsKeyDown(vbKeyUp) = True Or MouseY <= 2 Then rTerrain.Top = rTerrain.Top - 20
        If rTerrain.Top < 0 Then rTerrain.Top = 0
    If IsKeyDown(vbKeyDown) = True Or MouseY >= ResHeight - 2 Then rTerrain.Top = rTerrain.Top + 20
        If rTerrain.Top + ResHeight > (MapHeight * TerrainTileHeight) Then rTerrain.Top = (MapHeight * TerrainTileHeight) - ResHeight
End If
rTerrain.Right = rTerrain.Left + ResWidth
rTerrain.Bottom = rTerrain.Top + ResHeight

'center Key
If IsKeyDown(vbKeyC) = True Then
    If bCannonBallFlying = True Then
        CenterAtCBall = True
    End If
Else
    CenterAtCBall = False
End If

'------------
'show the sky
'------------
rDestSky.Left = 0
rDestSky.Right = ResWidth
rDestSky.Top = 0
rDestSky.Bottom = ResHeight

rSky.Left = ((SkyWidth - ResWidth) / ((MapWidth * TerrainTileWidth) - ResWidth)) * rTerrain.Left
rSky.Right = rSky.Left + ResWidth
rSky.Top = ((SkyHeight - ResHeight) / ((MapHeight * TerrainTileHeight) - ResHeight)) * rTerrain.Top
rSky.Bottom = rSky.Top + ResHeight

ddsBack.Blt rDestSky, ddsSky, rSky, DDBLT_WAIT
'----------------
'show the terrain
'----------------
rDestTerrain.Left = 0
rDestTerrain.Right = ResWidth
rDestTerrain.Top = 0
rDestTerrain.Bottom = ResHeight

ddsBack.Blt rDestTerrain, ddsTerrain, rTerrain, DDBLT_WAIT Or DDBLT_KEYSRC

'-------------------
'show the cannonball
'-------------------
If bCannonBallFlying = True Then

    Static C1 As Long
    C1 = C1 + mSperLoop
    If C1 >= 1 Then
        'wind
    
        'calculate the position of the cannonball for the current time
        If CurrentPlayer = 1 Then
            CBallX = (P1.Pos.X + ShootFromX) + (XPosAtTime(P1.GunPowder, P1.Angle, T) / 10) '+ (CannonBallWidth / 2)
            CBallY = (P1.Pos.Y + ShootFromY) - (YPosAtTime(P1.GunPowder, P1.Angle, EARTHS_GRAVITY, T) / 10) '+ (CannonBallHeight / 2)
        Else
            CBallX = (P2.Pos.X + ShootFromX) + (XPosAtTime(P2.GunPowder, P2.Angle, T) / 10) '+ (CannonBallWidth / 2)
            CBallY = (P2.Pos.Y + ShootFromY) - (YPosAtTime(P2.GunPowder, P2.Angle, EARTHS_GRAVITY, T) / 10) '+ (CannonBallHeight / 2)
        End If
    
        T = T + 0.4
    
        C1 = 0
    End If
    
    '-------------------------------------------------------------------------
    'find if we're shooting at nearby obstacles and do not destroy us if we do
    '-------------------------------------------------------------------------
    If CurrentPlayer = 1 Then
        If CBallX > P1.Pos.X And CBallX < (P1.Pos.X + CannonTileWidth) And _
        CBallY > P1.Pos.Y And CBallY < P1.Pos.Y + CannonTileHeight Then
        Else
        bCanBeDestroyed = True
        End If
    Else
        If CBallX > P2.Pos.X And CBallX < (P2.Pos.X + CannonTileWidth) And _
        CBallY > P2.Pos.Y And CBallY < P2.Pos.Y + CannonTileHeight Then
        Else
        bCanBeDestroyed = True
        End If
    End If
    '-----------------------------------
    'center the screen at the cannonball
    '-----------------------------------
    If CenterAtCBall = True Then
        If CBallX > (ResWidth / 2) And _
            CBallX < (MapWidth * TerrainTileWidth) - (ResWidth / 2) Then
                    
            rTerrain.Left = CBallX - (ResWidth / 2)
        End If
        If CBallY > (ResHeight / 2) And _
            CBallY < (MapHeight * TerrainTileHeight) - (ResHeight / 2) Then
                    
            rTerrain.Top = CBallY - (ResHeight / 2)
        End If
                
        If CBallX < (ResWidth / 2) Then
            rTerrain.Left = 0
            rTerrain.Top = CBallY - (ResHeight / 2)
        End If
        If CBallX > (MapWidth * TerrainTileWidth) - (ResWidth / 2) Then
            rTerrain.Left = (MapWidth * TerrainTileWidth) - ResWidth
            rTerrain.Top = CBallY - (ResHeight / 2)
        End If
        If CBallY < (ResHeight / 2) Then
            rTerrain.Left = CBallX - (ResWidth / 2)
            rTerrain.Top = 0
        End If
                        
        If rTerrain.Left < 0 Then rTerrain.Left = 0
        If rTerrain.Left + ResWidth > (MapWidth * TerrainTileWidth) Then rTerrain.Left = (MapWidth * TerrainTileWidth) - ResWidth
        If rTerrain.Top < 0 Then rTerrain.Top = 0
        If rTerrain.Top + ResHeight > (MapHeight * TerrainTileHeight) Then rTerrain.Top = (MapHeight * TerrainTileHeight) - ResHeight
            
        rTerrain.Right = rTerrain.Left + ResWidth
        rTerrain.Bottom = rTerrain.Top + ResHeight
    End If
            
    '-----------------------------------
    'check if the ball is out of the map
    '-----------------------------------
    If CBallX > 0 And CBallX < (MapWidth * TerrainTileWidth) And _
    CBallY > 10 And CBallY < (MapHeight * TerrainTileHeight) Then
        
        Static Col As Long
        Static Col2 As Long
        Static I As Long
        Static X As Long
        Static Y As Long
        Static TT As Double
        Static Hit As Boolean
        
        '-------------------------------------------------------------------
        'calculate the pixels within two changes of the cannonbal's position
        'for more precision
        '-------------------------------------------------------------------
        For I = 1 To 16 Step 1
            TT = T - (I * 0.025)
        
            If CurrentPlayer = 1 Then
                X = (P1.Pos.X + ShootFromX) + (XPosAtTime(P1.GunPowder, P1.Angle, TT) / 10) '+ (CannonBallWidth / 2)
                Y = (P1.Pos.Y + ShootFromY) - (YPosAtTime(P1.GunPowder, P1.Angle, EARTHS_GRAVITY, TT) / 10) '+ (CannonBallHeight / 2)
            Else
                X = (P2.Pos.X + ShootFromX) + (XPosAtTime(P2.GunPowder, P2.Angle, TT) / 10) '+ (CannonBallWidth / 2)
                Y = (P2.Pos.Y + ShootFromY) - (YPosAtTime(P2.GunPowder, P2.Angle, EARTHS_GRAVITY, TT) / 10) '+ (CannonBallHeight / 2)
            End If
            
            '------------------------
            'check if we hit a cannon
            '------------------------
            'cannon 1
            If X > P1.Pos.X And X < (P1.Pos.X + CannonTileWidth) And _
            Y > P1.Pos.Y And Y < (P1.Pos.Y + CannonTileHeight) Then
                
                'check the color under the pixel
                ddsCannon.Lock rFill, sdCannon, DDLOCK_WAIT, 0
                Col2 = ddsCannon.GetLockedPixel(P1.rLeft + (X - P1.Pos.X), P1.rTop + (Y - P1.Pos.Y))
                ddsCannon.Unlock rFill
                
                'we have a hit
                If Col2 <> keyMagenta.high Then
                    If bCanBeDestroyed = True Then
                        'destroy the cannon
                        bCannonBallFlying = False
                        CenterAtCBall = False
                        bCanBeDestroyed = False
                        
                        P1.Destroyed = True
                        Exit For
                    End If
                End If
            End If
            
            'cannon 2
            If X > P2.Pos.X And X < (P2.Pos.X + CannonTileWidth) And _
            Y > P2.Pos.Y And Y < (P2.Pos.Y + CannonTileHeight) Then
                'check the color under the pixel
                ddsCannon.Lock rFill, sdCannon, DDLOCK_WAIT, 0
                Col2 = ddsCannon.GetLockedPixel(P2.rLeft + (X - P2.Pos.X), P2.rTop + (Y - P2.Pos.Y))
                ddsCannon.Unlock rFill
                
                'we have a hit
                If Col2 <> keyMagenta.high Then
                    If bCanBeDestroyed = True Then
                        'destroy the cannon
                        bCannonBallFlying = False
                        CenterAtCBall = False
                        bCanBeDestroyed = False
                        
                        P2.Destroyed = True
                        Exit For
                    End If
                End If
            End If
            
            '------------------------------------------
            'check if we hit the terrain or an obstacle
            '------------------------------------------
            ddsTerrain.Lock rFill, sdTerrain, DDLOCK_WAIT, 0
            Col = ddsTerrain.GetLockedPixel(X, Y)
            ddsTerrain.Unlock rFill
            
            'we have a hit
            If Col = keyGreen.high Then
                Hit = False
            Else
                Hit = True
                Exit For
            End If
        Next
        
        If Hit = False Then
            '----------------------------------
            'the cannonball is in still the air
            '----------------------------------
            rDestCannonBall.Left = CBallX - rTerrain.Left - (CannonBallWidth / 2)
            rDestCannonBall.Top = CBallY - rTerrain.Top - (CannonBallHeight / 2)
            rDestCannonBall.Right = rDestCannonBall.Left + CannonBallWidth
            rDestCannonBall.Bottom = rDestCannonBall.Top + CannonBallHeight
            
            ddsBack.Blt rDestCannonBall, ddsCannonBall, rCannonBall, DDBLT_WAIT Or DDBLT_KEYSRC
        Else
            '----------------------------------------------------------
            'we have a hit on the ground or on an obstacle, draw a hole
            'and explosion
            '----------------------------------------------------------
            rDestHole.Left = X - (HoleWidth / 2)
            rDestHole.Top = Y - (HoleHeight / 2)
            rDestHole.Right = rDestHole.Left + HoleWidth
            rDestHole.Bottom = rDestHole.Top + HoleHeight

            ddsTerrain.Blt rDestHole, ddsHole, rHole, DDBLT_KEYSRC Or DDBLT_WAIT
            
            bCannonBallFlying = False
            CenterAtCBall = False
            bCanBeDestroyed = False
            bExplosion = True
        End If
    Else
        '----------------------------------
        'the cannonball is out of the map
        'check if it's out of the boundries
        '----------------------------------
        If CBallX < -300 Or CBallX > (MapWidth * TerrainTileWidth) + 300 Or _
        CBallY < -300 Or CBallY > (MapHeight * TerrainTileHeight) + 300 Then
        
            bCannonBallFlying = False
            CenterAtCBall = False
            bCanBeDestroyed = False
        End If
    End If
    
    '------------------------------------------------
    'check to see if the cannonball out of the screen
    'and display an arrow if it is
    '------------------------------------------------
    'arrow left
    If CBallX < rTerrain.Left And CBallY > rTerrain.Top Then
        rDestArrow.Left = 0
        rDestArrow.Top = (CBallY - rTerrain.Top) - (ArrowHeight / 2)
        rDestArrow.Right = rDestArrow.Left + ArrowWidth
        rDestArrow.Bottom = rDestArrow.Top + ArrowHeight
        
        rArrow.Left = 0
        rArrow.Right = rArrow.Left + ArrowWidth
        rArrow.Top = 0
        rArrow.Bottom = rArrow.Top + ArrowHeight
        
        ddsBack.Blt rDestArrow, ddsArrow, rArrow, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
    
    'arrow up-Left
    If CBallX < rTerrain.Left And CBallY < rTerrain.Top Then
        rDestArrow.Left = 0
        rDestArrow.Top = 0
        rDestArrow.Right = rDestArrow.Left + ArrowWidth
        rDestArrow.Bottom = rDestArrow.Top + ArrowHeight
        
        rArrow.Left = ArrowWidth
        rArrow.Right = rArrow.Left + ArrowWidth
        rArrow.Top = 0
        rArrow.Bottom = rArrow.Top + ArrowHeight
        
        ddsBack.Blt rDestArrow, ddsArrow, rArrow, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
    
    'arrow up
    If CBallX > rTerrain.Left And CBallY < rTerrain.Top Then
        rDestArrow.Left = (CBallX - rTerrain.Left) - (ArrowWidth / 2)
        rDestArrow.Top = 0
        rDestArrow.Right = rDestArrow.Left + ArrowWidth
        rDestArrow.Bottom = rDestArrow.Top + ArrowHeight
        
        rArrow.Left = ArrowWidth * 2
        rArrow.Right = rArrow.Left + ArrowWidth
        rArrow.Top = 0
        rArrow.Bottom = rArrow.Top + ArrowHeight
        
        ddsBack.Blt rDestArrow, ddsArrow, rArrow, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
    
    'arrow up-right
    If CBallX > rTerrain.Right And CBallY < rTerrain.Top Then
        rDestArrow.Left = ResWidth - ArrowWidth
        rDestArrow.Top = 0
        rDestArrow.Right = rDestArrow.Left + ArrowWidth
        rDestArrow.Bottom = rDestArrow.Top + ArrowHeight
        
        rArrow.Left = ArrowWidth * 3
        rArrow.Right = rArrow.Left + ArrowWidth
        rArrow.Top = 0
        rArrow.Bottom = rArrow.Top + ArrowHeight
        
        ddsBack.Blt rDestArrow, ddsArrow, rArrow, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
    
    'arrow right
    If CBallX > rTerrain.Right And CBallY > rTerrain.Top Then
        rDestArrow.Left = ResWidth - ArrowWidth
        rDestArrow.Top = (CBallY - rTerrain.Top) - (ArrowHeight / 2)
        rDestArrow.Right = rDestArrow.Left + ArrowWidth
        rDestArrow.Bottom = rDestArrow.Top + ArrowHeight
        
        rArrow.Left = ArrowWidth * 4
        rArrow.Right = rArrow.Left + ArrowWidth
        rArrow.Top = 0
        rArrow.Bottom = rArrow.Top + ArrowHeight
        
        ddsBack.Blt rDestArrow, ddsArrow, rArrow, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
End If
'-------------------------
'show the cannons
'show the cannons shooting
'-------------------------
Static C3 As Long
Static CannonFireNum As Long
If bShooting = True Then
    C3 = C3 + mSperLoop
    If C3 >= 40 Then
        If CannonFireNum >= (CannonTileSetHeight / CannonTileHeight) - 1 Then
            CannonFireNum = 0
            bShooting = False
        Else
            CannonFireNum = CannonFireNum + 1
        End If
        C3 = 0
    End If
End If

'player 1
rCannon.Left = Int(((CannonTileSetWidth / CannonTileWidth) / 90) * (P1.Angle - 45)) * CannonTileWidth
If rCannon.Left >= CannonTileSetWidth Then rCannon.Left = CannonTileSetWidth - CannonTileWidth
P1.rLeft = rCannon.Left
rCannon.Right = rCannon.Left + CannonTileWidth

If CurrentPlayer = 1 Then
    rCannon.Top = CannonFireNum * CannonTileHeight
Else
    rCannon.Top = 0
End If
P1.rTop = rCannon.Top
rCannon.Bottom = rCannon.Top + CannonTileHeight

rDestCannon.Left = P1.Pos.X - rTerrain.Left
rDestCannon.Top = P1.Pos.Y - rTerrain.Top
rDestCannon.Right = rDestCannon.Left + CannonTileWidth
rDestCannon.Bottom = rDestCannon.Top + CannonTileHeight

ddsBack.Blt rDestCannon, ddsCannon, rCannon, DDBLT_WAIT Or DDBLT_KEYSRC

'player 2
rCannon.Left = Int(((CannonTileSetWidth / CannonTileWidth) / 90) * (P2.Angle - 45)) * CannonTileWidth
If rCannon.Left >= CannonTileSetWidth Then rCannon.Left = CannonTileSetWidth - CannonTileWidth
P2.rLeft = rCannon.Left
rCannon.Right = rCannon.Left + CannonTileWidth

If CurrentPlayer = 2 Then
    rCannon.Top = CannonFireNum * CannonTileHeight
Else
    rCannon.Top = 0
End If
P2.rTop = rCannon.Top
rCannon.Bottom = rCannon.Top + CannonTileHeight

rDestCannon.Left = P2.Pos.X - rTerrain.Left
rDestCannon.Top = P2.Pos.Y - rTerrain.Top
rDestCannon.Right = rDestCannon.Left + CannonTileWidth
rDestCannon.Bottom = rDestCannon.Top + CannonTileHeight

ddsBack.Blt rDestCannon, ddsCannon, rCannon, DDBLT_WAIT Or DDBLT_KEYSRC

'------------------------------
'Show explosions on ground hits
'------------------------------
Static C2 As Long
Static ExpNum As Long
If bExplosion = True Then
     C2 = C2 + mSperLoop
    If C2 >= 40 Then
        If ExpNum >= (ExpTileWidth / ExpWidth) Then
            ExpNum = 0
            bExplosion = False
        Else
            ExpNum = ExpNum + 1
        End If
        C2 = 0
    End If

    rDestExp.Left = X - ExpCenterX - rTerrain.Left
    rDestExp.Top = Y - ExpCenterY - rTerrain.Top
    rDestExp.Right = rDestExp.Left + ExpWidth
    rDestExp.Bottom = rDestExp.Top + ExpHeight

    rExp.Left = ExpNum * ExpWidth
    rExp.Right = rExp.Left + ExpWidth
    rExp.Top = 0
    rExp.Bottom = rExp.Top + ExpHeight

    If bExplosion = True Then
        ddsBack.Blt rDestExp, ddsExp, rExp, DDBLT_KEYSRC Or DDBLT_WAIT
    End If

Else
    ExpNum = 0
End If

'----------------
'cannon destroyed
'----------------
If P1.Destroyed = True Then
    rDestExp.Left = P1.Pos.X - rTerrain.Left
    rDestExp.Top = P1.Pos.Y - rTerrain.Top
    rDestExp.Right = rDestExp.Left + CannonTileWidth
    rDestExp.Bottom = rDestExp.Top + CannonTileHeight
    
    bBigExplosion = True
End If

If P2.Destroyed = True Then
    rDestExp.Left = P2.Pos.X - rTerrain.Left
    rDestExp.Top = P2.Pos.Y - rTerrain.Top
    rDestExp.Right = rDestExp.Left + CannonTileWidth
    rDestExp.Bottom = rDestExp.Top + CannonTileHeight

    bBigExplosion = True
End If

'show explosion
Static C4 As Long
Static BigExpNum As Long
If bBigExplosion = True And bBigExplodeOnce = False Then
    C4 = C4 + mSperLoop
    If C4 >= 40 Then
        If BigExpNum >= (ExpTileWidth / ExpWidth) Then
            BigExpNum = 0
            bBigExplosion = False
            bBigExplodeOnce = True
        Else
            BigExpNum = BigExpNum + 1
        End If
        C4 = 0
    End If
    
    rExp.Left = BigExpNum * ExpWidth
    rExp.Right = rExp.Left + ExpWidth
    rExp.Top = 0
    rExp.Bottom = rExp.Top + ExpHeight
    
    If bBigExplosion = True Then
        ddsBack.Blt rDestExp, ddsExp, rExp, DDBLT_KEYSRC Or DDBLT_WAIT
    End If
End If

'--------------
'draw info text
'--------------
ddsBack.DrawText 10, 10, "FPS: " & CStr(FPS), False
ddsBack.DrawText 10, 30, "Mouse X/Y: " & CStr(MouseX) & "/" & CStr(MouseY), False
ddsBack.DrawText 10, 50, "RelMouse X/Y: " & CStr(MouseX + rTerrain.Left) & "/" & CStr(MouseY + rTerrain.Top), False
ddsBack.DrawText 10, 70, "Cannon 1 XY: " & CStr(P1.Pos.X) & "/" & CStr(P1.Pos.Y), False
ddsBack.DrawText 10, 90, "Cannon 2 XY: " & CStr(P2.Pos.X) & "/" & CStr(P2.Pos.Y), False
ddsBack.DrawText 10, 110, "CannonBallFlying: " & CStr(bCannonBallFlying), False
ddsBack.DrawText 10, 130, "Explosion - small: " & CStr(bExplosion), False
ddsBack.DrawText 10, 150, "Shooting: " & CStr(bShooting), False
ddsBack.DrawText 10, 170, "Ball XY: " & CStr(Int(CBallX)) & "/" & CStr(Int(CBallY)), False
ddsBack.DrawText 10, 190, "CanBeDestroyed: " & CStr(bCanBeDestroyed), False

ddsBack.DrawText 300, 10, "P1 Angle: " & CStr(Int(P1.Angle)), False
ddsBack.DrawText 300, 30, "P1 Gunpowder: " & CStr(Int(P1.GunPowder)), False

ddsBack.DrawText 500, 10, "Current Player: " & CStr(Int(CurrentPlayer)), False

ddsBack.DrawText 700, 10, "P2 Angle: " & CStr(Int(P2.Angle)), False
ddsBack.DrawText 700, 30, "P2 Gunpowder: " & CStr(Int(P2.GunPowder)), False

'----
'flip
'----
ddsPrim.Flip Nothing, DDFLIP_WAIT

End Sub
