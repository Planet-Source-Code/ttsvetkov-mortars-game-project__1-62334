Attribute VB_Name = "Globals"
Option Explicit

'============================================
'Place to store global variables or functions
'============================================

'DX
Public DX As DirectX7
Public DD As DirectDraw7

'=============================================================
'primary
Public ddsPrim As DirectDrawSurface7
Public sdPrim As DDSURFACEDESC2

'back
Public ddsBack As DirectDrawSurface7
Public sdBack As DDSURFACEDESC2
Public Clip As DirectDrawClipper
Public rClip(0) As RECT

'terrain
Public ddsTerrain As DirectDrawSurface7
Public sdTerrain As DDSURFACEDESC2
Public rTerrain As RECT
Public rDestTerrain As RECT

Public Const MapWidth = 20
Public Const MapHeight = 16

'terrain tile
Public ddsTerrainTile As DirectDrawSurface7
Public sdTerrainTile As DDSURFACEDESC2
Public rTerrainTile As RECT

Public Const TerrainTileWidth As Integer = 100
Public Const TerrainTileHeight As Integer = 100

'cannon
Public ddsCannon As DirectDrawSurface7
Public sdCannon As DDSURFACEDESC2
Public rCannon As RECT
Public rDestCannon As RECT

Public Const CannonTileWidth As Long = 120
Public Const CannonTileHeight As Long = 100
Public Const CannonTileSetWidth As Long = 1200
Public Const CannonTileSetHeight As Long = 600

'cannonball
Public ddsCannonBall As DirectDrawSurface7
Public sdCannonBall As DDSURFACEDESC2
Public rCannonBall As RECT
Public rDestCannonBall As RECT

Public Const CannonBallWidth As Long = 10
Public Const CannonBallHeight As Long = 10
Public Const ShootFromX As Long = 60
Public Const ShootFromY As Long = 75

'the blast hole
Public ddsHole As DirectDrawSurface7
Public sdHole As DDSURFACEDESC2
Public rHole As RECT
Public rDestHole As RECT

Public Const HoleWidth As Long = 30
Public Const HoleHeight As Long = 30

'the explosion
Public ddsExp As DirectDrawSurface7
Public sdExp As DDSURFACEDESC2
Public rExp As RECT
Public rDestExp As RECT

Public Const ExpWidth As Long = 45
Public Const ExpHeight As Long = 50
Public Const ExpTileWidth As Long = 765
Public Const ExpTileHeight As Long = 50
Public Const ExpCenterX As Long = 22
Public Const ExpCenterY As Long = 35

'the arrow
Public ddsArrow As DirectDrawSurface7
Public sdArrow As DDSURFACEDESC2
Public rArrow As RECT
Public rDestArrow As RECT

Public Const ArrowWidth As Long = 40
Public Const ArrowHeight As Long = 40
Public Const ArrowTileWidth As Long = 200
Public Const ArrowTileHeight As Long = 40

'the sky
Public ddsSky As DirectDrawSurface7
Public sdSky As DDSURFACEDESC2
Public rSky As RECT
Public rDestSky As RECT

Public Const SkyWidth As Long = 2100
Public Const SkyHeight As Long = 1700

'trees
Public ddsTree As DirectDrawSurface7
Public sdTree As DDSURFACEDESC2
Public rTree As RECT
Public rDestTree As RECT

Public Const TreeWidth As Long = 50
Public Const TreeHeight As Long = 100
Public Const MaxTreesAtElev As Long = 6
Public Const NoTreesInterval As Long = 30

Public Const P1NoTrees1 As Long = 10
Public Const P1NoTrees2 As Long = 16
Public Const P2NoTrees1 As Long = 16
Public Const P2NoTrees2 As Long = 10

'--------------
'the Color Keys
'--------------
Public keyMagenta As DDCOLORKEY
Public keyGreen As DDCOLORKEY

'================================================================

Public Const ResWidth As Integer = 1024
Public Const ResHeight As Integer = 768

'is the main loop running?
Public bLoop As Boolean

'the mouse position on the screen
Public MouseX As Long
Public MouseY As Long

'a COORDS private type
Public Type COORDS
    X As Long
    Y As Long
    Tile As Integer
End Type

'the player's cannon props
Public Type PLAYER
    Elev As Integer
    Angle As Double
    Pos As COORDS
    GunPowder As Double
    Destroyed As Boolean
    rLeft As Long
    rTop As Long
End Type
Public P1 As PLAYER
Public P2 As PLAYER

'the elevations
Public Elev(1 To (MapWidth / 2)) As Long
