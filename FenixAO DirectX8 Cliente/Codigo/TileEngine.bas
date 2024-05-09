Attribute VB_Name = "modTileEngine"
'FenixAO DirectX8
'Engine By ∑Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester



Option Explicit
Option Base 0

Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Private Const MAX_DIALOGOS = 300
Private Const MAXLONG = 15

Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'RGB Type
Public Type RGB
    r As Long
    G As Long
    b As Long
End Type

Dim map_current As map
Dim char_last As Long
Dim char_list() As Char
Dim char_count As Long

Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    Grh(1 To 3) As Grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    Blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
    
    exit_index As Long
    npc_index As Long
    item_index As Long
    
    Trigger As Byte
End Type

Private Type map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan MartÌn Sotuyo Dodero
    base_light_color As Long
End Type



Rem Mannakia .. Parituclas ORE 1.0.
 
Private Type Particle
    TimeAlpha As Single
    alpha As Single
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list As D3DCOLORVALUE
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Dim base_tile_size As Integer
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    speed As Single
    life_counter As Long
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
 
    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte
 
    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(3) As Long
    
    'Added by Juan MartÌn Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system
 
'Dim StreamData() As particle_group
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Public particle_group_count As Long
Dim particle_group_last As Long
Rem mannakia

Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherDoFog As Byte
Public WeatherFogCount As Byte
 
Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer
Public EndTime As Long
Private Const ScreenWidth As Long = 534
Private Const ScreenHeight As Long = 408

 

Public Const PI As Single = 3.14159265358979


'Textos>>>
Dim font_count As Long
Dim font_last As Long
Public font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'Cargas de texto desde GRH
Private Type tFont
    Caracteres(0 To 255) As Integer 'indice de cada letra
End Type
 
Private Fuentes(1) As tFont

'Textos<<<

Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private lFrameLimiter As Long
Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public timerTicksPerFrame As Double 'mmmm me encanta que sea Double jaja
Public timerElapsedTime As Single
Public particletimer As Single
Public engineBaseSpeed As Single

'Describes a transformable lit vertex
Private Type TLVERTEX
  X As Single
  Y As Single
  Z As Single
  rhw As Single
  color As Long
  Specular As Long
  tu As Single
  tv As Single
End Type


'********** Direct X ***********
Private Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Integer
    texheight As Integer
End Type

Private Type D3D8Textures2
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

'DirectX 8 Objects
Public Dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Font List
Public FontList As D3DXFont
Public FontDesc As IFont

Private Type tLight
    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_x As Byte
    map_y As Byte
    range As Byte
    id As Long
End Type
 
Private light_list() As tLight
Private NumLights As Byte
Dim light_count As Long
Dim light_last As Long

Public mFreeMemoryBytes As Long

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Base_Light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte
Type luzxhora
    r As Long
    G As Long
    b As Long
End Type
Public luz_dia(0 To 24) As luzxhora

Public Const ImgSize As Byte = 4

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020

Public Type Position
    X As Integer
    Y As Integer
End Type


Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    speed As Integer
    MiniMap_color As Long
    active As Boolean
End Type

Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

Public Type HeadData
    Head(1 To 4) As Grh
End Type

Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type

Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type

Public Type FxData
    FX As Grh
    OffsetX As Long
    OffsetY As Long
End Type

Public Type Char
    
    r As Byte
    G As Byte
    b As Byte
    Particula As Integer
    Aura_Index As Integer
    Aura As Grh
    active As Byte
    Heading As Byte
    POS As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    GM As Integer
    
    haciendoataque As Byte
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    
End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    particle_group_index As Integer
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    light_value(3) As Long
    
    luz As Integer
    color(3) As Long
    
    particle_group As Integer
End Type

Public IniPath As String
Public MapPath As String

Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public LastChar As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh
Public MapData() As MapBlock
Public CharList(1 To 10000) As Char

Public bRain        As Boolean
Public bTecho       As Boolean

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
    plFogata = 3
End Enum

'PARTICULAS MENDUZ
Private Enum PARTICLE_STATUS2
    Alive = 0
    Dead = 1
End Enum

Private Type PARTICLE2
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    lifeTime As Long    'How long Mr. Particle Exists
    created As Long 'When this particle was created...
    Status As PARTICLE_STATUS2 'Does he even exist?
End Type

Private Type pa_gro2
    PrtData() As PARTICLE2
    PrtVertList() As TLVERTEX
    Position As D3DVECTOR
    light As D3DLIGHT8
    type As Integer
    nParticles As Long
    ParticleSize As Single
    gravity As Single
    XWind As Single
    ZWind As Single
    YWind As Single
    XVariation As Single
    YVariation As Single
    ZVariation As Single
    X As Single
    Y As Single
    Z As Single
    activated As Boolean
    texture As Integer
    size As Single
    Life As Integer
End Type

Dim particle_group_list2() As pa_gro2
Dim particle_group_count2 As Integer
Dim particle_group_last2 As Integer
'PARTICULAS MENDUZ

'LLUVIA
Private Enum PARTICLE_STATUS3
    alive2 = 0
    dead2 = 1
End Enum

Private Type PARTICLE3
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    lifeTime As Long    'How long Mr. Particle Exists
    created As Long 'When this particle was created...
    Status As PARTICLE_STATUS3 'Does he even exist?
End Type

Private Type pa_gro3
    PrtData() As PARTICLE3
    PrtVertList() As TLVERTEX
    Position As D3DVECTOR
    light As D3DLIGHT8
    type As Integer
    nParticles As Long
    ParticleSize As Single
    gravity As Single
    XWind As Single
    ZWind As Single
    YWind As Single
    XVariation As Single
    YVariation As Single
    ZVariation As Single
    X As Single
    Y As Single
    Z As Single
    activated As Boolean
    texture As Integer
    size As Single
    Life As Integer
End Type

Dim particle_group_list3() As pa_gro3
Dim particle_group_count3 As Integer
Dim particle_group_last3 As Integer
'PARTICULAS MENDUZ

Private Type tCache
    Number        As Long
    SrcHeight     As Single
    SrcWidth      As Single
End Type: Private Cache As tCache

'BitBlt
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Function GetElapsedTime() As Single
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
Public Sub ShowNextFrame()

Dim OffsetCounterX As Single
Dim OffsetCounterY As Single
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Long
Dim loopc As Long

Do While prgRun

    Call RefreshAllChars
    
    If EngineRun Then
        If frmMain.WindowState <> 1 Then

        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.X))
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
    
    D3DDevice.BeginScene
     D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 1#, 0
            
            
            If UserCiego Then
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            Else
                RenderScreen UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY
            End If
            
            'FPS
            Dim color As Long
            If FramesPerSec >= 30 Then
            color = D3DColorXRGB(255, 255, 255)
            ElseIf FramesPerSec >= 15 Then
            color = D3DColorXRGB(255, 255, 0)
            ElseIf FramesPerSec >= 1 Then
            color = D3DColorXRGB(255, 0, 0)
            End If
            Call Grh_Text_Render(True, "FPS: " & FramesPerSec, 10, 10, color)
            '/FPS

        
            
            Engine_ActFPS
            
            If ModoTrabajo Then Grh_Text_Render True, "MODO TRABAJO", 40, 10, D3DColorARGB(255, 255, 0, 0)
            If Cartel Then DibujarCartel
            If Dialogos.CantidadDialogos <> 0 Then Dialogos.MostrarTexto
            RenderSounds
                
        D3DDevice.Present ByVal 0, ByVal 0, frmMain.Renderer.hwnd, ByVal 0
    D3DDevice.EndScene
    
            lFrameLimiter = GetTickCount
            FramesPerSecCounter = FramesPerSecCounter + 1
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
            particletimer = timerElapsedTime * 0.05
            
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    
    If Not Pausa And frmMain.Visible And Not frmForo.Visible Then
        CheckKeys
    End If

    If GetTickCount - lFrameTimer > 1000 Then
        FramesPerSec = FramesPerSecCounter
        If FPSFLAG Then frmMain.Caption = "FenixAO DirectX8" & " v" & App.Major & "." & App.Minor & "." & App.Revision
        frmMain.fpstext.Caption = FramesPerSec
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
    
    'Limitar FPS
    While (GetTickCount - lFrameTimer) \ 30 < FramesPerSecCounter
    Sleep 5
    Wend
    
    ' ### I N T E R V A L O S ###
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
        timers(loopc) = timers(loopc) + (esttick - ulttick)
        
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
    Next loopc
    ulttick = GetTickCount
    
    DoEvents
Loop

End Sub

Sub DDrawTransGrhtoSurface(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)

Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub


iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If

If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        False, 0

End Sub
Sub DrawGrhtoHdc(hdc As Long, GrhIndex As Integer)

    Dim hDCsrc As Long
 
    If GrhIndex <= 0 Then Exit Sub
        
        'If it's animated switch GrhIndex to first frame
        If GrhData(GrhIndex).NumFrames <> 1 Then
            GrhIndex = GrhData(GrhIndex).Frames(1)
        End If
           
        hDCsrc = CreateCompatibleDC(hdc)
        
        Call SelectObject(hDCsrc, LoadPicture(App.Path & "\Graficos\" & GrhData(GrhIndex).FileNum & ".bmp"))

        'Draw
        BitBlt hdc, 0, 0, _
        GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelWidth, _
        hDCsrc, _
        GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, _
        vbSrcCopy

        DeleteDC hDCsrc
End Sub
Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan MartÌn Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
On Error Resume Next 'Thusing
    
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim TempChar As Char
    Dim Moved    As Byte
    Dim iPPx     As Integer
    Dim iPPy     As Integer
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)
    
                Dim Blanco(3) As Long
                Blanco(0) = RGB(255, 255, 255)
                Blanco(1) = RGB(255, 255, 255)
                Blanco(2) = RGB(255, 255, 255)
                Blanco(3) = RGB(255, 255, 255)
    
                   Dim ColorAura(3) As Long
                ColorAura(0) = RGB(255, 255, 255)
                ColorAura(1) = RGB(255, 255, 255)
                ColorAura(2) = RGB(255, 255, 255)
                ColorAura(3) = RGB(255, 255, 255)
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
            
#If HARDCODED = True Then
                    If MapData(X, Y).Graphic(1).Started = 1 Then
                        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames / MapData(X, Y).Graphic(1).speed)
                        If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames Then
                            MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames) + 1
                            
                            If MapData(X, Y).Graphic(1).Loops <> -1 Then
                                If MapData(X, Y).Graphic(1).Loops > 0 Then
                                    MapData(X, Y).Graphic(1).Loops = MapData(X, Y).Graphic(1).Loops - 1
                                Else
                                    MapData(X, Y).Graphic(1).Started = 0
                                End If
                            End If
                        End If
                    End If

                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(MapData(X, Y).Graphic(1).FrameCounter)

                Device_Box_Textured_Render CurrentGrhIndex, _
                    (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, _
                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    MapData(X, Y).light_value, _
                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    False _
                    , 0
#Else
                Call Draw_Grh(MapData(X, Y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, MapData(X, Y).light_value(), , , , , X, Y)
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
                End If
#End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw floor layer 2
'    ScreenY = minYOffset
'    For Y = screenminY To screenmaxY
'        ScreenX = minXOffset
'        For X = screenminX To screenmaxX
                'Layer 2 **********************************
'                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
'#If HARDCODED = True Then
'                    If MapData(X, Y).Graphic(2).Started = 1 Then
'                        MapData(X, Y).Graphic(2).FrameCounter = MapData(X, Y).Graphic(2).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames / MapData(X, Y).Graphic(2).Speed)
'                        If MapData(X, Y).Graphic(2).FrameCounter > GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames Then
'                            MapData(X, Y).Graphic(2).FrameCounter = (MapData(X, Y).Graphic(2).FrameCounter Mod GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames) + 1
'
'                            If MapData(X, Y).Graphic(2).Loops <> -1 Then
'                                If MapData(X, Y).Graphic(2).Loops > 0 Then
'                                    MapData(X, Y).Graphic(2).Loops = MapData(X, Y).Graphic(2).Loops - 1
'                                Else
'                                    MapData(X, Y).Graphic(2).Started = 0
'                                End If
'                            End If
'                        End If
'                    End If
'
'                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(MapData(X, Y).Graphic(2).FrameCounter)
'
'                offx = 0
'                offy = 0
'                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                If GrhData(MapData(X, Y).Graphic(2).grhindex).TileHeight <> 1 Then _
'                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                Device_Box_Textured_Render CurrentGrhIndex, _
'                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
'                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                    MapData(X, Y).light_value, _
'                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                    False _
'                    , 0
'#Else
''                    Call Draw_Grh(MapData(X, Y).Graphic(2), _
'                            (ScreenX - 1) * 32 + PixelOffsetX, _
'                            (ScreenY - 1) * 32 + PixelOffsetY, _
'                            1, 1, , X, Y)
'#End If
''                End If
'
''            ScreenX = ScreenX + 1
''        Next X'
'
'        'Reset ScreenX to original value and increment ScreenY
'        'ScreenX = ScreenX - X + screenminX
'        'ScreenY = ScreenY + 1
'    'Next Y

    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                'Object Layer **********************************
         '       If .ObjGrh.GrhIndex <> 0 Then
         '       If Abs(nX - X) < 1 And (Abs(nY - Y)) < 1 And MapData(X, Y).Blocked = 0 Then
         '           Call Draw_Grh(.ObjGrh, _
         '                   PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
         '       End If


                If .ObjGrh.GrhIndex <> 0 Then
                Call Draw_Grh(.ObjGrh, _
                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
                            End If

                
                'Char layer ************************************
If MapData(X, Y).CharIndex > 0 Then
                TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
    
            If TempChar.MoveOffset.X <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                TempChar.MoveOffset.X = TempChar.MoveOffset.X - (8 * Sgn(TempChar.MoveOffset.X))
                Moved = 1
            End If

            If TempChar.MoveOffset.Y <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (8 * Sgn(TempChar.MoveOffset.Y))
                Moved = 1
            End If

            If Moved = 0 And TempChar.Moving = 1 Then
                TempChar.Moving = 0
                TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                TempChar.Body.Walk(TempChar.Heading).Started = 0
                TempChar.arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 0
                TempChar.escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                TempChar.escudo.ShieldWalk(TempChar.Heading).Started = 0
                TempChar.haciendoataque = 0
            End If
            
            If TempChar.haciendoataque = 0 And TempChar.MoveOffset.X = 0 And TempChar.MoveOffset.Y = 0 Then
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 0
                TempChar.arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                End If
            If TempChar.haciendoataque = 1 Then
                TempChar.arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.haciendoataque = 0
            End If
            
            
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp + 32
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp + 32
                
                If Len(TempChar.Nombre) = 0 Then
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                Else
                    If TempChar.Navegando = 1 Then
                                           'Aura
                        If TempChar.Aura_Index > 0 Then
                        Call Draw_Grh(TempChar.Aura, iPPx, iPPy + 40, 1, 1, ColorAura(), True, X, Y)
                        End If
                        'Cuerpo (Barca / Galeon / Galera)
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                    
                    ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        
                        'Cuerpo
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                        
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                        
                        'Casco
                        If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                        
                        'Arma
                        If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If
                        
                        'Escudo
                        If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If
                    
                    End If
                        
                    If Nombres Then
                        
                        If Not (TempChar.invisible Or TempChar.Navegando = 1) Then
                       
                            Dim lCenter As Long
                            If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                Dim sClan As String
                                lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan = Mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                Text_Render font_list(1), Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), iPPy + 30, iPPx - lCenter, 100, 20, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)), DT_TOP Or DT_LEFT, True
                                lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                Text_Render font_list(1), sClan, iPPy + 45, iPPx - lCenter, 200, 20, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)), DT_TOP Or DT_LEFT, True
                            Else
                                lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                Text_Render font_list(1), TempChar.Nombre, iPPy + 30, iPPx - lCenter, 100, 20, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)), DT_TOP Or DT_LEFT, True
                            End If
                      
                        End If
                       
                    End If
                End If
    
                If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + TempChar.Body.HeadOffset.X), (iPPy + TempChar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)
                
                CharList(MapData(X, Y).CharIndex) = TempChar

                If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call Draw_Grh(FxData(TempChar.FX).FX, iPPx + FxData(TempChar.FX).OffsetX, iPPy + FxData(TempChar.FX).OffsetY, 1, 1, Blanco(), True, , , MapData(X, Y).CharIndex)
                
            End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), _
                            ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , X, Y)
'#End If
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

            ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                If MapData(X, Y).particle_group_index Then
                    Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY  '+ (16)
                End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
 D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    Dim tmplng As Long
    Dim tmblng2 As Long
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            With MapData(X, Y)
                '***********************************************
                If .particle_group > 0 Then
                    Particles_Update .particle_group, ScreenY * 32 + PixelOffsetY - 32, ScreenX * 32 + PixelOffsetX - 32
        
                    
                    D3DDevice.SetTexture 0, GetTexture(GrhData(particle_group_list2(.particle_group).texture).FileNum, Val(tmplng), Val(tmblng2)) 'Textures(particle_group_list(.particle_group).Texture)
                    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(particle_group_list2(.particle_group).size)
                    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, particle_group_list2(.particle_group).nParticles, particle_group_list2(.particle_group).PrtVertList(0), Len(particle_group_list2(.particle_group).PrtVertList(0))

                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
                    
                    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
   
 
 
If Not bTecho Then
        'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                ScreenX * 32 + PixelOffsetX, _
                ScreenY * 32 + PixelOffsetY, _
                1, 1)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
        Else
        ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, True, X, Y)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
            
        'If LuzMouse Then
        '    Light_Move (Light_Find(20)), UserPos.X + frmMain.MouseX \ 32 - frmMain.Renderer.ScaleWidth \ 64, UserPos.Y + frmMain.MouseY / 32 - frmMain.Renderer.ScaleHeight \ 64
        'End If
        Light_Render_All
    
    If Niebla Then
    WeatherDoFog = 10
    Engine_Weather_UpdateFog
End If
 
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

    
End Sub
Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviain.wav", 0, 0, Enabled)
                    frmMain.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviaout.wav", 0, 0, Enabled)
                    frmMain.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer) As Boolean

IniPath = App.Path & "\Init\"


UserPos.X = MinXBorder
UserPos.Y = MinYBorder

TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

Call LoadGrhData
Call CargarParticulas
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs
Call CargarAnimArmas
Call CargarAnimEscudos

    HalfWindowTileHeight = (frmMain.Renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.Renderer.ScaleWidth / 32) \ 2

    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32




'Parra: Aca inician las variables globales del Directx8
                    

    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
    
    
If Not InitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING, setDisplayFormhWnd) Then
        If Not InitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING, setDisplayFormhWnd) Then
            MsgBox "El D3DDevice no pudo iniciar..."
            End
        End If
    End If
    
    D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    D3DDevice.SetVertexShader FVF
    
    
    
    'partÌculas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

        '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    




    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.Renderer.ScaleWidth
        .BackBufferHeight = frmMain.Renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.Renderer.hwnd
    End With
    DispMode.format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format
        End If
    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.Renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.Renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.Renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    


    engineBaseSpeed = 0.016
    
    'partÌculas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'Fuentes
    Font_Create "Tahoma", 8, True, 0
    Font_Create "Morpheus", 42, True, 0

    Call Engine_Font_Initialize
    
    
    'Set Memory Status
    GlobalMemoryStatus pUdtMemStatus
    mFreeMemoryBytes = pUdtMemStatus.dwAvailPhys
    
    'Luz del dia
    Call Base_Luz(125, 125, 125)


Light_Remove_All
    
    
    Light_Render_All


InitTileEngine = True
End Function
Private Function InitD3DDevice(ByVal mode As CONST_D3DCREATEFLAGS, ByRef setDisplayFormhWnd As Long) As Boolean

    'When there is an error, destroy the D3D device and get ready to make a new one
    On Error GoTo ErrOut
    
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3DWindow.Windowed = True
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.format
    
  '###################################
  '## CHECK THE DEVICE CAPABILITIES ##
  '###################################
    
    Dim DevCaps As D3DCAPS8
    
    D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps
    
    If Err.Number = D3DERR_INVALIDDEVICE Then
        'We couldn't get data from the hardware device - probably doesn't exist...
        D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, DevCaps
        Err.Clear
    End If
    
    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    'Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, setDisplayFormhWnd, mode, D3DWindow)
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.Renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    frmMain.Visible = False
    DoEvents
    
    'Everything was successful
    InitD3DDevice = True
    
Exit Function

ErrOut:
    MsgBox "Error Number Returned: " & Err.Number & vbNewLine & "Description: " & Err.Description
    
    'Return a failure
    InitD3DDevice = False
End Function
Public Sub DeInitTileEngine()

    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).texture = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i

    Set Dx = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set FontList = Nothing
    
    Erase CharList
    Erase Grh
    Erase GrhData
    Erase MapData
End Sub

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, G As Integer, b As Integer, a As Integer
    r = 127 * Vec.X + 128
    G = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, G, b)
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan MartÌn Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, range, id, Red, Green, Blue
    End If
End Function
 
Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function
 
Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan MartÌn Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addY As Byte
    Dim addX As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If
 
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
 
 
        Select Case Heading
            Case NORTH
                addY = -1
        
            Case EAST
                addX = 1
        
            Case SOUTH
                addY = 1
            
            Case WEST
                addX = -1
        End Select
        
        nX = map_x + addX
        nY = map_y + addY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index
 
            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function
 
Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                        ByVal range As Long, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
        'Le damos color
    light_list(light_index).RGBcolor.r = Red
    light_list(light_index).RGBcolor.G = Green
    light_list(light_index).RGBcolor.b = Blue
   
    'Alpha (Si borras esto RE KB!!)
    light_list(light_index).RGBcolor.a = 255
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub
 
Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function
 
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            LightRender loop_counter
        End If
    
    Next loop_counter
End Sub
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.G), Round(CurrentColor.b))
        'If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function
Private Sub LightRender(ByVal light_index As Integer)
 
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.r = ColorLuz.r
    AmbientColor.G = ColorLuz.G
    AmbientColor.b = ColorLuz.b

   
    LightColor = light_list(light_index).RGBcolor
       
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If
        Next Xa
    Next Ya
End Sub

 
Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function
 
Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function
 
Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function
 
Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As tLight
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub
 
Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan MartÌn Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub
Private Function CreateColorVal(a As Integer, r As Integer, G As Integer, b As Integer) As D3DCOLORVALUE
    CreateColorVal.a = a
    CreateColorVal.r = r
    CreateColorVal.G = G
    CreateColorVal.b = b
End Function
Public Function ARGB(ByVal r As Long, ByVal G As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    End If
    
    ARGB = c

End Function


Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)

    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = Base_Light
    If (light_value(1) = 0) Then light_value(1) = Base_Light
    If (light_value(2) = 0) Then light_value(2) = Base_Light
    If (light_value(3) = 0) Then light_value(3) = Base_Light
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function
Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Integer, Optional ByRef Textures_Height As Integer, Optional ByVal angle As Single)

    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If

End Sub
Private Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
If FileName = 0 Then Debug.Print "ERROR! GRH = 0": Exit Function

    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).texture
                Exit Function
            End If
        Next i
    End With

    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
    
            Set .texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\GRAFICOS\" & LTrim(Str(Archivo)) & ".bmp", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                
            .texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.width
            .texture_height = texture_info.height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .texture
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.size
        End With
    End With
    
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp"
End Function

Private Function RemoveLRU() As Boolean
    
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
    
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).texture.GetLevelDesc(0, surface_desc)
    
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function

Sub Draw_Grh(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef color() As Long, Optional alpha As Boolean, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
'***************************
'/////By Thusing/////
'***************************

Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub


iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If

If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        color(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        alpha, 0


End Sub

Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.bottom = Top + height
    TextRect.Right = Left + width
    
    If shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 1
        ShadowRect.bottom = (Top + height) - 1
        ShadowRect.Right = (Left + width) - 1
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
    End If
    
    D3DX.DrawText font, color, Text, TextRect, format
End Sub
Public Sub Text_Render_ext(ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)
    If center = True Then
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, DT_TOP Or DT_LEFT, shadow)
    End If
End Sub

Public Sub Text_Render_Bordes(ByVal Borde As Byte, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                            ByVal width As Long, ByVal height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)

    If center = True Then
        Call Text_Render(font_list(1), Text, Top - Borde, Left, width, height, D3DColorXRGB(0, 0, 0), fa_center, shadow)
        Call Text_Render(font_list(1), Text, Top, Left - Borde, width, height, D3DColorXRGB(0, 0, 0), fa_center, shadow)
        Call Text_Render(font_list(1), Text, Top + Borde, Left, width, height, D3DColorXRGB(0, 0, 0), fa_center, shadow)
        Call Text_Render(font_list(1), Text, Top, Left + Borde, width, height, D3DColorXRGB(0, 0, 0), fa_center, shadow)
        
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top - Borde, Left, width, height, D3DColorXRGB(0, 0, 0), DT_TOP Or DT_LEFT, shadow)
        Call Text_Render(font_list(1), Text, Top, Left - Borde, width, height, D3DColorXRGB(0, 0, 0), DT_TOP Or DT_LEFT, shadow)
        Call Text_Render(font_list(1), Text, Top + Borde, Left, width, height, D3DColorXRGB(0, 0, 0), DT_TOP Or DT_LEFT, shadow)
        Call Text_Render(font_list(1), Text, Top, Left + Borde, width, height, D3DColorXRGB(0, 0, 0), DT_TOP Or DT_LEFT, shadow)
    
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, DT_TOP Or DT_LEFT, shadow)
    End If
    
    
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal Bold As Boolean, _
                        ByVal Italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = style
    fnt.size = size
    fnt.Bold = Bold
    fnt.Italic = Italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal Bold As Boolean, _
                            ByVal Italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, Bold, Italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Public Sub Particle_Group_Make_Menduz( _
ByVal particle_group_index As Integer, _
ByVal map_z As Integer, _
ByVal map_x As Integer, _
ByVal map_y As Integer, _
ByVal var_z As Single, _
ByVal var_x As Single, _
ByVal var_y As Single, _
ByVal start_a As Integer, _
ByVal start_r As Integer, _
ByVal start_g As Integer, _
ByVal start_b As Integer, _
ByVal end_a As Integer, _
ByVal end_r As Integer, _
ByVal end_g As Integer, _
ByVal end_b As Integer, _
ByVal n_particles As Integer, _
ByVal gravity As Single, _
ByVal texture As Integer, _
ByVal size As Single, _
ByVal Life As Integer)
    'If particle_group_index > particle_group_last2 Then
    '    particle_group_last2 = particle_group_index
    
    ReDim Preserve particle_group_list2(0 To particle_group_index)
    With particle_group_list2(particle_group_index)
        ReDim Preserve .PrtData(0 To n_particles)
        ReDim Preserve .PrtVertList(0 To n_particles)
        'Map pos
        .X = map_x
        .Y = map_y
        .Z = map_z
        .Life = Life
        .size = size
        .gravity = gravity '-0.5
        'Sets alive counter
        .nParticles = n_particles
        .texture = texture
        Dim i As Integer
        .XVariation = var_x '0.5
        .YVariation = var_y '0.85
        .ZVariation = var_z '0.5
        Dim nowa As Long
        nowa = GetTickCount
        For i = 0 To .nParticles
            .PrtData(i).Status = Alive
            .PrtData(i).lifeTime = 1000 + (Rnd * 500)
            .PrtData(i).created = nowa
            .PrtData(i).X = map_x
            .PrtData(i).Y = map_y
            .PrtData(i).vX = (Rnd * var_x) - (var_x / 2)
            .PrtData(i).vY = (Rnd * var_y) - (var_y / 3)
            
            'Randomize
            .PrtData(i).StartColor = CreateColorVal(start_a, start_r, start_g, start_g)
            .PrtData(i).EndColor = CreateColorVal(end_a, end_r, end_g, end_b)
            
            .PrtData(i).CurrentColor = .PrtData(i).StartColor
            .PrtVertList(i).X = .PrtData(i).X
            .PrtVertList(i).rhw = 1
            .PrtVertList(i).Y = .PrtData(i).Y '
            .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.a, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.G, .PrtData(i).CurrentColor.b)
        Next i
        MapData(map_x, map_y).particle_group = particle_group_index
    End With
    
    particle_group_count2 = particle_group_count2 + 1
End Sub
Private Sub Particles_Update(ii As Integer, offy As Integer, offx As Integer)

    Dim i As Long
    Dim nowa As Long
    nowa = GetTickCount
    With particle_group_list2(ii)
        For i = 0 To .nParticles - 1
            If nowa - .PrtData(i).created >= .PrtData(i).lifeTime Then .PrtData(i).Status = Dead
            If .PrtData(i).Status = Alive Then
                .PrtData(i).X = .PrtData(i).X + ((.PrtData(i).vX / 5)) * particletimer
                .PrtData(i).Y = .PrtData(i).Y + ((.PrtData(i).vY / 5)) * particletimer
                .PrtData(i).vX = .PrtData(i).vX + (.XWind / 5) * particletimer
                .PrtData(i).vY = .PrtData(i).vY + (.gravity / 5) * particletimer
                D3DXColorLerp .PrtData(i).CurrentColor, .PrtData(i).StartColor, .PrtData(i).EndColor, _
                    (nowa - .PrtData(i).created) / .PrtData(i).lifeTime
                .PrtVertList(i).X = .PrtData(i).X + offx
                .PrtVertList(i).Y = .PrtData(i).Y + offy
                .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.a, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.G, .PrtData(i).CurrentColor.b)
            Else
                .PrtData(i).Status = Alive
                .PrtData(i).lifeTime = .Life + (Rnd * .Life / 2)
                .PrtData(i).created = nowa
                .PrtData(i).X = .X
                .PrtData(i).Y = .Y
                .PrtData(i).vX = (Rnd * .XVariation) - (.XVariation * Rnd)
                .PrtData(i).vY = (Rnd * .YVariation) - (.YVariation * Rnd)
                .PrtData(i).CurrentColor = .PrtData(i).StartColor
                .PrtVertList(i).X = .PrtData(i).X + offx
                .PrtVertList(i).Y = .PrtData(i).Y + offy
                .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.a, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.G, .PrtData(i).CurrentColor.b)
            End If
        Next i
    End With
End Sub

'********************************************************
'PARTICULAS ORE 1.0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer, _
                                        Optional ConLuz As Boolean = True)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan MartÌn Sotuyo Dodero
'Modified by Augusto JosÈ Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    
'If ConLuz = True Then
Light_Create map_x, map_y, 3
'End If
 
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            'For I = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
            '    If charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = particle_group_index Then
            '        charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = 0
            '
            '        Exit For
            '    End If
            'Next I
        End If
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub
 
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan MartÌn Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
    
End Sub
Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan MartÌn Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan MartÌn Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    If UserMinHP = 0 Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
        
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional alpha As Integer, Optional MoveX As Integer, Optional MoveY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan MartÌn Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
'On Error GoTo A:


    If no_move = False Then
    
    If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index
            
            If MoveX <> 0 Or MoveY <> 0 Then
            temp_particle.X = RandomNumber(X1, X2) - (base_tile_size / 2) + screen_x
            temp_particle.Y = RandomNumber(Y1, Y2) - (base_tile_size / 2) + screen_y
            Else

            temp_particle.X = RandomNumber(X1, X2) - (base_tile_size / 2)
            temp_particle.Y = RandomNumber(Y1, Y2) - (base_tile_size / 2)
            
            End If
            
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
            temp_particle.alpha = 255
            temp_particle.TimeAlpha = temp_particle.alive_counter * 0.5
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.angle = temp_particle.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
           
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
       
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
   
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    

    Dim Blanco(3) As Long
    Blanco(0) = RGB(255, 255, 255)
    Blanco(1) = RGB(255, 255, 255)
    Blanco(2) = RGB(255, 255, 255)
    Blanco(3) = RGB(255, 255, 255)
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.GrhIndex Then
            ' Grh_Render_Advance temp_particle.grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
             'DDrawTransGrhtoSurface temp_particle.Grh, temp_particle.X, temp_particle.Y, 1, 1, , , , , alpha_blend, , , temp_particle.angle, D3DColorARGB(temp_particle.Alpha, r, g, b)
            Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list(), alpha_blend
            Exit Sub
        End If
    End If
 
    If temp_particle.Grh.GrhIndex Then
       ' Draw_Grh temp_particle.grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, True, True, rgb_list(), alpha_blend, , , temp_particle.angle
            If (temp_particle.alpha > 0) And (temp_particle.alive_counter <= temp_particle.TimeAlpha) Then
    
    temp_particle.alpha = temp_particle.alpha - timerTicksPerFrame * 15
    
    End If
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list(), alpha_blend
        
        'Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(),  True, True, alpha_blend
    End If
    
a:
    
End Sub


 
'Sub CARGARMAP()
'Particle_Group_Remove_All
'End Sub
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function
Rem Mannakia .. Parituclas ORE 1.0.
 
Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = Mid$(Text, LastPos + 1)
    End If
End Function
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
Private Function Char_Check(ByVal char_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= char_last Then
        If char_list(char_index).active Then
            Char_Check = True
        End If
    End If
End Function
 Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
   
    Map_In_Bounds = True
End Function
'********************************************************
'PARTICULAS ORE 1.0

'TEXTOS CARGADOS DESDE GRH
Public Sub Engine_Font_Initialize()

Dim a As Long

Fuentes(1).Caracteres(48) = 19730 ' 0
Fuentes(1).Caracteres(49) = 19731 ' 1
Fuentes(1).Caracteres(50) = 19732 ' 2
Fuentes(1).Caracteres(51) = 19733 ' 3
Fuentes(1).Caracteres(52) = 19734 ' 4
Fuentes(1).Caracteres(53) = 19735 ' 5
Fuentes(1).Caracteres(54) = 19736 ' 6
Fuentes(1).Caracteres(55) = 19737 ' 7
Fuentes(1).Caracteres(56) = 19738 ' 8
Fuentes(1).Caracteres(57) = 19739 ' 9

For a = 0 To 25
Fuentes(1).Caracteres(a + 97) = 19779 + a
Next a

For a = 0 To 25
Fuentes(1).Caracteres(a + 65) = 19747 + a
Next a

Fuentes(1).Caracteres(32) = 19714 '
Fuentes(1).Caracteres(33) = 19715 ' !
Fuentes(1).Caracteres(34) = 19716 ' "
Fuentes(1).Caracteres(35) = 19717 ' #
Fuentes(1).Caracteres(36) = 19718 ' $
Fuentes(1).Caracteres(37) = 19719 ' %
Fuentes(1).Caracteres(38) = 19720 ' &
Fuentes(1).Caracteres(39) = 19721 ' '
Fuentes(1).Caracteres(40) = 19722 ' (
Fuentes(1).Caracteres(41) = 19723 ' )
Fuentes(1).Caracteres(42) = 19724 ' *
Fuentes(1).Caracteres(43) = 19725 ' +
Fuentes(1).Caracteres(44) = 19726 ' ,
Fuentes(1).Caracteres(45) = 19727 ' -
Fuentes(1).Caracteres(46) = 19728 ' .
Fuentes(1).Caracteres(47) = 19729 ' /
Fuentes(1).Caracteres(58) = 19740 ' :
Fuentes(1).Caracteres(59) = 19741 ' ;
Fuentes(1).Caracteres(60) = 19742 ' <
Fuentes(1).Caracteres(61) = 19743 ' =
Fuentes(1).Caracteres(62) = 19744 ' >
Fuentes(1).Caracteres(63) = 19745 ' ?
Fuentes(1).Caracteres(64) = 19746 ' @
Fuentes(1).Caracteres(91) = 19773 ' [
Fuentes(1).Caracteres(92) = 19774 ' \
Fuentes(1).Caracteres(93) = 19775 ' ]
Fuentes(1).Caracteres(94) = 19776 ' ^
Fuentes(1).Caracteres(95) = 19777 '
Fuentes(1).Caracteres(96) = 19778 ' `
Fuentes(1).Caracteres(123) = 19805 ' {
Fuentes(1).Caracteres(124) = 19806 ' |
Fuentes(1).Caracteres(125) = 19807 ' }
Fuentes(1).Caracteres(126) = 19808 ' ~
Fuentes(1).Caracteres(127) = 19809 ' 
Fuentes(1).Caracteres(63) = 19810 ' ?
Fuentes(1).Caracteres(129) = 19811 ' Å
Fuentes(1).Caracteres(63) = 19812 ' ?
Fuentes(1).Caracteres(63) = 19813 ' ?
Fuentes(1).Caracteres(63) = 19814 ' ?
Fuentes(1).Caracteres(63) = 19815 ' ?
Fuentes(1).Caracteres(63) = 19816 ' ?
Fuentes(1).Caracteres(63) = 19817 ' ?
Fuentes(1).Caracteres(63) = 19818 ' ?
Fuentes(1).Caracteres(63) = 19819 ' ?
Fuentes(1).Caracteres(63) = 19820 ' ?
Fuentes(1).Caracteres(63) = 19821 ' ?
Fuentes(1).Caracteres(63) = 19822 ' ?
Fuentes(1).Caracteres(141) = 19823 ' ç
Fuentes(1).Caracteres(63) = 19824 ' ?
Fuentes(1).Caracteres(143) = 19825 ' è
Fuentes(1).Caracteres(144) = 19826 ' ê
Fuentes(1).Caracteres(63) = 19827 ' ?
Fuentes(1).Caracteres(63) = 19828 ' ?
Fuentes(1).Caracteres(63) = 19829 ' ?
Fuentes(1).Caracteres(63) = 19830 ' ?
Fuentes(1).Caracteres(63) = 19831 ' ?
Fuentes(1).Caracteres(63) = 19832 ' ?
Fuentes(1).Caracteres(63) = 19833 ' ?
Fuentes(1).Caracteres(63) = 19834 ' ?
Fuentes(1).Caracteres(63) = 19835 ' ?
Fuentes(1).Caracteres(63) = 19836 ' ?
Fuentes(1).Caracteres(63) = 19837 ' ?
Fuentes(1).Caracteres(63) = 19838 ' ?
Fuentes(1).Caracteres(157) = 19839 ' ù
Fuentes(1).Caracteres(63) = 19840 ' ?
Fuentes(1).Caracteres(63) = 19841 ' ?
Fuentes(1).Caracteres(160) = 19842 '
Fuentes(1).Caracteres(161) = 19843 ' °
Fuentes(1).Caracteres(162) = 19844 ' ¢
Fuentes(1).Caracteres(163) = 19845 ' £
Fuentes(1).Caracteres(164) = 19846 ' §
Fuentes(1).Caracteres(165) = 19847 ' •
Fuentes(1).Caracteres(166) = 19848 ' ¶
Fuentes(1).Caracteres(167) = 19849 ' ß
Fuentes(1).Caracteres(168) = 19850 ' ®
Fuentes(1).Caracteres(169) = 19851 ' ©
Fuentes(1).Caracteres(170) = 19852 ' ™
Fuentes(1).Caracteres(171) = 19853 ' ´
Fuentes(1).Caracteres(172) = 19854 ' ¨
Fuentes(1).Caracteres(173) = 19855 '
Fuentes(1).Caracteres(174) = 19856 ' Æ
Fuentes(1).Caracteres(175) = 19857 ' Ø
Fuentes(1).Caracteres(176) = 19858 ' ∞
Fuentes(1).Caracteres(177) = 19859 ' ±
Fuentes(1).Caracteres(178) = 19860 ' ≤
Fuentes(1).Caracteres(179) = 19861 ' ≥
Fuentes(1).Caracteres(180) = 19862 ' ¥
Fuentes(1).Caracteres(181) = 19863 ' µ
Fuentes(1).Caracteres(182) = 19864 ' ∂
Fuentes(1).Caracteres(183) = 19865 ' ∑
Fuentes(1).Caracteres(184) = 19866 ' ∏
Fuentes(1).Caracteres(185) = 19867 ' π
Fuentes(1).Caracteres(186) = 19868 ' ∫
Fuentes(1).Caracteres(187) = 19869 ' ª
Fuentes(1).Caracteres(188) = 19870 ' º
Fuentes(1).Caracteres(189) = 19871 ' Ω
Fuentes(1).Caracteres(190) = 19872 ' æ
Fuentes(1).Caracteres(191) = 19873 ' ø
Fuentes(1).Caracteres(192) = 19874 ' ¿
Fuentes(1).Caracteres(193) = 19875 ' ¡
Fuentes(1).Caracteres(194) = 19876 ' ¬
Fuentes(1).Caracteres(195) = 19877 ' √
Fuentes(1).Caracteres(196) = 19878 ' ƒ
Fuentes(1).Caracteres(197) = 19879 ' ≈
Fuentes(1).Caracteres(198) = 19880 ' ∆
Fuentes(1).Caracteres(199) = 19881 ' «
Fuentes(1).Caracteres(200) = 19882 ' »
Fuentes(1).Caracteres(201) = 19883 ' …
Fuentes(1).Caracteres(202) = 19884 '  
Fuentes(1).Caracteres(203) = 19885 ' À
Fuentes(1).Caracteres(204) = 19886 ' Ã
Fuentes(1).Caracteres(205) = 19887 ' Õ
Fuentes(1).Caracteres(206) = 19888 ' Œ
Fuentes(1).Caracteres(207) = 19889 ' œ
Fuentes(1).Caracteres(208) = 19890 ' –
Fuentes(1).Caracteres(209) = 19891 ' —
Fuentes(1).Caracteres(210) = 19892 ' “
Fuentes(1).Caracteres(211) = 19893 ' ”
Fuentes(1).Caracteres(212) = 19894 ' ‘
Fuentes(1).Caracteres(213) = 19895 ' ’
Fuentes(1).Caracteres(214) = 19896 ' ÷
Fuentes(1).Caracteres(215) = 19897 ' ◊
Fuentes(1).Caracteres(216) = 19898 ' ÿ
Fuentes(1).Caracteres(217) = 19899 ' Ÿ
Fuentes(1).Caracteres(218) = 19900 ' ⁄
Fuentes(1).Caracteres(219) = 19901 ' €
Fuentes(1).Caracteres(220) = 19902 ' ‹
Fuentes(1).Caracteres(221) = 19903 ' ›
Fuentes(1).Caracteres(222) = 19904 ' ﬁ
Fuentes(1).Caracteres(223) = 19905 ' ﬂ
Fuentes(1).Caracteres(224) = 19906 ' ‡
Fuentes(1).Caracteres(225) = 19907 ' ·
Fuentes(1).Caracteres(226) = 19908 ' ‚
Fuentes(1).Caracteres(227) = 19909 ' „
Fuentes(1).Caracteres(228) = 19910 ' ‰
Fuentes(1).Caracteres(229) = 19911 ' Â
Fuentes(1).Caracteres(230) = 19912 ' Ê
Fuentes(1).Caracteres(231) = 19913 ' Á
Fuentes(1).Caracteres(232) = 19914 ' Ë
Fuentes(1).Caracteres(233) = 19915 ' È
Fuentes(1).Caracteres(234) = 19916 ' Í
Fuentes(1).Caracteres(235) = 19917 ' Î
Fuentes(1).Caracteres(236) = 19918 ' Ï
Fuentes(1).Caracteres(237) = 19919 ' Ì
Fuentes(1).Caracteres(238) = 19920 ' Ó
Fuentes(1).Caracteres(239) = 19921 ' Ô
Fuentes(1).Caracteres(240) = 19922 ' 
Fuentes(1).Caracteres(241) = 19923 ' Ò
Fuentes(1).Caracteres(242) = 19924 ' Ú
Fuentes(1).Caracteres(243) = 19925 ' Û
Fuentes(1).Caracteres(244) = 19926 ' Ù
Fuentes(1).Caracteres(245) = 19927 ' ı
Fuentes(1).Caracteres(246) = 19928 ' ˆ
Fuentes(1).Caracteres(247) = 19929 ' ˜
Fuentes(1).Caracteres(248) = 19930 ' ¯
Fuentes(1).Caracteres(249) = 19931 ' ˘
Fuentes(1).Caracteres(250) = 19932 ' ˙
Fuentes(1).Caracteres(251) = 19933 ' ˚
Fuentes(1).Caracteres(252) = 19934 ' ¸
Fuentes(1).Caracteres(253) = 19935 ' ˝
Fuentes(1).Caracteres(254) = 19936 ' ˛
Fuentes(1).Caracteres(255) = 19937 ' ˇ
End Sub

Public Sub Grh_Text_Render(ByVal Borde As Boolean, ByVal Texto As String, ByVal X As Integer, ByVal Y As Integer, ByRef color As Long, Optional ByVal alpha As Boolean = False, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh
Dim text_color(3) As Long
text_color(0) = color
text_color(1) = color
text_color(2) = color

Dim Negro(3) As Long
Negro(0) = D3DColorXRGB(0, 0, 0)
Negro(1) = D3DColorXRGB(0, 0, 0)
Negro(2) = D3DColorXRGB(0, 0, 0)
Negro(3) = D3DColorXRGB(0, 0, 0)

text_color(3) = color

If (Len(Texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
For a = 1 To Len(Texto)
b = Asc(Mid$(Texto, a, 1))
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If b <> 32 Then
If graf.GrhIndex <> 0 Then
'mega sombra O-matica
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If Borde Then
Grh_Render graf, (X + d) - 1, Y, Negro(), False, False, False
Grh_Render graf, (X + d), Y - 1, Negro(), False, False, False
End If

Grh_Render graf, (X + d), Y, text_color, False, False, alpha
'Draw_Grh graf, (x + d), y, 0, 0, text_color, Alpha, False, False
d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 2
End If
Else
d = d + 4
End If
Next a
Else
e = 0
f = 0
For a = 1 To Len(Texto)
b = Asc(Mid$(Texto, a, 1))
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If b = 32 Or b = 13 Then
If e >= 20 Then 'reemplazar por lo que os plazca
f = f + 1
e = 0
d = 0
Else
If b = 32 Then d = d + 4
End If
Else
If graf.GrhIndex > 12 Then
'mega sombra O-matica
graf.GrhIndex = Fuentes(font_index).Caracteres(b)
If Borde Then
Grh_Render graf, (X + d) - 1, Y + f * 13, Negro(), False, False, False
Grh_Render graf, (X + d), Y + f * 13 - 1, Negro(), False, False, False
End If

Grh_Render graf, (X + d), Y + f * 13, text_color, False, False, alpha
'Draw_Grh graf, (x + d), y + f * 13, 0, 0, text_color, Alpha, False, False
d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 2
End If
End If
e = e + 1
Next a
End If

End Sub

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/08/2006
'Modified by Juan MartÌn Sotuyo Dodero
'Modified by Augusto JosÈ Rando
'Added centering
'**************************************************************
Dim tile_width As Integer
Dim tile_height As Integer
Dim grh_index As Long
Dim timer_ticks_per_frame As Single
Dim base_tile_size As Integer
If Grh.GrhIndex = 0 Then Exit Sub

'Animation
If Grh.Started Then
Grh.FrameCounter = Grh.FrameCounter + (timer_ticks_per_frame * Grh.SpeedCounter)
If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
'If Grh.noloop Then
' Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
'Else
Grh.FrameCounter = 1
'End If
End If
End If

'particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
'If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
' particle_group_list(particle_group_index).frame_counter = 0
' no_move = False
'Else
' no_move = True
'End If

'Figure out what frame to draw (always 1 if not animated)
If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If grh_index <= 0 Then Exit Sub
If GrhData(grh_index).FileNum = 0 Then Exit Sub

'Modified by Augusto JosÈ Rando
'Simplier function - according to basic ORE engine
If h_centered Then
If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
End If
End If

If v_centered Then
If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
End If
End If

'Draw it to device
Device_Box_Textured_Render_Advance grh_index, _
screen_x, screen_y, _
GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
rgb_list(), _
GrhData(grh_index).sX, GrhData(grh_index).sY, _
alpha_blend

End Sub

Private Sub DDrawRenderGrhtoSurface(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef rgblist() As Long, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte, Optional ByVal angle As Single)
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                           
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                           
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If
 
If map_x Or map_y = 0 Then map_x = 1: map_y = 1
 
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        rgblist(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, angle
 
End Sub

Public Sub Engine_ActFPS()
 
If mode = True Then
        TechoDesv.AlphaX = TechoDesv.AlphaX + 1
        If TechoDesv.AlphaX > 50 And TechoDesv.AlphaX < 60 Then
            TechoDesv.AlphaX = 50
            mode = False
        End If
    Else
        TechoDesv.AlphaX = TechoDesv.AlphaX - 1
        If TechoDesv.AlphaX < 10 And TechoDesv.AlphaX > 5 Then
            TechoDesv.AlphaX = 5
            mode = True
        End If
    End If
 
    If bTecho Then
        If Not Val(AlphaY) = 10 Then AlphaY = Val(AlphaY) - 1
    Else
        If Not AlphaY = 50 Then AlphaY = AlphaY + 1
    End If
   
    temp_rgb(0) = D3DColorARGB(AlphaY, AlphaY, AlphaY, AlphaY)
    temp_rgb(1) = D3DColorARGB(AlphaY, AlphaY, AlphaY, AlphaY)
    temp_rgb(2) = D3DColorARGB(AlphaY, AlphaY, AlphaY, AlphaY)
    temp_rgb(3) = D3DColorARGB(AlphaY, AlphaY, AlphaY, AlphaY)
 
    LuzGrh(0) = D3DColorARGB(TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX)
    LuzGrh(1) = D3DColorARGB(TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX)
    LuzGrh(2) = D3DColorARGB(TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX)
    LuzGrh(3) = D3DColorARGB(TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX, TechoDesv.AlphaX)
 
End Sub

Sub Engine_Weather_UpdateFog()
'*****************************************************************
'Update the fog effects
'*****************************************************************
Dim TempGrh As Grh
Dim i As Long
Dim X As Long
Dim Y As Long
Dim cc(3) As Long
Dim ElapsedTime As Single
ElapsedTime = Engine_ElapsedTime
 
    If WeatherFogCount = 0 Then WeatherFogCount = 13
 
    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
   
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop
    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop
    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop
    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
   
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop
    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop
    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop
    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop
 
    TempGrh.FrameCounter = 1
   
    'Render fog 2
    TempGrh.GrhIndex = 18102
    X = 2
    Y = -1
 
    cc(1) = D3DColorARGB(75, 50, 50, 50)
    cc(2) = D3DColorARGB(75, 50, 50, 50)
    cc(3) = D3DColorARGB(75, 50, 50, 50)
    cc(0) = D3DColorARGB(75, 50, 50, 50)
    For i = 1 To WeatherFogCount
        Draw_Niebla TempGrh, (X * 512) + WeatherFogX2, (Y * 512) + WeatherFogY2, 1, 1
        X = X + 1
        If X > (1 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
           
    'Render fog 1
    TempGrh.GrhIndex = 18101
    X = 0
    Y = 0
    cc(1) = D3DColorARGB(75, 50, 50, 50)
    cc(2) = D3DColorARGB(75, 50, 50, 50)
    cc(3) = D3DColorARGB(75, 50, 50, 50)
    cc(0) = D3DColorARGB(75, 50, 50, 50)
    For i = 1 To WeatherFogCount
        Draw_Niebla TempGrh, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1, 1
        X = X + 1
        If X > (2 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
 
End Sub
 
Function Engine_PixelPosX(ByVal X As Integer) As Integer
    Engine_PixelPosX = (X - 1) * 32
End Function
 
Function Engine_PixelPosY(ByVal Y As Integer) As Integer
    Engine_PixelPosY = (Y - 1) * 32
End Function
 
Private Function Engine_ElapsedTime() As Long
Dim start_time As Long
    start_time = GetTickCount
    Engine_ElapsedTime = start_time - EndTime
    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    EndTime = start_time
End Function

Sub Draw_Niebla(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte, Optional ByVal angle As Single)
 
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                           
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                           
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If
 
If map_x Or map_y = 0 Then map_x = 1: map_y = 1
 
Dim cc(3) As Long
    cc(1) = D3DColorARGB(75, 50, 50, 50)
    cc(2) = D3DColorARGB(75, 50, 50, 50)
    cc(3) = D3DColorARGB(75, 50, 50, 50)
    cc(0) = D3DColorARGB(75, 50, 50, 50)
 
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        cc(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        20, angle
 
End Sub

