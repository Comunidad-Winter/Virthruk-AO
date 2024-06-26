Attribute VB_Name = "modDeclaraciones"
'FenixAO DirectX8
'Engine By �Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester


Option Explicit

Public Niebla As Boolean

Public mode As Boolean
Public temp_rgb(3) As Long
Public LuzGrh(3) As Long
Public AlphaY As Byte
 
Public Type tDesv
    Act As Byte
    AlphaX As Byte
End Type
 
Public TechoDesv As tDesv

Public RepMusica As Byte
Public NumReproduciendo As Integer
Public UltimaReproducida As Integer

'Particulas ORE v1.0a
Public TotalStreams As Integer
Public StreamData() As Stream
 
Public Type Stream
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
'/Particulas ORE v1.0a

Public ColorLuz As LuzClima

Type LuzClima
r As Byte
G As Byte
b As Byte
End Type

'Render
Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer
Public TileBufferSize As Integer
Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer
'/Render

Public Type tCabecera
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Public MiCabecera As tCabecera

Public CurMidi As String

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public Fide As Byte
Public Sincroniza As Single

Type tApuesta
    NumFichas As Integer
    QueSale As Integer
End Type

Public Ignorados(1 To 10) As String

Type tCasino
    FichasTotal As Byte
    NroApuestas As Byte
    Mesa As Byte
    ValorFicha As Long
    Apuesta(1 To 5) As tApuesta
End Type

Public Casino As tCasino

Type Mensajito
    Code As String
    mensaje As String
    Red As Byte
    Green As Byte
    Blue As Byte
    Bold As Byte
    Italic As Byte
End Type

Public Procesado As Boolean
Public Mensajes() As Mensajito

Type Clan
    Name As String
    Relation As Byte
End Type

Public oClan() As Clan
Public lista As Byte

Public Ayuda As Integer
Public SubAyuda As Integer
Public LastPos As Position

Public IzquierdaMapa As Integer
Public TopMapa As Integer

Public Const FONTTYPE_TALK = ",255,255,255,0,0"
Public Const FONTTYPE_FENIX = ",255,150,50,1,0"
Public Const FONTTYPE_ARMADA = ",12,0,240,1,0"
Public Const FONTTYPE_CAOS = ",222,22,22,1,1"
Public Const FONTTYPE_FIGHT = ",255,0,0,1,0"
Public Const FONTTYPE_WARNING = ",32,51,223,1,1"
Public Const FONTTYPE_INFO = ", 65, 190, 156, 0, 0"
Public Const FONTTYPE_VENENO = ",0,255,0,0,0"
Public Const FONTTYPE_GUILD = ",255,255,255,1,0"
Public Const FONTTYPE_LVLINFO = ",200,200,200,0,0"
Public Const FONTTYPE_LVLUP = ",255,150,25,1,0"
Public Const FONTTYPE_NUEVA = ",255,250,55,1,0"
Public Const FONTTYPE_MUERTO = ",255,155,55,1,0"

Public IntervaloGolpe As Single
Public IntervaloFlecha As Single
Public IntervaloSpell As Single

Public IntervaloPaso As Single
Public IntervaloUsar As Single
Public EligiendoWhispereo As Boolean

Public Golpeo As Single
Public Flecho As Single
Public Hechi As Single

Public LastHechizo As Single
Public LastGolpe As Single
Public LastFlecha As Single

Public LastPaso As Single

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6


Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public Dialogos As New cDialogos
Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public EscudosHerrero(0 To 100) As Integer
Public CascosHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer
Public ObjDruida(0 To 100) As Integer
Public ObjSastre(0 To 100) As Integer

Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public Const MAX_TIENDA_SLOTS = 10

Public NoMandoElMsg As Integer

Public Const LoopAdEternum = 999

Public Const NUMCIUDADES = 3


Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4


Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 25
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 35

Public Const NUMSKILLS = 22
Public Const NUMATRIBUTOS = 5
Public Const NUMCLASES = 58
Public Const NUMRAZAS = 5

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Const Magia = 1
Public Const Robar = 2
Public Const Tacticas = 3
Public Const Armas = 4
Public Const Meditar = 5
Public Const Apu�alar = 6
Public Const Ocultarse = 7
Public Const Supervivencia = 8
Public Const Talar = 9
Public Const Defensa = 10
Public Const Pesca = 11
Public Const Mineria = 12
Public Const Carpinteria = 13
Public Const Herreria = 14
Public Const Curacion = 15
Public Const Domar = 16
Public Const Proyectiles = 17
Public Const Wresterling = 18
Public Const Navegacion = 19
Public Const Sastreria = 20
Public Const Comerciar = 21
Public Const Resis = 22
Public Const Invita = 23

Public Const FundirMetal = 88
Public Const PescarR = 99


Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Long
    Equipped As Byte
    Valor As Long
    ObjType As Integer
    SubTipo As Byte
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxModificador As Integer
    MinModificador As Integer
    PuedeUsar As Byte
    TipoPocion As Integer
End Type

Type tReputacion
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticas
    Clase As String
    Raza As String
    VecesMurioUsuario As Long
    CiudadanosMatados As Long
    CriminalesMatados As Long
    NPCsMatados As Long
    UsuariosMatados As Long
End Type

Public ListaRazas() As String
Public ListaClases() As String

Public Nombres As Boolean

Public MostrarTextos As Boolean
Public MixedKey As Long


Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory
Public OtherInventory(1 To 40) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public codigo As String
Public UserMaxHP As Long
Public UserMinHP As Long
Public UserMaxMAN As Long
Public UserMinMAN As Long
Public UserMaxSTA As Long
Public UserMinSTA As Long
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserMontando As Boolean
Public UserEstado As Byte
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticas
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public Pausa As Boolean
Public ModoTrabajo As Boolean
Public UserParalizado As Boolean
Public CONGELADO As Boolean
Public UserNavegando As Boolean


Public Comerciando As Byte


Public UserHogar As Byte
Public UserSexo As Integer
Public UserRaza As Byte
Public UserEmail As String

Public UserSkills() As Integer
Public SkillsNames() As String
Public MiClase As Integer
Public UserAtributos() As Integer
Public AtributosNames() As String

Public Ciudades() As String
Public CityDesc() As String

Public Musica As Byte
Public FX As Byte

Public SkillPoints As Integer
Public Alocados As Integer
Public FLAGS() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

Public UsingSkill As Integer

Public Const CIUDADANO = 1
Public Const TRABAJADOR = 2
Public Const EXPERTO_MINERALES = 3
Public Const MINERO = 4
Public Const HERRERO = 8
Public Const EXPERTO_MADERA = 13
Public Const TALADOR = 14
Public Const CARPINTERO = 18
Public Const PESCADOR = 23
Public Const SASTRE = 27
Public Const ALQUIMISTA = 31
Public Const LUCHADOR = 35
Public Const CON_MANA = 36
Public Const HECHICERO = 37
Public Const MAGO = 38
Public Const NIGROMANTE = 39
Public Const ORDEN = 40
Public Const PALADIN = 41
Public Const CLERIGO = 42
Public Const NATURALISTA = 43
Public Const BARDO = 44
Public Const DRUIDA = 45
Public Const SIGILOSO = 46
Public Const ASESINO = 47
Public Const CAZADOR = 48
Public Const SIN_MANA = 49
Public Const ARQUERO = 50
Public Const GUERRERO = 51
Public Const CABALLERO = 52
Public Const BANDIDO = 53
Public Const PIRATA = 55
Public Const LADRON = 56

Public Enum E_MODO
    Normal = 1
    BorrarPJ = 2
    CrearNuevoPj = 3
    dados = 4
    RecuperarPAss = 5
    Activar = 6
End Enum


Public EstadoLogin As E_MODO

Public stxtbuffer As String
Public SendNewChar As Boolean
Public Connected As Boolean
Public DownloadingMap As Boolean
Public UserMap As Integer

Public ENDC As String

Public prgRun As Boolean
Public finpres As Boolean

Public IPdelServidor As String
Public PuertoDelServidor As Integer


Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Private Const LB_DIR As Long = &H18D
Private Const DDL_ARCHIVE As Long = &H20
Private Const DDL_EXCLUSIVE As Long = &H8000
Private Const DDL_FLAGS As Long = DDL_ARCHIVE Or DDL_EXCLUSIVE
 
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Dim MyPath As String


Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public bmoving As Boolean
Public Dx3 As Integer
Public dy As Integer


Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type


