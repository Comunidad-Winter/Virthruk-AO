Attribute VB_Name = "modCarga"
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester



'Parra: Modulo donde se encuentran todos los subs para cargar los recursos & info adicional

Option Explicit

Private NumBodies As Integer
Private Numheads As Integer
Private NumFxs As Integer
Private NumWeaponAnims As Integer
Private NumShieldAnims As Integer
Public Sub Setup_Ambient()

'Parra: ¿Recomendaciones para evitar este tipo de subs? Crear un archivo binario guardando el array completo.

    'Noche 87, 61, 43
    luz_dia(0).r = 87
    luz_dia(0).G = 61
    luz_dia(0).b = 43
    luz_dia(1).r = 87
    luz_dia(1).G = 61
    luz_dia(1).b = 43
    luz_dia(2).r = 87
    luz_dia(2).G = 61
    luz_dia(2).b = 43
    luz_dia(3).r = 87
    luz_dia(3).G = 61
    luz_dia(3).b = 43
    '4 am 124,117,91
    luz_dia(4).r = 124
    luz_dia(4).G = 127
    luz_dia(4).b = 91
    '5,6 am 143,137,135
    luz_dia(5).r = 143
    luz_dia(5).G = 137
    luz_dia(5).b = 135
    luz_dia(6).r = 143
    luz_dia(6).G = 137
    luz_dia(6).b = 135
    '7 am 212,205,207
    luz_dia(7).r = 212
    luz_dia(7).G = 205
    luz_dia(7).b = 207
    luz_dia(8).r = 212
    luz_dia(8).G = 205
    luz_dia(8).b = 207
    luz_dia(9).r = 212
    luz_dia(9).G = 205
    luz_dia(9).b = 207
    luz_dia(10).r = 212
    luz_dia(10).G = 205
    luz_dia(10).b = 207
    luz_dia(11).r = 212
    luz_dia(11).G = 205
    luz_dia(11).b = 207
    luz_dia(12).r = 212
    luz_dia(12).G = 205
    luz_dia(12).b = 207
    'Dia 255, 255, 255
    luz_dia(12).r = 255
    luz_dia(12).G = 255
    luz_dia(12).b = 255
    luz_dia(13).r = 255
    luz_dia(13).G = 255
    luz_dia(13).b = 255
    'Medio Dia 255, 200, 255
    luz_dia(14).r = 255
    luz_dia(14).G = 250
    luz_dia(14).b = 255
    luz_dia(15).r = 255
    luz_dia(15).G = 240
    luz_dia(15).b = 255
    luz_dia(16).r = 255
    luz_dia(16).G = 230
    luz_dia(16).b = 255
    '17/18 0, 100, 255
    luz_dia(17).r = 230
    luz_dia(17).G = 230
    luz_dia(17).b = 255
    '18/19 0, 100, 255
    luz_dia(18).r = 230
    luz_dia(18).G = 230
    luz_dia(18).b = 255
    '19/20 156, 142, 83
    luz_dia(19).r = 156
    luz_dia(19).G = 142
    luz_dia(19).b = 83
    luz_dia(20).r = 87
    luz_dia(20).G = 61
    luz_dia(20).b = 43
    luz_dia(21).r = 87
    luz_dia(21).G = 61
    luz_dia(21).b = 43
    luz_dia(22).r = 87
    luz_dia(22).G = 61
    luz_dia(22).b = 43
    luz_dia(23).r = 87
    luz_dia(23).G = 61
    luz_dia(23).b = 43
    luz_dia(24).r = 87
    luz_dia(24).G = 61
    luz_dia(24).b = 43
            
End Sub
Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Numheads


ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCascos


ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCuerpos


ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumFxs


ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

End Sub
Sub CargarArrayLluvia()
'On Error Resume Next
Dim n As Integer, i As Integer
Dim Nu As Integer

n = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Nu


ReDim bLluvia(1 To 230) As Byte

For i = 1 To 230
    Get #n, , bLluvia(i)
Next i

Close #n

End Sub
Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)

If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1

If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed

End Sub
Sub LoadGrhData()
On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer


ReDim GrhData(1 To 32000) As GrhData

Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

Get #1, , Grh

Do Until Grh <= 0
    
    
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > 32000 Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).speed
        If GrhData(Grh).speed <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If
    
    
    Get #1, , Grh

Loop


Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub
Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
'Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub
Sub CargarAnimsExtra()
Call CrearGrh(6580, 1)
Call CrearGrh(534, 2)
End Sub
Sub CargarAnimArmas()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
Next loopc

End Sub
Sub CargarAnimEscudos()
On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
Next loopc

End Sub
Sub SwitchMapNew(map As Integer)
On Error Resume Next
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer
Dim InfoTile As Byte
Dim i As Integer

Open App.Path & "\maps\Mapa" & map & ".mcl" For Binary As #1
Seek #1, 1
        

Get #1, , tempint

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        Get #1, , InfoTile
        
        MapData(X, Y).Blocked = (InfoTile And 1)
        
        Get #1, , MapData(X, Y).Graphic(1).GrhIndex
        
        For i = 2 To 4
            If InfoTile And (2 ^ (i - 1)) Then
                Get #1, , MapData(X, Y).Graphic(i).GrhIndex
                Call InitGrh(MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndex)
            Else: MapData(X, Y).Graphic(i).GrhIndex = 0
            End If
        Next
        
        MapData(X, Y).Trigger = 0
        
        For i = 4 To 6
            If (InfoTile And 2 ^ i) Then MapData(X, Y).Trigger = MapData(X, Y).Trigger Or 2 ^ (i - 4)
        Next
        
        Call InitGrh(MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex)
    
        If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
        MapData(X, Y).ObjGrh.GrhIndex = 0
    Next X
Next Y

Close #1

CurMap = map

Dim grh_list(1) As Long
Dim rgb_list(3) As Long
Particle_Group_Remove_All
Light_Remove_All
If map = 1 Then
'Fuente Grande
grh_list(1) = 16275
rgb_list(0) = D3DColorXRGB(0, 100, 255)
rgb_list(1) = D3DColorXRGB(0, 100, 255)
rgb_list(2) = D3DColorXRGB(0, 100, 255)
rgb_list(3) = D3DColorXRGB(0, 100, 255)
Particle_Group_Create 45, 46, grh_list(), rgb_list(), 50, 4, 1, -1, 0, , 0, 0, 0, -20, 20, -10, -50, 10, 50, 8, 1, 1, 2, -15, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0
End If
If map = 86 Then
'Antorcha
grh_list(1) = 16275
rgb_list(0) = D3DColorXRGB(255, 122, 0)
rgb_list(1) = D3DColorXRGB(255, 255, 255)
rgb_list(2) = D3DColorXRGB(255, 128, 0)
rgb_list(3) = D3DColorXRGB(255, 255, 0)
Particle_Group_Create 49, 48, grh_list(), rgb_list(), 12, 6, 1, -1, 0, , -5, 0, 0, 0, 0, -70, 0, 5, 10, 10, 5, 0, 0, 0, 5, 0, 1, -25, 25, 0, 0, 0, 5, 0, 0
End If


End Sub

