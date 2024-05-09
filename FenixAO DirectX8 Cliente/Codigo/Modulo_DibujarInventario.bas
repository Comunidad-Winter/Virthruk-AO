Attribute VB_Name = "modInventarioGrafico"
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester


Option Explicit

Public Const XCantItems = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Private bStaticInit  As Boolean
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer
Sub ActualizarOtherInventory(Slot As Integer)

If OtherInventory(Slot).OBJIndex = 0 Then
    frmComerciar.List1(0).List(Slot - 1) = "Nada"
Else
    frmComerciar.List1(0).List(Slot - 1) = OtherInventory(Slot).Name
End If

If frmComerciar.List1(0).ListIndex = Slot - 1 And lista = 0 Then Call ActualizarInformacionComercio(0)

End Sub
Sub ActualizarInventario(Slot As Integer)
Dim OBJIndex As Long
Dim NameSize As Byte

If UserInventory(Slot).Amount = 0 Then
    frmMain.imgObjeto(Slot).ToolTipText = "Nada"
    frmMain.lblObjCant(Slot).ToolTipText = "Nada"
    frmMain.lblObjCant(Slot).Caption = ""
    If ItemElegido = Slot Then frmMain.Shape1.Visible = False
Else
    frmMain.imgObjeto(Slot).ToolTipText = UserInventory(Slot).Name
    frmMain.lblObjCant(Slot).ToolTipText = UserInventory(Slot).Name
    frmMain.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)
    If ItemElegido = Slot Then frmMain.Shape1.Visible = True
End If

If UserInventory(Slot).GrhIndex > 0 Then
    frmMain.imgObjeto(Slot).Picture = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp")
Else
    frmMain.imgObjeto(Slot).Picture = LoadPicture()
End If

If UserInventory(Slot).Equipped > 0 Then
    frmMain.Label2(Slot).Visible = True
Else
    frmMain.Label2(Slot).Visible = False
End If

If frmComerciar.Visible Then
    If UserInventory(Slot).Amount = 0 Then
        frmComerciar.List1(1).List(Slot - 1) = "Nada"
     Else
        frmComerciar.List1(1).List(Slot - 1) = UserInventory(Slot).Name
    End If
    If frmComerciar.List1(1).ListIndex = Slot - 1 And lista = 1 Then Call ActualizarInformacionComercio(1)
End If

End Sub
