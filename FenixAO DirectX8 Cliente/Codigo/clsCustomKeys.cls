VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08/03/07
' Created and Implemented by Rapsodius
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Customizes User Keys.
'
'@author Rapsodius
'@version 1.1.0
'@date 20080413

'03/08/2007 - Rapsodius
'   - First Release
'
'13/04/2008 - Juan Mart�n Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Removed lots of unwanted buggy features keeping it clean. Revised documentation and coding standards.


Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 30

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\init\BindKeys.bin"

''
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyToggleCombatMode Key for toggling combat mode
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyToggleFPS Key for toggling the FPS counter
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKeyUp = 1 '1
    mKeyDown '2
    mKeyLeft '3
    mKeyRight '4
    
    mKeyToggleMusic '5
    mKeyRequestRefresh '6
    mKeyToggleNames '7
    mKeyGetObject '8
    mKeyEquipObject '9
    mKeyTamAnimal '10
    mKeySteal '11
    mKeyHide '12
    mKeyDropObject '13
    mKeyUseObject '14
    mKeyAttack '15
    mKeyTalk '16
    mKeyTakeScreenShot '17
    mKeyShowOptions '18
    mKeyMeditate '19
    mKeyExitGame '20
    mKeyUnlock '21
    mKeyParty '22
    mKeyInvi '23
    mKeyNormal
    mKeySusurrar
    mKeyClan
    mKeyGrito
    mKeyRol
    mKeyParti
    mKeyGlobal
End Enum

''
'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid

Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Enter"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
            
        Case KeyCodeConstants.vbKey1
            ReadableName = "1 Normal"
            
        Case KeyCodeConstants.vbKey2
            ReadableName = "2 Susurrar"

        Case KeyCodeConstants.vbKey3
            ReadableName = "3 Clan"
           
        Case KeyCodeConstants.vbKey4
            ReadableName = "4 Grito"

        Case KeyCodeConstants.vbKey5
            ReadableName = "5 Rol"

        Case KeyCodeConstants.vbKey6
            ReadableName = "6 Party"
            
        Case KeyCodeConstants.vbKey8
            ReadableName = "8 Global"
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

''
'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.

Public Sub LoadCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/21/07
'
'***************************************************
    Dim i As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    'Inits keyboard config with default configurations
    Call LoadDefaults
        
    Open App.Path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Get FileNum, , mappedKeys(i)
        If LenB(ReadableName(mappedKeys(i))) = 0 Then
            Close FileNum
            Call Err.Raise(ERROR_NO_KEY_FILE)
            Exit Sub
        End If
    Next i
    
    Close FileNum
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Dim FileNum As Integer
    Dim i As Long
    FileNum = FreeFile()
    
    Open App.Path & KEY_CONFIG_FILE For Binary Access Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Put FileNum, , mappedKeys(i)
    Next i
    
    Close #FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    BindedKey = mappedKeys(Index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 08/05/07
'
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(Index) = NewVal
End Property

''
'Loads default keys

Public Sub LoadDefaults()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM '5
    mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL '6
    mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyF12 '7
    mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA '8
    mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE '9
    mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD '10
    mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR '11
    mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO '12
    mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT '13
    mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU '14
    mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl '15
    mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn '16
    mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF5 '17 'X
    mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5 '18
    mappedKeys(eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF7 '19 'X
    mappedKeys(eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyEscape '20
    mappedKeys(eKeyType.mKeyUnlock) = KeyCodeConstants.vbKeyZ '21 'X
    mappedKeys(eKeyType.mKeyParty) = KeyCodeConstants.vbKeyF9 '22 'X
    mappedKeys(eKeyType.mKeyInvi) = KeyCodeConstants.vbKeyF3 '23 'X
    mappedKeys(eKeyType.mKeyNormal) = KeyCodeConstants.vbKey1 '24
    mappedKeys(eKeyType.mKeySusurrar) = KeyCodeConstants.vbKey2 '25
    mappedKeys(eKeyType.mKeyClan) = KeyCodeConstants.vbKey3 '26
    mappedKeys(eKeyType.mKeyGrito) = KeyCodeConstants.vbKey4 '27
    mappedKeys(eKeyType.mKeyRol) = KeyCodeConstants.vbKey5 '28
    mappedKeys(eKeyType.mKeyParti) = KeyCodeConstants.vbKey6 '29
    mappedKeys(eKeyType.mKeyGlobal) = KeyCodeConstants.vbKey8 '29
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get Count() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Count = NUM_CUSTOM_KEYS
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
On Error GoTo ErrHand:
    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.Number = ERROR_NO_KEY_FILE Then
        Call LoadDefaults
        Call SaveCustomKeys
        Resume Next
    End If
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Call SaveCustomKeys
End Sub
