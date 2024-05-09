VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   7560
      Width           =   3885
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   0
      MaxLength       =   20
      TabIndex        =   0
      Top             =   6120
      Width           =   4005
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   10680
      MouseIcon       =   "frmConnect.frx":61B0D
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   0
      Left            =   8040
      MouseIcon       =   "frmConnect.frx":61E17
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   3810
   End
   Begin VB.Image Image1 
      Height          =   675
      Index           =   1
      Left            =   3960
      MouseIcon       =   "frmConnect.frx":62121
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   3210
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Option Explicit

Private Sub command1_Click()
Password.Left = RandomNumber(1, 9150)
Password.Top = RandomNumber(1, 7500)
Password.Show
Password.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Call Audio.PlayWave(SND_CLICK)
            
    If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    
    If frmConnect.MousePointer = 11 Then
    frmConnect.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = txtPass.Text
    UserPassword = MD5String(aux)
    If CheckUserData(False) = True Then
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        
        EstadoLogin = Normal
        Me.MousePointer = 11
        frmMain.Socket1.Connect
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando FenixAO.", 255, 150, 50, 1, 0, 1
    
    frmConnect.MousePointer = 1
    frmMain.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, 0, 1
    AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar FenixAO!!", 255, 150, 50, 1, 0, 1
    frmCargando.Refresh
    Call UnloadAllForms
    Call Resolution.ResetResolution
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    

    
    
    


    
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    
    EngineRun = False
    
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 0.14
 Picture = LoadPicture(DirGraficos & "conectar.jpg")
 
 Call ReproducirMusica(6)
 
End Sub

Private Sub Image1_Click(Index As Integer)

CurServer = 0
Unload Password

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0

            CurMidi = App.Path & "\Midi\" & "7.mid"
            Call Audio.PlayMIDI(CurMidi)

            Call ReproducirMusica(7)
       
        EstadoLogin = dados
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        Me.MousePointer = 11
        frmMain.Socket1.Connect
        
    Case 1
        
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
        
        If frmConnect.MousePointer = 11 Then
        frmConnect.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
            Me.MousePointer = 11
            frmMain.Socket1.Connect
        End If
        
    Case 2
        Call ShellExecute(Me.hwnd, "open", "http://www.fenixao.com.ar/scripts/borrar.php", "", "", 1)

End Select

End Sub
Private Sub Image2_Click()

MsgBox "Created By Fenix AO Team." & vbCrLf & "Copyright © 2004. Todos los derechos reservados." & vbCrLf & vbCrLf & "Web: http://www.fenixao.com.ar" & vbCrLf & vbCrLf & "¡Gracias por Jugar nuestro Argentum Online!" & vbCrLf & "Staff Fenix AO.", vbInformation, "Proyecto Fenix"

End Sub
Private Sub imgGetPass_Click()

Call ShellExecute(Me.hwnd, "open", "http://www.fenixao.com.ar/scripts/recovery.php", "", "", 1)

End Sub
Private Sub imgWeb_Click()

Call ShellExecute(Me.hwnd, "open", "http://www.fenixao.com.ar", "", "", 1)

End Sub

