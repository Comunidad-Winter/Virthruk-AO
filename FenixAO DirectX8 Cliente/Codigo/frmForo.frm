VERSION 5.00
Begin VB.Form frmForo 
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   5265
   End
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002EB7EB&
      Height          =   345
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5115
      Index           =   0
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":0000
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002EB7EB&
      Height          =   5100
      ItemData        =   "frmForo.frx":0006
      Left            =   600
      List            =   "frmForo.frx":0008
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
   Begin VB.Image command2 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmForo.frx":000A
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmForo.frx":0314
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmForo.frx":061E
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T�tulo del Tema:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmForo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By �Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Option Explicit


Public ForoIndex As Integer
Private Sub command1_Click()

Dim i
For Each i In Text
    i.Visible = False
Next

If Not MiMensaje(0).Visible Then
    List.Visible = False
    MiMensaje(0).Text = ""
MiMensaje(1).Text = ""
    MiMensaje(0).Visible = True
    MiMensaje(1).Visible = True
    MiMensaje(0).SetFocus
    command1.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
Else
    Call SendData("DEMSG" & MiMensaje(0).Text & " [" & frmMain.Label8 & "]" & Chr(176) & "Fecha: " & Date & " || Hora: " & Time & " || " & MiMensaje(1).Text)

    List.AddItem MiMensaje(0).Text & " [" & UserName & "]"
    Load Text(List.ListCount)
    Text(List.ListCount - 1).Text = "Fecha: " & Date & " || Hora: " & Time & vbCrLf & "--------------------------------------------" & vbCrLf & vbCrLf & MiMensaje(1).Text
    List.Visible = True
    
    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    command1.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

MiMensaje(0).Visible = False
MiMensaje(1).Visible = False
command1.Enabled = True
Label1.Visible = False
Label2.Visible = False
Dim i
For Each i In Text
    i.Visible = False
Next
List.Visible = True
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "ForoMensajes.gif")

End Sub

Private Sub List_Click()
List.Visible = False
Text(List.ListIndex).Visible = True

End Sub

Private Sub MiMensaje_Change(Index As Integer)
If Len(MiMensaje(0).Text) <> 0 And Len(MiMensaje(1).Text) <> 0 Then
command1.Enabled = True
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    Dx3 = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Move Left + (X - Dx3), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub mensaje_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 209) And (KeyAscii <> 241) And (KeyAscii <> 8) And (KeyAscii <> 32) And (KeyAscii <> 164) And (KeyAscii <> 165) Then
    If (KeyAscii <> 6) And ((KeyAscii < 40 Or KeyAscii > 122) Or (KeyAscii > 90 And KeyAscii < 96)) Then
        KeyAscii = 0
    End If
End If

 KeyAscii = Asc((Chr(KeyAscii)))
End Sub
