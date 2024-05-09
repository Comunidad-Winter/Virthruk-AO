VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LOGO 
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8475
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   2340
      Left            =   2460
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3045
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   4128
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":830D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Private Sub command1_Click()

ddsd4.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set ddsAlphaPicture = DirectDraw.CreateSurfaceFromFile("C:\Windows\Escritorio\Noche.bmp", ddsd4)

End Sub
Private Sub Form_Load()

LOGO.Picture = LoadPicture(App.Path & "\graficos\Cargando.gif")

End Sub

