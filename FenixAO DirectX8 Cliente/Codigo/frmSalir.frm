VERSION 5.00
Begin VB.Form frmSalir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cancelar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmSalir.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image aceptar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmSalir.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmSalir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester



Private Sub Aceptar_Click()

Call SendData("/SALIR")
Unload Me
Unload frmMain

End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Salir.gif")


End Sub
