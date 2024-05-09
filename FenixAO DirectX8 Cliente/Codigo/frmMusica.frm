VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMusica 
   BorderStyle     =   0  'None
   Caption         =   "frmMusica"
   ClientHeight    =   90
   ClientLeft      =   1365
   ClientTop       =   16305
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   960
   End
   Begin WMPLibCtl.WindowsMediaPlayer Musica 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6376
      _cy             =   1508
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Musica.Controls.Play
End Sub
