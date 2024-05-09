VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   Picture         =   "FrmIntro.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":F85D
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":FB67
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":FE71
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":1017B
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":10485
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   3495
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\MenuRapido.jpg")

End Sub
Private Sub Image2_Click()

Call Main

End Sub

Private Sub Image3_Click()
ShellExecute Me.hwnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
ShellExecute Me.hwnd, "open", "http://www.fenixao.com.ar/public_html/Html/manual/", "", "", 1

End Sub

Private Sub Image5_Click()
ShellExecute Me.hwnd, "open", "http://www.fenixao.com.ar", "", "", 1

End Sub

Private Sub Image6_Click()
Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      Dx3 = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then

      Move Left + (X - Dx3), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub
