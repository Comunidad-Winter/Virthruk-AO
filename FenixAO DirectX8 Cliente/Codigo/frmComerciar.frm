VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox precio 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   285
      Left            =   5280
      TabIndex        =   11
      Text            =   "0"
      Top             =   6600
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   285
      Left            =   2445
      TabIndex        =   8
      Text            =   "1"
      Top             =   6600
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   735
      Width           =   480
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   1
      ItemData        =   "frmComerciar.frx":0000
      Left            =   3780
      List            =   "frmComerciar.frx":0002
      TabIndex        =   1
      Top             =   2070
      Width           =   2595
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   0
      ItemData        =   "frmComerciar.frx":0004
      Left            =   690
      List            =   "frmComerciar.frx":0006
      TabIndex        =   0
      Top             =   2070
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   6
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   60
   End
   Begin VB.Image Image2 
      Height          =   210
      Index           =   1
      Left            =   3960
      Top             =   6645
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image Image2 
      Height          =   195
      Index           =   0
      Left            =   1470
      Top             =   6645
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   975
      MouseIcon       =   "frmComerciar.frx":0008
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6120
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   10
      Top             =   1650
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image Command2 
      Height          =   270
      Left            =   6075
      MouseIcon       =   "frmComerciar.frx":0312
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7005
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   4095
      MouseIcon       =   "frmComerciar.frx":061C
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6120
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   1042
      Width           =   45
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By �Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Option Explicit
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Private Sub cantidad_Change()

If Val(cantidad.Text) < 0 Or Val(cantidad.Text) > MAX_INVENTORY_OBJS Then cantidad.Text = 1

End Sub
Private Sub cantidad_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0

End Sub
Private Sub Command2_Click()

SendData ("FINCOM")
Call Unload(Me)

End Sub
Private Sub Form_Deactivate()

Me.SetFocus

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyE Then
    If List1(1).ListIndex > -1 And List1(1).ListIndex < MAX_INVENTORY_SLOTS Then
        Call SendData("EQUI" & List1(1).ListIndex + 1)
        Call ActualizarInventario(List1(1).ListIndex + 1)
        Exit Sub
    End If
End If

End Sub

Private Sub Form_Load()

frmComerciar.Picture = LoadPicture(DirGraficos & "\comerciar.gif")
frmComerciar.Image2(0).Picture = LoadPicture(DirGraficos & "\Cantidad.gif")
frmComerciar.Image2(1).Picture = LoadPicture(DirGraficos & "\Precio.gif")

End Sub
Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then
   Picture1.Picture = Nothing
   Exit Sub
End If

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        Select Case Comerciando
            Case 1
                If UserGLD >= OtherInventory(List1(0).ListIndex + 1).Valor * Val(cantidad) Then
                    Call SendData("COMP" & List1(0).ListIndex + 1 & "," & cantidad.Text)
                Else
                    AddtoRichTextBox frmMain.rectxt, "No ten�s suficiente oro.", 2, 51, 223, 1, 1
                    Exit Sub
                End If
            Case 2
                Call SendData("RETI" & List1(0).ListIndex + 1 & "," & cantidad.Text)
            Case 3
                Call SendData("SAVE" & List1(0).ListIndex + 1 & "," & cantidad.Text)
        End Select
        If lista = 1 Then Call ActualizarInformacionComercio(0)
   Case 1
        LastIndex2 = List1(1).ListIndex
        Select Case Comerciando
            Case 1
                If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
                    Call SendData("VEND" & List1(1).ListIndex + 1 & "," & cantidad.Text)
                Else
                    AddtoRichTextBox frmMain.rectxt, "No podes vender el item porque lo est�s usando.", 2, 51, 223, 1, 1
                    Exit Sub
                End If
            Case 2
                If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
                    Call SendData("DEPO" & List1(1).ListIndex + 1 & "," & cantidad.Text)
                Else
                    AddtoRichTextBox frmMain.rectxt, "No podes depositar el item porque lo est�s usando.", 2, 51, 223, 1, 1
                    Exit Sub
                End If
            Case 3
                If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
                    If Val(precio.Text) > 0 Then
                        Call SendData("POVE" & List1(1).ListIndex + 1 & "," & cantidad.Text & "," & precio.Text)
                    Else
                        AddtoRichTextBox frmMain.rectxt, "�Debes elegir un precio de venta!", 2, 51, 223, 1, 1
                        Exit Sub
                    End If
                Else
                    AddtoRichTextBox frmMain.rectxt, "No puedes poner el item a la venta porque lo est�s usando.", 2, 51, 223, 1, 1
                    Exit Sub
                End If

        End Select
        If lista = 0 Then Call ActualizarInformacionComercio(1)
End Select

End Sub

Private Sub list1_Click(Index As Integer)

lista = Index
Call ActualizarInformacionComercio(Index)

End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
            Image1(0).Tag = 0
            Image1(1).Tag = 1
        End If
    Case 1
        If Image1(1).Tag = 1 Then
            Image1(1).Tag = 0
            Image1(0).Tag = 1
        End If
End Select

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