VERSION 5.00
Begin VB.Form frmGuildsNuevo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3480
      Width           =   5295
   End
   Begin VB.ListBox MembersList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.ListBox GuildList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   600
      TabIndex        =   0
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Image command5 
      Height          =   375
      Left            =   2640
      MouseIcon       =   "GuildNuevo.frx":0000
      MousePointer    =   99  'Custom
      Top             =   7170
      Width           =   1575
   End
   Begin VB.Image command4 
      Height          =   375
      Left            =   2640
      MouseIcon       =   "GuildNuevo.frx":030A
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image command8 
      Height          =   255
      Left            =   0
      MouseIcon       =   "GuildNuevo.frx":0614
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   735
   End
End
Attribute VB_Name = "frmGuildsNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By �Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester


Public Function ListaDeClanes(ByVal Data As String) As Integer
Dim a As Integer
Dim i As Integer

a = Val(ReadField(1, Data, Asc("�")))
ReDim oClan(1 To a) As Clan

For i = 1 To a
    oClan(i).Name = Left$(ReadField(i + 1, Data, Asc("�")), Len(ReadField(i + 1, Data, Asc("�"))) - 2)
    oClan(i).Relation = Right$(ReadField(1 + i, Data, Asc("�")), 1)
Next

For i = 1 To a
    If oClan(i).Relation = 4 Then
        Call GuildList.AddItem(oClan(i).Name)
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 1 Then
        Call GuildList.AddItem(oClan(i).Name & " (A)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 2 Then
        Call GuildList.AddItem(oClan(i).Name & " (E)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 0 Then
        Call GuildList.AddItem(oClan(i).Name)
    End If
Next

ListaDeClanes = a + 2

End Function
Public Sub ParseMemberInfo(ByVal Data As String)

GuildList.Clear
MembersList.Clear
Text1 = ""

If Me.Visible Then Exit Sub

Dim a As Integer
Dim b As Integer
Dim i As Integer

b = ListaDeClanes(Data)

a = Val(ReadField(b, Data, Asc("�")))

For i = 1 To a
    Call MembersList.AddItem(ReadField(b + i, Data, Asc("�")))
Next

b = b + a + 1

Text1 = Replace(ReadField(b, Data, Asc("�")), "�", vbCrLf)

Call Me.Show(vbModeless, frmMain)
Call Me.SetFocus

End Sub
Private Sub Command4_Click()

frmCharInfo.frmmiembros = 2
Call SendData("1HRINFO<" & MembersList.List(MembersList.ListIndex))

End Sub
Private Sub command5_Click()
Dim GuildName As String


GuildName = GuildList.List(GuildList.ListIndex)
If Right$(GuildName, 1) = ")" Then GuildName = Left$(GuildName, Len(GuildName) - 4)

Call SendData("CLANDETAILS" & GuildName)

End Sub
Private Sub Command8_Click()

Me.Visible = False
frmMain.SetFocus

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "GuildMember.gif")

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
