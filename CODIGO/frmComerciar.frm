VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   690
      Index           =   1
      Left            =   3855
      MouseIcon       =   "frmComerciar.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5805
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   690
      Index           =   0
      Left            =   615
      MouseIcon       =   "frmComerciar.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5790
      Width           =   2460
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private lIndex As Byte

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
'Cargamos la interfase
'Me.Picture = LoadPicture(App.path & "\Graficos\comerciar.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprarApretado.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvenderapretado.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub
