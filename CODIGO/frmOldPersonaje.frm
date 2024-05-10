VERSION 5.00
Begin VB.Form frmOldPersonaje 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "frmOldPersonaje.frx":0000
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2805
      TabIndex        =   0
      Top             =   645
      Width           =   3210
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2805
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1515
      Width           =   3210
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   480
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   3780
      MouseIcon       =   "frmOldPersonaje.frx":1E020
      MousePointer    =   99  'Custom
      Top             =   2790
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   660
      MouseIcon       =   "frmOldPersonaje.frx":1E172
      MousePointer    =   99  'Custom
      Top             =   2805
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   7200
      MouseIcon       =   "frmOldPersonaje.frx":1E2C4
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   360
   End
End
Attribute VB_Name = "frmOldPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.1 MENDUZ DX8 VERSION www.noicoder.com
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Const textoKeypad = ""
Private Const textoSeguir = "Conectarse"
Private Const textoSalir = "Volver"



Private Sub Form_Load()
Dim j
For Each j In Image1()
    j.Tag = "0"
Next

NameTxt.Text = ""
PasswordTxt.Text = ""

Image1(1).Picture = LoadPicture(App.path & "\Graficos\bvolver.jpg")
Image1(0).Picture = LoadPicture(App.path & "\Graficos\bsiguiente.jpg")
Image1(2).Picture = LoadPicture(App.path & "\Graficos\bteclas.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(0).Tag = "0"
            Image1(0).Picture = LoadPicture(App.path & "\Graficos\bsiguiente.jpg")
End If
If Image1(1).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(1).Tag = "0"
            Image1(1).Picture = LoadPicture(App.path & "\Graficos\bvolver.jpg")
End If
If Image1(2).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(2).Tag = "0"
            Image1(2).Picture = LoadPicture(App.path & "\Graficos\bteclas.jpg")
End If

End Sub

Private Sub Image1_Click(index As Integer)
lblInfo.Caption = "Espere..."
Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
       
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
#End If
        
        'update user info
        UserName = NameTxt.Text
        Dim aux As String
        aux = PasswordTxt.Text
#If SeguridadAlkon Then
        UserPassword = MD5.GetMD5String(aux)
        Call MD5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            EstadoLogin = Normal
            
#If UsarWrench = 1 Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
        
    Case 1
        Me.Visible = False
    Case 2
        Load frmKeypad
        frmKeypad.Show vbModal
        Unload frmKeypad
        Me.PasswordTxt.SetFocus
        
End Select
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoSeguir
            Image1(0).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(0).Picture = LoadPicture(App.path & "\Graficos\bsiguientea.jpg")
        End If
    Case 1
        If Image1(1).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoSalir
            Image1(1).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(1).Picture = LoadPicture(App.path & "\Graficos\bvolvera.jpg")
        End If
    Case 2
        If Image1(2).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoKeypad
            Image1(2).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(2).Picture = LoadPicture(App.path & "\Graficos\bteclasa.jpg")
        End If
        
End Select
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(0)
    End If
End Sub
