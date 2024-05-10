VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   4875
   ClientTop       =   2085
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10920
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   24
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10440
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   23
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9960
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   22
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9480
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   21
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9000
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   20
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Duelo BOT"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   8400
      Width           =   1215
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   6840
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   13
      Top             =   150
      Width           =   1545
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   315
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   529
      TabIndex        =   12
      Top             =   2100
      Width           =   7935
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   1200
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   600
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   11880
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   5400
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   4800
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   3600
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3120
      Top             =   2520
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6960
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   7935
      Left            =   8535
      ScaleHeight     =   529
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   1
      Top             =   360
      Width           =   3270
      Begin VB.CommandButton Command1 
         Caption         =   "PARTÍCULAS Y LUCES"
         Height          =   495
         Left            =   3840
         TabIndex        =   10
         Top             =   6360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MOTION BLUR"
         Height          =   495
         Left            =   5280
         TabIndex        =   9
         Top             =   6360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Text            =   "2"
         Top             =   6120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox hlst 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   420
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label LblNpcMana 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label LblNpcHp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   6360
         Width           =   1395
      End
      Begin VB.Label ManaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   7080
         Width           =   1515
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   7440
         Width           =   1515
      End
      Begin VB.Image Hpshp 
         Height          =   300
         Left            =   240
         Picture         =   "frmMain.frx":E711
         Top             =   7440
         Width           =   2805
      End
      Begin VB.Image MANShp 
         Height          =   300
         Left            =   240
         Picture         =   "frmMain.frx":11026
         Top             =   7080
         Width           =   2805
      End
      Begin VB.Image CmdLanzar 
         Height          =   645
         Left            =   330
         MouseIcon       =   "frmMain.frx":13A8E
         MousePointer    =   99  'Custom
         Top             =   4830
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "fdssdfsdfsdf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1800
         MouseIcon       =   "frmMain.frx":13BE0
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1290
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         MouseIcon       =   "frmMain.frx":13D32
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1305
         Width           =   1605
      End
      Begin VB.Image InvEqu 
         Height          =   4395
         Left            =   75
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   2
         Top             =   1965
         Width           =   30
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1515
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   150
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   2672
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":13E84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LvlLbl 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9120
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10080
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11040
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Image PicResu 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   9960
      Picture         =   "frmMain.frx":13F02
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image PicAU 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   9600
      Picture         =   "frmMain.frx":15204
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image PicMH 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   9600
      Picture         =   "frmMain.frx":16476
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image PicSeg 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   9255
      Picture         =   "frmMain.frx":17288
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   6240
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   8190
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NPCsemiparalizado As Boolean
Dim TiraSpell As Integer
Dim NPCtype As Integer
Public NPCparalizado As Boolean
Public tx As Byte
Public ty As Byte
Public BotDisponible As Boolean

Public NPCcount As Integer
Public NPCcont As Integer


Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub

Public Sub DibujarSatelite()
PicAU.Visible = True
End Sub

Public Sub DesDibujarSatelite()
PicAU.Visible = False
End Sub

Private Sub Command2_Click()
engine.Engine_Blur_Toggle
End Sub

Private Sub Command3_Click()

If BotDisponible = False Then
NPCtype = 1
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
BotDisponible = True
NPCparalizado = False
NPChp = 350
NPCmana = 3000
frmMain.LblNpcMana.Caption = NPCmana & "/" & NPCmanamax
frmMain.LblNpcHp.Caption = NPChp & "/" & NPChpMax
UserParalizado = False
MIhp = MihpMax
MImana = MimanaMax
Call Actualiza
End If

End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label3_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Picture2_Click()
UsarItemRojo
End Sub

Private Sub Picture3_Click()
UsarItemAzul
End Sub

Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Timer1_Timer()
Dim lugar As Byte
If NPCtype = 1 Then 'movimiento de mago
If NPCparalizado = False Then
'Timer1.Interval = CInt(RandomNumber(240, 350))
lugar = CInt(RandomNumber(1, 4))
If charlist(100).Pos.X - charlist(101).Pos.X > 8 Then lugar = EAST
If charlist(101).Pos.X - charlist(100).Pos.X > 8 Then lugar = WEST
If charlist(100).Pos.Y - charlist(101).Pos.Y > 5 Then lugar = SOUTH
If charlist(101).Pos.Y - charlist(100).Pos.Y > 5 Then lugar = NORTH
'9 92
'8 94
If charlist(101).Pos.X <= 9 And lugar = WEST Then lugar = EAST
If charlist(101).Pos.X >= 92 And lugar = EAST Then lugar = WEST
If charlist(101).Pos.Y <= 8 And lugar = NORTH Then lugar = SOUTH
If charlist(101).Pos.Y >= 94 And lugar = SOUTH Then lugar = NORTH

Select Case lugar
Case EAST
If MapData(charlist(101).Pos.X + 1, charlist(101).Pos.Y).CharIndex <> 0 Then Exit Sub
Case SOUTH
If MapData(charlist(101).Pos.X, charlist(101).Pos.Y + 1).CharIndex <> 0 Then Exit Sub
Case NORTH
If MapData(charlist(101).Pos.X, charlist(101).Pos.Y - 1).CharIndex <> 0 Then Exit Sub
Case WEST
If MapData(charlist(101).Pos.X - 1, charlist(101).Pos.Y).CharIndex <> 0 Then Exit Sub
End Select
'Call MoveCharbyHead(101, lugar)
engine.Char_Move_by_Head 101, lugar
Call Audio.PlayWave(SND_PASOS1)
End If
End If

If NPCtype = 2 Then 'movimiento de guerre
If charlist(100).Pos.X - charlist(101).Pos.X > 1 Then lugar = EAST
If charlist(101).Pos.X - charlist(100).Pos.X > 1 Then lugar = WEST
If charlist(100).Pos.Y - charlist(101).Pos.Y > 1 Then lugar = SOUTH
If charlist(101).Pos.Y - charlist(100).Pos.Y > 1 Then lugar = NORTH

Select Case lugar
Case EAST
If MapData(charlist(101).Pos.X + 1, charlist(101).Pos.Y).CharIndex <> 0 Then Exit Sub
Case SOUTH
If MapData(charlist(101).Pos.X, charlist(101).Pos.Y + 1).CharIndex <> 0 Then Exit Sub
Case NORTH
If MapData(charlist(101).Pos.X, charlist(101).Pos.Y - 1).CharIndex <> 0 Then Exit Sub
Case WEST
If MapData(charlist(101).Pos.X - 1, charlist(101).Pos.Y).CharIndex <> 0 Then Exit Sub
End Select
If (charlist(100).Pos.X - charlist(101).Pos.X = 0) And Abs(charlist(100).Pos.Y - charlist(101).Pos.Y) = 1 Then
    If (charlist(100).Pos.Y - charlist(101).Pos.Y) < 0 Then charlist(101).Heading = NORTH Else charlist(101).Heading = SOUTH
    Exit Sub
End If

If (charlist(100).Pos.Y - charlist(101).Pos.Y = 0) And Abs(charlist(100).Pos.X - charlist(101).Pos.X) = 1 Then
    If (charlist(100).Pos.X - charlist(101).Pos.X) < 0 Then charlist(101).Heading = WEST Else charlist(101).Heading = EAST
    Exit Sub
End If


If charlist(100).Pos.X - charlist(101).Pos.X = 1 And charlist(100).Pos.Y - charlist(101).Pos.Y = 1 Then lugar = SOUTH
If charlist(100).Pos.X - charlist(101).Pos.X = 1 And charlist(101).Pos.Y - charlist(100).Pos.Y = 1 Then lugar = NORTH
If charlist(101).Pos.X - charlist(100).Pos.X = 1 And charlist(100).Pos.Y - charlist(101).Pos.Y = 1 Then lugar = SOUTH
If charlist(101).Pos.X - charlist(100).Pos.X = 1 And charlist(101).Pos.Y - charlist(100).Pos.Y = 1 Then lugar = NORTH

If frmConnect.CheckHuellas.value = 1 Then
Call SusPasos
End If

engine.Char_Move_by_Head 101, lugar
End If

End Sub

Private Sub Timer2_Timer()
Dim act As Integer

'///////////////////////////////////////
If NPChp > NPChpMax Then
    NPChp = NPChpMax
End If
    
If NPCmana > NPCmanamax Then
    NPCmana = NPCmanamax
End If

NPCcount = NPCcount + 1
NPCcont = NPCcount
If (NPCcount >= 5) Then
NPCcount = 1
If NPCsemiparalizado = True Then
remover:
    Call Audio.PlayWave(SND_INMO)
    NPCmana = NPCmana - 300
    NPCparalizado = False
    NPCsemiparalizado = False
    Call Dialogos.CreateDialog("AN HOAX VORP", UserCharIndex + 1, D3DColorXRGB(0, 255, 255))
Else
 If (NPCparalizado = True) Then
 NPCsemiparalizado = True
 NPCcount = 2
 GoTo fina
 End If
 
 If NPCmana > 1000 Then
    If ((UserParalizado = True) Or (UserMoving = 0) Or (SameDir >= 4)) And (Abs(charlist(101).Pos.X - charlist(100).Pos.X) <= 9) And (Abs(charlist(101).Pos.Y - charlist(100).Pos.Y) <= 7) Then
        Call Audio.PlayWave(SND_APOCA)
        Call MiFX(13)
        NPCmana = NPCmana - 1000
        Call Dialogos.CreateDialog("Rahma Nañarak O'al", UserCharIndex + 1, D3DColorXRGB(0, 255, 255))
        RestarHP
        Misangre
    Else
        act = CInt(RandomNumber(0, 9))
        If ((act = 5) Or (act = 3)) And (Abs(charlist(101).Pos.X - charlist(100).Pos.X) <= 9) And (Abs(charlist(101).Pos.Y - charlist(100).Pos.Y) <= 7) Then
        Call Audio.PlayWave(SND_APOCA)
        NPCmana = NPCmana - 1000
        Call MiFX(13)
        Call Dialogos.CreateDialog("Rahma Nañarak O'al", UserCharIndex + 1, D3DColorXRGB(0, 255, 255))
        RestarHP
        Misangre
        Else
        If ((act = 2) Or (act = 4)) And (Abs(charlist(101).Pos.X - charlist(100).Pos.X) <= 9) And (Abs(charlist(101).Pos.Y - charlist(100).Pos.Y) <= 7) Then
        Call Audio.PlayWave(SND_INMO)
        Call MiFX(12)
        NPCmana = NPCmana - 300
        Call Dialogos.CreateDialog("Är Prop s'uo", UserCharIndex + 1, D3DColorXRGB(0, 200, 255))
        UserParalizado = True
        NPCcont = 0
        End If
        End If
    End If
End If
End If
End If

fina:
LblNpcMana.Caption = NPCmana & "/" & NPCmanamax
LblNpcHp.Caption = NPChp & "/" & NPChpMax
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" Then
        UsingSkill = Magia
        TiraSpell = hlst.ListIndex + 1
        frmMain.MousePointer = 2
    End If
    Call AddtoRichTextBox(frmMain.RecTxt, "Debes tirar el hechizo!", 200, 1, 1, False, True, False)
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub Form_Click()
        Call ConvertCPtoTP(MouseX, MouseY, tx, ty)

        If (MapData(tx, ty).CharIndex <> 0) Then
                        If charlist(MapData(tx, ty).CharIndex).Criminal = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Haz Clickiado al usuario " & charlist(MapData(tx, ty).CharIndex).nombre & " <CRIMINAL> ", 255, False, False, True) Else Call AddtoRichTextBox(frmMain.RecTxt, "Haz Clickiado al usuario " & charlist(MapData(tx, ty).CharIndex).nombre & " <CIUDADANO>", False, False, 255, True)
                          If UsingSkill > 0 Then
                                If UserCanAttack = 1 Then
                                UserCanAttack = 0
                                End If
                                End If
                   Else
                        If (MapData(tx, ty + 1).CharIndex <> 0) Then
                        ty = ty + 1
                            If charlist(MapData(tx, ty).CharIndex).Criminal = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Haz Clickiado al usuario " & charlist(MapData(tx, ty).CharIndex).nombre & " <CRIMINAL>", 255, False, False, True) Else Call AddtoRichTextBox(frmMain.RecTxt, "Haz Clickiado al usuario " & charlist(MapData(tx, ty).CharIndex).nombre & " <CIUDADANO>", False, False, 255, True)
                                If UsingSkill > 0 Then
                                If UserCanAttack = 1 Then
                                UserCanAttack = 0
                                End If
                                End If
                        End If
    End If
    
    frmMain.MousePointer = vbDefault
    
    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    hlst.Visible = True
    CmdLanzar.Visible = True
    
    'Creo mi inventario
    Call Grh_Render_To_Hdc(frmMain.Picture2.hdc, (542), 0, 0)
    Call Grh_Render_To_Hdc(frmMain.Picture3.hdc, (541), 0, 0)
    Call Grh_Render_To_Hdc(frmMain.Picture4.hdc, (681), 0, 0)
    Call Grh_Render_To_Hdc(frmMain.Picture5.hdc, (1018), 0, 0)
    Call Grh_Render_To_Hdc(frmMain.Picture6.hdc, (986), 0, 0)
    
    Dim RESTA As Integer
    
    Select Case TiraSpell

    Case 1 'tormenta
    If BotDisponible = True Then
    If MImana >= 150 Then
    If (MapData(tx, ty).CharIndex > 100) Then
    MImana = MImana - 150
    RESTA = CInt(RandomNumber(80, 130))
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Has lanzado Tormenta de fuego sobre " & charlist(MapData(tx, ty).CharIndex).nombre & "!!", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le haz quitado " & RESTA & " puntos de vida!!", 255, 0, 0, True, False, False)
    Call Audio.PlayWave(SND_APOCA)
    Call SuFX(7)
    ManaBar.Caption = MImana & "/" & MimanaMax
    MANShp.Width = (((MImana + 1 / 100) / (MimanaMax + 1 / 100)) * 94)
    QuitarvidaNPC RESTA
    MANShp.Width = (((MImana + 1 / 100) / (MimanaMax + 1 / 100)) * 94)
    Call Dialogos.CreateDialog("EN VAX ON TAR", UserCharIndex, D3DColorXRGB(0, 255, 255))
    SuSangre
    Call Dialogos.CreateDialog("-" & RESTA, UserCharIndex + 1, D3DColorXRGB(255, 0, 0))
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Target Invalido!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente mana!", False, 255, False, False)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡No puedes atacar a tu enemigo, esta descansando!!", 255, 0, 0, True, False, False)
    End If
    
    Case 2 'Descarga!
    If BotDisponible = True Then
    If MImana >= 460 Then
    If (MapData(tx, ty).CharIndex > 100) Then
    MImana = MImana - 460
    RESTA = CInt(RandomNumber(140, 190))
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Has lanzado Descarga Electrica sobre " & charlist(MapData(tx, ty).CharIndex).nombre & "!!", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le haz quitado " & RESTA & " puntos de vida!!", 255, 0, 0, True, False, False)
    QuitarvidaNPC RESTA
    Call Dialogos.CreateDialog("T'HY KOOOL", UserCharIndex, D3DColorXRGB(0, 255, 255))
    Call Audio.PlayWave(SND_DESC)
    Call SuFX(11)
    SuSangre
    Actualiza
    Call Dialogos.CreateDialog("-" & RESTA, UserCharIndex + 1, D3DColorXRGB(255, 0, 0))
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Target Invalido!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente mana!", False, 255, False, False)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡No puedes atacar a tu enemigo, esta descansando!!", 255, 0, 0, True, False, False)
    End If
    
    Case 3 'Apoca!
    If BotDisponible = True Then
    If MImana >= 1000 Then
    If (MapData(tx, ty).CharIndex > 100) Then
    MImana = MImana - 1000
    RESTA = CInt(RandomNumber(190, 220))
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Has lanzado Apocalipsis sobre " & charlist(MapData(tx, ty).CharIndex).nombre & "!!", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le haz quitado " & RESTA & " puntos de vida!!", 255, 0, 0, True, False, False)
    Call Audio.PlayWave(SND_APOCA)
    QuitarvidaNPC RESTA
    Call Dialogos.CreateDialog("Rahma Nañarak O'al", UserCharIndex, D3DColorXRGB(0, 255, 255))
    Actualiza
    SuSangre
    Call SuFX(13)
    Call Dialogos.CreateDialog("-" & RESTA, UserCharIndex + 1, D3DColorXRGB(255, 0, 0))
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Target Invalido!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente mana!", False, 255, False, False)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡No puedes atacar a tu enemigo, esta descansando!!", 255, 0, 0, True, False, False)
    End If
    
    Case 4 'Inmo!
    If BotDisponible = True Then
    If MImana >= 300 Then
    If (MapData(tx, ty).CharIndex > 100) Then
    MImana = MImana - 300
    Call Audio.PlayWave(SND_INMO)
    Call Dialogos.CreateDialog("Är Prop s'uo", UserCharIndex, D3DColorXRGB(0, 255, 255))
    NPCparalizado = True
    Call SuFX(12)
    Actualiza
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Target Invalido!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente mana!", False, 255, False, False)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡No puedes atacar a tu enemigo, esta descansando!!", 255, 0, 0, True, False, False)
    End If
    
    Case 5 'Remo!
    If MImana >= 300 Then
    If (MapData(tx, ty).CharIndex = 100) Then
    If UserParalizado = True Then
    MImana = MImana - 300
    Call Audio.PlayWave(SND_INMO)
    Call Dialogos.CreateDialog("AN HOAX VORP", UserCharIndex, D3DColorXRGB(0, 255, 255))
    UserParalizado = False
    Actualiza
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No te encuentras inmovilizado!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "Target Invalido!", False, 255, 50, 50)
    End If
    Else
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente mana!", False, 255, False, False)
    End If
    End Select

    TiraSpell = 0
End Sub

Private Sub Form_Load()
            
   Me.Left = 0
   Me.Top = 0
   
   Call AddtoRichTextBox(frmMain.RecTxt, "Bienvenido al servidor de agite desarrollado por Eduardo J. Moreno", 255, 255, 255, 255)
   Call AddtoRichTextBox(frmMain.RecTxt, "Gracias por jugar... Para mas info www.Google.com", 200, 200, 200, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tx >= MinXBorder And ty >= MinYBorder And _
    ty <= MaxYBorder And tx <= MaxXBorder Then
    If MapData(tx, ty).CharIndex > 0 Then
        If charlist(MapData(tx, ty).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tx, ty).CharIndex).nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tx, ty).CharIndex).nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + (32 * 17) Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + (32 * 13) Then Exit Function
    
    InGameArea = True
End Function



Private Sub UsarItemRojo()
    
    Call Audio.PlayWave(SND_ROJA)
    
    If MIhp > MihpMax Then
    MIhp = MihpMax
    End If
    
    If MIhp < MihpMax Then
    MIhp = MIhp + 30
    End If
    
    HpBar.Refresh
    Hpshp.Refresh
    
    HpBar.Caption = MIhp & "/" & MihpMax
    Hpshp.Width = (((MIhp + 1 / 100) / (MihpMax + 1 / 100)) * 94)
End Sub


Private Sub UsarItemAzul()
    
    Call Audio.PlayWave(SND_ROJA)
    
    If MImana > MimanaMax Then
    MImana = MimanaMax
    End If
    
    If MImana < 0 Then
    MImana = 0
    End If
    
    If MImana < MimanaMax Then
    MImana = MImana + 110
    End If
    
    ManaBar.Refresh
    MANShp.Refresh
    
    ManaBar.Caption = MImana & "/" & MimanaMax
    MANShp.Width = (((MImana + 1 / 100) / (MimanaMax + 1 / 100)) * 94)
End Sub


Sub RestarHP()

Dim RESTAr As Integer
RESTAr = CInt(RandomNumber(195, 230))
MIhp = MIhp - RESTAr

If MIhp <= 0 Then
MIhp = 0
Call AddtoRichTextBox(frmMain.RecTxt, "Haz muerto...", 255, 0, 0, True, False, False)
frmMain.HpBar.Caption = MIhp & "/" & MihpMax
Hpshp.Width = (((MIhp + 1 / 100) / (MihpMax + 1 / 100)) * 94)
Call Audio.PlayWave(SND_MUERE)
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Call Dialogos.CreateDialog("Eh ganado el duelo!", UserCharIndex + 1, D3DColorXRGB(255, 255, 255))
Call YoGanador
BotDisponible = False
Else
Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & charlist(101).nombre & " ha lanzado Apocalipsis sobre ti.!!", 255, 0, 0, True, False, False)
Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & charlist(101).nombre & " te ha quitado " & RESTAr & " puntos de vida!!", 255, 0, 0, True, False, False)
Call Dialogos.CreateDialog("-" & RESTAr, UserCharIndex, D3DColorXRGB(255, 0, 0))
frmMain.HpBar.Caption = MIhp & "/" & MihpMax
Hpshp.Width = (((MIhp + 1 / 100) / (MihpMax + 1 / 100)) * 94)
End If

End Sub


Sub QuitarvidaNPC(Vida As Integer)
NPChp = NPChp - Vida
If NPChp <= 0 Then
NPChp = 0
Call AddtoRichTextBox(frmMain.RecTxt, "¡¡ Haz matado a " & charlist(101).nombre & " !!", 255, 0, 0, True, False, False)
Call Audio.PlayWave(SND_MUERE)
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Call YoGanador
Call Dialogos.CreateDialog("Me haz derrotado u.u!", UserCharIndex + 1, D3DColorXRGB(255, 255, 255))
LblNpcHp.Caption = NPChp & "/" & NPChpMax
LblNpcHp.Refresh
BotDisponible = False
Else
LblNpcHp.Caption = NPChp & "/" & NPChpMax
LblNpcHp.Refresh
End If
End Sub


Sub Actualiza()
ManaBar.Caption = MImana & "/" & MimanaMax
HpBar.Caption = MIhp & "/" & MihpMax
Hpshp.Width = (((MIhp + 1 / 100) / (MihpMax + 1 / 100)) * 94)
MANShp.Width = (((MImana + 1 / 100) / (MimanaMax + 1 / 100)) * 94)
End Sub

Private Sub Timer3_Timer()

If NPChp > NPChpMax Then
    NPChp = NPChpMax
End If
If NPCmana > NPCmanamax Then
    NPCmana = NPCmanamax
End If

If NPChp < NPChpMax Then
    Call Audio.PlayWave(SND_ROJA)
    NPChp = NPChp + 40
End If

If NPCmana < NPCmanamax Then
    Call Audio.PlayWave(SND_ROJA)
    NPCmana = NPCmana + 110
End If

LblNpcMana.Caption = NPCmana & "/" & NPCmanamax
LblNpcHp.Caption = NPChp & "/" & NPChpMax
End Sub


Sub Misangre()
With charlist(UserCharIndex).Pos
MapData(.X, .Y).Blood.Active = 1
MapData(.X, .Y).Blood.Grh = 35
MapData(.X, .Y).Blood.LifeTime = engine.FPS * 10
MapData(.X, .Y).Blood.Alpha = 200
MapData(.X, .Y).Blood.Head = charlist(UserCharIndex).Heading
End With
End Sub

Sub SuSangre()
With charlist(UserCharIndex + 1).Pos
MapData(.X, .Y).Blood.Active = 1
MapData(.X, .Y).Blood.Grh = 35
MapData(.X, .Y).Blood.LifeTime = engine.FPS * 10
MapData(.X, .Y).Blood.Alpha = 200
MapData(.X, .Y).Blood.Head = charlist(UserCharIndex + 1).Heading
End With
End Sub

Sub MiFX(Cual As Byte)
charlist(UserCharIndex).FxIndex = Cual
charlist(UserCharIndex).fX.Loops = 0
Call SetCharacterFx(UserCharIndex, charlist(UserCharIndex).FxIndex, charlist(UserCharIndex).fX.Loops)
End Sub

Sub SuFX(Cual As Byte)
charlist(UserCharIndex + 1).FxIndex = Cual
charlist(UserCharIndex + 1).fX.Loops = 0
Call SetCharacterFx(UserCharIndex + 1, charlist(UserCharIndex + 1).FxIndex, charlist(UserCharIndex + 1).fX.Loops)
End Sub
