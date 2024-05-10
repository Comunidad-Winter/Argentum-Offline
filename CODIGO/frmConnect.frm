VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox CheckHuellas 
      Caption         =   "Pasos con huellas?"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox CheckLluvia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sistema de Lluvia"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.ComboBox ComboDificulti 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Text            =   "Normal"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Text            =   "Bot"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Text            =   "Ejmo"
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   2820
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2370
      Width           =   2820
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2820
      Width           =   2820
   End
   Begin VB.PictureBox HeadView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2025
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox HeadViewA1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2760
      ScaleHeight     =   735
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox HeadViewB1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1320
      ScaleHeight     =   735
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox bodyView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2025
      ScaleHeight     =   975
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox MasHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3480
      ScaleHeight     =   1185
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   1920
      Width           =   720
   End
   Begin VB.PictureBox MenosHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1185
      ScaleWidth      =   690
      TabIndex        =   1
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona nivel de Dificultad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Enemigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese un nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   4920
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   3045
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SkillPoints As Byte
Public Actual As Integer
Public Edu1 As Integer
Public Edu2 As Integer
Public Edu3 As Integer



Private Sub Form_Activate()

lstRaza.AddItem "Humano"
lstRaza.AddItem "Elfo"
lstRaza.AddItem "Elfo Oscuro"
lstRaza.AddItem "Gnomo"
lstRaza.AddItem "Enano"

lstGenero.AddItem "Hombre"
lstGenero.AddItem "Mujer"

lstProfesion.AddItem "Mago"
lstProfesion.AddItem "Clerigo"

ComboDificulti.AddItem "Imposible"
ComboDificulti.AddItem "Dificil"
ComboDificulti.AddItem "Normal"
ComboDificulti.AddItem "Facil"

    Call DameOpciones
    Call DoyCuerpoDesnudo
End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
 '[END]'

'Recordatorio para cumplir la licencia, por si borrás el botón sin leer el code...
Dim i As Long

For i = 0 To Me.Controls.Count - 1
    If Me.Controls(i).Name = "downloadServer" Then
        Exit For
    End If
Next i

Call Audio.MusicMP3Play(App.Path & "\MP3\1.mp3")

End Sub

Private Sub Image1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

If frmConnect.Actual = 0 Then
frmConnect.Actual = 1
End If

Call MakeChar(100, 56, frmConnect.Actual, 3, 40, 40, 10, 0, 4)
    
frmMain.Show

charlist(UserCharIndex).nombre = Text1.Text
charlist(UserCharIndex + 1).nombre = Text2.Text

frmMain.Label8.Caption = Text1.Text
If ComboDificulti.Text = "Imposible" Then
frmMain.Timer2.Interval = 150
ElseIf ComboDificulti.Text = "Dificil" Then
frmMain.Timer2.Interval = 500
ElseIf ComboDificulti.Text = "Normal" Then
frmMain.Timer2.Interval = 1000
ElseIf ComboDificulti.Text = "Facil" Then
frmMain.Timer2.Interval = 1500
End If

End Sub

Private Sub Image2_Click()
End
End Sub

Private Sub lstGenero_Click()
    Call DameOpciones
    Call DoyCuerpoDesnudo
End Sub

Private Sub lstProfesion_Click()
    Call DameOpciones
    Call DoyCuerpoDesnudo
End Sub

Private Sub MasHead_Click()
    Call Audio.PlayWave(SND_CLICK)
    Actual = Actual + 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 9, 16)
    HeadView.Refresh

    'HeadViewA(1 to 4)
    Call DrawGrhtoHdc(HeadViewA1.hdc, HeadData(Actual + 1).Head(3).grhindex, 9, 16)
    HeadViewA1.Refresh
    
    If Actual >= 1 Then
    Call CabezaAnterior
    End If
    
    Call DoyCuerpoDesnudo
End Sub

    Sub CabezaAnterior()
    Call DrawGrhtoHdc(HeadViewB1.hdc, HeadData(Actual - 1).Head(3).grhindex, 9, 16)
    HeadViewB1.Refresh
    End Sub

Private Sub MenosHead_Click()
    Call Audio.PlayWave(SND_CLICK)
    Actual = Actual - 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 9, 16)
    HeadView.Refresh

    
    'HeadViewA
    Call DrawGrhtoHdc(HeadViewA1.hdc, HeadData(Actual + 1).Head(3).grhindex, 9, 16)
    HeadViewA1.Refresh
    
    If Actual >= 1 Then
    Call CabezaAnterior
    End If
    Call DoyCuerpoDesnudo
End Sub
