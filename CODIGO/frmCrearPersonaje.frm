VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   43
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   42
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox Rota2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox Rota1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   40
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox MenosHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1185
      ScaleWidth      =   690
      TabIndex        =   39
      Top             =   5400
      Width           =   720
   End
   Begin VB.PictureBox MasHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3720
      ScaleHeight     =   1185
      ScaleWidth      =   690
      TabIndex        =   38
      Top             =   5400
      Width           =   720
   End
   Begin VB.PictureBox bodyView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2265
      ScaleHeight     =   975
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   5880
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
      Left            =   1560
      ScaleHeight     =   735
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   5400
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
      Left            =   3000
      ScaleHeight     =   735
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox HeadView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2265
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   5400
      Width           =   495
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
      ItemData        =   "frmCrearPersonaje.frx":C042
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":C044
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2700
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
      ItemData        =   "frmCrearPersonaje.frx":C046
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":C050
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2250
      Width           =   2820
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
      ItemData        =   "frmCrearPersonaje.frx":C063
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":C065
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1800
      Width           =   2820
   End
   Begin VB.ComboBox lstHogar 
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
      ItemData        =   "frmCrearPersonaje.frx":C067
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":C069
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3510
      Width           =   2565
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   3570
      TabIndex        =   0
      Top             =   1275
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4020
      TabIndex        =   33
      Top             =   4260
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   32
      Top             =   8535
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":C06B
      MousePointer    =   99  'Custom
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":C1BD
      MousePointer    =   99  'Custom
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":C30F
      MousePointer    =   99  'Custom
      Top             =   3885
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":C461
      MousePointer    =   99  'Custom
      Top             =   4110
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":C5B3
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":C705
      MousePointer    =   99  'Custom
      Top             =   4545
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":C857
      MousePointer    =   99  'Custom
      Top             =   4785
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":C9A9
      MousePointer    =   99  'Custom
      Top             =   4965
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":CAFB
      MousePointer    =   99  'Custom
      Top             =   5175
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":CC4D
      MousePointer    =   99  'Custom
      Top             =   5385
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":CD9F
      MousePointer    =   99  'Custom
      Top             =   5610
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":CEF1
      MousePointer    =   99  'Custom
      Top             =   5820
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":D043
      MousePointer    =   99  'Custom
      Top             =   6015
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":D195
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":D2E7
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":D439
      MousePointer    =   99  'Custom
      Top             =   3495
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":D58B
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":D6DD
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":D82F
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":D981
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":DAD3
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":DC25
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":DD77
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":DEC9
      MousePointer    =   99  'Custom
      Top             =   5220
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E01B
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E16D
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E2BF
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E411
      MousePointer    =   99  'Custom
      Top             =   6075
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E563
      MousePointer    =   99  'Custom
      Top             =   6285
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":E6B5
      MousePointer    =   99  'Custom
      Top             =   6270
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":E807
      MousePointer    =   99  'Custom
      Top             =   6495
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":E959
      MousePointer    =   99  'Custom
      Top             =   6465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":EAAB
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":EBFD
      MousePointer    =   99  'Custom
      Top             =   6690
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":ED4F
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":EEA1
      MousePointer    =   99  'Custom
      Top             =   6915
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":EFF3
      MousePointer    =   99  'Custom
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":F145
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":F297
      MousePointer    =   99  'Custom
      Top             =   7395
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":F3E9
      MousePointer    =   99  'Custom
      Top             =   7335
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":F53B
      MousePointer    =   99  'Custom
      Top             =   7590
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7815
      MouseIcon       =   "frmCrearPersonaje.frx":F68D
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   2
      Left            =   315
      MouseIcon       =   "frmCrearPersonaje.frx":F7DF
      MousePointer    =   99  'Custom
      Top             =   3015
      Width           =   660
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   1905
      MouseIcon       =   "frmCrearPersonaje.frx":F931
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   765
   End
   Begin VB.Image boton 
      Height          =   210
      Index           =   0
      Left            =   9075
      MouseIcon       =   "frmCrearPersonaje.frx":FA83
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   1200
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7995
      TabIndex        =   28
      Top             =   7515
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7995
      TabIndex        =   27
      Top             =   7306
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7995
      TabIndex        =   26
      Top             =   7092
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7995
      TabIndex        =   25
      Top             =   6878
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7995
      TabIndex        =   24
      Top             =   6664
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7995
      TabIndex        =   23
      Top             =   6450
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7995
      TabIndex        =   22
      Top             =   6236
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7995
      TabIndex        =   21
      Top             =   6022
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7995
      TabIndex        =   20
      Top             =   5808
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7995
      TabIndex        =   19
      Top             =   5594
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7995
      TabIndex        =   18
      Top             =   5380
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7995
      TabIndex        =   17
      Top             =   5166
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7995
      TabIndex        =   16
      Top             =   4952
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7995
      TabIndex        =   15
      Top             =   4738
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7995
      TabIndex        =   14
      Top             =   4524
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7995
      TabIndex        =   13
      Top             =   4310
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7995
      TabIndex        =   12
      Top             =   4096
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7995
      TabIndex        =   11
      Top             =   3882
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7995
      TabIndex        =   10
      Top             =   3668
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7995
      TabIndex        =   9
      Top             =   3240
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7995
      TabIndex        =   8
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   8775
      Picture         =   "frmCrearPersonaje.frx":FBD5
      Top             =   3900
      Width           =   2985
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   6
      Top             =   3570
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5325
      TabIndex        =   4
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   3
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   2
      Top             =   2910
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   1
      Top             =   2580
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte

Public jeje As Integer
Public Edu1 As Integer
Public Edu2 As Integer
Public Edu3 As Integer


Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim I As Integer
For I = 1 To NUMATRIBUTOS
    If UserAtributos(I) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next I

CheckData = True


End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub


Private Sub Form_Load()
SkillPoints = 10

'If MasHead_MouseMove = True Then
'Picture = LoadPicture(App.path & "\Interfaces\SiguienteMouse.jpg")
'End Sub

'If MenosHead_MouseMove = True Then
'Picture = LoadPicture(App.path & "\Interfaces\AnteriorMouse.jpg")
'End Sub

Puntos.Caption = SkillPoints

Dim I As Integer
lstProfesion.Clear
For I = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(I)
Next I

lstHogar.Clear

For I = LBound(Ciudades()) To UBound(Ciudades())
    lstHogar.AddItem Ciudades(I)
Next I


lstRaza.Clear

For I = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(I)
Next I

lstProfesion.Clear

For I = LBound(ListaClases()) To UBound(ListaClases())
    lstProfesion.AddItem ListaClases(I)
Next I

lstProfesion.ListIndex = 1

Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")

End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
End Sub

Private Sub Picture1_Click()
Call Izquierda
End Sub

Private Sub Picture2_Click()
Call Frente
End Sub

Private Sub Rota1_Click()
Call Derecha
End Sub

Private Sub Rota2_Click()
Call Arriba
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

    'HeadViewB(1 to 4)
    Sub CabezaAnterior()
    Call DrawGrhtoHdc(HeadViewB1.hdc, HeadData(Actual - 1).Head(3).grhindex, 9, 16)
    HeadViewB1.Refresh
    End Sub
    
    Sub Derecha()
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(2).grhindex, 9, 16)
    Call DrawGrhtoHdc(bodyView.hdc, Bodydata(CuerpoDesnudo).Walk(2).grhindex, 5, -7)
    HeadView.Refresh
    bodyView.Refresh
    End Sub

    Sub Arriba()
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(1).grhindex, 9, 16)
    Call DrawGrhtoHdc(bodyView.hdc, Bodydata(CuerpoDesnudo).Walk(1).grhindex, 5, -7)
    HeadView.Refresh
    bodyView.Refresh
    End Sub
    
    Sub Izquierda()
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(4).grhindex, 9, 16)
    Call DrawGrhtoHdc(bodyView.hdc, Bodydata(CuerpoDesnudo).Walk(4).grhindex, 5, -7)
    HeadView.Refresh
    bodyView.Refresh
    End Sub

    Sub Frente()
    Call DrawGrhtoHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 9, 16)
    Call DrawGrhtoHdc(bodyView.hdc, Bodydata(CuerpoDesnudo).Walk(3).grhindex, 5, -7)
    HeadView.Refresh
    bodyView.Refresh
    End Sub
