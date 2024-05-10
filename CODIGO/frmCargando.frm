VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox MP3Files 
      Height          =   285
      Left            =   360
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   840
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      _Version        =   393217
      BackColor       =   16512
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCargando.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8865
      Left            =   0
      Picture         =   "frmCargando.frx":007E
      ScaleHeight     =   591
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   536
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.Image Barra 
         Height          =   690
         Left            =   720
         Picture         =   "frmCargando.frx":585BB
         Top             =   7440
         Width           =   6780
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Barra.width = 1
End Sub

