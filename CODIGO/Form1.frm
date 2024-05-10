VERSION 5.00
Begin VB.Form frmListaDeAmigos 
   BorderStyle     =   0  'None
   Caption         =   "Listado de amigos"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Nickveinte 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   4800
      Width           =   3015
   End
   Begin VB.OptionButton Nickdiecinueve 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   4560
      Width           =   3015
   End
   Begin VB.OptionButton Nickdieciocho 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   4320
      Width           =   3015
   End
   Begin VB.OptionButton Nickdiecisiete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   4080
      Width           =   3015
   End
   Begin VB.OptionButton Nickdieciseis 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   3840
      Width           =   3015
   End
   Begin VB.OptionButton Nickquince 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   3600
      Width           =   3015
   End
   Begin VB.OptionButton Nickcatorce 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   3360
      Width           =   3015
   End
   Begin VB.OptionButton Nicktrece 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   3120
      Width           =   3015
   End
   Begin VB.OptionButton Nickdoce 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   2880
      Width           =   3015
   End
   Begin VB.OptionButton Nickonce 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agregar amigo"
      Height          =   1335
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command7 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Nick de usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   2415
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   3495
      Begin VB.CommandButton Command5 
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mandarle MP"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Borrar amigo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar amigo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.OptionButton Nickdiez 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
      Width           =   3015
   End
   Begin VB.OptionButton Nicknueve 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2160
      Width           =   3015
   End
   Begin VB.OptionButton Nickocho 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.OptionButton Nicksiete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   1680
      Width           =   3015
   End
   Begin VB.OptionButton Nickseis 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.OptionButton Nickcinco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.OptionButton Nickcuatro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.OptionButton Nicktres 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin VB.OptionButton Nickdos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.OptionButton Nickuno 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slot Vacío"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   4800
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amigos"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmListaDeAmigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frame2.Visible = True
Text1.Text = ""
End Sub

Private Sub Command2_Click()
If Nickuno.value = True And Not Nickuno.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickuno.Caption = "Slot Vacío"
 Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK1", "Slot Vacío")
Else
End If
End If

If Nickdos.value = True And Not Nickdos.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdos.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK2", "Slot Vacío")
Else
End If
End If

If Nicktres.value = True And Not Nicktres.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nicktres.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK3", "Slot Vacío")
Else
End If
End If

If Nickcuatro.value = True And Not Nickcuatro.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickcuatro.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK4", "Slot Vacío")
Else
End If
End If

If Nickcinco.value = True And Not Nickcinco.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickcinco.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK5", "Slot Vacío")
Else
End If
End If

If Nickseis.value = True And Not Nickseis.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickseis.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK6", "Slot Vacío")
Else
End If
End If

If Nicksiete.value = True And Not Nicksiete.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nicksiete.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK7", "Slot Vacío")
Else
End If
End If

If Nickocho.value = True And Not Nickocho.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickocho.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK8", "Slot Vacío")
Else
End If
End If

If Nicknueve.value = True And Not Nicknueve.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nicknueve.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK9", "Slot Vacío")
Else
End If
End If

If Nickdiez.value = True And Not Nickdiez.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdiez.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK10", "Slot Vacío")
Else
End If
End If

If Nickonce.value = True And Not Nickonce.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickonce.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK11", "Slot Vacío")
Else
End If
End If

If Nickdoce.value = True And Not Nickdoce.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdoce.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK12", "Slot Vacío")
Else
End If
End If

If Nicktrece.value = True And Not Nicktrece.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nicktrece.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK13", "Slot Vacío")
Else
End If
End If

If Nickcatorce.value = True And Not Nickcatorce.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickcatorce.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK14", "Slot Vacío")
Else
End If
End If

If Nickquince.value = True And Not Nickquince.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickquince.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK15", "Slot Vacío")
Else
End If
End If

If Nickdieciseis.value = True And Not Nickdieciseis.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdieciseis.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK16", "Slot Vacío")
Else
End If
End If

If Nickdiecisiete.value = True And Not Nickdiecisiete.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdiecisiete.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK17", "Slot Vacío")
Else
End If
End If

If Nickdieciocho.value = True And Not Nickdieciocho.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdieciocho.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK18", "Slot Vacío")
Else
End If
End If

If Nickdiecinueve.value = True And Not Nickdiecinueve.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickdiecinueve.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK19", "Slot Vacío")
Else
End If
End If


If Nickveinte.value = True And Not Nickveinte.Caption = "Slot Vacío" Then
If MsgBox("¿Esta seguro que desea Borrarlo de la lista de amigos?", vbYesNo, "") = vbYes Then
Nickveinte.Caption = "Slot Vacío"
Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK20", "Slot Vacío")
Else
End If
End If


End Sub

Private Sub Command3_Click()
If Nickuno.value = True And Not Nickuno.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickuno.Caption & " "
Unload Me
Else
End If

If Nickdos.value = True And Not Nickdos.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdos.Caption & " "
Unload Me
Else
End If

If Nicktres.value = True And Not Nicktres.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nicktres.Caption & " "
Unload Me
Else
End If

If Nickcuatro.value = True And Not Nickcuatro.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickcuatro.Caption & " "
Unload Me
Else
End If

If Nickcinco.value = True And Not Nickcinco.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickcinco.Caption & " "
Unload Me
Else
End If

If Nickseis.value = True And Not Nickseis.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickseis.Caption & " "
Unload Me
Else
End If

If Nicksiete.value = True And Not Nicksiete.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nicksiete.Caption & " "
Unload Me
Else
End If

If Nickocho.value = True And Not Nickocho.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickocho.Caption & " "
Unload Me
Else
End If

If Nicknueve.value = True And Not Nicknueve.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nicknueve.Caption & " "
Unload Me
Else
End If

If Nickdiez.value = True And Not Nickdiez.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdiez.Caption & " "
Unload Me
Else
End If

If Nickonce.value = True And Not Nickonce.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickonce.Caption & " "
Unload Me
Else
End If

If Nickdoce.value = True And Not Nickdoce.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdoce.Caption & " "
Unload Me
Else
End If

If Nicktrece.value = True And Not Nicktrece.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nicktrece.Caption & " "
Unload Me
Else
End If

If Nickcatorce.value = True And Not Nickcatorce.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickcatorce.Caption & " "
Unload Me
Else
End If

If Nickquince.value = True And Not Nickquince.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickquince.Caption & " "
Unload Me
Else
End If

If Nickdieciseis.value = True And Not Nickdieciseis.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdieciseis.Caption & " "
Unload Me
Else
End If

If Nickdiecisiete.value = True And Not Nickdiecisiete.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdiecisiete.Caption & " "
Unload Me
Else
End If

If Nickdieciocho.value = True And Not Nickdieciocho.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdieciocho.Caption & " "
Unload Me
Else
End If

If Nickdiecinueve.value = True And Not Nickdiecinueve.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickdiecinueve.Caption & " "
Unload Me
Else
End If

If Nickveinte.value = True And Not Nickveinte.Caption = "Slot Vacío" Then
frmMain.SendTxt.Visible = True
frmMain.SendTxt.Text = "\" & Nickveinte.Caption & " "
Unload Me
Else
End If

End Sub

Private Sub Command4_Click()
Me.Visible = False
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Frame2.Visible = False
End Sub

Private Sub Command7_Click()
If Nickuno.value = True And Nickuno.Caption = "Slot Vacío" Then
    Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK1", Text1.Text)
    Nickuno.Caption = Text1.Text
    Frame2.Visible = False
    ElseIf Nickuno.value = True And Not Nickuno.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickuno.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Nickuno.Caption = Text1.Text
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK1", Text1.Text)
    Frame2.Visible = False
    Else
    End If
    End If
    
If Nickdos.value = True And Nickdos.Caption = "Slot Vacío" Then
    Nickdos.Caption = Text1.Text
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK2", Text1.Text)
    Frame2.Visible = False
    ElseIf Nickdos.value = True And Not Nickdos.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdos.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK2", Text1.Text)
    Frame2.Visible = False
    Nickdos.Caption = Text1.Text
    Else
    End If
    End If
    
If Nicktres.value = True And Nicktres.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK3", Text1.Text)
    Nicktres.Caption = Text1.Text
    ElseIf Nicktres.value = True And Not Nicktres.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nicktres.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK3", Text1.Text)
    Frame2.Visible = False
    Nicktres.Caption = Text1.Text
    Else
    End If
    End If
    
If Nickcuatro.value = True And Nickcuatro.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK4", Text1.Text)
    Nickcuatro.Caption = Text1.Text
    ElseIf Nickcuatro.value = True And Not Nickcuatro.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickcuatro.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK4", Text1.Text)
    Frame2.Visible = False
    Nickcuatro.Caption = Text1.Text
    Else
    End If
    End If
    
If Nickcinco.value = True And Nickcinco.Caption = "Slot Vacío" Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK5", Text1.Text)
    Frame2.Visible = False
    Nickcinco.Caption = Text1.Text
    ElseIf Nickcinco.value = True And Not Nickcinco.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickcinco.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK5", Text1.Text)
    Frame2.Visible = False
    Nickcinco.Caption = Text1.Text
    Else
    End If
    End If
    
If Nickseis.value = True And Nickseis.Caption = "Slot Vacío" Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK6", Text1.Text)
    Frame2.Visible = False
    Nickseis.Caption = Text1.Text
    ElseIf Nickseis.value = True And Not Nickseis.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickseis.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK6", Text1.Text)
    Frame2.Visible = False
    Nickseis.Caption = Text1.Text
    Else
    End If
    End If
    
If Nicksiete.value = True And Nicksiete.Caption = "Slot Vacío" Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK7", Text1.Text)
    Frame2.Visible = False
    Nicksiete.Caption = Text1.Text
    ElseIf Nicksiete.value = True And Not Nicksiete.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nicksiete.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK7", Text1.Text)
    Frame2.Visible = False
    Nicksiete.Caption = Text1.Text
    Else
    End If
    End If
    
If Nickocho.value = True And Nickocho.Caption = "Slot Vacío" Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK8", Text1.Text)
    Frame2.Visible = False
    Nickocho.Caption = Text1.Text
    ElseIf Nickocho.value = True And Not Nickocho.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickocho.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK8", Text1.Text)
    Frame2.Visible = False
    Nickocho.Caption = Text1.Text
    Else
    End If
    End If
    
If Nicknueve.value = True And Nicknueve.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK9", Text1.Text)
    Nicknueve.Caption = Text1.Text
    ElseIf Nicknueve.value = True And Not Nicknueve.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nicknueve.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK9", Text1.Text)
    Nicknueve.Caption = Text1.Text
    Else
    End If
    End If
    
If Nickdiez.value = True And Nickdiez.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK10", Text1.Text)
    Nickdiez.Caption = Text1.Text
    ElseIf Nicknueve.value = True And Not Nickdiez.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdiez.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK10", Text1.Text)
    Nickdiez.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickonce.value = True And Nickonce.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK11", Text1.Text)
    Nickonce.Caption = Text1.Text
    ElseIf Nickonce.value = True And Not Nickonce.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickonce.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK11", Text1.Text)
    Nickonce.Caption = Text1.Text
    Else
    End If
    End If
    
    
    If Nickdoce.value = True And Nickdoce.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK12", Text1.Text)
    Nickdoce.Caption = Text1.Text
    ElseIf Nickdoce.value = True And Not Nickdoce.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdoce.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK12", Text1.Text)
    Nickdoce.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nicktrece.value = True And Nicktrece.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK13", Text1.Text)
    Nicktrece.Caption = Text1.Text
    ElseIf Nicktrece.value = True And Not Nicktrece.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nicktrece.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK13", Text1.Text)
    Nicktrece.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickcatorce.value = True And Nickcatorce.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK14", Text1.Text)
    Nickcatorce.Caption = Text1.Text
    ElseIf Nickcatorce.value = True And Not Nickcatorce.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickcatorce.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK14", Text1.Text)
    Nickcatorce.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickquince.value = True And Nickquince.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK15", Text1.Text)
    Nickquince.Caption = Text1.Text
    ElseIf Nickquince.value = True And Not Nickquince.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickquince.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK15", Text1.Text)
    Nickquince.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickdieciseis.value = True And Nickdieciseis.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK16", Text1.Text)
    Nickdieciseis.Caption = Text1.Text
    ElseIf Nickdieciseis.value = True And Not Nickdieciseis.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdieciseis.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK16", Text1.Text)
    Nickdieciseis.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickdiecisiete.value = True And Nickdiecisiete.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK17", Text1.Text)
    Nickdiecisiete.Caption = Text1.Text
    ElseIf Nickdiecisiete.value = True And Not Nickdiecisiete.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdiecisiete.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK17", Text1.Text)
    Nickdiecisiete.Caption = Text1.Text
    Else
    End If
    End If
    
    If Nickdieciocho.value = True And Nickdieciocho.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK18", Text1.Text)
    Nickdieciocho.Caption = Text1.Text
    ElseIf Nickdieciocho.value = True And Not Nickdieciocho.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdiciocho.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK18", Text1.Text)
    Nickdieciocho.Caption = Text1.Text
    Else
    End If
    End If
    
    
    If Nickdiecinueve.value = True And Nickdiecinueve.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK19", Text1.Text)
    Nickdiecinueve.Caption = Text1.Text
    ElseIf Nickdiecinueve.value = True And Not Nickdiecinueve.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickdiecinueve.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK19", Text1.Text)
    Nickdiecinueve.Caption = Text1.Text
    Else
    End If
    End If
    
    
    If Nickveinte.value = True And Nickveinte.Caption = "Slot Vacío" Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK20", Text1.Text)
    Nickveinte.Caption = Text1.Text
    ElseIf Nickveinte.value = True And Not Nickveinte.Caption = "Slot Vacío" Then
    If MsgBox("¿Estas seguro que deseas reemplazarlo a " & Nickveinte.Caption & ", por " & Text1.Text & "?", vbYesNo, "") = vbYes Then
    Frame2.Visible = False
     Call WriteVar(IniPath & "Amigos.ini", "FRIENDS", "NICK20", Text1.Text)
    Nickveinte.Caption = Text1.Text
    Else
    End If
    End If
    
    
End Sub

Private Sub Label2_Click(index As Integer)

End Sub

Private Sub Form_Load()
'Cargamos los amigos
Nickuno.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK1")
Nickdos.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK2")
Nicktres.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK3")
Nickcuatro.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK4")
Nickcinco.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK5")
Nickseis.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK6")
Nicksiete.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK7")
Nickocho.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK8")
Nicknueve.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK9")
Nickdiez.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK10")
Nickonce.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK11")
Nickdoce.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK12")
Nicktrece.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK13")
Nickcatorce.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK14")
Nickquince.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK15")
Nickdieciseis.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK16")
Nickdiecisiete.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK17")
Nickdieciocho.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK18")
Nickdiecinueve.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK19")
Nickveinte.Caption = GetVar(IniPath & "Amigos.ini", "FRIENDS", "NICK20")
End Sub

Private Sub Nickquience_Click()

End Sub

