Attribute VB_Name = "Mod_Briancito"
'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
 
'RGB Type
Public Type RGB
    r As Long
    g As Long
    B As Long
End Type
 
Public Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    Speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'index de la particula que debe ser = que le pusimos al server
Public Enum ParticulaMedit
    CHICO = 34
    MEDIANO = 20
    GRANDE = 27
    XGRANDE = 35
    GIGANTE = 36
End Enum


Public Actual As Byte

'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Resolucion Briancito.vd@hotmail.com
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Resolucion Briancito.vd@hotmail.com
Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1
 
Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
 
Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long
Private bNoResChange As Boolean
 
 
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Const EM_SETEVENTMASK = &H445
Const EN_LINK = &H70B
Const ENM_LINK = &H4000000
Const EM_AUTOURLDETECT = &H45B
Const EM_GETEVENTMASK = &H43B
Const GWL_WNDPROC = (-4)
Const WM_NOTIFY = &H4E
Const WM_LBUTTONDOWN = &H201
Const EM_GETTEXTRANGE = &H44B
 
Dim lOldProc As Long
Dim hWndRTB As Long
Dim hWndParent As Long

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type
 
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
Sub CargarParticulas()
Dim StreamFile As String
Dim loopc As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
    
StreamFile = App.path & "\INIT\Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).Speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, 44)
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
    Next loopc

End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long
    
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

End Function

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

End Function

Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(Text, LastPos + 1)
    End If
End Function

 
Public Sub SetResolucion()
 
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim CambiarResolucion As Boolean
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
    
    oldResWidth = Screen.width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.height \ Screen.TwipsPerPixelY
    
    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If
    
    If CambiarResolucion Then
        
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
            
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
        End With
        
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
        bNoResChange = True
    End If
End Sub
 
Public Sub ResetResolucion()
 
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    If Not bNoResChange Then
    
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub

Public Sub Detectar(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
  'Don't want to subclass twice!
  If lOldProc = 0 Then
    'Subclass!
    lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)
    SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
    SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
    hWndParent = hWndOwner
    hWndRTB = hWndTextbox
  End If
End Sub
 
Public Sub NoDetectar()
  If lOldProc Then
    SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
    'Reset the window procedure (stop the subclassing)
    SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
    'Set this to 0 so we can subclass again in future
    lOldProc = 0
  End If
End Sub
 
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    'Which message?
    Select Case uMsg
    Case WM_NOTIFY
        'Copy the notification header into our structure from the pointer
        CopyMemory uHead, ByVal lParam, Len(uHead)
        
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
            
            'What kind of message?
            Select Case eLink.msg
            
            Case WM_LBUTTONDOWN
              eText.chrg.cpMin = eLink.chrg.cpMin
              eText.chrg.cpMax = eLink.chrg.cpMax
              eText.lpstrText = Space$(1024)
              lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
              sText = Left$(eText.lpstrText, lLen)
              ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
                
            End Select
            
        End If
        
    End Select
    WndProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam)
End Function

Sub Relogueo()
UserName = NameRelogueo
UserPassword = PassRelogueo
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
        End If
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
If CheckUserData(False) = True Then
            EstadoLogin = loginaccount
            frmConnect.MousePointer = 99
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
End If
End Sub



