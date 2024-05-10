Attribute VB_Name = "modDX8FIFO"
Option Explicit



Sub CargarCabezas()
    Dim N As Integer
    Dim I As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For I = 1 To Numheads
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(HeadData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(HeadData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(HeadData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(HeadData(I).Head(4), Miscabezas(I).Head(4), 0)
        End If
    Next I
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim I As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For I = 1 To NumCascos
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0)
        End If
    Next I
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim I As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For I = 1 To NumCuerpos
        Get #N, , MisCuerpos(I)
        
        If MisCuerpos(I).Body(1) Then
            InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
            InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
            InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
            InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
            
            BodyData(I).HeadOffset.x = MisCuerpos(I).HeadOffsetX
            BodyData(I).HeadOffset.y = MisCuerpos(I).HeadOffsetY
        End If
    Next I
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim I As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For I = 1 To NumFxs
        Get #N, , FxData(I)
    Next I
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim I As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For I = 1 To NumTips
        Get #N, , Tips(I)
    Next I
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim I As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For I = 1 To Nu
        Get #N, , bLluvia(I)
    Next I
    
    Close #N
End Sub

Public Function LoadGrhData() As Boolean
On Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
   
    'Open files
    handle = FreeFile()
    Open App.path & "\INIT\Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
   
    'Get file version
    Get handle, , fileVersion
   
    'Get number of grhs
    Get handle, , grhCount
   
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
   
    While Not EOF(handle)
        Get handle, , Grh
       
        With GrhData(Grh)
            'Get number of frames
            GrhData(Grh).Active = True
           
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
           
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
               
                Get handle, , .Speed
               
                If .Speed <= 0 Then Resume Next
               
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then Resume Next
               
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then Resume Next
               
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then Resume Next
               
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then Resume Next
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
               
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then Resume Next
               
                Get handle, , .sY
                If .sY < 0 Then Resume Next
               
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then Resume Next
               
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then Resume Next
               
                'Compute width and height
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
               
                .Frames(1) = Grh
            End If
        End With
    Wend
   
    Close handle
   
Dim Count As Long
 
Open App.path & "\INIT\minimap.dat" For Binary As #1
    Seek #1, 1
    For Count = 1 To 15000
        If GrhData(Count).Active Then
            Get #1, , GrhData(Count).MiniMap_color
        End If
    Next Count
Close #1
 
    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
End Function
