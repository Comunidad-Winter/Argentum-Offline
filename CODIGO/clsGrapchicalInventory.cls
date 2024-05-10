VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item

Dim MouseDownSelectedItem As Long
Dim MouseX As Integer
Dim MouseY As Integer

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindowNpc As PictureBox
Attribute InventoryWindowNpc.VB_VarHelpID = -1
Dim WithEvents InventoryWindowUser As PictureBox
Attribute InventoryWindowUser.VB_VarHelpID = -1

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get grhindex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    grhindex = UserInventory(slot).grhindex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With

End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ 32 Or TilePixelWidth
    temp_y = Y \ 32 Or TilePixelHeight
    
'ejejjejeje
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
'jejejejeje
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
    ClickItem = TempItem
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhindex Then
            ClickItem = TempItem
        End If
    End If
    

End Function

Public Sub DrawInventory()
    Dim I As Byte
    Dim X As Integer
    Dim Y As Integer
    
    
    If MouseOK = True Then
Call engine.Draw_GrhIndex(MouseItem, MouseX, MouseY)
End If


    For I = 1 To UBound(UserInventory)
        If UserInventory(I).grhindex Then
            X = ((I - 1) Mod (InventoryWindow.Width / 32)) * 32 + 2
            Y = ((I - 1) \ (InventoryWindow.Width / 32)) * 32 + 2
            
            Call engine.Draw_GrhIndexAlpha(UserInventory(I).grhindex, X, Y)
                
                If InvSelectedItem = I Then
                    Call engine.Draw_GrhIndex(UserInventory(I).grhindex, X, Y)
                End If
                
            Call engine.Text_Render_ext(UserInventory(I).Amount, Y, X, 40, 40, D3DColorXRGB(255, 255, 255))
            If UserInventory(I).Equipped Then
                Call engine.Text_Render_ext("+", Y + 20, X + 20, 40, 40, D3DColorXRGB(255, 255, 255))
            End If
        End If
    Next I
End Sub

Public Sub Initialize(ByRef InvPic As PictureBox, ByRef InvNpc As PictureBox, ByRef InvUser As PictureBox)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the reference to the picture box on which to render the inventory
'***************************************************
    Set InventoryWindow = InvPic
    Set InventoryWindowNpc = InvNpc
    Set InventoryWindowUser = InvUser
    '\TODO
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
    End If
    
    If MouseItem = 0 And Button = 2 Then
MouseItem = UserInventory(TempItem).grhindex
End If
If MouseItem = 0 And PUEDO = True Then
MouseItem = UserInventory(TempItem).grhindex
End If
If Button = vbRightButton Then
MouseX = X
MouseY = Y
MouseOK = True
End If
If PUEDO = True Then
MouseX = X
MouseY = Y
MouseOK = True
ButtonIN = True
End If

End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 Then
            'Call Protocol.WriteSwapObjects(InvSelectedItem, MouseDownSelectedItem)
    End If
        
        If Button = vbRightButton Then
MouseOK = False
MouseItem = 0
End If
MouseItem = 0
    
#If ConMenuesConextuales = 1 Then
    'If it was a right click on the same item we had selected before, show popup menu
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then
        'Show the provided menu as a popup
        Call InventoryWindow.Parent.PopupMenu(ItemMenu, , X, Y)
    End If
#End If
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long
 
If Button = 2 Then 'click derecho
TempItem = ClickItem(X, Y)
           
If UserInventory(TempItem).grhindex Then
MouseDownSelectedItem = TempItem
Exit Sub
End If
End If
       
MouseDownSelectedItem = 0
End Sub




    Sub DrawInventoryUser()
    '---------------------------------------------------------------------------------------
    ' Procedure : DrawInventoryUser
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
        Dim I As Byte
        Dim X As Integer
        Dim Y As Integer
       
        For I = 1 To UBound(UserInventory)
            If UserInventory(I).grhindex Then
                X = ((I - 1) Mod (InventoryWindowUser.Width / 32)) * 32 + 2
                Y = ((I - 1) \ (InventoryWindowUser.Width / 32)) * 32 + 2
                If InvSelectedItemUser = I Then
                    Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(255, 0, 0))
                End If
                Call engine.Draw_GrhIndex(UserInventory(I).grhindex, X, Y)
                Call engine.Text_Render_ext(UserInventory(I).Amount, Y, X, 40, 40, D3DColorXRGB(255, 255, 255))
            End If
        Next I
    End Sub
     
    Private Function ClickItemUser(ByVal X As Long, ByVal Y As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : ClickItemNpc
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
        Dim TempItem As Long
        Dim temp_x As Long
        Dim temp_y As Long
       
        temp_x = X \ 32
        temp_y = Y \ 32
       
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindowUser.ScaleWidth \ 32) + 1
       
        'Make sure it's within limits
        If TempItem <= MAX_INVENTORY_SLOTS Then
            'Make sure slot isn't empty
            If UserInventory(TempItem).grhindex Then
                ClickItemUser = TempItem
            Else
                ClickItemUser = 0
            End If
        End If
    End Function
     
    Private Sub InventoryWindowUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '---------------------------------------------------------------------------------------
    ' Procedure : ClickItemNpc
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
        Dim prevSelItem As Long
       
        If X < 0 Or Y < 0 Or X > InventoryWindowUser.Width Or Y > InventoryWindowUser.Height Then _
        Exit Sub
       
        prevSelItem = InvSelectedItemUser
     
        InvSelectedItemUser = ClickItemUser(CInt(X), CInt(Y))
       
    End Sub
     
    Sub DrawInventoryNpc()
     Dim I As Byte
    Dim X As Integer
    Dim Y As Integer
   
    For I = 1 To UBound(NPCInventory)
        If NPCInventory(I).grhindex Then
            X = ((I - 1) Mod (InventoryWindowNpc.Width / 32)) * 32 + 2
           Y = ((I - 1) \ (InventoryWindowNpc.Width / 32)) * 32 + 2
            If InvSelectedItemNpc = I Then
                Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(200, 200, 200))
            End If
            Call engine.Draw_GrhIndex(NPCInventory(I).grhindex, X, Y)
            Call engine.Text_Render_ext(NPCInventory(I).Amount, Y, X, 40, 40, D3DColorXRGB(255, 255, 255))
        End If
    Next I
End Sub
     
    Private Function ClickItemNpc(ByVal X As Long, ByVal Y As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : ClickItemNpc
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
        Dim TempItem As Long
        Dim temp_x As Long
        Dim temp_y As Long
       
        temp_x = X \ 32
        temp_y = Y \ 32
       
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindowNpc.ScaleWidth \ 32) + 1
       
        'Make sure it's within limits
        If TempItem <= MAX_INVENTORY_SLOTS Then
            'Make sure slot isn't empty
            If NPCInventory(TempItem).grhindex Then
                ClickItemNpc = TempItem
            Else
                ClickItemNpc = 0
            End If
        End If
    End Function
     
    Private Sub InventoryWindowNpc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '---------------------------------------------------------------------------------------
    ' Procedure : ClickItemNpc
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
        Dim prevSelItem As Long
       
        If X < 0 Or Y < 0 Or X > InventoryWindowNpc.Width Or Y > InventoryWindowNpc.Height Then _
        Exit Sub
       
        prevSelItem = InvSelectedItemNpc
     
        InvSelectedItemNpc = ClickItemNpc(CInt(X), CInt(Y))
              
    End Sub

