Attribute VB_Name = "Application"
Option Explicit
Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Function IsAppActive() As Boolean
    IsAppActive = (GetActiveWindow <> 0)
End Function
