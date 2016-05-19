Attribute VB_Name = "MainModule"

Option Explicit
Public Gxx As Long, Gyy As Long, Gkx As Long, Gky As Long
Public Oxx As Long, Oyy As Long
Function RescaleX(x_ As Long) As Long
    RescaleX = x_ + Oxx
End Function

Function RescaleY(y_ As Long) As Long
    RescaleY = Oyy - y_
End Function



