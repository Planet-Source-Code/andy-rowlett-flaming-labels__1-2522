Attribute VB_Name = "Module1"
Public Const Flame_Height = 30  ' Higher the number the shorter the flame

Type pix
    r As Integer   ' Red
    g As Integer   ' Green
    b As Integer   ' Blue
    c As Boolean   ' Constant Colour
End Type

Public maxx As Integer   ' Array max x
Public maxy As Integer   ' Array max y

Public new_flame() As pix  ' Flames buffers
Public old_flame() As pix
