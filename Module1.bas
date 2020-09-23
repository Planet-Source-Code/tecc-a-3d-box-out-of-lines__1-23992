Attribute VB_Name = "Module1"
'I should change these bytes to boolean, but
'what the heck
Public ccc As Byte
'color setting (0 or 1) on or off
Public ccc1 As Byte
'redraw setting (0 or 1) on or off
Public cls11 As Byte
'Clear form setting (0 or 1) on or off
'
Public SCl As Long
'SCl is the variable for the scale of the
'object.. default scale is 100
Public H1 As Long
'H1 is the object's y origin
Public W1 As Long
'H1 is the object's x origin
Public x4 As Long
'x4 holds the current X coordinate to pass on to
'DD, which is not on the main form
Public y4 As Long
'y4 holds the current Y coordinate to pass on to
'DD, which is not on the main form
Public ChngPos As Boolean
'ChngPos declares whether or not to allow
'changes to the origin

Public Function DD()
'if the user changes the origin with the
'arrow keys, update the object with the
'new origin
Dim x3 As Long, y3 As Long
x3 = x4
y3 = y4
If cls11 <> 1 Then
Form1.Cls
Else
End If
If ccc <> 1 Then
Form1.Line (W1 / 2, H1 / 2)-(x3, y3), RGB(255, 0, 0)
Form1.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(x3 + SCl, y3 + SCl), RGB(255, 0, 0)
Form1.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(x3 + SCl * 2, y3 + SCl), RGB(255, 255, 255)
Form1.Line (W1 / 2 + SCl, H1 / 2)-(x3 + SCl, y3), RGB(255, 255, 255)

Form1.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2 + SCl), RGB(0, 255, 0)
Form1.Line (x3, y3)-(SCl + x3, SCl + y3), RGB(0, 255, 0)
Form1.Line (x3 + SCl, y3)-(x3, y3), RGB(0, 255, 0)
Form1.Line (x3 + SCl, y3)-(x3 + SCl * 2, y3 + SCl), RGB(0, 255, 0)
Form1.Line (x3 + SCl, y3 + SCl)-(x3 + SCl * 2, y3 + SCl), RGB(0, 255, 0)
Form1.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2), RGB(0, 255, 0)
Form1.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(W1 / 2 + SCl, H1 / 2), RGB(0, 255, 0)
Form1.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(W1 / 2 + SCl * 2, H1 / 2 + SCl), RGB(0, 255, 0)
Else
Form1.Line (W1 / 2, H1 / 2)-(x3, y3)
Form1.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(x3 + SCl, y3 + SCl)
Form1.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(x3 + SCl * 2, y3 + SCl)
Form1.Line (W1 / 2 + SCl, H1 / 2)-(x3 + SCl, y3)

Form1.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2 + SCl)
Form1.Line (x3, y3)-(SCl + x3, SCl + y3)
Form1.Line (x3 + SCl, y3)-(x3, y3)
Form1.Line (x3 + SCl, y3)-(x3 + SCl * 2, y3 + SCl)
Form1.Line (x3 + SCl, y3 + SCl)-(x3 + SCl * 2, y3 + SCl)
Form1.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2)
Form1.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(W1 / 2 + SCl, H1 / 2)
Form1.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(W1 / 2 + SCl * 2, H1 / 2 + SCl)
End If
End Function
