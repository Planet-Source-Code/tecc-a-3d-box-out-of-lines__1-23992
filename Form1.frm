VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D box o' Lines"
   ClientHeight    =   3915
   ClientLeft      =   1725
   ClientTop       =   1500
   ClientWidth     =   5025
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5025
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Off"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4020
      TabIndex        =   3
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Scale..."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4020
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Left            =   3960
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CLS on"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4020
      TabIndex        =   1
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No Redraw"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4020
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   3960
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   3960
      Top             =   420
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   3960
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 27 'escape key
'exit the program
Unload Me
End
'if the user presses z, add to the object scale
'if the user presses x, subtract from it
Case 90 'key z
SCl = SCl + 10
Case 88 'key x
SCl = SCl - 10
'move the shape's origin on the form
Case vbKeyLeft 'arrow left
W1 = W1 - 40
Case vbKeyRight 'arrow right
W1 = W1 + 40
Case vbKeyUp 'arrow up
H1 = H1 - 40
Case vbKeyDown 'arrow down
H1 = H1 + 40
End Select
'call DD to redraw the shape
DD
End Sub

Private Sub Form_Load()
'set default settings
ccc = 1
ccc1 = 0
SCl = 100
H1 = Me.ScaleHeight
W1 = Me.ScaleWidth

MsgBox "Sorry for the poor code documentation" & vbCrLf & vbCrLf & "Controls:" & vbCrLf & vbCrLf & "z - Add 10 pixels to object scale" & vbCrLf & "x - Subtract 10 pixels from object scale" & vbCrLf & "arrow keys - Change object origin" & vbCrLf & vbCrLf & "ESC to exit, have fun!", , "3-D Lines"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if the mouse button is clicked, allow the user
'to choose an origin with the mouse position
ChngPos = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
x4 = X
y4 = Y
'if chngpos is not true, then dont allow the user
'to change the origin
'if it is true, skip the normal object drawing,
'and change the position
If ChngPos <> True Then
'if chngpos isnt true, and the clear form setting
'is on (0), then clear the form
'if not, skip the clearing of the form
If cls11 <> 1 Then
'clear form
Me.Cls
Else
End If
'if colors are on, add color to lines drawn
'if not, leave them black
If ccc <> 1 Then
'code for the box
Me.Line (W1 / 2, H1 / 2)-(X, Y), RGB(255, 0, 0)
Me.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(X + SCl, Y + SCl), RGB(255, 0, 0)
Me.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(X + SCl * 2, Y + SCl), RGB(255, 255, 255)
Me.Line (W1 / 2 + SCl, H1 / 2)-(X + SCl, Y), RGB(255, 255, 255)

Me.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2 + SCl), RGB(0, 255, 0)
Me.Line (X, Y)-(SCl + X, SCl + Y), RGB(0, 255, 0)
Me.Line (X + SCl, Y)-(X, Y), RGB(0, 255, 0)
Me.Line (X + SCl, Y)-(X + SCl * 2, Y + SCl), RGB(0, 255, 0)
Me.Line (X + SCl, Y + SCl)-(X + SCl * 2, Y + SCl), RGB(0, 255, 0)
Me.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2), RGB(0, 255, 0)
Me.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(W1 / 2 + SCl, H1 / 2), RGB(0, 255, 0)
Me.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(W1 / 2 + SCl * 2, H1 / 2 + SCl), RGB(0, 255, 0)
Else
Me.Line (W1 / 2, H1 / 2)-(X, Y)
Me.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(X + SCl, Y + SCl)
Me.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(X + SCl * 2, Y + SCl)
Me.Line (W1 / 2 + SCl, H1 / 2)-(X + SCl, Y)

Me.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2 + SCl)
Me.Line (X, Y)-(SCl + X, SCl + Y)
Me.Line (X + SCl, Y)-(X, Y)
Me.Line (X + SCl, Y)-(X + SCl * 2, Y + SCl)
Me.Line (X + SCl, Y + SCl)-(X + SCl * 2, Y + SCl)
Me.Line (W1 / 2, H1 / 2)-(W1 / 2 + SCl, H1 / 2)
Me.Line (W1 / 2 + SCl * 2, H1 / 2 + SCl)-(W1 / 2 + SCl, H1 / 2)
Me.Line (W1 / 2 + SCl, H1 / 2 + SCl)-(W1 / 2 + SCl * 2, H1 / 2 + SCl)
End If
Else
'if the mouse button is down, move the origin
H1 = Y + SCl
W1 = X + SCl
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'once the mouse button is let go, do not change
'the origin anymore
ChngPos = False
End Sub

Private Sub Label1_Click()
'option for color
Select Case ccc
Case 0
ccc = 1
Label1.Caption = "Color off"
Case 1
Label1.Caption = "Color on"
ccc = 0
End Select

End Sub

Private Sub Label2_Click()
'option for auto-redraw
Select Case ccc1
Case 0
ccc1 = 1
Me.Cls
Label2.Caption = "Redraw"
Case 1
Label2.Caption = "No Redraw"
ccc1 = 0
Me.Cls
End Select
Form1.AutoRedraw = ccc1
Me.Cls
Me.Cls
End Sub

Private Sub Label3_Click()
'option for clear form
Select Case cls11
Case 0
cls11 = 1
Label3.Caption = "CLS off"
Case 1
Label3.Caption = "CLS on"
cls11 = 0
End Select

End Sub

Private Sub Label4_Click()
'set a new scale for the object
SCl = InputBox("Enter New Scale", "Scale", 100)
End Sub
