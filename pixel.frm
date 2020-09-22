VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3690
   ClientTop       =   1020
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   3480
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Source"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   maxx = Label1.Width                          'get label width
   maxy = Label1.Height + (Label1.Height / 2)   'get label height add extra height for flame
   ReDim new_flame(maxx, maxy)                  'resize array to label
   ReDim old_flame(maxx, maxy)
End Sub

Private Sub Timer1_Timer()
  'This is the main timer,  Displays and updates the flame
  Dim X, Y As Integer                           'store current x and y pos.
  Dim red, green, blue As Long                  'store colours
  
  'This part generates the flame :)
  For X = 1 To maxx - 1
     For Y = 1 To maxy - 1
        red = new_flame(X + 1, Y).r             'Add up the surrounding red colours
        red = red + new_flame(X - 1, Y).r
        red = red + new_flame(X, Y + 1).r
        red = red + new_flame(X, Y - 1).r
        
        green = new_flame(X + 1, Y).g           'Add up the surrounding green colours
        green = green + new_flame(X - 1, Y).g
        green = green + new_flame(X, Y + 1).g
        green = green + new_flame(X, Y - 1).g
  
'        blue = blue + new_flame(X + 1, Y).b    'Add up the surrounding blue colours
'        blue = blue + new_flame(X - 1, Y).b
'        blue = blue + new_flame(X, Y + 1).b
'        blue = blue + new_flame(X, Y - 1).b
        
        'uses the row above (y-1) to give the effect of moving up!
        If old_flame(X, Y - 1).c = False Then   'if pixel is part of flame update
          tmp = (Rnd * Flame_Height)                      'pick a number from the air!
          old_flame(X, Y - 1).r = red / 4 - (tmp) ' Average the red and decrease the colour
          old_flame(X, Y - 1).g = (green / 4) - (tmp + 8) ' Average the green and decrease the colour
    
'         old_flame(X, Y - 1).b = blue / 4 ' Average the blue
    
          If old_flame(X, Y - 1).r < 0 Then old_flame(X, Y - 1).r = 0  'Check colours haven`t gone below 0
          If old_flame(X, Y - 1).g < 0 Then old_flame(X, Y - 1).g = 0
'          If old_flame(X, Y - 1).b < 0 Then old_flame(X, Y - 1).b = 0
        End If
     Next Y
  Next X
  
  'This loop Displays and updates the array
  For X = 1 To maxx
     For Y = 1 To maxy
        new_flame(X, Y).r = old_flame(X, Y).r     ' update array
        new_flame(X, Y).g = old_flame(X, Y).g
'        new_flame(X, Y).b = old_flame(X, Y).b
        'put the pixel!
        Me.PSet (Label1.Left + X, Label1.Top + Y - Int(Label1.Height / 2)), RGB(new_flame(X - 1, Y).r, new_flame(X - 1, Y).g, new_flame(X - 1, Y).b)
     Next Y
  Next X
End Sub

Private Sub Timer2_Timer()
    'This timer only initializes the array colours
    
    For X = 1 To maxx
     For Y = 1 To maxy
          If Point(Label1.Left + X, Label1.Top + Label1.Height - Y) <> 0 Then ' is there any colour at this point
           new_flame(X, maxy - Y).r = 255   ' Set colour to Yellow
           new_flame(X, maxy - Y).g = 255
           new_flame(X, maxy - Y).b = 0
           new_flame(X, maxy - Y).c = True  ' Is a permenant colour
          Else
           new_flame(X, maxy - Y).r = 0
           new_flame(X, maxy - Y).g = 0
           new_flame(X, maxy - Y).b = 0
           new_flame(X, maxy - Y).c = False ' Can be any colour
          End If
          
          old_flame(X, maxy - Y).r = new_flame(X, maxy - Y).r  'old_flame=new_flame
          old_flame(X, maxy - Y).g = new_flame(X, maxy - Y).g
          old_flame(X, maxy - Y).b = new_flame(X, maxy - Y).b
          old_flame(X, maxy - Y).c = new_flame(X, maxy - Y).c
     Next Y
  Next X
  Label1.Visible = False
  Timer1.Enabled = True   ' Call the Fire brigade :)
  Timer2.Enabled = False  ' Turn off the taps!
End Sub
