VERSION 5.00
Begin VB.UserControl XPWIN 
   BackColor       =   &H00F7DED6&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   2850
   ToolboxBitmap   =   "XPWIN.ctx":0000
   Begin VB.PictureBox picmain 
      BackColor       =   &H00F7DED6&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   3
      Left            =   1320
      Picture         =   "XPWIN.ctx":0312
      ToolTipText     =   "Up"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   2
      Left            =   960
      Picture         =   "XPWIN.ctx":07C8
      ToolTipText     =   "Down"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   1
      Left            =   240
      Picture         =   "XPWIN.ctx":0C7E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "XPWIN.ctx":0CBF
      Top             =   120
      Width           =   45
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   3
      Left            =   2640
      Picture         =   "XPWIN.ctx":0D04
      Top             =   120
      Width           =   30
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   0
      Left            =   240
      Picture         =   "XPWIN.ctx":0D49
      ToolTipText     =   "Down"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   1
      Left            =   600
      Picture         =   "XPWIN.ctx":11CD
      ToolTipText     =   "Up"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgbutton 
      Height          =   285
      Left            =   2160
      Picture         =   "XPWIN.ctx":164C
      Tag             =   "0"
      Top             =   180
      Width           =   285
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   0
      Left            =   120
      Picture         =   "XPWIN.ctx":1AD0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   30
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   1
      Left            =   2640
      Picture         =   "XPWIN.ctx":1B07
      Stretch         =   -1  'True
      Top             =   600
      Width           =   30
   End
   Begin VB.Image imgborder 
      Height          =   30
      Index           =   2
      Left            =   120
      Picture         =   "XPWIN.ctx":1B3E
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2565
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "XPWIN.ctx":1B75
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "XPWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim xpwin_state As Boolean
Dim xpwin_height As Single
Dim xpwin_icon  As Boolean
Dim mousex As Single, mousey As Single
Public Event statechange(state As Boolean)
Public Event Click(x As Single, y As Single)
Public Event mousedown(x As Single, y As Single, Button As Integer)
Public Event mouseup(x As Single, y As Single, Button As Integer)
Public Property Let CanResize(resizeable As Boolean)
    imgbutton.Enabled = resizeable
End Property
Public Property Get CanResize() As Boolean
    CanResize = imgbutton.Enabled
End Property
Public Property Let state(state As Boolean)
    xpwin_state = state
    imgbutton_Click
End Property
Public Property Get state() As Boolean
    state = xpwin_state
End Property

Private Sub imgbutton_Click()
xpwin_state = Not xpwin_state
Select Case xpwin_state
     Case True 'was down, move to up
          UserControl.Height = xpwin_height
          imgbutton.Picture = imgstate(1).Picture
     Case False
          xpwin_height = UserControl.Height
          UserControl.Height = imghead(0).Height
          imgbutton.Picture = imgstate(0).Picture
End Select
RaiseEvent statechange(xpwin_state)
End Sub

Private Sub imghead_Click(index As Integer)
    RaiseEvent Click(mousex, mousey)
End Sub

Private Sub imghead_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mousedown(mousex, mousey, Button)
End Sub

Private Sub imghead_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
mousex = x + imghead(index).Left
mousey = y + imghead(index).Top
End Sub

Private Sub imghead_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mouseup(mousex, mousey, Button)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click(mousex, mousey)
End Sub

Private Sub UserControl_Initialize()
     xpwin_state = True
     xpwin_icon = False
     imgbutton.Picture = imgstate(1).Picture
End Sub
Private Sub imgbutton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If state = True Then
    imgbutton.Picture = imgstate(3).Picture
Else
    imgbutton.Picture = imgstate(2).Picture
End If
End Sub

Private Sub imgbutton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If state = True Then
    imgbutton.Picture = imgstate(1).Picture
Else
    imgbutton.Picture = imgstate(0).Picture
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mousedown(mousex, mousey, Button)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
mousey = y
mousex = x
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mouseup(mousex, mousey, Button)
End Sub

Private Sub UserControl_Resize()
imghead(0).Move 0, 0
imghead(1).Move imghead(0).Left, 0
imghead(3).Move UserControl.Width - imghead(3).Width, 0
imghead(2).Move imghead(1).Left + imghead(1).Width, 0, imghead(3).Left - imghead(1).Left - imghead(1).Width
imgbutton.Move UserControl.Width - imgbutton.Width * 1.3, imghead(0).Height / 2 - imgbutton.Height / 2
If UserControl.Height > imghead(0).Height Then
    imgborder(0).Move 0, imghead(0).Height, imgborder(0).Width, UserControl.Height - imghead(0).Height
    imgborder(1).Move UserControl.Width - imgborder(1).Width, imghead(0).Height, imgborder(1).Width, UserControl.Height - imghead(0).Height
    imgborder(2).Move 0, UserControl.Height - imgborder(2).Height, UserControl.Width
    picmain.Move imgborder(0).Width, imghead(0).Height, UserControl.Width - imgborder(0).Width * 2, UserControl.Height - imghead(0).Height - imgborder(2).Height
End If
End Sub
