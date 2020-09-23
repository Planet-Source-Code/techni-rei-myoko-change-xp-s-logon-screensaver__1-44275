VERSION 5.00
Begin VB.UserControl XPCMD 
   BackStyle       =   0  'Transparent
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   ControlContainer=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   1050
   ToolboxBitmap   =   "XPCMD.ctx":0000
   Begin VB.Image imgup 
      Height          =   60
      Index           =   8
      Left            =   840
      Picture         =   "XPCMD.ctx":0312
      Top             =   360
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   45
      Index           =   7
      Left            =   720
      Picture         =   "XPCMD.ctx":0384
      Stretch         =   -1  'True
      Top             =   360
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   6
      Left            =   600
      Picture         =   "XPCMD.ctx":03F6
      Top             =   360
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   5
      Left            =   840
      Picture         =   "XPCMD.ctx":0468
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   4
      Left            =   720
      Picture         =   "XPCMD.ctx":04DA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   3
      Left            =   600
      Picture         =   "XPCMD.ctx":054C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   2
      Left            =   840
      Picture         =   "XPCMD.ctx":05BE
      Top             =   120
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   1
      Left            =   720
      Picture         =   "XPCMD.ctx":0630
      Stretch         =   -1  'True
      Top             =   120
      Width           =   60
   End
   Begin VB.Image imgup 
      Height          =   60
      Index           =   0
      Left            =   600
      Picture         =   "XPCMD.ctx":06A2
      Top             =   120
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   8
      Left            =   360
      Picture         =   "XPCMD.ctx":0714
      Top             =   360
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   7
      Left            =   240
      Picture         =   "XPCMD.ctx":0786
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   6
      Left            =   120
      Picture         =   "XPCMD.ctx":07F8
      Top             =   360
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   5
      Left            =   360
      Picture         =   "XPCMD.ctx":086A
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   4
      Left            =   240
      Picture         =   "XPCMD.ctx":08DC
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   3
      Left            =   120
      Picture         =   "XPCMD.ctx":094E
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   2
      Left            =   360
      Picture         =   "XPCMD.ctx":09C0
      Top             =   120
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   1
      Left            =   240
      Picture         =   "XPCMD.ctx":0A32
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgdown 
      Height          =   60
      Index           =   0
      Left            =   120
      Picture         =   "XPCMD.ctx":0AA4
      Top             =   120
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "XPCMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click()

Public Property Let state(cmdstate As Boolean)
Dim count As Long
For count = 0 To 8
    imgup(count).Visible = cmdstate
    imgdown(count).Visible = Not cmdstate
Next
End Property
Private Sub imgdown_Click(index As Integer)
UserControl_Click
End Sub

Private Sub imgdown_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub imgdown_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub imgup_Click(index As Integer)
UserControl_Click
End Sub

Private Sub imgup_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub imgup_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseUp Button, Shift, x, y
If x >= 0 And x <= UserControl.Width Then
If y >= 0 And y <= UserControl.Height Then
RaiseEvent Click
End If
End If
End Sub

Public Sub UserControl_Click()
RaiseEvent Click
End Sub

Public Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
state = False
End Sub

Public Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
state = True
End Sub

Private Sub UserControl_Resize()
Dim count As Long
imgup(0).Move 0, 0
imgup(2).Move UserControl.Width - imgup(2).Width, 0
imgup(6).Move 0, UserControl.Height - imgup(6).Height
imgup(8).Move UserControl.Width - imgup(8).Width, UserControl.Height - imgup(8).Height

imgup(1).Move imgup(0).Width, 0, imgup(2).Left - imgup(0).Width
imgup(7).Move imgup(0).Width, UserControl.Height - imgup(7).Height, imgup(2).Left - imgup(0).Width

imgup(3).Move 0, imgup(0).Height, imgup(3).Width, UserControl.Height - imgup(0).Height - imgup(6).Height
imgup(5).Move UserControl.Width - imgup(5).Width, imgup(0).Height, imgup(5).Width, UserControl.Height - imgup(0).Height - imgup(6).Height

imgup(4).Move imgup(0).Width, imgup(0).Height, imgup(1).Width, imgup(3).Height

For count = 0 To 8
    imgdown(count).Move imgup(count).Left, imgup(count).Top, imgup(count).Width, imgup(count).Height
Next
End Sub
