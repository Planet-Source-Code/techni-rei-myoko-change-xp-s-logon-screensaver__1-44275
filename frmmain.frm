VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Logon Screensaver"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin prjmain.XPWIN XPWIN 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      Begin prjmain.XPCMD XPCMD 
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   9
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&xit"
            Height          =   195
            Index           =   2
            Left            =   825
            TabIndex        =   10
            Top             =   75
            Width           =   285
         End
      End
      Begin prjmain.XPCMD XPCMD 
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Restore to logon.scr"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   75
            Width           =   1455
         End
      End
      Begin prjmain.XPCMD XPCMD 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Set to selected"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   6
            Top             =   75
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lstmain 
         Height          =   2655
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList iml 
         Left            =   6600
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.PictureBox picmain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgmain 
         Height          =   240
         Left            =   120
         Picture         =   "frmmain.frx":0902
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblmain 
         BackStyle       =   0  'Transparent
         Caption         =   "Logon Screensaver"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   60
         Width           =   1935
      End
   End
   Begin VB.FileListBox Filemain 
      Height          =   480
      Left            =   6960
      Pattern         =   "*.scr"
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Public Function isrunningonxp() As Boolean
    Dim Ver As Long, WinVer2 As Long, winver As String
    Ver = GetVersion()
    WinVer2 = Ver And &HFFFF&
    winver = Format((WinVer2 Mod 256) + ((WinVer2 \ 256) / 100), "Fixed")
    isrunningonxp = Left(winver, 1) = "5"
End Function

Public Sub dragform(hWnd As Long)
On Error Resume Next
  ReleaseCapture
  SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Form_Load()
    'Get the system directory, set the filebox to point to it
    Dim temp As String * 255, currentscr As String
    Dim filedata As FILEPROPERTIE

    If isrunningonxp = False Then
        MsgBox "This program is only for Windows XP", vbCritical, "Invalid OS Version"
        Unload Me
        End
    End If
    
    GetSystemDirectory temp, 255
    Filemain.Pattern = "*.scr" 'to be sure
    Filemain.Path = Trim(temp)
    
    Call exeFileInfo(chkdir(Filemain.Path, "logon.scr"), filedata)
    currentscr = Trim(filedata.FileDescription)
    
    If GetSetting("XP LOGO SCR", "MAIN", "HAS BACKED UP", False) = False Then
        FileCopy chkdir(Filemain.Path, "logon.scr"), chkdir(Filemain.Path, "logon.bak")
        Call SaveSetting("XP LOGO SCR", "MAIN", "HAS BACKED UP", True)
    End If
    
    Dim count As Long
    With lstmain.ListItems
    For count = 0 To Filemain.ListCount - 1
        drawfileicon chkdir(Filemain.Path, Filemain.List(count)), SmallIcon, picmain.hDC, 0, 0
        picmain.Refresh
        iml.ListImages.Add , Filemain.List(count), picmain.Image
        
        Call exeFileInfo(chkdir(Filemain.Path, Filemain.List(count)), filedata)
        
        .Add , Filemain.List(count), Filemain.List(count)
        .Item(.count).SubItems(1) = filedata.FileDescription
    Next
    Set lstmain.SmallIcons = iml
    For count = 1 To .count
        .Item(count).SmallIcon = 1
        .Item(count).Selected = currentscr = .Item(count).SubItems(1)
    Next
    autosizeall lstmain
    End With
End Sub
Public Function chkdir(dirmainectory As String, filemainname As String) As String
On Error Resume Next
If Right(dirmainectory, 1) <> "\" Then chkdir = dirmainectory & "\" & filemainname Else chkdir = dirmainectory & filemainname
End Function

Private Sub imgmain_Click()
If MsgBox("Are you sure you want to quit?", vbYesNo, "Quit Logon Screensaver") = vbYes Then
    Unload Me
    End
End If
End Sub

Private Sub lblcmd_Click(index As Integer)
XPCMD_Click index
End Sub

Private Sub lblcmd_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    XPCMD(index).state = False
End Sub

Private Sub lblcmd_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    XPCMD(index).state = True
End Sub

Private Sub lblmain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    XPWIN_mousedown lblmain.Left + x, lblmain.Top + y, 0
End Sub

Private Sub lstmain_DblClick()
    Shell "start """ & chkdir(Filemain.Path, lstmain.SelectedItem) & ""
End Sub

Public Sub XPCMD_Click(index As Integer)
On Error Resume Next
Select Case index
    Case 0 'set
        Kill chkdir(Filemain.Path, "logon.scr")
        FileCopy chkdir(Filemain.Path, lstmain.SelectedItem), chkdir(Filemain.Path, "logon.scr")
        MsgBox "Your screen saver has been set to " & lstmain.SelectedItem, vbInformation, "Screensaver set"
    Case 1 'restore
        Kill chkdir(Filemain.Path, "logon.scr")
        FileCopy chkdir(Filemain.Path, "logon.bak"), chkdir(Filemain.Path, "logon.scr")
        MsgBox "Your screen saver has been restored to logon.scr", vbInformation, "Screensaver restored"
    Case 2 'exit
        Unload Me
        End
End Select
End Sub

Public Sub XPWIN_mousedown(x As Single, y As Single, Button As Integer)
If y <= 375 Then dragform Me.hWnd
End Sub

Private Sub XPWIN_statechange(state As Boolean)
Me.Move Me.Left, Me.Top, XPWIN.Width, XPWIN.Height
End Sub
