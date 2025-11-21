VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB6 SysTray demo"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show balloon tip"
      Height          =   525
      Left            =   4410
      TabIndex        =   2
      Top             =   1515
      Width           =   2370
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4410
      TabIndex        =   1
      Top             =   720
      Width           =   2370
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   285
      TabIndex        =   0
      Top             =   450
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tray icon events:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents sTray As SysTray
Attribute sTray.VB_VarHelpID = -1

Friend Sub SetButtonTitle()
  Command1.Caption = IIf(IsInTray() = True, "Remove icon from SysTray", "Create icon in SysTray")
End Sub

Private Function IsInTray() As Boolean
  If Not sTray Is Nothing Then IsInTray = sTray.InTray
End Function

Private Sub PutInTray()
  KillTray
  Set sTray = New SysTray
  With sTray
    'hWndParent must be set.
    'SysTray will subclass itself into the parent window's message stream.
    .hWndParent = Me.hWnd
    
    'Icon is optional.
    'Note that GDI+ is not supported!
    .Icon = Me.Icon
    
    'This is the tooltip that will pop up when the mouse hovers over the tray icon.
    .Tip = "This tray icon is a demo."
    
    'Put the icon in the tray.
    .InTray = True
  End With
End Sub

Private Sub KillTray()
  If Not sTray Is Nothing Then
    sTray.InTray = False
    Set sTray = Nothing
  End If
End Sub

Private Sub AddToList(NewText As String)
  List1.AddItem NewText
  List1.ListIndex = (List1.ListCount - 1)
End Sub

Private Sub Command1_Click()
  If IsInTray Then KillTray Else PutInTray
  SetButtonTitle
End Sub

Private Sub Command2_Click()
  If IsInTray Then
    sTray.ShowBalloonTip "SysTray VB6 Demo", "You should click on this. Or not. It's your life."
  Else
    MsgBox "Tray icon not active!", vbOKOnly Or vbInformation, "Error!"
  End If
End Sub

Private Sub Form_Load()
  SetButtonTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If IsInTray Then KillTray
End Sub

Private Sub sTray_BalloonTipClick()
  AddToList "User clicked the balloon tip."
End Sub

Private Sub sTray_BalloonTipTimeout()
  AddToList "Balloon tip timed out"
End Sub

Private Sub sTray_LeftMouseDblClick()
  AddToList "Left mouse dbl-click"
End Sub

Private Sub sTray_LeftMouseDown()
  AddToList "Left mouse down"
End Sub

Private Sub sTray_LeftMouseUp()
  AddToList "Left mouse up"
End Sub

Private Sub sTray_MiddleMouseDblClick()
  AddToList "Middle mouse dbl-click"
End Sub

Private Sub sTray_MiddleMouseDown()
  AddToList "Middle mouse down"
End Sub

Private Sub sTray_MiddleMouseUp()
  AddToList "Middle mouse up"
End Sub

Private Sub sTray_MouseHover()
  AddToList "Mouse hover"
End Sub

Private Sub sTray_MouseLeave()
  AddToList "Mouse leave"
End Sub

Private Sub sTray_RightMouseDblClick()
  AddToList "Right mouse dbl-click"
End Sub

Private Sub sTray_RightMouseDown()
  AddToList "Right mouse down"
End Sub

Private Sub sTray_RightMouseUp()
  AddToList "Right mouse up"
End Sub

Private Sub sTray_MouseMove()
  AddToList "Mouse move"
End Sub
