Attribute VB_Name = "modGlobal"
Option Explicit

Public Const WM_USER As Long = &H400
Public Const WM_MYHOOK As Long = WM_USER + 1
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10
Public Const WM_PRINTCLIENT = &H318

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MOUSEHOVER = &H2A1
Public Const WM_MOUSELEAVE = &H2A3

Public Const GWL_WNDPROC As Long = (-4)
Public Const GWL_HWNDPARENT As Long = (-8)
Public Const GWL_ID As Long = (-12)
Public Const GWL_STYLE As Long = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_USERDATA As Long = (-21)

Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2
Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_INFO = &H10

Public Const NIIF_NONE = &H0
Public Const NIIF_INFO = &H1
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_GUID = &H5
Public Const NIIF_ICON_MASK = &HF
Public Const NIIF_NOSOUND = &H10

Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Enum EVENT_Constants
  eMouseMove = WM_MOUSEMOVE
  eLeftMouseDown = WM_LBUTTONDOWN
  eLeftMouseUp = WM_LBUTTONUP
  eLeftMouseDblClick = WM_LBUTTONDBLCLK
  eMiddleMouseDown = WM_MBUTTONDOWN
  eMiddleMouseUp = WM_MBUTTONUP
  eMiddleMouseDblClick = WM_MBUTTONDBLCLK
  eRightMouseDown = WM_RBUTTONDOWN
  eRightMouseUp = WM_RBUTTONUP
  eRightMouseDblClick = WM_RBUTTONDBLCLK
  eMouseHover = WM_MOUSEHOVER
  eMouseLeave = WM_MOUSELEAVE
  eBalloonTipClicked = NIN_BALLOONUSERCLICK
  eBalloonTipTimedOut = NIN_BALLOONTIMEOUT
End Enum

Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function IsShellVersion(ByVal ShellVersion As Long) As Boolean
  Dim bSize As Long
  Dim nUnused As Long
  Dim lpBuffer As Long
  Dim nVerMajor As Integer
  Dim bBuffer() As Byte
  bSize = GetFileVersionInfoSize("shell32.dll", nUnused)
  If bSize > 0 Then
    ReDim bBuffer((bSize - 1)) As Byte
    GetFileVersionInfo "shell32.dll", 0&, bSize, bBuffer(0)
    If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
      CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
      IsShellVersion = (nVerMajor >= ShellVersion)
    End If
  End If
End Function
