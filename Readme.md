# VB6 SysTray Control

A compact, stand-alone SysTray control for Visual Basic 6.0 that provides a small, reliable wrapper around the Windows shell notify icon functionality.
It was designed to be simple, fast and responsive and to be easy to drop into existing VB6 projects.

This repository bundles everything you need in one place:
- `SysTray.cls` — a class module that raises events and exposes the public API.
- `modSysTray.bas` — implementation code that interacts with the Windows API.
- `modGlobal.bas` — global declares and constants.
- `frmMain.frm` — an example form that demonstrates how to use the `SysTray` object.

## Requirements

- Visual Basic 6.0 (VB6). The project uses plain Win32 API calls — **GDI+ not supported.**
- Works on modern Windows versions; balloon/toast behavior may vary with OS policy.

## Quick overview

- Small, focused SysTray control implemented in VB6.
- Events for all mouse actions (including double-clicks and hover/leave).
- Simple properties to control the icon, tooltip and whether it is in the tray.
- Balloon tip support (default timeout: 15 seconds).

## Supported events

The `SysTray` class raises the following events:

- Event **MouseMove**()
- Event **LeftMouseDown**()
- Event **LeftMouseUp**()
- Event **LeftMouseDblClick**()
- Event **MiddleMouseDown**()
- Event **MiddleMouseUp**()
- Event **MiddleMouseDblClick**()
- Event **RightMouseDown**()
- Event **RightMouseUp**()
- Event **RightMouseDblClick**()
- Event **MouseHover**()
- Event **MouseLeave**()
- Event **BalloonTipClick**()
- Event **BalloonTipTimeout**()

Public properties (for user use)
- Property **Icon** (StdPicture)
  - Set this with `LoadPicture` or a `Picture` property from a PictureBox/Image.
  - The control converts the StdPicture to an HICON internally for the tray.
- Property **Tip** (String)
  - Tooltip text shown when hovering over the tray icon.
- Property **hWndParent** (Long)
  - The owner window handle. This MUST be set (for example, `Me.hWnd` from a form).
  - Set this before putting the icon into the tray.
- Property **InTray** (Boolean)
  - `True` to add the icon to the tray; `False` to remove it.

Internal properties (visible in the project but not for user modification)
- Property OldWindowProc
- Property TrayID
**Do NOT change these — they are required for correct operation.**

Public methods
- **ShowBalloonTip** (Title As String, Text As String)
  - Pops up a balloon tip (default timeout is 15 seconds on this control).
  - Note: actual appearance and behavior of notifications may differ on modern Windows (Action Center / Toasts).

Exposed internal function
- **DoThisEvent**()
  - Internal helper to cause the class to call a specific event. Not intended for regular user code.

## How to use it
Add the 3 files, **SysTray.cls*, **modSysTray.Bas**, **modGlobal.bas** to your project.

## Notes and Tips

Notes and tips
- **hWndParent** must be a valid window handle (Me.hWnd from a form or the handle of a hidden message window you create). If events do not fire, check hWndParent.
- Keep event handlers small and quick. If you need to perform long work, queue it or run it in a worker/background approach to keep UI responsive.
- Icon quality on high-DPI systems may vary; this implementation targets classic API behavior — if you want hi-DPI/alpha-blended icons, consider extending the project to use GDI+ (not included).
- Balloons/toasts: modern Windows may suppress or convert balloons into native notifications. Behavior is controlled by the shell and user settings.
- Multi-icon support: this class was designed around a single icon instance. If you need multiple tray icons, instantiate multiple SysTray objects.

## License

[MIT License](LICENSE)  
Copyright © Ubehage

---

## Credits

Created by Ubehage  

[GitHub Profile](https://github.com/Ubehage)
