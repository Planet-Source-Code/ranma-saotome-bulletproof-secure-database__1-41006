Attribute VB_Name = "modDesign"
Option Explicit

'PURE API MODULE DESIGN STUFF

' Public API declarations
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


' Public constants for GetWindowLong API declaration
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Public constants for window styles
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10

' Private constants for extended window styles
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

' Public constants for SetWindowPos API declaration
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

' Public constants for ShowWindow API declaration
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5


Private m_sPattern As String
Private m_lhFind As Long



Public Sub MakeFlat(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    RemoveBorder lhWnd
End Sub



Public Sub RemoveBorder(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
    'MakeFlat Text1.hwnd
End Sub
