VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   3480
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2640
      Top             =   1680
   End
   Begin VB.Menu Menu 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      Private Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uid As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
      Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Private Const WM_LBUTTONDOWN = &H201     'Button down
      Private Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Private Const WM_RBUTTONDOWN = &H204     'Button down
      Private Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
    
      Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long

Private Type MARGINS
m_Left As Long
m_Right As Long
m_Top As Long
   m_Button As Long
End Type

Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const MF_REMOVE = &H1000&
Private Const SC_CLOSE = &HF060
Private Const SC_MAXIMIZE = &HF030


Dim Inied As Boolean
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern void DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS margins);
Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, margin As MARGINS) As Long
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern bool DwmIsCompositionEnabled();
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

      
      Dim nid As NOTIFYICONDATA
      
      Dim PX As Long, PY As Long
      Dim lPos As POINTAPI


      Private Sub Form_Load()
      
          nid.cbSize = Len(nid)
         nid.hWnd = Form1.hWnd
         nid.uid = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = Form1.Icon
         nid.szTip = "RyanWei's Program" & vbCrLf & "China has wise codes" & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid
         If GetWinVersion = "Windows 7" Or GetWinVersion = "Windows Vista" Then IsWindows7 = True
         
         
      End Sub

      Private Sub Form_Terminate()
         'Delete the added icon from the taskbar status area when the
         'program ends.
         Shell_NotifyIcon NIM_DELETE, nid
      End Sub

      Private Sub Form_MouseMove _
         (Button As Integer, _
          Shift As Integer, _
          X As Single, _
          Y As Single)
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          Dim msg As Long
          Dim sFilter As String
          msg = X / Screen.TwipsPerPixelX
          
          Dim P As POINTAPI
                    
          Select Case msg
             Case WM_LBUTTONDOWN
               If Me.WindowState = 0 Then
               Me.WindowState = 1
               Else
               Me.WindowState = 0
               SetForegroundWindow Me.hWnd
               End If
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK
             
             Case WM_RBUTTONDOWN
                SetForegroundWindow Me.hWnd
                PopupMenu Menu
             
'                Dim ToolTipString As String
'                ToolTipString = InputBox("Enter the new ToolTip:", _
'                                  "Change ToolTip")
'                If ToolTipString <> "" Then
'                   nid.szTip = ToolTipString & vbNullChar
'                   Shell_NotifyIcon NIM_MODIFY, nid
'                End If
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
             
             Case WM_MOUSEMOVE
             GetCursorPos P
             PX = P.X     'remember the mouse position in taskbar icon tray
             PY = P.Y
             Timer1.Enabled = True
             
             
          End Select
      End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

Dim P As POINTAPI
Dim i As Long, j As Long
GetCursorPos P
i = P.X - PX 'zero means mouse position doesn't change, it's a mouse hover event
j = P.Y - PY
If i = 0 And j = 0 And lPos.X = 0 And lPos.Y = 0 Then
lPos.X = P.X ' as the mouse hovers, remember the location of tray icon
lPos.Y = P.Y

If IsWindows7 Then DrawAero
SetFormToAlpha hWnd, 120

Me.Caption = "MouseHover"
Timer1.Interval = 200
ElseIf CheckMouseOut(P.X, P.Y, lPos.X, lPos.Y) = True Then
lPos.X = 0
lPos.Y = 0

If IsWindows7 Then DisableAero
SetFormToAlpha hWnd, 255

Me.Caption = "MouseExit"
Timer1.Interval = 800
Timer1.Enabled = False
End If


End Sub


Private Sub DrawAero()
Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long

    hBrush = CreateSolidBrush(RGB(0, 0, 0))
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hWnd, m_Rect
    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld
    DeleteObject hBrush
    
    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
     
    DwmIsCompositionEnabled en
    If en Then
        
        DwmExtendFrameIntoClientArea Me.hWnd, mg

    End If
    
    BorderStyle = 0
End Sub

Private Sub DisableAero()
Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long

    hBrush = CreateSolidBrush(RGB(240, 240, 240))
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hWnd, m_Rect

    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld

    DeleteObject hBrush
    
    Dim mg As MARGINS, en As Long
    mg.m_Left = 0
    mg.m_Button = 0
    mg.m_Right = 0
    mg.m_Top = 0
     
    DwmIsCompositionEnabled en
    If en Then
        
        DwmExtendFrameIntoClientArea Me.hWnd, mg
       
    End If
End Sub
