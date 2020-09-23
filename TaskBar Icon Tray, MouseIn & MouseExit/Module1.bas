Attribute VB_Name = "Module1"
'herewith "China has wise codes" by ryanwei2005@gmail.com
'with this code you can make a simulation of windows 7 show desktop function easily now
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
   End Type

Type POINTAPI
   X As Long
   Y As Long
   End Type

Private Declare Function SetWindowLong Lib "User32" _
                Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
Private Declare Function GetWindowLong Lib "User32" _
                Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) _
                As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Long, _
                ByVal dwFlags As Long) _
                As Long
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA As Long = &H2
Const WS_EX_LAYERED As Long = &H80000
Const WS_EX_NOACTIVATE As Long = &H8000000

Private Declare Function GetVersion Lib "kernel32" () As Long

Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const TB_GETBUTTONSIZE = (WM_USER + 58)

Public IsWindows7 As Boolean
Public vRect As RECT
Public IconWidth As Integer, IconHeight As Integer

Public Function CheckMouseOut(ByVal nx As Long, ByVal ny As Long, ByVal ix As Long, ByVal iy As Long) As Boolean

  Dim Fhwnd As Long
  Dim BtnSize As Long
  Dim offset As Integer

If ix <> 0 And vRect.Top = 0 Then
Fhwnd = WindowFromPoint(ix, iy)
GetWindowRect Fhwnd, vRect
BtnSize = SendMessage(Fhwnd, TB_GETBUTTONSIZE, 0&, 0&)
IconWidth = BtnSize And &HFFFF&
IconHeight = (BtnSize / &H10000) And &HFFFF&
End If

If IsWindows7 Then offset = 2
' a little difference between newer 7, Vista and other OS

If IconWidth > 0 And IconHeight > 0 Then
    If ix <> 0 And Int((nx - vRect.Left - offset) / IconWidth) <> Int((ix - vRect.Left - offset) / IconWidth) Or _
    Int((ny - vRect.Top) / IconHeight) <> Int((iy - vRect.Top) / IconHeight) Then
        CheckMouseOut = True
        vRect.Top = 0
    End If

End If

End Function

Public Sub SetFormToAlpha(hWnd As Long, lngAlpha As Long)
    Dim tmpLog As Long
    
    If hWnd = 0 Then Exit Sub
    If lngAlpha >= 0 And lngAlpha <= 255 Then
        tmpLog = GetWindowLong(hWnd, GWL_EXSTYLE)
        Call SetWindowLong(hWnd, GWL_EXSTYLE, tmpLog Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hWnd, 0, lngAlpha, LWA_ALPHA)
    End If
End Sub

Public Function GetWinVersion() As String
    Dim lngRetval, lngMajor, lngMinor As Long
    
    lngRetval = GetVersion()
    lngMinor = (lngRetval And &HFF00) \ &H100
    lngMinor = lngMinor And &HFF
    lngMajor = (lngRetval And &HFF)
    If CBool((lngRetval And &H80000000) = 0&) Then
    If lngMajor = 6 And lngMinor = 1 Then GetWinVersion = "Windows 7"
    If lngMajor = 6 And lngMinor = 0 Then GetWinVersion = "Windows Vista"
    If lngMajor = 5 And lngMinor = 1 Then GetWinVersion = "Windows XP"
    If lngMajor = 5 And lngMinor = 0 Then GetWinVersion = "Windows 2000"

    ElseIf CBool(lngMajor = 4&) And CBool(lngMinor = 90&) Then
        'Windows ME
        GetWinVersion = "Windows ME"
    Else
    End If
End Function


