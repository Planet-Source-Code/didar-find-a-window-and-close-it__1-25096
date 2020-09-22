Attribute VB_Name = "Module1"

Option Explicit

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_OWNER = 4
Const GWL_STYLE = -16
Const WS_DISABLED = &H8000000
Const WS_CANCELMODE = &H1F
Const WM_CLOSE = &H10

Private Declare Function GetWindow Lib "user32" ( _
    ByVal hwnd As Integer, _
    ByVal wCmd As Integer) As Integer
    
Private Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

Private Declare Function PostMessage _
    Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Declare Function GetWindowTextLength _
    Lib "user32" Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

Private Declare Function GetWindowText _
    Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

Private Declare Function IsWindow Lib "user32" ( _
    ByVal hwnd As Integer) As Integer


Sub CloseWindow(ByVal partialWindowCaption$)
    Dim Whnd&, L&, Nam$
    
    partialWindowCaption = LCase$(partialWindowCaption)
    Whnd = GetWindow(Form1.hwnd, GW_HWNDFIRST)
    
    Do While Whnd <> 0
        If IsWindow(Whnd) Then
            L = GetWindowTextLength(Whnd)
            If L > 0 Then
                Nam = Space$(L + 1)
                L = GetWindowText(Whnd, Nam, L + 1)
                Nam = LCase$(Left$(Nam, Len(Nam) - 1))
                If InStr(Nam, partialWindowCaption) Then
                    EndTask Whnd
                    Exit Do
                End If
            End If
        End If
        Whnd = GetWindow(Whnd, GW_HWNDNEXT)
        DoEvents
    Loop
End Sub



Sub EndTask(Whnd As Long)
    If Whnd = Form1.hwnd Or _
        GetWindow(Whnd, GW_OWNER) _
            = Form1.hwnd Then End
    
    If (GetWindowLong(Whnd, GWL_STYLE) _
        And WS_DISABLED) Then Exit Sub
    
    PostMessage Whnd, WS_CANCELMODE, 0, 0&
    PostMessage Whnd, WM_CLOSE, 0, 0&
End Sub




