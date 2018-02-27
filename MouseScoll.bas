Attribute VB_Name = "MouseScoll"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type


'the goal of these functions is to takes different variables and distill them down to a single number
'some are locations of cursors, some are location of windows

Private Declare Function FindWindow Lib "user32" _
                                        Alias "FindWindowA" ( _
                                                        ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32.dll" _
                                        Alias "GetWindowLongA" ( _
                                                        ByVal hwnd As Long, _
                                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" _
                                        Alias "SetWindowsHookExA" ( _
                                                        ByVal idHook As Long, _
                                                        ByVal lpfn As Long, _
                                                        ByVal hmod As Long, _
                                                        ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" ( _
                                                        ByVal hHook As Long, _
                                                        ByVal nCode As Long, _
                                                        ByVal wParam As Long, _
                                                        lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                                        ByVal hHook As Long) As Long


Private Declare Function WindowFromPoint Lib "user32" ( _
                                                        ByVal xPoint As Long, _
                                                        ByVal yPoint As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                        ByRef lpPoint As POINTAPI) As Long

'set up the variables used for the mousewheel
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
Private mLngMouseHook As Long
Private mListBoxHwnd As Long

'set up variable to determine wheter mousewheel is hooked
Private mbHook As Boolean

'variable representing the current control form fed into the scrolling module, but specific to this module
Private mCtl As MSForms.Control
Dim n As Long

Sub HookListBoxScroll(frm As Object, ctl As MSForms.Control)
    Dim lngAppInst As Long
    Dim hwndUnderCursor As Long
    Dim tPT As POINTAPI
    
    'calls the get cursor position using the POINTAPI tPT
     GetCursorPos tPT
     
     'gets the current X,Y coordinate of the cursor as it relates to the whole window
     'windowfrompoint feeds in an X location, Y location and returns a single long
     hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
     
     'if the current control form under active control is not the one as fed into this subroutine
     If Not frm.ActiveControl Is ctl Then
            'set the focus on the control form
             ctl.SetFocus
     End If
     
     'if the current form being controlled is not the one under hand
     If mListBoxHwnd <> hwndUnderCursor Then
            'call the unhook sub routine
             UnhookListBoxScroll
             
             'set the form to be controlled to the one fed into this subroutine
             Set mCtl = ctl
             
             'set the location of "mListBoxHwnd" to be the current point of the cursor
             mListBoxHwnd = hwndUnderCursor
             
             'gets the current window long...
             lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
             ' PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
             'if the mouseball is not hooked then hook it
             If Not mbHook Then
                    'used to actually do the scrolling
                     mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                     mbHook = mLngMouseHook <> 0
             End If
     End If
End Sub


Sub UnhookListBoxScroll()
    'if the mouseball is hooked
     If mbHook Then
             'set the mouse control to nothing
             Set mCtl = Nothing
             
             'unhook the numbers too
             UnhookWindowsHookEx mLngMouseHook
             
             'set numbers to 0 and booleans to false
             mLngMouseHook = 0
             mListBoxHwnd = 0
             mbHook = False
        End If
End Sub

Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As Long
            Dim idx As Long
        On Error GoTo errH
     
     'if the mousewheel function is the 14th action of mousewheel.... then
     If (nCode = HC_ACTION) Then
             'if the window point is on the control box fed into this whole procedure....
             If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mListBoxHwnd Then
                     'if the parameter is mousewheel...
                     If wParam = WM_MOUSEWHEEL Then
                                'it is doing a mouse procedure
                                MouseProc = True
'                                If lParam.hwnd > 0 Then
'                                        PostMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
'                                Else
'                                        PostMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
'                                End If
'                                PostMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                                'go either up or down
                                If lParam.hwnd > 0 Then idx = -1 Else idx = 1
                             
                             'check the index in the function correspondingly
                             idx = idx + mCtl.ListIndex
                             If idx >= 0 Then mCtl.ListIndex = idx
                                Exit Function
                     End If
             Else
                    'if the cursor is not located in the services form
                     UnhookListBoxScroll
             End If
     End If
     'call the next hook
     MouseProc = CallNextHookEx( _
                             mLngMouseHook, nCode, wParam, ByVal lParam)
     Exit Function
errH:
     UnhookListBoxScroll
End Function

