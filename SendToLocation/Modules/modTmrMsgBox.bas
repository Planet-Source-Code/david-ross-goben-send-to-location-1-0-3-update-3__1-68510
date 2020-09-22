Attribute VB_Name = "modTmrMsgBox"
Option Explicit
'~modTmrMsgBox.bas;
'MsgBox that does not suspend running timers. Optional Timeout value available
'**********************************************************************
' modTmrMsgBox:
' The TmrMsgBox() function is like the MsgBox() function, except that
' timer control functions are NOT suspended, you can specify which form will
' act as the message's parent, and you can specify in the ResponseType
' parameter mbApplicationModal, where the user must respond to the message
' box before continuing work in the current application, or mbSystemModal,
' which will suspend ALL applications until the user responds. An ADDITIONAL
' feature is the ability to automatically close the MessageBox after a pre-
' determined number of seconds.
'
' ResponseType parameters are (can be OR'd to mix multiples):
'  vbOKOnly:           Display OK button only (default).      BUTTON
'  vbOKCancel:         Display OK, CANCEL buttons.            BUTTON
'  vbAbortRetryIgnore: Display ABORT, RETRY, IGNORE buttons.  BUTTON
'  vbYesNoCancel:      Display YES, NO, CANCEL buttons.       BUTTON
'  vbYesNo:            Display YES, NO buttons.               BUTTON
'  vbRetryCancel:      Display RETRY, CANCEL buttons.         BUTTON
'  vbCritical:         Display CRITICAL MESSAGE Icon.         ICON STYLE
'  vbQuestion:         Display WARNING QUERY Icon.            ICON STYLE
'  vbExclamation:      Display WARNING MESSAGE Icon.          ICON STYLE
'  vbInformation:      Display INFORMATION MESSAGE Icon.      ICON STYLE
'  vbDefaultButton1:   First button is default (default).     DEFAULT BUTTON
'  vbDefaultButton2:   Second button is default.              DEFAULT BUTTON
'  vbDefaultButton3:   Third button is default.               DEFAULT BUTTON
'  vbDefaultButton4:   Fourth button is default.              DEFAULT BUTTON
'  vbApplicationModal: Suspend application until user responds.     MODALITY
'  vbSystemModal       Suspend ALL application until user responds. MODALITY
'
' Returned values are:
' 1: OK button pressed.       vbOK
' 2: CANCEL button pressed.   vbCancel
' 3: ABORT button pressed.    vbAbort
' 4: RETRY button pressed.    vbRetry
' 5: IGNORE button pressed.   vbIgnore
' 6: YES button pressed.      vbYes
' 7: NO button pressed.       vbNo
'
'EXAMPLE;
'  Dim Resp As Integer
'  Resp = TmrMsgBox(Me.hwnd, "Normal Message", mbAbortRetryIgnore, "Normal Title")
'  Resp = TmrMsgBox(Me.hwnd, "Timed Message", mbAbortRetryIgnore, "Timed Title", 5)
'
' NOTE: Multiple lines are possible by inserting a vbLF or vbCrLf code between
'       lines in the message string.
' NOTE: If the dialog box displays a CANCEL button, pressing the ESC key
'       has the same effect as selecting CANCEL.
'**********************************************************************

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Const WM_QUIT As Long = &H12
Private Const PM_REMOVE As Long = &H1

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type Msg
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

Public Function TmrMsgBox(ByVal ParenthWnd As Long, message As String, _
                          Optional flags As VbMsgBoxStyle = vbOKOnly Or vbInformation, _
                          Optional Title As String = vbNullString, _
                          Optional SecondsCount As Long = 0) As VbMsgBoxResult
  Dim uiFlags As Long   'long version of flags
  Dim uiResult As Long  'result from MessageBox
  Dim dwTimeOut As Long 'timeout of seconds converted to miliseconds
  Dim idTimer As Long   'timer handle
  Dim WPMsg As Msg      'message structure
  Dim FocushWnd As Long 'window which has focus
'
' convert VbMsgBoxStyle flags to long
'
  uiFlags = flags
'
' convert seconds to miliseconds
'
  dwTimeOut = SecondsCount * 1000&
'
' if a timer value present, process MessageBox through a callback routine
'
  If SecondsCount Then
'
' get window which has focus
'
    FocushWnd = GetFocus()
'
' get handle to timer
'
    idTimer = SetTimer(0&, 0&, dwTimeOut, AddressOf MessageBoxTimer)
'
' display message, get result
'
    uiResult = MessageBox(ParenthWnd, message, Title, uiFlags)
'
' stop the timer
'
    KillTimer 0&, idTimer
'
' if message was WM_QUIT for the current app. Discard it if so
'
    If PeekMessage(WPMsg, 0&, WM_QUIT, WM_QUIT, PM_REMOVE) Then
      uiResult = 0                'indicate no button key selected
      EnableWindow ParenthWnd, 1&
    End If
'
' get focus back to where it was before
'
    If FocushWnd Then SetFocusAPI (FocushWnd)
  Else
    uiResult = MessageBox(ParenthWnd, message, Title, uiFlags)
  End If
'
' return result
'
  TmrMsgBox = uiResult

End Function

'***************************************
' go here only if message box times out
'***************************************
Private Sub MessageBoxTimer(hwnd As Long, uiMsg As Long, idEvent As Long, dwTime As Long)
  PostQuitMessage 0     'post a quit message for the current thread
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

