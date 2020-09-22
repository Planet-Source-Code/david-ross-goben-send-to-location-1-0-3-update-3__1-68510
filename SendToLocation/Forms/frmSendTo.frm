VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "Msmapi32.ocx"
Begin VB.Form frmSendTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send To Other Location..."
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "frmSendTo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstFilesFull 
      Height          =   1035
      ItemData        =   "frmSendTo.frx":08CA
      Left            =   2280
      List            =   "frmSendTo.frx":08CC
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "frmSendTo.frx":08CE
      Left            =   120
      List            =   "frmSendTo.frx":08D8
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   420
      Width           =   7095
   End
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   8100
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   7500
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cmdAdrBk 
      Caption         =   "Address Book..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      ToolTipText     =   "Browse Address Book"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   4560
      Width           =   4695
   End
   Begin VB.CheckBox chkEmail 
      Caption         =   "Attach to &Email Recipeint(s):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Place a shortcut to the selected item(s) instead of copying/moving them (separate recipients with semicolon ';')"
      Top             =   4560
      Width           =   2715
   End
   Begin VB.CheckBox chkOverWrite 
      Caption         =   "&Prompt if destination already exists"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Prompt user for over-writes if files/folders already exist"
      Top             =   3780
      Width           =   3015
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   15
      Top             =   4200
      Width           =   2595
   End
   Begin VB.CheckBox chkUse 
      Caption         =   "&Use this title:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "Include ""Shortcut to"" at the beginning of the filename"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "&Add ""Shortcut to"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1740
      TabIndex        =   13
      ToolTipText     =   "Include ""Shortcut to"" at the beginning of the filename"
      Top             =   4200
      Width           =   1755
   End
   Begin VB.CheckBox chkShortcut 
      Caption         =   "Send as &Shortcut"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Place a shortcut to the selected item(s) instead of copying/moving file(s)"
      Top             =   4200
      Width           =   1635
   End
   Begin VB.Timer tmrFader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   1020
   End
   Begin VB.TextBox txtFocus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7620
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "txtFocus"
      Top             =   1020
      Width           =   915
   End
   Begin VB.OptionButton optCopyMove 
      Caption         =   "&Move"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6540
      TabIndex        =   8
      ToolTipText     =   "Move files/folders"
      Top             =   3180
      Width           =   795
   End
   Begin VB.OptionButton optCopyMove 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5820
      TabIndex        =   7
      ToolTipText     =   "Copy files/folders"
      Top             =   3180
      Width           =   795
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "Browse for a destination path"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear &History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      ToolTipText     =   "Remove all items from the TO list"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      ToolTipText     =   "Cancel selections"
      Top             =   540
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      ToolTipText     =   "Accept selections and process them"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboHistory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3420
      Width           =   7095
   End
   Begin VB.CheckBox chkSpecial 
      Caption         =   "Include Special System Folders in Send To List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   23
      ToolTipText     =   "Include special folders such as Desktop, QuickLaunch, etc."
      Top             =   3780
      Width           =   3555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   7440
      X2              =   7440
      Y1              =   60
      Y2              =   4860
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   7200
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(DEL or Drag && Drop OK)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   5460
      TabIndex        =   19
      Top             =   180
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   8040
      Picture         =   "frmSendTo.frx":0963
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblSize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3060
      TabIndex        =   21
      Top             =   180
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send &To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3180
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Files && Folders:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   1755
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   180
      X2              =   7200
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   7440
      X2              =   7440
      Y1              =   60
      Y2              =   4860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Operation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   6
      Top             =   3180
      Width           =   900
   End
End
Attribute VB_Name = "frmSendTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SendPath As String         'selected destination path
Public SrcPath As String          'set source path for file or files
Public coLSendTo As Collection    'collection for checking item pre-existence
Public Fso As FileSystemObject
Public FormLoaded As Boolean      'True when form is loaded

'*******************************************************************************
' Subroutine Name   : chkSpecial_Click
' Purpose           : Toggle Include Special Folders Option
'*******************************************************************************
Private Sub chkSpecial_Click()
  Dim Idx As Integer, I As Integer
  Dim S As String
  
  SaveSetting App.Title, "SendTo", "InclSpecial", CStr(Me.chkSpecial.Value)
  With Me.cboHistory
    If Me.chkSpecial.Value = vbChecked Then                 'add new entries if checked
      On Error Resume Next
      S = "Desktop"                                         'apply Desktop
      coLSendTo.Add S, S
      Me.cboHistory.AddItem S                               'add to combo
      S = "Favorites"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "My Documents"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "Quick Launch"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "Send To"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "Start Menu"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "Startup"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      S = "Templates"
      coLSendTo.Add S
      Me.cboHistory.AddItem S
      On Error GoTo 0
    Else                                                      'remove specials from lists
      For Idx = .ListCount - 1 To 0 Step -1
        S = .List(Idx)                                        'grab an item
        I = InStr(1, S, "\")
        If CBool(I) Then S = Left$(S, I - 1)                  'strip any backslash
        Select Case LCase$(S)                                 'check for special folders
          Case "desktop"
          Case "favorites"
          Case "my documents"
          Case "quick launch"
          Case "send to"
          Case "start menu"
          Case "startup"
          Case "templates"
          Case Else
            S = vbNullString                                  'no special folder
        End Select
        If CBool(Len(S)) Then                                 'if special folder found...
          .RemoveItem Idx                                     'then remove
        End If
      Next Idx
      
      With coLSendTo                                          'now empty collection
        Do While .Count
          .Remove 1
        Loop
      End With
      For Idx = 0 To .ListCount - 1                           'and repolpulate it with existing items
        S = .List(Idx)
        coLSendTo.Add S, S
      Next Idx
    End If
    Me.cmdOK.Enabled = CBool(Len(Trim$(.Text))) And CBool(Me.lstFiles.ListCount)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up XP buttons
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Load the form, load the history list
'*******************************************************************************
Private Sub Form_Load()
  Dim Count As Long, Idx As Long, I As Long
  Dim S As String, Keys As String
  Dim Ary() As String               'list of files to copy/move
  Dim Cmd As String
  Dim InQuote As Boolean
  
  Me.Caption = Me.Caption & " (Version " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision) & ")"
  Set Fso = New FileSystemObject                                  'create file I/O object
  
  Cmd = Command$                                                  'get command line argument
  Call CheckSendTo(Cmd)                                           'create shortcut as needed
  If StrComp(Cmd, "/MakeShortcut", vbTextCompare) = 0 Then        'if we are just creating shortcut
    Cmd = vbNullString
  End If
  
  Me.chkInclude.Value = CInt(GetSetting(App.Title, "Settings", "ShortcutTo", "1"))
  Me.chkOverWrite.Value = CInt(GetSetting(App.Title, "Settings", "OverWritePmt", "1"))
  
  Me.lstFiles.Clear                                               'erase developer help note
  Me.lstFilesFull.Clear
  
  If CBool(Len(Cmd)) Then                                         'if data to copy
    InQuote = False                                               'init flag for not inside quoted text
    For Idx = 1 To Len(Cmd)                                       'scan each character
      If Mid$(Cmd, Idx, 1) = Chr$(34) Then InQuote = Not InQuote  'flip flag on quote encounteres
      If Mid$(Cmd, Idx, 1) = " " Then                             'space?
        If Not InQuote Then                                       'yes, and not in quote?
          Mid$(Cmd, Idx, 1) = Chr$(0)                             'right, so we can create separator
        End If
      End If
    Next Idx
    
    Ary = Split(Cmd, Chr$(0))                                     'now set aside list array
    If InStr(1, Cmd, Chr$(0)) Then                                'did we make an array?
      S = Ary(0)                                                  'yes, so get first entry
      I = InStrRev(Ary(0), "\")                                   'find base path
      SrcPath = Left$(Ary(0), I)                                  'and save it
      Count = UBound(Ary)                                         'set number to do, offset 0
      Cmd = vbNullString                                          'init accumulator
      For Idx = 0 To Count
        S = Ary(Idx)                                              'get a string
        If Left$(S, 1) = """" Then S = Mid$(S, 2, Len(S) - 2)     'strip quotes
        Me.lstFilesFull.AddItem S                                 'add full path
        S = Mid$(S, InStrRev(S, "\") + 1)                         'strip base path
        Me.lstFiles.AddItem S                                     'add to display list
        Cmd = Cmd & S & vbCrLf                                    'accumulate list
      Next Idx
    Else
      If Left$(Cmd, 1) = """" Then Cmd = Mid$(Cmd, 2, Len(Cmd) - 2)
      Me.lstFilesFull.AddItem Cmd
      I = InStrRev(Cmd, "\")
      SrcPath = Left$(Cmd, I)                                     'save base path
      Cmd = Mid$(Cmd, I + 1)                                      'and grab object to process
      Me.lstFiles.AddItem Cmd
    End If
    Me.lstFiles.ListIndex = -1                                    'hide selection box
  Else
    TmrMsgBox Me.hwnd, "No files selected. Exiting " & App.Title, vbOKOnly Or vbInformation, "No Files Selected"
    Unload Me
    Exit Sub
  End If
'
' now set up display
'
  Me.txtFocus.Left = -2880                                    'hide focus box
  Me.lblSize.Top = -2880
  Set coLSendTo = New Collection                              'create duplicity col.
  SendPath = vbNullString                                     'init result to nada
  Me.optCopyMove(0).Value = True                              'ensure Copy option set
  Count = CLng(GetSetting(App.Title, "SendTo", "Count", "0")) 'get history count
  With Me.cboHistory
    .Clear                                                    'clear combo box
    If Count Then                                             'somthing to do?
      On Error Resume Next                                    'trap possible collection errors
      For Idx = 1 To Count                                    'yes, fill it
        S = Trim$(GetSetting(App.Title, "SendTo", "Path" & CStr(Idx), vbNullString))
        If Len(S) Then
          Err.Clear
          coLSendTo.Add S, UCase$(S)                          'add to collection
          If Err.Number = 0 Then .AddItem S                   'add to list if no errors
        End If
      Next Idx
      On Error GoTo 0
'
' see if we want to include special folders
'
      Me.chkSpecial.Value = CInt(GetSetting(App.Title, "SendTo", "InclSpecial", "0"))
'
' get last destination selected
'
      S = GetSetting(App.Title, "SendTo", "LastSel", vbNullString)
      If CBool(Len(S)) Then
      For Idx = 0 To .ListCount - 1
        If StrComp(S, .List(Idx), vbTextCompare) = 0 Then
          .ListIndex = Idx                                    'set as default
          Exit For
        End If
      Next Idx
      End If
    End If
    Me.cmdClear.Enabled = CBool(Count)
'
' if anything in list, place the topmost one in the combo textbox
'
    Me.cmdClear.Enabled = CBool(.ListCount)
    If .ListCount Then                                        'if something there
      Me.cboHistory.Text = Me.cboHistory.List(0)              'show topmost item
    Else
      Me.cmdOK.Enabled = False
    End If
  End With
  FormLoaded = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : If user hit window's X button
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Or UnloadMode = 1 Then
    Me.cmdCancel.Value = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Close out form. Save target list
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Long, Count As Long
  
  With Me.cboHistory
'
' remove any special folders
'
    For Idx = .ListCount - 1 To 0 Step -1
      If InStr(1, .List(Idx), "\") = 0 Then 'special folders do not have a backslash
        .RemoveItem Idx
      End If
    Next Idx
'
' save the current history list
'
    Count = .ListCount
    If CBool(Count) Then
      SaveSetting App.Title, "SendTo", "Count", CStr(Count)
      For Idx = 0 To .ListCount - 1
        SaveSetting App.Title, "SendTo", "Path" & CStr(Idx + 1), .List(Idx)
      Next Idx
    End If
  End With
'
' remove resources
'
  Set coLSendTo = Nothing
  Set Fso = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : check for user deleting selection(s)
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Idx As Integer
  
  If KeyCode = 46 Then                            'DEL
    With Me.lstFiles
      If CBool(.SelCount) Then                    'if anything selected
        For Idx = .ListCount - 1 To 0 Step -1     'scan backward through list
          If .Selected(Idx) Then
            .RemoveItem Idx                       'remove selected items
            Me.lstFilesFull.RemoveItem Idx        'also delete from full-path list
          End If
        Next Idx
        .ListIndex = -1                           'hide selection box
        Me.cmdOK.Enabled = CBool(Len(Trim$(Me.cboHistory.Text))) And CBool(.ListCount)
      End If
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkEmail_Click
' Purpose           : Toggle email reciient
'*******************************************************************************
Private Sub chkEmail_Click()
  Dim Bol As Boolean
  
  Bol = Me.chkEmail.Value = vbUnchecked
  Me.txtEmail.Enabled = Not Bol
  If Bol Then
    Me.cboHistory.BackColor = &H80000005
    Me.txtEmail.BackColor = &H80000016
  Else
    Me.cboHistory.BackColor = &H80000016
    Me.chkShortcut.Value = vbUnchecked
    Me.txtEmail.BackColor = &H80000005
    Me.txtEmail.SetFocus
  End If
  Me.optCopyMove(0).Enabled = Bol         'disable Copy/Move if checked
  Me.optCopyMove(1).Enabled = Bol
  Me.chkOverWrite.Enabled = Bol
  Me.cboHistory.Enabled = Bol
  Me.cmdAdrBk.Enabled = Not Bol
  chkUse_Click
End Sub

'*******************************************************************************
' Subroutine Name   : chkOverWrite_Click
' Purpose           : Toggle over-writing prompt
'*******************************************************************************
Private Sub chkOverWrite_Click()
  If FormLoaded Then
    SaveSetting App.Title, "Settings", "OverWritePmt", CStr(Me.chkOverWrite.Value)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAdrBk_Click
' Purpose           : Open the user's address book
'*******************************************************************************
Private Sub cmdAdrBk_Click()
  Dim Idx As Integer, Cnt As Integer
  Dim S As String
  
  MAPISession.DownLoadMail = False      'init session
  MAPISession.SignOn
'
' show address book
'
  With MAPIMessages
    .SessionID = MAPISession.SessionID
    .Compose
    .MsgIndex = -1                      'allow assigning properties
    On Error Resume Next
    .Show False
'
' get address(es)
'
    If Not CBool(Err.Number) Then       'if OK
      Cnt = .RecipCount                 'get count of recipients
      If CBool(Cnt) Then                'if something selected
        S = vbNullString                'init accumulator
        For Idx = Cnt - 1 To 0 Step -1  'pass through list in reverse (entered that way)
          .RecipIndex = Idx             'set an index to a recipient
          S = S & "; " & .RecipAddress  'accumulate addressees
        Next Idx
        Me.txtEmail.Text = Mid$(S, 3)   'stuff result
      End If
    End If
  End With
  
  MAPISession.SignOff                   'close out session
  On Error GoTo 0
End Sub

'*******************************************************************************
' Subroutine Name   : chkInclude_Click
' Purpose           : Save Setting of "Shortcut to" prepend option
'*******************************************************************************
Private Sub chkInclude_Click()
  If FormLoaded Then
    SaveSetting App.Title, "Settings", "ShortcutTo", CStr(Me.chkInclude.Value)
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkShortcut_Click
' Purpose           : Toggle Send ShortCut option
'*******************************************************************************
Private Sub chkShortcut_Click()
  Dim Bol As Boolean
  
  Bol = Me.chkShortcut.Value = vbChecked  'set Bol=True if the checkbox is checked
  If Bol Then
    Bol = Me.lstFiles.ListCount = 1       'enable other items only if 1 file
  End If
  Me.chkInclude.Enabled = Bol             'enable other shortcut checkboxes
  Me.chkUse.Enabled = Bol
  Me.txtTitle.Enabled = Bol And Me.chkUse.Value = vbChecked
  If Bol Then
    Me.chkEmail.Value = vbUnchecked       'ensure email unchecked if this option selected
  End If
  Bol = Me.chkShortcut.Value = Unchecked
  Me.optCopyMove(0).Enabled = Bol         'disable Copy/Move if checked
  Me.optCopyMove(1).Enabled = Bol
  Me.chkOverWrite.Enabled = Bol
  Call chkUse_Click
End Sub

'*******************************************************************************
' Subroutine Name   : chkUse_Click
' Purpose           : Allow user to change title of shortcuts
'*******************************************************************************
Private Sub chkUse_Click()
  Dim S As String
  Dim I As Integer
  
  With Me.txtTitle
    .Enabled = Me.chkUse.Value = vbChecked And Me.chkUse.Enabled
    If .Enabled Then
      .BackColor = &H80000005
      If Len(.Text) = 0 Then
        S = Me.lstFiles.List(0)
        I = InStrRev(S, ".")
        If I = 0 Then I = Len(S) + 1
        .Text = Left$(S, I - 1)
      End If
      .SetFocus
    Else
      .BackColor = &H80000016
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cboHistory_Change
' Purpose           : When the textbox for the combo changes, enable OK button
'                   : as required.
'*******************************************************************************
Private Sub cboHistory_Change()
  Me.cmdOK.Enabled = CBool(Len(Trim$(Me.cboHistory.Text))) And CBool(Me.lstFiles.ListCount)
  Me.lblSize.Caption = Me.cboHistory.Text
  If Me.lblSize.Width > Me.cboHistory.Width - 375 Then
    Me.cboHistory.ToolTipText = Me.cboHistory.Text
  Else
    Me.cboHistory.ToolTipText = vbNullString
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cboHistory_Click
' Purpose           : When an item in the combo box is selected, make sure it is
'                   : moved to the top of the combo list, and is displayed in the
'                   : combo's text box
'*******************************************************************************
Private Sub cboHistory_Click()
  Call CheckList              'verify list order
  With Me.cboHistory
    .ListIndex = 0            'ensure if moved, reselected
    .Text = .List(0)          'grab top of list for display
    Me.cmdOK.Enabled = CBool(Len(Trim$(Me.cboHistory.Text))) And CBool(Me.lstFiles.ListCount)
    SaveSetting App.Title, "SendTo", "LastSel", .List(.ListIndex)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowse_Click
' Purpose           : Browse for a destination location
'*******************************************************************************
Private Sub cmdBrowse_Click()
  Dim Path As String
  
  Path = DirBrowser(Me.hwnd, ViewDirsOnly, "Select Destination Folder")
  If CBool(Len(Path)) Then
    Me.cboHistory.Text = Path
    SaveSetting App.Title, "SendTo", "LastSel", Path
    Call CheckList
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Abandon ship...
'*******************************************************************************
Private Sub cmdCancel_Click()
  SendPath = vbNullString       'cancelling, so nothing to return
  On Error Resume Next
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Purpose           : Erase the history list
'*******************************************************************************
Private Sub cmdClear_Click()
  Me.cboHistory.Clear                   'erase combo data
  Me.cboHistory.Text = vbNullString
  With coLSendTo                        'erase the contents of the collection
    Do While .Count
      .Remove 1
    Loop
  End With
  SaveSetting App.Title, "SendTo", "LastSel", vbNullString
  Me.cmdOK.Enabled = False
  On Error Resume Next
  DeleteSetting App.Title, "SendTo"     'remove the history from the registry
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Try to accept selections
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim Path As String, S As String, T As String
  Dim I As Integer, J As Integer
  Dim sc As cShortcut
  Dim Bol As Boolean
  
  Call CheckList                              'see if user typed or modified the destination path
'
' Emailing?
'
  If Me.chkEmail.Value = vbChecked Then
    S = vbNullString
    With Me.lstFilesFull
      For I = 0 To .ListCount - 1             'check each entry to process
        T = .List(I)
        If GetAttr(T) And vbDirectory Then
          TmrMsgBox Me.hwnd, "You cannot email a folder specification. This often results in huge" & vbCrLf & _
                             "overhead. Consider Zipping the folder and emailing the ZIP file.", _
                             vbOKOnly Or vbExclamation, "Cannot Email Folders"
          Exit Sub
        End If
        S = S & "; " & .List(I)               'grab an item
      Next I
    End With
    If SendEMail(Trim$(Me.txtEmail.Text), "Sending Attachments", "See Attachments", Mid$(S, 3)) Then
      Unload Me                               'unload app
    End If
    Exit Sub
  End If
'
' get the destination path
'
  Path = Trim$(Me.cboHistory.Text)            'get path selection
'
' expand as required in case a special folder specified
'
  I = InStr(1, Path, "\")                     'check for path spec (user could have added on)
  If I Then
    S = Left$(Path, I - 1)                    'grab all before backslash
  Else
    S = Path                                  'else grab all
  End If
  
  Select Case LCase$(S)                       'check for special folders
    Case "desktop"
    Case "favorites"
    Case "my documents"
      S = "mydocuments"
    Case "quick launch"
      S = "quicklaunch"                       'strip spaces
    Case "send to"
      S = "sendto"
    Case "start menu"
      S = "startmenu"
    Case "startup"
    Case "templates"
    Case Else
      S = vbNullString                        'no special folder
  End Select
  
  If CBool(Len(S)) Then                       'if special folder specified...
    S = GetSpecialFolder(S)                   'get path to it
    If CBool(I) Then                          'if user had added on specs
      Path = S & Mid$(Path, I)                'append to new path
    Else
      Path = S                                'else simply grab new path
    End If
  End If
'
' Ensure that the destination path exists
'
  On Error Resume Next
  I = 0
  Do While I = 0
    I = Len(Dir$(Path, vbDirectory Or vbHidden Or vbSystem))  'path exists?
    If Err.Number Then I = 0                  'no
    On Error GoTo 0
    If I = 0 Then                           'oopsage message
      If TmrMsgBox(Me.hwnd, "The path '" & Path & "' does not exist. Retry?", vbRetryCancel Or vbQuestion, "Path Does Not Exist") = vbCancel Then
        Path = vbNullString                   'cancelling
        Exit Do
      End If
    End If
  Loop
  
  Path = AddSlash(Path)                       'ensure Path has trailing backslash
  Set sc = New cShortcut                      'creat shortcut object
  
  For I = 0 To Me.lstFilesFull.ListCount - 1  'check each entry to process
    S = Me.lstFilesFull.List(I)               'grab an item
    J = InStrRev(S, "\") + 1                  'find target filename/foldername
    T = Mid$(S, J)                            'save it to T var
    If Me.chkShortcut.Value = vbChecked Then  'if sending shortcut
      If Me.chkUse.Value = vbChecked And CBool(Len(Trim$(Me.txtTitle.Text))) Then
        T = Trim$(Me.txtTitle.Text)           'use substitution title for shortcut
      End If
      If Me.chkInclude.Value = vbChecked Then
        T = "Shortcut to " & T                'prepend "Shortcut to
      End If
      Call sc.CreateShortcutAt(Path, T, S)    'create shortcut
    Else
      If Fso.FolderExists(Path & T) Or Fso.FileExists(Path & T) Then
        Select Case TmrMsgBox(Me.hwnd, "'" & T & "' already exists at the destination path." & vbCrLf & vbCrLf & _
                  "Over-write it?", vbYesNoCancel Or vbQuestion, "Over-Write Prompt")
          Case vbYes
            Bol = True
          Case vbNo
            Bol = False
          Case vbCancel
            Set sc = Nothing                  'remove allocated resources
            Exit Sub
        End Select
      Else
        Bol = True
      End If
      
      If Bol Then
        If Me.optCopyMove(0).Value Then       'copying
          If Fso.FolderExists(S) Then         'if source is folder
            Fso.CopyFolder S, Path & T, True  'copy folder
          Else
            Fso.CopyFile S, Path, True        'else copy file
          End If
        Else                                  'else moving
          On Error Resume Next
          If Fso.FolderExists(S) Then         'source exists
            Fso.CopyFolder S, Path & T, True  'copy folder
            If Not CBool(Err.Number) Then
              Fso.DeleteFolder S              'delete source folder if no error
            End If
          Else
            Fso.CopyFile S, Path, True        'do it this way, because MoveFile does not allow over-write
            If Not CBool(Err.Number) Then
              Fso.DeleteFile S, True          'delete source file if no error
            End If
          End If
          On Error GoTo 0
        End If
      End If
    End If
  Next I
  Set sc = Nothing                          'remove allocated resources
  Unload Me                                 'unload app
End Sub

'*******************************************************************************
' Subroutine Name   : lstFiles_MouseMove
' Purpose           : Display tooltip as item mouse is over
'*******************************************************************************
Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Idx As Long
  Dim S As String
  
  With Me.lstFiles
  Idx = ListItemByCoordinate(Me.lstFiles, X, Y)
    If Idx = -1 Then
      S = vbNullString
    Else
      S = Me.lstFilesFull.List(Idx)
    End If
    If .ToolTipText <> S Then .ToolTipText = S
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstFiles_OLEDragDrop
' Purpose           : Drop additional files onto list
'*******************************************************************************
Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Idx As Long, I As Long
  Dim iIdx As Integer
  Dim S As String
  Dim col As Collection
'
' build list of current files
'
  Set col = New Collection
  With Me.lstFilesFull
    For iIdx = 0 To .ListCount - 1
      col.Add .List(iIdx), .List(iIdx)
    Next iIdx
  End With
'
' now parse for files to add
'
  For Idx = 1 To Data.Files.Count
    S = Data.Files(Idx)                                     'grab path path
    If Left$(S, 1) = """" Then S = Mid$(S, 2, Len(S) - 2)   'presently not needed, but who knows?
    On Error Resume Next
    col.Add S, S                                            'add to collection
    If Not CBool(Err.Number) Then                           'not already there?
      Me.lstFilesFull.AddItem S                             'no, so add full path
      I = InStrRev(S, "\")
      Me.lstFiles.AddItem Mid$(S, I + 1)                    'add just filename to display
    End If
  Next Idx
  
  Set col = Nothing                                         'done with object
End Sub

'*******************************************************************************
' Subroutine Name   : optCopyMove_Click
' Purpose           : Select option Copy or Move
'*******************************************************************************
Private Sub optCopyMove_Click(Index As Integer)
  Me.optCopyMove(0).FontBold = False
  Me.optCopyMove(1).FontBold = False
  Me.optCopyMove(Index).FontBold = True
End Sub

'*******************************************************************************
' Subroutine Name   : txtEmail_GotFocus
' Purpose           : Select all text when it gets focus
'*******************************************************************************
Private Sub txtEmail_GotFocus()
  With Me.txtEmail
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtEmail_Change
' Purpose           : Duplicate the text contents to the tooltip
'*******************************************************************************
Private Sub txtEmail_Change()
  Me.txtEmail.ToolTipText = Me.txtEmail.Text
End Sub

'*******************************************************************************
' Subroutine Name   : txtFiles_GotFocus
' Purpose           : When the textbox gets focus, hide focus by diverting
'                   : focus to the hidden locked textbox
'*******************************************************************************
Private Sub txtFiles_GotFocus()
  Me.txtFocus.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : CheckList
' Purpose           : Adjust the selected entry in the combo list, always ensuring that
'                   : the most recent entry is top-most
'*******************************************************************************
Private Sub CheckList()
  Dim Idx As Long
  Dim Txt As String
  Dim NewTxt As String
  
  With Me.cboHistory
    NewTxt = Trim$(.Text)                         'master plate of selection
    Txt = UCase$(NewTxt)                          'uppercase version
    With coLSendTo
      On Error Resume Next                        'trap duplication errors
      .Add Txt, Txt                               'add to collection
      Idx = Err.Number                            'set non-zero if already in col.
      On Error GoTo 0
    End With
'
' if already in the colleciton, find the entry in the combo list and remove it
'
    If CBool(Idx) Then
      For Idx = 0 To .ListCount - 1
        If Txt = UCase$(.List(Idx)) Then Exit For
      Next Idx
      If Idx = 0 Then Exit Sub                    'nothing more if already at top
      .RemoveItem Idx
    End If
'
' set new selection to top of list
'
    .AddItem NewTxt, 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : CheckSendTo
' Purpose           : Save Shortcut to SendTo Folder
'*******************************************************************************
Private Sub CheckSendTo(Txt As String)
 Dim SendTo As String
  Dim sc As cShortcut
'
' if any command line parameter set, do not check if the shortcut has already been made
'
  If CBool(Len(Txt)) Then
    If CBool(StrComp(Txt, "/MakeShortcut", vbTextCompare)) Then  'if not forcing creation...
      If GetSetting(App.Title, "Settings", "MadeSendTo", "0") = "1" Then Exit Sub 'see if created
    End If
  End If
  
  SendTo = GetSpecialFolder("SendTo")   'get full path to the user's SENDTO folder
  Set sc = New cShortcut                'create shortcut object
  On Error Resume Next
  Call sc.CreateShortcutAt(SendTo, "Send To Location...", AddSlash(App.Path) & App.EXEName & ".exe")
  If CBool(Err.Number) Then                    'd'oh!
    Set sc = Nothing
    TmrMsgBox Me.hwnd, "Cannot save a shortcut to this program to your SendTo folder.", vbOKOnly Or vbExclamation, "Error"
    Exit Sub
  End If
  Set sc = Nothing                      'all is well, so remove created object
  If GetSetting(App.Title, "Settings", "MadeSendTo", "0") = "0" Then
    TmrMsgBox Me.hwnd, "Saved a shortcut to ""Send To Location..."" to your System's SendTo folder.", vbOKOnly Or vbInformation, "ShortCut Saved to SendTo Folder"
  End If
  SaveSetting App.Title, "Settings", "MadeSendTo", "1"  'indicate item saved
End Sub

'*******************************************************************************
' Subroutine Name   : txtTitle_GotFocus
' Purpose           : Select all when textbox selected
'*******************************************************************************
Private Sub txtTitle_GotFocus()
  With Me.txtTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*************************************************
' SendEMail(): Send composed email message to someone. An Email dialog box is
'              opened up with the provided fields filled in. The user need only
'              press the SEND button.
'
' Parameter values:
'   Addresses:      Email address of addressee
'   Subject:        The subject of the message
'   Message:        Body of text
'   Attachments:    Optional filepath to a file to attach to the message
'
'EXAMPLE:
'   SendEmail Me, "www.merrymech.com", "Help Me!", "Where is the ANY KEY?"
'
'*************************************************
Public Function SendEMail(Addresses As String, _
                     Subject As String, message As String, _
                     Optional Attachments As String = vbNullString) As Boolean
  Dim Ary() As String, S As String
  Dim Cnt As Integer, Idx As Integer
  
  SendEMail = True
  MAPISession.DownLoadMail = False                'init session
  MAPISession.SignOn
  With MAPIMessages
    .SessionID = MAPISession.SessionID
    .Compose
    .MsgIndex = -1                                'allow assigning properties
    
    If CBool(Len(Addresses)) Then                 'addrersses specified?
      Ary = Split(Addresses, ";")                 'yes, break up
      Cnt = UBound(Ary)                           'get count-1
      For Idx = 0 To Cnt
        .RecipIndex = Idx                         'set current address
        .RecipAddress = Trim$(Ary(Idx))           'set recipient address
      Next Idx
    Else
      .RecipAddress = vbNullString
    End If
    
    .MsgSubject = Subject                         'set subject
    If CBool(Len(message)) Then
      .MsgNoteText = message                      'set message
    Else
      .MsgNoteText = " "
    End If
    
    If CBool(Len(Attachments)) Then               'attachments specified?
      Ary = Split(Attachments, ";")               'break up
      Cnt = UBound(Ary)                           'get count-1
      For Idx = 0 To Cnt
        .AttachmentIndex = Idx                    'set current
        .AttachmentPathName = Trim$(Ary(Idx))     'appy
      Next Idx
    Else
      .AttachmentPathName = vbNullString
    End If
    
    On Error Resume Next
    .Send True                                    'display dialog
  End With
  
  Select Case Err.Number
    Case 0              'no error
    Case 32001          'user abort
      SendEMail = False
    Case Else           'all other errors
      Err.Clear
      On Error GoTo MailError
'
' create Microsoft Outlook object in case the error was caused by user using MS Outlook
'
      Dim mailProg As Object
      Dim mailmsg As Object
      Set mailProg = CreateObject("Outlook.Application") 'late-bound Outlook object
      Set mailmsg = mailProg.CreateItem(0)               '0=olMailItem
      With mailmsg
        .To = Addresses                           'set up OUTLOOK eMail interface
        .Subject = Subject
        .BODY = message
        
        If CBool(Len(Attachments)) Then               'attachments specified?
          Ary = Split(Attachments, ";")               'break up
          Cnt = UBound(Ary)                           'get count-1
          For Idx = 0 To Cnt
            .Attachments.Add Trim$(Ary(Idx))          'appy attachment
          Next Idx
        Else
          .AttachmentPathName = vbNullString
        End If
        
        
        
        .Attachments.Add Attachments
        .Display                                  'display eMail interface. Allow user to hit SEND
      End With
      Set mailmsg = Nothing
      Set mailProg = Nothing
  End Select
  On Error GoTo 0
  MAPISession.SignOff                             'close session
  Exit Function

MailError:
  If Err.Number <> 32001 Then
     MsgBox "Unable to send Email:" & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description & vbCrLf, _
            vbOKOnly Or vbCritical, "eMail Send Error"
  End If
  MAPISession.SignOff                             'close session
  SendEMail = False
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

