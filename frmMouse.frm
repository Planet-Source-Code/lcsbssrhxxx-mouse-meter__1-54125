VERSION 5.00
Begin VB.Form frmMouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Mouse Meter"
   ClientHeight    =   4695
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7095
   Icon            =   "frmMouse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtRClicks 
         Height          =   525
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   4
         Text            =   "frmMouse.frx":0442
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtLClicks 
         Height          =   525
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         Text            =   "frmMouse.frx":0444
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtClicks 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         Text            =   "frmMouse.frx":0446
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtMClicks 
         Height          =   525
         Left            =   5160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   1
         Text            =   "frmMouse.frx":0448
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Clicks :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Left Clicks :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Right Clicks :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle / Wheel Clicks :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   2655
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
      Begin VB.TextBox SpeedKm 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   21
         Text            =   "frmMouse.frx":044A
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox SpeedCm 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         Text            =   "frmMouse.frx":0452
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox SpeedM 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Text            =   "frmMouse.frx":045A
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilometers/Hour :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Centimeters/Second :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Meters/Second :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
      Begin VB.TextBox txtDistanceKm 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   19
         Text            =   "frmMouse.frx":0462
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtDistanceM 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   16
         Text            =   "frmMouse.frx":046A
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtDistanceCm 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   15
         Text            =   "frmMouse.frx":0472
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Kilometers) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Meters) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Centimeters) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   3720
      TabIndex        =   23
      Top             =   1920
      Width           =   1575
      Begin VB.TextBox txtDistanceFt 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   26
         Text            =   "frmMouse.frx":047A
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtDistanceIn 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   25
         Text            =   "frmMouse.frx":0482
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDistanceMi 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   24
         Text            =   "frmMouse.frx":048A
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Feet) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Inches) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance (Miles) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2655
      Left            =   5400
      TabIndex        =   30
      Top             =   1920
      Width           =   1575
      Begin VB.TextBox SpeedMi 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   35
         Text            =   "frmMouse.frx":0492
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox SpeedFt 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   32
         Text            =   "frmMouse.frx":049A
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox SpeedIn 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   31
         Text            =   "frmMouse.frx":04A2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Miles/Hour :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Feet/Second :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Inches/Second :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   5400
      TabIndex        =   37
      Top             =   960
      Width           =   1575
      Begin VB.TextBox SpeedPix 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   38
         Text            =   "frmMouse.frx":04AA
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels/Second :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   3720
      TabIndex        =   40
      Top             =   960
      Width           =   1575
      Begin VB.TextBox txtDistance 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   41
         Text            =   "frmMouse.frx":04B2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance    (Pixels) :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRevert 
         Caption         =   "Revert To Saved"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuClear 
         Caption         =   "Delete File"
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'           ***************************************************
'           *                  Mouse Meter                    *
'           *      By Mike Plaehn (LCSBSSRHXXX) 4/27/04       *
'           ***************************************************

Option Explicit
'[Type PointAPI For Mouse Position And Mouse Distance]
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'[Type NotifyIconData For Tray Icon]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'[Constants]
Const NIM_ADD = &H0 'Add to Tray
Const NIM_MODIFY = &H1 'Modify Details
Const NIM_DELETE = &H2 'Remove From Tray
Const NIF_MESSAGE = &H1 'Message
Const NIF_ICON = &H2 'Icon
Const NIF_TIP = &H4 'TooTipText
Const WM_MOUSEMOVE = &H200 'On Mousemove
Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Const WM_RBUTTONDOWN = &H204 'Right Button Down
Const WM_RBUTTONUP = &H205 'Right Button Up
Const WM_RBUTTONDBLCLK = &H206 'Right Double Click
'[API]
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'[Variables For Tray Icon]
Dim TrayIco As NOTIFYICONDATA
Dim InTray As Boolean
'[Old X, Old Y]
Dim oldX As Integer
Dim oldY As Integer
'[Variables For Speed]
Dim xStart As Long
Dim xEnd As Long
'[Variables For Saving / Loading]
Dim C As Long
Dim LC As Long
Dim RC As Long
Dim MC As Long
Dim D, DIn, DFt, DCm, DM As Long
Private Sub Form_Load()
    Call SetCursorPos(0, 0)
    '[Load Data]
    On Error Resume Next
    'open file for input then load data from file
    Open App.Path & "\MOUSEDATA.dat" For Input As #1
        'load variables
        Input #1, C, LC, RC, MC, D, DIn, DFt, DCm, DM
        txtClicks.Text = C
        txtLClicks.Text = LC
        txtRClicks.Text = RC
        txtMClicks.Text = MC
        txtDistance.Text = D
        txtDistanceIn.Text = DIn
        txtDistanceFt.Text = DFt
        txtDistanceCm.Text = DCm
        txtDistanceM.Text = DM
    Close #1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case InTray
        Case True
            'if me is in in tray and you click me then
            If Button = 1 Then
                'restore the form
                Me.WindowState = vbNormal
                'show the form
                Me.Show
            End If
        'if me isn't in the try and you click me then
        Case False
            'exit sub
            Exit Sub
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, TrayIco
    '[Save Data]
    'define variables
    C = txtClicks.Text
    LC = txtLClicks.Text
    RC = txtRClicks.Text
    MC = txtMClicks.Text
    D = txtDistance.Text
    DIn = txtDistanceIn.Text
    DFt = txtDistanceFt.Text
    DCm = txtDistanceCm.Text
    DM = txtDistanceM.Text
    'open file for output then write data to file
    Open App.Path & "\MOUSEDATA.dat" For Output As #1
        Write #1, C, LC, RC, MC, D, DIn, DFt, DCm, DM
    Close #1
    MsgBox "Data Saved!", vbInformation, "Data Saved!"
End Sub
Private Sub mnuAbout_Click()
    MsgBox "About Mouse Meter:" & vbCrLf & "Mouse Meter was written by LCSBSSRHXXX" & vbCrLf & "With Microsoft Visual Basic 6.0" & vbCrLf & "On 5/25/04 to 5/27/04" & vbCrLf & vbCrLf & "If you have any questions contact me at:" & vbCrLf & "AIM : LCSBSSRHXXX" & vbCrLf & "MSN / Email : kodbooger@hotmail.com", vbInformation, "About Mouse Meter"
End Sub
Private Sub mnuClear_Click()
Dim xRes As VbMsgBoxResult
On Error Resume Next
    'message box (yes, no)
    xRes = MsgBox("Are you sure you want" & vbCrLf & "to delete file?", vbYesNo + vbQuestion, "Mouse Meter")
    'if yes
    If xRes = vbYes Then
        'delete file
        Kill (App.Path & "\MOUSEDATA.dat")
        MsgBox "File sucuessfully deleted." & vbCrLf & "Please reastart Mouse Meter.", vbExclamation, "Mouse Meter"
        End
    'if no
    Else
        Exit Sub
    End If
End Sub
Private Sub mnuExit_Click()
    End
End Sub
Private Sub mnuRevert_Click()
Dim xRes As VbMsgBoxResult
On Error Resume Next
    'message box (yes, no)
    xRes = MsgBox("Are you sure you want" & vbCrLf & "to revert to saved?", vbYesNo + vbQuestion, "Mouse Meter")
    'if yes
    If xRes = vbYes Then
        '[Load Data]
        'open file for input then load data from file
        Open App.Path & "\MOUSEDATA.dat" For Input As #1
            Input #1, C, LC, RC, MC, D, DIn, DFt, DCm, DM
            'load variables
            txtClicks.Text = C
            txtLClicks.Text = LC
            txtRClicks.Text = RC
            txtMClicks.Text = MC
            txtDistance.Text = D
            txtDistanceIn.Text = DIn
            txtDistanceFt.Text = DFt
            txtDistanceCm.Text = DCm
            txtDistanceM.Text = DM
        Close #1
    'if no
    Else
        Exit Sub
    End If
End Sub
Private Sub mnuSave_Click()
    '[Save Data]
    'define variables
    C = txtClicks.Text
    LC = txtLClicks.Text
    RC = txtRClicks.Text
    MC = txtMClicks.Text
    D = txtDistance.Text
    DIn = txtDistanceIn.Text
    DFt = txtDistanceFt.Text
    DCm = txtDistanceCm.Text
    DM = txtDistanceM.Text
    'open file for output then write data to file
    Open App.Path & "\MOUSEDATA.dat" For Output As #1
        Write #1, C, LC, RC, MC, D, DIn, DFt, DCm, DM
        MsgBox "Data Saved!", vbInformation, "Data Saved!"
    Close #1
End Sub
Private Sub tmrCheck_Timer()
Dim Posit As POINTAPI
Dim distance As Long
Dim keyresult As Long
On Error Resume Next
    
    '[Toatal Clicks]
    txtClicks.Text = Val(txtLClicks.Text) + Val(txtRClicks.Text) + Val(txtMClicks.Text)
    '[Left Click]
    'if you press the left mouse then
    keyresult = GetAsyncKeyState(vbLeftButton)
    If keyresult = -32767 Then
        'add to the value of the left click text box
        txtLClicks.Text = Val(txtLClicks.Text + 1)
    End If
    '[Right Click]
    'if you press the right mouse then
    keyresult = GetAsyncKeyState(vbRightButton)
    If keyresult = -32767 Then
        'add to the value of the right click text box
        txtRClicks.Text = Val(txtRClicks.Text + 1)
    End If
    '[Middle Click]
    'if you press middle mouse then
    keyresult = GetAsyncKeyState(vbMiddleButton)
    If keyresult = -32767 Then
        'add to the value of the middle click text box
        txtMClicks.Text = Val(txtMClicks.Text + 1)
    End If
    
    '[Position]
    '[Get Cursor Position]
    'get the cursor position
    GetCursorPos Posit
    'display the cursor position on the form's caption
    Me.Caption = "(X: " & Posit.X & ", Y:" & Posit.Y & ")"
    
    '[Distance]
    'distance = absolute value(Y1 - Y2) + (X1 - X2)
    distance = Abs((Val(Posit.Y) - Val(oldY)) + (Val(Posit.X) - Val(oldX)))
    '[Pixels]
    txtDistance.Text = Val(txtDistance.Text) + Val(distance)
    
    'Distance / 102.7618606 = Inches
    'Distance / 41.10474424 = Centimeters
    'The pixel to inches / centimeters was really hard to figure out!
    '.4 Inches Per Centimeter
    '2.5 Centimeters Per Inch
    
    '[Starndard]
    
    '[Inches]
    txtDistanceIn.Text = Format(Val(txtDistanceIn.Text) + Val(distance) / 102.7618606, "00.00")
    '[Feet]
    txtDistanceFt.Text = Format(Val(txtDistanceIn.Text) / 12, "00.00")
    '[Miles]
    txtDistanceMi.Text = Format(Val(txtDistanceFt.Text) / 5280, "00.00000")
    
    '[Meteric]
    
    '[Centimeters]
    txtDistanceCm.Text = Format(Val(txtDistanceCm.Text) + Val(Int(distance)) / 41.10474424, "00.00")
    '[Meters]
    txtDistanceM.Text = Format(Val(txtDistanceCm.Text / 100), "00.00")
    '[Kilometers]
    txtDistanceKm = Format(Val(txtDistanceM.Text / 1000), "00.00000")
    
    '[Tray]
    'if minimize
    If Me.WindowState = 1 Then
        'make variable InTray = true
        InTray = True
        'Hide frmMain
        Me.Hide
        With TrayIco
            .cbSize = Len(TrayIco)
            'tray icon hwnd = me.hwnd
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            'call back message on mouse move
            .uCallBackMessage = WM_MOUSEMOVE
            'tray icon
            .hIcon = Me.Icon
            'tray ToolTipText
            .szTip = "Mouse Meter" & vbNullChar
        End With
        'add tray icon with the properties of TrayIcon
        Shell_NotifyIcon NIM_ADD, TrayIco
    Else
        'remove from tray if the window is not minimized or in the tray
        InTray = False
        Shell_NotifyIcon NIM_DELETE, TrayIco
    End If
    
    '[Get OldX And OldY]
    oldX = Posit.X
    oldY = Posit.Y
End Sub
Private Sub tmrSpeed_Timer()
    xEnd = txtDistance.Text
    '[Pixels Per Second]
    SpeedPix.Text = Format(Val(xEnd) - Val(xStart), "00.00")
    '[Inches Per Second]
    SpeedIn.Text = Format(Val(SpeedPix.Text) / 102.7618606, "00.00")
    '[Feet Per Second]
    SpeedFt.Text = Format(SpeedIn.Text / 12, "00.00")
    '[Miles Per Hour]
    SpeedMi.Text = Format(SpeedFt.Text / 5280, "00.00000") * 3600
    '[Centimeters Per Second]
    SpeedCm.Text = Format(Val(SpeedPix.Text) / 41.10474424, "00.00")
    '[Meters Per Second]
    SpeedM.Text = Format(SpeedCm.Text / 100, "00.00")
    '[Kilometers Per Hour]
    SpeedKm.Text = Format(SpeedM.Text / 1000, "00.00000") * 3600
    xStart = txtDistance.Text
End Sub

