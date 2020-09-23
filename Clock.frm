VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClock 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "rsa's clock"
   ClientHeight    =   5985
   ClientLeft      =   3195
   ClientTop       =   3480
   ClientWidth     =   7305
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtfont 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtbackground 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Timer Timer5 
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   3240
      Top             =   3360
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2760
      Top             =   3360
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3000
      Picture         =   "Clock.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   3840
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2400
      Picture         =   "Clock.frx":074C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   3840
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "Clock.frx":0A56
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   3840
      Width           =   540
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2280
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   3360
   End
   Begin VB.Label lblHours2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblMinutes1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl10seconds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblHours1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours1"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1260
   End
   Begin VB.Label lblseconds1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblSeconds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblMinutes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label lblHours 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2100
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu MnuTime1 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuDate1 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit1 
         Caption         =   "E&xit"
      End
      Begin VB.Menu MnuHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu MnuShow 
         Caption         =   "&Show"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuTimer 
         Caption         =   "Seconds &Since Midnight..."
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuProgressBar 
         Caption         =   "Show &Progress Bar"
         Shortcut        =   ^P
      End
      Begin VB.Menu Mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDefault 
         Caption         =   "&Default"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuColour 
         Caption         =   "&Colour"
         Begin VB.Menu MnuFont 
            Caption         =   "&Font..."
            Shortcut        =   ^F
         End
         Begin VB.Menu MnuBackground 
            Caption         =   "&Background..."
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu MnuFont1 
         Caption         =   "&Font"
         Begin VB.Menu MnuBold 
            Caption         =   "&Bold"
            Shortcut        =   ^C
         End
         Begin VB.Menu MnuItalics 
            Caption         =   "&Italics"
            Shortcut        =   ^I
         End
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuTime 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDate 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    On Error Resume Next
    
    Timer1.Interval = 1
    Timer2.Interval = 1000
    Timer3.Interval = 1000
    Timer4.Interval = 1000
    
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = False
    Timer4.Enabled = False
    
    lbl10seconds.Caption = "00"
    lblSeconds.Caption = "00"
    lblMinutes.Caption = "00"
    lblHours.Caption = "00"
    
    lblSeconds.Visible = False
    lblMinutes.Visible = False
    lblHours.Visible = False
    
    lblseconds1.Caption = lblSeconds.Caption
    lblMinutes1.Caption = lblMinutes.Caption
    lblHours1.Caption = lblHours.Caption
    lblHours2.Caption = lblHours1.Caption
    
    frmClock.Height = "1845"
    frmClock.Width = "2910"
    Screen1 = Screen.Height
    Screen2 = frmClock.Height + 450
    
    frmClock.Left = "0"
    frmClock.Top = Screen1 - Screen2
    frmClock.Icon = Picture1.Picture
    frmClock.Caption = "rsa's clock"
    
    Pb1.Enabled = False
    Pb1.Max = "60"
    Pb1.Min = "0"
    MnuBold.Checked = True
    txtfont.Text = lblHours.ForeColor
    txtbackground.Text = frmClock.BackColor
    
    With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Msg As Integer
    
    On Error Resume Next
    
    Msg = MsgBox("Are you sure that you want to exit?", vbExclamation + vbYesNo, "Warning")
    
    If Msg = vbYes Then
    Shell_NotifyIcon NIM_DELETE, IconData
    End
    End If
    
    If Msg = vbNo Then
    Cancel = True
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If frmClock.WindowState = 1 Then
    With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmClock.Hide
    frmClock.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    End If
    
End Sub

Private Sub lbl10seconds_Change()
    On Error Resume Next
    
    If lbl10seconds.Caption = "98" Then
    lbl10seconds.Caption = "00"
    End If
    
End Sub


Private Sub lblHours_Change()
    Dim Hours2 As String
    
    If lblHours.Caption >= 13 Then
    lblHours.Caption = lblHours.Caption - 12
    End If
    
    Hours2 = lblHours.Caption
    
    If Len(lblHours.Caption) < 2 Then
    lblHours.Caption = "0" + Hours2
    End If
    
End Sub

Private Sub lblMinutes_Change()
    Dim Minutes2 As String
    
    On Error Resume Next
    
    Minutes2 = lblMinutes.Caption
    
    If Len(lblMinutes.Caption) < 2 Then
    lblMinutes.Caption = "0" + Minutes2
    End If
    
End Sub

Private Sub lblSeconds_Change()
    Dim Seconds2 As String
    
    On Error Resume Next
    
    Seconds2 = lblSeconds.Caption
    
    If Len(lblSeconds.Caption) < 2 Then
    lblSeconds.Caption = "0" + Seconds2
    End If
    
End Sub

Private Sub lblseconds1_Change()
    On Error Resume Next
    
    lbl10seconds.Caption = "00"
    
    Pb1.Value = Pb1.Value + 1
    
    If Pb1.Value = "60" Then
    Pb1.Value = "0"
    End If
    
End Sub

Private Sub MnuAbout_Click()
    On Error Resume Next
    
    frmAbout.Show
    
End Sub

Private Sub MnuBackground_Click()
    On Error Resume Next
    
    cd.CancelError = True
    cd.Action = 3
    
    If Err.Number = cdlCancel Then
    Err.Clear
    cd.Color = frmClock.BackColor
    Else
    frmClock.BackColor = cd.Color
    txtbackground.Text = frmClock.BackColor
    End If
    
End Sub

Private Sub MnuBold_Click()
    On Error Resume Next

    If MnuBold.Checked = False Then
    MnuBold.Checked = True
    lblHours2.FontBold = True
    lblMinutes1.FontBold = True
    lblseconds1.FontBold = True
    lbl10seconds.FontBold = True
    lblDate.FontBold = True
    Label1(0).FontBold = True
    Label1(1).FontBold = True
    Label1(2).FontBold = True
    Else
    MnuBold.Checked = False
    lblHours2.FontBold = False
    lblMinutes1.FontBold = False
    lblseconds1.FontBold = False
    lbl10seconds.FontBold = False
    lblDate.FontBold = False
    Label1(0).FontBold = False
    Label1(1).FontBold = False
    Label1(2).FontBold = False
    End If
    
End Sub

Private Sub MnuDefault_Click()
    On Error Resume Next
    
    frmClock.BackColor = &H0&
    lblHours2.ForeColor = &HFF&
    lblMinutes1.ForeColor = &HFF&
    lblseconds1.ForeColor = &HFF&
    lbl10seconds.ForeColor = &HFF&
    lblDate.ForeColor = &HFF&
    Label1(0).ForeColor = &HFF&
    Label1(1).ForeColor = &HFF&
    Label1(2).ForeColor = &HFF&
    
End Sub

Private Sub MnuExit_Click()
    On Error Resume Next
    
    Unload Me
    
End Sub

Private Sub MnuExit1_Click()
    On Error Resume Next
    
    Call MnuExit_Click
    
End Sub

Private Sub MnuFont_Click()
    On Error Resume Next
    
    cd.CancelError = True
    cd.Action = 3
    
    If Err.Number = cdlCancel Then
    Err.Clear
    cd.Color = lblHours.ForeColor
    Else
    lblHours2.ForeColor = cd.Color
    lblMinutes1.ForeColor = cd.Color
    lblseconds1.ForeColor = cd.Color
    lbl10seconds.ForeColor = cd.Color
    lblDate.ForeColor = cd.Color
    Label1(0).ForeColor = cd.Color
    Label1(1).ForeColor = cd.Color
    Label1(2).ForeColor = cd.Color
    txtfont.Text = lblHours2.ForeColor
    End If
End Sub

Private Sub MnuHide_Click()
    On Error Resume Next
    
    frmClock.WindowState = 1
    
End Sub

Private Sub MnuItalics_Click()
    On Error Resume Next
    
    If MnuItalics.Checked = False Then
    MnuItalics.Checked = True
    lblHours2.FontItalic = True
    lblMinutes1.FontItalic = True
    lblseconds1.FontItalic = True
    lbl10seconds.FontItalic = True
    lblDate.FontItalic = True
    Label1(0).FontItalic = True
    Label1(1).FontItalic = True
    Label1(2).FontItalic = True
    Else
    MnuItalics.Checked = False
    lblHours2.FontItalic = False
    lblMinutes1.FontItalic = False
    lblseconds1.FontItalic = False
    lbl10seconds.FontItalic = False
    lblDate.FontItalic = False
    Label1(0).FontItalic = False
    Label1(1).FontItalic = False
    Label1(2).FontItalic = False
    End If
    
End Sub

Private Sub MnuProgressBar_Click()
    Dim Screen1, Screen2 As String
  
    If MnuProgressBar.Checked = False Then
    MnuProgressBar.Checked = True
    Pb1.Value = lblseconds1.Caption
    Pb1.Max = "60"
    Pb1.Min = "0"
    Pb1.Enabled = True
    frmClock.Height = "2325"
    
    Screen1 = Screen.Height
    Screen2 = frmClock.Height + 450
    frmClock.Top = Screen1 - Screen2
    
    Else
    MnuProgressBar.Checked = False
    Pb1.Enabled = False
    frmClock.Height = "1845"
    
    Screen1 = Screen.Height
    Screen2 = frmClock.Height + 450
    frmClock.Top = Screen1 - Screen2
    End If
    
End Sub

Private Sub MnuShow_Click()
    On Error Resume Next
    
    frmClock.WindowState = vbNormal
    frmClock.Show
      
      With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmClock.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    
End Sub

Private Sub MnuTimer_Click()
    Dim Msg As Integer
    
    On Error Resume Next
    
    Msg = MsgBox(Timer, vbExclamation + vbOKOnly, "Warning")
    
    
    
End Sub

Private Sub Timer1_Timer()
    Dim Time1 As String
    
    On Error Resume Next
    
    lblSeconds.Caption = Second(Time)
    lblMinutes.Caption = Minute(Time)
    lblHours.Caption = Hour(Time)
    lblseconds1.Caption = lblSeconds.Caption
    lblMinutes1.Caption = lblMinutes.Caption
    lblHours1.Caption = lblHours.Caption
    lblHours2.Caption = lblHours1.Caption
    lblDate.Caption = Date
    lbl10seconds.Caption = lbl10seconds.Caption + 2
    
    Time1 = lblHours1.Caption + ":" + lblMinutes1.Caption + ":" + lblseconds1.Caption
    
    frmClock.Caption = "rsa's clock - " + Time1
    frmAbout.Caption = "about - " & frmClock.Caption
    
    MnuTime.Caption = Time1
    MnuDate.Caption = Date
    MnuTime1.Caption = Time1
    MnuDate1.Caption = Date
    
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    
    If frmClock.Icon = Picture1.Picture Then
    Shell_NotifyIcon NIM_MODIFY, IconData
    frmClock.Icon = Picture2.Picture
    frmAbout.Icon = frmClock.Icon
    
      With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmClock.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    
    Timer2.Enabled = False
    Timer3.Enabled = True
    End If
    
    
End Sub

Private Sub Timer3_Timer()
    On Error Resume Next
    
    If frmClock.Icon = Picture2.Picture Then
    Shell_NotifyIcon NIM_MODIFY, IconData
    frmClock.Icon = Picture3.Picture
    frmAbout.Icon = frmClock.Icon
    
      With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmClock.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    
    Timer3.Enabled = False
    Timer4.Enabled = True
    End If
    
End Sub

Private Sub Timer4_Timer()
    On Error Resume Next
    
    If frmClock.Icon = Picture3.Picture Then
    Shell_NotifyIcon NIM_MODIFY, IconData
    frmClock.Icon = Picture1.Picture
    frmAbout.Icon = frmClock.Icon
    
      With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmClock.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmClock.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    
    Timer4.Enabled = False
    Timer2.Enabled = True
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long

    On Error Resume Next

    Msg = X

    If Msg = WM_LBUTTONDBLCLK Then
        Call MnuShow_Click
    Else
        If Msg = WM_RBUTTONDOWN Then
        PopupMenu MnuPopup
        End If
    End If
    
End Sub
